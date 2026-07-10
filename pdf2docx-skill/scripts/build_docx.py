#!/usr/bin/env python3
"""
build_docx.py —— DOCX 重建（核心模块）

职责：遍历对齐合并后的 blocks，按 type 分发，用 python-docx 重建 DOCX。
这是工作量所在——把 MinerU 的版面结构 + PyMuPDF 的字符样式，
还原成接近"用户手动用 Word 另存的 DOCX"。

验收标准（P0 必须达成）：
  - 标题层级正确：Heading 1/2/3 + OutlineLevel（下游章节分割强依赖）
  - 正文段落完整：无大面积丢失
  - 表格结构还原：行列基本对齐（复杂合并单元格留 TODO）
  - 字号字体基本还原

依赖：python-docx

用法（被 pdf2docx.py 调用）:
    from build_docx import build_docx
    build_docx(merged_data, images_dir, "output/结果.docx")

⚠️ 骨架阶段实现边界：
  - title（含 OutlineLevel）：完整实现 ✅
  - text：完整实现 ✅
  - table：HTML 解析框架搭好，rowspan/colspan 合并留 TODO ⏳
  - image：基本实现（按顺序插入）✅
  - list：降级为普通段落 ✅
"""

from __future__ import annotations

import re
import sys
from pathlib import Path
from typing import Any

# 标题层级正则（用于修正 MinerU text_level 不准的情况）
# 优先级：正则命中 > MinerU text_level
# TODO(调优): 见 references/tuning-guide.md，根据真实招标 PDF 扩充规则
_RE_CHAPTER = re.compile(r"^第[一二三四五六七八九十百千\d]+[章节篇部]")
_RE_LEVEL1 = re.compile(r"^\d+[\.、]")           # "1." / "1、"
_RE_LEVEL2 = re.compile(r"^\d+\.\d+")            # "1.1"
_RE_LEVEL3 = re.compile(r"^\d+\.\d+\.\d+")       # "1.1.1"


# ═══════════════════════════════════════════════════════════════
#  文本提取辅助
# ═══════════════════════════════════════════════════════════════

def _extract_text(block: dict[str, Any]) -> str:
    """
    从 MinerU block 提取纯文本。
    兼容多种结构：block.lines.spans.content / block.content
    """
    # 优先从 lines.spans 拼接
    parts = []
    for line in block.get("lines", []):
        for span in line.get("spans", []):
            content = span.get("content") or span.get("text") or ""
            parts.append(content)
    if parts:
        return "".join(parts)

    # 回退：顶层 content 字段
    return block.get("content", "")


def _get_dominant_style(block: dict[str, Any]) -> dict[str, Any] | None:
    """
    取 block 内占比最大的样式（用于 title 等需要整体设样式的块）。
    优先用 block._style（align.py 兜底贴的），否则统计内部 span。
    """
    # align.py 兜底贴的 block 级样式
    if block.get("_style"):
        return block["_style"]

    # 统计 lines.spans 内的样式
    style_counter: dict[str, int] = {}
    style_map: dict[str, dict] = {}
    for line in block.get("lines", []):
        for span in line.get("spans", []):
            st = span.get("_style")
            if not st:
                continue
            # 用 size 作为指纹（标题通常统一字号）
            key = str(st.get("size", 0))
            style_counter[key] = style_counter.get(key, 0) + 1
            style_map[key] = st

    if not style_counter:
        return None
    dominant_key = max(style_counter, key=style_counter.get)
    return style_map.get(dominant_key)


def _guess_heading_level(text: str, mineru_level: int = 1) -> int:
    """
    推断标题层级。
    优先用文本正则，正则不命中才用 MinerU 的 text_level。
    """
    text = text.strip()
    # 正则优先（更可靠）
    if _RE_CHAPTER.match(text):
        return 1
    if _RE_LEVEL3.match(text):
        return 3
    if _RE_LEVEL2.match(text):
        return 2
    if _RE_LEVEL1.match(text):
        return 1
    # 回退 MinerU text_level
    try:
        level = int(mineru_level)
    except (TypeError, ValueError):
        level = 1
    return min(max(level, 1), 9)


# ═══════════════════════════════════════════════════════════════
#  OOXML 辅助：OutlineLevel（P0 关键）
# ═══════════════════════════════════════════════════════════════

def _set_outline_level(paragraph, level: int) -> None:
    """
    显式设置段落的 w:outlineLvl。

    ⚠️ P0 关键：下游 splitDocumentByChapters 的 getOutlineLevel()
    强依赖此属性（ChapterSplitServiceImpl.java:268/354/1485-1497）。
    add_heading() 会设 Heading 样式，但显式设 outlineLvl 是双保险。

    依据：python-docx #746，OOXML 规范 w:outlineLvl。
    """
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    pPr = paragraph._p.get_or_add_pPr()
    # 先移除已存在的 outlineLvl（避免重复）
    for existing in pPr.findall(qn("w:outlineLvl")):
        pPr.remove(existing)
    outline = OxmlElement("w:outlineLvl")
    outline.set(qn("w:val"), str(level))
    pPr.append(outline)


# ═══════════════════════════════════════════════════════════════
#  样式应用
# ═══════════════════════════════════════════════════════════════

# PDF 字体名 → 常用中文字体名映射
# 已根据真实招标 PDF（新疆磋商文件）实测补充
# TODO(调优): 见 references/tuning-guide.md，遇到新字体继续扩充
_FONT_MAP = {
    # 宋体类
    "simsun": "宋体", "song": "宋体", "songti": "宋体",
    "stsong": "华文宋体", "stzhongs": "华文中宋",
    "fzshusong": "方正书宋", "fzsongs": "方正宋体",
    # 黑体类
    "simhei": "黑体", "hei": "黑体", "heiti": "黑体",
    "stheiti": "华文黑体", "stxihei": "华文细黑",
    "fzhei": "方正黑体", "fzht": "方正黑体",
    # 标题用方正小标宋（招标文件封面大标题常见，实测：XJHY磋商文件用了 FZXBSJW）
    "fzxbsjw": "方正小标宋简体", "fzxbs": "方正小标宋简体",
    "fzxiaobiaosong": "方正小标宋简体",
    # 楷体类
    "kaiti": "楷体", "kai": "楷体", "simkai": "楷体",
    "stkaiti": "华文楷体", "fzkai": "方正楷体",
    # 仿宋类（公文正文常见）
    "fangsong": "仿宋", "fs": "仿宋", "fangsong_gb2312": "仿宋",
    "simfang": "仿宋", "stfangsong": "华文仿宋", "fzfangsong": "方正仿宋",
    # 其他
    "microsoft yahei": "微软雅黑", "yahei": "微软雅黑",
    # 英文/数字部分（Times/Arial 等保留原样，Word 自带）
}


def _map_font_name(pdf_font: str) -> str:
    """PDF 字体名 → 常用字体名映射。"""
    name = pdf_font.lower()
    for key, val in _FONT_MAP.items():
        if key in name:
            return val
    return pdf_font  # 兜底原样保留


def _apply_style(run, style: dict[str, Any] | None) -> None:
    """
    把 PyMuPDF 样式应用到 python-docx run。
    style 格式：{"font": str, "size": float, "bold": bool, "italic": bool, "color": int}
    """
    if not style:
        return

    from docx.shared import Pt, RGBColor

    size = style.get("size")
    if size:
        run.font.size = Pt(size)

    if style.get("bold"):
        run.font.bold = True
    if style.get("italic"):
        run.font.italic = True

    color = style.get("color")
    if color is not None:
        # sRGB 整数 → RGBColor
        # 0=黑，16777215=白
        r = (color >> 16) & 0xFF
        g = (color >> 8) & 0xFF
        b = color & 0xFF
        # 纯黑（0）不设，用默认即可，避免不必要的 XML
        if color != 0:
            run.font.color.rgb = RGBColor(r, g, b)

    font = style.get("font")
    if font:
        from docx.oxml.ns import qn
        from docx.oxml import OxmlElement

        font_name = _map_font_name(font)
        run.font.name = font_name
        # 中文字体需设 EastAsia 属性
        rPr = run._element.get_or_add_rPr()
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is None:
            rFonts = OxmlElement("w:rFonts")
            rPr.append(rFonts)
        rFonts.set(qn("w:eastAsia"), font_name)


# ═══════════════════════════════════════════════════════════════
#  各 block 类型的重建函数
# ═══════════════════════════════════════════════════════════════

def _build_title(doc, block: dict[str, Any]) -> None:
    """
    标题：Heading 样式 + OutlineLevel + 字号字体。
    P0 验收项：层级必须正确。
    """
    text = _extract_text(block).strip()
    if not text:
        return

    mineru_level = block.get("text_level") or block.get("level") or 1
    level = _guess_heading_level(text, mineru_level)

    para = doc.add_heading(text, level=level)

    # 补充字符样式（字号字体粗体）
    style = _get_dominant_style(block)
    if style:
        for run in para.runs:
            _apply_style(run, style)

    # 显式设置 OutlineLevel（双保险，确保下游能识别）
    _set_outline_level(para, level)


def _build_text(doc, block: dict[str, Any]) -> None:
    """
    正文段落。
    按 MinerU 的 lines.spans 粒度设置 run 样式（同段不同样式）。
    """
    text = _extract_text(block)
    if not text.strip():
        return

    para = doc.add_paragraph()

    # 按 lines.spans 粒度建 run
    has_lines = bool(block.get("lines"))
    if has_lines:
        for line in block.get("lines", []):
            line_text = ""
            for mspan in line.get("spans", []):
                content = mspan.get("content") or mspan.get("text") or ""
                line_text += content
            if line_text:
                # 用第一个 span 的样式代表整行（MVP 简化）
                first_style = None
                for mspan in line.get("spans", []):
                    if mspan.get("_style"):
                        first_style = mspan["_style"]
                        break
                run = para.add_run(line_text)
                _apply_style(run, first_style)
    else:
        # 无 lines 结构，整块作为一个 run
        run = para.add_run(text)
        _apply_style(run, block.get("_style"))


def _build_table(doc, block: dict[str, Any]) -> None:
    """
    表格：解析 table_body（HTML，含 rowspan/colspan）→ python-docx 表格。

    MinerU 的 table_body 是 HTML 格式，含 <td rowspan="" colspan="">。

    ⚠️ 骨架阶段：HTML 解析框架搭好，合并单元格处理留 TODO。
    复杂表格（跨页表、嵌套表、深度合并）需真实 PDF 实测后完善。
    """
    table_html = block.get("table_body") or block.get("html") or ""
    if not table_html:
        return

    rows = _parse_html_table(table_html)
    if not rows:
        return

    # 计算实际列数（考虑 colspan）
    max_cols = max(sum(c.get("colspan", 1) for c in row) for row in rows)
    table = doc.add_table(rows=len(rows), cols=max_cols)
    try:
        table.style = "Table Grid"
    except KeyError:
        pass  # 无 Table Grid 样式时用默认

    _fill_table_cells(table, rows)


def _parse_html_table(html: str) -> list[list[dict[str, Any]]]:
    """
    解析 HTML 表格为 [[{text, rowspan, colspan}, ...], ...]。

    ⚠️ MVP 用正则粗解析（够用），生产级建议用 lxml/html.parser。
    TODO(完善): 处理 <thead>/<tbody>、嵌套表、跨页表头重复等边界情况。
    """
    rows: list[list[dict[str, Any]]] = []

    # 按行分割 <tr>...</tr>
    tr_pattern = re.compile(r"<tr[^>]*>(.*?)</tr>", re.DOTALL | re.IGNORECASE)
    # 匹配整个 <td>/<th> 单元格：捕获组1=标签名，组2=属性串，组3=单元格内容
    td_pattern = re.compile(
        r"<(td|th)([^>]*)>(.*?)</\1>",
        re.DOTALL | re.IGNORECASE,
    )
    # 从属性串里提取 rowspan / colspan 值
    rowspan_pat = re.compile(r"rowspan\s*=\s*['\"]?(\d+)", re.IGNORECASE)
    colspan_pat = re.compile(r"colspan\s*=\s*['\"]?(\d+)", re.IGNORECASE)

    for tr_match in tr_pattern.finditer(html):
        row: list[dict[str, Any]] = []
        tr_content = tr_match.group(1)
        for td_match in td_pattern.finditer(tr_content):
            attrs = td_match.group(2) or ""
            rowspan_m = rowspan_pat.search(attrs)
            colspan_m = colspan_pat.search(attrs)
            rowspan = int(rowspan_m.group(1)) if rowspan_m else 1
            colspan = int(colspan_m.group(1)) if colspan_m else 1
            # 去除单元格内的 HTML 标签
            cell_html = td_match.group(3)
            cell_text = re.sub(r"<[^>]+>", "", cell_html).strip()
            cell_text = re.sub(r"&nbsp;", " ", cell_text)
            cell_text = re.sub(r"&amp;", "&", cell_text)
            row.append({"text": cell_text, "rowspan": rowspan, "colspan": colspan})
        if row:
            rows.append(row)

    return rows


def _fill_table_cells(table, rows: list[list[dict[str, Any]]]) -> None:
    """
    填充表格单元格，处理 rowspan/colspan 合并。

    ⚠️ 骨架阶段：基础合并已实现，复杂场景（深度嵌套合并）留 TODO。
    python-docx 的合并通过 cell.merge() 实现。
    """
    # 跟踪哪些单元格已被合并占用
    # occupied[row][col] = True 表示该位置被上方/左边的合并单元格占用
    occupied: dict[tuple[int, int], bool] = {}

    for r_idx, row in enumerate(rows):
        c_idx = 0
        for cell_data in row:
            # 跳过被占用的列
            while occupied.get((r_idx, c_idx)):
                c_idx += 1
            if c_idx >= len(table.columns):
                break

            rowspan = cell_data.get("rowspan", 1)
            colspan = cell_data.get("colspan", 1)
            text = cell_data.get("text", "")

            try:
                target_cell = table.cell(r_idx, c_idx)
                target_cell.text = text

                # 处理 colspan（水平合并）
                if colspan > 1:
                    end_col = min(c_idx + colspan - 1, len(table.columns) - 1)
                    if end_col > c_idx:
                        merge_cell = table.cell(r_idx, end_col)
                        target_cell = target_cell.merge(merge_cell)

                # 处理 rowspan（垂直合并）
                if rowspan > 1:
                    end_row = min(r_idx + rowspan - 1, len(table.rows) - 1)
                    if end_row > r_idx:
                        merge_cell = table.cell(end_row, c_idx)
                        target_cell = target_cell.merge(merge_cell)

                # 标记被占用的位置
                for dr in range(rowspan):
                    for dc in range(colspan):
                        occupied[(r_idx + dr, c_idx + dc)] = True

            except (IndexError, Exception) as e:
                # 超出表格范围的单元格，跳过
                # TODO(完善): 真实 PDF 上观察是否需要动态扩展表格
                print(f"  ⚠️ 表格单元格填充跳过 ({r_idx},{c_idx}): {e}",
                      file=sys.stderr)
                pass

            c_idx += colspan


def _build_image(doc, block: dict[str, Any], images_dir: str | None) -> None:
    """
    图片：从 images/ 读图插入。
    MVP 按顺序插入，位置精度后续优化。
    """
    from docx.shared import Inches

    img_path = block.get("img_path") or block.get("image_path") or ""
    if not img_path or not images_dir:
        return

    full_path = Path(images_dir) / Path(img_path).name
    if not full_path.exists():
        print(f"  ⚠️ 图片不存在: {full_path}", file=sys.stderr)
        return

    doc.add_picture(str(full_path), width=Inches(5.5))


def _build_list(doc, block: dict[str, Any]) -> None:
    """
    列表：MVP 降级为普通段落（带编号文本）。
    TODO(完善): 根据 MinerU 的 list sub_type 还原编号/项目符号格式。
    """
    for line in block.get("lines", []):
        text = "".join(
            s.get("content") or s.get("text") or ""
            for s in line.get("spans", [])
        )
        if text.strip():
            try:
                doc.add_paragraph(text, style="List Bullet")
            except KeyError:
                doc.add_paragraph(text)


# ═══════════════════════════════════════════════════════════════
#  对外主接口
# ═══════════════════════════════════════════════════════════════

# block type → 重建函数的分发表
# 兼容 MinerU 2.0 多种可能的类型名
_BLOCK_BUILDERS = {
    "title": _build_title,
    "text": _build_text,
    "paragraph": _build_text,      # 别名
    "table": _build_table,
    "image": _build_image,
    "picture": _build_image,       # 别名
    "list": _build_list,
}


def build_docx(
    merged_data: dict[str, Any],
    images_dir: str | None,
    output_path: str | Path,
) -> str:
    """
    遍历合并后的 blocks，按 type 分发重建 DOCX。

    参数：
        merged_data:  align.align_and_merge 的输出（含 _style 的 MinerU 数据）
        images_dir:   图片目录路径（可能为 None）
        output_path:  输出 docx 路径

    返回：
        输出文件的绝对路径。

    处理顺序：按 page_idx 顺序遍历，每页内按 block 顺序。
    """
    from docx import Document

    doc = Document()
    pdf_info = merged_data.get("pdf_info", [])

    block_count = 0
    for page in pdf_info:
        page_idx = page.get("page_idx", 0)
        blocks = page.get("para_blocks") or page.get("blocks") or []

        for block in blocks:
            btype = (block.get("type") or block.get("block_type") or "text").lower()
            builder = _BLOCK_BUILDERS.get(btype, _build_text)  # 未知类型降级为正文

            try:
                if btype in ("image", "picture"):
                    builder(doc, block, images_dir)
                else:
                    builder(doc, block)
                block_count += 1
            except Exception as e:
                # 单个 block 失败不中断整体流程
                print(f"  ⚠️ block 重建失败 (page={page_idx}, type={btype}): {e}",
                      file=sys.stderr)

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))

    print(f"[4/4] DOCX 重建完成: {output_path}（{block_count} 个 block）",
          file=sys.stderr)
    return str(output_path.resolve())


# ═══════════════════════════════════════════════════════════════
#  CLI（可独立运行测试）
# ═══════════════════════════════════════════════════════════════

if __name__ == "__main__":
    import argparse
    import json

    parser = argparse.ArgumentParser(
        description="python-docx DOCX 重建"
    )
    parser.add_argument("merged", help="align.py 输出的 merged.json 路径")
    parser.add_argument("-o", "--output", default="./output/重建结果.docx",
                        help="输出 docx 路径")
    parser.add_argument("--images-dir", default=None,
                        help="图片目录路径")

    args = parser.parse_args()

    merged_data = json.loads(Path(args.merged).read_text(encoding="utf-8"))
    result = build_docx(merged_data, args.images_dir, args.output)

    print(json.dumps({"output": result}, ensure_ascii=False, indent=2))
