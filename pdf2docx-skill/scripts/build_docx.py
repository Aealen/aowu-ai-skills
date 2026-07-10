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

# PDF 字体名 → 系统可用字体名映射
# ⚠️ 只映射到 Windows 确定自带的基础字体（宋体/黑体/仿宋/楷体/微软雅黑）
# 系统没装的字体（方正系列、华文系列）映射到最接近的系统字体
# TODO(调优): 如用户装了方正/华文字体，可扩充映射表
_FONT_MAP = {
    # 宋体类 → 宋体（系统自带 simsun.ttc）
    "simsun": "宋体", "song": "宋体", "songti": "宋体",
    "stsong": "宋体", "stzhongs": "宋体",          # 华文宋体→宋体
    "fzshusong": "宋体", "fzsongs": "宋体",         # 方正书宋→宋体
    # 黑体类 → 黑体（系统自带 simhei.ttf）
    "simhei": "黑体", "hei": "黑体", "heiti": "黑体",
    "stheiti": "黑体", "stxihei": "黑体",           # 华文黑体→黑体
    "fzhei": "黑体", "fzht": "黑体",                # 方正黑体→黑体
    # 方正小标宋（封面大标题）→ 黑体（视觉上比宋体更接近方正小标宋的庄重感）
    "fzxbsjw": "黑体", "fzxbs": "黑体",
    "fzxiaobiaosong": "黑体",
    # 楷体类 → 楷体（系统自带 simkai.ttf）
    "kaiti": "楷体", "kai": "楷体", "simkai": "楷体",
    "stkaiti": "楷体", "fzkai": "楷体",
    # 仿宋类 → 仿宋（系统自带 simfang.ttf）
    "fangsong": "仿宋", "fs": "仿宋", "fangsong_gb2312": "仿宋",
    "simfang": "仿宋", "stfangsong": "仿宋", "fzfangsong": "仿宋",
    # 微软雅黑（系统自带 msyh.ttc）
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


def _get_dominant_size(block: dict[str, Any] | None) -> float:
    """
    从 block 内 span 样式取占比最大的字号。
    用于动态计算行距（字号 × 1.4）。
    """
    if not block:
        return 12.0
    size_counter: dict[float, int] = {}
    for line in block.get("lines", []):
        for sp in line.get("spans", []):
            st = sp.get("_style")
            if st and st.get("size"):
                sz = st["size"]
                size_counter[sz] = size_counter.get(sz, 0) + 1
    if not size_counter:
        return 12.0
    return max(size_counter, key=size_counter.get)


def _apply_paragraph_format(para, block: dict[str, Any] | None = None) -> None:
    """
    根据 block 的 bbox 位置推断段落格式。
    不再硬编码固定格式——从 PDF 的真实坐标还原。

    推断规则（基于 A4 595pt 宽）：
      - 居中：首行中心 ≈ 页面中心(297pt) → CENTER
      - 左对齐 + 首行缩进：首行 x0 > 左边距 + 15pt → 有缩进
      - 左对齐无缩进：首行 x0 ≈ 左边距 → LEFT 无缩进
    """
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    pf = para.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)

    # 行距：固定值24pt
    # 源PDF实测行距约23pt（段内行），但docx渲染需微调
    pf.line_spacing = Pt(24)

    if not block:
        # 无 block 信息时，默认正文格式
        pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        pf.first_line_indent = Pt(24)
        return

    # 从 block 的首行 bbox 推断
    bbox = block.get("bbox", [])
    lines = block.get("lines", [])

    if not bbox or not lines:
        pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        pf.first_line_indent = Pt(24)
        return

    # 首行的 span bbox
    first_spans = lines[0].get("spans", [])
    if not first_spans:
        pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        pf.first_line_indent = Pt(24)
        return

    first_x0 = first_spans[0].get("bbox", [0])[0]
    block_x0 = bbox[0]
    block_x1 = bbox[2]
    block_width = block_x1 - block_x0
    page_width = 595  # A4 宽度 pt
    page_center = page_width / 2

    # 判断居中：首行起始 x0 接近页面中心（容差 40pt）
    # 注意：用首行 span 的 x0 判断，不用 block 中心（block 右边界不可靠）
    if abs(first_x0 - page_center) < 40:
        pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pf.first_line_indent = Pt(0)
    else:
        # 左对齐/两端对齐
        pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        # 首行缩进：用绝对 x0 位置判断（相对页面左边距）
        # 源PDF左边距约82pt，文字x0>100pt 说明有缩进（约2字符）
        page_left_margin = 82  # 源PDF左边距（pt）
        indent = first_x0 - page_left_margin
        if indent > 15:
            pf.first_line_indent = Pt(24)  # 2字符
        else:
            pf.first_line_indent = Pt(0)


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
    保留多行结构（封面标题常有多个独立行，各有不同字号）。
    """
    text = _extract_text(block).strip()
    if not text:
        return

    mineru_level = block.get("text_level") or block.get("level") or 1
    level = _guess_heading_level(text, mineru_level)

    # 手动建段落（不用 add_heading，因为它把整个 text 当一个 run，丢失多行结构）
    from docx.enum.style import WD_STYLE_TYPE
    heading_style_name = f"Heading {level}"
    try:
        para = doc.add_paragraph(style=heading_style_name)
    except KeyError:
        para = doc.add_paragraph()  # 无 Heading 样式时用默认

    # 标题段落格式
    _apply_title_format(para, level)

    # 逐行建 run，保留多行结构（每行末尾加换行）
    lines = block.get("lines", [])
    if lines:
        for li, line in enumerate(lines):
            for sp in line.get("spans", []):
                content = sp.get("content") or sp.get("text") or ""
                if not content:
                    continue
                run = para.add_run(content)
                _apply_style(run, sp.get("_style"))
            # 行之间加换行（最后一行不加）
            if li < len(lines) - 1:
                run = para.add_run()
                run.add_break()
    else:
        run = para.add_run(text)
        _apply_style(run, block.get("_style"))

    # 显式设置 OutlineLevel（双保险，确保下游能识别）
    _set_outline_level(para, level)

    # 强制标题颜色为黑色（覆盖 Heading 样式默认蓝色）
    _force_heading_color_black(para)


def _apply_title_format(para, level: int) -> None:
    """
    标题段落格式。
    H1（章标题）：居中、段前段后 12pt
    H2+（节标题）：左对齐、无缩进、段前 6pt
    行距用单倍（多行标题时行间紧凑，不用正文的 28pt 固定行距）
    """
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING

    pf = para.paragraph_format
    pf.first_line_indent = Pt(0)  # 标题无首行缩进
    pf.line_spacing_rule = WD_LINE_SPACING.SINGLE  # 单倍行距（标题行间紧凑）
    pf.space_before = Pt(0)

    if level == 1:
        # H1（章标题）：源PDF实测标题到下方正文约40pt间距，单倍行距只给21pt
        # 需补约20pt段后间距
        pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pf.space_after = Pt(10)
    else:
        # H2（节标题）：源PDF实测标题到正文约23pt，单倍行距已接近，不需额外间距
        pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf.space_after = Pt(0)


def _force_heading_color_black(para) -> None:
    """
    强制标题段落所有 run 的颜色为黑色。
    python-docx 的 Heading 样式默认是蓝色（accent1: 365F91/4F81BD），
    源 PDF 的标题都是纯黑（color=0），必须覆盖。
    注意：_apply_style 对 color=0 跳过设置，所以需要单独强制。
    """
    from docx.shared import RGBColor
    for run in para.runs:
        run.font.color.rgb = RGBColor(0, 0, 0)


def _build_text(doc, block: dict[str, Any]) -> None:
    """
    正文段落。
    按 MinerU 的 lines.spans 粒度设置 run 样式（同段不同样式）。
    """
    text = _extract_text(block)
    if not text.strip():
        return

    para = doc.add_paragraph()

    # 设置段落格式（根据 block bbox 推断对齐/缩进）
    _apply_paragraph_format(para, block)

    # 按 lines.spans 粒度逐个建 run，保留行内混排
    has_lines = bool(block.get("lines"))
    if has_lines:
        for line in block.get("lines", []):
            for mspan in line.get("spans", []):
                content = mspan.get("content") or mspan.get("text") or ""
                if not content:
                    continue
                run = para.add_run(content)
                _apply_style(run, mspan.get("_style"))
    else:
        # 无 lines 结构，整块作为一个 run
        run = para.add_run(text)
        _apply_style(run, block.get("_style"))


def _extract_table_html(block: dict[str, Any]) -> str:
    """
    从 MinerU table block 提取 HTML。
    MinerU 3.x 的 table HTML 不在 table_body 字段里，
    而是在 blocks[].lines[].spans[].html 里。
    """
    # 路径1：顶层 table_body（部分版本可能有）
    html = block.get("table_body") or ""
    if html:
        return html
    # 路径2：blocks[].lines[].spans[].html（MinerU 3.x 实际位置）
    for sb in block.get("blocks", []):
        for line in sb.get("lines", []):
            for sp in line.get("spans", []):
                h = sp.get("html", "")
                if h:
                    return h
    return ""


def _build_table(doc, block: dict[str, Any]) -> None:
    """
    表格：解析 HTML（含 rowspan/colspan）→ python-docx 表格。
    MinerU 3.x 的 table HTML 在 blocks[].lines[].spans[].html 里。
    """
    table_html = _extract_table_html(block)
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

    _fill_table_cells(table, rows, block.get("_style"))


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


def _fill_table_cells(table, rows: list[list[dict[str, Any]]],
                      style: dict[str, Any] | None = None) -> None:
    """
    填充表格单元格，处理 rowspan/colspan 合并，并应用样式。
    style 是 block._style（align.py 贴的兜底样式），应用到每个单元格的文字。
    """
    occupied: dict[tuple[int, int], bool] = {}

    for r_idx, row in enumerate(rows):
        c_idx = 0
        for cell_data in row:
            while occupied.get((r_idx, c_idx)):
                c_idx += 1
            if c_idx >= len(table.columns):
                break

            rowspan = cell_data.get("rowspan", 1)
            colspan = cell_data.get("colspan", 1)
            text = cell_data.get("text", "")

            try:
                target_cell = table.cell(r_idx, c_idx)
                # 用 run 设文字（而不是 cell.text），这样可以应用样式
                target_cell.text = ""  # 清空
                para = target_cell.paragraphs[0]
                run = para.add_run(text)
                if style:
                    _apply_style(run, style)

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


def _build_index(doc, block: dict[str, Any]) -> None:
    """
    目录（MinerU 的 type=index）：每行一个章节条目，独立段落输出。
    保留行结构，避免所有条目被拼成一个段落。
    TODO: TOC 跳转（书签/超链接）在 MinerU 的 index block 里不保留，无法还原。
    """
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    for line in block.get("lines", []):
        text = "".join(
            s.get("content") or s.get("text") or ""
            for s in line.get("spans", [])
        )
        if not text.strip():
            continue
        para = doc.add_paragraph()
        pf = para.paragraph_format
        pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf.first_line_indent = Pt(0)
        pf.line_spacing = Pt(24)  # 目录行距稍紧凑
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)

        # 逐 span 建 run，保留样式
        for sp in line.get("spans", []):
            content = sp.get("content") or sp.get("text") or ""
            if not content:
                continue
            run = para.add_run(content)
            _apply_style(run, sp.get("_style"))


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
    "index": _build_index,         # 目录（每行独立段落）
}


def _detect_page_margins(pdf_path: str, pdf_info: list) -> dict[str, float] | None:
    """
    用 PyMuPDF 精确提取 PDF 页面边距。
    取多页（跳过首页封面）所有 span 的坐标范围，算出上下左右边距。
    比 MinerU block bbox 精确（MinerU bbox 底部系统性偏低）。
    """
    try:
        import fitz
    except ImportError:
        return None

    try:
        doc = fitz.open(pdf_path)
    except Exception:
        return None

    page_count = doc.page_count
    top_margins = []
    bottom_margins = []
    left_margins = []
    right_margins = []

    # 取第2-10页采样（跳过首页封面）
    for pi in range(1, min(10, page_count)):
        page = doc[pi]
        d = page.get_text("dict")
        pw = page.rect.width
        ph = page.rect.height
        y0s, y1s, x0s, x1s = [], [], [], []
        for block in d.get("blocks", []):
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                for sp in line.get("spans", []):
                    if not sp["text"].strip():
                        continue
                    bx = sp["bbox"]
                    y0s.append(bx[1])
                    y1s.append(bx[3])
                    x0s.append(bx[0])
                    x1s.append(bx[2])
        if y0s and y1s:
            top_margins.append(min(y0s))
            bottom_margins.append(ph - max(y1s))
            left_margins.append(min(x0s))
            right_margins.append(pw - max(x1s))

    doc.close()

    if not top_margins:
        return None

    top_margins.sort()
    bottom_margins.sort()
    left_margins.sort()
    right_margins.sort()
    mid = len(top_margins) // 2

    return {
        "top": top_margins[mid],         # 中位数
        "bottom": min(bottom_margins),   # 最小值（内容最满的页反映真实下边距）
        "left": left_margins[mid],
        "right": right_margins[mid],
    }


def build_docx(
    merged_data: dict[str, Any],
    images_dir: str | None,
    output_path: str | Path,
    pdf_path: str | None = None,
) -> str:
    """
    遍历合并后的 blocks，按 type 分发重建 DOCX。

    参数：
        merged_data:  align.align_and_merge 的输出（含 _style 的 MinerU 数据）
        images_dir:   图片目录路径（可能为 None）
        output_path:  输出 docx 路径
        pdf_path:     源 PDF 路径（用于精确提取页面边距，可选）

    返回：
        输出文件的绝对路径。

    参数：
        merged_data:  align.align_and_merge 的输出（含 _style 的 MinerU 数据）
        images_dir:   图片目录路径（可能为 None）
        output_path:  输出 docx 路径

    返回：
        输出文件的绝对路径。

    处理顺序：按 page_idx 顺序遍历，每页内按 block 顺序。
    利用 block 的 bbox Y 坐标差还原垂直间距（留白）。
    """
    from docx import Document
    from docx.shared import Mm, Pt

    doc = Document()
    pdf_info = merged_data.get("pdf_info", [])

    # ── 页面设置：从 PDF 真实尺寸设置纸张大小和边距 ──
    # python-docx 默认用 US Letter，需改为 PDF 的实际尺寸（通常是 A4）
    if pdf_info:
        first_page = pdf_info[0]
        page_size = first_page.get("page_size", [595, 842])
        pdf_w_pt = page_size[0] if len(page_size) > 0 else 595
        pdf_h_pt = page_size[1] if len(page_size) > 1 else 842

        sec = doc.sections[0]
        # pt → mm（1pt = 0.3528mm）
        sec.page_width = Mm(pdf_w_pt * 0.3528)
        sec.page_height = Mm(pdf_h_pt * 0.3528)

        # 边距：优先用 PyMuPDF 精确提取（MinerU bbox 底部偏低）
        margins = _detect_page_margins(pdf_path, pdf_info) if pdf_path else None

        if margins:
            sec.left_margin = Pt(max(margins["left"], 28))
            sec.right_margin = Pt(max(margins["right"], 28))
            sec.top_margin = Pt(max(margins["top"], 28))
            sec.bottom_margin = Pt(max(margins["bottom"], 20))
        else:
            # 无 pdf_path 时从 MinerU block bbox 推断（精度较低）
            first_blocks = first_page.get("para_blocks") or first_page.get("blocks") or []
            if first_blocks:
                all_x0 = [b["bbox"][0] for b in first_blocks if b.get("bbox")]
                all_x1 = [b["bbox"][2] for b in first_blocks if b.get("bbox")]
                if all_x0 and all_x1:
                    sec.left_margin = Pt(max(min(all_x0), 28))
                    sec.right_margin = Pt(max(pdf_w_pt - max(all_x1), 28))
            sec.top_margin = Mm(25.4)
            sec.bottom_margin = Mm(25.4)

    block_count = 0
    prev_block_bottom = None  # 上一个 block 的 Y1 底部坐标

    for page in pdf_info:
        page_idx = page.get("page_idx", 0)
        page_size = page.get("page_size", [595, 842])
        page_height = page_size[1] if len(page_size) > 1 else 842
        blocks = page.get("para_blocks") or page.get("blocks") or []

        # 跨页处理：换页时重置 Y 坐标追踪
        if page_idx > 0:
            prev_block_bottom = None

        # 分页：每页第一个 block 是 title 时，在该页内容输出前插入分页符
        # 这样封面/目录/章节标题各自从新页开始，正文续页不分页
        if page_idx > 0 and blocks:
            first_btype = (blocks[0].get("type") or blocks[0].get("block_type") or "").lower()
            if first_btype == "title":
                _add_page_break(doc)

        for block in blocks:
            # 根据与上一个 block 的 Y 坐标差插入垂直留白
            bbox = block.get("bbox", [])
            if bbox and len(bbox) >= 4:
                y0 = bbox[1]
                if prev_block_bottom is not None:
                    gap = y0 - prev_block_bottom
                    # Y 差超过 50pt（约 1.8cm）时，插入空段落还原留白
                    # 50pt 阈值排除正常段间距（20-30pt），只还原明显的页面留白
                    if gap > 50:
                        _add_vertical_gap(doc, gap)
                prev_block_bottom = bbox[3]

            btype = (block.get("type") or block.get("block_type") or "text").lower()
            builder = _BLOCK_BUILDERS.get(btype, _build_text)

            try:
                if btype in ("image", "picture"):
                    builder(doc, block, images_dir)
                else:
                    builder(doc, block)
                block_count += 1
            except Exception as e:
                print(f"  ⚠️ block 重建失败 (page={page_idx}, type={btype}): {e}",
                      file=sys.stderr)

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))

    print(f"[4/4] DOCX 重建完成: {output_path}（{block_count} 个 block）",
          file=sys.stderr)
    return str(output_path.resolve())


def _add_vertical_gap(doc, gap_pt: float) -> None:
    """
    插入垂直留白（用空段落的段前距实现）。
    gap_pt: 需要的间距（PDF 点，1pt ≈ 0.35mm）。
    """
    from docx.shared import Pt
    # gap_pt 是 PDF 坐标的间距，直接转为磅值（1pt = 1磅）
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after = Pt(0)
    para.paragraph_format.line_spacing = Pt(1)  # 最小行高
    # 用段前距实现间距
    para.paragraph_format.space_before = Pt(gap_pt)


def _add_page_break(doc) -> None:
    """插入分页符。"""
    from docx.enum.text import WD_BREAK
    para = doc.add_paragraph()
    run = para.add_run()
    run.add_break(WD_BREAK.PAGE)


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
