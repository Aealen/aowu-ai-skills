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

import html
import re
import sys
from pathlib import Path
from typing import Any


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


def _detect_page_layout(pdf_info: list, margins: dict | None = None) -> None:
    """
    从每页的 block bbox 数据中检测页面布局参数，注入到每个 block 中。

    检测内容：
      - _page_x0: 正文左边距（优先用 PyMuPDF 精确检测的 margins，回退到 bbox 统计）
      - _page_width: 页宽（page_size[0]）
      - _page_center: 页面水平中心
    """
    from collections import Counter

    # 文档级左边距：优先用 PyMuPDF 精确检测的值，回退到 bbox 统计
    if margins and "left" in margins:
        doc_x0 = margins["left"]
    else:
        all_text_x0s = []
        for page in pdf_info:
            blocks = page.get("para_blocks") or page.get("blocks") or []
            for block in blocks:
                bbox = block.get("bbox")
                if not (bbox and len(bbox) >= 4):
                    continue
                btype = (block.get("type") or block.get("block_type") or "").lower()
                if btype not in ("text", "paragraph", "list"):
                    continue
                raw_text = _extract_text(block).strip()
                if len(raw_text) < 15:
                    continue
                all_text_x0s.append(bbox[0])
        doc_x0 = 82
        if all_text_x0s:
            all_text_x0s.sort()
            doc_x0 = round(all_text_x0s[len(all_text_x0s) // 4] / 3) * 3

    for page in pdf_info:
        page_size = page.get("page_size", [595, 842])
        page_width = page_size[0] if len(page_size) > 0 else 595
        page_center = page_width / 2

        blocks = page.get("para_blocks") or page.get("blocks") or []
        for block in blocks:
            block["_page_x0"] = doc_x0
            block["_page_width"] = page_width
            block["_page_center"] = page_center


def _apply_paragraph_format(para, block: dict[str, Any] | None = None) -> None:
    """
    根据 block 的 bbox 位置推断段落格式。
    所有参数均从 PDF 实际数据提取（_page_x0 / _page_width / _page_center），
    不硬编码任何值。
    """
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    pf = para.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)

    line_height = (block.get("_line_height") if block else None) or (block.get("_doc_line_height") if block else None)
    if line_height:
        pf.line_spacing = Pt(line_height)

    if not block:
        return

    page_x0 = block.get("_page_x0", 82)
    page_width = block.get("_page_width", 595)
    page_center = block.get("_page_center", page_width / 2)

    bbox = block.get("bbox", [])
    lines = block.get("lines", [])

    if not bbox or not lines:
        return

    first_spans = lines[0].get("spans", [])
    if not first_spans:
        return

    first_x0 = first_spans[0].get("bbox", [0])[0]

    # 居中判断：检查各行中心是否接近页面中心
    # 多行 block 中，第一行可能占满全宽（拉宽 block bbox），但后续行是居中的
    # 因此看"多数行的中心是否接近页面中心"，而非 block 整体 bbox
    # 注意：全宽行（占满内容区）的中心天然接近页面中心，但不是居中
    # 只有"行宽 < 内容区宽度且中心接近页面中心"的行才算居中行
    # 额外排除：左对齐文本恰好中心接近页面中心的情况——
    # 居中行的左侧应有明显留白（x0 远离左边距），左对齐的 x0 接近左边距
    content_width = page_width - page_x0 * 2
    centered_lines = 0
    total_lines = 0
    for line in lines:
        spans = line.get("spans", [])
        if not spans:
            continue
        line_bbox = line.get("bbox") or spans[0].get("bbox", [])
        if len(line_bbox) >= 4:
            line_center = (line_bbox[0] + line_bbox[2]) / 2
            line_width = line_bbox[2] - line_bbox[0]
            # 居中行的期望 x0 = 页面中心 - 行宽/2
            expected_x0 = page_center - line_width / 2
            total_lines += 1
            # 居中行：中心接近页面中心 + 行宽未占满 + 实际 x0 接近期望 x0
            if (abs(line_center - page_center) < page_width * 0.05 and
                    line_width < content_width * 0.90 and
                    abs(line_bbox[0] - expected_x0) < page_width * 0.05):
                centered_lines += 1

    if total_lines > 0 and centered_lines >= total_lines / 2:
        is_centered = True
    elif bbox and len(bbox) >= 4:
        # 单行 block 回退到 bbox 对称性判断
        left_gap = bbox[0] - page_x0
        right_gap = page_width - page_x0 - bbox[2]
        gap_diff = abs(left_gap - right_gap)
        is_centered = (gap_diff < page_width * 0.05 and
                       left_gap > page_width * 0.02 and
                       right_gap > page_width * 0.02)
    else:
        is_centered = False

    if is_centered:
        pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pf.first_line_indent = Pt(0)
    else:
        pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        # 缩进：分两层
        # left_indent = 续行 x0（所有行的最小 x0）相对左边距的偏移
        # first_line_indent = 首行 x0 相对续行 x0 的额外偏移（首行缩进）
        min_x0 = first_x0
        for line in lines:
            spans = line.get("spans", [])
            if spans:
                lx0 = spans[0].get("bbox", [min_x0])[0]
                if lx0 < min_x0:
                    min_x0 = lx0
        # left_indent：续行的起始位置
        left_indent_pt = max(0, min_x0 - page_x0)
        if left_indent_pt > 6:
            pf.left_indent = Pt(round(left_indent_pt))
        # first_line_indent：首行相对续行的额外缩进
        first_indent_pt = max(0, first_x0 - min_x0)
        if first_indent_pt > 6:
            pf.first_line_indent = Pt(round(first_indent_pt))
        else:
            pf.first_line_indent = Pt(0)


def _apply_style(run, style: dict[str, Any] | None = None) -> None:
    """
    把 PyMuPDF 样式应用到 python-docx run。
    style 格式：{"font": str, "size": float, "bold": bool, "italic": bool, "color": int}

    ⚠️ CID 字体加粗检测：PyMuPDF 的 font flags 对某些中文字体（如 SimHei）
    无法准确报告加粗状态（字形轮廓被加粗但 flags 未变）。这里用字体名做补充判断。
    """
    if not style:
        return

    from docx.shared import Pt, RGBColor

    size = style.get("size")
    if size:
        run.font.size = Pt(size)

    # 加粗：完全依赖提取层的判断（flags + 像素密度检测），不再做字体名猜测
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

        font_name = font  # 保留 PDF 原始字体名，由 Word/LibreOffice 自动匹配系统字体
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

    # 标题层级：优先用 PDF 大纲（_toc_level），这是作者在 Word 里设的权威层级
    # 无大纲数据时回退到 MinerU text_level
    level = block.get("_toc_level")
    is_structural = level is not None  # 是否是真·结构标题（在 PDF 大纲中）
    if level is None:
        level = block.get("text_level") or block.get("level") or 1
    try:
        level = int(level)
    except (TypeError, ValueError):
        level = 1
    level = min(max(level, 1), 9)

    # 建段落：结构标题用 Heading 样式 + TOC 绑定，视觉标题用普通段落
    # （视觉标题保留加粗/字号等视觉强调，但不污染 DOCX 目录）
    if is_structural:
        heading_style_name = f"Heading {level}"
        try:
            para = doc.add_paragraph(style=heading_style_name)
        except KeyError:
            para = doc.add_paragraph()
    else:
        para = doc.add_paragraph()

    # 段落格式：所有标题都用 _apply_title_format（LEFT/CENTER，不用 JUSTIFY）
    # _apply_title_format 已支持从 bbox 检测缩进，视觉标题也能正确缩进
    _apply_title_format(para, level, block)

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

    # 只有结构标题才设 OutlineLevel（绑定到 DOCX TOC）
    if is_structural:
        _set_outline_level(para, level)

    # 强制标题颜色为黑色（覆盖 Heading 样式默认蓝色）
    _force_heading_color_black(para)


def _apply_title_format(para, level: int, block: dict | None = None) -> None:
    """
    标题段落格式。所有参数从 PDF 实际数据提取。
    
    对齐：根据 bbox 位置判断（居中 / 左对齐），不再按 level 写死
    缩进：根据 bbox 实际 x0 计算，不再写死为 0
    行距：优先用 PDF 测量值，无数据时单倍
    段后间距：根据同页下一个 block 的 Y 间距计算
    """
    from docx.shared import Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING

    pf = para.paragraph_format
    pf.space_before = Pt(0)

    # 行距
    line_height = (block.get("_line_height") if block else None) or (block.get("_doc_line_height") if block else None)
    if line_height:
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.line_spacing = Pt(line_height)
    else:
        pf.line_spacing_rule = WD_LINE_SPACING.SINGLE

    if not block:
        return

    # 对齐和缩进：从 bbox 位置推断
    page_x0 = block.get("_page_x0", 82)
    page_width = block.get("_page_width", 595)
    page_center = block.get("_page_center", page_width / 2)

    bbox = block.get("bbox", [])
    lines = block.get("lines", [])
    first_x0 = page_x0

    if bbox and lines:
        first_spans = lines[0].get("spans", [])
        if first_spans:
            first_x0 = first_spans[0].get("bbox", [bbox[0]])[0]

    # 居中检测：看多数行的中心是否接近页面中心（排除全宽行）
    content_width = page_width - page_x0 * 2
    centered_lines = 0
    total_lines = 0
    for line in lines:
        spans = line.get("spans", [])
        if not spans:
            continue
        line_bbox = line.get("bbox") or spans[0].get("bbox", [])
        if len(line_bbox) >= 4:
            line_center = (line_bbox[0] + line_bbox[2]) / 2
            line_width = line_bbox[2] - line_bbox[0]
            expected_x0 = page_center - line_width / 2
            total_lines += 1
            if (abs(line_center - page_center) < page_width * 0.05 and
                    line_width < content_width * 0.90 and
                    abs(line_bbox[0] - expected_x0) < page_width * 0.05):
                centered_lines += 1

    if total_lines > 0 and centered_lines >= total_lines / 2:
        is_centered = True
    elif bbox and len(bbox) >= 4:
        left_gap = bbox[0] - page_x0
        right_gap = page_width - page_x0 - bbox[2]
        gap_diff = abs(left_gap - right_gap)
        is_centered = (gap_diff < page_width * 0.05 and
                       left_gap > page_width * 0.02 and
                       right_gap > page_width * 0.02)
    else:
        is_centered = False

    if is_centered:
        pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
        # 标题缩进：用 left_indent（统一所有行），取首行 x0 差值
        # 但如果首行文字已接近内容区宽度（>85%），不设缩进避免溢出换行
        indent_pt = max(0, first_x0 - page_x0)
        content_width = page_width - page_x0 * 2  # 粗估内容宽度
        if indent_pt > 6 and bbox and len(bbox) >= 4:
            block_width = bbox[2] - bbox[0]
            if block_width < content_width * 0.85:
                pf.left_indent = Pt(round(indent_pt))
    pf.first_line_indent = Pt(0)

    # 段后间距：从同页下一个 block 的 Y 坐标差获取，而非硬编码
    gap_after = block.get("_gap_after")
    if gap_after is not None and gap_after > 0:
        pf.space_after = Pt(round(gap_after))
    else:
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


def _find_matching_pymupdf_table(
    fitz_doc, page_idx: int, table_bbox: list[float],
    iou_threshold: float = 0.3,
):
    """
    在指定页面用 IoU 匹配 MinerU table bbox 对应的 PyMuPDF 表格。
    返回匹配的 fitz.Table 对象，匹配失败返回 None。
    """
    page = fitz_doc[page_idx]
    tabs = page.find_tables()
    target_x0, target_y0, target_x1, target_y1 = table_bbox

    best_table = None
    best_iou = 0.0
    for t in tabs.tables:
        tx0, ty0, tx1, ty1 = t.bbox
        ix0 = max(target_x0, tx0)
        iy0 = max(target_y0, ty0)
        ix1 = min(target_x1, tx1)
        iy1 = min(target_y1, ty1)
        if ix0 < ix1 and iy0 < iy1:
            inter = (ix1 - ix0) * (iy1 - iy0)
            area_target = (target_x1 - target_x0) * (target_y1 - target_y0)
            area_detected = (tx1 - tx0) * (ty1 - ty0)
            union = area_target + area_detected - inter
            iou = inter / union if union > 0 else 0
            if iou > best_iou:
                best_iou = iou
                best_table = t

    if best_table is None or best_iou < iou_threshold:
        return None
    return best_table


def _extract_table_col_widths_from_table(pymupdf_table) -> list[int] | None:
    """
    从 PyMuPDF Table 对象提取列宽（pt）。
    收集所有 cell 的 x 边界，排序去重 → 列宽。
    """
    xs: set[int] = set()
    for cell in pymupdf_table.cells:
        if cell:
            xs.add(round(cell[0]))
            xs.add(round(cell[2]))
    xs_sorted = sorted(xs)
    if len(xs_sorted) < 2:
        return None
    return [xs_sorted[i + 1] - xs_sorted[i] for i in range(len(xs_sorted) - 1)]


def _extract_table_cell_texts(pymupdf_table) -> list[list[str]] | None:
    """
    从 PyMuPDF Table 对象提取单元格文本矩阵（保留换行 \\n）。

    PyMuPDF 的 extract() 返回 list[list[str]]，每个 cell 的文本
    保留了原文的换行（\\n）。这是恢复表格内换行格式的关键数据源——
    MinerU 的 HTML 把单元格内多行内容合并成了无换行的扁平字符串。

    返回 None 表示提取失败。
    """
    try:
        return pymupdf_table.extract()
    except Exception:
        return None


def _measure_block_line_height(
    fitz_doc, page_idx: int, bbox: list[float],
) -> float | None:
    """
    测量 block 区域内文字行的实际行高（y0 差值中位数），单位 pt。

    通用函数，适用于所有 block 类型（text/title/index/table/list）。
    从 PDF 直接提取每个 block 的真实行高，而非用统一默认值。
    这样不同段落、不同表格各自还原为原 PDF 的真实行距。

    方法：在 block bbox 内收集所有文字行的 y0，按列分组后
    计算同列连续行的 y0 差值，取中位数作为行高。

    返回 None 表示提取失败（调用方回退到默认行距）。
    """
    page = fitz_doc[page_idx]
    d = page.get_text("dict")
    x0_min, y0_min, x1_max, y1_max = bbox

    # 先收集 block 内所有 span 的字号，取众数作为主导字号
    # 用于动态计算行高过滤范围（而非写死 12-45pt）
    size_counter: dict[float, int] = {}
    for block in d.get("blocks", []):
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            lx0 = line["bbox"][0]
            ly0 = line["bbox"][1]
            if not (x0_min - 5 <= lx0 <= x1_max + 5 and
                    y0_min - 5 <= ly0 <= y1_max + 5):
                continue
            for sp in line.get("spans", []):
                if sp["text"].strip():
                    sz = round(sp["size"], 1)
                    size_counter[sz] = size_counter.get(sz, 0) + 1
    dominant_size = max(size_counter, key=size_counter.get) if size_counter else 12.0
    # 行高合理范围：字号 × [0.8, 3.5]（覆盖单倍到多倍行距）
    gap_min = dominant_size * 0.8
    gap_max = dominant_size * 3.5

    # 收集 block bbox 内所有文字行
    # 按 x0 粗分列（同一列的文字 x0 接近），便于只算同列内连续行距
    col_y0s: dict[int, list[float]] = {}  # 列桶 → y0 列表
    for block in d.get("blocks", []):
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            spans = [sp for sp in line.get("spans", []) if sp["text"].strip()]
            if not spans:
                continue
            ly0 = line["bbox"][1]
            lx0 = line["bbox"][0]
            # 限定在 block bbox 内
            if not (x0_min - 5 <= lx0 <= x1_max + 5 and
                    y0_min - 5 <= ly0 <= y1_max + 5):
                continue
            text = "".join(sp["text"] for sp in spans)
            if len(text) < 3:
                continue  # 跳过短文本（条款号等），避免干扰行高
            # 列桶：x0 每 20pt 一个桶（同一列文字 x0 差异通常 < 20pt）
            col_bucket = int(lx0 // 20)
            col_y0s.setdefault(col_bucket, []).append(ly0)

    # 收集所有列内的连续 y0 差值
    all_gaps: list[float] = []
    for y0_list in col_y0s.values():
        y0_list.sort()
        for i in range(1, len(y0_list)):
            gap = y0_list[i] - y0_list[i - 1]
            if gap_min <= gap <= gap_max:
                all_gaps.append(gap)

    if len(all_gaps) < 2:
        return None

    all_gaps.sort()
    return all_gaps[len(all_gaps) // 2]  # 中位数


def _match_blocks_with_toc(fitz_doc, pdf_info: list) -> int:
    """
    用 PDF 大纲（书签/TOC）确定哪些 block 是真正的标题。

    PDF 大纲是作者在 Word 里设置的大纲层级（通过 get_toc 获取），
    包含每个标题条目的层级、文本和精确坐标（页码 + y 位置）。
    这是文档结构的权威定义，比 MinerU 的 type 判断可靠得多。

    对每个大纲条目，按坐标匹配 MinerU block，注入 block["_toc_level"]。
    返回成功匹配的大纲条目数。

    下游效果：
      - 有 _toc_level 的 block → 真标题，用大纲层级
      - 无 _toc_level 但 MinerU 标记为 title → 降级为正文（MinerU 误判）
    """
    toc = fitz_doc.get_toc(simple=False)
    if not toc:
        return 0

    # 构建大纲查找表：{(page_idx, y_bucket): level}
    # y_bucket 精度 10pt（容差匹配）
    # 注意：PyMuPDF get_toc 的 to.y 已经是 top-down 坐标（与 MinerU bbox 一致），
    # 不需要做 PDF→top-down 转换
    toc_lookup: dict[tuple[int, int], int] = {}
    for entry in toc:
        level = entry[0]
        dest = entry[3] if len(entry) > 3 else {}
        # entry[2] 是 1-based 页码，dest["page"] 是 0-based 页码
        pi = dest.get("page", entry[2] - 1)
        to = dest.get("to")
        if to is None:
            continue
        y_bucket = int(to.y // 10)
        toc_lookup[(pi, y_bucket)] = level
        # 相邻 bucket 也设上（容差 ±10pt）
        toc_lookup[(pi, y_bucket + 1)] = level

    matched = 0
    for page in pdf_info:
        pi = page.get("page_idx", 0)
        page_blocks = page.get("para_blocks") or page.get("blocks") or []
        for block in page_blocks:
            bbox = block.get("bbox")
            if not (bbox and len(bbox) >= 4):
                continue
            # 用 block 的 y0 匹配大纲
            y_bucket = int(bbox[1] // 10)
            level = toc_lookup.get((pi, y_bucket))
            if level is not None:
                block["_toc_level"] = level
                matched += 1

    return matched


def _extract_table_html(block: dict[str, Any]) -> str:
    """
    从 MinerU table block 提取 HTML。
    MinerU 3.x 的 table HTML 不在 table_body 字段里，
    而是在 blocks[].lines[].spans[].html 里。
    """
    # 路径1：顶层 table_body（部分版本可能有）
    table_body = block.get("table_body") or ""
    if table_body:
        return table_body
    # 路径2：blocks[].lines[].spans[].html（MinerU 3.x 实际位置）
    for sb in block.get("blocks", []):
        for line in sb.get("lines", []):
            for sp in line.get("spans", []):
                h = sp.get("html", "")
                if h:
                    return h
    return ""


def _set_table_fixed_width(table, col_widths_pt: list[float]) -> None:
    """
    设置表格为固定列宽布局，直接操作 OOXML。

    python-docx 的 table.autofit=False + cell.width 只设了 tblLayout 和 tcW，
    但 **tblGrid 的 gridCol 仍是创建时的等宽值**，且 tblW 是 auto。
    LibreOffice 在 fixed 布局下优先用 tblGrid 决定列宽，导致 cell.width 被忽略。

    本函数补全三处 OOXML：
      1. tblLayout type=fixed（关闭自动调整）
      2. tblGrid gridCol（列宽定义，渲染器实际依据）
      3. tblW type=dxa（表格总宽度=各列宽之和）
      4. 每个 cell 的 tcW（逐 cell 宽度，与 gridCol 一致）
    """
    from docx.oxml.ns import qn
    from docx.shared import Pt
    from docx.oxml import OxmlElement

    tbl = table._tbl

    # 1. tblLayout type=fixed
    tblPr = tbl.tblPr
    tblLayout = tblPr.find(qn('w:tblLayout'))
    if tblLayout is None:
        tblLayout = OxmlElement('w:tblLayout')
        tblPr.append(tblLayout)
    tblLayout.set(qn('w:type'), 'fixed')

    # 2. 更新 tblGrid 的 gridCol（列宽定义）
    # pt → twips(dxa)：1pt = 20 twips
    col_widths_dxa = [round(w * 20) for w in col_widths_pt]
    total_dxa = sum(col_widths_dxa)

    tblGrid = tbl.find(qn('w:tblGrid'))
    if tblGrid is not None:
        # 清除旧 gridCol，按新列宽重建
        for gc in tblGrid.findall(qn('w:gridCol')):
            tblGrid.remove(gc)
        for w in col_widths_dxa:
            gridCol = OxmlElement('w:gridCol')
            gridCol.set(qn('w:w'), str(w))
            tblGrid.append(gridCol)

    # 3. tblW type=dxa（表格总宽度固定）
    tblW = tblPr.find(qn('w:tblW'))
    if tblW is None:
        tblW = OxmlElement('w:tblW')
        tblPr.append(tblW)
    tblW.set(qn('w:type'), 'dxa')
    tblW.set(qn('w:w'), str(total_dxa))

    # 4. 逐 cell 设 tcW（与 gridCol 一致）
    for row in table.rows:
        for ci, cell in enumerate(row.cells):
            if ci < len(col_widths_pt):
                cell.width = Pt(col_widths_pt[ci])


def _normalize_text(s: str) -> str:
    """去除空白用于文本匹配（空格/换行差异不影响内容比对）。"""
    return re.sub(r"\s+", "", s or "")


def _enrich_rows_with_pymupdf_text(
    rows: list[list[dict[str, Any]]],
    pymupdf_cells: list[list[str]],
) -> None:
    """
    用 PyMuPDF extract() 的带换行文本，替换 HTML rows 里的扁平文本。

    MinerU HTML 把单元格内多行内容合并成无换行的连续字符串，
    PyMuPDF extract() 保留了原文换行（\\n）。

    两种策略：
      1. 行列完全一致 → 按位置直接替换
      2. 行列不一致（如跨页表格 PyMuPDF 只检测到部分行）→
         按"去空白文本前缀匹配"逐 cell 查找替换，匹配不到的保留 HTML 原文

    匹配规则：HTML cell 去空白文本 == PyMuPDF cell 去空白文本的前 N 字符
    （双向取较短者比较），避免空格/换行差异导致漏匹配。
    """
    n_rows_html = len(rows)
    n_rows_pm = len(pymupdf_cells)
    n_cols_pm = len(pymupdf_cells[0]) if pymupdf_cells else 0

    # 策略1：行列完全一致，按位置替换
    if n_rows_pm == n_rows_html:
        # HTML 的列数可能因 colspan 变化，逐行检查
        all_match = True
        for ri, row in enumerate(rows):
            html_cols = sum(c.get("colspan", 1) for c in row)
            if html_cols != n_cols_pm:
                all_match = False
                break
        if all_match:
            for ri, row in enumerate(rows):
                for ci, cell_data in enumerate(row):
                    pm_text = pymupdf_cells[ri][ci] if ci < n_cols_pm else None
                    if pm_text and pm_text.strip():
                        cell_data["text"] = pm_text.strip()
            return

    # 策略2：行列不一致，按文本相似度匹配
    # PyMuPDF 和 MinerU 对同一文字的提取可能有字符差异（如"）"vs"〕"、"繫"vs"系"），
    # 用 difflib.SequenceMatcher 计算整体相似度，容忍少量 OCR 差异。
    import difflib

    pm_texts: list[tuple[str, str]] = []  # (去空白文本, 原始带换行文本)
    for pm_row in pymupdf_cells:
        for pm_text in pm_row:
            if pm_text and pm_text.strip():
                norm = _normalize_text(pm_text)
                if norm:
                    pm_texts.append((norm, pm_text.strip()))

    for row in rows:
        for cell_data in row:
            html_text = cell_data.get("text", "")
            if not html_text:
                continue
            html_norm = _normalize_text(html_text)
            if len(html_norm) < 8:
                continue  # 短文本（如条款号"1.1"）不做模糊匹配，避免误匹配

            best_match = None
            best_ratio = 0.0
            for pm_norm, pm_orig in pm_texts:
                if html_norm == pm_norm:
                    best_match = pm_orig
                    break
                # 整体相似度（容忍 OCR 字符差异）
                ratio = difflib.SequenceMatcher(None, html_norm, pm_norm).ratio()
                # 动态阈值：文本越长容忍度越高
                # ≥30字符用0.8，10-30字符用0.9，避免短文本误匹配
                threshold = 0.8 if len(html_norm) >= 30 else 0.9
                if ratio > threshold and ratio > best_ratio:
                    best_ratio = ratio
                    best_match = pm_orig
            if best_match:
                cell_data["text"] = best_match


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

    # ── 设置固定列宽（匹配原文 PDF）──
    # MinerU HTML 不含列宽，_col_widths 由 build_docx 预扫描用 PyMuPDF
    # find_tables 提取后注入。无此字段时保持默认等宽。
    col_widths = block.get("_col_widths")
    if col_widths and len(col_widths) == max_cols:
        _set_table_fixed_width(table, col_widths)
    else:
        # 列宽数与 max_cols 不匹配（可能 PyMuPDF 检测的列数与 HTML 不一致）
        if col_widths and len(col_widths) != max_cols:
            print(f"  ⚠️ 表格列宽不匹配: PyMuPDF={len(col_widths)}列 vs HTML={max_cols}列，回退等宽",
                  file=sys.stderr)
        table.autofit = True

    # ── 恢复单元格内换行（PyMuPDF extract() 保留 \n）──
    # MinerU HTML 把单元格内多行内容合并成无换行的扁平字符串，
    # 用 PyMuPDF 的 extract() 文本恢复换行格式。
    pymupdf_cells = block.get("_pymupdf_cells")
    if pymupdf_cells:
        _enrich_rows_with_pymupdf_text(rows, pymupdf_cells)

    # ── 修复跨页表格的单元格错位（在 PyMuPDF 文本增强之后）──
    # 用 PyMuPDF 数据作为基准：如果 PyMuPDF 显示 col0 为空但 HTML 有文本 → 错位
    _fix_crosspage_cell_misplacement(rows, pymupdf_cells)

    _fill_table_cells(table, rows, block.get("_style"),
                      block.get("_line_height") or block.get("_doc_line_height"))


def _parse_html_table(table_html: str) -> list[list[dict[str, Any]]]:
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

    for tr_match in tr_pattern.finditer(table_html):
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
            cell_text = html.unescape(cell_text)  # 反转义所有 HTML 实体
            row.append({"text": cell_text, "rowspan": rowspan, "colspan": colspan})
        if row:
            rows.append(row)

    return rows


def _fix_crosspage_cell_misplacement(
    rows: list[list[dict[str, Any]]],
    pymupdf_cells: list[list[str]] | None = None,
) -> None:
    """
    修复 MinerU 跨页表格的单元格错位。

    问题：跨页表格在换页处，MinerU 常把内容列的文字错放到条款号列（col0）。
    例如条款号列出现 "商自行承担。应商自行承担。可能导致投标..."（整句正文）。

    修复逻辑（两层）：
      1. PyMuPDF 数据覆盖的行：对比 col0，PyMuPDF 显示为空但 HTML 有文本 → 错位
      2. PyMuPDF 数据未覆盖的行（行数不匹配）：判断 col0 文本长度
         —— 条款号通常 ≤10 字符，col0 超过 20 字符几乎一定是正文错位
    """
    if not pymupdf_cells:
        return  # 无 PyMuPDF 数据时不做修正

    for ri, row in enumerate(rows):
        if len(row) < 2:
            continue
        col0_text = row[0].get("text", "").strip()
        if not col0_text:
            continue

        is_misplaced = False

        if ri < len(pymupdf_cells):
            # 策略 1：PyMuPDF 数据覆盖到的行，对比 col0
            pm_col0 = ""
            if len(pymupdf_cells[ri]) > 0:
                pm_col0 = (pymupdf_cells[ri][0] or "").strip()
            if not pm_col0 and col0_text:
                is_misplaced = True
        else:
            # 策略 2：PyMuPDF 未覆盖的行，用文本特征判断
            # 条款号特征：短（≤10字符）、无中文标点、通常是数字+点/顿号格式
            # 正文特征：长（>20字符）或含中文标点（句号、逗号等）
            import re as _re
            has_cn_punct = bool(_re.search(r'[。，；！？、：]', col0_text))
            if len(col0_text) > 20 or has_cn_punct:
                is_misplaced = True

        if is_misplaced:
            col1_text = row[1].get("text", "").strip()
            if col1_text:
                row[1]["text"] = col0_text + "\n" + col1_text
            else:
                row[1]["text"] = col0_text
            row[0]["text"] = ""


def _fill_table_cells(table, rows: list[list[dict[str, Any]]],
                      style: dict[str, Any] | None = None,
                      line_height: float | None = None) -> None:
    """
    填充表格单元格，处理 rowspan/colspan 合并，并应用样式。
    style 是 block._style（align.py 贴的兜底样式），应用到每个单元格的文字。
    line_height 是从 PDF 测量的表格内文字行高（pt），无则回退默认。
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
                # 设置段落行距：从 PDF 测量的真实行高，无数据时回退 22pt
                # python-docx 默认继承 Word 的 1.15 倍多倍行距（约17pt），
                # 比源 PDF 表格紧凑，需用固定行距匹配原文。
                from docx.shared import Pt as _Pt
                from docx.enum.text import WD_LINE_SPACING
                pf = para.paragraph_format
                pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
                pf.line_spacing = _Pt(line_height or 22)  # 22 = 全文档无行高数据时的最后兜底
                pf.space_before = _Pt(0)
                pf.space_after = _Pt(0)
                # 支持单元格内换行：文本含 \n 时，用 run.add_break() 渲染
                if "\n" in text:
                    lines = text.split("\n")
                    run = para.add_run()
                    for li, line in enumerate(lines):
                        if li > 0:
                            run.add_break()  # <w:br/> 单元格内换行
                        if line:
                            run.add_text(line)
                else:
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
    图片宽度从 block 的 bbox 提取（pt→inch），保持与 PDF 原始尺寸一致。
    超出页面内容宽度时等比缩放。
    """
    from docx.shared import Inches

    img_path = block.get("img_path") or block.get("image_path") or ""
    if not img_path or not images_dir:
        return

    full_path = Path(images_dir) / Path(img_path).name
    if not full_path.exists():
        print(f"  ⚠️ 图片不存在: {full_path}", file=sys.stderr)
        return

    # 从 bbox 提取图片宽度（pt），转换为英寸（1 inch = 72pt）
    bbox = block.get("bbox", [])
    width_in = None
    if bbox and len(bbox) >= 4:
        width_pt = bbox[2] - bbox[0]
        if width_pt > 10:  # 过滤异常值
            width_in = width_pt / 72.0

    # 上限保护：不超过页面内容宽度
    page_width = block.get("_page_width", 595)
    page_x0 = block.get("_page_x0", 72)
    content_width_in = (page_width - page_x0 * 2) / 72.0  # 粗估内容宽度
    if width_in is None:
        width_in = content_width_in  # 无 bbox 时用内容宽度
    elif width_in > content_width_in:
        width_in = content_width_in  # 超出时截断到内容宽度

    doc.add_picture(str(full_path), width=Inches(width_in))


def _build_list(doc, block: dict[str, Any]) -> None:
    """
    列表：MVP 降级为普通段落（带编号文本）。
    TODO(完善): 根据 MinerU 的 list sub_type 还原编号/项目符号格式。
    """
    line_height = block.get("_line_height") or block.get("_doc_line_height")
    for line in block.get("lines", []):
        text = "".join(
            s.get("content") or s.get("text") or ""
            for s in line.get("spans", [])
        )
        if text.strip():
            try:
                para = doc.add_paragraph(text, style="List Bullet")
            except KeyError:
                para = doc.add_paragraph(text)
            # 行距：从 PDF 测量的真实行高
            if line_height:
                from docx.shared import Pt
                from docx.enum.text import WD_LINE_SPACING
                pf = para.paragraph_format
                pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
                pf.line_spacing = Pt(line_height)


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
        pf.line_spacing = Pt(block.get("_line_height") or block.get("_doc_line_height", 22))
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

        # 边距：直接用 PyMuPDF 从 PDF 提取的真实边距，不做任何写死裁剪
        margins = _detect_page_margins(pdf_path, pdf_info) if pdf_path else None

        if margins:
            sec.left_margin = Pt(margins["left"])
            sec.right_margin = Pt(margins["right"])
            sec.top_margin = Pt(margins["top"])
            sec.bottom_margin = Pt(margins["bottom"])
        else:
            # 无 pdf_path 时从 MinerU block bbox 推断
            first_blocks = first_page.get("para_blocks") or first_page.get("blocks") or []
            if first_blocks:
                all_x0 = [b["bbox"][0] for b in first_blocks if b.get("bbox")]
                all_x1 = [b["bbox"][2] for b in first_blocks if b.get("bbox")]
                if all_x0 and all_x1:
                    sec.left_margin = Pt(min(all_x0))
                    sec.right_margin = Pt(pdf_w_pt - max(all_x1))
            if not sec.top_margin:
                sec.top_margin = Pt(72)  # 1 inch fallback
            if not sec.bottom_margin:
                sec.bottom_margin = Pt(72)

    # ── 预扫描：从 PDF 提取排版信息（行高、列宽、单元格换行）──
    # 对所有 block 类型提取真实行高（block["_line_height"]），
    # 使后续不同 PDF 都能各自还原真实行距，而非用统一默认值。
    # table 类型额外提取列宽和带换行的单元格文本。
    if pdf_path:
        try:
            import fitz
            fdoc = fitz.open(pdf_path)
            line_h_found = 0
            widths_found = 0
            texts_found = 0
            for page in pdf_info:
                pi = page.get("page_idx", 0)
                if pi >= fdoc.page_count:
                    continue
                page_blocks = page.get("para_blocks") or page.get("blocks") or []
                for block in page_blocks:
                    btype = (block.get("type") or block.get("block_type") or "").lower()
                    bbox = block.get("bbox")
                    if not (bbox and len(bbox) >= 4):
                        continue

                    # 所有类型：提取真实行高
                    line_h = _measure_block_line_height(fdoc, pi, bbox)
                    if line_h:
                        block["_line_height"] = line_h
                        line_h_found += 1

                    # table 类型：额外提取列宽和单元格文本
                    if btype == "table":
                        mtable = _find_matching_pymupdf_table(fdoc, pi, bbox)
                        if mtable is None:
                            continue
                        widths = _extract_table_col_widths_from_table(mtable)
                        if widths:
                            block["_col_widths"] = widths
                            widths_found += 1
                        cell_texts = _extract_table_cell_texts(mtable)
                        if cell_texts:
                            block["_pymupdf_cells"] = cell_texts
                            texts_found += 1

            # ── 用 PDF 大纲（书签）确定标题层级 ──
            # PDF 大纲是作者在 Word 里设置的大纲层级，是文档结构的权威定义。
            # MinerU 的 type=title 判断经常误判（如把"1.采购人信息"判为标题），
            # 用大纲坐标匹配来校验：只有在大纲里的 block 才是真标题。
            toc_matched = _match_blocks_with_toc(fdoc, pdf_info)
            fdoc.close()
            if line_h_found or widths_found or texts_found or toc_matched:
                print(f"  📐 排版信息提取: {line_h_found}个行高, "
                      f"{widths_found}个列宽, {texts_found}个单元格文本, "
                      f"{toc_matched}个大纲标题",
                      file=sys.stderr)
        except Exception as e:
            print(f"  ⚠️ 预扫描失败（回退默认）: {e}", file=sys.stderr)

    # ── 页面布局检测：从每页 block bbox 提取左边距、页宽等 ──
    # 替代所有硬编码值（82pt 左边距、595pt 页宽），支持不同版式的 PDF
    _detect_page_layout(pdf_info, margins)

    # ── 文档级行高中位数（兜底值，替代硬编码的 24pt）──
    all_lh = []
    for page in pdf_info:
        for b in (page.get("para_blocks") or page.get("blocks") or []):
            lh = b.get("_line_height")
            if lh:
                all_lh.append(lh)
    doc_median_lh = 22.0  # 极端兜底
    if all_lh:
        all_lh.sort()
        doc_median_lh = all_lh[len(all_lh) // 2]
    for page in pdf_info:
        for b in (page.get("para_blocks") or page.get("blocks") or []):
            b["_doc_line_height"] = doc_median_lh

    # ── 段后间距：从同页相邻 block 的 Y 坐标差计算 ──
    for page in pdf_info:
        blocks = page.get("para_blocks") or page.get("blocks") or []
        for bi, block in enumerate(blocks):
            if bi + 1 < len(blocks):
                this_bbox = block.get("bbox", [])
                next_bbox = blocks[bi + 1].get("bbox", [])
                if this_bbox and next_bbox and len(this_bbox) >= 4 and len(next_bbox) >= 4:
                    gap = next_bbox[1] - this_bbox[3]  # 下一个 block 的 y0 - 当前 block 的 y1
                    if gap > 0:
                        block["_gap_after"] = gap

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

        # 分页：每页第一个 block 是标题时插分页符
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
                    # Y 差超过页高 6% 时，插入空段落还原留白
                    # 比例阈值替代硬编码 50pt，适应不同页面尺寸的 PDF
                    gap_threshold = page_height * 0.06
                    if gap > gap_threshold:
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
