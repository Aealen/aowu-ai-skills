#!/usr/bin/env python3
"""
parse_pymupdf.py —— PyMuPDF 字符级样式提取

职责：提取 PDF 每个 span 的字符级样式（font/size/color/flags + bbox）。
PyMuPDF 给字号字体颜色，但给不了版面语义（不知道哪段是标题）。
和 MinerU 配合：MinerU 给结构，PyMuPDF 给样式。

为什么选 PyMuPDF 不选 PDFBox：
  - PyMuPDF 直接给 color 字段（sRGB 整数），无需自建颜色状态机
  - flags 位掩码直接判断粗体/斜体
  - 底层轻量，提取速度快

依赖：PyMuPDF（fitz）

用法（被 pdf2docx.py 调用）:
    from parse_pymupdf import extract_spans, save_spans
    spans = extract_spans("input.pdf")
    save_spans(spans, "output/_spans.json")

PyMuPDF span 字段说明（page.get_text("dict") 的 span）:
    {
        "text":   str,       # 文本内容
        "font":   str,       # 字体名，如 "SimSun"、"Helvetica-Bold"
        "size":   float,     # 字号（磅）
        "color":  int,       # sRGB 整数，0=黑，16777215=白
        "flags":  int,       # 位掩码：bit4(16)=粗体, bit1(2)=斜体
        "bbox":   (x0,y0,x1,y1),  # PDF 点坐标，top-left 原点，y 向下
        "origin": (x, y),    # 文本基线原点
    }

    span 定义（官方）："adjacent characters with identical font properties:
    name, size, flags, and color" —— 相同样式的连续字符聚成一个 span。

flags 位掩码（PyMuPDF Appendix 1）:
    bit0(1)  = 上标 superscript
    bit1(2)  = 斜体 italic
    bit2(4)  = 衬线 serifed
    bit3(8)  = 等宽 monospaced
    bit4(16) = 粗体 bold
"""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any


# ═══════════════════════════════════════════════════════════════
#  核心提取函数
# ═══════════════════════════════════════════════════════════════
#  粗体判断（flags 位标记 + 字体名推断）
# ═══════════════════════════════════════════════════════════════

# 字体名明确标识为粗体变体的（字体名里含 bold/heavy/black）。
# 仅用于非 CID 字体（如 TimesNewRomanPS-BoldMT、Arial-Bold）的常规检测。
# CID 字体（中文）的粗体判断由像素密度检测（_check_cid_bold_by_density）处理，
# 不在此列表中维护字体名白名单。
_BOLD_FONT_PATTERNS = (
    "bold",     # TimesNewRomanPS-BoldMT, Arial-Bold 等
    "heavy",
    "black",
)

# CID 字体：flags 可靠度低，需要像素密度辅助判断粗体
# 只要字体名属于 CJK 字体范围，就纳入密度检测。不再手工维护白名单。
# 判断方式：字体名含已知 CJK 字体族关键词
_CJK_FONT_KEYWORDS = (
    "song", "hei", "kai", "fang", "yahei", "ming", "gothic", "mincho",
    "batang", "gulim", "dotum", "simsun", "simhei", "simkai", "simfang",
    "stsong", "stheiti", "stkaiti", "stfangsong", "fz", "ms",
    "source", "noto",  # 思源/Noto 系列
)


def _is_cjk_font(font_name: str) -> bool:
    """判断是否为 CJK 字体（需要像素密度辅助检测粗体）。"""
    name = font_name.lower()
    return any(kw in name for kw in _CJK_FONT_KEYWORDS)


def _is_bold(flags: int, font_name: str, linewidth: float | None = None) -> bool:
    """
    判断 span 是否粗体。三级判断：
      1. flags 位标记（bit4=16）—— 标准 PDF 字体标记
      2. linewidth > 0 —— PDF 内容流中的描边宽度（Tr=2 伪粗体的直接信号）
      3. 字体名推断 —— 仅当字体名明确含 bold/heavy/black 字样
    """
    if flags & (1 << 4):  # bit4(16)=粗体
        return True
    if linewidth is not None and linewidth > 0.01:  # 描边渲染 = 伪粗体
        return True
    name = font_name.lower()
    return any(p in name for p in _BOLD_FONT_PATTERNS)


def _measure_span_pixel_density(pix, span_bbox, zoom: int) -> float:
    """
    从预渲染的页面灰度图中采样 span 区域的暗像素比例。
    使用 numpy 向量化操作，单个 span 测量 < 1ms。
    
    参数：
        pix: 预渲染的 fitz.Pixmap（灰度图）
        span_bbox: [x0, y0, x1, y1]（与 pix 同坐标系）
        zoom: 渲染倍率
    
    返回：暗像素比例 [0, 1]
    """
    import numpy as np
    
    x0, y0, x1, y1 = span_bbox
    px0 = max(0, int(x0 * zoom))
    py0 = max(0, int(y0 * zoom))
    px1 = min(pix.width, int(x1 * zoom) + 1)
    py1 = min(pix.height, int(y1 * zoom) + 1)
    
    if px1 <= px0 or py1 <= py0:
        return 0.0
    
    # 零拷贝转为 numpy 数组，向量化计数
    arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.stride)
    region = arr[py0:py1, px0:px1]
    dark = int(np.sum(region < 128))
    total = region.size
    
    return dark / total if total > 0 else 0.0


def _check_cid_bold_by_density(page, all_page_spans: list[dict], zoom: int = 4) -> None:
    """
    对单页 CID 字体 span 测量像素密度（渲染整页一次）。
    密度值写入 span['_density']，后续由 _apply_document_density_baseline 统一判断粗体。

    注意：此函数只负责测量密度，不做粗体判断（判断需要文档级基线）。
    """
    import fitz

    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat, colorspace=fitz.csGRAY)

    for sp in all_page_spans:
        font = (sp.get('font') or '').lower()
        if not _is_cjk_font(font):
            continue
        text = sp.get('text', '')
        if len(text.strip()) < 2:
            continue
        density = _measure_span_pixel_density(pix, sp['bbox'], zoom)
        sp['_density'] = density


def _apply_document_density_baseline(all_spans: list[dict]) -> None:
    """
    文档级像素密度粗体检测。
    在所有页面密度测量完成后调用，用全文档的正文密度建立基线，统一判断粗体。

    优势：即使封面页 span 很少，也能用正文页的基线做对比。
    """
    from collections import defaultdict, Counter

    # 统计每种 (font, size) 的出现次数
    size_counts: Counter = Counter()
    for sp in all_spans:
        font = (sp.get('font') or '').lower()
        if not _is_cjk_font(font):
            continue
        if '_density' not in sp:
            continue
        text = sp.get('text', '')
        if len(text) < 8:
            continue
        size = sp.get('size', 12)
        size_counts[(font, size)] += 1

    # 建立密度基线：按 (font, size) 分组，只取出现≥5次的正文字号
    body_densities: dict[tuple[str, float], list[float]] = defaultdict(list)
    for sp in all_spans:
        font = (sp.get('font') or '').lower()
        if not _is_cjk_font(font):
            continue
        if '_density' not in sp:
            continue
        text = sp.get('text', '')
        if len(text) < 8:
            continue
        size = sp.get('size', 12)
        key = (font, size)
        if size_counts.get(key, 0) < 5:
            continue
        d = sp['_density']
        if 0.02 < d < 0.50:
            body_densities[key].append(d)

    # 同字号基线 + 字体级基线 + 全局基线
    baseline: dict[tuple[str, float], float] = {}
    font_densities: dict[str, list[tuple[float, float]]] = defaultdict(list)
    all_body_densities = []

    for key, densities in body_densities.items():
        if len(densities) >= 2:
            densities.sort()
            med = densities[len(densities) // 2]
            baseline[key] = med
            font_densities[key[0]].append((key[1], med))
            all_body_densities.append(med)

    font_baseline: dict[str, float] = {}
    for font_name, sd_list in font_densities.items():
        best = max(sd_list, key=lambda sd: size_counts.get((font_name, sd[0]), 0))
        font_baseline[font_name] = best[1]

    # 全局基线：用文档中出现次数最多的 (font, size) 组合的密度
    # 而非所有字体的中位数——因为大字标题会拉高中位数
    # 最常见的字号/字体组合几乎一定是正文（非粗体）
    global_baseline = None
    if size_counts:
        dominant_key = size_counts.most_common(1)[0][0]
        global_baseline = baseline.get(dominant_key)
        if global_baseline is None:
            # 众数字号可能没进 baseline（样本不足），取最低密度兜底
            if all_body_densities:
                global_baseline = min(all_body_densities)

    if not baseline and not font_baseline and global_baseline is None:
        return

    # 对比每个 CID span
    for sp in all_spans:
        if sp.get('bold'):
            continue
        font = (sp.get('font') or '').lower()
        if not _is_cjk_font(font):
            continue
        density = sp.get('_density')
        if density is None:
            continue
        text = sp.get('text', '')
        if len(text.strip()) < 2:
            continue

        size = sp.get('size', 12)
        key = (font, size)
        same_font_bl = baseline.get(key)
        font_bl = font_baseline.get(font)

        # 基线选择策略（关键！）：
        # 密度检测的核心是"同字体内部的粗体/常规对比"——即检测同一字体是否被加粗渲染。
        # 不同字体之间的笔画密度差异（如 SimHei 比 FangSong 密）是字体设计差异，
        # 不是"加粗"。因此：
        # - 优先用同字体同字号基线（最精确：同字体同字号的粗体/常规对比）
        # - 次选用字体级基线（同字体不同字号的对比）
        # - 最后用全局基线（仅在字体首次出现、无自身基线时使用）
        bl = same_font_bl
        if bl is None:
            bl = font_bl
        if bl is None:
            bl = global_baseline

        if bl is None:
            continue

        if density > bl * 1.3 and density > 0.05:
            sp['bold'] = True

    # 清理临时字段
    for sp in all_spans:
        sp.pop('_density', None)


# ═══════════════════════════════════════════════════════════════

def extract_spans(
    pdf_path: str | Path,
    *,
    start_page: int = 0,
    end_page: int | None = None,
) -> list[dict[str, Any]]:
    """
    提取 PDF 每个 span 的字符级样式。

    参数：
        pdf_path:   输入 PDF 路径
        start_page: 起始页（0-based，默认 0）
        end_page:   结束页（None=全部）

    返回：
        扁平 span 列表，每项含：
        {
            "page_idx": int,              # 页码（0-based）
            "bbox": [x0, y0, x1, y1],     # PDF 点坐标
            "text": str,                  # 文本
            "font": str,                  # 字体名
            "size": float,                # 字号
            "color": int,                 # sRGB 整数
            "flags": int,                 # 原始 flags
            "bold": bool,                 # 粗体（bit4）
            "italic": bool,               # 斜体（bit1）
        }

    坐标系：PDF 点（1/72 英寸），top-left 原点，Y 轴向下。
    与 MinerU 坐标系完全一致，可直接用 bbox 对齐。
    """
    try:
        import fitz  # PyMuPDF
    except ImportError:
        print(
            "✗ PyMuPDF 未安装。请运行: pip install PyMuPDF",
            file=sys.stderr,
        )
        sys.exit(1)

    doc = fitz.open(str(pdf_path))
    all_spans: list[dict[str, Any]] = []

    total_pages = doc.page_count
    if end_page is None:
        end_page = total_pages
    end_page = min(end_page, total_pages)

    for page_idx in range(start_page, end_page):
        page = doc[page_idx]
        page_dict = page.get_text("dict")

        # 提取 texttrace 的 linewidth（PDF 描边宽度 = 伪粗体的直接信号）
        # 用 bbox 做查找表，与 dict 的 span 匹配
        lw_lookup: dict[tuple, float] = {}  # (x0,y0,x1,y1) → linewidth
        try:
            for t in page.get_texttrace():
                lw = t.get("linewidth")
                if lw is not None and lw > 0.01:
                    bbox = tuple(round(v, 1) for v in t.get("bbox", ()))
                    if bbox:
                        lw_lookup[bbox] = lw
        except Exception:
            pass

        # 先收集当前页所有 span
        page_spans: list[dict[str, Any]] = []
        for block in page_dict.get("blocks", []):
            # type 0 = 文本块，type 1 = 图片块（跳过）
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    flags = span["flags"]
                    font_name = span["font"]
                    # 匹配 linewidth：先精确匹配 bbox，再模糊匹配
                    bbox_key = tuple(round(v, 1) for v in span["bbox"])
                    linewidth = lw_lookup.get(bbox_key)
                    if linewidth is None:
                        # 模糊匹配：bbox 中心点接近的
                        cx = (span["bbox"][0] + span["bbox"][2]) / 2
                        cy = (span["bbox"][1] + span["bbox"][3]) / 2
                        for k, v in lw_lookup.items():
                            tcx = (k[0] + k[2]) / 2
                            tcy = (k[1] + k[3]) / 2
                            if abs(tcx - cx) < 3 and abs(tcy - cy) < 3:
                                linewidth = v
                                break
                    page_spans.append({
                        "page_idx": page_idx,
                        "bbox": list(span["bbox"]),  # [x0, y0, x1, y1]
                        "text": span["text"],
                        "font": font_name,
                        "size": round(span["size"], 1),
                        "color": span["color"],
                        "flags": flags,
                        "bold": _is_bold(flags, font_name, linewidth),
                        "italic": bool(flags & (1 << 1)),   # bit1(2)=斜体
                    })

        # CID 字体像素密度测量（仅测量，不判断——判断需要文档级基线）
        if page_spans:
            try:
                _check_cid_bold_by_density(page, page_spans, zoom=4)
            except Exception:
                pass  # 密度测量失败不影响整体流程

        all_spans.extend(page_spans)

    doc.close()

    # 文档级密度粗体判断（用全文档正文建立基线，统一判断所有 span）
    try:
        _apply_document_density_baseline(all_spans)
    except Exception:
        pass

    return all_spans


# ═══════════════════════════════════════════════════════════════
#  持久化辅助
# ═══════════════════════════════════════════════════════════════

def save_spans(spans: list[dict[str, Any]], output_path: str | Path) -> str:
    """
    将 span 列表保存为 JSON 文件。
    返回保存路径。
    """
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(
        json.dumps(spans, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return str(output_path)


def load_spans(path: str | Path) -> list[dict[str, Any]]:
    """从 JSON 文件加载 span 列表。"""
    return json.loads(Path(path).read_text(encoding="utf-8"))


# ═══════════════════════════════════════════════════════════════
#  CLI（可独立运行测试）
# ═══════════════════════════════════════════════════════════════

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        description="PyMuPDF 字符级样式提取"
    )
    parser.add_argument("pdf", help="输入 PDF 路径")
    parser.add_argument("-o", "--output", default="./output/_spans.json",
                        help="输出 JSON 路径")
    parser.add_argument("--start", type=int, default=0, help="起始页")
    parser.add_argument("--end", type=int, default=None, help="结束页")

    args = parser.parse_args()

    print(f"[2/4] PyMuPDF 样式提取: {args.pdf}", file=sys.stderr)
    spans = extract_spans(args.pdf, start_page=args.start, end_page=args.end)
    saved = save_spans(spans, args.output)

    print(f"  → 提取 {len(spans)} 个 span", file=sys.stderr)
    print(f"  → 保存到: {saved}", file=sys.stderr)

    # 输出 JSON 供上层调用解析
    result = {"spans_count": len(spans), "spans_path": saved}
    print(json.dumps(result, ensure_ascii=False, indent=2))
