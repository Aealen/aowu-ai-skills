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
# 注意：方正小标宋(FZXBSJW)、黑体(SimHei) 等是标题字体但不是粗体字体，
# 不应强制推断为 bold——应尊重源 PDF 的 flags 值。
# 只对字体名里明确含粗体标识的才推断。
_BOLD_FONT_PATTERNS = (
    "bold",     # TimesNewRomanPS-BoldMT, Arial-Bold 等
    "heavy",
    "black",
)


def _is_bold(flags: int, font_name: str) -> bool:
    """
    判断 span 是否粗体。
    两级判断：
      1. flags 位标记（bit4=16）—— 标准方式，优先信任
      2. 字体名推断 —— 仅当字体名明确含 bold/heavy/black 字样
    """
    if flags & (1 << 4):  # bit4(16)=粗体
        return True
    name = font_name.lower()
    return any(p in name for p in _BOLD_FONT_PATTERNS)


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

        for block in page_dict.get("blocks", []):
            # type 0 = 文本块，type 1 = 图片块（跳过）
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    flags = span["flags"]
                    font_name = span["font"]
                    all_spans.append({
                        "page_idx": page_idx,
                        "bbox": list(span["bbox"]),  # [x0, y0, x1, y1]
                        "text": span["text"],
                        "font": font_name,
                        "size": round(span["size"], 1),
                        "color": span["color"],
                        "flags": flags,
                        "bold": _is_bold(flags, font_name),
                        "italic": bool(flags & (1 << 1)),   # bit1(2)=斜体
                    })

    doc.close()
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
