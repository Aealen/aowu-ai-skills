#!/usr/bin/env python3
"""
inspect.py —— 字段结构检查工具（先头部队）

⚠️ 这个脚本必须最先写、最先跑。
原因：MinerU 2.0 重构过架构，middle.json 的真实字段结构（块类型分布、
title 的 text_level、table 的 table_body 格式、image 的字段名）有不确定性。
官方文档描述到块级，细节必须实测确认，才能让 align.py / build_docx.py 写得准。

用法:
    # 检查 MinerU 输出的 middle.json（需要先跑 parse_mineru.py 产出）
    python3 inspect.py middle <middle.json路径>

    # 检查 PyMuPDF 的 span 真实样式字段（直接读 PDF）
    python3 inspect.py spans <pdf路径> [--page 0]

    # 检查对齐合并后的 merged.json
    python3 inspect.py merged <merged.json路径>

    # 一键全检（middle + spans）
    python3 inspect.py all <pdf路径> <middle.json路径>
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any


# ═══════════════════════════════════════════════════════════════
#  MinerU middle.json 检查
# ═══════════════════════════════════════════════════════════════

def _preview(block: dict, max_len: int = 60) -> str:
    """从 block 中提取预览文本。"""
    # MinerU 的文本可能在 content_list 或 middle.json 的不同层级
    # middle.json: blocks -> lines -> spans -> content
    # 先尝试多种已知字段
    if "content" in block:
        return str(block["content"])[:max_len]
    # 尝试从 lines.spans 拼接
    text_parts = []
    for line in block.get("lines", []):
        for span in line.get("spans", []):
            if "content" in span:
                text_parts.append(span["content"])
            elif "text" in span:
                text_parts.append(span["text"])
    if text_parts:
        return "".join(text_parts)[:max_len]
    return "(无文本)"


def _collect_blocks(data: dict) -> list[tuple[int, dict]]:
    """
    从 middle.json 中收集所有 block，附带 page_idx。
    兼容 MinerU 2.0 的多种可能结构：
      - pdf_info[].para_blocks[]（pipeline backend 旧结构）
      - pdf_info[].blocks[]（可能的 VLM 结构）
      - 顶层 blocks[]
    返回 [(page_idx, block), ...]
    """
    results: list[tuple[int, dict]] = []

    # 结构1: pdf_info[].para_blocks[]
    pdf_info = data.get("pdf_info", [])
    if pdf_info:
        for page in pdf_info:
            page_idx = page.get("page_idx", 0)
            # 优先 para_blocks，回退 blocks
            blocks = page.get("para_blocks") or page.get("blocks") or []
            for b in blocks:
                results.append((page_idx, b))
        return results

    # 结构2: 顶层 blocks[]
    if "blocks" in data:
        for b in data["blocks"]:
            results.append((0, b))
        return results

    return results


def inspect_middle_json(path: str, verbose: bool = False) -> None:
    """
    检查 MinerU middle.json 的真实字段结构。
    打印：总页数、各页块类型分布、各类型样例。
    """
    p = Path(path)
    if not p.exists():
        print(f"✗ 文件不存在: {path}", file=sys.stderr)
        sys.exit(1)

    data = json.loads(p.read_text(encoding="utf-8"))
    print(f"📄 检查 middle.json: {path}")
    print(f"   文件大小: {p.stat().st_size / 1024:.1f} KB")
    print()

    # ── 顶层 key ──
    print("【顶层字段】")
    top_keys = list(data.keys())
    print(f"   keys: {top_keys}")
    print()

    # ── pdf_info 结构 ──
    pdf_info = data.get("pdf_info", [])
    print(f"【页数】{len(pdf_info)}")
    if pdf_info:
        print(f"   第1页 keys: {list(pdf_info[0].keys())}")
    print()

    # ── 收集所有 block ──
    all_blocks = _collect_blocks(data)
    print(f"【总块数】{len(all_blocks)}")
    print()

    if not all_blocks:
        print("⚠️ 未找到 block 数据。可能 middle.json 结构与预期不同。")
        print("   请手动检查 JSON 结构。以下是完整顶层结构预览：")
        print(json.dumps(data, ensure_ascii=False, indent=2)[:2000])
        return

    # ── 块类型分布 ──
    print("【块类型分布】")
    type_counter: dict[str, int] = {}
    for _, b in all_blocks:
        # type 字段可能有多种写法
        t = b.get("type") or b.get("block_type") or "unknown"
        type_counter[t] = type_counter.get(t, 0) + 1
    for t, cnt in sorted(type_counter.items(), key=lambda x: -x[1]):
        print(f"   {t}: {cnt}")
    print()

    # ── 各类型样例（前2页） ──
    print("【各类型样例（前2页）】")
    seen_pages = set()
    for page_idx, b in all_blocks:
        if page_idx > 1:
            break
        seen_pages.add(page_idx)

    for target_page in sorted(seen_pages):
        print(f"  --- page_idx={target_page} ---")
        page_blocks = [(pi, b) for pi, b in all_blocks if pi == target_page]
        # 找各类样例
        type_samples: dict[str, dict] = {}
        for _, b in page_blocks:
            t = b.get("type") or b.get("block_type") or "unknown"
            if t not in type_samples:
                type_samples[t] = b

        for t, b in type_samples.items():
            if t == "title":
                print(f"  [title]   level={b.get('text_level')} "
                      f"text_level_alt={b.get('level')} "
                      f"text={_preview(b)[:40]}")
                if verbose:
                    print(f"            keys: {list(b.keys())}")
            elif t == "table":
                tbody = b.get("table_body") or b.get("html") or ""
                print(f"  [table]   table_body前80字符: {str(tbody)[:80]}")
                if verbose:
                    print(f"            keys: {list(b.keys())}")
            elif t == "image":
                print(f"  [image]   img_path={b.get('img_path')} "
                      f"bbox={b.get('bbox')}")
                if verbose:
                    print(f"            keys: {list(b.keys())}")
            elif t in ("text", "paragraph"):
                preview = _preview(b)
                print(f"  [{t}]  text: {preview[:50]}")
                if verbose:
                    # 打印一个 text block 的完整 lines/spans 结构
                    print(f"            完整结构: "
                          f"{json.dumps(b, ensure_ascii=False)[:300]}")
            elif t == "list":
                print(f"  [list]    text={_preview(b)[:40]}")
                if verbose:
                    print(f"            keys: {list(b.keys())}")
            else:
                print(f"  [{t}]  text={_preview(b)[:40]}")
                if verbose:
                    print(f"            keys: {list(b.keys())}")
    print()

    # ── 一个 text block 的完整 lines/spans 结构（关键！）──
    print("【text block 完整 lines/spans 结构（第1个找到的）】")
    for _, b in all_blocks:
        t = b.get("type") or b.get("block_type") or "unknown"
        if t in ("text", "paragraph") and b.get("lines"):
            print(json.dumps(b, ensure_ascii=False, indent=2)[:800])
            print()
            print("   ⚠️ 确认 span 是否有 bbox 字段（对齐算法依赖此字段）")
            break
    else:
        print("   （未找到含 lines 的 text block）")


# ═══════════════════════════════════════════════════════════════
#  PyMuPDF spans 检查
# ═══════════════════════════════════════════════════════════════

def inspect_pymupdf_spans(pdf_path: str, page_num: int = 0, verbose: bool = False) -> None:
    """
    检查 PyMuPDF span 的真实样式字段。
    打印前 10 个 span 的 font/size/color/flags。
    """
    try:
        import fitz  # PyMuPDF
    except ImportError:
        print("✗ PyMuPDF 未安装。请运行: pip install PyMuPDF", file=sys.stderr)
        sys.exit(1)

    p = Path(pdf_path)
    if not p.exists():
        print(f"✗ 文件不存在: {pdf_path}", file=sys.stderr)
        sys.exit(1)

    doc = fitz.open(pdf_path)
    print(f"📄 检查 PyMuPDF spans: {pdf_path}")
    print(f"   总页数: {doc.page_count}")
    print()

    if page_num >= doc.page_count:
        print(f"⚠️ page {page_num} 超出范围，使用第0页")
        page_num = 0

    page = doc[page_num]
    page_rect = page.rect
    print(f"【页面尺寸】{page_rect.width:.1f} × {page_rect.height:.1f} pt")
    print()

    d = page.get_text("dict")
    print(f"【dict 顶层 keys】{list(d.keys())}")
    print()

    # ── blocks 结构 ──
    blocks = d.get("blocks", [])
    print(f"【blocks 数量】{len(blocks)}")
    type_dist = {}
    for b in blocks:
        bt = b.get("type", -1)
        type_dist[bt] = type_dist.get(bt, 0) + 1
    print(f"   block type 分布: {type_dist}  (0=文本, 1=图片)")
    print()

    # ── span 样例 ──
    print("【前 10 个 span 的样式字段】")
    print(f"   {'font':<25} {'size':>6} {'color':>8} {'flags':>5} "
          f"{'bold':>5} {'italic':>6}  text")
    print(f"   {'-'*25} {'-'*6} {'-'*8} {'-'*5} {'-'*5} {'-'*6}  {'-'*20}")

    count = 0
    for block in blocks:
        if block.get("type") != 0:  # 只看文本块
            continue
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                flags = span["flags"]
                bold = bool(flags & (1 << 4))     # bit4=16=粗体
                italic = bool(flags & (1 << 1))   # bit1=2=斜体
                print(f"   {span['font'][:25]:<25} {span['size']:>6.1f} "
                      f"{span['color']:>8} {flags:>5} "
                      f"{'Y' if bold else '-':>5} "
                      f"{'Y' if italic else '-':>6}  "
                      f"{span['text'][:30]}")
                count += 1
                if count >= 10:
                    break
            if count >= 10:
                break
        if count >= 10:
            break

    print()
    print("【flags 位掩码说明】")
    print("   bit0(1)=上标  bit1(2)=斜体  bit2(4)=衬线")
    print("   bit3(8)=等宽  bit4(16)=粗体")
    print()

    # ── 一个 span 的完整字段（关键！）──
    if verbose:
        print("【一个 span 的完整字段】")
        for block in blocks:
            if block.get("type") != 0:
                continue
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    print(json.dumps(span, ensure_ascii=False, indent=2))
                    print()
                    print("   ⚠️ 确认有 bbox 字段（对齐算法依赖此字段）")
                    doc.close()
                    return

    doc.close()


# ═══════════════════════════════════════════════════════════════
#  合并后的 merged.json 检查
# ═══════════════════════════════════════════════════════════════

def inspect_merged(path: str) -> None:
    """
    检查对齐合并后的 merged.json，统计 PyMuPDF 样式贴回的成功率。
    """
    p = Path(path)
    if not p.exists():
        print(f"✗ 文件不存在: {path}", file=sys.stderr)
        sys.exit(1)

    data = json.loads(p.read_text(encoding="utf-8"))
    print(f"📄 检查 merged.json: {path}")
    print()

    all_blocks = _collect_blocks(data)
    total_spans = 0
    styled_spans = 0

    for _, b in all_blocks:
        for line in b.get("lines", []):
            for span in line.get("spans", []):
                total_spans += 1
                if span.get("_style"):
                    styled_spans += 1

    print(f"【对齐命中率统计】")
    print(f"   总 span 数: {total_spans}")
    print(f"   已贴样式: {styled_spans}")
    if total_spans > 0:
        rate = styled_spans / total_spans * 100
        print(f"   命中率: {rate:.1f}%")
        if rate < 80:
            print(f"   ⚠️ 命中率低于 80%，需调整 align.py 的 IoU 阈值")
        else:
            print(f"   ✓ 命中率达标")


# ═══════════════════════════════════════════════════════════════
#  CLI
# ═══════════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(
        description="检查 MinerU middle.json / PyMuPDF spans 的真实字段结构"
    )
    sub = parser.add_subparsers(dest="command", required=True)

    p_mid = sub.add_parser("middle", help="检查 MinerU middle.json")
    p_mid.add_argument("path", help="middle.json 文件路径")
    p_mid.add_argument("-v", "--verbose", action="store_true", help="打印完整字段")

    p_spans = sub.add_parser("spans", help="检查 PyMuPDF span 样式")
    p_spans.add_argument("pdf", help="PDF 文件路径")
    p_spans.add_argument("--page", type=int, default=0, help="页码（0-based）")
    p_spans.add_argument("-v", "--verbose", action="store_true", help="打印完整 span 字段")

    p_merged = sub.add_parser("merged", help="检查对齐合并后的 merged.json")
    p_merged.add_argument("path", help="merged.json 文件路径")

    p_all = sub.add_parser("all", help="一键全检（middle + spans）")
    p_all.add_argument("pdf", help="PDF 文件路径")
    p_all.add_argument("middle", help="middle.json 文件路径")
    p_all.add_argument("-v", "--verbose", action="store_true", help="详细输出")

    args = parser.parse_args()

    if args.command == "middle":
        inspect_middle_json(args.path, args.verbose)
    elif args.command == "spans":
        inspect_pymupdf_spans(args.pdf, args.page, args.verbose)
    elif args.command == "merged":
        inspect_merged(args.path)
    elif args.command == "all":
        inspect_middle_json(args.middle, args.verbose)
        print("\n" + "=" * 60 + "\n")
        inspect_pymupdf_spans(args.pdf, verbose=args.verbose)


if __name__ == "__main__":
    main()
