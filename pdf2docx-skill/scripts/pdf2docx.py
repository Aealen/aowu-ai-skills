#!/usr/bin/env python3
"""
pdf2docx.py —— PDF 转 DOCX 统一 CLI 入口

串联四步管线：MinerU 解析 → PyMuPDF 样式提取 → bbox 对齐合并 → DOCX 重建
各子命令也可单独调用，便于调试。

用法:
    # 一键全流程
    python3 pdf2docx.py convert input.pdf -o output.docx

    # 环境检查
    python3 pdf2docx.py env.check

    # 字段检查（先头部队，确认 MinerU 2.0 真实结构）
    python3 pdf2docx.py inspect input.pdf

    # 分步调试（每步单独跑）
    python3 pdf2docx.py parse input.pdf -o ./output/       # 仅 MinerU
    python3 pdf2docx.py extract input.pdf -o ./output/     # 仅 PyMuPDF
    python3 pdf2docx.py align middle.json spans.json -o ./output/  # 仅对齐
    python3 pdf2docx.py build merged.json -o output.docx   # 仅重建
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

# 确保能 import 同目录下的模块
_SCRIPTS_DIR = Path(__file__).parent.resolve()
if str(_SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(_SCRIPTS_DIR))


# ═══════════════════════════════════════════════════════════════
#  env.check 子命令
# ═══════════════════════════════════════════════════════════════

def cmd_env_check(args) -> None:
    """检查环境依赖。"""
    print("🔍 检查环境依赖...\n")

    deps = [
        ("Python 3.10+", None, True),   # 特殊处理
    ]

    # Python 版本
    import platform
    py_ver = sys.version_info
    py_ok = py_ver >= (3, 10)
    print(f"  {'✓' if py_ok else '✗'} Python {platform.python_version()} "
          f"({'满足 3.10+' if py_ok else '需要 3.10+'})")

    # Python 包
    py_deps = [
        ("fitz", "PyMuPDF", "pip install PyMuPDF"),
        ("docx", "python-docx", "pip install python-docx"),
    ]
    all_ok = py_ok

    for mod, name, install_hint in py_deps:
        try:
            m = __import__(mod)
            ver = getattr(m, "__version__", "installed")
            print(f"  ✓ {name} ({ver})")
        except ImportError:
            print(f"  ✗ {name} 未安装  →  {install_hint}")
            all_ok = False

    # MinerU（import 名 mineru 或 magic_pdf）
    mineru_ok = True
    try:
        import mineru  # noqa: F401
        print(f"  ✓ mineru (installed)")
    except ImportError:
        try:
            import magic_pdf  # noqa: F401  # 旧版兼容
            print(f"  ○ magic_pdf (旧版包名，建议升级到 mineru 2.0)")
        except ImportError:
            print(f"  ✗ mineru 未安装  →  pip install -U \"mineru[all]\"")
            print(f"    安装后需运行: mineru-models-download")
            mineru_ok = False
            all_ok = False

    print()
    if all_ok:
        print("✅ 环境就绪，可以执行转换。")
    else:
        print("❌ 环境不完整，请按提示安装缺失依赖。")
        print("   或运行: bash setup.sh")


# ═══════════════════════════════════════════════════════════════
#  convert 子命令（一键全流程）
# ═══════════════════════════════════════════════════════════════

def cmd_convert(args) -> None:
    """一键全流程：PDF → DOCX。"""
    pdf_path = Path(args.pdf).resolve()
    output_path = Path(args.output).resolve()
    work_dir = Path(args.work_dir).resolve() if args.work_dir \
        else output_path.parent / f"_pdf2docx_work_{pdf_path.stem}"

    work_dir.mkdir(parents=True, exist_ok=True)

    print(f"═══════════════════════════════════════════════", file=sys.stderr)
    print(f"  PDF 转 DOCX", file=sys.stderr)
    print(f"  输入: {pdf_path}", file=sys.stderr)
    print(f"  输出: {output_path}", file=sys.stderr)
    print(f"  中间产物: {work_dir}", file=sys.stderr)
    print(f"═══════════════════════════════════════════════\n", file=sys.stderr)

    if not pdf_path.exists():
        print(f"✗ 输入文件不存在: {pdf_path}", file=sys.stderr)
        sys.exit(1)

    # ── Step 1: MinerU 解析（同步 API，进程内设环境变量）──
    from parse_mineru import parse_with_mineru
    middle_path, images_dir = parse_with_mineru(
        pdf_path=str(pdf_path),
        output_dir=str(work_dir),
        parse_method=args.method,
        formula_enable=not args.no_formula,
        table_enable=not args.no_table,
        language=args.lang,
    )

    # ── Step 2: PyMuPDF 样式提取 ──
    from parse_pymupdf import extract_spans, save_spans, extract_underline_lines
    spans = extract_spans(str(pdf_path))
    spans_path = save_spans(spans, work_dir / "_spans.json")
    # B 类下划线提取（填空区域空白下划线，A 类已在 extract_spans 中标注到 span）
    underline_lines = extract_underline_lines(str(pdf_path), spans)
    print(f"[2/4] PyMuPDF 样式提取: {len(spans)} 个 span, "
          f"{len(underline_lines)} 条填空下划线 → {spans_path}",
          file=sys.stderr)

    # ── Step 3: bbox 对齐合并 ──
    from align import align_and_merge, load_mineru, save_merged
    mineru_data = load_mineru(middle_path)
    merged_data, stats = align_and_merge(
        mineru_data, spans, iou_threshold=args.iou,
        underline_lines=underline_lines,
    )
    merged_path = save_merged(merged_data, work_dir / "_merged.json")

    rate = (stats["matched"] / stats["total"] * 100) if stats["total"] else 0
    print(f"[3/4] bbox 对齐: 命中率 {rate:.1f}% "
          f"({stats['matched']}/{stats['total']}) → {merged_path}",
          file=sys.stderr)
    if rate < 80 and stats["total"] > 0:
        print(f"  ⚠️ 命中率低于 80%，样式还原可能不完整", file=sys.stderr)

    # ── Step 4: DOCX 重建 ──
    from build_docx import build_docx
    result_path = build_docx(merged_data, images_dir, str(output_path),
                             pdf_path=str(pdf_path))

    # ── 清理中间产物（除非 --keep-work）──
    if not args.keep_work:
        import shutil
        shutil.rmtree(work_dir, ignore_errors=True)
        print(f"\n  → 已清理中间产物（--keep-work 可保留）", file=sys.stderr)

    print(f"\n✅ 转换完成: {result_path}", file=sys.stderr)

    # 输出 JSON 结果（供程序化调用）
    result = {
        "output": result_path,
        "stats": {
            "spans_total": len(spans),
            "align_matched": stats["matched"],
            "align_total": stats["total"],
            "align_rate": round(rate, 1),
        },
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))


# ═══════════════════════════════════════════════════════════════
#  inspect 子命令
# ═══════════════════════════════════════════════════════════════

def cmd_inspect(args) -> None:
    """字段检查（先头部队）。"""
    from pdf_inspect import inspect_pymupdf_spans
    inspect_pymupdf_spans(args.pdf, page_num=args.page, verbose=args.verbose)


# ═══════════════════════════════════════════════════════════════
#  分步调试子命令
# ═══════════════════════════════════════════════════════════════

def cmd_parse(args) -> None:
    """仅 MinerU 解析。"""
    from parse_mineru import parse_with_mineru
    middle_path, images_dir = parse_with_mineru(
        pdf_path=args.pdf,
        output_dir=args.output,
        parse_method=args.method,
        formula_enable=not args.no_formula,
        table_enable=not args.no_table,
        language=args.lang,
    )
    print(json.dumps(
        {"middle_json": middle_path, "images_dir": images_dir},
        ensure_ascii=False, indent=2
    ))


def cmd_extract(args) -> None:
    """仅 PyMuPDF 样式提取。"""
    from parse_pymupdf import extract_spans, save_spans
    spans = extract_spans(args.pdf, start_page=args.start, end_page=args.end)
    saved = save_spans(spans, Path(args.output) / "_spans.json")
    print(json.dumps(
        {"spans_count": len(spans), "spans_path": saved},
        ensure_ascii=False, indent=2
    ))


def cmd_align(args) -> None:
    """仅 bbox 对齐。"""
    from align import align_and_merge, load_mineru, save_merged
    mineru_data = load_mineru(args.middle)
    spans = json.loads(Path(args.spans).read_text(encoding="utf-8"))
    merged, stats = align_and_merge(mineru_data, spans, iou_threshold=args.iou)
    saved = save_merged(merged, Path(args.output) / "_merged.json")
    print(json.dumps(
        {"merged_path": saved, "stats": stats},
        ensure_ascii=False, indent=2
    ))


def cmd_build(args) -> None:
    """仅 DOCX 重建。"""
    from build_docx import build_docx
    merged_data = json.loads(Path(args.merged).read_text(encoding="utf-8"))
    result = build_docx(merged_data, args.images_dir, args.output,
                        pdf_path=getattr(args, 'pdf', None))
    print(json.dumps({"output": result}, ensure_ascii=False, indent=2))


# ═══════════════════════════════════════════════════════════════
#  CLI 主入口
# ═══════════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(
        prog="pdf2docx.py",
        description="PDF 转 DOCX（MinerU 版面结构 + PyMuPDF 字符样式）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  # 一键全流程
  python3 pdf2docx.py convert input.pdf -o output.docx

  # 环境检查
  python3 pdf2docx.py env.check

  # 字段检查（先跑这个确认 MinerU 真实结构）
  python3 pdf2docx.py inspect input.pdf -v

  # 分步调试
  python3 pdf2docx.py parse input.pdf -o ./work/      # 仅 MinerU
  python3 pdf2docx.py extract input.pdf -o ./work/    # 仅 PyMuPDF
  python3 pdf2docx.py align work/_middle.json work/_spans.json -o ./work/
  python3 pdf2docx.py build work/_merged.json -o output.docx
        """,
    )
    sub = parser.add_subparsers(dest="command", required=True)

    # ── convert ──
    p_conv = sub.add_parser("convert", help="一键全流程：PDF → DOCX")
    p_conv.add_argument("pdf", help="输入 PDF 路径")
    p_conv.add_argument("-o", "--output", required=True, help="输出 docx 路径")
    p_conv.add_argument("--work-dir", default=None,
                        help="中间产物目录（默认在输出旁边）")
    p_conv.add_argument("--keep-work", action="store_true",
                        help="保留中间产物（默认清理）")
    p_conv.add_argument("-m", "--method", default="auto",
                        choices=["auto", "txt", "ocr"],
                        help="解析方法（auto=自动判断）")
    p_conv.add_argument("-l", "--lang", default="ch", help="语言")
    p_conv.add_argument("--no-formula", action="store_true",
                        help="关闭公式识别（加快速度）")
    p_conv.add_argument("--no-table", action="store_true",
                        help="关闭表格识别（加快速度）")
    p_conv.add_argument("--iou", type=float, default=0.3,
                        help="bbox 对齐 IoU 阈值（默认 0.3）")
    p_conv.set_defaults(func=cmd_convert)

    # ── env.check ──
    p_env = sub.add_parser("env.check", help="检查环境依赖")
    p_env.set_defaults(func=cmd_env_check)

    # ── inspect ──
    p_ins = sub.add_parser("inspect", help="字段检查（PyMuPDF spans）")
    p_ins.add_argument("pdf", help="PDF 路径")
    p_ins.add_argument("--page", type=int, default=0, help="页码")
    p_ins.add_argument("-v", "--verbose", action="store_true")
    p_ins.set_defaults(func=cmd_inspect)

    # ── parse（仅 MinerU）──
    p_par = sub.add_parser("parse", help="仅 MinerU 解析")
    p_par.add_argument("pdf", help="PDF 路径")
    p_par.add_argument("-o", "--output", default="./output", help="输出目录")
    p_par.add_argument("-m", "--method", default="auto",
                       choices=["auto", "txt", "ocr"])
    p_par.add_argument("-l", "--lang", default="ch")
    p_par.add_argument("--no-formula", action="store_true")
    p_par.add_argument("--no-table", action="store_true")
    p_par.set_defaults(func=cmd_parse)

    # ── extract（仅 PyMuPDF）──
    p_ext = sub.add_parser("extract", help="仅 PyMuPDF 样式提取")
    p_ext.add_argument("pdf", help="PDF 路径")
    p_ext.add_argument("-o", "--output", default="./output", help="输出目录")
    p_ext.add_argument("--start", type=int, default=0)
    p_ext.add_argument("--end", type=int, default=None)
    p_ext.set_defaults(func=cmd_extract)

    # ── align（仅对齐）──
    p_aln = sub.add_parser("align", help="仅 bbox 对齐")
    p_aln.add_argument("middle", help="middle.json 路径")
    p_aln.add_argument("spans", help="spans.json 路径")
    p_aln.add_argument("-o", "--output", default="./output", help="输出目录")
    p_aln.add_argument("--iou", type=float, default=0.3)
    p_aln.set_defaults(func=cmd_align)

    # ── build（仅重建）──
    p_bld = sub.add_parser("build", help="仅 DOCX 重建")
    p_bld.add_argument("merged", help="merged.json 路径")
    p_bld.add_argument("-o", "--output", default="./output/结果.docx",
                       help="输出 docx 路径")
    p_bld.add_argument("--images-dir", default=None)
    p_bld.add_argument("--pdf", default=None,
                       help="源 PDF 路径（用于提取表格列宽和精确边距）")
    p_bld.set_defaults(func=cmd_build)

    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
