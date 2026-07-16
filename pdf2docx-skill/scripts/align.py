#!/usr/bin/env python3
"""
align.py —— bbox 对齐合并（核心模块）

职责：把 PyMuPDF 的字符级样式（font/size/color）"贴回"到 MinerU 的版面 block 上。
这是双数据源方案的核心：MinerU 给结构，PyMuPDF 给样式，用 bbox 坐标对齐合并。

为什么能对齐（已核实）：
  - MinerU 底层用 PyMuPDF 做文本提取
  - 两者坐标系完全一致：PDF 点（1/72 英寸），top-left 原点，Y 轴向下
  - MinerU block 的 bbox 与 PyMuPDF span 的 bbox 直接可比较
  - 用矩形相交（IoU 或包含关系）判断即可

对齐策略：
  对 MinerU 的每个 block：
    1. 按 page_idx 找出同页的 PyMuPDF spans
    2. 对 block 内每个 MinerU span，找 bbox 最匹配的 PyMuPDF span
    3. 用 IoU（交并比）衡量匹配度，阈值 0.3 起步
    4. 匹配上则把 PyMuPDF 的 font/size/color/bold/italic 贴到 _style 字段

依赖：无第三方依赖（纯几何计算）

用法（被 pdf2docx.py 调用）:
    from align import align_and_merge
    merged = align_and_merge(mineru_data, spans)
"""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any

# IoU 匹配阈值：MinerU span 与 PyMuPDF span 的 bbox IoU 超过此值才贴样式。
# 0.3 起步（MVP 调参项，真实 PDF 上实测后调整）。
# TODO(调优): 见 references/tuning-guide.md，根据真实数据调整此阈值
DEFAULT_IOU_THRESHOLD = 0.3


# ═══════════════════════════════════════════════════════════════
#  几何工具：bbox 相交与 IoU
# ═══════════════════════════════════════════════════════════════

def bbox_overlap(b1: list | tuple, b2: list | tuple) -> bool:
    """
    判断两个 bbox [x0,y0,x1,y1] 是否相交（有面积重叠）。
    """
    x0 = max(b1[0], b2[0])
    y0 = max(b1[1], b2[1])
    x1 = min(b1[2], b2[2])
    y1 = min(b1[3], b2[3])
    return x0 < x1 and y0 < y1


def bbox_area(b: list | tuple) -> float:
    """bbox 面积。"""
    w = b[2] - b[0]
    h = b[3] - b[1]
    return max(0.0, w) * max(0.0, h)


def calc_iou(b1: list | tuple, b2: list | tuple) -> float:
    """
    计算两个 bbox 的 IoU（交并比）。
    返回 0.0~1.0。
    """
    # 交集
    ix0 = max(b1[0], b2[0])
    iy0 = max(b1[1], b2[1])
    ix1 = min(b1[2], b2[2])
    iy1 = min(b1[3], b2[3])
    iw = max(0.0, ix1 - ix0)
    ih = max(0.0, iy1 - iy0)
    intersection = iw * ih

    if intersection == 0:
        return 0.0

    union = bbox_area(b1) + bbox_area(b2) - intersection
    if union <= 0:
        return 0.0
    return intersection / union


def calc_containment(b_small: list | tuple, b_big: list | tuple) -> float:
    """
    计算小 bbox 被大 bbox 包含的比例（小面积中被大覆盖的比例）。
    用于处理 MinerU span 比 PyMuPDF span 大的情况（block 整体匹配）。
    """
    ix0 = max(b_small[0], b_big[0])
    iy0 = max(b_small[1], b_big[1])
    ix1 = min(b_small[2], b_big[2])
    iy1 = min(b_small[3], b_big[3])
    iw = max(0.0, ix1 - ix0)
    ih = max(0.0, iy1 - iy0)
    intersection = iw * ih

    small_area = bbox_area(b_small)
    if small_area <= 0:
        return 0.0
    return intersection / small_area


# ═══════════════════════════════════════════════════════════════
#  MinerU 数据结构提取（兼容多种结构）
# ═══════════════════════════════════════════════════════════════

def collect_page_blocks(data: dict) -> dict[int, list[dict]]:
    """
    从 MinerU middle.json 中提取按 page_idx 分组的 block 列表。
    兼容 MinerU 2.0 多种结构：
      - pdf_info[].para_blocks[]
      - pdf_info[].blocks[]
    返回 {page_idx: [block, ...]}
    """
    by_page: dict[int, list[dict]] = {}
    pdf_info = data.get("pdf_info", [])

    for page in pdf_info:
        page_idx = page.get("page_idx", 0)
        blocks = page.get("para_blocks") or page.get("blocks") or []
        by_page.setdefault(page_idx, []).extend(blocks)

    return by_page


# ═══════════════════════════════════════════════════════════════
#  核心：样式贴回
# ═══════════════════════════════════════════════════════════════

def _find_best_match(
    target_bbox: list | tuple,
    candidates: list[dict[str, Any]],
    iou_threshold: float,
) -> dict[str, Any] | None:
    """
    在候选 spans 中找与 target_bbox 最匹配的 span。
    匹配策略：优先 IoU，IoU 不足时看包含关系。
    """
    best = None
    best_score = 0.0

    for cand in candidates:
        cand_bbox = cand["bbox"]
        # 优先 IoU
        iou = calc_iou(target_bbox, cand_bbox)
        if iou > best_score:
            best_score = iou
            best = cand
            continue
        # IoU 不足时，看 target 是否被候选包含（target 是候选的一部分）
        containment = calc_containment(target_bbox, cand_bbox)
        if containment > best_score and containment > iou_threshold:
            best_score = containment
            best = cand

    if best and best_score > iou_threshold:
        return best
    return None


def _attach_styles_to_block(
    block: dict[str, Any],
    page_spans: list[dict[str, Any]],
    iou_threshold: float,
    stats: dict[str, int],
) -> None:
    """
    把 PyMuPDF 样式贴到 MinerU block 的 lines.spans 上。
    MinerU 的 span 有 bbox + content，PyMuPDF 的 span 有 bbox + 样式，
    用 bbox 匹配把样式贴过去。

    MinerU block 结构（middle.json）:
      block.lines[].spans[] —— 每个 span 有 bbox 和 content

    直接修改 block（in-place），给匹配到的 span 加 _style 字段。
    """
    for line in block.get("lines", []):
        for mspan in line.get("spans", []):
            m_bbox = mspan.get("bbox")
            stats["total"] += 1
            if not m_bbox:
                stats["no_bbox"] += 1
                continue

            best = _find_best_match(m_bbox, page_spans, iou_threshold)
            if best:
                mspan["_style"] = {
                    "font": best["font"],
                    "size": best["size"],
                    "bold": best["bold"],
                    "italic": best["italic"],
                    "color": best["color"],
                    "underline": best.get("underline", False),
                }
                stats["matched"] += 1
            else:
                stats["unmatched"] += 1


def _attach_style_to_block_fallback(
    block: dict[str, Any],
    page_spans: list[dict[str, Any]],
) -> None:
    """
    兜底：block 无 lines.spans 结构时（如 title/image/table），
    用 block 整体 bbox 找占比最大的样式，贴到 block._style。
    """
    block_bbox = block.get("bbox")
    if not block_bbox or not page_spans:
        return

    # 统计落在此 block 范围内的 spans 样式，取占比最大的
    style_counter: dict[str, int] = {}
    for s in page_spans:
        if not bbox_overlap(block_bbox, s["bbox"]):
            continue
        # 用 "size|bold|font" 作为样式指纹统计
        key = f"{s['size']}|{s['bold']}|{s['font']}"
        style_counter[key] = style_counter.get(key, 0) + 1

    if not style_counter:
        return

    # 取出现最多的样式指纹
    dominant_key = max(style_counter, key=style_counter.get)
    size_str, bold_str, font = dominant_key.split("|")
    # 再找对应的具体 span 拿完整样式
    for s in page_spans:
        if (f"{s['size']}|{s['bold']}|{s['font']}" == dominant_key
                and bbox_overlap(block_bbox, s["bbox"])):
            block["_style"] = {
                "font": s["font"],
                "size": s["size"],
                "bold": s["bold"],
                "italic": s["italic"],
                "color": s["color"],
                "underline": s.get("underline", False),
            }
            break


# ═══════════════════════════════════════════════════════════════
#  对外主接口
# ═══════════════════════════════════════════════════════════════


def _inject_underline_spans(
    mineru_data: dict[str, Any],
    underline_lines: list[dict[str, Any]],
) -> int:
    """
    把 B 类下划线（填空区域空白下划线）作为虚拟 span 注入 MinerU 数据。

    每条下划线查找它所在的 MinerU block 和 line，创建一个带下划线样式的
    空格 span，按 x 坐标插入到 line.spans 的正确位置。

    返回注入的 span 数量。
    """
    if not underline_lines:
        return 0

    page_blocks = collect_page_blocks(mineru_data)
    injected = 0

    for ul in underline_lines:
        page_idx = ul["page_idx"]
        ul_x0, ul_y, ul_x1 = ul["x0"], ul["y"], ul["x1"]
        blocks = page_blocks.get(page_idx, [])

        # 找包含此下划线的 block（bbox 覆盖线条位置）
        target_block = None
        for b in blocks:
            bbox = b.get("bbox") or []
            if len(bbox) < 4:
                continue
            # 线条中心在 block bbox 内
            cx = (ul_x0 + ul_x1) / 2
            if bbox[0] - 5 <= cx <= bbox[2] + 5 and bbox[1] - 10 <= ul_y <= bbox[3] + 10:
                target_block = b
                break

        if not target_block or not target_block.get("lines"):
            continue

        # 找包含此下划线的 line（line bbox 的 y 范围包含线条 y）
        target_line = None
        for ln in target_block.get("lines", []):
            lbbox = ln.get("bbox") or []
            if len(lbbox) < 4:
                continue
            # 线条 y 在 line 的 y 范围内（容差 10pt，因为下划线可能在行底部稍下方）
            if lbbox[1] - 10 <= ul_y <= lbbox[3] + 5:
                target_line = ln
                break

        if not target_line:
            continue

        # 创建虚拟 span：空格填充下划线长度
        # 每字符约12pt（FangSong 12pt 全角空格宽度），估算空格数
        line_width = ul_x1 - ul_x0
        n_spaces = max(1, round(line_width / 11))

        # 获取同 line 主导样式（首个有 _style 的 span）
        dominant_style = None
        for sp in target_line.get("spans", []):
            st = sp.get("_style")
            if st and st.get("font"):
                dominant_style = dict(st)
                break
        if not dominant_style:
            dominant_style = {"font": "FangSong", "size": 12.0, "bold": False,
                              "italic": False, "color": 0}
        dominant_style["underline"] = True

        virtual_span = {
            "bbox": [ul_x0, ul_y - 2, ul_x1, ul_y + 2],
            # 用全角空格(U+3000)而非半角空格——WPS 对半角空格的下划线
            # 可能不渲染，全角空格更可靠地显示下划线填空区域
            "content": "\u3000" * n_spaces,
            "type": "text",
            "_style": dominant_style,
        }

        # 按 x 坐标插入到 spans 正确位置
        spans_list = target_line.setdefault("spans", [])

        # 检查下划线是否落在某个 span 的 x 范围内（MinerU 常把整行合并成一个 span）
        # 若是，需要拆分该 span：在下划线位置切开，前半段保留、中间插下划线、后半段另起
        split_info = None  # (span_idx, front_text, back_text, front_bbox, back_bbox)
        for i, sp in enumerate(spans_list):
            sbbox = sp.get("bbox") or []
            if len(sbbox) < 4:
                continue
            sp_x0, sp_x1 = sbbox[0], sbbox[2]
            # 下划线落在此 span 内部（不是在 span 边界之外）
            if sp_x0 + 2 <= ul_x0 and ul_x1 <= sp_x1 - 2:
                content = sp.get("content", "")
                sp_w = sp_x1 - sp_x0
                if sp_w <= 0 or not content:
                    break
                # 按 x 比例拆分文本（等宽字体假设：字符宽度均匀）
                front_ratio = (ul_x0 - sp_x0) / sp_w
                n_chars = len(content)
                front_n = round(front_ratio * n_chars)
                front_text = content[:front_n]
                back_text = content[front_n:]
                # 拆分后的 bbox（前段到下划线起点，后段从下划线终点）
                front_bbox = [sp_x0, sbbox[1], ul_x0, sbbox[3]]
                back_bbox = [ul_x1, sbbox[1], sp_x1, sbbox[3]]
                split_info = (i, front_text, back_text, front_bbox, back_bbox)
                break

        if split_info is not None:
            # 拆分目标 span：原 span 缩为前半段，后半段作为新 span 插入到下划线后
            si, front_text, back_text, front_bbox, back_bbox = split_info
            orig_span = spans_list[si]
            # 前半段：复用原 span，改文本和 bbox
            orig_span["content"] = front_text
            orig_span["bbox"] = front_bbox
            # 插入虚拟下划线 span
            spans_list.insert(si + 1, virtual_span)
            # 后半段：克隆原 span 样式，新文本和 bbox
            if back_text:
                back_span = {
                    "bbox": back_bbox,
                    "content": back_text,
                    "type": orig_span.get("type", "text"),
                }
                if orig_span.get("_style"):
                    back_span["_style"] = dict(orig_span["_style"])
                spans_list.insert(si + 2, back_span)
        else:
            # 正常情况：下划线在 span 之间，按 x0 找插入位置
            insert_idx = len(spans_list)
            for i, sp in enumerate(spans_list):
                sp_x0 = (sp.get("bbox") or [0])[0]
                if ul_x0 < sp_x0:
                    insert_idx = i
                    break
            spans_list.insert(insert_idx, virtual_span)
        injected += 1

    return injected


def align_and_merge(
    mineru_data: dict[str, Any],
    spans: list[dict[str, Any]],
    *,
    iou_threshold: float = DEFAULT_IOU_THRESHOLD,
    underline_lines: list[dict[str, Any]] | None = None,
) -> tuple[dict[str, Any], dict[str, int]]:
    """
    将 PyMuPDF 的 span 样式对齐贴回 MinerU 的 block。

    参数：
        mineru_data:    MinerU middle.json 解析后的 dict
        spans:          parse_pymupdf.extract_spans 的输出
        iou_threshold:  IoU 匹配阈值（默认 0.3）

    返回：
        (merged_data, stats)
        - merged_data: 原 mineru_data（已 in-place 贴上 _style）
        - stats:       {"total": n, "matched": n, "unmatched": n, "no_bbox": n}

    处理逻辑：
        1. spans 按 page_idx 分组
        2. 遍历 MinerU 的每个 page 的 blocks
        3. 对每个 block：
           - 有 lines.spans 结构 → 逐 span 匹配（精细）
           - 无 lines.spans 结构 → block 整体匹配（兜底）
    """
    # 按 page_idx 分组 spans，加速查询
    spans_by_page: dict[int, list[dict[str, Any]]] = {}
    for s in spans:
        spans_by_page.setdefault(s["page_idx"], []).append(s)

    page_blocks = collect_page_blocks(mineru_data)
    stats = {"total": 0, "matched": 0, "unmatched": 0, "no_bbox": 0}

    for page_idx, blocks in page_blocks.items():
        page_spans = spans_by_page.get(page_idx, [])
        if not page_spans:
            continue

        for block in blocks:
            # 精细匹配：block 有 lines.spans 结构
            if block.get("lines"):
                # 跨页 block 检测（规则19）：
                # MinerU 把跨页段落的续行（坐标在下一页）也收进同一个 block 的 lines，
                # 但 block 所在页的 spans 里没有这些续行坐标 → IoU=0 匹配不上（无 _style）
                # 或 containment 误匹配到本页错误 span（字体串扰，如 SimHei 15.9pt）。
                # 检测判据：lines 的 y 跨度 > block bbox 高度 × 1.5
                #   正常 block：lines_span ≈ block_h
                #   跨页 block：续行 y0 是下一页坐标，lines_span ≈ 满页高（≈660pt）
                bbox = block.get("bbox") or []
                lines = block.get("lines") or []
                is_crosspage = False
                if bbox and len(bbox) >= 4 and lines:
                    block_h = bbox[3] - bbox[1]
                    if block_h > 0:
                        line_bboxes = [ln["bbox"] for ln in lines
                                       if ln.get("bbox") and len(ln["bbox"]) >= 4]
                        if line_bboxes:
                            lines_span = (max(b[3] for b in line_bboxes)
                                          - min(b[1] for b in line_bboxes))
                            is_crosspage = lines_span > block_h * 1.5

                if is_crosspage:
                    # 跨页：合并相邻页（page±1）的 spans，让续行 span 能匹配到正确页
                    # IoU 坐标唯一性保障：续行 bbox 只会与同坐标的 span 高 IoU，
                    # 不会因合并多页而误匹配（不同页同坐标的 span 极罕见，仅页眉页脚）
                    nearby_spans: list[dict[str, Any]] = []
                    for delta in (-1, 0, 1):
                        nearby_spans.extend(spans_by_page.get(page_idx + delta, []))
                    _attach_styles_to_block(block, nearby_spans, iou_threshold, stats)
                else:
                    _attach_styles_to_block(block, page_spans, iou_threshold, stats)
            # 兜底：block 无 lines 结构（title/image/table 等），
            # 用 block 整体 bbox 找占比最大样式
            else:
                _attach_style_to_block_fallback(block, page_spans)

    # B 类下划线注入：在样式贴完后，把填空区域的空白下划线作为虚拟 span 插入
    if underline_lines:
        injected = _inject_underline_spans(mineru_data, underline_lines)
        stats["underline_injected"] = injected

    return mineru_data, stats


# ═══════════════════════════════════════════════════════════════
#  持久化辅助
# ═══════════════════════════════════════════════════════════════

def save_merged(data: dict[str, Any], output_path: str | Path) -> str:
    """保存合并后的数据为 JSON。"""
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(
        json.dumps(data, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return str(output_path)


def load_mineru(path: str | Path) -> dict[str, Any]:
    """加载 MinerU middle.json。"""
    return json.loads(Path(path).read_text(encoding="utf-8"))


# ═══════════════════════════════════════════════════════════════
#  CLI（可独立运行测试）
# ═══════════════════════════════════════════════════════════════

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        description="bbox 对齐合并（MinerU 结构 + PyMuPDF 样式）"
    )
    parser.add_argument("middle", help="MinerU middle.json 路径")
    parser.add_argument("spans", help="PyMuPDF spans.json 路径")
    parser.add_argument("-o", "--output", default="./output/_merged.json",
                        help="输出 merged.json 路径")
    parser.add_argument("--iou", type=float, default=DEFAULT_IOU_THRESHOLD,
                        help=f"IoU 匹配阈值（默认 {DEFAULT_IOU_THRESHOLD}）")

    args = parser.parse_args()

    print(f"[3/4] bbox 对齐合并", file=sys.stderr)
    print(f"  → middle: {args.middle}", file=sys.stderr)
    print(f"  → spans:  {args.spans}", file=sys.stderr)

    mineru_data = load_mineru(args.middle)
    spans = json.loads(Path(args.spans).read_text(encoding="utf-8"))

    merged, stats = align_and_merge(mineru_data, spans, iou_threshold=args.iou)
    saved = save_merged(merged, Path(args.output) / "_merged.json")

    # 打印命中率
    print(f"\n  【对齐命中率统计】", file=sys.stderr)
    print(f"  总 span 数:    {stats['total']}", file=sys.stderr)
    print(f"  已贴样式:      {stats['matched']}", file=sys.stderr)
    print(f"  未匹配:        {stats['unmatched']}", file=sys.stderr)
    print(f"  无 bbox:       {stats['no_bbox']}", file=sys.stderr)
    if stats["total"] > 0:
        rate = stats["matched"] / stats["total"] * 100
        print(f"  命中率:        {rate:.1f}%", file=sys.stderr)
        if rate < 80:
            print(f"  ⚠️ 命中率低于 80%，建议调整 IoU 阈值或检查数据", file=sys.stderr)
    print(f"\n  → 保存到: {saved}", file=sys.stderr)

    result = {"merged_path": saved, "stats": stats}
    print(json.dumps(result, ensure_ascii=False, indent=2))
