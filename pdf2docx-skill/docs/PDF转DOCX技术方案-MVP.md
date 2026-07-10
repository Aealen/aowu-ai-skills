# PDF 转 DOCX 技术方案（MVP）

> **目标**：在标书生成系统入口处，将用户上传的 PDF 招标文件转换为 DOCX，转换完成后下游完全对齐现有 DOCX 处理流程，无需任何改动。
>
> **本文档定位**：技术可行性 MVP 实施方案，供先实现一个最小可用 Demo 验证转换质量。
>
> - 创建日期：2026-07-10
> - 适用范围：PDF 转 DOCX 的转换环节（入口前置），不涉及下游解析逻辑改造
> - 前置核实：所有"下游依赖"结论均经源码核实（见附录 B）

---

## 一、方案定位与边界

### 1.1 核心目标

```
用户上传 PDF 招标文件
        │
        ▼
  [PDF → DOCX 转换]   ← 本方案只做这一步
        │
        ▼
  转换出的 DOCX 进入现有流程（与用户直接传 .docx 走同一条路）
        │
        ▼
  search-agent 解析 / 章节分割 / 模板提取 ... （完全不改动）
```

**转换的唯一验收标准**：产出的 DOCX 越接近"用户原本会手动用 Word 另存的那份 DOCX"越好。保真度越高，下游越不容易出问题。

### 1.2 为什么要"保真还原"而非"凑一个能跑的 docx"

下游环节（search-agent 智能体、`splitDocumentByChapters` 章节分割等）是按"原生 DOCX"设计的。转换产出的 DOCX 与原生 DOCX 越接近，下游天然越能正确处理。因此转换必须尽量保真——原文 14 号字就还原 14 号，原文表格什么样就还原什么样，而不是用默认值凑一个结构上能跑的文档。

> ⚠️ 这意味着字号、字体、粗体、颜色**都属于转换保真的范畴**，需要还原。

### 1.3 不在 MVP 范围内

| 不做的事 | 原因 |
|---|---|
| 改动下游任何 Java 代码 | 转换是入口前置环节，转换完对齐现有流程即可 |
| 集成进 Java 主项目 | MVP 纯 Python 验证转换质量；最终落地形态（Java 重建 / Python 微服务 / 沙箱）待 MVP 通过后另行决策 |
| search-agent 适配 | search-agent 吃的是 DOCX，转换质量达标即可，与 PDF 转换无关 |
| 扫描件 PDF（OCR） | 用户明确先不考虑扫描件，只处理文字版 PDF |

---

## 二、技术选型与依据

### 2.1 为什么用 MinerU + PyMuPDF

**两个工具职责分工明确，数据互补**：

| 工具 | 职责 | 提供的数据 | 为什么需要它 |
|---|---|---|---|
| **MinerU** | 版面结构分析 | 标题层级（`text_level` 1/2/3）、块类型（title/text/table/image/list）、表格 HTML（含合并单元格）、图片位置 + bbox、段落层级 | AI 模型驱动的版面理解，识别"这块是表格""那块是 H2 标题"，开源成熟 |
| **PyMuPDF** | 字符级样式提取 | 每个 span 的 font（字体名）、size（字号）、color（颜色）、flags（粗体/斜体位标记）+ bbox | MinerU 的 `middle.json` 不携带字号字体（已核实，见附录 A.2），必须额外取 |

**核心要点**：MinerU 给"版面结构"，PyMuPDF 给"字符样式"，两者用 bbox 坐标对齐合并，即得全量数据。

### 2.2 两者数据是否可对齐（已核实）

✅ **可直接对齐，无需坐标转换。**

经核实（[PyMuPDF Appendix 3](https://pymupdf.readthedocs.io/en/latest/app3.html)、[MinerU issue #3867](https://github.com/opendatalab/MinerU/issues/3867)）：

- **MinerU 底层就是用 PyMuPDF 做文本提取**
- 两者坐标系**完全一致**：单位都是 PDF 点（1/72 英寸），原点都是页面 top-left，Y 轴都向下
- 因此 MinerU block 的 bbox 与 PyMuPDF span 的 bbox **直接可比较**，用矩形相交判断即可把 PyMuPDF 的样式"贴回"到 MinerU 对应的 block

> 对齐算法：对 MinerU 的每个 block，找出所有 bbox 与之相交（IoU 或包含关系）的 PyMuPDF span，将这些 span 的样式赋予该 block 的对应文本段。

### 2.3 为什么不用单一工具

| 单一工具 | 缺什么 |
|---|---|
| 仅 MinerU | 无字符级字号/字体/颜色（`middle.json` 的 span 只有 content/bbox，无 font/size） |
| 仅 PyMuPDF | 无版面语义（要自己写规则判断"哪些字号大+粗体的段落是标题"，调参无止境；表格/图片识别要从零写） |
| Aspose.Words for Java | 不支持加载 PDF（已核实官方论坛，Java 版抛 `UnsupportedFileFormatException`，PDF 导入仅在 .NET/Python 版） |
| pdf2docx | 维护者已在 PyPI 建议用户迁移走，长期风险；黑盒一次转换不可控 |

---

## 三、MVP 架构

### 3.1 处理流程

```
真实招标 PDF
    │
    ├──[1] MinerU CLI 解析
    │      mineru -p input.pdf -o output/ -m auto -l zh
    │      → output/input_middle.json   （版面结构）
    │      → output/input.md            （Markdown，参考用）
    │      → output/images/             （抽取的图片）
    │
    ├──[2] PyMuPDF 解析
    │      fitz.open(input.pdf)
    │      page.get_text("dict") → blocks → lines → spans
    │      每个 span: { text, font, size, color, flags, bbox }
    │      → spans.json （字符级样式，含 page_idx + bbox）
    │
    ├──[3] 数据对齐合并
    │      遍历 MinerU middle.json 的 para_blocks
    │      对每个 block，用 bbox 相交找出 PyMuPDF 对应 spans
    │      合并：block 结构 + span 样式
    │      → merged_blocks.json
    │
    └──[4] python-docx 重建
           遍历 merged_blocks，按 type 分发：
             title  → Heading 样式 + OutlineLevel + 字号字体
             text   → 普通段落 + 字号字体
             table  → 表格（解析 HTML rowspan/colspan）
             image  → 插入图片
           → output/重建结果.docx
```

### 3.2 目录结构

```
pdf-to-docx-mvp/
├── input/                # 放真实招标 PDF（选 2-3 份不同排版风格）
├── output/               # 中间产物 + 最终 docx
│   ├── _middle.json      # MinerU 输出
│   ├── _spans.json       # PyMuPDF 输出
│   ├── _merged.json      # 对齐合并结果（调试用）
│   └── 重建结果.docx
├── src/
│   ├── parse_mineru.py   # [1] 调 MinerU CLI（薄封装）
│   ├── parse_pymupdf.py  # [2] PyMuPDF 取字符级样式
│   ├── align.py          # [3] bbox 对齐合并（核心）
│   ├── build_docx.py     # [4] python-docx 重建（核心）
│   └── inspect.py        # 工具：检查 middle.json/spans.json 字段
├── run.py                # 一键串联 1→2→3→4
└── requirements.txt
```

### 3.3 依赖

```
# requirements.txt
mineru[all]          # 含模型，pip install -U "mineru[all]"
PyMuPDF              # fitz，字符级样式提取
python-docx          # docx 重建
```

环境：Python 3.10+。首次需运行 `mineru-models-download` 下载模型权重。

> **Windows 坑提醒**：[GitHub Issue #4433](https://github.com/opendatalab/MinerU/issues/4433) 指出 Windows 上 pip 装 MinerU 可能不生成 CLI 入口。如踩到，用官方 Docker 镜像跑 MinerU（脚本逻辑不受影响）。

---

## 四、各模块实现要点

### 4.1 `parse_mineru.py` — MinerU 解析（薄封装）

MVP 阶段直接调 CLI，不封装 Python API。CLI 输出 middle.json 到磁盘，便于检查。

```python
import subprocess
from pathlib import Path

def parse_with_mineru(pdf_path: str, output_dir: str) -> str:
    """
    调用 MinerU CLI 解析 PDF。
    返回生成的 middle.json 路径。
    """
    cmd = [
        "mineru", "-p", pdf_path,
        "-o", output_dir,
        "-m", "auto",    # auto: 自动判断是否需 OCR（文字版走 pipeline）
        "-l", "zh",      # 中文
    ]
    subprocess.run(cmd, check=True)

    # MinerU 输出文件名 = 输入文件名去后缀 + _middle.json
    stem = Path(pdf_path).stem
    return str(Path(output_dir) / stem / f"{stem}_middle.json")
```

**关键参数说明**：
- `-m auto`：文字版 PDF 走 pipeline（快、准），扫描版自动走 OCR（MVP 虽不主攻扫描件，但保留能力）
- `-l zh`：中文优化
- 输出结构：`output/<文件名>/<文件名>_middle.json` + `_content_list.json` + `images/`

### 4.2 `parse_pymupdf.py` — 字符级样式提取

PyMuPDF 的 `page.get_text("dict")` 返回 `blocks → lines → spans` 三级结构，每个 span 携带完整样式。

```python
import fitz  # PyMuPDF
import json

def extract_spans(pdf_path: str) -> list[dict]:
    """
    提取每个 span 的字符级样式。
    返回扁平 span 列表，每项含 page_idx + bbox + 样式。
    """
    doc = fitz.open(pdf_path)
    all_spans = []

    for page_idx, page in enumerate(doc):
        page_dict = page.get_text("dict")
        page_height = page.rect.height   # 对齐时备用

        for block in page_dict.get("blocks", []):
            if block.get("type") != 0:   # type 0=文本，type 1=图片（跳过）
                continue
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    flags = span["flags"]
                    all_spans.append({
                        "page_idx": page_idx,
                        "bbox": span["bbox"],           # (x0, y0, x1, y1) PDF点，top-left原点
                        "text": span["text"],
                        "font": span["font"],           # 字体名，如 "SimSun"、"Helvetica-Bold"
                        "size": round(span["size"], 1), # 字号（磅）
                        "color": span["color"],         # sRGB 整数，0=黑
                        "flags": flags,
                        "bold": bool(flags & (1 << 4)),     # bit4=粗体
                        "italic": bool(flags & (1 << 1)),   # bit1=斜体
                    })
    doc.close()
    return all_spans
```

**flags 位掩码说明**（[PyMuPDF Appendix 1](https://pymupdf.readthedocs.io/en/latest/app1.html)）：

| bit | 值 | 含义 |
|---|---|---|
| 0 | 1 | 上标 superscript |
| 1 | 2 | 斜体 italic |
| 2 | 4 | 衬线 serifed |
| 3 | 8 | 等宽 monospaced |
| 4 | 16 | **粗体 bold** |

**span 定义**（官方原文）："a span consists of adjacent characters with identical font properties: name, size, flags, and color" —— 相同样式的连续字符聚成一个 span。

> **颜色补充**：PyMuPDF 直接给 `color` 字段（sRGB 整数），无需像 PDFBox 那样自建颜色状态机。这是选 PyMuPDF 不选 PDFBox 的关键优势之一。

### 4.3 `align.py` — bbox 对齐合并（核心）

这是双数据源方案的核心不确定性所在，MVP 重点验证。

```python
def bbox_overlap(b1: tuple, b2: tuple) -> bool:
    """判断两个 bbox（x0,y0,x1,y1）是否相交。"""
    x0 = max(b1[0], b2[0])
    y0 = max(b1[1], b2[1])
    x1 = min(b1[2], b2[2])
    y1 = min(b1[3], b2[3])
    return x0 < x1 and y0 < y1


def align_blocks(mineru_blocks: list[dict], spans: list[dict]) -> list[dict]:
    """
    将 PyMuPDF 的 span 样式对齐贴回 MinerU 的 block。
    mineru_blocks: middle.json 的 pdf_info[].para_blocks
    spans: parse_pymupdf 的输出
    """
    # 按 page_idx 分组 spans，加速查询
    spans_by_page = {}
    for s in spans:
        spans_by_page.setdefault(s["page_idx"], []).append(s)

    for block in mineru_blocks:
        page_idx = block.get("page_idx", 0)
        block_bbox = block["bbox"]
        page_spans = spans_by_page.get(page_idx, [])

        # 找出落在本 block 范围内的 spans
        matched = [s for s in page_spans if bbox_overlap(block_bbox, s["bbox"])]

        # 将 matched spans 按 MinerU 的 lines → spans 结构重新组织
        # （MinerU span 有 bbox 和 content，用坐标匹配把 PyMuPDF 样式贴上去）
        _attach_styles_to_block(block, matched)

    return mineru_blocks


def _attach_styles_to_block(block: dict, matched_spans: list[dict]):
    """
    把 PyMuPDF 样式贴到 MinerU block 的 lines.spans 上。
    MinerU 的 span 有 bbox + content，PyMuPDF 的 span 有 bbox + 样式，
    用 bbox 相交把样式贴过去。
    """
    for line in block.get("lines", []):
        for mspan in line.get("spans", []):
            m_bbox = mspan.get("bbox")
            if not m_bbox:
                continue
            # 找最佳匹配（IoU 最大或包含关系）
            best = None
            best_iou = 0
            for ps in matched_spans:
                iou = _calc_iou(m_bbox, ps["bbox"])
                if iou > best_iou:
                    best_iou = iou
                    best = ps
            if best and best_iou > 0.3:   # 阈值可调
                mspan["_style"] = {
                    "font": best["font"],
                    "size": best["size"],
                    "bold": best["bold"],
                    "italic": best["italic"],
                    "color": best["color"],
                }
```

**对齐策略说明**：
- MinerU 与 PyMuPDF 坐标系一致（都是 PDF 点、top-left、y-down），**无需坐标转换**
- 采用 IoU（交并比）匹配，阈值 0.3 起步（MVP 调参项）
- 边界情况：跨 block 的 span、文本框浮动定位需特殊处理（MVP 先观察日志）

### 4.4 `build_docx.py` — DOCX 重建（核心）

这是工作量所在。遍历合并后的 blocks，按 type 分发重建。

```python
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from pathlib import Path


def build_docx(merged_blocks_by_page: dict, images_dir: str, output_path: str):
    doc = Document()

    for page_idx in sorted(merged_blocks_by_page.keys()):
        for block in merged_blocks_by_page[page_idx]:
            btype = block.get("type", "text")
            if btype == "title":
                _build_title(doc, block)
            elif btype == "table":
                _build_table(doc, block)
            elif btype == "image":
                _build_image(doc, block, images_dir)
            elif btype == "list":
                _build_list(doc, block)
            else:  # text / 其他
                _build_text(doc, block)

    doc.save(output_path)


def _build_title(doc, block):
    """标题：Heading 样式 + OutlineLevel + 字号字体。"""
    level = block.get("text_level", 1)
    level = min(max(level, 1), 9)   # 限制 1-9

    text = _extract_text(block)
    para = doc.add_heading(text, level=level)

    # 补充字符样式（字号字体粗体）
    style = _get_dominant_style(block)   # 取该 block 内占比最大的样式
    if style:
        for run in para.runs:
            _apply_style(run, style)

    # 显式设置 OutlineLevel（双保险，确保下游 splitDocumentByChapters 能识别）
    _set_outline_level(para, level)


def _build_text(doc, block):
    """正文段落。"""
    text = _extract_text(block)
    if not text.strip():
        return

    para = doc.add_paragraph()
    # 按 MinerU 的 lines.spans 粒度设置 run 样式（同段不同样式）
    for line in block.get("lines", []):
        for mspan in line.get("spans", []):
            run = para.add_run(mspan.get("content", ""))
            _apply_style(run, mspan.get("_style"))


def _build_table(doc, block):
    """
    表格：解析 table_body（HTML，含 rowspan/colspan）→ python-docx 表格。
    MinerU 的 table_body 是 HTML 格式，可用 html.parser 或 lxml 解析。
    """
    table_html = block.get("table_body", "")
    if not table_html:
        return
    rows = _parse_html_table(table_html)   # 返回 [[{text, rowspan, colspan}, ...], ...]
    if not rows:
        return

    # 计算实际列数（考虑 colspan）
    max_cols = max(sum(c.get("colspan", 1) for c in row) for row in rows)
    table = doc.add_table(rows=len(rows), cols=max_cols)
    table.style = "Table Grid"

    _fill_table_cells(table, rows)   # 处理 rowspan/colspan 合并


def _build_image(doc, block, images_dir):
    """图片：从 images/ 读图插入。MVP 按顺序插入，位置精度后续优化。"""
    img_path = block.get("img_path", "")
    if img_path:
        full_path = Path(images_dir) / Path(img_path).name
        if full_path.exists():
            doc.add_picture(str(full_path), width=Inches(5.5))


def _build_list(doc, block):
    """列表：MVP 降级为普通段落（带编号文本）。"""
    for line in block.get("lines", []):
        text = "".join(s.get("content", "") for s in line.get("spans", []))
        if text.strip():
            doc.add_paragraph(text, style="List Bullet")
```

**关键辅助函数**：

```python
def _set_outline_level(paragraph, level: int):
    """
    显式设置段落的 w:outlineLvl。
    双保险：add_heading 已设 Heading 样式，这里再显式设 outlineLvl，
    确保下游 splitDocumentByChapters 的 getOutlineLevel() 能读到。
    （依据：ChapterSplitServiceImpl.java:268/354/1485-1497 强依赖 OutlineLevel）
    """
    pPr = paragraph._p.get_or_add_pPr()
    # 先移除已存在的 outlineLvl
    for existing in pPr.findall(qn('w:outlineLvl')):
        pPr.remove(existing)
    outline = OxmlElement('w:outlineLvl')
    outline.set(qn('w:val'), str(level))
    pPr.append(outline)


def _apply_style(run, style: dict):
    """把 PyMuPDF 样式应用到 python-docx run。"""
    if not style:
        return
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
        r = (color >> 16) & 0xFF
        g = (color >> 8) & 0xFF
        b = color & 0xFF
        run.font.color.rgb = RGBColor(r, g, b)
    font = style.get("font")
    if font:
        # PDF 字体名映射到常用中文字体（MVP 简化）
        font_name = _map_font_name(font)
        run.font.name = font_name
        # 中文字体需设 EastAsia
        rPr = run._element.get_or_add_rPr()
        rFonts = rPr.find(qn('w:rFonts'))
        if rFonts is None:
            rFonts = OxmlElement('w:rFonts')
            rPr.append(rFonts)
        rFonts.set(qn('w:eastAsia'), font_name)


def _map_font_name(pdf_font: str) -> str:
    """PDF 字体名 → 常用字体名映射（MVP 简化，后续扩充）。"""
    name = pdf_font.lower()
    if "simsun" in name or "song" in name:
        return "宋体"
    if "simhei" in name or "hei" in name:
        return "黑体"
    if "kaiti" in name or "kai" in name:
        return "楷体"
    if "fangsong" in name or "fs" in name:
        return "仿宋"
    return pdf_font   # 兜底原样保留
```

### 4.5 `inspect.py` — 字段检查工具（先于 build 跑）

**这个脚本必须最先写、最先跑**。原因：middle.json 和 PyMuPDF dict 的真实字段结构是方案地基，官方文档描述到块级，细节必须实测确认。在写 build_docx.py 之前，先 inspect 一份真实输出。

```python
def inspect_middle_json(path: str):
    """检查 MinerU middle.json 的真实字段结构。"""
    data = json.load(open(path, encoding="utf-8"))
    pdf_info = data.get("pdf_info", [])

    print(f"总页数: {len(pdf_info)}")
    for page in pdf_info[:2]:   # 只看前2页
        blocks = page.get("para_blocks", [])
        type_counter = {}
        for b in blocks:
            t = b.get("type", "unknown")
            type_counter[t] = type_counter.get(t, 0) + 1
        print(f"\npage_idx={page.get('page_idx')} 块类型分布: {type_counter}")

        # 打印各类型样例
        for b in blocks:
            if b.get("type") == "title":
                print(f"  [title] level={b.get('text_level')} text={_preview(b)[:40]}")
                break
        for b in blocks:
            if b.get("type") == "table":
                print(f"  [table] table_body前80字符: {str(b.get('table_body',''))[:80]}")
                break
        for b in blocks:
            if b.get("type") == "image":
                print(f"  [image] img_path={b.get('img_path')} bbox={b.get('bbox')}")
                break
        # 打印一个 text block 的完整 lines/spans 结构
        for b in blocks:
            if b.get("type") == "text":
                print(f"  [text] 完整结构: {json.dumps(b, ensure_ascii=False)[:300]}")
                break


def inspect_pymupdf_spans(pdf_path: str):
    """检查 PyMuPDF span 的真实样式字段。"""
    import fitz
    doc = fitz.open(pdf_path)
    page = doc[0]
    d = page.get_text("dict")
    count = 0
    for block in d.get("blocks", []):
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                print(f"font={span['font']} size={span['size']} "
                      f"color={span['color']} flags={span['flags']} "
                      f"text={span['text'][:20]}")
                count += 1
                if count >= 10:
                    break
        if count >= 10:
            break
    doc.close()
```

---

## 五、实施步骤

### 5.1 推荐执行顺序

```
Step 0：环境搭建
  ├─ pip install -U "mineru[all]" PyMuPDF python-docx
  ├─ mineru-models-download（下载模型权重）
  └─ 准备 2-3 份真实招标 PDF（不同排版风格）放 input/

Step 1：inspect 先行（必须最先做）
  ├─ 跑 MinerU 解析一份 PDF，看 middle.json 真实结构
  ├─ 跑 PyMuPDF dict，看 span 真实样式字段
  └─ 确认字段与预期是否一致（不一致先调整 align/build 逻辑）

Step 2：实现对齐 align.py
  ├─ 用 inspect 确认的字段结构写对齐逻辑
  ├─ 跑一份，输出 merged.json
  └─ 人工检查对齐命中率（PyMuPDF span 贴回 MinerU block 的成功率）

Step 3：实现重建 build_docx.py
  ├─ 按优先级实现：title → text → table → image → list
  ├─ 每类 block 实现后立即用真实 PDF 验证
  └─ 产出 docx，人工对比原 PDF

Step 4：验收
  ├─ 对比转换 docx 与原 PDF 的还原度（见 5.2）
  └─ 若达标 → MVP 通过；若不达标 → 定位问题（MinerU 解析差？对齐命中率低？重建丢信息？）
```

### 5.2 验收标准

**核心标准**：转换出的 docx 与原 PDF 对比，保真度越高越好。

| 优先级 | 验收项 | 验证方式 |
|---|---|---|
| **P0** | 标题层级正确：第X章/1.1/1.1.1 在 docx 里是 Heading 1/2/3 + 对应 OutlineLevel | Word 大纲视图看层级树是否正确 |
| **P0** | 正文段落完整：无大面积文本丢失或乱序 | 对比原 PDF 逐段核对 |
| **P0** | 评分表/资质表还原为表格：行列结构基本对齐，合并单元格保留 | 对比原 PDF 表格 |
| **P0** | 字号字体基本还原：标题大字粗体、正文小字 | 肉眼对比 docx 与原 PDF |
| P1 | 图片位置基本正确（顺序/大致位置） | 人工看 |
| P1 | 颜色基本还原（非纯黑白的重要文字） | 人工看 |
| P2 | 分栏排版阅读顺序正确 | 对比原 PDF |

**P0 是 MVP 必须达成的**。若 P0 任一不达标，需定位根因：

| 不达标现象 | 可能根因 | 应对 |
|---|---|---|
| 标题层级错乱 | MinerU text_level 不准 | 加规则修正（结合正则"第X章"） |
| 文本丢失/乱序 | MinerU 版面解析遗漏 或 分栏串行 | 换 PDF 样本对比，评估是普遍还是个例 |
| 表格崩 | MinerU 表格识别失败 或 HTML 解析 bug | 单独调试 _build_table |
| 对齐命中率低 | bbox 匹配阈值/算法问题 | 调整 align.py 的 IoU 阈值 |

---

## 六、关键风险点

### 6.1 表格复杂度（最高风险）

招标文件的评分表、资质要求表常有**合并单元格、跨页表、嵌套表**。

- MinerU 的 `table_body` 是 HTML 格式，理论上含 rowspan/colspan
- 但**复杂表格的实际识别率必须重点测**——这是 MinerU 对真实招标 PDF 解析能力的最大考验
- 若评分表解析崩，影响 scoreItem 解析（但已核实 `scoreItemParseFlag` 智能体模式下不阻塞目录，降级影响有限）

### 6.2 对齐命中率（核心不确定性）

PyMuPDF span 贴回 MinerU block 的成功率，决定样式还原质量。

- 绝大多数情况 bbox 相交可靠
- 边界情况：跨 block 的 span、浮动文本框、竖排文字
- MVP 重点观察对齐日志（命中率统计），低于 80% 需调算法

### 6.3 text_level 可靠性

MinerU 给的标题层级是 AI 推断，不一定 100% 准。

- 若层级乱，章节分割会受影响
- MVP 应对：结合文本正则（`^第[一二三四五六七八九十]+章`、`^\d+\.\d+`）修正 text_level
- 修正优先级：正则命中 > MinerU text_level

### 6.4 页眉页脚干扰

招标文件页眉常有"XX招标文件"。MinerU 一般能过滤，但要确认没把页眉误判为标题。

### 6.5 Windows 环境坑

MinerU 在 Windows 上 pip 安装可能不生成 CLI 入口（[Issue #4433](https://github.com/opendatalab/MinerU/issues/4433)）。应对：用 Docker 跑 MinerU。

---

## 七、最终落地形态预告（MVP 通过后决策）

MVP 纯 Python 验证转换质量。**最终落地形态待 MVP 通过后另行决策**，有几个方向：

| 方向 | 说明 | 适用场景 |
|---|---|---|
| **A. 重建逻辑迁 Java POI** | MinerU 解析留 Python（CLI/微服务），docx 重建用主项目 POI | 主项目保持纯净，重建与下游同语言 |
| **B. Python 微服务** | MinerU + PyMuPDF + python-docx 全在独立服务，Java HTTP 调用 | 转换整体外置，Java 零 Python 依赖 |
| **C. 沙箱执行** | 转换脚本打包交由另一系统的沙箱执行 | 用户提到有沙箱能力（属另一系统） |

**MVP 阶段不用纠结落地形态**——python-docx 重建逻辑与 POI 重建逻辑是等价的（都操作 OOXML），迁移成本可控。

---

## 附录 A：关键事实核实记录

### A.1 MinerU middle.json 不携带字号字体（已核实）

MinerU 的 span 字段定义（[官方文档](https://opendatalab.github.io/MinerU/reference/output_files/)）只有 `bbox` / `type` / `content` / `score`，**无 font / size / bold 字段**。这是 MinerU 定位（RAG/知识库版面理解）决定的。`middle.json` 保留版面结构（para_blocks/lines/spans + bbox），但不保留字符级排版属性。

### A.2 PyMuPDF 能取字符级样式（已核实）

`page.get_text("dict")` 的 span 含 `font` / `size` / `color` / `flags` / `bbox`（[Appendix 1](https://pymupdf.readthedocs.io/en/latest/app1.html)）。flags 位掩码：bit4(16)=bold, bit1(2)=italic。颜色直接给 sRGB 整数，无需像 PDFBox 自建颜色状态机。

### A.3 MinerU 与 PyMuPDF 坐标系一致（已核实）

两者都是 PDF 点单位、top-left 原点、Y 轴向下（[Appendix 3](https://pymupdf.readthedocs.io/en/latest/app3.html)、[MinerU #3867](https://github.com/opendatalab/MinerU/issues/3867)）。bbox 直接可比较，无需坐标转换。MinerU 底层即用 PyMuPDF 做文本提取。

### A.4 python-docx 设 OutlineLevel（已核实）

OutlineLevel 是段落的 `w:outlineLvl` XML 属性（[OOXML 规范](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.outlinelevel)）。python-docx 通过 `OxmlElement('w:outlineLvl')` + `qn('w:val')` 设置（[python-docx #746](https://github.com/python-openxml/python-docx/issues/746)）。`add_heading(text, level)` 会设 Heading 样式，但显式设 outlineLvl 是双保险。

### A.5 Aspose.Words for Java 不能加载 PDF（已核实）

官方论坛（2025-11-14，针对 25.9 版本）确认：Aspose.Words for Java 不支持加载 PDF，PDF 导入仅在 .NET/Python 版本实现。Java 版抛 `UnsupportedFileFormatException: Pdf format is not supported on this platform`。项目用的 24.12 版本同理。

---

## 附录 B：下游依赖核实结论（源码级，仅供理解"为什么转换要保真"）

> 本附录记录转换质量对下游的影响分析。MVP 实现时无需关心，转换只管保真还原即可。

| 下游环节 | 对 docx 的依赖 | 转换需保证 |
|---|---|---|
| search-agent 智能体 | URL 主路径自取完整 docx，返回扁平标题字符串列表，`titleLevel` 本地硬编码 1 | docx 文本完整即可 |
| `splitDocumentByChapters` 章节分割 | 强依赖 `OutlineLevel` + Heading 样式名（`ChapterSplitServiceImpl.java:268/354/1485-1497`）；字号/粗体仅作弱加分 | **必须设 Heading 样式 + OutlineLevel** |
| `DirectoryTitleSourceLookup` 目录回查 | 纯 `getText()`，零样式依赖（`:214`） | 文本完整即可 |
| `doParseChapterTemplate` 章节模板 | `DocxReader` 转 Markdown 喂 AI，丢字号字体（`DocxReader.java:65`） | 文本完整即可 |
| `extractStyleTemplate` 样式模板 | 读 styles/numbering 定义，但缺失时 fallback 内置 `myTitleFormat.docx`（`FileButtonOperationServiceImpl.java:797`） | 可选，不阻塞 |

**结论**：转换保真度越高越好。最硬的保真要求是**标题层级（Heading + OutlineLevel）**和**表格结构**，这两项直接影响章节分割和评分项解析。字号字体粗体颜色作为保真的一部分尽量还原。

---

## 文档修订记录

| 版本 | 日期 | 作者 | 修订内容 |
|---|---|---|---|
| v1.0 | 2026-07-10 | ZCode | 初版。基于多轮源码核实与技术调研，确定 MinerU + PyMuPDF 双数据源方案，含完整 MVP 实施步骤、代码骨架、验收标准。 |
