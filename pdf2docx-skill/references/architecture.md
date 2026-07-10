# 架构原理 —— 四模块数据流与坐标对齐

> 本文档说明 PDF 转 DOCX 的四步管线设计原理。理解"为什么这么设计"时加载。

## 一、为什么用双数据源（MinerU + PyMuPDF）

单一工具都不完整：

| 单一工具 | 能给什么 | 缺什么 |
|----------|----------|--------|
| 仅 MinerU | 版面语义（标题层级、表格、图片、段落块类型） | 无字符级字号/字体/颜色（middle.json 的 span 只有 content/bbox） |
| 仅 PyMuPDF | 字符级样式（font/size/color/flags） | 无版面语义（要自己写规则判断"哪些大字粗体的是标题"，调参无止境） |

**两者数据互补**：MinerU 给"结构"，PyMuPDF 给"样式"，合并即得全量数据。

### 为什么不用其他方案

| 方案 | 问题 |
|------|------|
| Aspose.Words for Java | 不支持加载 PDF（已核实官方论坛，Java 版抛 UnsupportedFileFormatException） |
| pdf2docx | 维护者已建议用户迁移，黑盒一次转换不可控 |
| 仅 MinerU | 无字号字体颜色（middle.json 不携带） |
| 仅 PyMuPDF | 无版面语义，要自己写标题识别规则 |

## 二、为什么能对齐（坐标系统一）

这是双数据源方案成立的关键前提，已核实：

- **MinerU 底层就是用 PyMuPDF 做文本提取**（[MinerU issue #3867](https://github.com/opendatalab/MinerU/issues/3867)）
- 两者坐标系**完全一致**：
  - 单位都是 PDF 点（1/72 英寸）
  - 原点都是页面 top-left
  - Y 轴都向下
- 因此 MinerU block 的 bbox 与 PyMuPDF span 的 bbox **直接可比较**

对齐算法：对 MinerU 的每个 block，找出所有 bbox 与之相交的 PyMuPDF span，
将这些 span 的样式赋予该 block 的对应文本段。

## 三、四步管线数据流

### Step 1: MinerU 解析（parse_mineru.py）

```
输入: PDF
输出: middle.json（版面结构）+ images/（图片）

middle.json 结构（MinerU 2.0）:
{
  "pdf_info": [
    {
      "page_idx": 0,
      "para_blocks": [   // 或 blocks（VLM backend 可能用此名）
        {
          "type": "title",       // title/text/table/image/list
          "text_level": 1,        // 标题层级（仅 title 有）
          "bbox": [x0, y0, x1, y1],
          "lines": [
            { "spans": [
                { "content": "第一章", "bbox": [...] }
            ]}
          ]
        },
        {
          "type": "table",
          "table_body": "<table><tr><td>...</td></tr></table>",
          "bbox": [...]
        }
      ]
    }
  ]
}
```

> ⚠️ MinerU 2.0 的字段结构因 backend（pipeline/vlm/hybrid）而异。
> **首次使用前必须跑 pdf_inspect.py 确认真实字段**。

### Step 2: PyMuPDF 样式提取（parse_pymupdf.py）

```
输入: PDF
输出: spans.json（字符级样式，扁平列表）

每个 span:
{
  "page_idx": 0,
  "bbox": [x0, y0, x1, y1],     // PDF 点，与 MinerU 坐标系一致
  "text": "第一章",
  "font": "SimHei",              // 字体名
  "size": 18.0,                  // 字号（磅）
  "color": 0,                    // sRGB 整数，0=黑
  "flags": 16,                   // 位掩码
  "bold": true,                  // bit4(16)=粗体
  "italic": false                // bit1(2)=斜体
}
```

### Step 3: bbox 对齐合并（align.py）—— 核心不确定性

```
输入: middle.json + spans.json
输出: merged.json（原 middle.json，每个 span 多了 _style 字段）

对齐策略:
  对 MinerU 的每个 block:
    1. 按 page_idx 找出同页的 PyMuPDF spans
    2. 对 block 内每个 MinerU span，找 bbox 最匹配的 PyMuPDF span
    3. 用 IoU（交并比）衡量匹配度，阈值 0.3 起步
    4. 匹配上则贴 _style 字段

匹配策略（两级）:
  - 精细匹配：block 有 lines.spans 结构 → 逐 span IoU 匹配
  - 兜底匹配：block 无 lines 结构（title/image/table）→ block 整体 bbox 找占比最大样式
```

命中率统计：低于 80% 需调 IoU 阈值或检查数据。

### Step 4: DOCX 重建（build_docx.py）

```
输入: merged.json + images/
输出: DOCX

按 type 分发:
  title  → Heading 样式 + OutlineLevel + 字号字体
  text   → 普通段落 + 按 span 粒度设 run 样式
  table  → python-docx 表格（解析 HTML rowspan/colspan）
  image  → 插入图片
  list   → List Bullet 段落
```

## 四、为什么必须设 OutlineLevel（P0 关键）

下游 `splitDocumentByChapters` 章节分割的 `getOutlineLevel()` 强依赖段落的
`w:outlineLvl` 属性（`ChapterSplitServiceImpl.java:268/354/1485-1497`）。

`add_heading(text, level)` 会设 Heading 样式，但显式设 `w:outlineLvl` 是双保险。
python-docx 通过 `OxmlElement('w:outlineLvl')` + `qn('w:val')` 设置。

## 五、MinerU 3.x 同步 API（与原方案文档的差异）

原方案文档（MVP.md）基于 MinerU CLI 调用。实际开发中发现：

### 为什么不用 CLI / server 模式

MinerU 3.x 的 CLI 和 `mineru-api` server 都通过**子进程**运行。子进程
**不会继承** `MINERU_MODEL_SOURCE` 环境变量，导致国内环境默认连 HuggingFace
下载模型时**卡死**（worker 在等网络 I/O，CPU 不消耗，内存不增长，无报错）。

### 最终方案：同步 API 直接调用

`parse_mineru.py` 直接调用 pipeline backend 的 `doc_analyze_streaming()` 同步函数，
**在进程内**设置 `MINERU_MODEL_SOURCE=modelscope`，彻底绕开子进程环境变量问题。

| 项 | 原方案假设 | 最终实现 |
|----|------------|----------|
| 调用方式 | CLI subprocess | **同步 API**（`doc_analyze_streaming`） |
| 环境变量 | 外部设置 | 进程内 `os.environ` 设置（100% 生效） |
| 输出获取 | 下载 zip 解压 | **回调函数** `on_doc_ready` 直接拿 middle_json dict |
| middle.json 命名 | `{stem}_middle.json` | ✅ 一致 |

### 关键 API 签名

```python
doc_analyze_streaming(
    pdf_bytes_list,      # [pdf_bytes]
    image_writer_list,   # [ImageWriter 实例]
    lang_list,           # ["ch"]
    on_doc_ready,        # 回调：(doc_index, model_list, middle_json, ocr_enable)
    parse_method="auto", # auto/txt/ocr
    formula_enable=True,
    table_enable=True,
)
```

### 实测性能（88 页招标 PDF，CPU）

| 指标 | 结果 |
|------|------|
| 解析耗时 | 80-145 秒（视是否开公式/表格识别） |
| 对齐命中率 | **97.6%**（1351/1384 span） |
| DOCX block 数 | 697 |
| 图片提取 | 28 张 |
