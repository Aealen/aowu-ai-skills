# PDF 转 DOCX 格式还原规范共识

> 本文档总结了在开发 PDF 转 DOCX 通用转换工具过程中积累的格式还原规则。
> 所有规则均遵循一个核心原则：**从 PDF 实际数据提取格式，不做任何硬编码假设。**

---

## 核心原则

### 1. 一切格式从 PDF 数据提取，不写死任何值

**禁止的做法**：
- 写死页面尺寸（如 595×842 = A4）
- 写死边距（如 left=82pt）
- 写死行距（如 24pt）
- 写死缩进（如 first_line_indent=24pt）
- 写死字号/字体名映射（如 FangSong→仿宋）
- 写死图片宽度（如 5.5 英寸）
- 写死过滤范围（如行高 12-45pt）

**正确的做法**：
- 页面尺寸：从 MinerU `page_size` 或 PyMuPDF `page.rect` 获取
- 边距：用 PyMuPDF 精确提取 `_detect_page_margins`
- 行距：从 block bbox 高度 ÷ MinerU 行数 直接提取（`_measure_block_line_height`）
- 缩进：从 span bbox 的 x0 坐标差值计算
- 字体：保留 PDF 原始字体名，由 Word/LibreOffice 自动匹配
- 图片：从 block bbox 提取原始尺寸
- 过滤范围：基于实际字号动态计算（如 `dominant_size * 0.8` ~ `dominant_size * 3.5`）

### 2. 不用正则匹配文本内容做格式判断

**禁止的做法**：
- 用正则判断标题层级（如 `^第[一二三四五六七八九十]+章` → Heading 1）
- 用正则判断条款号（如 `^\d+\.\d+` → label）
- 用正则判断表格列错位（如 `_looks_like_label` 函数）

**正确的做法**：
- 标题层级：用 PDF 大纲（`get_toc`）的坐标匹配确定
- 表格错位：用 PyMuPDF 单元格坐标或文本长度/标点特征判断

### 3. 不维护字体名白名单

**禁止的做法**：
- 维护 `_FONT_MAP` 把字体名映射到系统字体名
- 维护 `_CID_BOLD_FONTS` 白名单判断哪些字体天生粗体

**正确的做法**：
- 保留 PDF 原始字体名写入 DOCX
- 粗体检测用 PDF 原生信号（flags + linewidth + 像素密度），不靠字体名猜

---

## 具体规则

### 规则 1：粗体检测——三层信号，优先 PDF 原生数据

| 优先级 | 信号 | 来源 | 适用场景 |
|--------|------|------|----------|
| 1 | flags bit4 | PyMuPDF `span["flags"]` | 非 CID 字体（英文 Times-Bold 等） |
| 2 | linewidth > 0 | PyMuPDF `get_texttrace()` 的 `linewidth` 字段 | CID 字体伪粗体（PDF 用 Tr=2 描边+填充） |
| 3 | 像素密度 > 同字体基线 × 1.3 | 渲染页面灰度图采样 | CID 字体同字体内部对比 |

**关键约束**：
- 像素密度检测**只在同字体内部对比**（如 FangSong-Bold vs FangSong-Regular）
- **不跨字体对比**（如 SimHei vs FangSong）——不同字体的笔画密度天生不同，跨字体对比必然误判
- 像素密度基线用**文档中出现最多的字号**（众数字号 = 正文），不用所有字体的中位数
- 像素密度检测是**文档级**的：跨页累积数据，即使封面页 span 少也能用正文页的基线

### 规则 2：居中对齐——逐行检查 + 留白对称约束

**判断条件**（四个都满足才算居中行）：
1. 行中心接近页面中心（容差 < 页宽 5%）
2. 行宽 < 内容区宽度 × 90%（排除全宽正文）
3. 实际 x0 接近居中期望 x0（`expected_x0 = page_center - line_width/2`，容差 < 页宽 5%）
4. **左右留白对称且显著**（`min(left_gap, right_gap) > 内容区宽度 × 8%` 且 `min/max > 0.4`）

**多数行居中 → 整个 block 居中**（`centered_lines >= total_lines / 2`）

**为什么不用 block 整体 bbox**：
- 多行 block 中第一行可能占满全宽（如封面标题），拉宽 block bbox
- 宽正文段落 bbox 中心天然接近页面中心，但实际是左对齐
- 居中行应有双侧留白且 x0 远离左边距，左对齐行 x0 接近左边距

**为什么需要第4个条件（留白对称约束）**（曾导致的问题）：
- 三条件阈值偏宽，"占内容区 60%-90% 宽度且中心恰好接近页中心"的左对齐编号条目
  （如"1.1 采购人：..."占 87% 宽）被误判为居中 → 41 个条目错误设为居中对齐
- 真居中行（标题"第三章"）左右隙 ≈145pt 且对称（min/max ≈ 1.0）；
  误判的编号条目左右隙仅 20-30pt（首行缩进量，不是居中留白）
- 第4条件要求留白"显著（>8%内容区）且对称（min/max>0.4）"，精准区分两者
- 验证：收紧后真居中 title 零误伤，58 个被误判的编号条目回归左对齐

### 规则 3：缩进——分层 left_indent + first_line_indent + 单行块首行缩进识别

| 参数 | 计算方式 | 含义 |
|------|----------|------|
| `left_indent` | 所有行中最小 x0 − 左边距 | 续行的起始位置（所有行的基准） |
| `first_line_indent` | 首行 x0 − 续行最小 x0 | 首行相对续行的额外偏移 |

**为什么分两层**：
- 编号列表（如"6.本项目采用..."）首行有缩进、续行贴边 → `left_indent=0, first_line_indent=23pt`
- 联系信息（如"采购人：新疆..."）所有行统一缩进 → `left_indent=23pt, first_line_indent=0`
- 不分层会导致多行 block 首行溢出或缩进不一致

**单行 block 的首行缩进识别**（坐标对齐法，符合核心原则2）：
- 单行 block（`min_x0 == first_x0`）无法从自身区分"首行缩进条目"与"整段缩进"
- 用**同页多行 block 的首行 x0** 作基准（`_page_first_line_x0`，取众数，聚类到 3pt）：
  - 多行 block 的首行 x0 可靠地反映"首行缩进"位置（续行 x0 < 首行 x0 有缩进差）
  - 单行 block 的 first_x0 若与同页基准对齐（±5pt 容差），即为首行缩进条目
  - 设为 `first_line_indent = first_x0 - page_x0`，`left_indent = 0`
- 例：page15 多行 block（1.2/1.3/1.4）首行 x0=113 → 基准=114；
  单行 block（1.1/1.3.3/1.5.1）first_x0=112-113 匹配基准 → 统一首行缩进 26pt
- **纯坐标对齐，不读文本内容**——不依赖正则判断 `^\d+\.\d+` 编号格式

### 规则 4：行距——block bbox 高度 / 行数，精确还原

每个 block 的 `line_spacing` = **block bbox 高度 ÷ MinerU 识别行数**。

**原理**：`line_spacing × 行数 = bbox 高度`（零误差），配合 `space_after = gap_after`，
每个 block 的垂直占用 = bbox 高 + gap = 原 PDF 中该 block 的精确垂直空间。

| block 类型 | line_spacing 计算 | 示例 |
|------------|-------------------|------|
| 单行 | bbox 高度 / 1 | bbox=14pt → lh=14pt（字符高度） |
| 多行 | bbox 高度 / 行数 | bbox=63pt, 3行 → lh=21pt → DOCX: 21×3=63pt ✓ |

**禁止的做法**：
- 用 y0 差值（baseline 间距）作为行高——它 ≠ DOCX 的 EXACTLY 行高语义，导致多行段落总高度偏大
- 用文档/同页中位数作为回退——不同段落行距各异，中位数是"猜"而非"取"
- 单行 block 无法测行高 → 回退中位数（应直接用 bbox 高度）

**兜底**：仅在 bbox 数据无效时用文档中位数（极端情况，正常 PDF 覆盖率 100%）。

### 规则 5：页面边距——用 PyMuPDF 精确提取，排除页眉页脚，合并表格 bbox

- 用 `_detect_page_margins` 从 PDF 提取上下左右边距
- **必须排除页眉/页脚**：`ly0 > 页高 - 60`（页脚）或 `ly0 < 60`（页眉）的行不参与统计
- 直接用提取值，**不加 `max(X, 28)` 裁剪**
- 边距设好后**不被后续代码覆盖**

**为什么必须排除页脚**（曾导致的问题）：
- 页脚"—1—"固定在 y1=794.7pt，如果不排除，`max(y1s)` 取到页脚坐标
- 下边距被低估为 47pt（页脚到页底），而真实正文下边距约 84pt
- 内容区被错误放大 37pt ≈ 每页多放 2 行，随页数累积越来越严重

**四边边距都取 `min()`，且都要合并表格 bbox**：
- `min(top_margins)` / `min(bottom_margins)` / `min(left_margins)` / `min(right_margins)`
- 语义：内容最靠四边的页反映真实可用内容区。内容区必须能容纳**最靠边的内容**，
  否则那一页（如表格页）会内容溢出。
- **表格 bbox 必须纳入上下左右边距统计**——表格边框线比 cell 内文字 span 更靠外
  （cell 有内边距）。只看文字 span 会低估表格页的真实占用范围。
  - 左右：不合并 → 内容区比表格窄，cell 被 Word/WPS 在 fixed 布局下强制压缩、文字换行
  - 上下：不合并 → 内容区高度比 PDF 实际小，atLeast 行高的表格在 WPS 下
    因剩余空间不足把整行推到下一页（见下方"上边距取中位数"反模式）

**为什么上边距也用 `min()` 而非中位数**（曾导致的问题）：
- 表格页的表格边框顶到 y=72pt，但同一页文字 span 顶在 y=80pt（cell 内边距+首行文字）
- 上边距取中位数（80pt）会丢失"表格从 72pt 开始"的信息，内容区高度比 PDF 实际小 7.8pt
- 当某页表格总高恰好填满 PDF 内容区（如 1.20-1.24 五行=690pt=PDF内容区），
  DOCX 内容区只有 682pt，WPS 把放不下的最后一行（1.24）整体推到下一页
- LibreOffice 分页较宽松能放下，但 WPS 严格执行 → 渲染引擎差异暴露了边距偏小

### 规则 6：标题分类——结构标题 vs 视觉标题

| 类型 | 判断条件 | 格式处理 |
|------|----------|----------|
| 结构标题 | PDF 大纲（`get_toc`）中有对应条目 | Heading 样式 + OutlineLevel + TOC 绑定 |
| 视觉标题 | MinerU 判为 title 但不在 PDF 大纲中 | 普通段落 + 视觉强调（字号/加粗），不绑定 TOC |

**不使用正则猜标题层级**。标题层级完全由 PDF 大纲的 level 字段决定。

### 规则 7：表格列宽——从 PyMuPDF 表格对象提取

- 用 `page.find_tables()` 获取 PyMuPDF Table 对象
- 从 `table.cells` 的 x 坐标边界计算列宽
- 通过 OOXML `tblGrid` + `tblLayout type=fixed` + `tblW type=auto` 设置
  （`tblW=dxa` 固定值可能因舍入误差被 Word 压缩，`auto` 更可靠）
- 列宽乘以 1.014 补偿系数（见规则 17）

### 规则 8：表格行高与 cell 行距——atLeast 模式，不裁剪内容

trHeight（OOXML `trHeight`）用 `atLeast` 模式——行高至少为 PDF 值，
DOCX 渲染时字体宽度差异可能导致实际行数多于 PDF（见规则 17），
atLeast 允许行高自适应增长，避免底部内容被裁剪。

**实现要点**：
- **先设 trHeight，再调 `_fill_table_cells`**——cell 填充时需读取 trHeight 做安全约束
- line_spacing 安全约束：`safe_lh = min(文字间距, trHeight / (cell文字行数+0.5))`
- row_h 微缩 2%（`scaled_h = row_h * 0.98`），trHeight 和 line_spacing 统一用此值

**禁止的做法**：
- 用 block bbox 高度 / 1 作为 table block 的 line_height（= 表格总高 600pt）
- 用文档中位数作为表格行距（与正文行距混淆）
- `exact` 模式（字体宽度差异导致实际行数多于预期时裁剪底部内容）

### 规则 9：跨页表格数据合并——首页主 block 收集全部 PyMuPDF 数据

MinerU 把跨页表格的全部行放在第一个 table block 的 HTML 中，
但 PyMuPDF 表格数据（每行高度、cell 文本含换行、cell 样式）分散在各页。

**必须合并**：主 block（有 HTML）向后扫描续页 block（无 HTML），
收集各页的 `_pymupdf_cells` / `_table_row_heights` / `_table_row_line_heights` / `_table_cell_styles`。

续页 block 标记 `_table_merged = True`，`_build_table` 跳过避免重复渲染。

**验证**：合并后 `_table_row_heights` 行数应等于 HTML 行数（如 33 行），
否则 trHeight 匹配失败，回退为不设行高。

### 规则 10：MinerU 漏检表格——用 PyMuPDF find_tables 交叉校验

MinerU 有时把整个表格误判为多个 text block（如 14 个编号段落），
导致 DOCX 中丢失表格结构。

**交叉校验**：对每页，如果 MinerU 没有 table block 但 PyMuPDF `find_tables()` 检测到表格，
从 PyMuPDF 数据重建 table block，替换落入表格 bbox 内的 text block。

### 规则 11：表格 cell 文字样式——从 PyMuPDF 提取，不用 block._style 统一

每个 cell 的字体/字号/粗体应从 PDF dict 中该 cell bbox 内的 span 数据提取，
不能用 block._style 一种样式覆盖所有 cell。

**实现**：`_extract_table_cell_styles` 按 cell bbox 收集 span 样式，
`_fill_table_cells` 用每个 cell 的主导字体（首个 span）替代 block._style。
- col0 条款号（如 "1.1"）→ TimesNewRomanPSMT
- col1 中文内容 → FangSong
- cell 文本中的 `\n`（PyMuPDF extract() 保留）→ 用 `run.add_break()` 还原换行

### 规则 12：跨页表格错位修复——两层策略

| 策略 | 条件 | 判断方式 |
|------|------|----------|
| 策略 1 | PyMuPDF 数据覆盖的行 | PyMuPDF col0 为空但 HTML 有文本 → 错位 |
| 策略 2 | PyMuPDF 未覆盖的行 | col0 超 20 字符**或含中文标点** → 错位 |

**关键约束**：
- 只对 `ri < len(pymupdf_cells)` 的行用策略 1
- 策略 2 中"含中文标点"比"正则匹配编号格式"更通用——条款号不会含句号/逗号

### 规则 13：跨页表格——直接用 PyMuPDF 原始行结构重建，不合并续行

MinerU HTML 是 AI 推理结果，跨页大单元格常被错误拆成多余 `<tr>`
（如 1.15 条目跨页，MinerU 拆成3行 + 文本错位到 col0 + OCR 文本重复）。
PyMuPDF `extract()` 是 PDF 结构直接提取，文字 100% 准确。

**核心原则：复刻 PDF，而非绘制新表格**。
跨页大单元格在 PDF 中物理上就是分段的（如 1.15 在第8页底部67pt +
第9页顶部432pt），PyMuPDF 正确地把它们识别为两行（续行 col0 为空）。
直接用这个原始结构重建表格，每行用 PDF 真实行高（exact 模式），
Word 会自然分页——主行在上一页底部、续行在下一页顶部，完美复刻 PDF。

**触发条件**：block 标记为 `_table_crosspage`（跨页表格主 block）且为 2 列。
多列表格（评分表等）不走此路径，避免误伤 rowspan 结构。

**禁止的做法**：
- **合并续行**——合并后行高累加超过单页（1.15→498pt），Word 的 exact 模式
  不允许行内分页，整行被推到下一页（1.15另起一页）；即使不设 trHeight，
  cell 内 23 行文本被压缩在极矮区域（safe_lh 索引错位 bug → line_spacing=1pt）
- 对多列表格（4列+）走此路径——其 col0 空行可能是 rowspan 合并单元格
- 盲信 MinerU HTML 行结构——AI 推理有 OCR 错误和续行拆分错误

### 规则 14：复杂合并表格——HTML 含 colspan/rowspan 时走 HTML 路径，不走 PyMuPDF

PyMuPDF `find_tables()` 对**含 colspan/rowspan 的复杂表格** cell 边界检测不可靠：
会把每个 cell 的 bbox 都报告成跨满表格宽度，导致 `_infer_colspan_rowspan` 把
所有 cell 都误判为 colspan=2（如第14页表格被全部合并成一列，文本错乱混合）。

**MinerU 的 AI 视觉对此类表格的合并结构识别更可靠**——HTML 准确标注了哪些
cell 是 colspan=2、哪些是 colspan=1。

**路径选择策略**（`_build_table` 由 `_html_has_merged_cells` 决定）：
| 表格特征 | 路径 | 原因 |
|---------|------|------|
| HTML 含 colspan>1 或 rowspan>1 | **HTML 路径** | PyMuPDF cell 边界对复杂表格不可靠，HTML 结构更准 |
| HTML 不含合并（简单表格） | **PyMuPDF 路径**（默认） | 避免 MinerU 的 OCR 错误/续行拆分 |

**HTML 路径仍用 PyMuPDF 增强文本**：`_enrich_rows_with_pymupdf_text` 用 PyMuPDF
精确文本替换 HTML 的 OCR 文本（结构用 HTML，文字用 PyMuPDF，各取所长）。

**判别信号可靠性**：跨页条款表（1.1-1.28）的 HTML 全是 colspan=1（简单表格），
不会被误判为复杂表格。全文仅 5 个表格（第14/54/56/70/83页）含合并单元格。

### 规则 15：图片尺寸——从 block bbox 提取

- `width_in = (bbox[2] - bbox[0]) / 72`
- 加上限保护：不超过页面内容宽度
- 无 bbox 时回退到内容宽度

### 规则 16：HTML 实体——用标准库，不全手写

- 用 `html.unescape()` 替代手写 `re.sub(r"&nbsp;", ...)` 逐个替换
- 避免遗漏 `&lt;`、`&gt;`、`&quot;` 等实体

### 规则 17：表格列宽字体度量补偿——×1.014 系数

Windows 安装的 FangSong 和 TimesNewRoman 字符宽度比 PDF 内嵌字体平均宽约 **1.4%**
（实测比例：FangSong=1.0143，TimesNewRoman=1.0142，SimHei=0.9506）。
不补偿会导致接近 cell 宽度的行（如33字=396pt）超出 cell 列宽（395pt）触发换行。

**实现**：在 `_set_table_fixed_width` 中，gridCol 列宽统一乘以 1.014：
```python
col_widths_dxa = [round(w * 20 * 1.014) for w in col_widths_pt]
```

**为什么对所有列统一补偿**：
- FangSong 和 TimesNewRoman 比例几乎相同（1.014），文档主要用这两种字体
- SimHei 比例 0.95（更窄），多补偿只会让它更宽松，不会换行
- 逐字体按比例补偿过于复杂且收益有限

**验证**：补偿后 cell1 从 395pt→400.6pt，33字 FangSong（396pt）能完整放下一行。

### 规则 18：跨页文本段落——拆分 block + 插分页符，复刻 PDF 分页

MinerU 把**跨页段落**的所有行（含下一页续行）都塞进同一个 block 的 `lines` 数组，
但 block 的 `bbox` 只覆盖**本页部分**的高度。

**核心目标**：复刻 PDF 的原始分页——PDF 中段落跨两页（本页底部 N 行 + 下页顶部 M 行），
DOCX 中必须保持这个跨页状态，而非让 Word 自行决定分页。

**问题演进（三轮修复）**：

**第一轮**（已废弃）：`bbox高 / len(lines)` 用全部行数做除数 → 极小行高（15pt/5行=3pt），
文字被 EXACTLY 行距压扁。

**第二轮**（已废弃）：用续行 y0 差值（24.7pt）做行高。行距虽准，但 5行×24.7pt=123.5pt
远超 bbox 的 15pt，段落总高度暴涨，把后续内容推后 1-2 页。

**根因**：DOCX 的行框模型（总高=lh×行数）≠ PDF 固定坐标模型（被分页符切断）。
一个连续的 DOCX 段落无法表达"被分页符切成两段，各自有不同高度"。

**最终方案**（`_split_crosspage_blocks`，本规则）：**按 PDF 页边界拆分 block + 插分页符**。

数据模式（13 个跨页 text block，全部遵循）：
- 跨页 block 在 page_idx=N，bbox 只覆盖本页部分，`lines` 含全部行
  （本页行 y0 在 bbox 内 + 续页行 y0 突然变小≈80，是下一页坐标）
- **下一页（page_idx=N+1）的 block#0 一定是空占位 block**：
  `type=text, lines=[], lines_deleted=true, bbox 精确覆盖续行区域`

拆分流程（`_split_crosspage_blocks`）：
1. 检测跨页 block（`lines y 跨度 > bbox 高 × 1.5`）
2. 本页行 = y0 落在 bbox 范围内（容差 5pt）的 line；续页行 = 其余
3. **原 block 只保留本页行**（lines 截断），bbox 不变
4. **下一页空占位 block#0 被替换为续页 block**：塞入续页行，继承占位 block 的 bbox
   （已精确覆盖续行区域），标记 `_crosspage_continuation=True`
5. 主循环遍历到 `_crosspage_continuation` block 时，**先插分页符**（`_add_page_break`）

拆分后每个部分都是独立 block，bbox 准确覆盖自己的行 → `bbox高/行数` 直接算出正确行高：

| block | 本页部分 | 续页部分 | PDF真实行距 |
|-------|---------|---------|------------|
| `1.5.2` | bbox高15/1行 = **15pt** ✓ | bbox高92/4行 = **23pt** ✓ | 25.0pt |
| page38可研 | bbox高359/13行 = **27.6pt** ✓ | bbox高301/11行 = **27.4pt** ✓ | 28.0pt |
| pi=3 | bbox高63/3行 = **21pt** ✓ | bbox高16/1行 = **16pt** ✓ | — |

**检测判据**：`lines 跨度 > block bbox 高度 × 1.5`
- 正常 block：lines_span ≈ block_h（行紧贴 bbox 内）
- 跨页 block：lines_span ≈ 660pt（接近满页高，续行坐标是下一页的），
  block_h 仅 15-360pt，比值 ≥ 1.89 → 阈值 1.5 有充分余量，不误判正常 block

**为何分页符比调行高可靠**：分页符是硬分页，所有渲染引擎（WPS/Word/LibreOffice）
都严格遵守，不依赖行高猜测。而调行高是在猜 Word 会怎么分页——不同引擎规则不同，不可靠。

**注意**：
- align.py 在 build_docx.py 之前运行，读的是原始未拆分数据，跨页匹配逻辑（规则19）不受影响
- 拆分后续行 line 携带的 `_style`（align 阶段附加）跟着 line 对象走，无需额外处理
- 只拆 `text`/`paragraph`/`list` 类型，`table` 跨页用独立的 `_table_merged` 机制（规则9）

### 规则 19：跨页 block 续行 span 样式——邻页 spans 匹配 + 主导样式兜底

跨页 block 的续行 span 坐标在**下一页**，但 `align.py` 默认用 block 所在页的 spans 匹配。

**问题**：
- 用本页 spans 匹配下一页坐标的续行 → IoU≈0 匹配不上 → span 无 `_style` → DOCX run 无字体（继承 Word 默认）
- 或 containment 误匹配到本页错误 span → 字体串扰（如 FangSong 12pt 续行被串扰成 SimHei 15.9pt）

**修复（两层）**：

1. **align.py 跨页匹配**：检测到跨页 block（`lines 跨度 > bbox 高 × 1.5`，同规则18判据）时，
   合并相邻页（page_idx±1）的 spans 供匹配。IoU 坐标唯一性保障不会误匹配（续行 bbox 只会与
   同坐标的 span 高 IoU，跨页同坐标 span 极罕见）。

2. **build_docx.py 主导样式兜底**：对仍无 `_style` 的 span（匹配失败的残留），用 block 首个有效
   span 的样式兜底。在 `_build_text` 和 `_build_index` 中均实现。

| 数据源 | 跨页续行样式 | 可靠性 |
|--------|-------------|--------|
| PyMuPDF（PDF 结构直接提取） | 准确（FangSong 12pt） | ★★★★★ |
| MinerU _merged.json（align 本页匹配） | 错误/缺失 | ★ |
| MinerU _merged.json（align 邻页匹配） | 准确 | ★★★★★ |

**验证**：1.5.2 全文统一 FangSong 12pt（修复前 run1 是 SimHei 15.5pt、run2-4 无字体）；
13 个跨页 block 续行样式全部正确。

---

## 反模式清单（禁止再次出现）

| 反模式 | 曾导致的问题 | 正确做法 |
|--------|-------------|----------|
| 写死 `Pt(24)` 行距 | 不同 PDF 行距差异大 | 从 PDF 测量 |
| 写死 `Inches(5.5)` 图片宽度 | 图片变形 | 从 bbox 提取 |
| 写死 `_FONT_MAP` 字体映射 | 非中文 Windows 字体全丢失 | 保留原始字体名 |
| 用正则 `_RE_CHAPTER` 判断标题 | 无法适配不同文档格式 | 用 PDF 大纲 |
| 用正则 `_looks_like_label` 判断条款号 | 误判评分表短内容 | 用坐标/长度/标点判断 |
| 像素密度跨字体对比 | SimHei 常规被判为粗体 | 同字体内部对比 |
| 边距加 `max(X, 28)` 裁剪 | 精确值被放大 | 直接用提取值 |
| `font_baseline` 取最小字号 | 脚注密度当正文基线 | 取众数字号 |
| 变量名 `html` 遮蔽 `html` 模块 | `html.unescape()` 报错 | 重命名为 `table_html` |
| `_guess_heading_level` 死代码 | 文档与实现脱节 | 删除并更新文档 |
| 行高用 y0 差值（baseline 间距） | 多行段落总高度偏大 | 用 bbox 高度/行数 |
| 跨页段落塞进单个 DOCX 段落 | DOCX 行框模型(lh×行数)≠PDF固定坐标，行高怎么调都回归（3pt压扁/24.7pt推后）；1.5.2被推到下页 | 拆分 block + 插分页符，复刻 PDF 分页（规则18） |
| 跨页段落用续行y0差值做行高 | 单block行距准(24.7pt)但5行总高123pt远超bbox 15pt，内容连锁后移 | 拆分后各部分用 bbox高/行数，总高=PDF |
| 行高用文档/同页中位数兜底 | 单行 block 行距被"猜"而非"取" | 用 bbox 高度（100%覆盖） |
| 页边距统计含页脚行 | 下边距 47pt（页脚位）而非 84pt（正文位），每页多放2行 | 排除页眉页脚后统计 |
| 正文段落 `space_after` 写死为 0 | 段间距丢失，每页少 ~150pt | 用 `_gap_after` 设置 |
| table block 用 bbox高/1 做行高 | 600pt 行高，每行占满一页 | table 类型跳过行高提取 |
| cell 统一用 block._style | 条款号 "1.1" 显示为 FangSong 而非 TimesNewRomanPSMT | 按 cell 从 PDF 提取主导字体 |
| trHeight 和 line_spacing 不同源 | line_spacing × 行数 > trHeight，exact 模式内容裁剪 | 统一用缩放后的 row_h |
| 先填 cell 再设 trHeight | 安全约束读不到 trHeight，溢出检测失效 | 先设 trHeight 再填 cell |
| `atLeast` 行高模式 | ~~内容溢出时行高膨胀~~ → 实际是 exact 模式裁剪内容 | 统一用 `atLeast`，允许行高自适应 |
| 跨页续页 block 不标记 | 同一表格被渲染两次（94页） | 续页标记 `_table_merged` |
| MinerU 漏检表格不补救 | 整页表格变散落文本段落 | PyMuPDF find_tables 交叉校验重建 |
| cell 行间距含跨列 y0 | 不同列文字的 y0 差被误算为行间距（13.2pt） | 按列分组后算同列内 y0 差值 |
| 盲信 MinerU HTML 行结构 | 跨页续行被拆成多余行 + OCR 文本错位重复（"商自行承担。应商自行承担。"） | PyMuPDF 为权威基准，直接用原始行结构重建 |
| 合并跨页续行 | 行高累加超单页（498pt），exact 模式整行推到下一页；safe_lh 索引错位致 line_spacing=1pt | 不合并，保留 PyMuPDF 原始分页行结构 |
| `_fill_table_cells` 的 row_tr_heights 跳过无 trHeight 行 | 索引错位，行N读到行N+1的 trHeight，safe_lh 算出 1pt 行距 | row_tr_heights 与 table.rows 一一对应（含 None） |
| 用行数差异触发 PyMuPDF 重建 | 跨页合并后 HTML 和 PyMuPDF 行数可能恰好接近（33 vs 33） | 用 `_table_crosspage` 标记触发 |
| 简单表格用 MinerU HTML 重建 | AI 推理有 OCR 错误、续行拆分、文本错位 | 简单表格用 PyMuPDF 重建路径（文字100%准确） |
| 含合并单元格的复杂表格用 PyMuPDF 重建 | PyMuPDF 把每个 cell bbox 都报成跨满宽，_infer_colspan_rowspan 全判 colspan=2，表格被压成一列 | HTML 含 colspan/rowspan 时走 HTML 路径（结构用HTML+文字用PyMuPDF增强） |
| PyMuPDF 全量适用假设 | 跨页简单表格 PyMuPDF 更准，但含 colspan/rowspan 复杂表格 HTML 更准 | 按 HTML 合并单元格特征分路径（规则14） |
| `_enrich_rows_with_pymupdf_text` 不检查 None | 无 PyMuPDF 数据的表格 block 传入 None 触发 `len(None)` 崩溃 | 函数开头 `if not pymupdf_cells: return` |
| Table Grid 样式的默认 cell margin | 左右各5.4pt 吃掉 cell 宽度导致换行 | 不用 TableGrid 样式，手动加边框 + tcMar=0 |
| span 间空格未去除 | PyMuPDF 分隔符空格在 CJK 字体下占12pt，每行多出12pt | 渲染 span 时 `sp_text.strip()` |
| 表格列宽不加字体度量补偿 | Windows FangSong 比 PDF 内嵌宽1.4%，接近行宽的文字超出cell换行 | gridCol × 1.014 补偿系数 |
| 上边距取中位数 + 不合并表格bbox顶部 | 表格边框顶72pt但文字span在80pt，内容区比PDF小7.8pt；填满PDF的表格页在WPS下最后一行被推到下一页 | 四边边距都取min()且合并表格bbox |
| 只合并表格bbox的左右边界，忽略上下边界 | 表格页内容区高度不足，atLeast表格行被WPS整体推到下页 | 上下左右都合并表格bbox |
| 跨页block用全部lines数算行高 | bbox只覆盖本页(15pt)但lines含续页行(5行)→3pt行距，文字被EXACTLY压扁挤一行（1.5.2/可研正文8处）| 检测lines跨度>bbox高×1.5，优先用续行y0差值，回退本页行数(规则18) |
| 跨页block用本页bbox高/本页行数（本页行≤2） | 算出首行字符高度(15pt)而非行间距(25pt)，续行被压扁、分页错位（1.5.2/1.5.7/3.成交单位3处）| 续行≥2时用续行y0差值均值（精确行间距）(规则18第二层) |
| 居中判断无留白对称约束 | 占87%宽的左对齐编号条目(1.1/5.6/12.3)中心恰好接近页中心被误判居中，41个条目对齐错误 | 增加第4条件：min(左右隙)>内容区8%且对称(规则2) |
| 单行block一律设left_indent | 单行编号条目(1.1/1.3.3/1.5.1)无法区分首行缩进与整段缩进，设成整段缩进与多行同类条目不一致 | 用同页多行block首行x0基准做坐标对齐(规则3) |
| 跨页block续行用本页spans匹配 | 续行坐标在下一页，本页IoU=0匹配不上→无样式(继承Word默认)；或containment误匹配→字体串扰(SimHei15.9pt) | 检测跨页后合并邻页spans匹配 + 无样式span用主导样式兜底(规则19) |

---

## 数据源优先级

当多个数据源可用时，优先级从高到低：

```
PDF 原生数据（PyMuPDF 直接提取）
  > PDF 渲染数据（像素密度等需要渲染的信号）
    > MinerU 结构数据（block type / level / bbox）
      > 统计推断（众数 / 中位数 / P25）
        > 固定兜底值（仅在以上全部失败时）
```

**示例**：
- 粗体：`flags` (PDF原生) > `linewidth` (PDF原生) > 像素密度 (渲染) > 不加粗
- 边距：`_detect_page_margins` (PDF原生) > bbox P25 统计 > 82pt 兜底
- 标题层级：`get_toc` level (PDF原生) > MinerU `text_level` > 默认 level 1
