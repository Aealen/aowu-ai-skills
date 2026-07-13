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
- 行距：从 PDF 逐 block 测量 `_measure_block_line_height`，兜底用文档中位数
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

### 规则 2：居中对齐——逐行检查，不看 block 整体

**判断条件**（三个都满足才算居中行）：
1. 行中心接近页面中心（容差 < 页宽 5%）
2. 行宽 < 内容区宽度 × 90%（排除全宽正文）
3. 实际 x0 接近居中期望 x0（`expected_x0 = page_center - line_width/2`，容差 < 页宽 5%）

**多数行居中 → 整个 block 居中**（`centered_lines >= total_lines / 2`）

**为什么不用 block 整体 bbox**：
- 多行 block 中第一行可能占满全宽（如封面标题），拉宽 block bbox
- 宽正文段落 bbox 中心天然接近页面中心，但实际是左对齐
- 居中行应有双侧留白且 x0 远离左边距，左对齐行 x0 接近左边距

### 规则 3：缩进——分层 left_indent + first_line_indent

| 参数 | 计算方式 | 含义 |
|------|----------|------|
| `left_indent` | 所有行中最小 x0 − 左边距 | 续行的起始位置（所有行的基准） |
| `first_line_indent` | 首行 x0 − 续行最小 x0 | 首行相对续行的额外偏移 |

**为什么分两层**：
- 编号列表（如"6.本项目采用..."）首行有缩进、续行贴边 → `left_indent=0, first_line_indent=23pt`
- 联系信息（如"采购人：新疆..."）所有行统一缩进 → `left_indent=23pt, first_line_indent=0`
- 不分层会导致多行 block 首行溢出或缩进不一致

### 规则 4：行距——逐 block 从 PDF 测量，文档中位数兜底

| 层级 | 来源 | 适用场景 |
|------|------|----------|
| 优先 | `block["_line_height"]` | 有 PDF 测量值的 block |
| 兜底 | `block["_doc_line_height"]` | 无测量值时用全文档中位数 |
| 极端 | 固定值（22pt） | 全文档无任何行高数据 |

**行高测量范围**：基于 block 内主导字号动态计算（`size × 0.8` ~ `size × 3.5`），不写死 12-45pt。

### 规则 5：页面边距——用 PyMuPDF 精确提取，不裁剪不覆盖

- 用 `_detect_page_margins` 从 PDF 提取上下左右边距（采样多页取中位数）
- 直接用提取值，**不加 `max(X, 28)` 裁剪**
- 边距设好后**不被后续代码覆盖**（之前的"内容边距修正"代码会覆盖精确值为不准确的统计值）

### 规则 6：标题分类——结构标题 vs 视觉标题

| 类型 | 判断条件 | 格式处理 |
|------|----------|----------|
| 结构标题 | PDF 大纲（`get_toc`）中有对应条目 | Heading 样式 + OutlineLevel + TOC 绑定 |
| 视觉标题 | MinerU 判为 title 但不在 PDF 大纲中 | 普通段落 + 视觉强调（字号/加粗），不绑定 TOC |

**不使用正则猜标题层级**。标题层级完全由 PDF 大纲的 level 字段决定。

### 规则 7：表格列宽——从 PyMuPDF 表格对象提取

- 用 `page.find_tables()` 获取 PyMuPDF Table 对象
- 从 `table.cells` 的 x 坐标边界计算列宽
- 通过 OOXML `tblGrid` + `tblLayout type=fixed` 设置（不是 `cell.width`，LibreOffice 忽略后者）

### 规则 8：跨页表格错位修复——两层策略

| 策略 | 条件 | 判断方式 |
|------|------|----------|
| 策略 1 | PyMuPDF 数据覆盖的行 | PyMuPDF col0 为空但 HTML 有文本 → 错位 |
| 策略 2 | PyMuPDF 未覆盖的行 | col0 超 20 字符**或含中文标点** → 错位 |

**关键约束**：
- 只对 `ri < len(pymupdf_cells)` 的行用策略 1
- 策略 2 中"含中文标点"比"正则匹配编号格式"更通用——条款号不会含句号/逗号

### 规则 9：图片尺寸——从 block bbox 提取

- `width_in = (bbox[2] - bbox[0]) / 72`
- 加上限保护：不超过页面内容宽度
- 无 bbox 时回退到内容宽度

### 规则 10：HTML 实体——用标准库，不全手写

- 用 `html.unescape()` 替代手写 `re.sub(r"&nbsp;", ...)` 逐个替换
- 避免遗漏 `&lt;`、`&gt;`、`&quot;` 等实体

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
