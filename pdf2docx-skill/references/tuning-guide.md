# 调参指引 —— IoU 阈值 / 字体映射 / 表格解析

> 本文档汇总需要在真实 PDF 上实测调优的参数。调优转换质量时加载。

## 一、bbox 对齐 IoU 阈值

**位置**：`align.py` 的 `DEFAULT_IOU_THRESHOLD`，CLI 参数 `--iou`

**默认值**：0.3（起步值，未经验证）

### 调优方法

1. **跑一份真实 PDF，看命中率统计**：
   ```bash
   python3 pdf2docx.py convert input.pdf -o output.docx --keep-work
   ```
   输出会打印：`命中率 XX.X% (matched/total)`

2. **根据命中率调整**：
   - 命中率 < 80%：阈值可能偏高，尝试降低（0.3 → 0.2）
   - 命中率 > 95% 但样式错乱：阈值可能偏低导致误匹配，尝试升高（0.3 → 0.4）
   - 用 `pdf_inspect.py merged` 检查贴的样式是否正确

3. **边界情况**（需观察日志）：
   - 跨 block 的 span：一个 PyMuPDF span 横跨两个 MinerU block
   - 浮动文本框：定位坐标偏离常规
   - 竖排文字：bbox 匹配可能失效

### 匹配策略说明

当前实现两级匹配：
- **IoU（交并比）**：主策略，衡量两个 bbox 的重叠程度
- **Containment（包含比）**：兜底，处理 MinerU span 比 PyMuPDF span 大的情况

如果命中率持续不理想，可考虑：
- 改用"中心点距离"匹配（对偏移更鲁棒）
- 按 line 级别先聚合，再按 span 级别精匹配

## 二、字体映射表

**位置**：`build_docx.py` 的 `_FONT_MAP` 字典 + `_map_font_name()` 函数

**当前覆盖**（已根据真实招标 PDF 实测补充）：
- 正文类：宋体(SimSun)、仿宋(FangSong)、楷体(KaiTi)
- 标题类：黑体(SimHei)、方正小标宋(FZXBSJW，封面大标题常见)
- 华文系列：华文宋体/黑体/楷体/仿宋
- 方正系列：方正书宋/黑体/楷体/仿宋

### 真实招标 PDF 字体分布（实测样本：新疆磋商文件 88 页）

| 字体名 | 占比 | 用途 | 映射状态 |
|--------|------|------|----------|
| FangSong | 65% | 正文主体（公文标准） | ✅ → 仿宋 |
| TimesNewRomanPSMT | 20% | 英文/数字 | 兜底原样保留 |
| SimSun | 10% | 部分正文 | ✅ → 宋体 |
| TimesNewRomanPS-BoldMT | 2.6% | 加粗英文/数字 | flags bold 位正确识别 |
| SimHei | 0.7% | 标题/重点 | ✅ → 黑体 + 粗体推断 |
| KaiTi | 0.7% | 附注/说明 | ✅ → 楷体 |
| FZXBSJW | 少量 | 封面大标题 | ✅ → 方正小标宋 + 粗体推断 |

### 扩充方法

1. **用 pdf_inspect.py 查看真实 PDF 的字体名**：
   ```bash
   python3 pdf_inspect.py spans input.pdf
   # Windows uv: uv run python scripts/pdf_inspect.py spans input.pdf
   ```
   输出的 `font` 列就是 PDF 内嵌的字体名

2. **把遇到的字体名加到 `_FONT_MAP`**：
   ```python
   _FONT_MAP = {
       # 已有...
       "新字体名小写": "中文名",
   }
   ```

3. **字体名匹配规则**：用 `in` 子串匹配（不区分大小写），所以 `"simsun"` 能匹配
   `"SimSun"`、`"SimSun-Bold"`、`"SimSun-Regular"` 等。

### 粗体判断增强（重要）

实测发现：**招标文件标题不靠 flags bold 位标记粗体**。方正小标宋(FZXBSJW)、黑体(SimHei)
等字体本身是粗字面，flags 只有衬线位(4)没有粗体位(16)。

`parse_pymupdf.py` 的 `_is_bold()` 函数做两级判断：
1. flags 位标记（bit4=16）—— 标准方式，识别 `TimesNewRomanPS-BoldMT` 这类
2. 字体名推断（`_BOLD_FONT_PATTERNS`）—— 识别方正小标宋、黑体等粗字面字体

遇到新的粗字面标题字体，加到 `_BOLD_FONT_PATTERNS` 元组即可。

### 兜底策略

未命中映射的字体名会**原样保留**（PDF 字体名直接用）。Word 如果没有对应字体会回退到
默认字体，不影响内容但影响视觉保真度。

## 三、表格解析增强

**位置**：`build_docx.py` 的 `_build_table()` / `_parse_html_table()` / `_fill_table_cells()`

**当前状态**：骨架实现，基础 rowspan/colspan 合并已搭框架

### 已知 TODO

1. **HTML 解析用正则**（`_parse_html_table`）：
   - 当前用正则粗解析，复杂 HTML（嵌套标签、属性顺序变化）可能漏
   - **建议升级**：改用 `html.parser` 或 `lxml`

2. **合并单元格**（`_fill_table_cells`）：
   - 基础 rowspan/colspan 已实现
   - **复杂场景待测**：深度嵌套合并、跨页表、表头重复
   - python-docx 的 `cell.merge()` 在某些合并场景有边界 bug，需真实数据验证

3. **表格样式**：
   - 当前用 `Table Grid`（网格线）
   - 未还原原 PDF 的表格底纹、边框样式

4. **表格内文本样式**：
   - 当前未设单元格内文字的字号字体
   - **建议增强**：从 merged.json 的对应 span 取 _style 贴到单元格 run

### 调试方法

```bash
# 检查 middle.json 里 table_body 的真实 HTML 格式
python3 pdf_inspect.py middle work/_middle.json -v
```

看 `table_body` 字段的 HTML 结构，确认 rowspan/colspan 属性的写法，
再针对性调整 `_parse_html_table()` 的正则。

## 四、标题层级修正

**位置**：`build_docx.py` 的 `_guess_heading_level()` + 正则常量

**当前策略**：正则优先 > MinerU text_level

```python
_RE_CHAPTER = 第X章/节/篇/部
_RE_LEVEL1  = 1. / 1、
_RE_LEVEL2  = 1.1
_RE_LEVEL3  = 1.1.1
```

### 扩充方法

如果真实招标 PDF 有特殊标题格式（如 `一、` `（一）` `1）`），加到正则常量并补充
`_guess_heading_level()` 的判断逻辑。

**原则**：正则命中的层级优先于 MinerU 的 AI 推断，因为正则对招标文件的
标准章节编号更可靠。

## 五、验证检查清单

调优后，用以下方式验证转换质量：

1. **标题层级**：Word 打开转换后的 docx → 大纲视图，看层级树是否正确
2. **文本完整性**：对比原 PDF 逐段核对，检查有无丢失/乱序
3. **表格结构**：对比原 PDF 的评分表/资质表
4. **字号字体**：肉眼对比 docx 与原 PDF
5. **对齐命中率**：`--keep-work` 后跑 `pdf_inspect.py merged work/_merged.json`

不达标时按 `SKILL.md` 的"调试指引"表定位根因。
