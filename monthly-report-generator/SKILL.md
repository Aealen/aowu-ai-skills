---
name: monthly-report-generator
description: 工作月报生成器。当用户提供每日工作日志或日报文档时，自动分析并生成结构化的工作月报。按项目分类整理，输出Markdown格式。触发场景：用户说"生成月报"、"总结本月工作"、"写月报"、"帮我整理工作日志"等。支持从直接文本、飞书文档、AFFINE文档、Obsidian笔记、本地文件等多种来源读取工作日志。
---

# 工作月报生成器

根据每日工作日志，自动生成结构化的工作月报。

## 触发条件

- 用户提供每日工作日志、日报文档
- 用户要求生成月报、总结本月工作
- 用户说"帮我写月报"、"整理一下这个月的工作"
- 用户提供AFFINE文档链接要求总结
- 用户提供Obsidian笔记（vault内笔记名、路径、日记或搜索关键词）要求总结

## 支持的文档来源

### 1. 直接文本输入
用户直接粘贴工作日志文本。

### 2. 飞书文档
提供飞书文档链接，使用 `feishu-doc` skill 读取内容。

### 3. AFFINE文档
提供AFFINE文档链接：
- **公开分享链接**：`https://app.affine.pro/share/xxx` 或 `https://affine.pro/share/xxx`
- **私有链接**：`https://app.affine.pro/workspace/xxx/xxx` 或 `https://affine.xxx.tech/workspace/xxx/xxx`

**处理方式（按优先级选择）：**

#### 方式1：OpenClaw Browser Relay（推荐，支持私有链接）
当用户提供AFFINE私有链接时，优先使用浏览器控制方式获取内容：

1. **确认Chrome扩展已连接**
   - 用户需在Chrome中安装OpenClaw Browser Relay扩展
   - 在目标标签页点击扩展图标，使badge变为ON状态

2. **获取标签页列表**
   ```
   browser(action="tabs", profile="chrome")
   ```

3. **导航到目标页面**
   ```
   browser(action="navigate", profile="chrome", targetId="xxx", targetUrl="AFFINE文档URL")
   ```

4. **获取页面快照**
   ```
   browser(action="snapshot", profile="chrome", snapshotFormat="ai", targetId="xxx")
   ```

5. **解析快照内容**
   - AFFINE是SPA应用，内容通过JavaScript动态渲染
   - 快照返回的是可访问性树结构，需提取文本内容
   - 重点关注列表项、标题、代码块等结构化内容

**优点**：
- 支持私有链接，无需公开分享
- 可获取完整渲染后的页面内容
- 适用于需要登录或私有工作区的AFFINE实例

#### 方式2：web_fetch（仅限公开分享链接）
对于公开分享的AFFINE文档，可直接使用：
```
web_fetch(url: "https://app.affine.pro/share/xxx")
```

**注意**：此方式仅适用于公开分享链接，私有链接只能获取到页面框架，无法读取实际内容。

### 4. Obsidian 文档（vault 内笔记）

复用 **`obsidian-cli`** skill 提供的 `obsidian` CLI 读取 vault 内容。前置条件：Obsidian 已运行（CLI 通过本地协议与运行中的 Obsidian 实例通信）。默认 vault 为 `Aowu`（路径：`C:\Users\Administrator\Documents\Obsidian-Work-DIr\Aowu`）。

用户可能以下列任一形式提供 Obsidian 来源：

1. **笔记名 / wikilink**：如 `工作日志` 或 `[[工作日志]]`
   ```
   obsidian read file="工作日志"
   ```
2. **相对路径**：如 `日报/2026-06/第一周.md`
   ```
   obsidian read path="日报/2026-06/第一周.md"
   ```
3. **日记 / 每日笔记**：读取当日或指定日期的 daily note
   ```
   obsidian daily:read
   ```
4. **搜索关键词**：用户只给了关键词（如"六月的开发记录"），先用 `search` 定位，再逐个读取
   ```
   obsidian search query="2026-06" limit=20
   obsidian search query="#日报" limit=30
   obsidian tags sort=count counts
   ```
   若需在指定 vault 中操作：
   ```
   obsidian vault="Aowu" search query="工作日志" limit=20
   ```
5. **多笔记合并**：用户给出多个笔记名时，依次 `obsidian read` 后合并内容。

**读取要点**：
- `file=` 按 wikilink 方式解析（仅名称，无需路径/扩展名）；`path=` 为 vault 根下的精确路径。
- Obsidian 笔记可能含 wikilinks、callouts、frontmatter（properties）等 Obsidian Flavored Markdown 语法，提取正文时按普通 Markdown 处理即可（参考 `obsidian-markdown` skill）。
- 若 `obsidian read` 返回错误（如笔记不存在 / Obsidian 未运行），向用户复述原始报错并询问是否改用其他来源。
- 完整命令列表随时可用 `obsidian help` 查询。

### 5. 本地文件
提供本地文件路径（vault 外的任意本地文件），直接读取文件内容。

## 工作流程

### 1. 收集工作日志

首先确认用户提供的工作日志内容来源，按以下优先级处理：

1. **AFFINE文档链接** → 优先使用 Browser Relay（支持私有链接），其次使用 web_fetch（仅公开链接）
2. **飞书文档链接** → 使用 feishu-doc skill 获取内容
3. **Obsidian 笔记** → 使用 obsidian CLI 读取（笔记名 / 路径 / 日记 / 搜索关键词）
4. **本地文件路径** → 直接读取文件
5. **直接文本** → 直接处理

### 2. 分析日志内容

从日志中提取关键信息：
- **项目名称**：识别涉及的不同项目
- **任务类型**：开发、测试、会议、文档、沟通等
- **时间信息**：日期、耗时
- **成果产出**：完成的功能、解决的问题、提交的代码等
- **进展状态**：进行中、已完成、待跟进

### 3. 按项目分类整理

将所有工作按项目分组，每个项目作为一个一级标题。

### 4. 输出月报

## 输出格式规范

**严格遵守以下格式要求：**

1. **不使用代码块** - 所有内容直接以文本呈现，不包裹在 \`\`\` 中
2. **Markdown格式** - 使用标题、列表、加粗等Markdown语法
3. **项目为一级标题** - 每个项目使用 # 作为一级标题

## 输出模板

# 月度工作总结
**报告周期：** YYYY年MM月

---

## 本月工作概览

简要描述本月整体工作情况（2-3句话）。

---

# 项目A名称

## 主要工作内容

- 完成了XXX功能的开发
- 解决了XXX问题
- 参与了XXX会议

## 产出成果

- 功能模块：描述完成的功能
- 文档产出：如有相关文档
- 数据指标：如有可量化的成果

## 当前状态

描述项目当前进展状态。

---

# 项目B名称

## 主要工作内容

- 任务1
- 任务2
- 任务3

## 产出成果

...

## 当前状态

...

---

## 其他工作

- 日常会议
- 团队协作
- 学习提升
- 其他杂项

---

## 下月计划

1. 计划1
2. 计划2
3. 计划3

## 备注

如有需要特别说明的事项。

---

## 处理原则

1. **简洁明了**：避免冗余描述，突出重点
2. **成果导向**：强调完成的工作和产出，而非过程
3. **结构清晰**：按项目分类，便于阅读
4. **实事求是**：基于日志内容总结，不夸大不虚构
5. **可读性强**：使用Markdown格式，层次分明

## 特殊情况处理

- **跨月项目**：标注"延续项目"，说明本月进展
- **临时任务**：归类到"其他工作"
- **未完成任务**：如实记录，并在下月计划中跟进
- **日志不完整**：基于已有信息总结，不臆测缺失内容

## AFFINE文档处理流程

当用户提供AFFINE文档链接时：

### 步骤1：识别链接类型

检查链接格式：
- 包含 `/share/` → 公开分享链接，可使用 web_fetch 或 Browser Relay
- 包含 `/workspace/` → 私有链接，必须使用 Browser Relay

### 步骤2：选择获取方式

#### 方式A：Browser Relay（推荐，支持私有链接）

1. **检查Chrome扩展连接状态**
   ```
   browser(action="tabs", profile="chrome")
   ```
   如果返回"No tab is connected"，提示用户：
   > 请在Chrome中打开一个标签页，然后点击OpenClaw Browser Relay扩展图标（让badge变成ON状态）

2. **导航到AFFINE文档**
   ```
   browser(action="navigate", profile="chrome", targetId="xxx", targetUrl="AFFINE文档URL")
   ```

3. **获取页面快照**
   等待页面加载完成后：
   ```
   browser(action="snapshot", profile="chrome", snapshotFormat="ai", targetId="xxx")
   ```

4. **解析快照内容**
   - AFFINE是SPA应用，内容通过JavaScript动态渲染
   - 快照返回可访问性树结构，需提取文本内容
   - 重点关注列表项、标题、代码块等结构化内容
   - 示例快照结构：
     ```
     - generic [ref=e1]:
       - heading "文档标题" [level=1]
       - list:
         - listitem: "1. 任务项1"
         - listitem: "2. 任务项2"
     ```

#### 方式B：web_fetch（仅限公开分享链接）

```
web_fetch(url: "https://app.affine.pro/share/xxx")
```

**注意**：此方式仅适用于公开分享链接，私有链接只能获取到页面框架。

### 步骤3：解析内容

AFFINE文档获取后需要：
1. 提取纯文本内容
2. 识别文档结构（标题、列表、段落、代码块）
3. 转换为统一格式进行分析

### 步骤4：生成月报

按照标准流程分析日志并生成月报。

## 示例交互

**示例1：直接文本输入**
> 这是我这周的日志：
> 周一：完成了用户登录功能的前端页面
> 周二：和后端对接登录接口，修复了两个bug
> 周三：参加项目评审会议，开始做用户注册功能
> 周四：用户注册功能开发完成
> 周五：代码review，修复review发现的问题

**输出：**
> （按上述模板生成月报，将工作归类到对应用户系统项目下）

**示例2：Obsidian 笔记（按笔记名）**
> 帮我总结这个月的工作，日志在 Obsidian 的"工作日志-202606"笔记里

**处理流程：**
1. 识别为 Obsidian 笔记来源
2. 使用 `obsidian read file="工作日志-202606"` 读取笔记内容
3. 解析并提取工作日志（保留 Markdown 结构）
4. 按项目分类生成月报

**示例3：Obsidian 笔记（按关键词搜索）**
> 帮我整理六月的开发记录，应该都在我的 Obsidian 里，标题带"日报"

**处理流程：**
1. 识别为 Obsidian 来源（仅给关键词）
2. 使用 `obsidian search query="日报" limit=30` 定位笔记列表
3. 从搜索结果中筛选 2026-06 的日报笔记
4. 逐个 `obsidian read` 读取并合并内容
5. 按项目分类生成月报

**示例4：AFFINE公开分享链接**
> 帮我总结这个月的工作：https://app.affine.pro/share/abc123

**处理流程：**
1. 识别为AFFINE公开分享链接
2. 使用 web_fetch 获取文档内容
3. 解析并提取工作日志
4. 按项目分类生成月报

**示例5：AFFINE私有链接（使用Browser Relay）**
> 这些是我这个月的工作日志：https://affine.xxx.tech/workspace/xxx/doc1
> https://affine.xxx.tech/workspace/xxx/doc2

**处理流程：**
1. 识别为AFFINE私有链接
2. 检查Browser Relay连接状态
3. 使用 browser 工具导航到每个文档
4. 获取页面快照并解析内容
5. 合并多个文档内容
6. 按项目分类生成月报

**示例6：多个来源组合**
> 这是我在AFFINE上的日志：https://app.affine.pro/share/abc123
> Obsidian 里还有一篇补充：[[六月会议纪要]]
> 还有一些零散内容：[粘贴文本]

**处理流程：**
1. 获取AFFINE文档内容（web_fetch）
2. 使用 `obsidian read file="六月会议纪要"` 读取 Obsidian 笔记
3. 合并补充文本
4. 统一分析生成月报

## 注意事项

1. **AFFINE私有文档**：优先使用 Browser Relay 获取，无需公开分享
2. **Browser Relay前置条件**：用户需在Chrome中安装并启用OpenClaw Browser Relay扩展
3. **SPA内容渲染**：AFFINE是SPA应用，必须等待JavaScript渲染完成后才能获取内容
4. **文档格式**：AFFINE支持富文本，提取时保留结构信息
5. **网络依赖**：获取在线文档需要网络连接
6. **内容验证**：获取后确认内容完整性，如有缺失提示用户
7. **多文档处理**：用户可能提供多个AFFINE链接，需逐个获取并合并分析
8. **Obsidian 前置条件**：使用 `obsidian` CLI 读取 vault 前需确认 Obsidian 已运行；CLI 通过本地协议与运行中的 Obsidian 通信，未运行时会报错，此时提示用户启动 Obsidian 或改用其他来源
9. **Obsidian 来源多样性**：用户可能给出笔记名、wikilink、相对路径、日记或仅关键词；关键词需先用 `obsidian search` 定位再读取
10. **Obsidian 笔记格式**：笔记可能含 wikilinks、callouts、frontmatter 等 Obsidian Flavored Markdown 扩展语法，提取正文时按标准 Markdown 处理即可（wikilinks 可视为内部引用，通常不进入月报正文）
