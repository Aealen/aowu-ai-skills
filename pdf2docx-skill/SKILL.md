---
name: pdf2docx
metadata:
  author: aowu
  version: "1.0"
  description: >
    将文字版 PDF（尤其招标文件）高保真转换为 DOCX。采用 MinerU 版面结构分析
    + PyMuPDF 字符级样式提取的双数据源方案，用 bbox 坐标对齐合并后用 python-docx
    重建。转换产出的 DOCX 尽量接近"用户手动用 Word 另存的 DOCX"，可直接接入
    现有 DOCX 处理流程。适用于 PDF转Word、招标文件转换、PDF文档还原等场景。
---

# PDF 转 DOCX —— 高保真文档还原

## 这个 Skill 做什么

将文字版 PDF 高保真转换为 DOCX。核心是**保真还原**——不是凑一个能跑的 docx，
而是尽量还原原文的标题层级、字号字体、表格结构，让转换结果可直接接入下游 DOCX 流程。

**适用场景**：招标文件、合同、报告、说明书等文字版 PDF（不含扫描件）。

## 快速开始

> **⚠️ 关键环境变量（国内必须）**：MinerU 默认从 HuggingFace 下载模型，国内连不上会卡死。
> 必须设置 `MINERU_MODEL_SOURCE=modelscope`（脚本内部已自动设置，但建议同时设为系统环境变量）：
> ```powershell
> # PowerShell 永久设置（执行一次，重启终端生效）
> [Environment]::SetEnvironmentVariable("MINERU_MODEL_SOURCE", "modelscope", "User")
> ```
> ```bash
> # bash 永久设置
> echo 'export MINERU_MODEL_SOURCE=modelscope' >> ~/.bashrc
> ```

> **Windows 用户注意**：本机若未装真实 Python（`python3` 命中 Windows Store 占位符），
> 推荐用 [uv](https://docs.astral.sh/uv/) 管理。本仓库已配好 `.venv`：
> ```bash
> uv venv --python 3.13                    # 建虚拟环境（仅首次）
> uv pip install -r requirements.txt       # 装依赖
> # 之后所有命令用 uv run python：
> uv run python "$SKILL_DIR/scripts/pdf2docx.py" convert input.pdf -o output.docx
> ```
> macOS/Linux 有系统 Python 3.10+ 的，直接用 `python3` 即可。

### 1. 环境准备（首次使用）

```bash
# 装依赖
uv pip install -r requirements.txt

# 下载 MinerU 模型权重（约 2GB，国内用 modelscope 源）
MINERU_MODEL_SOURCE=modelscope uv run mineru-models-download -s modelscope -m pipeline

# 检查环境
python3 "$SKILL_DIR/scripts/pdf2docx.py" env.check
# Windows uv: uv run python "$SKILL_DIR/scripts/pdf2docx.py" env.check
```

### 2. 一键转换

```bash
python3 "$SKILL_DIR/scripts/pdf2docx.py" convert input.pdf -o output.docx
# Windows uv: uv run python "$SKILL_DIR/scripts/pdf2docx.py" convert input.pdf -o output.docx
```

> 88 页招标 PDF 在 CPU 上约 2-3 分钟（含公式+表格识别）。
> 加 `--no-formula --no-table` 可关掉公式/表格识别加快速度。

### 3. 字段检查（首次调试建议先跑）

```bash
# 检查 PDF 的 PyMuPDF span 样式字段
python3 "$SKILL_DIR/scripts/pdf2docx.py" inspect input.pdf -v

# 跑完 MinerU 后，检查 middle.json 真实结构
python3 "$SKILL_DIR/scripts/pdf_inspect.py" middle <middle.json路径>
```

## MinIO 集成（从对象存储读写）

支持直接从 MinIO 下载源 PDF、转换后上传 DOCX。MinIO 连接信息全部走环境变量，
缺失即报错（本地路径模式不需要 MinIO env）。

### 环境变量（MinIO 模式必填）

| 变量 | 说明 |
|------|------|
| `MINIO_ENDPOINT` | MinIO 地址，如 `http://127.0.0.1:9000` |
| `MINIO_ACCESS_KEY` | 访问密钥 |
| `MINIO_SECRET_KEY` | 秘密密钥 |
| `MINIO_BUCKET` | bucket 名 |
| `MINIO_PUBLIC_URL` | 可选，对外 URL；配了才能用 URL 形式传 `--minio-input` |

> **env 缺失行为**：必填项任一缺失，立即抛 `MinioConfigError`，退出码 2，
> stdout 输出 `{"status":"error","code":2,...}`。外部 Agent 读 stderr/stdout
> 即可判断"未配 MinIO"。源码不内置任何默认连接信息。

### 用法

```bash
# 1. 内联 JSON（短 spec）
python3 "$SKILL_DIR/scripts/pdf2docx.py" convert \
  --minio-input '{"object_key":"generation/demo/x.pdf"}' \
  --minio-output '{"object_key":"generation/demo/x.docx"}'

# 2. 文件 JSON（长 spec，避 shell 转义）
python3 "$SKILL_DIR/scripts/pdf2docx.py" convert \
  --minio-input-file input.json \
  --minio-output-file output.json

# 3. stdin（agent 管道注入）
cat input.json | python3 "$SKILL_DIR/scripts/pdf2docx.py" convert \
  --minio-input - --minio-output-file output.json

# 4. URL 形式输入（需 MINIO_PUBLIC_URL 已配）
python3 "$SKILL_DIR/scripts/pdf2docx.py" convert \
  --minio-input '{"url":"http://domain/upload/tender/generation/demo/x.pdf"}' \
  --minio-output '{"object_key":"generation/demo/x.docx"}'

# 5. 混用：本地 PDF → 上传 MinIO
python3 "$SKILL_DIR/scripts/pdf2docx.py" convert input.pdf \
  --minio-output '{"object_key":"generation/demo/output.docx"}'
```

### JSON Schema

**输入**（`--minio-input`）：
```json
{
  "url": "http://.../tender/generation/demo/x.pdf",
  "object_key": "generation/demo/x.pdf",
  "filename": "招标文件.pdf",
  "metadata": {"file_id": "abc", "source": "tender-parse"}
}
```
- `url`/`object_key` 二选一（都给则 `object_key` 优先）
- `metadata` 透传到转换日志和返回值，**不落 MinIO 对象**

**输出**（`--minio-output`）：
```json
{
  "object_key": "generation/demo/x.docx",
  "content_type": "application/vnd...wordprocessingml.document",
  "metadata": {"file_id": "abc", "converted_from": "pdf"}
}
```
- `object_key` 可省略（自动推导：输入同目录 + `.pdf→.docx`）
- `metadata` **写入 MinIO 对象 header**（`x-amz-meta-*`），下游 `stat_object` 可读
- 自动注入溯源字段：`converter`、`converted_from`、`converted_at`

### 退出码

| code | 场景 |
|------|------|
| 0 | 成功 |
| 1 | 转换本身失败 |
| 2 | MinIO env 缺失 |
| 3 | 下载失败 |
| 4 | 上传失败 |
| 5 | JSON 解析/参数校验失败 |

完整设计文档：`docs/2026-08-03-minio-integration-design.md`

## 共享缓存协议（沙箱环境）

沙箱容器每次冷启动会重新下载 ~4GB（依赖 1.5GB + MinerU 模型 2.5GB），
准备环境耗时 5-15 分钟。若平台支持**共享目录**（注入环境变量 `SHARED_DIR`
并读写挂载进所有沙箱），本 Skill 自动复用全局缓存，秒级启动。

**平台配置（一次性）**：
- 注入环境变量 `SHARED_DIR=/共享目录根路径`（所有沙箱同一路径）
- 将共享目录读写挂载进所有沙箱容器

**共享目录布局**（Skill 自动维护，无需人工干预）：
```
$SHARED_DIR/pdf2docx/
  venv-<hash>/       完整 venv（依赖，hash = requirements 内容哈希）
  modelscope/        MinerU 模型缓存
  cache/uv/          uv 构建缓存
  manifest.json      版本指针（Skill 更新自动触发重建）
  .lock              并发锁文件
```

**使用流程**（每次转换前必跑，幂等）：
```bash
python3 "$SKILL_DIR/scripts/env_setup.py"
```
- 成功 → stdout 输出 `SHARED_VENV_PY=...` 和 `MODELSCOPE_CACHE=...`
- 后续命令用共享 venv 执行：
  ```bash
  "$SHARED_VENV_PY" "$SKILL_DIR/scripts/pdf2docx.py" convert input.pdf -o output.docx
  ```
- 未设置 `SHARED_DIR` 或平台不支持共享锁（flock）→ 自动降级本地模式
  （现有 `uv run python` 流程）

**并发安全**：多沙箱同时首启，仅第一个执行初始化（flock 锁 + double-check），
其余等待后直接复用，无重复下载。版本变更自动重建新 venv（哈希目录原子切换），
旧版本保留最近 2 个后清理。

## 处理流程（四步管线）

```
PDF 输入
    │
    ├──[1] MinerU 解析 ──→ 版面结构（标题层级、表格、图片、段落块类型）
    │                      产出 middle.json
    │
    ├──[2] PyMuPDF 解析 ─→ 字符级样式（font/size/color/flags）
    │                      产出 spans.json
    │
    ├──[3] bbox 对齐合并 ─→ 用坐标相交把样式贴回结构
    │                      产出 merged.json
    │
    └──[4] python-docx 重建 → 按 type 分发：title/text/table/image/list
                             产出 DOCX
```

**核心思路**：MinerU 给"版面语义"（这块是表格、那块是标题），PyMuPDF 给"字符样式"
（字号字体粗体颜色）。两者坐标系一致（都是 PDF 点），用 bbox 相交对齐合并，即得
既有结构又有样式的完整数据。

## 分步调试

每一步可单独运行，便于定位问题：

```bash
# 仅 MinerU 解析（版面结构）
python3 "$SKILL_DIR/scripts/pdf2docx.py" parse input.pdf -o ./work/

# 仅 PyMuPDF 样式提取
python3 "$SKILL_DIR/scripts/pdf2docx.py" extract input.pdf -o ./work/

# 仅 bbox 对齐（需要先有 middle.json 和 spans.json）
python3 "$SKILL_DIR/scripts/pdf2docx.py" align work/_middle.json work/_spans.json -o ./work/

# 仅 DOCX 重建
python3 "$SKILL_DIR/scripts/pdf2docx.py" build work/_merged.json -o output.docx
```

一键转换加 `--keep-work` 可保留中间产物供调试：

```bash
python3 "$SKILL_DIR/scripts/pdf2docx.py" convert input.pdf -o output.docx --keep-work
```

## 脚本路径设置（调用前必做）

所有脚本用绝对路径调用，`$SKILL_DIR` 指向本 SKILL.md 所在目录：

```bash
SKILL_DIR="<skill_directory>"   # ← 本 SKILL.md 的父目录
python3 "$SKILL_DIR/scripts/pdf2docx.py" convert input.pdf -o output.docx
```

**⚠️ 不要用 `python3 scripts/pdf2docx.py`** —— 这只在 cwd 是 skill 目录时才生效。

## 关键参数

| 参数 | 说明 | 默认值 |
|------|------|--------|
| `-b/--backend` | MinerU 后端：`hybrid-engine`（混合）/ `pipeline`（纯规则，快）/ `vlm-engine` | `hybrid-engine` |
| `-m/--method` | 解析方法：`auto`（自动判断 OCR）/ `txt`（纯文字版）/ `ocr` | `auto` |
| `-l/--lang` | 语言 | `ch`（中文） |
| `--iou` | bbox 对齐 IoU 阈值 | `0.3` |
| `--start/--end` | 页码范围（0-based） | 全部 |
| `--api-url` | 外部 MinerU API 地址（默认启动本地 server） | 无 |

## 验收标准

转换质量按优先级验收：

| 优先级 | 验收项 |
|--------|--------|
| **P0** | 标题层级正确：第X章 / 1.1 / 1.1.1 → Heading 1/2/3 + OutlineLevel |
| **P0** | 正文段落完整：无大面积丢失或乱序 |
| **P0** | 表格结构还原：行列基本对齐 |
| **P0** | 字号字体基本还原：标题大字粗体、正文小字 |
| P1 | 图片位置基本正确 |
| P1 | 颜色基本还原 |
| P2 | 分栏排版阅读顺序正确 |

## 适用边界

**适用**：文字版 PDF（原生数字 PDF，文本可复制选中）。

**不适用**：
- 扫描件 PDF（需 OCR，本 Skill 不主攻，但 MinerU `-m ocr` 可保留能力）
- 加密 PDF（需先解密）
- 纯图片 PDF（无文本层）

## 调试指引

转换质量不达标时，按以下顺序定位：

| 现象 | 可能根因 | 排查方法 |
|------|----------|----------|
| 标题层级错乱 | MinerU text_level 不准 | 检查 middle.json，看 text_level 字段 |
| 文本丢失/乱序 | MinerU 版面解析遗漏 | 对比原 PDF，评估是普遍还是个例 |
| 表格崩 | MinerU 表格识别失败 | 检查 middle.json 的 table_body |
| 对齐命中率低 | bbox 匹配阈值问题 | 看对齐统计，调 `--iou` 参数 |
| 字体不对 | 字体映射表不全 | 扩充 `build_docx.py` 的 `_FONT_MAP` |

详见 `references/tuning-guide.md`。

## 加载协议

1. **必读**：本 SKILL.md
2. **按需加载**：
   - `references/architecture.md` —— 四模块数据流、坐标对齐原理（理解方案时加载）
   - `references/tuning-guide.md` —— IoU 阈值/字体映射/表格解析的调参指引（调优时加载）

## 文件结构

```
SKILL.md                          ← 你在这里（路由入口）
setup.sh                          ← 环境检查 + 依赖安装
requirements.txt                  ← Python 依赖清单
scripts/
  pdf2docx.py                     ← 统一 CLI 入口（子命令式）
  parse_mineru.py                 ← [模块1] MinerU 版面结构分析
  parse_pymupdf.py                ← [模块2] PyMuPDF 字符级样式提取
  align.py                        ← [模块3] bbox 对齐合并（核心）
  build_docx.py                   ← [模块4] python-docx 重建（核心）
  pdf_inspect.py                  ← 字段检查工具（先头部队）
references/
  architecture.md                 ← 数据流 + 坐标对齐原理
  tuning-guide.md                 ← 调参指引
docs/
  PDF转DOCX技术方案-MVP.md        ← 完整技术方案文档
```
