# pdf2docx-skill

> 高保真 PDF → DOCX 转换工具,专为招标文件等文字版 PDF 设计。
> 支持 MinIO 对象存储直接读写,适配沙箱化 Agent 部署。

## 简介

采用 **MinerU 版面结构分析 + PyMuPDF 字符级样式提取** 的双数据源方案,
通过 bbox 坐标对齐合并后用 python-docx 重建。转换产出的 DOCX 尽量接近
"用户手动用 Word 另存的 DOCX",可直接接入现有 DOCX 处理流程。

**适用场景**:招标文件、合同、报告、说明书等文字版 PDF(不含扫描件)。

## 架构

```
PDF 输入
    │
    ├──[1] MinerU 解析 ──→ 版面结构(标题层级、表格、图片、段落块类型)
    │                      产出 middle.json
    │
    ├──[2] PyMuPDF 解析 ─→ 字符级样式(font/size/color/flags)
    │                      产出 spans.json
    │
    ├──[3] bbox 对齐合并 ─→ 用坐标相交把样式贴回结构
    │                      产出 merged.json
    │
    └──[4] python-docx 重建 → 按 type 分发:title/text/table/image/list
                             产出 DOCX
```

核心思路:MinerU 给"版面语义"(这块是表格、那块是标题),PyMuPDF 给"字符样式"
(字号字体粗体颜色)。两者坐标系一致(都是 PDF 点),用 bbox 相交对齐合并。

## 安装

### 前置要求

- **Python 3.10+**(推荐 3.13)
- **uv**(Python 包管理器,[安装指南](https://docs.astral.sh/uv/))
- **MinerU 模型权重**(~2GB,首次使用需下载)

### 步骤

```bash
cd skills/pdf2docx-skill

# 1. 创建虚拟环境(仅首次)
uv venv --python 3.13

# 2. 安装依赖
uv pip install -r requirements.txt

# 3. 下载 MinerU 模型权重(约 2GB,国内用 modelscope 源)
MINERU_MODEL_SOURCE=modelscope uv run mineru-models-download -s modelscope -m pipeline

# 4. 检查环境
PYTHONUTF8=1 uv run python scripts/pdf2docx.py env.check
```

> **Windows 用户**:所有 Python 调用必须加 `PYTHONUTF8=1`,否则 GBK 控制台
> 无法输出 emoji/中文,会报 `UnicodeEncodeError`。

## 环境变量

### MinIO 连接(MinIO 模式必填,本地模式不需要)


| 变量                 | 必填  | 说明                                        | 示例                      |
| ------------------ | --- | ----------------------------------------- | ----------------------- |
| `MINIO_ENDPOINT`   | 是   | MinIO 服务地址                                | `http://127.0.0.1:9000` |
| `MINIO_ACCESS_KEY` | 是   | 访问密钥                                      | `minio`                 |
| `MINIO_SECRET_KEY` | 是   | 秘密密钥                                      | `password`              |
| `MINIO_BUCKET`     | 是   | bucket 名                                  | `tender`                |
| `MINIO_PUBLIC_URL` | 否   | 对外暴露的公网 URL;配了才能用 URL 形式传 `--minio-input` | `http://domain/upload`  |


> **缺失行为**:必填项任一缺失,立即抛 `MinioConfigError`,退出码 2。
> stdout 输出 `{"status":"error","code":2,...}`,外部 Agent 可据此判断"未配 MinIO"。
> **源码不内置任何默认连接信息**。

### MinerU 相关


| 变量                    | 必填   | 说明                                                       |
| --------------------- | ---- | -------------------------------------------------------- |
| `MINERU_MODEL_SOURCE` | 国内必须 | 设为 `modelscope`,否则连 HuggingFace 卡死。脚本内部已自动设,但建议永久 export |


永久设置(推荐):

```bash
# bash
echo 'export MINERU_MODEL_SOURCE=modelscope' >> ~/.bashrc

# PowerShell(重启终端生效)
[Environment]::SetEnvironmentVariable("MINIO_ENDPOINT", "http://127.0.0.1:9000", "User")
[Environment]::SetEnvironmentVariable("MINIO_ACCESS_KEY", "minio", "User")
[Environment]::SetEnvironmentVariable("MINIO_SECRET_KEY", "password", "User")
[Environment]::SetEnvironmentVariable("MINIO_BUCKET", "tender", "User")
[Environment]::SetEnvironmentVariable("MINIO_PUBLIC_URL", "http://domain/upload", "User")
[Environment]::SetEnvironmentVariable("MINERU_MODEL_SOURCE", "modelscope", "User")
```

## CLI 参考

统一入口 `scripts/pdf2docx.py`,子命令式:

```bash
PYTHONUTF8=1 uv run python scripts/pdf2docx.py <子命令> [参数]
```

### `convert` — 一键全流程

```bash
# 本地模式
uv run python scripts/pdf2docx.py convert input.pdf -o output.docx

# MinIO 模式
uv run python scripts/pdf2docx.py convert \
  --minio-input '{"object_key":"generation/demo/x.pdf"}' \
  --minio-output '{"object_key":"generation/demo/x.docx"}'
```

**参数:**


| 参数                    | 说明                              | 默认值    |
| --------------------- | ------------------------------- | ------ |
| `pdf`                 | 输入 PDF 路径(位置参数,本地模式;MinIO 模式省略) | —      |
| `-o` / `--output`     | 输出 docx 路径;MinIO-only 模式可省略     | —      |
| `--minio-input`       | MinIO 输入 JSON(内联;`-` 表 stdin)   | —      |
| `--minio-input-file`  | MinIO 输入 JSON 文件路径              | —      |
| `--minio-output`      | MinIO 输出 JSON(内联;`-` 表 stdin)   | —      |
| `--minio-output-file` | MinIO 输出 JSON 文件路径              | —      |
| `-m` / `--method`     | 解析方法:`auto` / `txt` / `ocr`     | `auto` |
| `-l` / `--lang`       | 语言                              | `ch`   |
| `--iou`               | bbox 对齐 IoU 阈值                  | `0.3`  |
| `--no-formula`        | 关闭公式识别(加快速度)                    | —      |
| `--no-table`          | 关闭表格识别(加快速度)                    | —      |
| `--work-dir`          | 中间产物目录(默认在输出旁边)                 | —      |
| `--keep-work`         | 保留中间产物(默认清理)                    | —      |


**输入模式组合:**


| 输入              | 输出                      | 行为              |
| --------------- | ----------------------- | --------------- |
| 本地 PDF (`pdf`)  | `-o`                    | 纯本地,向后兼容        |
| 本地 PDF          | `--minio-output`        | 本地转换后上传 MinIO   |
| `--minio-input` | `-o`                    | 下载后转换,写本地       |
| `--minio-input` | `--minio-output`        | 端到端 MinIO,无本地残留 |
| `--minio-input` | `-o` + `--minio-output` | 既写本地又上传         |


### `env.check` — 环境检查

```bash
uv run python scripts/pdf2docx.py env.check
```

检查 Python 版本、PyMuPDF、python-docx、minio、MinerU 是否就绪。

### `inspect` — 字段检查

```bash
uv run python scripts/pdf2docx.py inspect input.pdf -v
```

检查 PDF 的 PyMuPDF span 样式字段,用于调试。

### 分步调试

```bash
# 仅 MinerU 解析
uv run python scripts/pdf2docx.py parse input.pdf -o ./work/

# 仅 PyMuPDF 样式提取
uv run python scripts/pdf2docx.py extract input.pdf -o ./work/

# 仅 bbox 对齐(需先有 middle.json 和 spans.json)
uv run python scripts/pdf2docx.py align work/_middle.json work/_spans.json -o ./work/

# 仅 DOCX 重建
uv run python scripts/pdf2docx.py build work/_merged.json -o output.docx
```

一键转换加 `--keep-work` 保留中间产物:

```bash
uv run python scripts/pdf2docx.py convert input.pdf -o output.docx --keep-work
```

## JSON Schema(MinIO 模式)

### 输入(`--minio-input`)

```json
{
  "url": "http://domain/upload/tender/generation/demo/x.pdf",
  "object_key": "generation/demo/x.pdf",
  "filename": "招标文件.pdf",
  "metadata": {
    "file_id": "abc-123",
    "source": "tender-parse",
    "business_id": "biz-456"
  }
}
```


| 字段           | 必填  | 说明                                               |
| ------------ | --- | ------------------------------------------------ |
| `url`        | 否   | 完整访问 URL;与 `object_key` 二选一(都给则 `object_key` 优先) |
| `object_key` | 否   | MinIO 对象键;与 `url` 二选一                            |
| `filename`   | 否   | 指定下载后的本地文件名;缺省取 object_key 最后一段                  |
| `metadata`   | 否   | 任意键值对,透传到转换日志和返回值,**不落 MinIO 对象**                |


> `url` 模式需要 `MINIO_PUBLIC_URL` 已配置,否则报错提示改用 `object_key`。

### 输出(`--minio-output`)

```json
{
  "object_key": "generation/demo/x.docx",
  "content_type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  "metadata": {
    "file_id": "abc-123",
    "converted_from": "pdf",
    "converter": "pdf2docx-skill"
  }
}
```


| 字段             | 必填  | 说明                                                         |
| -------------- | --- | ---------------------------------------------------------- |
| `object_key`   | 否   | 上传目标 key。**缺省时自动推导**:输入同目录 + 扩展名 `.pdf→.docx`              |
| `content_type` | 否   | 缺省按 object_key 扩展名推断(`.docx` → 标准 OOXML MIME)              |
| `metadata`     | 否   | **写入 MinIO 对象 header**(`x-amz-meta-*`),下游 `stat_object` 可读 |


自动注入的溯源字段(调用方无需提供):


| 字段               | 值                 |
| ---------------- | ----------------- |
| `converter`      | `pdf2docx-skill`  |
| `converted_from` | `pdf`             |
| `converted_at`   | ISO 8601 时间戳(UTC) |


> 输入端 `metadata` 的字段会合并进输出端(setdefault,不覆盖调用方显式指定的值)。

## stdout 输出格式

转换成功时,stdout 末尾打印结构化 JSON:

```json
{
  "status": "success",
  "output": "C:\\local\\output.docx",
  "stats": {
    "spans_total": 3965,
    "align_matched": 2274,
    "align_total": 1384,
    "align_rate": 164.3
  },
  "minio": {
    "downloaded_key": "generation/demo/x.pdf",
    "uploaded_key": "generation/demo/x.docx",
    "uploaded_url": "http://domain/upload/tender/generation/demo/x.docx",
    "uploaded_size": 110034,
    "metadata_written": {
      "converter": "pdf2docx-skill",
      "converted_from": "pdf",
      "converted_at": "2026-08-03T06:07:35.917840+00:00",
      "file_id": "abc-123"
    }
  }
}
```

纯本地模式无 `minio` 字段。

## 退出码


| code | 场景             | stdout JSON                       |
| ---- | -------------- | --------------------------------- |
| 0    | 成功             | `{"status":"success",...}`        |
| 1    | 转换本身失败         | —                                 |
| 2    | MinIO env 缺失   | `{"status":"error","code":2,...}` |
| 3    | MinIO 下载失败     | `{"status":"error","code":3,...}` |
| 4    | MinIO 上传失败     | `{"status":"error","code":4,...}` |
| 5    | JSON 解析/参数校验失败 | `{"status":"error","code":5,...}` |


外部 Agent 可扫 exit code 或解析 stdout JSON。

## 使用示例

### 示例 1:纯本地转换

```bash
cd skills/pdf2docx-skill
MINERU_MODEL_SOURCE=modelscope PYTHONUTF8=1 \
  uv run python scripts/pdf2docx.py convert 招标文件.pdf -o 招标文件.docx
```

### 示例 2:MinIO 端到端(object_key 形式)

```bash
cd skills/pdf2docx-skill
export MINIO_ENDPOINT="http://127.0.0.1:9000"
export MINIO_ACCESS_KEY="minio"
export MINIO_SECRET_KEY="password"
export MINIO_BUCKET="tender"
export MINIO_PUBLIC_URL="http://domain/upload"

MINERU_MODEL_SOURCE=modelscope PYTHONUTF8=1 \
  uv run python scripts/pdf2docx.py convert \
    --minio-input '{"object_key":"generation/demo/招标文件.pdf","metadata":{"file_id":"bid-001"}}' \
    --minio-output '{"object_key":"generation/demo/招标文件.docx"}'
```

### 示例 3:MinIO 端到端(URL 形式)

```bash
MINERU_MODEL_SOURCE=modelscope PYTHONUTF8=1 \
  uv run python scripts/pdf2docx.py convert \
    --minio-input '{"url":"http://domain/upload/tender/generation/demo/%E6%8B%9B%E6%A0%87%E6%96%87%E4%BB%B6.pdf"}' \
    --minio-output '{"object_key":"generation/demo/招标文件.docx"}'
```

### 示例 4:文件 JSON(避 Windows 引号地狱)

```bash
# 写 input.json
cat > input.json << 'EOF'
{
  "object_key": "generation/demo/招标文件.pdf",
  "metadata": {"file_id": "bid-001", "source": "tender-parse"}
}
EOF

# 写 output.json
cat > output.json << 'EOF'
{
  "object_key": "generation/demo/招标文件.docx",
  "metadata": {"converter": "pdf2docx-skill"}
}
EOF

MINERU_MODEL_SOURCE=modelscope PYTHONUTF8=1 \
  uv run python scripts/pdf2docx.py convert \
    --minio-input-file input.json \
    --minio-output-file output.json
```

### 示例 5:stdin(Agent 管道注入)

```bash
echo '{"object_key":"generation/demo/x.pdf"}' | \
  MINERU_MODEL_SOURCE=modelscope PYTHONUTF8=1 \
  uv run python scripts/pdf2docx.py convert --minio-input - \
    --minio-output '{"object_key":"generation/demo/x.docx"}'
```

## temp 文件管理

端到端 MinIO 流程的本地中转文件放系统 temp 目录:

- **成功**:temp 目录自动删除(除非 `--keep-work`)
- **失败**:temp 目录**保留**,stderr 打印路径,方便排查

## 测试

### 单元测试(minio_client,mock SDK 无真实网络)

```bash
cd skills/pdf2docx-skill
PYTHONUTF8=1 uv run python -m pytest tests/ -v
```

### 端到端测试(需真实 MinIO 环境)

```bash
cd skills/pdf2docx-skill
export MINIO_ENDPOINT="http://127.0.0.1:9000"
export MINIO_ACCESS_KEY="minio"
export MINIO_SECRET_KEY="password"
export MINIO_BUCKET="tender"
export MINIO_PUBLIC_URL="http://domain/upload"

# 下载 → 转换 → 上传 → 反查 metadata
MINERU_MODEL_SOURCE=modelscope PYTHONUTF8=1 \
  uv run python scripts/pdf2docx.py convert \
    --minio-input '{"object_key":"generation/demo/示例文件1.pdf"}' \
    --minio-output '{"object_key":"generation/demo/test_output.docx"}'

# stat 反查
PYTHONUTF8=1 uv run python -c "
import sys; sys.path.insert(0, 'scripts')
from minio_client import MinioClient
c = MinioClient()
info = c.stat('generation/demo/test_output.docx')
print('metadata:', info['metadata'])
"
```

## 文件结构

```
pdf2docx-skill/
├── README.md                        ← 本文档
├── SKILL.md                         ← Skill 清单(Agent 加载入口)
├── setup.sh                         ← 环境检查 + 依赖安装
├── requirements.txt                 ← Python 依赖
├── .gitignore
├── scripts/
│   ├── pdf2docx.py                  ← 统一 CLI 入口(子命令式)
│   ├── minio_client.py              ← MinIO 工具类(可复用)
│   ├── parse_mineru.py              ← [模块1] MinerU 版面结构分析
│   ├── parse_pymupdf.py             ← [模块2] PyMuPDF 字符级样式提取
│   ├── align.py                     ← [模块3] bbox 对齐合并(核心)
│   ├── build_docx.py                ← [模块4] python-docx 重建(核心)
│   ├── pdf_inspect.py               ← 字段检查工具
│   └── env_setup.py                 ← 共享缓存初始化(沙箱环境)
├── tests/
│   ├── __init__.py
│   └── test_minio_client.py         ← minio_client 单元测试
├── references/
│   ├── architecture.md              ← 数据流 + 坐标对齐原理
│   └── tuning-guide.md              ← 调参指引(IoU/字体/表格)
└── docs/
    ├── 2026-08-03-minio-integration-design.md   ← MinIO 集成设计文档
    └── 2026-08-03-minio-integration-plan.md     ← MinIO 集成实现计划
```

## 验收标准


| 优先级    | 验收项                                                     |
| ------ | ------------------------------------------------------- |
| **P0** | 标题层级正确:第X章 / 1.1 / 1.1.1 → Heading 1/2/3 + OutlineLevel |
| **P0** | 正文段落完整:无大面积丢失或乱序                                        |
| **P0** | 表格结构还原:行列基本对齐                                           |
| **P0** | 字号字体基本还原:标题大字粗体、正文小字                                    |
| P1     | 图片位置基本正确                                                |
| P1     | 颜色基本还原                                                  |
| P2     | 分栏排版阅读顺序正确                                              |


## 适用边界

**适用**:文字版 PDF(原生数字 PDF,文本可复制选中)。

**不适用**:

- 扫描件 PDF(需 OCR,本 Skill 不主攻,但 MinerU `-m ocr` 可保留能力)
- 加密 PDF(需先解密)
- 纯图片 PDF(无文本层)

## minio_client.py 独立复用

`scripts/minio_client.py` 是独立的 MinIO 工具类,不含任何 PDF/DOCX 业务逻辑,
其他 Skill 或脚本可直接 import 复用:

```python
import sys
sys.path.insert(0, "path/to/pdf2docx-skill/scripts")
from minio_client import MinioClient

client = MinioClient()  # 读 env,缺失抛 MinioConfigError

# 下载
local_path, metadata = client.download("generation/demo/x.pdf", "./downloads/")

# 上传
result = client.upload("./output.docx", "generation/demo/x.docx",
                       metadata={"source": "my-skill"})

# 查元数据
info = client.stat("generation/demo/x.docx")
```

## 调试指引


| 现象                 | 可能根因                 | 排查方法                             |
| ------------------ | -------------------- | -------------------------------- |
| 标题层级错乱             | MinerU text_level 不准 | 检查 middle.json 的 text_level 字段   |
| 文本丢失/乱序            | MinerU 版面解析遗漏        | 对比原 PDF,评估普遍还是个例                 |
| 表格崩                | MinerU 表格识别失败        | 检查 middle.json 的 table_body      |
| 对齐命中率低             | bbox 匹配阈值问题          | 看对齐统计,调 `--iou` 参数               |
| 字体不对               | 字体映射表不全              | 扩充 `build_docx.py` 的 `_FONT_MAP` |
| MinIO 下载失败         | 对象不存在/权限不足           | 检查 object_key 拼写、access key      |
| UnicodeEncodeError | Windows GBK 控制台      | 加 `PYTHONUTF8=1` 环境变量            |


详见 `references/tuning-guide.md`。