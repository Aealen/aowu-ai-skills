# pdf2docx-skill MinIO 集成设计

> 日期: 2026-08-03
> 状态: 设计稿(待评审)

## 1. 背景与目标

当前 pdf2docx-skill 的 `convert` 子命令只支持本地路径输入输出。在沙箱化部署场景下,
外部 Agent 调用本 Skill 时,源 PDF 与目标 DOCX 通常都存放在 MinIO 对象存储中,
通过本地中转文件交互会带来额外的 IO 与清理负担。

本设计为 Skill 增加 **MinIO 下载/上传** 能力,使其能直接:

1. 从 MinIO 拉取源 PDF
2. 本地完成转换
3. 将结果 DOCX 推回 MinIO,并把溯源元数据写入对象 header

**核心约束**:

- MinIO 连接信息**全部从环境变量获取**,源码不内置任何 endpoint/key(避免密钥泄露到 git)
- 上传/下载能力封装为**独立工具类 `minio_client.py`**,供本 Skill 及未来其他 Skill 复用
- 向后兼容:既有的本地路径用法零改动

## 2. 架构

```
┌─────────────────────────────────────────────────────────┐
│ scripts/pdf2docx.py convert                             │
│   1. 解析 --minio-input / --minio-output (JSON)         │
│   2. 若有 minio 输入 → minio_client.download 到 temp    │
│   3. 四步管线: parse → extract → align → build          │
│   4. 若有 minio 输出 → minio_client.upload              │
│   5. stdout 打印结构化 JSON 结果                        │
└───────────────┬─────────────────────────────────────────┘
                │ 调用(纯函数式工具层)
                ▼
┌─────────────────────────────────────────────────────────┐
│ scripts/minio_client.py  (独立工具类,可被任何脚本 import)│
│   class MinioClient:                                    │
│     download(source, dest_dir) → Path                   │
│     upload(local_path, key, content_type, metadata)     │
│     _load_config() ← env,缺失抛 MinioConfigError        │
└─────────────────────────────────────────────────────────┘
```

**分层原则**:

- `minio_client.py` 是纯工具层,只管连接 MinIO、收发字节,不含任何 PDF/DOCX 业务逻辑
- `pdf2docx.py` 是业务层,负责 JSON 解析、管线编排、调用工具层
- 两者通过明确的函数签名解耦,未来其他 Skill 可直接 `import minio_client` 复用

**文件放置**: `skills/pdf2docx-skill/scripts/minio_client.py`。理由:Skill 最终整体打包
交给沙箱运行,单目录自包含比维护跨 Skill 的共享路径更简单。其他 Skill 需要时直接 copy
单文件即可(无外部包依赖之外的耦合)。

## 3. 环境变量约定

全部连接信息走 env,与 Java 侧 `application-dev.yml` 的 key 名一一对应(仅大小写与分隔符差异):

| 环境变量 | 对应 yml key | 必填 | 用途 |
|---------|-------------|------|------|
| `MINIO_ENDPOINT` | `minio.endpoint` | 是 | SDK 连接地址 |
| `MINIO_ACCESS_KEY` | `minio.accessKey` | 是 | 访问密钥 |
| `MINIO_SECRET_KEY` | `minio.secretkey` | 是 | 秘密密钥 |
| `MINIO_BUCKET` | `minio.bucket` | 是 | bucket 名 |
| `MINIO_PUBLIC_URL` | `minio.publicUrl` | 否 | URL 反解 object key 用;不设则只接受 object_key,拒收 URL |

**安全**:任何 env 缺失(必填项),`MinioClient.__init__` 立即抛 `MinioConfigError`,
消息一次性列出**所有**缺失变量,外部 Agent 读 stderr 即可得出"未配 MinIO"的判断。

源码内**禁止**硬编码任何真实连接信息。

## 4. `minio_client.py` 接口

```python
from pathlib import Path

class MinioConfigError(RuntimeError):
    """MinIO 配置缺失。message 列出所有缺失 env 变量。"""

class MinioClient:
    def __init__(self):
        """读 env。必填项缺失 → 抛 MinioConfigError。"""

    def download(self, source: str | dict, dest_dir: Path) -> tuple[Path, dict]:
        """
        下载对象到本地。
        source:
          - str: object_key 或完整 URL(需 MINIO_PUBLIC_URL 已配)
          - dict: 已解析的 JSON,{"object_key":..., "url":..., "metadata":...}
        dest_dir: 下载目标目录(自动创建)
        返回: (下载后的本地路径, metadata dict)
        异常: 对象不存在/签名错/网络错 → 抛 RuntimeError 含原始 MinIO 错误
        """

    def upload(self, local_path: Path, key: str,
               content_type: str | None = None,
               metadata: dict | None = None) -> dict:
        """
        上传本地文件到 MinIO。
        metadata: 写入对象自定义 header(x-amz-meta-*)
        content_type: 缺省按 local_path 扩展名推断
        返回: {"bucket":..., "key":..., "size":..., "etag":..., "url":...}
        异常: 上传失败 → 抛 RuntimeError 含原始 MinIO 错误
        """

    def stat(self, key: str) -> dict:
        """查对象元数据(校验用)。返回 {size, content_type, metadata, etag}。"""
```

**URL → object_key 反解规则**:

- URL 形如 `{MINIO_PUBLIC_URL}/{MINIO_BUCKET}/{object_key}`
- 反解时 strip 掉 `public_url` 与 `bucket` 前缀,剩余即为 object_key
- URL 编码自动 decode(`示例文件1.pdf` 会被正确还原)
- 若 URL 不以已配置的 `public_url` 开头 → 抛 ValueError 提示"URL 不属于当前 MinIO 实例"
- 若 `MINIO_PUBLIC_URL` 未配但调用方传了 URL → 抛 ValueError 提示"未配 public_url,请改用 object_key"

## 5. JSON Schema

### 5.1 输入 spec(`--minio-input`)

```json
{
  "url": "http://domain/upload/tender/generation/demo/示例文件1.pdf",
  "object_key": "generation/demo/示例文件1.pdf",
  "filename": "招标文件.pdf",
  "metadata": {
    "file_id": "abc-123",
    "source": "tender-parse",
    "business_id": "biz-456"
  }
}
```

字段规则:

| 字段 | 必填 | 说明 |
|------|------|------|
| `url` | 否 | 完整访问 URL;与 `object_key` 二选一 |
| `object_key` | 否 | MinIO 对象键;与 `url` 二选一;**同时给则 object_key 优先** |
| `filename` | 否 | 指定下载后的本地文件名;缺省取 object_key 最后一段 |
| `metadata` | 否 | 任意键值对,透传给 convert 流程的日志/返回值,**不写进 MinIO 对象** |

校验:`url` 与 `object_key` 至少一个;都缺 → JSON 解析阶段报错。

### 5.2 输出 spec(`--minio-output`)

```json
{
  "object_key": "generation/demo/示例文件1_converted.docx",
  "content_type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  "metadata": {
    "file_id": "abc-123",
    "converted_from": "pdf",
    "converter": "pdf2docx-skill",
    "converted_at": "2026-08-03T11:30:00+08:00"
  }
}
```

字段规则:

| 字段 | 必填 | 说明 |
|------|------|------|
| `object_key` | 否 | 上传目标 key。**缺省时自动推导**:输入同目录 + 扩展名 `.pdf→.docx` |
| `content_type` | 否 | 缺省按 object_key 扩展名推断(`.docx` → 标准 OOXML MIME) |
| `metadata` | 否 | **写入 MinIO 对象 header**(`x-amz-meta-*`),下游 `stat_object` 可读 |

`converted_at` 等动态字段由 pdf2docx.py 自动注入,调用方无需提供。

## 6. CLI 表面

### 6.1 新增 flag

| flag | 说明 |
|------|------|
| `--minio-input <JSON>` | 内联 JSON 字符串 |
| `--minio-input-file <path>` | 从文件读 JSON |
| `--minio-input -` | 从 stdin 读 JSON |
| `--minio-output <JSON>` | 同上,输出端 |
| `--minio-output-file <path>` | 同上 |
| `--minio-output -` | 同上 |

**优先级**:同侧(input 或 output)只能用一个;混用报错。
**可混用**:输入走文件、输出走 stdin 合法。

### 6.2 调用示例

```bash
# 1. 全本地(向后兼容,零改动)
pdf2docx.py convert input.pdf -o output.docx

# 2. flag 收短 JSON
pdf2docx.py convert \
  --minio-input '{"object_key":"generation/demo/示例文件1.pdf"}' \
  --minio-output '{"object_key":"generation/demo/示例文件1.docx"}'

# 3. 文件收长 JSON(避 Windows 引号地狱)
pdf2docx.py convert --minio-input-file in.json --minio-output-file out.json

# 4. stdin(管道 / Agent 注入)
cat in.json | pdf2docx.py convert --minio-input - --minio-output-file out.json
```

### 6.3 与既有 flag 的关系

- `-o output.docx`(本地输出)与 `--minio-output` **可共存**:既写本地又上传
- `-o` 缺省且给了 `--minio-output`:只上传,不保留本地 DOCX(temp 转换产物清理掉)
- 本地输入 + `--minio-output`:本地 PDF 转 DOCX 后上传

## 7. temp 文件管理

端到端 MinIO 流程的本地中转文件统一放 `./work/minio_<timestamp>/`:

```
work/minio_20260803_113000/
  ├── source.pdf        # 下载的源文件
  ├── output.docx       # 转换产物
  └── (管线中间产物,如启 --keep-work)
```

**清理策略**:

- **成功**:整个 `work/minio_<timestamp>/` 目录删除(除非 `--keep-work` 显式保留)
- **失败**:目录**保留**,日志打印绝对路径,方便排查
- `--keep-work` 永远保留,无视成功失败

## 8. stdout 输出格式

转换成功时,stdout 末尾打印结构化 JSON(既有行为,扩展字段):

```json
{
  "status": "success",
  "output": "C:\\local\\output.docx",
  "minio": {
    "downloaded_key": "generation/demo/示例文件1.pdf",
    "uploaded_key": "generation/demo/示例文件1.docx",
    "uploaded_url": "http://domain/upload/tender/generation/demo/示例文件1.docx",
    "uploaded_size": 110034,
    "metadata_written": {"file_id":"abc-123","converted_from":"pdf"}
  },
  "stats": {
    "spans_total": 3965,
    "align_matched": 2274,
    "align_total": 1384,
    "align_rate": 164.3
  }
}
```

纯本地模式无 `minio` 字段(向后兼容)。

## 9. 错误处理与退出码

| 场景 | exit code | stderr | stdout JSON |
|------|-----------|--------|-------------|
| MinIO env 缺失 | 2 | `MinIOConfigError: 缺失 MINIO_ENDPOINT, MINIO_ACCESS_KEY...` | `{"status":"error","code":2,...}` |
| 下载失败(对象不存在/签名错/网络) | 3 | MinIO 原始错误 + object_key | 同上 |
| 上传失败 | 4 | 同上 | 同上 |
| JSON 解析/校验失败 | 5 | 字段缺失明细 | 同上 |
| 转换本身失败 | 1 | 现有逻辑 | 现有逻辑 |

外部 Agent 可二选一:扫 exit code 或解析 stdout JSON。

## 10. 依赖

`requirements.txt` 新增:

```
# MinIO 对象存储 SDK —— 下载源文件 / 上传转换结果
minio>=7.2.0
```

纯 Python 包,无原生编译,跨平台。`setup.sh` 与 `env.check` 同步增加 `minio` import 检查。

## 11. 测试计划

### 11.1 单元测试(`tests/test_minio_client.py`)

- `_load_config`:全配齐 / 缺一个 / 全缺(校验错误消息完整性)
- `download`:object_key 模式 / URL 模式 / URL 反解失败 / 对象不存在
- `upload`:正常上传 + metadata 落 header / content_type 推断 / 网络错
- mock `minio.Minio` client,不发真实网络请求

### 11.2 端到端集成测试

用提供的测试环境实测:

- **下载**: `generation/demo/示例文件1.pdf`(URL: `http://domain/upload/tender/generation/demo/示例文件1.pdf`)
- **转换**: 四步管线
- **上传**: `generation/demo/示例文件1_converted.docx`,metadata 写 `{"converter":"pdf2docx-skill","converted_from":"pdf"}`
- **校验**: 用 `mc stat` 或 SDK `stat_object` 反查 metadata 是否落 header
- **清理**: 测试产物可保留在 `generation/demo/` 下供人工抽查

### 11.3 回归测试

- 本地路径模式(`convert input.pdf -o output.docx`)跑一遍,确认零破坏
- 既有 `--keep-work` / `--start` / `--end` 等参数与新 flag 共存无冲突

## 12. 文件改动清单

| 文件 | 改动 |
|------|------|
| `scripts/minio_client.py` | **新建** — MinioClient 工具类 |
| `scripts/pdf2docx.py` | 修改 `cmd_convert`:解析 minio flag,集成 download/upload,扩展 stdout JSON |
| `requirements.txt` | 加 `minio>=7.2.0` |
| `setup.sh` | 加 `minio` import 检查 |
| `scripts/pdf2docx.py env.check` | 加 `minio` 包检查 |
| `tests/test_minio_client.py` | **新建** — 单元测试 |
| `SKILL.md` | 文档:env 变量说明、新 flag 用法、JSON schema |

## 13. 非目标(YAGNI)

明确**不做**的事:

- ❌ 不做 MinIO 目录浏览/list 操作(当前用例用不到)
- ❌ 不做断点续传(单文件 < 100MB,SDK 内置分片已够)
- ❌ 不做 MinIO 连接池配置(SDK 自管)
- ❌ 不做跨 Skill 共享库目录(`skills/common/`),保持单 Skill 自包含
- ❌ 不内置任何 MinIO 连接默认值(env 必填)

## 14. 风险与对策

| 风险 | 对策 |
|------|------|
| Windows GBK 控制台 + emoji/中文 → UnicodeEncodeError | 所有调用加 `PYTHONUTF8=1`;SKILL.md 文档强调 |
| JSON 在 Windows shell 转义地狱 | 提供 `--minio-input-file` 与 stdin 两条逃生路径 |
| 大文件下载超时 | 依赖 SDK 默认重试;不额外配置 |
| MinIO 测试 bucket 被污染测试产物 | 统一放 `generation/demo/` 前缀,便于一键清理 |
