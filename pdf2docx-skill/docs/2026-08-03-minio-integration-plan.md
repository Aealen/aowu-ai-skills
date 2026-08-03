# pdf2docx-skill MinIO 集成 — 实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 为 pdf2docx-skill 增加 MinIO 下载/上传能力,使 `convert` 子命令能直接从 MinIO 拉源 PDF、转换后推回 DOCX(含溯源 metadata 落对象 header)。

**Architecture:** 新增独立工具类 `scripts/minio_client.py`(env 驱动,无内置默认值,缺失即抛异常),`pdf2docx.py convert` 加 `--minio-input`/`--minio-output` 系列 flag 调用它。MinIO SDK 用官方 `minio>=7.2.0`。本地路径模式零改动,向后兼容。

**Tech Stack:** Python 3.13、minio SDK 7.2+、pytest(测试)、uv(环境管理)

## Global Constraints

- **env 必填,无默认值**:`MINIO_ENDPOINT`/`MINIO_ACCESS_KEY`/`MINIO_SECRET_KEY`/`MINIO_BUCKET` 缺失即抛 `MinioConfigError`;`MINIO_PUBLIC_URL` 可选(仅 URL 反解用)。源码**禁止**硬编码任何真实连接信息。
- **Windows 编码**:所有 Python 调用加 `PYTHONUTF8=1`,避免 GBK 控制台炸 emoji/中文。
- **工作目录**:所有命令 `cd skills/pdf2docx-skill` 后执行。venv 已存在(`.venv/`)。
- **测试运行**:`PYTHONUTF8=1 uv run python -m pytest tests/ -v`
- **不提交 git**(遵守 CLAUDE.md "如非授权不提交")。

---

## File Structure

| 文件 | 责任 | 操作 |
|------|------|------|
| `scripts/minio_client.py` | MinIO 连接 + download/upload/stat 工具类,env 驱动 | **新建** |
| `tests/test_minio_client.py` | minio_client 单元测试(mock SDK,无真实网络) | **新建** |
| `tests/__init__.py` | 包标记(空文件) | **新建** |
| `scripts/pdf2docx.py` | CLI 入口;`cmd_convert` 集成 MinIO 流程 | **修改** |
| `requirements.txt` | 加 `minio>=7.2.0` | **修改** |
| `setup.sh` | 加 `minio` import 检查 | **修改** |
| `SKILL.md` | 文档:env 变量、新 flag、JSON schema | **修改** |
| `.gitignore` | 加 `tests/__pycache__/`(如需) | **修改** |

---

## Task 1: MinioClient 工具类 + 单元测试

**Files:**
- Create: `skills/pdf2docx-skill/scripts/minio_client.py`
- Create: `skills/pdf2docx-skill/tests/__init__.py`
- Create: `skills/pdf2docx-skill/tests/test_minio_client.py`

**Interfaces:**
- Consumes: 无(纯标准库 + minio SDK)
- Produces:
  - `class MinioConfigError(RuntimeError)` — 配置缺失异常
  - `class MinioClient` — `__init__(client_factory=None)`, `download(source, dest_dir) -> (Path, dict)`, `upload(local_path, key, content_type=None, metadata=None) -> dict`, `stat(key) -> dict`

- [ ] **Step 1: 安装 pytest + minio SDK**

```bash
cd skills/pdf2docx-skill
VIRTUAL_ENV="$(pwd)/.venv" uv pip install pytest minio>=7.2.0
```

验证:
```bash
PYTHONUTF8=1 uv run python -c "import minio; import pytest; print('ok', minio.__version__, pytest.__version__)"
```
期望: `ok 7.x.x x.x.x`

- [ ] **Step 2: 建 tests 包**

```bash
# 创建空文件
touch tests/__init__.py   # bash
# 或 PowerShell: New-Item tests/__init__.py -ItemType File
```

- [ ] **Step 3: 写失败测试(tests/test_minio_client.py)**

```python
"""minio_client 单元测试 —— mock SDK,无真实网络。"""
import sys
from pathlib import Path
from unittest.mock import MagicMock

import pytest

sys.path.insert(0, str(Path(__file__).parent.parent / "scripts"))
from minio_client import MinioClient, MinioConfigError


REQUIRED = ["MINIO_ENDPOINT", "MINIO_ACCESS_KEY", "MINIO_SECRET_KEY",
            "MINIO_BUCKET", "MINIO_PUBLIC_URL"]


@pytest.fixture
def env_set(monkeypatch):
    for k, v in {"MINIO_ENDPOINT": "http://127.0.0.1:9000",
                 "MINIO_ACCESS_KEY": "minio",
                 "MINIO_SECRET_KEY": "password",
                 "MINIO_BUCKET": "tender",
                 "MINIO_PUBLIC_URL": "http://domain/upload"}.items():
        monkeypatch.setenv(k, v)


@pytest.fixture
def env_unset(monkeypatch):
    for k in REQUIRED:
        monkeypatch.delenv(k, raising=False)


def _mock_factory(*args):
    return MagicMock()


# ── 配置加载 ──

def test_config_missing_all(env_unset):
    with pytest.raises(MinioConfigError) as exc:
        MinioClient(client_factory=_mock_factory)
    msg = str(exc.value)
    for k in MinioClient.REQUIRED_ENV:
        assert k in msg, f"{k} 应在错误消息中"


def test_config_missing_one(env_unset, monkeypatch):
    monkeypatch.setenv("MINIO_ENDPOINT", "http://x:9000")
    monkeypatch.setenv("MINIO_ACCESS_KEY", "a")
    monkeypatch.setenv("MINIO_SECRET_KEY", "s")
    # MINIO_BUCKET 未设
    with pytest.raises(MinioConfigError) as exc:
        MinioClient(client_factory=_mock_factory)
    assert "MINIO_BUCKET" in str(exc.value)
    assert "MINIO_ENDPOINT" not in str(exc.value)


# ── URL 反解 ──

def test_url_to_key_plain(env_set):
    c = MinioClient(client_factory=_mock_factory)
    assert c._url_to_object_key(
        "http://domain/upload/tender/generation/demo/示例文件1.pdf"
    ) == "generation/demo/示例文件1.pdf"


def test_url_to_key_encoded(env_set):
    c = MinioClient(client_factory=_mock_factory)
    assert c._url_to_object_key(
        "http://domain/upload/tender/generation/demo/%E7%A4%BA%E4%BE%8B%E6%96%87%E4%BB%B61.pdf"
    ) == "generation/demo/示例文件1.pdf"


def test_url_to_key_wrong_host(env_set):
    c = MinioClient(client_factory=_mock_factory)
    with pytest.raises(ValueError, match="不属于当前 MinIO"):
        c._url_to_object_key("http://other.com/upload/tender/x.pdf")


def test_url_to_key_no_public_url(env_set, monkeypatch):
    monkeypatch.delenv("MINIO_PUBLIC_URL")
    c = MinioClient(client_factory=_mock_factory)
    with pytest.raises(ValueError, match="MINIO_PUBLIC_URL 未配置"):
        c._url_to_object_key("http://domain/upload/tender/x.pdf")


# ── source 解析 ──

def test_resolve_str_object_key(env_set):
    c = MinioClient(client_factory=_mock_factory)
    key, meta = c._resolve_object_key("generation/demo/x.pdf")
    assert key == "generation/demo/x.pdf"
    assert meta == {}


def test_resolve_str_url(env_set):
    c = MinioClient(client_factory=_mock_factory)
    key, meta = c._resolve_object_key(
        "http://domain/upload/tender/generation/demo/x.pdf"
    )
    assert key == "generation/demo/x.pdf"


def test_resolve_dict_with_object_key(env_set):
    c = MinioClient(client_factory=_mock_factory)
    key, meta = c._resolve_object_key(
        {"object_key": "generation/demo/x.pdf", "metadata": {"file_id": "abc"}}
    )
    assert key == "generation/demo/x.pdf"
    assert meta == {"file_id": "abc"}


def test_resolve_dict_missing_both(env_set):
    c = MinioClient(client_factory=_mock_factory)
    with pytest.raises(ValueError, match="url 或 object_key"):
        c._resolve_object_key({"metadata": {}})


# ── download ──

def test_download_object_key(env_set, tmp_path):
    mock = MagicMock()
    c = MinioClient(client_factory=lambda *a: mock)
    dest, meta = c.download("generation/demo/x.pdf", tmp_path / "dl")
    mock.fget_object.assert_called_once_with(
        "tender", "generation/demo/x.pdf", str(dest)
    )
    assert dest.name == "x.pdf"
    assert meta == {}


def test_download_dict_with_filename(env_set, tmp_path):
    mock = MagicMock()
    c = MinioClient(client_factory=lambda *a: mock)
    dest, meta = c.download({
        "object_key": "generation/demo/x.pdf",
        "filename": "招标文件.pdf",
        "metadata": {"file_id": "abc"},
    }, tmp_path / "dl")
    assert dest.name == "招标文件.pdf"
    assert meta == {"file_id": "abc"}


def test_download_failure_wraps_error(env_set, tmp_path):
    mock = MagicMock()
    mock.fget_object.side_effect = Exception("NoSuchKey")
    c = MinioClient(client_factory=lambda *a: mock)
    with pytest.raises(RuntimeError, match="MinIO 下载失败"):
        c.download("generation/demo/x.pdf", tmp_path / "dl")


# ── upload ──

def test_upload_writes_metadata_header(env_set, tmp_path):
    mock = MagicMock()
    mock.fput_object.return_value = MagicMock(etag="abc123")
    local = tmp_path / "x.docx"
    local.write_bytes(b"fake docx content")
    c = MinioClient(client_factory=lambda *a: mock)
    result = c.upload(local, "generation/demo/x.docx",
                      metadata={"converted_from": "pdf"})
    mock.fput_object.assert_called_once()
    kwargs = mock.fput_object.call_args.kwargs
    assert kwargs["metadata"] == {"converted_from": "pdf"}
    assert kwargs["content_type"].startswith("application/vnd")
    assert result == {
        "bucket": "tender",
        "key": "generation/demo/x.docx",
        "size": len(b"fake docx content"),
        "etag": "abc123",
        "url": "http://domain/upload/tender/generation/demo/x.docx",
    }


def test_upload_content_type_inferred(env_set, tmp_path):
    mock = MagicMock()
    mock.fput_object.return_value = MagicMock(etag="x")
    local = tmp_path / "x.pdf"
    local.write_bytes(b"%PDF-1.4")
    c = MinioClient(client_factory=lambda *a: mock)
    c.upload(local, "k/x.pdf")
    assert mock.fput_object.call_args.kwargs["content_type"] == "application/pdf"


def test_upload_failure_wraps_error(env_set, tmp_path):
    mock = MagicMock()
    mock.fput_object.side_effect = Exception("QuotaExceeded")
    local = tmp_path / "x.docx"
    local.write_bytes(b"x")
    c = MinioClient(client_factory=lambda *a: mock)
    with pytest.raises(RuntimeError, match="MinIO 上传失败"):
        c.upload(local, "generation/demo/x.docx")


def test_upload_missing_file(env_set, tmp_path):
    c = MinioClient(client_factory=lambda *a: MagicMock())
    with pytest.raises(FileNotFoundError):
        c.upload(tmp_path / "nonexistent.docx", "k/x.docx")


# ── stat ──

def test_stat(env_set):
    mock = MagicMock()
    obj = MagicMock(size=100, content_type="application/pdf", etag="e1")
    obj.metadata = {"x-amz-meta-file_id": "abc"}
    mock.stat_object.return_value = obj
    c = MinioClient(client_factory=lambda *a: mock)
    result = c.stat("generation/demo/x.pdf")
    assert result["size"] == 100
    assert result["content_type"] == "application/pdf"
    assert result["metadata"] == {"x-amz-meta-file_id": "abc"}
```

- [ ] **Step 4: 运行测试验证失败**

```bash
cd skills/pdf2docx-skill
PYTHONUTF8=1 uv run python -m pytest tests/test_minio_client.py -v 2>&1 | tail -20
```
期望: 全部 ERROR(ImportError: No module named 'minio_client')

- [ ] **Step 5: 实现 scripts/minio_client.py**

```python
"""minio_client.py —— MinIO 对象存储工具类。

env 驱动,无内置默认值。缺失必填 env 抛 MinioConfigError。
供 pdf2docx.py 及未来其他 Skill 复用。
"""
from __future__ import annotations

import os
from pathlib import Path
from urllib.parse import unquote


class MinioConfigError(RuntimeError):
    """MinIO 配置缺失。message 列出所有缺失 env 变量。"""


class MinioClient:
    """MinIO 客户端封装。download / upload / stat。

    通过环境变量配置:
      必填: MINIO_ENDPOINT, MINIO_ACCESS_KEY, MINIO_SECRET_KEY, MINIO_BUCKET
      可选: MINIO_PUBLIC_URL(URL 反解 object_key 用)
    """

    REQUIRED_ENV = ["MINIO_ENDPOINT", "MINIO_ACCESS_KEY",
                    "MINIO_SECRET_KEY", "MINIO_BUCKET"]

    def __init__(self, client_factory=None):
        """
        client_factory: 可选 callable(endpoint, access_key, secret_key) → client。
                        测试注入用。生产留空,内部构造 minio.Minio。
        """
        config = self._load_config()
        self.endpoint = config["MINIO_ENDPOINT"]
        self.access_key = config["MINIO_ACCESS_KEY"]
        self.secret_key = config["MINIO_SECRET_KEY"]
        self.bucket = config["MINIO_BUCKET"]
        self.public_url = (config.get("MINIO_PUBLIC_URL") or "").rstrip("/")

        if client_factory is not None:
            self._client = client_factory(
                self.endpoint, self.access_key, self.secret_key
            )
        else:
            from minio import Minio
            secure = self.endpoint.startswith("https://")
            host = self.endpoint.replace("https://", "").replace("http://", "")
            self._client = Minio(
                host, access_key=self.access_key,
                secret_key=self.secret_key, secure=secure,
            )

    @classmethod
    def _load_config(cls):
        missing = [k for k in cls.REQUIRED_ENV if not os.getenv(k)]
        if missing:
            raise MinioConfigError(
                f"MinIO 未配置,缺失环境变量: {', '.join(missing)}。"
                f"请在沙箱注入或本地 export 后重试。"
            )
        keys = cls.REQUIRED_ENV + ["MINIO_PUBLIC_URL"]
        return {k: os.getenv(k) for k in keys}

    # ──────────────────────────────────────────────────────
    #  URL / source 解析
    # ──────────────────────────────────────────────────────

    def _url_to_object_key(self, url: str) -> str:
        """完整 URL → object_key。需 MINIO_PUBLIC_URL 已配。"""
        if not self.public_url:
            raise ValueError(
                f"传入了 URL({url})但 MINIO_PUBLIC_URL 未配置。"
                f"请改用 object_key,或先 export MINIO_PUBLIC_URL。"
            )
        if not url.startswith(self.public_url):
            raise ValueError(
                f"URL({url})不属于当前 MinIO 实例"
                f"(public_url={self.public_url})。"
            )
        rest = url[len(self.public_url):].lstrip("/")
        if rest.startswith(self.bucket + "/"):
            object_key = rest[len(self.bucket) + 1:]
        else:
            object_key = rest
        return unquote(object_key)

    def _resolve_object_key(self, source) -> tuple[str, dict]:
        """
        source: str(object_key/URL) 或 dict(已解析 JSON)。
        返回: (object_key, metadata)
        """
        if isinstance(source, str):
            if source.startswith(("http://", "https://")):
                return self._url_to_object_key(source), {}
            return source, {}
        if isinstance(source, dict):
            metadata = source.get("metadata") or {}
            if source.get("object_key"):
                return source["object_key"], metadata
            if source.get("url"):
                return self._url_to_object_key(source["url"]), metadata
            raise ValueError("minio input JSON 必须含 url 或 object_key 字段")
        raise TypeError(f"source 必须为 str 或 dict,实际: {type(source)}")

    # ──────────────────────────────────────────────────────
    #  公开 API
    # ──────────────────────────────────────────────────────

    def download(self, source, dest_dir: Path) -> tuple[Path, dict]:
        """
        下载对象到本地。
        source: str | dict(见 _resolve_object_key)
        dest_dir: 下载目标目录(自动创建)
        返回: (本地文件路径, metadata)
        异常: RuntimeError(含 bucket/object_key/原始错误)
        """
        object_key, metadata = self._resolve_object_key(source)
        dest_dir = Path(dest_dir)
        dest_dir.mkdir(parents=True, exist_ok=True)

        if isinstance(source, dict) and source.get("filename"):
            filename = source["filename"]
        else:
            filename = Path(object_key).name

        dest_path = dest_dir / filename
        try:
            self._client.fget_object(self.bucket, object_key, str(dest_path))
        except Exception as e:
            raise RuntimeError(
                f"MinIO 下载失败: bucket={self.bucket}, "
                f"object_key={object_key}, error={e}"
            ) from e
        return dest_path, metadata

    def upload(self, local_path, key: str,
               content_type: str | None = None,
               metadata: dict | None = None) -> dict:
        """
        上传本地文件。metadata 落 x-amz-meta-* header。
        返回: {bucket, key, size, etag, url}
        异常: FileNotFoundError / RuntimeError
        """
        local_path = Path(local_path)
        if not local_path.exists():
            raise FileNotFoundError(f"上传源文件不存在: {local_path}")

        content_type = content_type or self._guess_content_type(local_path)
        metadata = metadata or {}

        try:
            result = self._client.fput_object(
                self.bucket, key, str(local_path),
                content_type=content_type, metadata=metadata,
            )
        except Exception as e:
            raise RuntimeError(
                f"MinIO 上传失败: bucket={self.bucket}, "
                f"object_key={key}, error={e}"
            ) from e

        url = (f"{self.public_url}/{self.bucket}/{key}"
               if self.public_url else None)
        return {
            "bucket": self.bucket,
            "key": key,
            "size": local_path.stat().st_size,
            "etag": getattr(result, "etag", None),
            "url": url,
        }

    def stat(self, key: str) -> dict:
        """查对象元数据。返回 {size, content_type, etag, metadata}。"""
        try:
            obj = self._client.stat_object(self.bucket, key)
        except Exception as e:
            raise RuntimeError(
                f"MinIO stat 失败: object_key={key}, error={e}"
            ) from e
        return {
            "size": obj.size,
            "content_type": obj.content_type,
            "etag": obj.etag,
            "metadata": dict(obj.metadata) if obj.metadata else {},
        }

    @staticmethod
    def _guess_content_type(path: Path) -> str:
        ext = Path(path).suffix.lower()
        mapping = {
            ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            ".doc": "application/msword",
            ".pdf": "application/pdf",
            ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            ".png": "image/png",
            ".jpg": "image/jpeg",
            ".jpeg": "image/jpeg",
        }
        return mapping.get(ext, "application/octet-stream")
```

- [ ] **Step 6: 运行测试验证通过**

```bash
cd skills/pdf2docx-skill
PYTHONUTF8=1 uv run python -m pytest tests/test_minio_client.py -v 2>&1 | tail -30
```
期望: 全部 PASS(20+ tests)

---

## Task 2: 集成到 pdf2docx.py convert 子命令

**Files:**
- Modify: `skills/pdf2docx-skill/scripts/pdf2docx.py`(改 `cmd_convert` + 加 helper + 改 CLI flags)

**Interfaces:**
- Consumes: Task 1 的 `MinioClient`、`MinioConfigError`
- Produces: `convert` 子命令新增 `--minio-input`/`--minio-input-file`/`--minio-output`/`--minio-output-file` flag

- [ ] **Step 1: 加 helper 函数(在 cmd_convert 定义之前插入)**

在 `pdf2docx.py` 文件顶部 import 区之后,`cmd_convert` 之前,插入:

```python
import tempfile
from datetime import datetime, timezone


def _load_minio_spec(inline, file_path):
    """
    解析 minio 输入/输出 JSON spec。
    inline: --minio-input 的值(内联 JSON 字符串,或 '-' 表 stdin)
    file_path: --minio-input-file 的值
    返回: dict 或 None(都未给)。出错 sys.exit(5)。
    """
    if inline is not None and file_path is not None:
        print("✗ 不能同时指定内联 JSON(--minio-input)和文件(--minio-input-file)",
              file=sys.stderr)
        sys.exit(5)
    if inline is not None:
        raw = sys.stdin.read() if inline == "-" else inline
    elif file_path is not None:
        raw = Path(file_path).read_text(encoding="utf-8")
    else:
        return None
    try:
        spec = json.loads(raw)
    except json.JSONDecodeError as e:
        print(f"✗ JSON 解析失败: {e}", file=sys.stderr)
        sys.exit(5)
    if not isinstance(spec, dict):
        print("✗ MinIO spec 必须是 JSON 对象", file=sys.stderr)
        sys.exit(5)
    return spec


def _print_error_json(code, message):
    """stdout 打错误 JSON,供 agent 解析。"""
    print(json.dumps({
        "status": "error",
        "code": code,
        "message": message,
    }, ensure_ascii=False))


def _derive_output_key(input_object_key, fallback_name):
    """
    从输入 object_key 推导输出 key:同目录 + 换扩展名 .pdf→.docx。
    input_object_key 为 None 时用 fallback_name 放 generation/demo/ 下。
    """
    if input_object_key:
        p = Path(input_object_key)
        return str(p.with_suffix(".docx"))
    return f"generation/demo/{Path(fallback_name).stem}.docx"


def _build_upload_metadata(spec, input_meta):
    """组装上传 metadata:spec.metadata + 溯源默认值 + 输入端 metadata。"""
    metadata = dict(spec.get("metadata") or {})
    metadata.setdefault("converted_from", "pdf")
    metadata.setdefault("converter", "pdf2docx-skill")
    metadata.setdefault("converted_at", datetime.now(timezone.utc).isoformat())
    for k, v in (input_meta or {}).items():
        metadata.setdefault(k, v)
    return metadata
```

- [ ] **Step 2: 重写 cmd_convert,集成 MinIO 流程**

将 `pdf2docx.py` 中整个 `cmd_convert` 函数(原 L100-180)替换为:

```python
def cmd_convert(args) -> None:
    """一键全流程:PDF → DOCX。支持本地或 MinIO 输入输出。"""
    # ── 解析 MinIO spec ──
    minio_input_spec = _load_minio_spec(
        getattr(args, "minio_input", None),
        getattr(args, "minio_input_file", None),
    )
    minio_output_spec = _load_minio_spec(
        getattr(args, "minio_output", None),
        getattr(args, "minio_output_file", None),
    )

    # ── 确定 PDF 输入 ──
    temp_dirs = []          # 待清理的 temp 目录
    downloaded_key = None   # MinIO 下载的 object_key(推导输出 key 用)
    input_meta = {}

    if minio_input_spec is not None:
        if args.pdf is not None:
            print("✗ 不能同时指定本地 PDF 路径和 --minio-input", file=sys.stderr)
            sys.exit(5)
        from minio_client import MinioClient, MinioConfigError
        try:
            client = MinioClient()
        except MinioConfigError as e:
            print(f"✗ {e}", file=sys.stderr)
            _print_error_json(2, str(e))
            sys.exit(2)
        dl_dir = Path(tempfile.mkdtemp(prefix="pdf2docx_dl_"))
        temp_dirs.append(dl_dir)
        try:
            pdf_path, input_meta = client.download(minio_input_spec, dl_dir)
        except (RuntimeError, ValueError) as e:
            print(f"✗ {e}", file=sys.stderr)
            _print_error_json(3, str(e))
            sys.exit(3)
        pdf_path = pdf_path.resolve()
        if isinstance(minio_input_spec, dict):
            downloaded_key = minio_input_spec.get("object_key")
        elif not str(minio_input_spec).startswith(("http://", "https://")):
            downloaded_key = minio_input_spec
    else:
        if args.pdf is None:
            print("✗ 必须提供 PDF 路径(位置参数)或 --minio-input", file=sys.stderr)
            sys.exit(5)
        pdf_path = Path(args.pdf).resolve()
        if not pdf_path.exists():
            print(f"✗ 输入文件不存在: {pdf_path}", file=sys.stderr)
            sys.exit(1)

    # ── 确定输出路径 ──
    keep_local_output = args.output is not None
    if minio_output_spec is not None and args.output is None:
        # 只上传,产出到 temp
        out_dir = Path(tempfile.mkdtemp(prefix="pdf2docx_out_"))
        temp_dirs.append(out_dir)
        output_path = out_dir / "output.docx"
    elif args.output is not None:
        output_path = Path(args.output).resolve()
    else:
        print("✗ 必须提供 -o/--output 或 --minio-output", file=sys.stderr)
        sys.exit(5)

    work_dir = Path(args.work_dir).resolve() if args.work_dir \
        else output_path.parent / f"_pdf2docx_work_{pdf_path.stem}"
    work_dir.mkdir(parents=True, exist_ok=True)

    print(f"═══════════════════════════════════════════════", file=sys.stderr)
    print(f"  PDF 转 DOCX", file=sys.stderr)
    print(f"  输入: {pdf_path}", file=sys.stderr)
    print(f"  输出: {output_path}", file=sys.stderr)
    print(f"  中间产物: {work_dir}", file=sys.stderr)
    if downloaded_key:
        print(f"  MinIO 下载: {downloaded_key}", file=sys.stderr)
    if minio_output_spec is not None:
        print(f"  MinIO 上传: 已启用", file=sys.stderr)
    print(f"═══════════════════════════════════════════════\n", file=sys.stderr)

    # ── Step 1: MinerU 解析 ──
    from parse_mineru import parse_with_mineru
    middle_path, images_dir = parse_with_mineru(
        pdf_path=str(pdf_path),
        output_dir=str(work_dir),
        parse_method=args.method,
        formula_enable=not args.no_formula,
        table_enable=not args.no_table,
        language=args.lang,
    )

    # ── Step 2: PyMuPDF 样式提取 ──
    from parse_pymupdf import extract_spans, save_spans, extract_underline_lines
    spans = extract_spans(str(pdf_path))
    spans_path = save_spans(spans, work_dir / "_spans.json")
    underline_lines = extract_underline_lines(str(pdf_path), spans)
    print(f"[2/4] PyMuPDF 样式提取: {len(spans)} 个 span, "
          f"{len(underline_lines)} 条填空下划线 → {spans_path}",
          file=sys.stderr)

    # ── Step 3: bbox 对齐合并 ──
    from align import align_and_merge, load_mineru, save_merged
    mineru_data = load_mineru(middle_path)
    merged_data, stats = align_and_merge(
        mineru_data, spans, iou_threshold=args.iou,
        underline_lines=underline_lines,
    )
    merged_path = save_merged(merged_data, work_dir / "_merged.json")

    rate = (stats["matched"] / stats["total"] * 100) if stats["total"] else 0
    print(f"[3/4] bbox 对齐: 命中率 {rate:.1f}% "
          f"({stats['matched']}/{stats['total']}) → {merged_path}",
          file=sys.stderr)
    if rate < 80 and stats["total"] > 0:
        print(f"  ⚠️ 命中率低于 80%,样式还原可能不完整", file=sys.stderr)

    # ── Step 4: DOCX 重建 ──
    from build_docx import build_docx
    result_path = build_docx(merged_data, images_dir, str(output_path),
                             pdf_path=str(pdf_path))

    # ── MinIO 上传 ──
    uploaded_info = None
    if minio_output_spec is not None:
        from minio_client import MinioClient, MinioConfigError
        try:
            client = MinioClient()
        except MinioConfigError as e:
            print(f"✗ {e}", file=sys.stderr)
            _print_error_json(2, str(e))
            sys.exit(2)
        upload_key = minio_output_spec.get("object_key") or \
            _derive_output_key(downloaded_key, pdf_path.name)
        content_type = minio_output_spec.get("content_type")
        metadata = _build_upload_metadata(minio_output_spec, input_meta)
        try:
            uploaded_info = client.upload(
                result_path, upload_key,
                content_type=content_type, metadata=metadata,
            )
            print(f"  ✅ MinIO 上传成功: {uploaded_info['key']}", file=sys.stderr)
        except (RuntimeError, FileNotFoundError) as e:
            print(f"✗ {e}", file=sys.stderr)
            _print_error_json(4, str(e))
            sys.exit(4)
        # 只上传模式(无 -o):删 temp 输出
        if not keep_local_output and not args.keep_work:
            Path(result_path).unlink(missing_ok=True)

    # ── 清理中间产物(除非 --keep-work)──
    if not args.keep_work:
        import shutil
        shutil.rmtree(work_dir, ignore_errors=True)
        print(f"\n  → 已清理中间产物(--keep-work 可保留)", file=sys.stderr)

    print(f"\n✅ 转换完成: {result_path}", file=sys.stderr)

    # ── 输出 JSON ──
    result = {
        "status": "success",
        "output": result_path,
        "stats": {
            "spans_total": len(spans),
            "align_matched": stats["matched"],
            "align_total": stats["total"],
            "align_rate": round(rate, 1),
        },
    }
    if uploaded_info is not None:
        result["minio"] = {
            "downloaded_key": downloaded_key,
            "uploaded_key": uploaded_info["key"],
            "uploaded_url": uploaded_info["url"],
            "uploaded_size": uploaded_info["size"],
            "metadata_written": metadata,
        }
    print(json.dumps(result, ensure_ascii=False, indent=2))

    # ── 成功后清 temp 目录(失败则保留供排查)──
    if not args.keep_work:
        import shutil
        for d in temp_dirs:
            shutil.rmtree(d, ignore_errors=True)
```

- [ ] **Step 3: 修改 CLI flag 定义(在 main() 的 convert 子解析器)**

将 `main()` 中 convert 子解析器部分(原 L277-294)替换为:

```python
    # ── convert ──
    p_conv = sub.add_parser("convert", help="一键全流程:PDF → DOCX")
    p_conv.add_argument("pdf", nargs="?", default=None,
                        help="输入 PDF 路径(本地模式);MinIO 模式省略")
    p_conv.add_argument("-o", "--output", default=None,
                        help="输出 docx 路径;MinIO-only 模式可省略")
    p_conv.add_argument("--work-dir", default=None,
                        help="中间产物目录(默认在输出旁边)")
    p_conv.add_argument("--keep-work", action="store_true",
                        help="保留中间产物(默认清理)")
    p_conv.add_argument("-m", "--method", default="auto",
                        choices=["auto", "txt", "ocr"],
                        help="解析方法(auto=自动判断)")
    p_conv.add_argument("-l", "--lang", default="ch", help="语言")
    p_conv.add_argument("--no-formula", action="store_true",
                        help="关闭公式识别(加快速度)")
    p_conv.add_argument("--no-table", action="store_true",
                        help="关闭表格识别(加快速度)")
    p_conv.add_argument("--iou", type=float, default=0.3,
                        help="bbox 对齐 IoU 阈值(默认 0.3)")
    # MinIO 输入
    p_conv.add_argument("--minio-input", default=None,
                        help="MinIO 输入 JSON(内联;'-' 表 stdin)")
    p_conv.add_argument("--minio-input-file", default=None,
                        help="MinIO 输入 JSON 文件路径")
    # MinIO 输出
    p_conv.add_argument("--minio-output", default=None,
                        help="MinIO 输出 JSON(内联;'-' 表 stdin)")
    p_conv.add_argument("--minio-output-file", default=None,
                        help="MinIO 输出 JSON 文件路径")
    p_conv.set_defaults(func=cmd_convert)
```

- [ ] **Step 4: 冒烟测试 — 本地模式回归**

```bash
cd skills/pdf2docx-skill
MINERU_MODEL_SOURCE=modelscope PYTHONUTF8=1 uv run python scripts/pdf2docx.py convert "示例文件1.pdf" -o "_test_output/regression.docx" 2>&1 | tail -10
```
期望: 转换成功,stdout JSON 有 `output` 和 `stats`,**无** `minio` 字段(向后兼容)。

- [ ] **Step 5: 冒烟测试 — flag 校验(无 env)**

```bash
cd skills/pdf2docx-skill
# 不设 env,跑 MinIO 模式,期望 exit 2
PYTHONUTF8=1 uv run python scripts/pdf2docx.py convert \
  --minio-input '{"object_key":"generation/demo/x.pdf"}' \
  --minio-output '{"object_key":"generation/demo/x.docx"}' 2>&1 | tail -5
echo "exit=$?"
```
期望: stderr 含 `MinIOConfigError: ...缺失环境变量...`,stdout 有 `{"status":"error","code":2,...}`,exit=2。

- [ ] **Step 6: 冒烟测试 — 缺输入报错**

```bash
cd skills/pdf2docx-skill
PYTHONUTF8=1 uv run python scripts/pdf2docx.py convert 2>&1 | tail -3
echo "exit=$?"
```
期望: `必须提供 PDF 路径...或 --minio-input`,exit=5。

---

## Task 3: 支撑文件(requirements / setup.sh / env.check / SKILL.md)

**Files:**
- Modify: `skills/pdf2docx-skill/requirements.txt`
- Modify: `skills/pdf2docx-skill/setup.sh`
- Modify: `skills/pdf2docx-skill/scripts/pdf2docx.py`(env.check 段)
- Modify: `skills/pdf2docx-skill/SKILL.md`

- [ ] **Step 1: requirements.txt 加 minio SDK**

在 `requirements.txt` 末尾加:

```
# MinIO 对象存储 SDK —— 下载源文件 / 上传转换结果
minio>=7.2.0
```

- [ ] **Step 2: setup.sh 加 minio 检查**

在 `setup.sh` 的 `PY_PKGS` 数组(L58-61)中加一项:

```bash
PY_PKGS=(
    "fitz:PyMuPDF"
    "docx:python-docx"
    "minio:minio"
)
```

并在 MinerU 检查块之后(L93 之后)插入 MinIO env 提示块:

```bash
# MinIO 环境变量(可选,仅 MinIO 模式需要)
echo ""
echo "--- MinIO 环境变量(MinIO 模式必填,本地模式不需要)---"
MINIO_REQUIRED=("MINIO_ENDPOINT" "MINIO_ACCESS_KEY" "MINIO_SECRET_KEY" "MINIO_BUCKET")
MINIO_MISSING=()
for var in "${MINIO_REQUIRED[@]}"; do
    if [ -z "${!var:-}" ]; then
        MINIO_MISSING+=("$var")
    fi
done
if [ ${#MINIO_MISSING[@]} -gt 0 ]; then
    warn "MinIO env 未设置: ${MINIO_MISSING[*]} (仅影响 --minio-input/--minio-output)"
    info "本地路径模式不需要 MinIO env"
else
    ok "MinIO env 已配置(本地模式可忽略)"
fi
```

- [ ] **Step 3: pdf2docx.py env.check 加 minio 包检查**

在 `cmd_env_check` 函数的 `py_deps` 列表(L58-61)中加:

```python
    py_deps = [
        ("fitz", "PyMuPDF", "pip install PyMuPDF"),
        ("docx", "python-docx", "pip install python-docx"),
        ("minio", "minio", "pip install minio>=7.2.0"),
    ]
```

- [ ] **Step 4: SKILL.md 加 MinIO 用法文档**

在 `SKILL.md` 的"快速开始"之后("处理流程"之前),插入新章节:

```markdown
## MinIO 集成(从对象存储读写)

支持直接从 MinIO 下载源 PDF、转换后上传 DOCX。MinIO 连接信息全部走环境变量,
缺失即报错(本地路径模式不需要)。

### 环境变量(必填,MinIO 模式)

| 变量 | 说明 |
|------|------|
| `MINIO_ENDPOINT` | MinIO 地址,如 `http://127.0.0.1:9000` |
| `MINIO_ACCESS_KEY` | 访问密钥 |
| `MINIO_SECRET_KEY` | 秘密密钥 |
| `MINIO_BUCKET` | bucket 名 |
| `MINIO_PUBLIC_URL` | 可选,对外 URL;配了才能用 URL 形式传 `--minio-input` |

### 用法

```bash
# 1. 内联 JSON(短 spec)
python3 "$SKILL_DIR/scripts/pdf2docx.py" convert \
  --minio-input '{"object_key":"generation/demo/x.pdf"}' \
  --minio-output '{"object_key":"generation/demo/x.docx"}'

# 2. 文件 JSON(长 spec,避 shell 转义)
python3 "$SKILL_DIR/scripts/pdf2docx.py" convert \
  --minio-input-file input.json \
  --minio-output-file output.json

# 3. stdin(agent 管道注入)
cat input.json | python3 "$SKILL_DIR/scripts/pdf2docx.py" convert \
  --minio-input - --minio-output-file output.json
```

### JSON Schema

**输入**(`--minio-input`):`url`/`object_key` 二选一,`metadata` 透传不落对象。
**输出**(`--minio-output`):`object_key` 可省略(自动推导),`metadata` 落对象 header。

详见 `docs/2026-08-03-minio-integration-design.md`。
```

- [ ] **Step 5: 验证 setup.sh + env.check**

```bash
cd skills/pdf2docx-skill
bash setup.sh 2>&1 | grep -E "minio|MinIO" | head -10
PYTHONUTF8=1 uv run python scripts/pdf2docx.py env.check 2>&1 | grep -i minio
```
期望: setup.sh 有 `MinIO env` 检查块输出;env.check 有 `✓ minio`。

---

## Task 4: 端到端集成测试(真实 MinIO)

**Files:** 无新建(命令行验证)

**测试环境:**
- endpoint: `http://127.0.0.1:9000`
- 测试 PDF object key: `generation/demo/示例文件1.pdf`(URL: `http://domain/upload/tender/generation/demo/示例文件1.pdf`)
- 上传目标前缀: `generation/demo`

- [ ] **Step 1: 设环境变量并跑端到端转换**

```bash
cd skills/pdf2docx-skill
export MINIO_ENDPOINT="http://127.0.0.1:9000"
export MINIO_ACCESS_KEY="minio"
export MINIO_SECRET_KEY="password"
export MINIO_BUCKET="tender"
export MINIO_PUBLIC_URL="http://domain/upload"

MINERU_MODEL_SOURCE=modelscope PYTHONUTF8=1 uv run python scripts/pdf2docx.py convert \
  --minio-input '{"object_key":"generation/demo/示例文件1.pdf","metadata":{"file_id":"e2e-test"}}' \
  --minio-output '{"object_key":"generation/demo/示例文件1_e2e.docx","metadata":{"converter":"pdf2docx-skill"}}' 2>&1 | tail -20
```
期望:
- stderr 有 `MinIO 下载: generation/demo/示例文件1.pdf`、`MinIO 上传: 已启用`、`✅ MinIO 上传成功: generation/demo/示例文件1_e2e.docx`
- stdout JSON 含 `minio.uploaded_key`、`minio.metadata_written`

- [ ] **Step 2: 反查上传对象 metadata**

```bash
cd skills/pdf2docx-skill
MINIO_ENDPOINT="http://127.0.0.1:9000" MINIO_ACCESS_KEY="minio" \
MINIO_SECRET_KEY="password" MINIO_BUCKET="tender" \
PYTHONUTF8=1 uv run python -c "
import sys; sys.path.insert(0, 'scripts')
from minio_client import MinioClient
c = MinioClient()
info = c.stat('generation/demo/示例文件1_e2e.docx')
print('size:', info['size'])
print('content_type:', info['content_type'])
print('metadata:', info['metadata'])
"
```
期望:
- `size > 0`
- `content_type` 含 `wordprocessingml.document`
- `metadata` 含 `x-amz-meta-converter=pdf2docx-skill`、`x-amz-meta-converted_from=pdf`、`x-amz-meta-file_id=e2e-test`、`x-amz-meta-converted_at=...`

- [ ] **Step 3: 测试 URL 形式输入(若 MINIO_PUBLIC_URL 可达)**

```bash
cd skills/pdf2docx-skill
MINERU_MODEL_SOURCE=modelscope PYTHONUTF8=1 uv run python scripts/pdf2docx.py convert \
  --minio-input '{"url":"http://domain/upload/tender/generation/demo/示例文件1.pdf"}' \
  --minio-output '{"object_key":"generation/demo/示例文件1_e2e_url.docx"}' 2>&1 | tail -10
```
期望: URL 被反解为 object_key,转换上传成功。

- [ ] **Step 4: 回归本地模式**

```bash
cd skills/pdf2docx-skill
MINERU_MODEL_SOURCE=modelscope PYTHONUTF8=1 uv run python scripts/pdf2docx.py convert "示例文件1.pdf" -o "_test_output/final_regression.docx" 2>&1 | tail -5
```
期望: 本地模式零破坏,stdout JSON 无 `minio` 字段。

---

## Self-Review 记录

**Spec coverage:**
- §3 env 变量 → Task 1 `_load_config` + Task 3 setup.sh ✓
- §4 minio_client 接口 → Task 1 全覆盖 ✓
- §5 JSON schema → Task 2 `_load_minio_spec` + Task 4 测试 ✓
- §6 CLI flag → Task 2 Step 3 ✓
- §7 temp 文件管理 → Task 2 `temp_dirs` 清理逻辑 ✓
- §8 stdout JSON → Task 2 输出结构 ✓
- §9 退出码 → Task 2 各 sys.exit(N) ✓
- §10 依赖 → Task 3 Step 1 ✓
- §11 测试计划 → Task 1(单元)+ Task 4(端到端)✓
- §13 非目标 → 计划未引入 list/断点续传/连接池/common 目录 ✓

**Type consistency:**
- `MinioClient.download(source, dest_dir) -> (Path, dict)` — Task 1 定义,Task 2 调用一致 ✓
- `MinioClient.upload(local_path, key, content_type, metadata) -> dict` — 一致 ✓
- `_load_minio_spec` / `_print_error_json` / `_derive_output_key` / `_build_upload_metadata` — Task 2 定义即调用 ✓
- `MinioConfigError` — Task 1 定义,Task 2 import 一致 ✓

**Placeholder scan:** 无 TBD/TODO,所有步骤含具体代码或命令。
