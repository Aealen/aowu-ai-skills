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
