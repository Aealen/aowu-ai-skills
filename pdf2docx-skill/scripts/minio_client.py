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
        client_factory: 可选 callable(endpoint, access_key, secret_key) -> client。
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
        """完整 URL -> object_key。需 MINIO_PUBLIC_URL 已配。"""
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
