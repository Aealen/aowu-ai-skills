#!/usr/bin/env python3
"""
env_setup.py —— 共享缓存初始化（幂等，沙箱环境）

职责：在共享目录中构建/复用 Python venv 和 MinerU 模型缓存，
避免每个沙箱冷启动重复下载 ~4GB 依赖和模型。

环境变量（平台注入）：
  SHARED_DIR   共享目录根。未设置 → 降级本地模式（现有 uv 流程）。

共享目录布局（约定，SKILL 用子目录避免多 Skill 冲突）：
  $SHARED_DIR/pdf2docx/
    venv-<hash>/       完整 venv，hash = requirements.txt 内容哈希
    modelscope/        MinerU 模型缓存（MODELSCOPE_CACHE 指向）
    cache/uv/          uv 下载缓存（加速 venv 构建）
    manifest.json      版本指针 {"venv_hash": ..., "skill_version": ...}
    .lock              flock 锁文件（内容为空，锁状态在内核）

并发安全（多个沙箱同时首启）：
  - flock 独占锁串行化初始化（内核级，bind mount 跨容器有效）
  - double-check：抢到锁后重读 manifest，别人已建好则直接用
  - venv 带 hash 目录，构建期间无读者，天然原子；manifest 最后原子切换
  - 构建失败 → 删除半成品目录，放锁；下次启动重试

保持最新：
  Skill 更新（requirements.txt 变化）→ hash 变化 → 下次启动自动重建新 venv
  （旧 venv 保留最近 N 个，超出清理）

用法（被 SKILL.md 缓存协议调用）:
  python3 scripts/env_setup.py
  成功 → stderr 打印就绪信息，stdout 打印:
      SHARED_VENV_PY=/shared/pdf2docx/venv-<hash>/bin/python
      MODELSCOPE_CACHE=/shared/pdf2docx/modelscope
  后续命令用 SHARED_VENV_PY 执行。
  降级（无 SHARED_DIR / 平台无 flock）→ stderr 警告，返回 0，走本地流程。
"""

from __future__ import annotations

import hashlib
import json
import os
import shutil
import subprocess
import sys
import time
from pathlib import Path

# ═══════════════════════════════════════════════════════════════
#  常量
# ═══════════════════════════════════════════════════════════════

SHARED_ENV_VAR = "SHARED_DIR"      # 平台注入的共享目录环境变量
SKILL_SUBDIR = "pdf2docx"          # 共享目录内本 Skill 的子目录
SKILL_VERSION = "1.0"
LOCK_TIMEOUT_SEC = 1800            # 等待锁总超时 30min
POLL_INTERVAL_SEC = 5              # 锁轮询间隔
KEEP_OLD_VENVS = 2                 # 保留 venv 版本数（含当前）
MANIFEST_NAME = "manifest.json"
LOCK_NAME = ".lock"


# ═══════════════════════════════════════════════════════════════
#  基础工具
# ═══════════════════════════════════════════════════════════════

def _skill_dir() -> Path:
    """SKILL 根目录（env_setup.py 的上一级）。"""
    return Path(__file__).resolve().parent.parent


def _requirements_hash() -> str:
    """requirements.txt 内容哈希 —— venv 版本指纹。"""
    req = _skill_dir() / "requirements.txt"
    h = hashlib.sha256()
    h.update(req.read_bytes())
    return h.hexdigest()[:12]


def _read_manifest(shared_root: Path) -> dict | None:
    p = shared_root / MANIFEST_NAME
    if not p.exists():
        return None
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return None


def _write_manifest(shared_root: Path, manifest: dict) -> None:
    """manifest 原子写入：先写 tmp 再 rename（读方永远看到完整版本）。"""
    tmp = shared_root / (MANIFEST_NAME + ".tmp")
    tmp.write_text(json.dumps(manifest, ensure_ascii=False, indent=2),
                   encoding="utf-8")
    os.replace(tmp, shared_root / MANIFEST_NAME)


def _run(cmd: list, env: dict | None = None) -> None:
    print(f"  $ {' '.join(cmd)}", file=sys.stderr)
    subprocess.run(cmd, env=env, check=True)


# ═══════════════════════════════════════════════════════════════
#  构建（锁内执行）
# ═══════════════════════════════════════════════════════════════

def _build_venv(venv_dir: Path) -> None:
    """在共享目录构建完整 venv（hash 目录，构建期间无读者）。"""
    cache_dir = venv_dir.parent / "cache" / "uv"
    cache_dir.mkdir(parents=True, exist_ok=True)
    env = dict(os.environ)
    env["UV_CACHE_DIR"] = str(cache_dir)
    req = str(_skill_dir() / "requirements.txt")
    venv_py = str(venv_dir / "bin" / "python")

    try:
        _run(["uv", "venv", str(venv_dir)], env=env)
        _run(["uv", "pip", "install", "--python", venv_py, "-r", req], env=env)
    except FileNotFoundError:
        # 沙箱无 uv → 回退 venv + pip
        print("  ○ uv 不可用，回退 venv+pip", file=sys.stderr)
        _run([sys.executable, "-m", "venv", str(venv_dir)])
        _run([str(venv_dir / "bin" / "pip"), "install", "-r", req], env=env)


def _build_models(venv_dir: Path, shared_root: Path) -> None:
    """下载 MinerU 模型到共享目录（modelscope 自身幂等，已存在则跳过）。"""
    env = dict(os.environ)
    env["MODELSCOPE_CACHE"] = str(shared_root / "modelscope")
    env["MINERU_MODEL_SOURCE"] = "modelscope"
    dl = str(venv_dir / "bin" / "mineru-models-download")
    _run([dl, "-s", "modelscope", "-m", "pipeline"], env=env)


def _cleanup_old_venvs(shared_root: Path, current_hash: str) -> None:
    """manifest 切换后清理旧 venv，保留最近 KEEP_OLD_VENVS 个。"""
    venvs = sorted(shared_root.glob("venv-*"),
                   key=lambda p: p.stat().st_mtime, reverse=True)
    keep = {f"venv-{current_hash}"}
    for v in venvs:
        if v.name == f"venv-{current_hash}":
            continue
        if len(keep) < KEEP_OLD_VENVS:
            keep.add(v.name)
            continue
        print(f"  ○ 清理旧缓存: {v.name}", file=sys.stderr)
        shutil.rmtree(v, ignore_errors=True)


# ═══════════════════════════════════════════════════════════════
#  锁
# ═══════════════════════════════════════════════════════════════

def _with_exclusive_lock(lock_path: Path, timeout: float, fn) -> None:
    """
    独占锁临界区（跨容器，内核级互斥）。
    抢不到锁 → 轮询等待；总超时 → TimeoutError。
    fcntl 仅 Unix（Docker 容器可用），Windows 由调用方先行降级。
    """
    import fcntl

    fd = os.open(lock_path, os.O_CREAT | os.O_RDWR)
    try:
        deadline = time.monotonic() + timeout
        while True:
            try:
                fcntl.flock(fd, fcntl.LOCK_EX | fcntl.LOCK_NB)  # type: ignore[attr-defined]
                break
            except OSError:
                if time.monotonic() > deadline:
                    raise TimeoutError(
                        f"等待共享锁超时（{int(timeout)}s）: {lock_path}")
                time.sleep(POLL_INTERVAL_SEC)
        try:
            fn()
        finally:
            fcntl.flock(fd, fcntl.LOCK_UN)  # type: ignore[attr-defined]
    finally:
        os.close(fd)


# ═══════════════════════════════════════════════════════════════
#  主流程
# ═══════════════════════════════════════════════════════════════

def _cache_hit(shared_root: Path, req_hash: str) -> Path | None:
    """缓存完整且版本匹配 → 返回共享 venv 的 python 路径；否则 None。"""
    manifest = _read_manifest(shared_root)
    if not manifest or manifest.get("venv_hash") != req_hash:
        return None
    venv_py = shared_root / f"venv-{req_hash}" / "bin" / "python"
    if not venv_py.exists():
        return None
    # 模型完整性软检查：modelscope 缓存非空即视为就绪（modelscope 自身幂等）
    model_dir = shared_root / "modelscope"
    if not model_dir.exists() or not any(model_dir.rglob("*")):
        return None
    return venv_py


def _print_ready(shared_root: Path, venv_py: Path, manifest: dict | None) -> None:
    """输出 SKILL.md 协议所需的环境变量（stdout，供 Agent 捕获）。"""
    skill_ver = manifest.get("skill_version", "?") if manifest else "?"
    print(f"✅ 共享缓存就绪 (skill v{skill_ver})", file=sys.stderr)
    print(f"SHARED_VENV_PY={venv_py}")
    print(f"MODELSCOPE_CACHE={shared_root / 'modelscope'}")


def main() -> int:
    shared_dir = os.environ.get(SHARED_ENV_VAR)
    if not shared_dir:
        print("○ 未设置 SHARED_DIR，使用本地模式（现有 uv 流程）", file=sys.stderr)
        return 0

    try:
        import fcntl  # type: ignore[attr-defined]  # Unix-only，Windows 已提前降级
    except ImportError:
        print("○ 当前平台不支持共享锁（flock），使用本地模式", file=sys.stderr)
        return 0

    shared_root = Path(shared_dir) / SKILL_SUBDIR
    shared_root.mkdir(parents=True, exist_ok=True)
    req_hash = _requirements_hash()

    # 命中缓存 → 秒级完成
    venv_py = _cache_hit(shared_root, req_hash)
    if venv_py:
        _print_ready(shared_root, venv_py, _read_manifest(shared_root))
        return 0

    # 未命中 → 锁内构建（double-check 防重复下载）
    def build():
        manifest = _read_manifest(shared_root)
        if manifest and manifest.get("venv_hash") == req_hash:
            return  # 排队期间别人已建好
        venv_dir = shared_root / f"venv-{req_hash}"
        if venv_dir.exists():
            # 上次构建失败残留 → 清掉重来
            shutil.rmtree(venv_dir, ignore_errors=True)
        try:
            print("🔨 构建共享缓存（首次或版本变更）...", file=sys.stderr)
            _build_venv(venv_dir)
            _build_models(venv_dir, shared_root)
            _write_manifest(shared_root, {
                "skill_version": SKILL_VERSION,
                "requirements_hash": req_hash,
                "venv_hash": req_hash,
                "model_status": "ready",
                "updated_at": time.strftime("%Y-%m-%d %H:%M:%S"),
            })
            _cleanup_old_venvs(shared_root, req_hash)
            print("✅ 共享缓存构建完成", file=sys.stderr)
        except Exception:
            shutil.rmtree(venv_dir, ignore_errors=True)  # 半成品清掉
            raise

    try:
        _with_exclusive_lock(shared_root / LOCK_NAME, LOCK_TIMEOUT_SEC, build)
    except TimeoutError as e:
        print(f"✗ {e}", file=sys.stderr)
        return 1

    venv_py = shared_root / f"venv-{req_hash}" / "bin" / "python"
    _print_ready(shared_root, venv_py, _read_manifest(shared_root))
    return 0


if __name__ == "__main__":
    sys.exit(main())
