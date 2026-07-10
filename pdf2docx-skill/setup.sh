#!/usr/bin/env bash
# ---
# name: pdf2docx-setup
# author: aowu
# version: "1.0"
# description: Environment setup for the PDF-to-DOCX skill. Checks and installs all required dependencies.
# ---
#
# 检查并安装 PDF 转 DOCX Skill 所需依赖：
#   Python 3.10+、PyMuPDF、python-docx、MinerU 2.0（含模型权重）
set -euo pipefail

RED='\033[0;31m'; GREEN='\033[0;32m'; YELLOW='\033[1;33m'; BLUE='\033[0;34m'; NC='\033[0m'
ok()   { echo -e "  ${GREEN}✓${NC} $1"; }
fail() { echo -e "  ${RED}✗${NC} $1"; }
warn() { echo -e "  ${YELLOW}○${NC} $1"; }
info() { echo -e "  ${BLUE}→${NC} $1"; }

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

echo "============================================"
echo "  PDF-to-DOCX Skill — Environment Setup"
echo "============================================"
echo ""

# ── 1. Python 3.10+ ──
echo "--- Python ---"
if command -v python3 &>/dev/null; then
    PY_VER=$(python3 -version 2>&1 || python3 --version 2>&1)
    PY_MAJOR=$(python3 -c "import sys; print(sys.version_info[0])" 2>/dev/null || echo "0")
    PY_MINOR=$(python3 -c "import sys; print(sys.version_info[1])" 2>/dev/null || echo "0")
    if [ "$PY_MAJOR" -ge 3 ] && [ "$PY_MINOR" -ge 10 ]; then
        ok "python3 ($PY_VER)"
    else
        fail "python3 $PY_VER 版本过低，需要 3.10+"
        info "Install: https://www.python.org/downloads/"
    fi
else
    fail "python3 not found"
    info "Install: https://www.python.org/downloads/"
fi

# ── 2. pip ──
echo ""
echo "--- pip ---"
if python3 -m pip --version &>/dev/null 2>&1; then
    PIP_VER=$(python3 -m pip --version 2>/dev/null | head -1)
    ok "pip ($PIP_VER)"
else
    fail "pip not found"
    info "Install: python3 -m ensurepip --upgrade"
fi

# ── 3. Python packages ──
echo ""
echo "--- Python Packages ---"
# 格式: import名:包名
PY_PKGS=(
    "fitz:PyMuPDF"
    "docx:python-docx"
)

MISSING_PY=()
for entry in "${PY_PKGS[@]}"; do
    mod="${entry%%:*}"
    pkg="${entry##*:}"
    if python3 -c "import $mod" 2>/dev/null; then
        ver=$(python3 -c "import $mod; print(getattr($mod, '__version__', 'installed'))" 2>/dev/null)
        ok "$pkg ($ver)"
    else
        fail "$pkg not installed"
        MISSING_PY+=("$pkg")
    fi
done

# MinerU 单独检查（import 名为 mineru，历史名 magic_pdf）
echo ""
echo "--- MinerU (版面分析) ---"
if python3 -c "import mineru" 2>/dev/null; then
    ok "mineru (installed)"
else
    fail "mineru not installed"
    MISSING_PY+=("mineru[core]==3.0.4")
    info "Install: uv pip install \"mineru[core]==3.0.4\""
fi

# albumentations —— MinerU 3.0.4 [core] 打包遗漏的依赖
if python3 -c "import albumentations" 2>/dev/null; then
    ok "albumentations (installed)"
else
    fail "albumentations not installed (MinerU 3.0.4 必需)"
    MISSING_PY+=("albumentations")
fi

# ⚠️ 关键：MINERU_MODEL_SOURCE 环境变量检查
echo ""
echo "--- MINERU_MODEL_SOURCE 环境变量（国内必须）---"
if [ "$MINERU_MODEL_SOURCE" = "modelscope" ]; then
    ok "MINERU_MODEL_SOURCE=modelscope"
else
    warn "MINERU_MODEL_SOURCE 未设置为 modelscope"
    info "国内环境不设置会导致 MinerU 连 HuggingFace 卡死！"
    info "永久设置："
    info "  bash:  echo 'export MINERU_MODEL_SOURCE=modelscope' >> ~/.bashrc"
    info "  PowerShell: [Environment]::SetEnvironmentVariable('MINERU_MODEL_SOURCE','modelscope','User')"
fi

if [ ${#MISSING_PY[@]} -gt 0 ]; then
    echo ""
    if [ -t 0 ]; then
        read -p "  Install missing Python packages? [Y/n] " -n 1 -r REPLY
        echo ""
        REPLY=${REPLY:-Y}
    else
        warn "Non-interactive mode — skipping auto-install. Run interactively or install manually."
        REPLY=N
    fi
    if [[ ! $REPLY =~ ^[Nn]$ ]]; then
        python3 -m pip install -q "${MISSING_PY[@]}" 2>/dev/null \
            || python3 -m pip install -q --user "${MISSING_PY[@]}" 2>/dev/null \
            || python3 -m pip install -q --break-system-packages "${MISSING_PY[@]}" 2>/dev/null \
            || { fail "pip install failed. Try manually: pip install ${MISSING_PY[*]}"; }
        ok "Installed: ${MISSING_PY[*]}"
        echo ""
        warn "如果刚装了 mineru，还需下载模型权重:"
        info "  MINERU_MODEL_SOURCE=modelscope mineru-models-download -s modelscope -m pipeline"
    fi
fi

# ── 4. MinerU 模型权重检查（软检查：仅提示，不阻塞） ──
echo ""
echo "--- MinerU 模型权重 ---"
# MinerU 模型默认缓存在 HuggingFace/ModelScope 缓存目录，路径因配置而异
# 这里只做提示性检查，不强制
if python3 -c "import mineru" 2>/dev/null || python3 -c "import magic_pdf" 2>/dev/null; then
    warn "请确认已运行过 mineru-models-download 下载模型权重（首次使用必需）"
    info "首次转换 PDF 时会自动下载，但建议提前执行以避免转换时等待"
else
    warn "MinerU 未安装，模型权重检查跳过"
fi

# ── Summary ──
echo ""
echo "============================================"
echo "  Setup complete."
echo "  Run 'python3 pdf2docx.py env.check' for detailed status."
echo "============================================"
