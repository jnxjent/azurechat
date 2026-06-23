#!/bin/bash
# Azure App Service startup script
# NOTE: Python install failure must NOT prevent the app from starting.

PDIR=/home/site/python-packages
PIP_CMD=/home/.local/bin/pip

install_pip_if_needed() {
  # ファイル存在だけでなく実際に動くかも確認（/home マウント由来の壊れたバイナリ対策）
  if [ ! -f "$PIP_CMD" ] || ! "$PIP_CMD" --version &>/dev/null; then
    echo "[startup] pip not found or broken, bootstrapping with get-pip.py..."
    curl -sS https://bootstrap.pypa.io/get-pip.py \
      | python3 - --user --break-system-packages 2>/dev/null \
      && echo "[startup] pip bootstrapped." \
      || echo "[startup] WARNING: pip bootstrap failed."
  fi
}

mkdir -p "$PDIR"

# ── Step 0: LibreOffice（PPTX→PDF変換用） ────────────────────────────────
if ! command -v libreoffice &>/dev/null && ! command -v soffice &>/dev/null; then
  echo "[startup] Installing LibreOffice..."
  apt-get update -qq && apt-get install -y --no-install-recommends \
    libreoffice-impress libreoffice-writer libreoffice-calc fonts-noto-cjk \
    2>/dev/null \
    && echo "[startup] LibreOffice installed." \
    || echo "[startup] WARNING: LibreOffice install failed (non-fatal)."
else
  echo "[startup] LibreOffice already available."
fi

# xlsx→PDF変換には libreoffice-calc が必須（既存インストールに含まれていない場合に追加）
if ! dpkg -s libreoffice-calc &>/dev/null 2>&1; then
  echo "[startup] libreoffice-calc not found, installing..."
  apt-get install -y --no-install-recommends libreoffice-calc 2>/dev/null \
    && echo "[startup] libreoffice-calc installed." \
    || echo "[startup] WARNING: libreoffice-calc install failed (non-fatal)."
else
  echo "[startup] libreoffice-calc already available."
fi

# ── Step 1: コア Office 編集ライブラリ ──────────────────────────────────
# python-pptx / openpyxl / docx / pdfplumber / Doc Intelligence SDK
install_pip_if_needed

MISSING=""
PYTHONPATH="$PDIR" python3 -c "import pptx"                        2>/dev/null || MISSING="$MISSING python-pptx lxml"
PYTHONPATH="$PDIR" python3 -c "import openpyxl"                    2>/dev/null || MISSING="$MISSING openpyxl xlrd"
PYTHONPATH="$PDIR" python3 -c "import docx"                        2>/dev/null || MISSING="$MISSING python-docx"
PYTHONPATH="$PDIR" python3 -c "import pdfplumber"                  2>/dev/null || MISSING="$MISSING pdfplumber"
PYTHONPATH="$PDIR" python3 -c "import fitz"                        2>/dev/null || MISSING="$MISSING pymupdf"
PYTHONPATH="$PDIR" python3 -c "import azure.ai.documentintelligence" 2>/dev/null || MISSING="$MISSING azure-ai-documentintelligence"
PYTHONPATH="$PDIR" python3 -c "import pdf2docx"                    2>/dev/null || MISSING="$MISSING pdf2docx"
PYTHONPATH="$PDIR" python3 -c "import matplotlib"                  2>/dev/null || MISSING="$MISSING matplotlib"
PYTHONPATH="$PDIR" python3 -c "import PIL"                         2>/dev/null || MISSING="$MISSING pillow"
PYTHONPATH="$PDIR" python3 -c "import cairosvg"                    2>/dev/null || MISSING="$MISSING cairosvg"

if [ -n "$MISSING" ]; then
  echo "[startup] Installing missing packages:$MISSING"
  if [ -f "$PIP_CMD" ]; then
    "$PIP_CMD" install --quiet --target="$PDIR" $MISSING \
      && echo "[startup] Packages installed." \
      || echo "[startup] WARNING: Some packages failed to install."
  else
    echo "[startup] WARNING: pip not found, cannot install packages."
  fi
else
  echo "[startup] All Python packages already available. Skipping."
fi

# PYTHONPATH を node プロセスに継承させる
export PYTHONPATH="$PDIR"

# Start Next.js standalone server
exec node /home/site/wwwroot/server.js
