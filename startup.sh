#!/bin/bash
# Azure App Service startup script
# NOTE: Python install failure must NOT prevent the app from starting.

PDIR=/home/site/python-packages
PIP_CMD=/home/.local/bin/pip

install_pip_if_needed() {
  if [ ! -f "$PIP_CMD" ]; then
    echo "[startup] pip not found, bootstrapping with get-pip.py..."
    curl -sS https://bootstrap.pypa.io/get-pip.py \
      | python3 - --user --break-system-packages 2>/dev/null \
      && echo "[startup] pip bootstrapped." \
      || echo "[startup] WARNING: pip bootstrap failed."
  fi
}

mkdir -p "$PDIR"

# ── Step 1: コア Office 編集ライブラリ ──────────────────────────────────
# python-pptx / openpyxl / docx / pdfplumber が入っていなければインストール
if ! PYTHONPATH="$PDIR" python3 -c "import pptx, lxml, openpyxl, docx, pdfplumber, fitz" 2>/dev/null; then
  echo "[startup] Core libs not found, installing..."
  install_pip_if_needed
  if [ -f "$PIP_CMD" ]; then
    "$PIP_CMD" install --quiet --target="$PDIR" \
      python-pptx lxml openpyxl xlrd python-docx pdfplumber pymupdf \
      && echo "[startup] Core libs installed." \
      || echo "[startup] WARNING: Core lib install failed."
  fi
else
  echo "[startup] Core libs already available. Skipping."
fi

# ── Step 2: PaddleOCR（任意。ディスク容量に余裕がある場合のみ） ──────────
# paddleocr が入っていなければインストールを試みる。
# 失敗してもアプリ起動は継続する（pymupdf フォールバックあり）。
if ! PYTHONPATH="$PDIR" python3 -c "import paddleocr" 2>/dev/null; then
  echo "[startup] PaddleOCR not found, attempting install (may take several minutes)..."
  install_pip_if_needed
  if [ -f "$PIP_CMD" ]; then
    "$PIP_CMD" install --quiet --target="$PDIR" paddlepaddle paddleocr \
      && echo "[startup] PaddleOCR installed." \
      || echo "[startup] WARNING: PaddleOCR install failed. PDF→Excel will use pymupdf fallback."
  fi
else
  echo "[startup] PaddleOCR already available. Skipping."
fi

# PYTHONPATH を node プロセスに継承させる
export PYTHONPATH="$PDIR"

# Start Next.js standalone server
exec node /home/site/wwwroot/server.js
