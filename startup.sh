#!/bin/bash
# Azure App Service startup script
# NOTE: Python install failure must NOT prevent the app from starting.

PDIR=/home/site/python-packages
PIP_CMD=/home/.local/bin/pip

# python-pptx が PDIR に入っているか確認
if ! PYTHONPATH="$PDIR" python3 -c "import pptx, lxml" 2>/dev/null; then
  echo "[startup] python-pptx not found in $PDIR, installing..."
  mkdir -p "$PDIR"

  # pip が /home/.local/bin になければブートストラップ
  if [ ! -f "$PIP_CMD" ]; then
    echo "[startup] pip not found, bootstrapping with get-pip.py..."
    curl -sS https://bootstrap.pypa.io/get-pip.py \
      | python3 - --user --break-system-packages 2>/dev/null \
      && echo "[startup] pip bootstrapped." \
      || echo "[startup] WARNING: pip bootstrap failed. PPT edit unavailable."
  fi

  if [ -f "$PIP_CMD" ]; then
    "$PIP_CMD" install --quiet --target="$PDIR" python-pptx lxml \
      && echo "[startup] python-pptx installed." \
      || echo "[startup] WARNING: Install failed. PPT edit unavailable."
  else
    echo "[startup] WARNING: pip not available. PPT edit unavailable."
  fi
else
  echo "[startup] python-pptx already available in $PDIR. Skipping."
fi

# PYTHONPATH を node プロセスに継承させる
export PYTHONPATH="$PDIR"

# Start Next.js standalone server
exec node /home/site/wwwroot/server.js
