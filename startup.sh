#!/bin/bash
set -e

# Azure App Service startup script

# pip が使えるか確認
if ! python3 -m pip --version 2>/dev/null; then
  echo "[startup] ERROR: python3 -m pip is not available. Cannot install python-pptx."
  exit 1
fi

# python-pptx / lxml が未インストールの場合のみインストール
if ! python3 -c "import pptx, lxml" 2>/dev/null; then
  echo "[startup] Installing python-pptx and lxml..."
  python3 -m pip install --quiet --user python-pptx lxml
  echo "[startup] Installation done."
else
  echo "[startup] python-pptx and lxml already installed. Skipping."
fi

# Start Next.js standalone server
node /home/site/wwwroot/server.js
