#!/bin/bash
# Azure App Service startup script
# NOTE: Python install failure must NOT prevent the app from starting.

# python-pptx / lxml のインストールを試みる（失敗しても起動は続行）
if python3 -m pip --version >/dev/null 2>&1; then
  if ! python3 -c "import pptx, lxml" 2>/dev/null; then
    echo "[startup] Installing python-pptx and lxml..."
    python3 -m pip install --quiet --user python-pptx lxml \
      && echo "[startup] Installation done." \
      || echo "[startup] WARNING: Installation failed. PPT edit feature will be unavailable."
  else
    echo "[startup] python-pptx and lxml already installed. Skipping."
  fi
else
  echo "[startup] WARNING: python3 -m pip not available. PPT edit feature will be unavailable."
fi

# Start Next.js standalone server
exec node /home/site/wwwroot/server.js
