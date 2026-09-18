#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$0")"
if [[ ! -d node_modules ]]; then
  echo "Устанавливаю зависимости…"
  npm ci --omit=dev 2>/dev/null || npm install --omit=dev
fi
export PORT="${PORT:-3847}"
echo "Старт на порту $PORT → http://127.0.0.1:${PORT}/karta-dnya"
exec node server.js
