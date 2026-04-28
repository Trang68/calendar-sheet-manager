#!/usr/bin/env bash
set -Eeuo pipefail

APP_DIR="${APP_DIR:-/root/russweb/calendar-sheet}"
APP_NAME="${APP_NAME:-calendar-sheet-manager}"
PORT="${PORT:-8080}"

cd "$APP_DIR"

# Load NVM to make npm available
[[ -s "$HOME/.nvm/nvm.sh" ]] && \. "$HOME/.nvm/nvm.sh"

echo "[deploy] Installing dependencies..."
if [[ -f package-lock.json ]]; then
  npm ci --omit=dev
else
  npm install --production
fi

echo "[deploy] Starting/restarting app..."
if command -v pm2 >/dev/null 2>&1; then
  if pm2 describe "$APP_NAME" >/dev/null 2>&1; then
    pm2 restart "$APP_NAME" --update-env
  else
    pm2 start src/server.js --name "$APP_NAME"
  fi
  pm2 save || true
else
  pkill -f "node src/server.js" || true
  nohup node src/server.js > app.log 2>&1 &
fi

echo "[deploy] Health check..."
curl -fsS "http://127.0.0.1:${PORT}/api/health" --retry 20 --retry-all-errors --retry-delay 2 >/dev/null

echo "[deploy] OK"
