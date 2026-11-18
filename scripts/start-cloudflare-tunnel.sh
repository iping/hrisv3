#!/usr/bin/env bash
# Expose local HRIS (port 5055) via Cloudflare Tunnel.
# Quick URL (no account):   ./scripts/start-cloudflare-tunnel.sh
# Named domain (after login + setup): ./scripts/start-cloudflare-tunnel.sh named

set -euo pipefail
APP_URL="${APP_URL:-http://127.0.0.1:5055}"
MODE="${1:-quick}"

if ! command -v cloudflared >/dev/null 2>&1; then
  echo "Install cloudflared first: brew install cloudflared" >&2
  exit 1
fi

if ! curl -fsS "$APP_URL" >/dev/null 2>&1; then
  echo "HRIS is not reachable at $APP_URL" >&2
  echo "Start it first: .venv/bin/python app.py" >&2
  exit 1
fi

if [[ "$MODE" == "named" ]]; then
  if [[ ! -f "$HOME/.cloudflared/config.yml" ]]; then
    echo "Missing ~/.cloudflared/config.yml" >&2
    echo "After 'cloudflared tunnel login', create a named tunnel and config for hris.product.biz.id" >&2
    exit 1
  fi
  exec cloudflared tunnel run
fi

echo "Starting quick Cloudflare tunnel → $APP_URL"
exec cloudflared tunnel --url "$APP_URL"
