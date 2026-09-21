#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$0")"
export VITE_APP_BASE="${VITE_APP_BASE:-/}"
export VITE_BACKEND_PROXY_URL="${VITE_BACKEND_PROXY_URL:-https://127.0.0.1:8443}"
export VITE_BACKEND_HOST="${VITE_BACKEND_HOST:-belle.iliadboxos.it}"
npm install
npm run dev -- --host 0.0.0.0
