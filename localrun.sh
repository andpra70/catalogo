#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$0")"
export VITE_APP_BASE="${VITE_APP_BASE:-/}"
npm install
npm run dev -- --host 0.0.0.0
