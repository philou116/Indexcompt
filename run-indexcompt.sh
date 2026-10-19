#!/usr/bin/env bash
set -euo pipefail
APP_DIR="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
DATA_DIR="${XDG_DATA_HOME:-$HOME/.local/share}/IndexCompt"
mkdir -p "$DATA_DIR"
export INDEXCOMPT_DATA_DIR="$DATA_DIR"
if [[ ! -x "$APP_DIR/IndexCompt/IndexCompt" ]]; then
  echo "Binaire PyInstaller absent: $APP_DIR/IndexCompt/IndexCompt" >&2
  echo "Construisez d'abord avec build-bazzite-package.sh." >&2
  exit 2
fi
cd "$APP_DIR"
exec "$APP_DIR/IndexCompt/IndexCompt" "$@"
