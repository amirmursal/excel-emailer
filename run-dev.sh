#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$ROOT_DIR"

PORT="${PORT:-5004}"
lsof -ti ":$PORT" | xargs kill -9 2>/dev/null || true
sleep 0.5

if [[ -x "$HOME/.asdf/installs/python/3.11.2/bin/python3" ]]; then
  exec "$HOME/.asdf/installs/python/3.11.2/bin/python3" app.py
elif [[ -x "/usr/bin/python3" ]]; then
  exec /usr/bin/python3 app.py
else
  exec python3 app.py
fi
