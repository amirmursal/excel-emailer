#!/usr/bin/env bash
set -euo pipefail

# ---- Config ----
REPO_DIR="${REPO_DIR:-$PWD}"
PYTHON_VERSION="${PYTHON_VERSION:-3.11}"

echo "==> Repo dir: $REPO_DIR"
cd "$REPO_DIR"

# ---- Basic checks ----
command -v git >/dev/null 2>&1 || { echo "git is not installed"; exit 1; }

if command -v py >/dev/null 2>&1; then
  PY_CMD=(py -"$PYTHON_VERSION")
elif [[ -x "/usr/bin/python3" ]]; then
  # Prefer system Python on macOS to avoid asdf shim errors.
  PY_CMD=(/usr/bin/python3)
elif command -v python3 >/dev/null 2>&1 && python3 -V >/dev/null 2>&1; then
  PY_CMD=(python3)
elif command -v python >/dev/null 2>&1 && python -V >/dev/null 2>&1; then
  PY_CMD=(python)
else
  echo "Python not found in PATH"
  exit 1
fi

echo "==> Using Python command: ${PY_CMD[*]}"

# ---- Ensure this is a git repo ----
git rev-parse --is-inside-work-tree >/dev/null 2>&1 || {
  echo "Current folder is not a git repository"
  exit 1
}

# ---- Update code ----
echo "==> Fetching latest..."
git fetch --all --prune

CURRENT_BRANCH="$(git rev-parse --abbrev-ref HEAD)"
echo "==> Pulling latest on branch: $CURRENT_BRANCH"
git pull --ff-only origin "$CURRENT_BRANCH"

# ---- Create venv if missing ----
if [[ ! -d ".venv" ]]; then
  echo "==> Creating virtual environment..."
  if ! "${PY_CMD[@]}" -m venv .venv; then
    echo "==> venv creation failed, falling back to run-dev.sh"
    exec bash ./run-dev.sh
  fi
fi

# ---- Activate venv ----
# shellcheck disable=SC1091
if [[ -f ".venv/Scripts/activate" ]]; then
  # Windows (Git Bash)
  source .venv/Scripts/activate
elif [[ -f ".venv/bin/activate" ]]; then
  # macOS/Linux
  source .venv/bin/activate
else
  echo "==> venv activate script missing, falling back to run-dev.sh"
  exec bash ./run-dev.sh
fi

# ---- Install missing dependencies only ----
echo "==> Checking dependencies..."
python -m pip --version >/dev/null

declare -a REQUIRED_MODULES=("flask" "pandas" "openpyxl")
declare -a REQUIRED_PACKAGES=("flask" "pandas" "openpyxl")

# pywin32 is Windows-only.
if [[ "${OS:-}" == "Windows_NT" || "$(uname -s)" =~ ^(MINGW|MSYS|CYGWIN) ]]; then
  REQUIRED_MODULES+=("win32com")
  REQUIRED_PACKAGES+=("pywin32")
fi

MISSING_PACKAGES="$(python - <<'PY'
import importlib.util
checks = [
    ("flask", "flask"),
    ("pandas", "pandas"),
    ("openpyxl", "openpyxl"),
    ("win32com", "pywin32"),
]
import os, platform
is_windows = os.environ.get("OS") == "Windows_NT" or platform.system().lower().startswith("win")
missing = []
for module_name, package_name in checks:
    if package_name == "pywin32" and not is_windows:
        continue
    if importlib.util.find_spec(module_name) is None:
        missing.append(package_name)
print(" ".join(missing))
PY
)"

if [[ -n "$MISSING_PACKAGES" ]]; then
  echo "==> Installing missing packages: $MISSING_PACKAGES"
  python -m pip install $MISSING_PACKAGES
else
  echo "==> All required packages already installed. Skipping install."
fi

# ---- Run app ----
echo "==> Starting app on http://127.0.0.1:5004"
python app.py