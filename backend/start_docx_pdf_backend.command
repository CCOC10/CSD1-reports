#!/bin/zsh
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/.." && pwd)"

cd "$PROJECT_ROOT"
echo "Starting CSD1 DOCX/PDF backend on http://127.0.0.1:8765"
exec python3 "$SCRIPT_DIR/docx_pdf_backend.py" serve --host 127.0.0.1 --port 8765
