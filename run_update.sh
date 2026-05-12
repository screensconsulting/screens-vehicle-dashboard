#!/bin/bash
# ════════════════════════════════════════════════════════════════
#  Screens Studio — Dashboard Weekly Updater (Mac / Linux)
#  Run manually: bash run_update.sh
#  Schedule with cron (every Monday 08:00):
#    crontab -e
#    Add line: 0 8 * * 1 /bin/bash /path/to/run_update.sh
# ════════════════════════════════════════════════════════════════

set -euo pipefail
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

echo ""
echo "  ┌──────────────────────────────────────────────────────────┐"
echo "  │  Screens Studio — Dashboard Updater                      │"
echo "  │  $(date '+%Y-%m-%d %H:%M:%S')"
echo "  └──────────────────────────────────────────────────────────┘"
echo ""

# ── Check Python ─────────────────────────────────────────────────
PYTHON=$(which python3 2>/dev/null || which python 2>/dev/null)
if [ -z "$PYTHON" ]; then
    echo "ERROR: Python not found. Install with: brew install python3"
    exit 1
fi

# ── Install packages if missing ──────────────────────────────────
"$PYTHON" -c "import openpyxl, requests, bs4" 2>/dev/null || {
    echo "Installing required packages…"
    "$PYTHON" -m pip install openpyxl requests beautifulsoup4 --quiet
}

# ── Run updater ──────────────────────────────────────────────────
echo "Running update_dashboard.py…"
"$PYTHON" "$SCRIPT_DIR/update_dashboard.py"

echo ""
echo "  ✓  Dashboard updated. Open dashboard.html in your browser."
echo ""
