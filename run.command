#!/bin/bash
# Double-click to start the Mouse Trap dashboard (opens in your browser).
cd "$(dirname "$0")"
if [ ! -x .venv/bin/streamlit ]; then
  echo "First run: installing Streamlit, Plotly, pandas, openpyxl…"
  python3 -m venv .venv && ./.venv/bin/pip install -q -r requirements.txt || exit 1
fi
exec ./.venv/bin/streamlit run app.py --browser.gatherUsageStats false
