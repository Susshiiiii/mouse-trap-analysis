# Mouse Trap Analysis

Streamlit dashboard that rebuilds the *Mouse Trap Charting* deck from the SEM/EDS raw-data workbook.
Upload the Excel file and every chart, pivot table and summary is generated in the browser.

## What it reads

Sheet **EDS data** (one row = one particle):

| Column | Field | Used for |
| --- | --- | --- |
| A | Process | grouping |
| B | Machine | grouping, filters |
| C | Location | location charts |
| J | Category by rule (New) | particle category (recalculated from element % if the cell has no saved value) |
| M | Particle Size (µm) | size bands: <6, 6–10, 10–20, 20–30, 30–50, 50–100, >100 |
| D, E, G, N–AG | Location code, particle count, visual, element wt% | Insight tab |

## Tabs

- **Overview** — category, size, category × size, location, and location × category charts by machine
- **By Process** — category and location charts for each process
- **Insight** — hotspot Pareto, defect mix heatmap, input/process/output zone, size distribution,
  large-particle rate, EDS fingerprint, metal source signature, visual vs EDS, heavy locations
- **Particle Log** — every particle with its key element %, downloadable as CSV

## Run locally

```bash
./run.command
```

or

```bash
python3 -m venv .venv && ./.venv/bin/pip install -r requirements.txt
./.venv/bin/streamlit run app.py
```

The uploaded workbook is processed in memory and is not saved. Raw data files are excluded from this repo by `.gitignore`.
