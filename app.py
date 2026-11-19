"""
Mouse Trap Analysis — Streamlit dashboard.

Upload the EDS raw-data workbook and every chart from the
"Mouse Trap Charting" deck is rebuilt from columns
A (Process), B (Machine), C (Location), J (Category by rule(New)), M (Particle Size).

Run:  streamlit run app.py
"""
from __future__ import annotations

import io
import itertools
import math
import re

import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from openpyxl import load_workbook

# ──────────────────────────────────────────────────────────────
# Constants
# ──────────────────────────────────────────────────────────────
CATEGORY_ORDER = ["Al Defect", "Ba Defect", "C,O,Si", "Fe Defect", "Organic", "Other Metal",
                  "SiN", "Si", "Other", "No particle"]
CAT_COLOR = {"Al Defect": "#2F7DD1", "Ba Defect": "#F28C28", "C,O,Si": "#1E8C6E", "Fe Defect": "#3EB8E0",
             "Organic": "#B04FC9", "Other Metal": "#34A853", "SiN": "#6A5ACD", "Si": "#8E8E93",
             "Other": "#A2845E", "No particle": "#C7C7CC"}
SIZE_BINS = [("<6um", 6), ("6-10um", 10), ("10-20um", 20), ("20-30um", 30),
             ("30-50um", 50), ("50-100um", 100), (">100um", float("inf"))]
SIZE_LABELS = [b[0] for b in SIZE_BINS]
SIZE_COLOR = dict(zip(SIZE_LABELS, ["#BFE1F6", "#8CC8EE", "#58ABE0", "#2F8BCB", "#1F6DAA", "#164F84", "#0E345C"]))
# Location code prefix (column D) → machine zone: A = input side, B = process area, C = output side
ZONES = {"A": "Input", "B": "Process", "C": "Output"}
ZONE_COLOR = {"Input": "#3EB8E0", "Process": "#B04FC9", "Output": "#F28C28", "Other": "#8E8E93"}
PALETTE = ["#2F7DD1", "#F28C28", "#34A853", "#B04FC9", "#3EB8E0", "#E5484D", "#6A5ACD",
           "#A2845E", "#1E8C6E", "#D6409F", "#8E8E93", "#C9A227", "#0E345C", "#7FB800"]
FONT = "-apple-system, BlinkMacSystemFont, 'SF Pro Text', 'Helvetica Neue', Arial, sans-serif"

# ──────────────────────────────────────────────────────────────
# Page + styling
# ──────────────────────────────────────────────────────────────
st.set_page_config(page_title="Mouse Trap Analysis", page_icon="📊", layout="wide")
st.markdown(f"""
<style>
html, body, [class*="css"], .stMarkdown, .stMetric {{ font-family:{FONT}; }}
.block-container {{ padding-top:2rem; max-width:1400px; }}
h1 {{ letter-spacing:-.025em; font-weight:700 !important; }}
h2, h3 {{ letter-spacing:-.015em; }}
.kpis {{ display:grid; grid-template-columns:repeat(5,minmax(0,1fr)); gap:12px; margin:8px 0 18px; }}
.kpi {{ background:rgba(128,128,128,.10); border:1px solid rgba(128,128,128,.14); border-radius:14px; padding:14px 16px; min-width:0; }}
.kpi .k {{ font-size:12px; opacity:.65; font-weight:500; }}
.kpi .v {{ font-size:28px; font-weight:700; letter-spacing:-.02em; font-variant-numeric:tabular-nums; }}
.kpi .s {{ font-size:12px; opacity:.75; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }}
@media (max-width:900px) {{ .kpis {{ grid-template-columns:repeat(2,minmax(0,1fr)); }} }}
.summary {{ background:rgba(128,128,128,.10); border-radius:10px; padding:10px 16px; margin:4px 0 8px; font-size:14.5px; }}
.summary ul {{ margin:0; padding-left:18px; }}
.summary .note {{ opacity:.6; font-size:12px; margin-top:4px; }}
.eyebrow {{ font-size:12px; opacity:.6; font-weight:500; margin-bottom:-12px; }}
.side-brand {{ display:flex; align-items:center; gap:10px; font-weight:650; font-size:16px; margin:-8px 0 6px; }}
.side-brand .logo {{ width:22px; height:22px; border-radius:6px; background:linear-gradient(135deg,#2F7DD1,#B04FC9); }}
.side-h {{ font-size:11px !important; font-weight:600; letter-spacing:.07em; text-transform:uppercase; opacity:.55;
          margin:22px 0 6px !important; padding-top:14px; border-top:1px solid rgba(128,128,128,.18); }}
section[data-testid="stSidebar"] [data-testid="stCaptionContainer"] {{ font-size:12px; }}
.finding {{ border-left:3px solid #409CFF; background:rgba(128,128,128,.08); border-radius:0 10px 10px 0; padding:10px 14px; margin:6px 0; font-size:14.5px; }}
.finding b.t {{ display:block; font-size:12px; text-transform:uppercase; letter-spacing:.06em; opacity:.6; margin-bottom:2px; }}
</style>
""", unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────
# Parsing
# ──────────────────────────────────────────────────────────────
def norm(v) -> str:
    return re.sub(r"\s+", " ", str(v)).strip() if v is not None else ""


def canon_category(c: str) -> str:
    t = norm(c)
    for x in CATEGORY_ORDER:
        if x.lower() == t.lower():
            return x
    return t


def canon_process(p: str) -> str:
    # "Wire bond" and "Wire Bond" are the same process
    return " ".join(w[:1].upper() + w[1:] for w in norm(p).split(" "))


def category_by_rule(e: dict) -> str:
    """Python version of the Excel formula in column J (Category by rule(New))."""
    g = lambda k: float(e.get(k)) if isinstance(e.get(k), (int, float)) else 0.0
    if g("Ba") > 0: return "Ba Defect"
    if g("Al") > 0: return "Al Defect"
    if g("Fe") > 0: return "Fe Defect"
    if g("Si") >= 30 and g("N") >= 20: return "SiN"
    if g("Si") > 50: return "Si"
    if g("Cu") >= 50 or g("Zn") >= 50: return "Other Metal"
    if g("Cu") >= 20 or g("Zn") >= 20: return "Other Metal"
    if g("Si") >= 20 and g("C") + g("O") >= 50: return "C,O,Si"
    if g("C") + g("O") >= 50: return "Organic"
    return "Other"


def size_bin(v: float) -> str:
    for label, hi in SIZE_BINS:
        if v < hi:
            return label
    return SIZE_LABELS[-1]


def key_element(cat: str, e: dict) -> str:
    g = lambda k: float(e.get(k) or 0) if isinstance(e.get(k), (int, float)) else 0.0
    simple = {"Al Defect": "Al", "Fe Defect": "Fe", "Ba Defect": "Ba", "C,O,Si": "Si", "Si": "Si", "SiN": "Si"}
    if cat in simple:
        k = simple[cat]; return f"{k} {g(k):.1f}%"
    if cat == "Organic":
        return f"C+O {g('C') + g('O'):.1f}%"
    if cat == "Other Metal":
        k = max(["Cu", "Ni", "Zn", "Sn", "Cr"], key=g); return f"{k} {g(k):.1f}%"
    return ""


@st.cache_data(show_spinner="Reading workbook…")
def parse_workbook(data: bytes) -> tuple[pd.DataFrame, str]:
    wb = load_workbook(io.BytesIO(data), data_only=True, read_only=True)
    ws_name = next((n for n in wb.sheetnames if n.strip().lower() == "eds data"), None)
    if ws_name is None:
        for n in wb.sheetnames:
            hdr = [norm(c).lower() for c in next(wb[n].iter_rows(min_row=1, max_row=1, values_only=True), [])]
            if "process" in hdr and "machine" in hdr:
                ws_name = n; break
    if ws_name is None:
        raise ValueError("Couldn’t find the “EDS data” sheet (a sheet with Process, Machine and Location headers).")

    rows = list(wb[ws_name].iter_rows(values_only=True))
    raw_hdr = [norm(h) for h in rows[0]]
    hdr = [h.lower() for h in raw_hdr]

    def col(names, fallback):
        for n in names:
            if n in hdr:
                return hdr.index(n)
        return fallback

    iP, iM, iL = col(["process"], 0), col(["machine"], 1), col(["location"], 2)
    iJ = col(["category by rule(new)", "category by rule (new)"], 9)
    iS = col(["particle size (um)", "particle size (µm)", "particle size"], 12)
    iN = col(["particle#", "particle #"], 5)
    iD = col(["location code"], 3)                       # extra columns used by the Insight tab
    iE = col(["particle count (ea>6um only)"], 4)
    iG = col(["particle visual"], 6)
    elem = {h: i for i, h in enumerate(raw_hdr) if i > iS and re.fullmatch(r"[A-Z][a-z]?", h)}

    out, P, M, L, pn, LC, LN = [], None, None, None, 0, "", ""
    for r, row in enumerate(rows[1:], start=2):
        if row is None:
            continue
        row = list(row) + [None] * (len(raw_hdr) - len(row))
        p, m, l = norm(row[iP]), norm(row[iM]), norm(row[iL])
        if p: P = p
        if m: M, L = m, None          # new machine → reset location
        if l: L, pn, LC, LN = l, 0, norm(row[iD]), norm(row[iE])
        try:
            size = float(row[iS])
        except (TypeError, ValueError):
            continue                   # no particle measured (e.g. "NA")
        if not P or not M or P.lower() == "na" or M.lower() == "na":
            continue
        e = {k: row[i] for k, i in elem.items()}
        cat = canon_category(row[iJ]) if row[iJ] is not None else ""
        derived = False
        if not cat or cat.startswith("#"):
            cat, derived = category_by_rule(e), True
        try:
            pn = int(row[iN])
        except (TypeError, ValueError):
            pn += 1
        out.append(dict(Process=canon_process(P), Machine=M, Location=L or "(unspecified)", Category=cat,
                        Size=size, Band=size_bin(size), Particle=pn, KeyElement=key_element(cat, e),
                        Recalculated=derived, ExcelRow=r,
                        LocCode=LC, Zone=ZONES.get(LC[:1].upper(), "Other"), LocCount=LN,
                        Visual=norm(row[iG]) or "Not recorded",
                        **{f"el_{k}": (float(v) if isinstance(v, (int, float)) else 0.0) for k, v in e.items()}))
    if not out:
        raise ValueError("No particle rows found. Check columns A, B, C, J and M on the EDS data sheet.")
    return pd.DataFrame(out), ws_name


# ──────────────────────────────────────────────────────────────
# Ordering helpers
# ──────────────────────────────────────────────────────────────
def proc_order(df):  # largest process first, as in the deck
    return df["Process"].value_counts().index.tolist() if len(df) else []


def mach_order(df):  # machines in workbook order, grouped by process
    out = []
    for p in proc_order(df):
        out += df.loc[df.Process == p, "Machine"].drop_duplicates().tolist()
    return out


def cat_order(df):
    present = set(df["Category"])
    return [c for c in CATEGORY_ORDER if c in present] + sorted(present - set(CATEGORY_ORDER))


def pct(a, b):
    return round(a / b * 100) if b else 0


def join_and(items):
    items = list(items)
    return items[0] if len(items) == 1 else ", ".join(items[:-1]) + " & " + items[-1]


# ──────────────────────────────────────────────────────────────
# Chart helpers
# ──────────────────────────────────────────────────────────────
def text_color(hex_):
    h = hex_.lstrip("#"); r, g, b = (int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))
    lin = lambda v: v / 12.92 if v <= .03928 else ((v + .055) / 1.055) ** 2.4
    return "#1D1D1F" if .2126 * lin(r) + .7152 * lin(g) + .0722 * lin(b) > .4 else "#FFFFFF"


def style(fig, *, height=460, x_title=None, y_title="Particle count", stacked=False, n_x=0, legend_right=False):
    fig.update_layout(
        barmode="stack" if stacked else "group", height=height, bargap=.35 if stacked else .2,
        font=dict(family=FONT, size=12), margin=dict(l=10, r=10, t=40, b=10),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0, title=None),
        hoverlabel=dict(font_family=FONT), plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)",
        uniformtext=dict(minsize=9, mode="hide"),
    )
    if legend_right:  # half-width charts: legend under the chart so bars keep their width
        fig.update_layout(legend=dict(orientation="h", x=0, y=-0.28, yanchor="top", xanchor="left", traceorder="normal"),
                          margin=dict(t=20, b=10))
    fig.update_yaxes(title=y_title, gridcolor="rgba(128,128,128,.18)", zeroline=False, rangemode="tozero")
    fig.update_xaxes(title=x_title, showgrid=False, tickangle=-60 if n_x > 8 else (-35 if n_x > 3 else 0))
    return fig


def x_axis(df_keys: pd.DataFrame, group_col: str | None, label_col: str):
    """Plotly multicategory axis (Excel-style grouped labels) when a group column is given."""
    if group_col:
        return [df_keys[group_col].tolist(), df_keys[label_col].tolist()]
    return df_keys[label_col].tolist()


def stacked_bars(keys, series, counts, color_of, group_col, label_col, *, show_totals=True):
    """keys: DataFrame of x positions; counts: dict[(key_tuple, series)] -> n."""
    fig = go.Figure()
    x = x_axis(keys, group_col, label_col)
    key_cols = [c for c in (group_col, label_col) if c]
    tuples = [tuple(r) for r in keys[key_cols].itertuples(index=False)]
    for s in series:
        ys = [counts.get((t, s), 0) for t in tuples]
        c = color_of(s)
        fig.add_bar(x=x, y=ys, name=s, marker_color=c, marker_line_width=0,
                    text=[v if v else "" for v in ys], textposition="inside", insidetextanchor="middle",
                    textfont=dict(color=text_color(c)), hovertemplate=f"<b>{s}</b><br>%{{x}}<br>%{{y}} particles<extra></extra>")
    if show_totals:
        totals = [sum(counts.get((t, s), 0) for s in series) for t in tuples]
        fig.add_scatter(x=x, y=totals, mode="text", text=totals, textposition="top center",
                        textfont=dict(size=12), showlegend=False, hoverinfo="skip")
    return fig


def hbars(keys, series, counts, color_of):
    """Horizontal stacked bars with a two-level label axis (group, label). Stays readable with many rows."""
    y = [[k[0] for k in keys], [k[1] for k in keys]]
    fig = go.Figure()
    for s_ in series:
        xs = [counts.get((k, s_), 0) for k in keys]
        if not any(xs):
            continue
        c = color_of(s_)
        fig.add_bar(y=y, x=xs, orientation="h", name=s_, marker_color=c, marker_line_width=0,
                    text=[v or "" for v in xs], textposition="inside", insidetextanchor="middle",
                    textfont=dict(color=text_color(c)),
                    hovertemplate=f"<b>{s_}</b><br>%{{y}}<br>%{{x}} particles<extra></extra>")
    totals = [sum(counts.get((k, s_), 0) for s_ in series) for k in keys]
    fig.add_scatter(y=y, x=totals, mode="text", text=[f"  {t}" for t in totals], textposition="middle right",
                    showlegend=False, hoverinfo="skip", textfont=dict(size=12))
    fig.update_layout(barmode="stack", height=max(320, 24 * len(keys) + 150), bargap=.28,
                      font=dict(family=FONT, size=12), margin=dict(l=10, r=30, t=40, b=10),
                      legend=dict(orientation="h", yanchor="bottom", y=1.0, xanchor="left", x=0, title=None, traceorder="normal"),
                      plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
    fig.update_yaxes(autorange="reversed", showgrid=False, tickfont=dict(size=12))
    fig.update_xaxes(title="Particle count", gridcolor="rgba(128,128,128,.18)", zeroline=False,
                     range=[0, max(totals) * 1.12 + .5] if totals else None, dtick=1 if max(totals or [0]) <= 12 else None)
    return fig


def summary(lines, note=None):
    body = "".join(f"<li>{l}</li>" for l in lines)
    st.markdown(f'<div class="summary"><ul>{body}</ul>{f"<div class=note>{note}</div>" if note else ""}</div>',
                unsafe_allow_html=True)


_card_ids = itertools.count()


def card(where=None):
    """Bordered surface for one chart section; the st-key-card_* class is styled as a glass card in CSS."""
    return (where or st).container(border=True, key=f"card_{next(_card_ids)}")


def section_title(eyebrow, title):
    st.markdown(f'<div class="eyebrow">{eyebrow}</div>', unsafe_allow_html=True)
    st.subheader(title)


def pivot(df, index, columns, col_order=None):
    t = pd.crosstab(index=[df[i] for i in index], columns=df[columns], margins=True, margins_name="Grand Total")
    if col_order:
        t = t[[c for c in col_order if c in t.columns] + ["Grand Total"]]
    return t


def show_table(t, label="Data table"):
    with st.expander(label):
        st.dataframe(t.replace(0, "–") if hasattr(t, "replace") else t, use_container_width=True)


PLOT_CFG = {"displaylogo": False, "toImageButtonOptions": {"format": "png", "scale": 2}}

# ──────────────────────────────────────────────────────────────
# Sidebar: upload + filters
# ──────────────────────────────────────────────────────────────
def _apply_theme():
    st._config.set_option("theme.base", "dark" if st.session_state.get("dark", True) else "light")


# keep the server theme in step with this session's toggle (dark by default)
if st._config.get_option("theme.base") != ("dark" if st.session_state.get("dark", True) else "light"):
    _apply_theme()
    st.rerun()


# ──────────────────────────────────────────────────────────────
# Glass theme: content sits on frosted cards above a soft colour backdrop;
# the sidebar and tab bar (controls) use a lighter, more translucent glass.
# ──────────────────────────────────────────────────────────────
DARK = st.session_state.get("dark", True)
T = dict(
    bg="#05070C", g1="rgba(64,156,255,.20)", g2="rgba(176,79,201,.16)", g3="rgba(52,168,83,.08)",
    card="rgba(26,28,36,.72)", card_b="rgba(255,255,255,.09)", hi="rgba(255,255,255,.06)",
    shadow="0 12px 32px rgba(0,0,0,.40)", side="rgba(16,18,24,.55)", side_b="rgba(255,255,255,.07)",
    bar="rgba(30,32,40,.62)", soft="rgba(255,255,255,.05)", soft_b="rgba(255,255,255,.08)",
    pill="rgba(255,255,255,.06)", pill_b="rgba(255,255,255,.12)", pill_on="rgba(64,156,255,.22)",
    pill_on_b="rgba(64,156,255,.65)", pill_on_t="#CFE4FF", solid="#16181F",
) if DARK else dict(
    bg="#EEF1F6", g1="rgba(64,156,255,.18)", g2="rgba(176,79,201,.12)", g3="rgba(52,168,83,.07)",
    card="rgba(255,255,255,.80)", card_b="rgba(15,23,42,.08)", hi="rgba(255,255,255,.9)",
    shadow="0 10px 28px rgba(22,34,66,.10)", side="rgba(255,255,255,.55)", side_b="rgba(15,23,42,.08)",
    bar="rgba(255,255,255,.70)", soft="rgba(15,23,42,.035)", soft_b="rgba(15,23,42,.08)",
    pill="rgba(255,255,255,.75)", pill_b="rgba(15,23,42,.12)", pill_on="rgba(0,102,204,.12)",
    pill_on_b="rgba(0,102,204,.55)", pill_on_t="#0057B8", solid="#FFFFFF",
)
st.markdown(f"""
<style>
[data-testid="stApp"] {{
  background:
    radial-gradient(1100px 620px at 8% -10%, {T['g1']}, transparent 60%),
    radial-gradient(900px 560px at 100% 0%, {T['g2']}, transparent 60%),
    radial-gradient(900px 600px at 50% 110%, {T['g3']}, transparent 60%),
    {T['bg']};
  background-attachment: fixed;
}}
header[data-testid="stHeader"] {{ background:transparent; }}

/* sidebar: functional layer → translucent glass */
section[data-testid="stSidebar"] {{
  background:{T['side']} !important; border-right:1px solid {T['side_b']};
  backdrop-filter:blur(28px) saturate(160%); -webkit-backdrop-filter:blur(28px) saturate(160%);
}}
section[data-testid="stSidebar"] > div {{ background:transparent; }}

/* cards: content layer → frosted but mostly opaque, for legible charts */
[class*="st-key-card_"], .kpi {{
  background:{T['card']} !important; border:1px solid {T['card_b']} !important; border-radius:20px !important;
  box-shadow:inset 0 1px 0 {T['hi']}, {T['shadow']};
  backdrop-filter:blur(22px) saturate(150%); -webkit-backdrop-filter:blur(22px) saturate(150%);
}}
[class*="st-key-card_"] {{ padding:22px 24px 18px !important; margin-bottom:10px; }}
[class*="st-key-card_"] h3 {{ padding-top:0; }}

/* tab bar: floating glass segmented control */
.stTabs [data-baseweb="tab-list"] {{
  position:sticky; top:3.4rem; z-index:50; gap:4px; width:fit-content; padding:5px; margin-bottom:14px;
  background:{T['bar']}; border:1px solid {T['card_b']}; border-radius:14px;
  backdrop-filter:blur(24px) saturate(160%); -webkit-backdrop-filter:blur(24px) saturate(160%);
  box-shadow:inset 0 1px 0 {T['hi']}, 0 6px 18px rgba(0,0,0,.18);
}}
.stTabs [data-baseweb="tab"] {{ height:34px; padding:0 16px; border-radius:10px; }}
.stTabs [data-baseweb="tab"][aria-selected="true"] {{ background:{T['pill_on']}; }}
.stTabs [data-baseweb="tab-highlight"], .stTabs [data-baseweb="tab-border"] {{ display:none; }}

/* inner surfaces */
.summary, .finding {{ background:{T['soft']} !important; border:1px solid {T['soft_b']}; }}
.finding {{ border-left:3px solid #409CFF; }}
[data-testid="stExpander"] details {{ border:1px solid {T['soft_b']} !important; border-radius:12px; background:{T['soft']}; }}
[data-testid="stDataFrame"] {{ border-radius:12px; overflow:hidden; }}

/* machine filter box */
.st-key-machbox {{
  background:{T['soft']} !important; border:1px solid {T['soft_b']} !important; border-radius:14px !important;
  padding:12px 12px 10px !important; gap:.45rem !important;
}}
.mach-head {{ display:flex; justify-content:space-between; align-items:baseline; font-size:14px; margin:6px 2px 6px; }}
.mach-head .count {{ font-size:12px; opacity:.6; font-variant-numeric:tabular-nums; }}
.mach-group {{ font-size:11px; font-weight:600; letter-spacing:.06em; text-transform:uppercase; opacity:.55; margin:8px 0 0; padding-bottom:8px; line-height:1; }}
.st-key-machbox [data-testid="stBaseButton-pills"], .st-key-machbox [data-testid="stBaseButton-pillsActive"] {{
  border-radius:9px !important; min-height:28px; padding:1px 8px !important; font-variant-numeric:tabular-nums;
  background:{T['pill']} !important; border:1px solid {T['pill_b']} !important;
}}
.st-key-machbox [data-testid="stBaseButton-pillsActive"] {{
  background:{T['pill_on']} !important; border-color:{T['pill_on_b']} !important; color:{T['pill_on_t']} !important;
}}
.st-key-machbox [data-testid="stBaseButton-pillsActive"] p {{ color:{T['pill_on_t']} !important; font-weight:600; }}
.st-key-machbox [data-testid^="stBaseButton-pills"] p {{ font-size:12.5px; }}
.st-key-machbox button[kind="secondary"] {{ border-radius:10px; min-height:34px; }}

/* people who ask the system for less transparency get solid surfaces */
@media (prefers-reduced-transparency: reduce) {{
  [class*="st-key-card_"], .kpi, section[data-testid="stSidebar"], .stTabs [data-baseweb="tab-list"] {{
    background:{T['solid']} !important; backdrop-filter:none !important; -webkit-backdrop-filter:none !important;
  }}
}}
</style>
""", unsafe_allow_html=True)


with st.sidebar:
    st.markdown('<div class="side-brand"><span class="logo"></span>Mouse Trap Analysis</div>', unsafe_allow_html=True)
    st.markdown('<p class="side-h">Data source</p>', unsafe_allow_html=True)
    up = st.file_uploader("EDS Excel workbook", type=["xlsx", "xlsm"], label_visibility="collapsed")
    st.caption("Sheet **EDS data** · columns **A** Process · **B** Machine · **C** Location · "
               "**J** Category · **M** Size (µm). One row = one particle.")
    filter_box = st.container()          # filled once a file is loaded
    st.markdown('<p class="side-h">Appearance</p>', unsafe_allow_html=True)
    st.toggle("Dark mode", value=True, key="dark", on_change=_apply_theme)

if up is None:
    st.title("Mouse Trap Analysis")
    st.markdown("#### Upload the EDS Excel workbook in the sidebar to rebuild every chart from the deck.")
    st.info("The file stays on this computer — nothing is sent anywhere. Merged cells in A, B and C are filled "
            "down automatically. If column J has no saved values, categories are recalculated from the element "
            "columns using the same rule as the Excel formula.", icon="📄")
    st.stop()

try:
    ALL, sheet = parse_workbook(up.getvalue())
except Exception as exc:  # noqa: BLE001
    st.error(f"This file couldn’t be read. {exc}")
    st.stop()

def _set_machines(groups, select_all):
    for key, values in groups:
        st.session_state[key] = list(values) if select_all else []


with filter_box:
    st.markdown('<p class="side-h">Filters</p>', unsafe_allow_html=True)
    proc_opts = ["All processes"] + proc_order(ALL)
    proc_label = st.selectbox("Process", proc_opts, format_func=lambda p: p if p == "All processes"
                              else f"{p}  ·  {int((ALL.Process == p).sum())} particles")
    proc = "All" if proc_label == "All processes" else proc_label
    pool = ALL if proc == "All" else ALL[ALL.Process == proc]
    groups = [(f"mach::{proc}::{p}", mach_order(pool[pool.Process == p])) for p in proc_order(pool)]
    for key, ms in groups:
        st.session_state.setdefault(key, ms)
    n_avail = sum(len(ms) for _, ms in groups)
    n_sel = sum(len(st.session_state.get(k) or []) for k, _ in groups)
    st.markdown(f'<div class="mach-head"><span>Machines</span><span class="count">{n_sel} / {n_avail}</span></div>',
                unsafe_allow_html=True)
    sel = []
    with st.container(border=True, key="machbox"):
        for key, ms in groups:
            pname = key.split("::")[-1]
            if len(groups) > 1:
                st.markdown(f'<div class="mach-group">{pname}</div>', unsafe_allow_html=True)
            sel += st.pills(pname, ms, selection_mode="multi", key=key, label_visibility="collapsed") or []
        c1, c2 = st.columns(2)
        c1.button("Select all", on_click=_set_machines, args=(groups, True), use_container_width=True)
        c2.button("Clear", on_click=_set_machines, args=(groups, False), use_container_width=True)
    st.caption(f"{int(pool.Machine.isin(sel).sum())} particles in view · the deck’s first section uses WB-KNS-611 + 612")

df = pool[pool.Machine.isin(sel)]
MACH_COLOR = {m: PALETTE[i % len(PALETTE)] for i, m in enumerate(mach_order(ALL))}

st.title("Particle Overview" if proc == "All" else proc)
n_recalc = int(ALL.Recalculated.sum())
st.caption(f"{len(ALL)} particles from “{sheet}” in **{up.name}**"
           + (f" · {n_recalc} categories recalculated from element data" if n_recalc else ""))

if df.empty:
    st.warning("No machines selected. Choose at least one machine in the sidebar.")
    st.stop()

N = len(df)
procs = proc_order(df)
machines = mach_order(df)
cats = cat_order(df)

# ── KPIs ──────────────────────────────────────────────────────
top_cat = df.Category.value_counts()
top_m = df.Machine.value_counts()
g20 = int((df.Size >= 20).sum())
kpis = [("Particles", N, f"{len(procs)} process{'es' if len(procs) > 1 else ''}"),
        ("Machines", len(machines), f"Most: {top_m.index[0]} ({top_m.iloc[0]})"),
        ("Locations", df[["Process", "Location"]].drop_duplicates().shape[0], "machine areas sampled"),
        ("Top Category", f"{pct(top_cat.iloc[0], N)}%", top_cat.index[0]),
        ("Size ≥ 20 µm", f"{pct(g20, N)}%", f"{g20} of {N} particles")]
st.markdown('<div class="kpis">' + "".join(
    f'<div class="kpi"><div class="k">{a}</div><div class="v">{b}</div><div class="s">{c}</div></div>' for a, b, c in kpis)
    + "</div>", unsafe_allow_html=True)

tabs = st.tabs(["Overview", "By Process", "Insight", "Particle Log"])
mkeys = pd.DataFrame([(p, m) for p in procs for m in mach_order(df[df.Process == p])], columns=["Process", "Machine"])
grp = "Process" if len(procs) > 1 else None

# ══════════════════════════════════════════════════════════════
# Overview tab
# ══════════════════════════════════════════════════════════════
with tabs[0]:
    with card():
        # 1 ─ Category by machine
        section_title("Column J × Column B", "Particle Count & Category by Machine")
        mc = top_m.sort_values(ascending=False)
        tied = mc[mc == mc.iloc[0]].index.tolist()
        lines = []
        if len(machines) > 1:
            lines.append("All machines have the same particle count." if len(tied) == len(machines) else
                         f"Highest particle count observed at {join_and([f'<b>{m}</b>' for m in tied])} ({mc.iloc[0]} particles{' each' if len(tied) > 1 else ''}).")
        lines.append("Highest category: " + join_and([f"<b>{c}</b> {v} ({pct(v, N)}%)" for c, v in top_cat.head(3).items()]) + ".")
        summary(lines, "Note: min. particle count is 1 and highest count is 5 per location.")
        cnt = df.groupby(["Process", "Machine", "Category"]).size()
        counts = {((p, m) if grp else (m,), c): v for (p, m, c), v in cnt.items()}
        fig = stacked_bars(mkeys, cats, counts, lambda c: CAT_COLOR.get(c, "#999"), grp, "Machine")
        st.plotly_chart(style(fig, stacked=True, x_title=None if grp else "Machine", n_x=len(mkeys)), use_container_width=True, config=PLOT_CFG)
        show_table(pivot(df, ["Process", "Machine"], "Category", cats))

    with card():
        # 2 ─ Size by machine
        section_title("Column M × Column B", "Particle Size Category by Process & Machine")
        bins = [b for b in SIZE_LABELS if b in set(df.Band)]
        bc = df.Band.value_counts()
        summary([f"<b>{pct((df.Size >= 20).sum(), N)}%</b> of particles are ≥ 20 µm; <b>{pct((df.Size >= 30).sum(), N)}%</b> are ≥ 30 µm.",
                 f"Most common size band: <b>{bc.index[0]}</b> ({bc.iloc[0]}). Median {df.Size.median():.1f} µm, largest {df.Size.max():.1f} µm."],
                "Size bands: <6, 6–10, 10–20, 20–30, 30–50, 50–100, >100 µm (lower bound inclusive).")
        cnt = df.groupby(["Process", "Machine", "Band"]).size()
        counts = {((p, m) if grp else (m,), b): v for (p, m, b), v in cnt.items()}
        fig = stacked_bars(mkeys, bins, counts, SIZE_COLOR.get, grp, "Machine")
        st.plotly_chart(style(fig, stacked=True, x_title=None if grp else "Machine", n_x=len(mkeys)), use_container_width=True, config=PLOT_CFG)
        show_table(pivot(df, ["Process", "Machine"], "Band", SIZE_LABELS))

    with card():
        # 3 ─ Category × Size × Machine (clustered)
        section_title("Column J × Column M × Column B", "Particle Count by Category – Size – Machine")
        cs = df.groupby(["Category", "Band"]).size()
        combos = pd.DataFrame([(c, b) for c in cats for b in SIZE_LABELS if (c, b) in cs.index], columns=["Category", "Band"])
        top = cs.idxmax()
        summary([f"Largest cell: <b>{top[0]} · {top[1]}</b> with {cs.max()} particles."],
                "Showing stacked bars because more than 4 machines are selected. Pick up to 4 machines for the deck’s side-by-side layout." if len(machines) > 4 else None)
        cnt = df.groupby(["Category", "Band", "Machine"]).size()
        combo_t = list(combos.itertuples(index=False, name=None))
        if len(machines) <= 4:   # deck style: clustered columns, one colour per machine
            fig = go.Figure()
            for m in machines:
                ys = [cnt.get((c, b, m), 0) for c, b in combo_t]
                fig.add_bar(x=[combos.Category.tolist(), combos.Band.tolist()], y=ys, name=m, marker_color=MACH_COLOR[m],
                            text=[v if v else "" for v in ys], textposition="outside")
            fig = style(fig, n_x=len(combos))
        else:                    # many machines: horizontal stacked bars stay readable
            fig = hbars(combo_t, machines, {((c, b), m): v for (c, b, m), v in cnt.items()}, MACH_COLOR.get)
        st.plotly_chart(fig, use_container_width=True, config=PLOT_CFG)
        show_table(pivot(df, ["Category", "Band"], "Machine", machines))

    with card():
        # 4 ─ Count by machine location
        section_title("Column C × Column B", "Particle Count by Machine Location")
        lines, lk = [], []
        for p in procs:
            vc = df[df.Process == p].Location.value_counts()
            lines.append(f"{p + ': ' if len(procs) > 1 else ''}highest particle count at <b>{vc.index[0]}</b> ({vc.iloc[0]} of {vc.sum()}).")
            lk += [(p, l) for l in vc.index]
        summary(lines)
        lkeys = pd.DataFrame(lk, columns=["Process", "Location"])
        cnt = df.groupby(["Process", "Location", "Machine"]).size()
        counts = {((p, l) if grp else (l,), m): v for (p, l, m), v in cnt.items()}
        fig = stacked_bars(lkeys, machines, counts, MACH_COLOR.get, grp, "Location")
        st.plotly_chart(style(fig, stacked=True, height=520, x_title=None if grp else "Machine location", n_x=len(lkeys)),
                        use_container_width=True, config=PLOT_CFG)
        show_table(pivot(df, ["Process", "Location"], "Machine", machines))

    with card():
        # 5 ─ Location × Category × Machine (clustered)
        section_title("Column C × Column J × Column B", "Particle Count by Category Defect & Machine Location")
        lc = df.groupby(["Location", "Category"]).size()
        (tl, tc), tv = lc.idxmax(), lc.max()
        summary([f"Highest particle count located at <b>{tl}</b> under <b>{tc}</b> category ({tv})."],
                None)
        cnt = df.groupby(["Process", "Location", "Category", "Machine"]).size()

        def loc_cat_chart(p):
            sub = df[df.Process == p]
            pm_ = mach_order(sub)
            keys = [(l, c) for l in sub.Location.value_counts().index
                    for c in sub[sub.Location == l].Category.value_counts().index]
            if len(pm_) <= 4 and len(keys) <= 24:   # deck style
                fig = go.Figure()
                for m in pm_:
                    ys = [cnt.get((p, l, c, m), 0) for l, c in keys]
                    fig.add_bar(x=[[k[0] for k in keys], [k[1] for k in keys]], y=ys, name=m, marker_color=MACH_COLOR[m],
                                text=[v if v else "" for v in ys], textposition="outside")
                return style(fig, n_x=len(keys), height=500)
            return hbars(keys, pm_, {((l, c), m): v for (pp, l, c, m), v in cnt.items() if pp == p}, MACH_COLOR.get)

        if len(procs) == 1:
            st.plotly_chart(loc_cat_chart(procs[0]), use_container_width=True, config=PLOT_CFG)
        else:
            st.caption("One chart per process — locations are listed with their defect categories.")
            for t, p in zip(st.tabs(procs), procs):
                with t:
                    st.plotly_chart(loc_cat_chart(p), use_container_width=True, config=PLOT_CFG, key=f"lc_{p}")
        show_table(pivot(df, ["Process", "Location", "Category"], "Machine", machines))

# ══════════════════════════════════════════════════════════════
# By Process tab (one block per process, two charts each)
# ══════════════════════════════════════════════════════════════
with tabs[1]:
    for i, p in enumerate(procs):
        with card():
            pr = df[df.Process == p]
            pm = mach_order(pr)
            section_title(f"{len(pr)} particles · {len(pm)} machine{'s' if len(pm) > 1 else ''}",
                          f"Particle Count by Machine Location – {p}")
            vc = pr.Location.value_counts()
            lead = vc[vc >= max(vc.iloc[0] * .7, 2)].index[:3].tolist() or [vc.index[0]]
            tc2 = pr.Category.value_counts().head(2)
            summary([f"Highest particle count observed at {join_and([f'<b>{l}</b>' for l in lead])}.",
                     "Highest particle category " + join_and([f"<b>{c}</b> ({pct(v, len(pr))}%)" for c, v in tc2.items()]) + "."])
            keys = pd.DataFrame({"Machine": pm})
            a, b = st.columns(2)
            with a:
                cnt = pr.groupby(["Machine", "Category"]).size()
                fig = stacked_bars(keys, cat_order(pr), {((m,), c): v for (m, c), v in cnt.items()},
                                   lambda c: CAT_COLOR.get(c, "#999"), None, "Machine")
                st.markdown(f"**Particle Category by Machine – {p}**")
                st.plotly_chart(style(fig, stacked=True, x_title=None, height=520, legend_right=True, n_x=len(pm)), use_container_width=True,
                                config=PLOT_CFG, key=f"pa{i}")
            with b:
                locs = vc.index[::-1].tolist()
                lcol = {l: PALETTE[j % len(PALETTE)] for j, l in enumerate(locs)}
                cnt = pr.groupby(["Machine", "Location"]).size()
                fig = stacked_bars(keys, locs, {((m,), l): v for (m, l), v in cnt.items()}, lcol.get, None, "Machine")
                st.markdown(f"**Particle Count by Machine – {p}**")
                st.plotly_chart(style(fig, stacked=True, x_title=None, height=520, legend_right=True, n_x=len(pm)), use_container_width=True,
                                config=PLOT_CFG, key=f"pb{i}")

# ══════════════════════════════════════════════════════════════
# Insight tab — extra analysis using columns D, E, G and the EDS element columns
# ══════════════════════════════════════════════════════════════
METAL_CATS = {"Al Defect", "Fe Defect", "Ba Defect", "Other Metal"}
METAL_ELEMS = ["Al", "Fe", "Cu", "Ni", "Zn", "Sn", "Cr", "Ba"]
ELEM_COLOR = {"Al": "#2F7DD1", "Fe": "#3EB8E0", "Cu": "#C9A227", "Ni": "#34A853", "Zn": "#6A5ACD",
              "Sn": "#8E8E93", "Cr": "#1E8C6E", "Ba": "#F28C28"}


def visual_group(v: str) -> str:
    v = v.lower().replace("greysih", "greyish").replace("color", "").strip()
    if v.startswith("not recorded"):
        return "Not recorded"
    colour = v.split()[0].capitalize() if v else "Other"
    return f"{colour} metallic" if "metallic" in v else colour


def finding(title, text):
    st.markdown(f'<div class="finding"><b class="t">{title}</b>{text}</div>', unsafe_allow_html=True)


def heat(z: pd.DataFrame, text, *, colorscale, height, zfmt, x_title=None, y_title=None, zmax=None):
    fig = go.Figure(go.Heatmap(z=z.values, x=z.columns.tolist(), y=z.index.tolist(), colorscale=colorscale,
                               text=text, texttemplate="%{text}", zmin=0, zmax=zmax, xgap=2, ygap=2,
                               hovertemplate=f"%{{y}} · %{{x}}<br>{zfmt}<extra></extra>",
                               colorbar=dict(thickness=10, outlinewidth=0)))
    fig.update_layout(height=height, font=dict(family=FONT, size=12), margin=dict(l=10, r=10, t=10, b=10),
                      plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
    fig.update_yaxes(autorange="reversed", title=y_title, showgrid=False)
    fig.update_xaxes(side="top", title=x_title, showgrid=False)
    return fig


with tabs[2]:
    section_title("Analyst view · uses Location code (D), Particle count (E), Particle visual (G) and EDS element %",
                  "Insight")

    # ---------- numbers behind the key findings ----------
    hot = df.groupby(["Process", "Location"]).size().sort_values(ascending=False)
    cum = hot.cumsum() / hot.sum() * 100
    n80 = int((cum < 80).sum()) + 1
    zone = df.Zone.value_counts()
    metal = df[df.Category.isin(METAL_CATS)].copy()
    metal_share = pct(len(metal), N)
    _sz = df.groupby("Category").Size.agg(["median", "size"])
    med = (_sz[_sz["size"] >= 3] if (_sz["size"] >= 3).sum() >= 2 else _sz)["median"].sort_values(ascending=False)
    big = df.assign(big=df.Size >= 50).groupby("Machine").big.agg(["mean", "sum", "size"]).sort_values("mean", ascending=False)
    loc_tbl = df.drop_duplicates(["Process", "Machine", "Location"])
    heavy = loc_tbl[loc_tbl.LocCount.astype(str).str.contains(">")]
    df_vis = df.assign(VisualGroup=df.Visual.map(visual_group))
    if len(metal):
        metal["Dominant"] = metal[[f"el_{e}" for e in METAL_ELEMS if f"el_{e}" in metal]].idxmax(axis=1).str[3:]

    # ---------- key findings ----------
    st.markdown("#### Key findings")
    f1, f2 = st.columns(2)
    with f1:
        finding("Concentration", f"<b>{n80} of {len(hot)}</b> machine locations ({pct(n80, len(hot))}%) hold <b>80%</b> of all particles. "
                f"Top hotspot: <b>{hot.index[0][1]}</b> ({hot.index[0][0]}, {hot.iloc[0]} particles).")
        finding("Machine zone", f"<b>{zone.index[0]}</b> zone carries the most particles ({pct(zone.iloc[0], N)}%). "
                + " · ".join(f"{z} {pct(v, N)}%" for z, v in zone.items()) + ".")
        finding("Particle size", f"<b>{med.index[0]}</b> particles are the largest (median {med.iloc[0]:.0f} µm); "
                f"<b>{med.index[-1]}</b> the smallest ({med.iloc[-1]:.0f} µm). Categories with fewer than 3 particles are left out.")
    with f2:
        if len(metal):
            dom = metal.Dominant.value_counts()
            finding("Metallic contamination", f"<b>{metal_share}%</b> of particles are metallic (Al / Fe / Ba / other metal). "
                    f"Dominant element: <b>{dom.index[0]}</b> ({pct(dom.iloc[0], len(metal))}% of metallic particles).")
        if len(big):
            finding("Large-particle risk", f"<b>{big.index[0]}</b> has the highest share of particles ≥ 50 µm "
                    f"({big['mean'].iloc[0] * 100:.0f}%, {int(big['sum'].iloc[0])} of {int(big['size'].iloc[0])}).")
        finding("Heavy locations", f"<b>{len(heavy)} of {len(loc_tbl)}</b> sampled locations showed <b>more than 5</b> particles "
                "at 400X (column E). At most 5 particles per location go to EDS, so these spots are under-counted in the charts.")

    with card():
        # ---------- 1. Pareto ----------
        st.markdown("#### 1 · Hotspot Pareto")
        st.caption("Where to clean first. Bars = particles per machine location (all selected machines); line = cumulative share.")
        # single-level labels keep the strict largest-first Pareto order across processes
        labels = [f"{l} · {p}" if len(procs) > 1 else l for p, l in hot.index]
        fig = go.Figure()
        fig.add_bar(x=labels, y=hot.values, name="Particles", marker_color=["#409CFF" if i < n80 else "rgba(128,128,128,.45)" for i in range(len(hot))],
                    text=hot.values, textposition="outside", hovertemplate="%{x}<br>%{y} particles<extra></extra>")
        fig.add_scatter(x=labels, y=cum.values, name="Cumulative %", yaxis="y2", mode="lines+markers",
                        line=dict(color="#F28C28", width=2), marker=dict(size=5), hovertemplate="%{y:.0f}%<extra></extra>")
        fig = style(fig, n_x=len(hot), height=560)
        fig.update_xaxes(tickfont=dict(size=11))
        fig.update_layout(yaxis2=dict(overlaying="y", side="right", range=[0, 105], ticksuffix="%", showgrid=False, title=None),
                          shapes=[dict(type="line", xref="paper", x0=0, x1=1, yref="y2", y0=80, y1=80,
                                       line=dict(color="#F28C28", width=1, dash="dot"))])
        st.plotly_chart(fig, use_container_width=True, config=PLOT_CFG)

    # ---------- 2. Heatmap + zone ----------
    a, b = st.columns([3, 2])
    with card(a):
        st.markdown("#### 2 · Defect mix per machine")
        st.caption("Share of each category within a machine (row = 100%). Number = particle count.")
        ct = pd.crosstab(df.Machine, df.Category).reindex(index=machines, columns=cats, fill_value=0)
        share = ct.div(ct.sum(axis=1), axis=0) * 100
        st.plotly_chart(heat(share.round(0), ct.where(ct > 0, "").values, colorscale="Blues", zfmt="%{z:.0f}% of machine",
                             height=60 + 34 * len(machines), zmax=100), use_container_width=True, config=PLOT_CFG)
    with card(b):
        st.markdown("#### 3 · Input vs process vs output zone")
        st.caption("Zone from Location code (column D): A = input side, B = process area, C = output side.")
        zt = pd.crosstab(df.Process, df.Zone).reindex(index=procs, fill_value=0)
        zs = zt.div(zt.sum(axis=1), axis=0) * 100
        fig = go.Figure()
        for z_ in [z for z in ["Input", "Process", "Output", "Other"] if z in zs.columns]:
            fig.add_bar(y=procs, x=zs[z_], name=z_, orientation="h", marker_color=ZONE_COLOR[z_],
                        text=[f"{v:.0f}%" if v >= 6 else "" for v in zs[z_]], textposition="inside",
                        customdata=zt[z_], hovertemplate=f"{z_}: %{{x:.0f}}% (%{{customdata}} particles)<extra></extra>")
        fig = style(fig, stacked=True, height=60 + 70 * len(procs), y_title=None)
        fig.update_xaxes(title=None, ticksuffix="%", range=[0, 100])
        fig.update_yaxes(autorange="reversed", showgrid=False)
        st.plotly_chart(fig, use_container_width=True, config=PLOT_CFG)

    # ---------- 4. size by category + large-particle rate ----------
    a, b = st.columns(2)
    with card(a):
        st.markdown("#### 4 · Size distribution by category")
        st.caption("Each dot is a particle; box = middle 50%. Log scale, because sizes range from a few µm to 800+ µm.")
        fig = go.Figure()
        for c in cats:
            v = df[df.Category == c]
            fig.add_box(y=v.Size, name=c, marker_color=CAT_COLOR.get(c, "#999"), boxpoints="all", jitter=.45, pointpos=0,
                        marker=dict(size=5, opacity=.75), line=dict(width=1.5), fillcolor="rgba(0,0,0,0)",
                        customdata=v[["Machine", "Location"]].values,
                        hovertemplate="%{y:.1f} µm<br>%{customdata[0]} · %{customdata[1]}<extra></extra>")
        fig = style(fig, height=440, y_title="Particle size (µm, log)")
        fig.update_yaxes(type="log", rangemode="normal", autorange=True)
        fig.update_layout(showlegend=False)
        for t in (20, 50, 100):   # guide lines; annotations on a log axis take log10 positions
            fig.add_hline(y=t, line=dict(color="rgba(128,128,128,.45)", width=1, dash="dot"))
            fig.add_annotation(x=1, xref="paper", y=math.log10(t), yref="y", text=f"{t} µm", showarrow=False,
                               xanchor="left", font=dict(size=10), opacity=.7)
        st.plotly_chart(fig, use_container_width=True, config=PLOT_CFG)
    with card(b):
        st.markdown("#### 5 · Large-particle rate by machine")
        st.caption("Share of particles ≥ 50 µm. Large particles carry the highest risk of a visible defect.")
        bb = big.sort_values("mean")
        fig = go.Figure(go.Bar(y=bb.index, x=bb["mean"] * 100, orientation="h",
                               marker_color=[MACH_COLOR[m] for m in bb.index],
                               text=[f"{v * 100:.0f}%  ({int(s_)}/{int(n_)})" for v, s_, n_ in zip(bb["mean"], bb["sum"], bb["size"])],
                               textposition="outside", hovertemplate="%{y}: %{x:.0f}%<extra></extra>"))
        fig = style(fig, height=440, y_title=None)
        fig.update_xaxes(range=[0, 115], ticksuffix="%", title=None)
        fig.update_yaxes(showgrid=False)
        st.plotly_chart(fig, use_container_width=True, config=PLOT_CFG)

    # ---------- 6. EDS fingerprint + dominant metal ----------
    a, b = st.columns(2)
    el_cols = [c for c in df.columns if c.startswith("el_")]
    with card(a):
        st.markdown("#### 6 · EDS chemical fingerprint")
        st.caption("Average element weight % for each category (elements below 1% everywhere are hidden).")
        fp = df.groupby("Category")[el_cols].mean().reindex(cats)
        fp.columns = [c[3:] for c in fp.columns]
        fp = fp.loc[:, fp.max() >= 1]
        st.plotly_chart(heat(fp.round(1), fp.round(1).astype(str).replace("0.0", "").values, colorscale="Viridis",
                             zfmt="%{z:.1f} wt%", height=80 + 40 * len(cats)), use_container_width=True, config=PLOT_CFG)
    with card(b):
        st.markdown("#### 7 · Metal source signature")
        st.caption("Metallic particles only, coloured by their highest metal element. Points to the part that is wearing.")
        if len(metal):
            mm = [m for m in machines if m in set(metal.Machine)]
            dc = metal.groupby(["Machine", "Dominant"]).size()
            fig = stacked_bars(pd.DataFrame({"Machine": mm}), [e for e in METAL_ELEMS if e in set(metal.Dominant)],
                               {((m,), e): v for (m, e), v in dc.items()}, lambda e: ELEM_COLOR.get(e, "#999"), None, "Machine")
            st.plotly_chart(style(fig, stacked=True, height=80 + 40 * len(cats) if len(cats) > 6 else 380, n_x=len(mm)),
                            use_container_width=True, config=PLOT_CFG)
        else:
            st.info("No metallic particles in the current selection.")

    # ---------- 8. visual vs EDS ----------
    a, b = st.columns([3, 2])
    with card(a):
        st.markdown("#### 8 · Can visual inspection predict the category?")
        st.caption("Particle visual (column G) against the EDS category. Row = 100%.")
        vt = pd.crosstab(df_vis.VisualGroup, df_vis.Category).reindex(columns=cats, fill_value=0)
        vt = vt.loc[vt.sum(axis=1).sort_values(ascending=False).index]
        vs = vt.div(vt.sum(axis=1), axis=0) * 100
        st.plotly_chart(heat(vs.round(0), vt.where(vt > 0, "").values, colorscale="Purples", zfmt="%{z:.0f}% of visual group",
                             height=60 + 34 * len(vt), zmax=100), use_container_width=True, config=PLOT_CFG)
        best = vs.max(axis=1)
        ok = best[(best >= 75) & (vt.sum(axis=1) >= 3)]
        finding("Reading", "Visual groups (3+ particles) where one category explains ≥ 75%: "
                + (", ".join(f"<b>{v}</b> → {vs.loc[v].idxmax()} ({best[v]:.0f}%, n={int(vt.loc[v].sum())})" for v in ok.index[:4]) or "none")
                + ". For the rest, colour alone cannot tell the category — EDS is still needed.")
    with card(b):
        st.markdown("#### 9 · Locations with more than 5 particles")
        st.caption("Column E at 400X. These spots had more particles than were analysed.")
        if len(heavy):
            st.dataframe(heavy[["Process", "Machine", "Location", "LocCode"]].rename(columns={"LocCode": "Code"}),
                         hide_index=True, use_container_width=True, height=min(420, 38 + 35 * len(heavy)))
        else:
            st.success("No location exceeded 5 particles in this selection.")

    with card():
        st.markdown("#### Recommended actions")
        acts = [f"Prioritise cleaning / PM at the top {n80} hotspot locations highlighted in blue on the Pareto — they cover 80% of particles."]
        if len(metal):
            top_m_el = metal.groupby("Machine").Dominant.agg(lambda x: x.value_counts().index[0])
            worst = metal.Machine.value_counts().index[0]
            acts.append(f"Check wear of <b>{top_m_el[worst]}</b>-bearing parts on <b>{worst}</b> (most metallic particles); "
                        "Ni or Cu signatures usually point to plated or brass components, Al to anodised fixtures or carriers.")
        acts.append(f"Focus on the <b>{zone.index[0].lower()}</b> zone, which carries {pct(zone.iloc[0], N)}% of particles.")
        if len(heavy):
            acts.append(f"Re-sample the {len(heavy)} heavy (> 5) locations with a larger EDS sample to confirm the category mix.")
        if "Organic" in top_cat.index and top_cat.index[0] == "Organic":
            acts.append(f"Organic is the largest category ({pct(top_cat['Organic'], N)}%) — review gloves, wipes, tapes and "
                        "lubricant handling around these machines.")
        summary(acts)

# ══════════════════════════════════════════════════════════════
# Particle log (replaces the SEM image slides)
# ══════════════════════════════════════════════════════════════
with tabs[3]:
    with card():
        section_title("Every particle, as listed on the SEM image slides", "Particle Log")
        q = st.text_input("Filter", placeholder="Machine, location or category", label_visibility="collapsed")
        log = df[["Process", "Machine", "Location", "Particle", "Category", "KeyElement", "Size", "Band", "ExcelRow"]].rename(
            columns={"KeyElement": "Key Element", "Size": "Size (µm)", "Band": "Size Band", "ExcelRow": "Excel Row"})
        if q:
            mask = log[["Process", "Machine", "Location", "Category"]].apply(lambda s: s.str.contains(q, case=False, regex=False)).any(axis=1)
            log = log[mask]
        st.dataframe(log, use_container_width=True, hide_index=True, height=560,
                     column_config={"Size (µm)": st.column_config.NumberColumn(format="%.2f"),
                                    "Particle": st.column_config.NumberColumn(format="#%d")})
        st.download_button("Download as CSV", log.to_csv(index=False).encode(), file_name="particle_log.csv", mime="text/csv")
