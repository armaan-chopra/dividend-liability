"""
PORTFOLIO TRUE EXPOSURE TERMINAL
---------------------------------
A Bloomberg-terminal styled dashboard that computes and visualizes an
investor's ACTUAL (look-through) exposure to individual stocks held
indirectly via a portfolio of ETFs.

Run with:
    streamlit run app.py

Expected columns (case/spacing-insensitive, both CSV and Excel supported):
    SL. NO.
    ETF IN FUND
    WEIGHT OF ETF IN FUND
    ETF'S HOLDING
    WEIGHT  IN ETF
    TRUE EXPOSURE IN PORTFOLIO
"""

import io
import re
from datetime import datetime

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import networkx as nx
import streamlit as st

# ======================================================================================
# PAGE CONFIG
# ======================================================================================
st.set_page_config(
    page_title="EXPOSURE TERMINAL",
    page_icon=None,
    layout="wide",
    initial_sidebar_state="expanded",
)

# ======================================================================================
# THEME / CSS  -- Bloomberg-terminal look: black background, yellow + white text
# ======================================================================================
YELLOW = "#FFCC00"
AMBER = "#FF9900"
WHITE = "#F5F5F5"
BLACK = "#000000"
PANEL = "#0A0A0A"
GRID = "#262626"
GREEN = "#00E676"
RED = "#FF3B30"

st.markdown(f"""
<style>
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500;600;700&display=swap');

    html, body, [class*="css"] {{
        font-family: 'IBM Plex Mono', 'Consolas', monospace !important;
    }}

    .stApp {{
        background-color: {BLACK};
        color: {WHITE};
    }}

    section[data-testid="stSidebar"] {{
        background-color: {PANEL};
        border-right: 1px solid {GRID};
    }}

    h1, h2, h3, h4 {{
        color: {YELLOW} !important;
        font-family: 'IBM Plex Mono', monospace !important;
        letter-spacing: 1px;
        text-transform: uppercase;
    }}

    p, span, label, div, li {{
        color: {WHITE};
    }}

    .terminal-header {{
        background-color: {PANEL};
        border: 1px solid {YELLOW};
        padding: 14px 22px;
        margin-bottom: 18px;
        display: flex;
        justify-content: space-between;
        align-items: center;
    }}
    .terminal-title {{
        color: {YELLOW};
        font-size: 26px;
        font-weight: 700;
        letter-spacing: 3px;
        margin: 0;
    }}
    .terminal-sub {{
        color: {AMBER};
        font-size: 12px;
        letter-spacing: 2px;
    }}

    /* Stat boxes */
    .stat-box {{
        background-color: {PANEL};
        border: 1px solid {GRID};
        border-left: 3px solid {YELLOW};
        padding: 14px 16px;
        margin-bottom: 12px;
        height: 100%;
    }}
    .stat-label {{
        color: {AMBER};
        font-size: 11px;
        letter-spacing: 1.5px;
        text-transform: uppercase;
        margin-bottom: 6px;
    }}
    .stat-value {{
        color: {YELLOW};
        font-size: 26px;
        font-weight: 700;
        line-height: 1.1;
    }}
    .stat-sub {{
        color: #999999;
        font-size: 11px;
        margin-top: 4px;
    }}
    .stat-value.white {{ color: {WHITE}; }}
    .stat-value.green {{ color: {GREEN}; }}
    .stat-value.red {{ color: {RED}; }}

    .risk-flag-high {{ color: {RED}; font-weight: 700; }}
    .risk-flag-med  {{ color: {AMBER}; font-weight: 700; }}
    .risk-flag-low  {{ color: {GREEN}; font-weight: 700; }}

    /* Tabs */
    .stTabs [data-baseweb="tab-list"] {{
        gap: 4px;
        border-bottom: 1px solid {GRID};
    }}
    .stTabs [data-baseweb="tab"] {{
        background-color: {PANEL};
        color: {WHITE};
        border: 1px solid {GRID};
        border-bottom: none;
        padding: 10px 18px;
        font-size: 12px;
        letter-spacing: 1.5px;
        text-transform: uppercase;
    }}
    .stTabs [aria-selected="true"] {{
        background-color: #1A1400 !important;
        color: {YELLOW} !important;
        border-color: {YELLOW} !important;
    }}

    /* DataFrames */
    .stDataFrame {{ border: 1px solid {GRID}; }}

    /* Buttons */
    .stButton>button, .stDownloadButton>button {{
        background-color: {BLACK};
        color: {YELLOW};
        border: 1px solid {YELLOW};
        border-radius: 0px;
        letter-spacing: 1px;
        text-transform: uppercase;
        font-size: 12px;
    }}
    .stButton>button:hover, .stDownloadButton>button:hover {{
        background-color: {YELLOW};
        color: {BLACK};
    }}

    /* File uploader */
    [data-testid="stFileUploader"] {{
        border: 1px dashed {GRID};
        padding: 8px;
    }}

    /* Metric widget override */
    div[data-testid="stMetric"] {{
        background-color: {PANEL};
        border: 1px solid {GRID};
        border-left: 3px solid {YELLOW};
        padding: 10px 14px;
    }}
    div[data-testid="stMetricLabel"] {{ color: {AMBER} !important; }}
    div[data-testid="stMetricValue"] {{ color: {YELLOW} !important; }}

    hr {{ border-color: {GRID}; }}

    ::-webkit-scrollbar {{ width: 10px; height: 10px; }}
    ::-webkit-scrollbar-track {{ background: {BLACK}; }}
    ::-webkit-scrollbar-thumb {{ background: {GRID}; }}
</style>
""", unsafe_allow_html=True)

PLOTLY_LAYOUT = dict(
    paper_bgcolor=BLACK,
    plot_bgcolor=BLACK,
    font=dict(color=WHITE, family="IBM Plex Mono, Consolas, monospace"),
    xaxis=dict(gridcolor=GRID, zerolinecolor=GRID),
    yaxis=dict(gridcolor=GRID, zerolinecolor=GRID),
    margin=dict(l=10, r=10, t=50, b=10),
    legend=dict(bgcolor="rgba(0,0,0,0)"),
)


def themed_layout(**overrides):
    """Merge chart-specific overrides on top of PLOTLY_LAYOUT without duplicate-keyword
    errors — nested dict keys like xaxis/yaxis are merged rather than replaced wholesale."""
    merged = dict(PLOTLY_LAYOUT)
    for key in ("xaxis", "yaxis", "font", "legend"):
        if key in overrides:
            merged_sub = dict(merged.get(key, {}))
            merged_sub.update(overrides.pop(key))
            merged[key] = merged_sub
    merged.update(overrides)
    return merged

# ======================================================================================
# COLUMN NORMALIZATION & DATA LOADING
# ======================================================================================
STD_COLS = ["SL_NO", "ETF", "ETF_WEIGHT_IN_FUND", "HOLDING", "HOLDING_WEIGHT_IN_ETF", "TRUE_EXPOSURE_REPORTED"]


def _clean_header(c):
    c = str(c).strip().upper()
    c = c.replace("’", "'")
    c = re.sub(r"\s+", " ", c)
    return c


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Map arbitrarily-spaced/cased source headers onto a standard internal schema."""
    df = df.copy()
    df.columns = [_clean_header(c) for c in df.columns]
    rename = {}
    for c in df.columns:
        if "SL" in c and "NO" in c:
            rename[c] = "SL_NO"
        elif "EXPOSURE" in c:
            rename[c] = "TRUE_EXPOSURE_REPORTED"
        elif "HOLDING" in c and "WEIGHT" not in c:
            rename[c] = "HOLDING"
        elif "WEIGHT" in c and "FUND" in c:
            rename[c] = "ETF_WEIGHT_IN_FUND"
        elif "WEIGHT" in c and "ETF" in c and "FUND" not in c:
            rename[c] = "HOLDING_WEIGHT_IN_ETF"
        elif "ETF" in c and "FUND" in c and "WEIGHT" not in c:
            rename[c] = "ETF"
        elif c == "ETF":
            rename[c] = "ETF"
    df = df.rename(columns=rename)

    for col in STD_COLS:
        if col not in df.columns:
            df[col] = np.nan

    df = df[STD_COLS]
    df["ETF"] = df["ETF"].astype(str).str.strip()
    df["HOLDING"] = df["HOLDING"].astype(str).str.strip()
    for c in ["ETF_WEIGHT_IN_FUND", "HOLDING_WEIGHT_IN_ETF", "TRUE_EXPOSURE_REPORTED"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    df = df.dropna(subset=["ETF", "HOLDING"])
    df = df[(df["ETF"] != "") & (df["ETF"] != "nan") & (df["HOLDING"] != "") & (df["HOLDING"] != "nan")]
    return df.reset_index(drop=True)


def read_any(uploaded_file) -> pd.DataFrame:
    """Read CSV or Excel transparently."""
    name = uploaded_file.name.lower()
    raw = uploaded_file.read()
    buf = io.BytesIO(raw)
    if name.endswith(".csv") or name.endswith(".txt"):
        df = pd.read_csv(buf)
    elif name.endswith(".xlsx") or name.endswith(".xls"):
        df = pd.read_excel(buf)
    else:
        # try csv first, fall back to excel
        try:
            buf.seek(0)
            df = pd.read_csv(buf)
        except Exception:
            buf.seek(0)
            df = pd.read_excel(buf)
    return normalize_columns(df)


def compute_exposure(df: pd.DataFrame) -> pd.DataFrame:
    """Use TRUE EXPOSURE IN PORTFOLIO exactly as provided in the source file (no recalculation).
    Auto-detects whether weight/exposure columns are fractions (0-1) or percentages (0-100)
    so everything is displayed on a consistent 0-100 scale."""
    df = df.copy()

    def to_pct_scale(s):
        s = s.astype(float)
        if s.dropna().empty:
            return s
        return s * 100 if s.dropna().max() <= 1.5 else s

    df["ETF_WEIGHT_IN_FUND"] = to_pct_scale(df["ETF_WEIGHT_IN_FUND"])
    df["HOLDING_WEIGHT_IN_ETF"] = to_pct_scale(df["HOLDING_WEIGHT_IN_ETF"])
    df["EXPOSURE"] = to_pct_scale(df["TRUE_EXPOSURE_REPORTED"])
    return df


def merge_data(existing: pd.DataFrame, new: pd.DataFrame, mode: str) -> pd.DataFrame:
    if existing is None or existing.empty or mode == "Replace all data":
        return new.reset_index(drop=True)
    combined = pd.concat([existing[STD_COLS], new[STD_COLS]], ignore_index=True)
    # upsert on (ETF, HOLDING): keep the LAST occurrence (i.e. the newly-uploaded row wins)
    combined["_key"] = combined["ETF"] + "||" + combined["HOLDING"]
    combined = combined.drop_duplicates(subset="_key", keep="last").drop(columns="_key")
    return combined.reset_index(drop=True)


# ======================================================================================
# SIDEBAR — DATA INGESTION
# ======================================================================================
if "master_df" not in st.session_state:
    st.session_state.master_df = None
if "upload_log" not in st.session_state:
    st.session_state.upload_log = []

with st.sidebar:
    st.markdown(f"<div class='terminal-sub'>DATA INGESTION</div>", unsafe_allow_html=True)
    st.markdown("### Load Portfolio Data")

    mode = st.radio(
        "Upload mode",
        ["Replace all data", "Merge / update existing data"],
        index=0 if st.session_state.master_df is None else 1,
        help="Merge mode upserts rows by (ETF, Holding) — matching rows are updated, new ones appended, everything else is kept.",
    )

    uploaded = st.file_uploader(
        "Upload CSV or Excel (.csv, .xlsx, .xls)",
        type=["csv", "xlsx", "xls"],
        accept_multiple_files=False,
    )

    if uploaded is not None:
        try:
            new_df = read_any(uploaded)
            st.session_state.master_df = merge_data(st.session_state.master_df, new_df, mode)
            st.session_state.upload_log.append(
                f"{datetime.now().strftime('%H:%M:%S')} — {uploaded.name} ({mode}) — {len(new_df)} rows"
            )
            st.success(f"Loaded {len(new_df)} rows from {uploaded.name}")
        except Exception as e:
            st.error(f"Failed to parse file: {e}")

    if st.session_state.master_df is not None and not st.session_state.master_df.empty:
        if st.button("Clear all data"):
            st.session_state.master_df = None
            st.session_state.upload_log = []
            st.rerun()

    if st.session_state.upload_log:
        st.markdown("---")
        st.markdown("<div class='terminal-sub'>UPLOAD LOG</div>", unsafe_allow_html=True)
        for entry in st.session_state.upload_log[-8:][::-1]:
            st.caption(entry)

    st.markdown("---")
    st.markdown(
        "<div class='terminal-sub'>EXPECTED COLUMNS</div>"
        "<div style='font-size:11px; color:#999'>"
        "SL. NO.<br>ETF IN FUND<br>WEIGHT OF ETF IN FUND<br>"
        "ETF'S HOLDING<br>WEIGHT IN ETF<br>TRUE EXPOSURE IN PORTFOLIO"
        "</div>",
        unsafe_allow_html=True,
    )

# ======================================================================================
# HEADER
# ======================================================================================
st.markdown(f"""
<div class="terminal-header">
    <div>
        <div class="terminal-title">PORTFOLIO TRUE EXPOSURE TERMINAL</div>
        <div class="terminal-sub">LOOK-THROUGH STOCK EXPOSURE ANALYTICS — ETF PORTFOLIO</div>
    </div>
    <div style="text-align:right;">
        <div class="terminal-sub">SESSION</div>
        <div style="color:{WHITE}; font-size:13px;">{datetime.now().strftime('%d %b %Y  %H:%M')}</div>
    </div>
</div>
""", unsafe_allow_html=True)

if st.session_state.master_df is None or st.session_state.master_df.empty:
    st.info("Upload a CSV or Excel file from the sidebar to begin. Columns can be in any order / spacing — "
            "they are matched automatically to SL. NO., ETF IN FUND, WEIGHT OF ETF IN FUND, "
            "ETF'S HOLDING, WEIGHT IN ETF, TRUE EXPOSURE IN PORTFOLIO.")
    st.stop()

df = compute_exposure(st.session_state.master_df)

# ======================================================================================
# CORE AGGREGATIONS (used across tabs)
# ======================================================================================
by_holding = (
    df.groupby("HOLDING", as_index=False)
    .agg(EXPOSURE=("EXPOSURE", "sum"), N_ETFS=("ETF", "nunique"))
    .sort_values("EXPOSURE", ascending=False)
    .reset_index(drop=True)
)

by_etf = (
    df.groupby("ETF", as_index=False)
    .agg(
        FUND_WEIGHT=("ETF_WEIGHT_IN_FUND", "first"),
        N_HOLDINGS=("HOLDING", "nunique"),
        TOTAL_EXPOSURE=("EXPOSURE", "sum"),
    )
    .sort_values("FUND_WEIGHT", ascending=False)
    .reset_index(drop=True)
)

total_etfs = df["ETF"].nunique()
total_holdings = df["HOLDING"].nunique()
total_rows = len(df)
fund_weight_sum = by_etf["FUND_WEIGHT"].sum()
total_true_exposure = by_holding["EXPOSURE"].sum()

shares = by_holding["EXPOSURE"].values
hhi = float(np.sum(shares ** 2))  # shares already in % points (0-100 scale) => HHI in 0-10000 range

top1 = by_holding["EXPOSURE"].iloc[0] if len(by_holding) else 0
top5 = by_holding["EXPOSURE"].iloc[:5].sum()
top10 = by_holding["EXPOSURE"].iloc[:10].sum()

overlap_holdings = by_holding[by_holding["N_ETFS"] > 1]
avg_etfs_per_holding = by_holding["N_ETFS"].mean() if len(by_holding) else 0

if hhi >= 1500:
    conc_level, conc_class = "HIGH", "risk-flag-high"
elif hhi >= 800:
    conc_level, conc_class = "MODERATE", "risk-flag-med"
else:
    conc_level, conc_class = "LOW", "risk-flag-low"

overlap_pct = (len(overlap_holdings) / total_holdings * 100) if total_holdings else 0
if overlap_pct >= 40:
    overlap_level, overlap_class = "HIGH", "risk-flag-high"
elif overlap_pct >= 15:
    overlap_level, overlap_class = "MODERATE", "risk-flag-med"
else:
    overlap_level, overlap_class = "LOW", "risk-flag-low"


def stat_box(label, value, sub="", color="yellow"):
    st.markdown(f"""
    <div class="stat-box">
        <div class="stat-label">{label}</div>
        <div class="stat-value {color}">{value}</div>
        <div class="stat-sub">{sub}</div>
    </div>
    """, unsafe_allow_html=True)


# ======================================================================================
# TABS
# ======================================================================================
tab_lookup, tab_stats, tab_viz, tab_raw = st.tabs(
    ["STOCK EXPOSURE LOOKUP", "KEY STATISTICS", "VISUALIZATIONS", "RAW DATA"]
)

# --------------------------------------------------------------------------------------
# TAB 1 — STOCK LOOKUP + RISK METRICS
# --------------------------------------------------------------------------------------
with tab_lookup:
    st.markdown("#### Actual Exposure to a Specific Stock")
    search = st.selectbox(
        "Select / search holding",
        options=by_holding["HOLDING"].tolist(),
        index=0,
    )

    row = by_holding[by_holding["HOLDING"] == search].iloc[0]
    detail = df[df["HOLDING"] == search][["ETF", "ETF_WEIGHT_IN_FUND", "HOLDING_WEIGHT_IN_ETF", "EXPOSURE"]] \
        .sort_values("EXPOSURE", ascending=False).reset_index(drop=True)
    detail.columns = ["ETF", "ETF Weight in Fund (%)", "Weight in ETF (%)", "True Exposure (%)"]

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        stat_box("Total True Exposure", f"{row['EXPOSURE']:.3f}%", f"of overall portfolio")
    with c2:
        stat_box("Contributing ETFs", f"{int(row['N_ETFS'])}", "ETF(s) holding this stock")
    with c3:
        rank = int(by_holding.index[by_holding["HOLDING"] == search][0]) + 1
        stat_box("Concentration Rank", f"#{rank}", f"of {total_holdings} holdings")
    with c4:
        pct_of_portfolio_value = row["EXPOSURE"] / total_true_exposure * 100 if total_true_exposure else 0
        stat_box("Share of Total Look-Through", f"{pct_of_portfolio_value:.2f}%", "vs. all holdings combined")

    left, right = st.columns([1.3, 1])
    with left:
        st.markdown("##### Breakdown by ETF")
        st.dataframe(detail, width="stretch", hide_index=True)
    with right:
        fig = px.bar(
            detail, x="True Exposure (%)", y="ETF", orientation="h",
            color_discrete_sequence=[YELLOW],
        )
        fig.update_layout(**themed_layout(height=max(260, 40 * len(detail)), title="Exposure Contribution by ETF"))
        fig.update_traces(marker_line_color=AMBER, marker_line_width=1)
        st.plotly_chart(fig, width="stretch")

    st.markdown("---")
    st.markdown("#### Key Risk Metrics — Whole Portfolio")

    r1, r2, r3, r4 = st.columns(4)
    with r1:
        stat_box("HHI (Look-Through)", f"{hhi:,.0f}", "Herfindahl-Hirschman Index")
    with r2:
        stat_box("Concentration Risk", f"<span class='{conc_class}'>{conc_level}</span>".replace("<span", "<span "), "")
    with r3:
        stat_box("ETF Overlap Risk", f"<span class='{overlap_class}'>{overlap_level}</span>", f"{overlap_pct:.1f}% of holdings shared")
    with r4:
        stat_box("Fund Weight Check", f"{fund_weight_sum:.1f}%", "should ≈ 100%",
                  color="green" if 98 <= fund_weight_sum <= 102 else "red")

    st.markdown("##### Top 10 Concentrated Holdings (True Exposure)")
    top10_df = by_holding.head(10).copy()
    top10_df["EXPOSURE"] = top10_df["EXPOSURE"].round(3)
    fig2 = px.bar(top10_df, x="EXPOSURE", y="HOLDING", orientation="h",
                  color="EXPOSURE", color_continuous_scale=["#332900", YELLOW])
    fig2.update_layout(**themed_layout(height=380, yaxis=dict(autorange="reversed", gridcolor=GRID),
                        coloraxis_showscale=False, title="Top 10 Holdings by True Portfolio Exposure (%)"))
    st.plotly_chart(fig2, width="stretch")

# --------------------------------------------------------------------------------------
# TAB 2 — KEY STATISTICS
# --------------------------------------------------------------------------------------
with tab_stats:
    st.markdown("#### Portfolio Snapshot")
    a1, a2, a3, a4 = st.columns(4)
    with a1:
        stat_box("Total ETFs", f"{total_etfs}", "distinct ETFs in portfolio")
    with a2:
        stat_box("Unique Underlying Holdings", f"{total_holdings}", "distinct stocks / assets")
    with a3:
        stat_box("Total Data Rows", f"{total_rows}", "ETF × holding pairs")
    with a4:
        stat_box("Total Look-Through Exposure", f"{total_true_exposure:.1f}%", "sum of all true exposures")

    b1, b2, b3, b4 = st.columns(4)
    with b1:
        stat_box("Top 1 Holding Weight", f"{top1:.2f}%", by_holding['HOLDING'].iloc[0] if len(by_holding) else "")
    with b2:
        stat_box("Top 5 Holdings Weight", f"{top5:.2f}%", "combined exposure")
    with b3:
        stat_box("Top 10 Holdings Weight", f"{top10:.2f}%", "combined exposure")
    with b4:
        stat_box("Avg ETFs per Holding", f"{avg_etfs_per_holding:.2f}", "overlap intensity")

    st.markdown("---")
    st.markdown("#### ETF Breakdown")
    etf_disp = by_etf.copy()
    etf_disp.columns = ["ETF", "Weight in Fund (%)", "# Holdings Disclosed", "Total Look-Through Exposure (%)"]
    etf_disp["Weight in Fund (%)"] = etf_disp["Weight in Fund (%)"].round(2)
    etf_disp["Total Look-Through Exposure (%)"] = etf_disp["Total Look-Through Exposure (%)"].round(3)

    s1, s2 = st.columns([1, 1])
    with s1:
        st.dataframe(etf_disp, width="stretch", hide_index=True, height=380)
    with s2:
        fig3 = px.pie(by_etf, names="ETF", values="FUND_WEIGHT", hole=0.55,
                       color_discrete_sequence=px.colors.sequential.YlOrBr_r)
        fig3.update_layout(**themed_layout(height=380, title="ETF Weight Composition of Fund"))
        fig3.update_traces(textfont_color=BLACK, marker=dict(line=dict(color=BLACK, width=1)))
        st.plotly_chart(fig3, width="stretch")

    st.markdown("---")
    st.markdown("#### Full Holdings Table (Aggregated Across All ETFs)")
    disp_holdings = by_holding.copy()
    disp_holdings["EXPOSURE"] = disp_holdings["EXPOSURE"].round(3)
    disp_holdings.columns = ["Holding", "True Exposure (%)", "# ETFs Providing Exposure"]
    st.dataframe(disp_holdings, width="stretch", hide_index=True, height=420)

    csv_out = disp_holdings.to_csv(index=False).encode("utf-8")
    st.download_button("Download aggregated holdings (CSV)", csv_out, "aggregated_true_exposure.csv", "text/csv")

# --------------------------------------------------------------------------------------
# TAB 3 — VISUALIZATIONS
# --------------------------------------------------------------------------------------
with tab_viz:
    st.markdown("#### Exposure Flow — ETF → Underlying Holding")

    top_n_viz = st.slider("Limit to top N holdings (for readability)", 5, min(60, total_holdings),
                           value=min(20, total_holdings))
    top_holdings_set = set(by_holding.head(top_n_viz)["HOLDING"])
    viz_df = df[df["HOLDING"].isin(top_holdings_set)]

    # --- Sankey ---
    etfs_list = sorted(viz_df["ETF"].unique().tolist())
    holdings_list = sorted(viz_df["HOLDING"].unique().tolist())
    node_labels = etfs_list + holdings_list
    node_idx = {name: i for i, name in enumerate(node_labels)}
    node_colors = [YELLOW] * len(etfs_list) + [AMBER] * len(holdings_list)

    link_src, link_tgt, link_val = [], [], []
    for _, r in viz_df.iterrows():
        link_src.append(node_idx[r["ETF"]])
        link_tgt.append(node_idx[r["HOLDING"]])
        link_val.append(max(r["EXPOSURE"], 0.001))

    sankey = go.Figure(data=[go.Sankey(
        node=dict(
            pad=12, thickness=14,
            line=dict(color=GRID, width=0.5),
            label=node_labels, color=node_colors,
        ),
        link=dict(source=link_src, target=link_tgt, value=link_val,
                   color="rgba(255,204,0,0.15)"),
    )])
    sankey.update_layout(**themed_layout(height=560, title="ETF → Holding Exposure Flow (width = true exposure)"))
    st.plotly_chart(sankey, width="stretch")

    st.markdown("---")
    v1, v2 = st.columns(2)

    with v1:
        st.markdown("##### Exposure Network Graph")
        G = nx.Graph()
        for e in etfs_list:
            G.add_node(e, kind="ETF")
        for h in holdings_list:
            G.add_node(h, kind="HOLDING")
        for _, r in viz_df.iterrows():
            G.add_edge(r["ETF"], r["HOLDING"], weight=r["EXPOSURE"])

        pos = nx.spring_layout(G, k=0.6, seed=42)
        edge_x, edge_y = [], []
        for u, v in G.edges():
            x0, y0 = pos[u]; x1, y1 = pos[v]
            edge_x += [x0, x1, None]
            edge_y += [y0, y1, None]
        edge_trace = go.Scatter(x=edge_x, y=edge_y, line=dict(width=0.7, color="#555555"),
                                 hoverinfo="none", mode="lines")

        node_x, node_y, node_txt, node_col, node_size = [], [], [], [], []
        exp_by_node = by_holding.set_index("HOLDING")["EXPOSURE"].to_dict()
        for n in G.nodes():
            x, y = pos[n]
            node_x.append(x); node_y.append(y); node_txt.append(n)
            if G.nodes[n]["kind"] == "ETF":
                node_col.append(YELLOW)
                node_size.append(22)
            else:
                node_col.append(AMBER)
                node_size.append(8 + min(exp_by_node.get(n, 0) * 3, 30))
        node_trace = go.Scatter(x=node_x, y=node_y, mode="markers", hoverinfo="text", text=node_txt,
                                 marker=dict(color=node_col, size=node_size, line=dict(color=BLACK, width=1)))

        netfig = go.Figure(data=[edge_trace, node_trace])
        netfig.update_layout(**themed_layout(height=480, showlegend=False,
                              xaxis=dict(visible=False), yaxis=dict(visible=False),
                              title="ETF (yellow) — Holding (amber) Network"))
        st.plotly_chart(netfig, width="stretch")

    with v2:
        st.markdown("##### Exposure Treemap")
        tdf = viz_df.copy()
        tdf["EXPOSURE"] = tdf["EXPOSURE"].clip(lower=0.0001)
        fig_tree = px.treemap(
            tdf, path=["ETF", "HOLDING"], values="EXPOSURE",
            color="EXPOSURE", color_continuous_scale=["#1a1400", YELLOW],
        )
        fig_tree.update_layout(**themed_layout(height=480, title="ETF → Holding Exposure Treemap",
                                coloraxis_showscale=False))
        fig_tree.update_traces(marker=dict(line=dict(color=BLACK, width=1)), textfont_color=BLACK)
        st.plotly_chart(fig_tree, width="stretch")

    st.markdown("---")
    st.markdown("##### Stock-Level Overlap Across ETFs")
    st.caption("Which individual stocks are held by more than one ETF, and how much true exposure each ETF contributes to that stock.")

    overlap_holdings_viz = (
        by_holding[(by_holding["N_ETFS"] > 1) & (by_holding["HOLDING"].isin(top_holdings_set))]
        .sort_values("EXPOSURE", ascending=False)["HOLDING"].tolist()
    )

    if overlap_holdings_viz:
        pivot = (
            viz_df[viz_df["HOLDING"].isin(overlap_holdings_viz)]
            .pivot_table(index="HOLDING", columns="ETF", values="EXPOSURE", aggfunc="sum", fill_value=0)
        )
        pivot = pivot.reindex(overlap_holdings_viz)

        stock_overlap_heat = go.Figure(data=go.Heatmap(
            z=pivot.values, x=pivot.columns.tolist(), y=pivot.index.tolist(),
            colorscale=[[0, "#0A0A0A"], [1, YELLOW]],
            colorbar=dict(title="Exposure %", tickfont=dict(color=WHITE)),
            hovertemplate="Holding: %{y}<br>ETF: %{x}<br>True Exposure: %{z:.3f}%<extra></extra>",
        ))
        stock_overlap_heat.update_layout(**themed_layout(
            height=max(360, 26 * len(pivot)),
            title="Stocks Held by Multiple ETFs (cell = true exposure %)",
        ))
        st.plotly_chart(stock_overlap_heat, width="stretch")
    else:
        st.info("None of the currently displayed holdings are shared across more than one ETF. "
                 "Increase the 'Limit to top N holdings' slider above to check more of the portfolio.")

    st.markdown("---")
    st.markdown("##### ETF Overlap Matrix (shared holdings, Jaccard similarity)")
    etf_holdings = df.groupby("ETF")["HOLDING"].apply(set)
    etfs_all = etf_holdings.index.tolist()
    mat = np.zeros((len(etfs_all), len(etfs_all)))
    for i, e1 in enumerate(etfs_all):
        for j, e2 in enumerate(etfs_all):
            a, b = etf_holdings[e1], etf_holdings[e2]
            union = len(a | b)
            mat[i, j] = len(a & b) / union if union else 0

    heat = go.Figure(data=go.Heatmap(
        z=mat, x=etfs_all, y=etfs_all,
        colorscale=[[0, "#0A0A0A"], [1, YELLOW]],
        colorbar=dict(title="Jaccard", tickfont=dict(color=WHITE)),
    ))
    heat.update_layout(**themed_layout(height=max(360, 28 * len(etfs_all)), title="Overlap Between ETFs (1.0 = identical holdings)"))
    st.plotly_chart(heat, width="stretch")

# --------------------------------------------------------------------------------------
# TAB 4 — RAW DATA
# --------------------------------------------------------------------------------------
with tab_raw:
    st.markdown("#### Raw / Combined Dataset (post-merge)")
    show_df = df.copy()
    show_df = show_df[["SL_NO", "ETF", "ETF_WEIGHT_IN_FUND", "HOLDING", "HOLDING_WEIGHT_IN_ETF", "EXPOSURE"]]
    show_df.columns = ["SL No", "ETF", "ETF Weight in Fund (%)", "Holding", "Weight in ETF (%)",
                        "True Exposure (%)"]
    st.dataframe(show_df, width="stretch", hide_index=True, height=520)

    csv_full = show_df.to_csv(index=False).encode("utf-8")
    st.download_button("Download full dataset (CSV)", csv_full, "portfolio_true_exposure_full.csv", "text/csv")