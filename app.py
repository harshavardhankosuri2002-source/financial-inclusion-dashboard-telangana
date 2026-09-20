"""
Financial Inclusion & MSME Credit Access — Telangana
District Intelligence Dashboard
Group 10 | MBA Analytics Project

A production-quality interactive Plotly Dash dashboard for analyzing
financial inclusion metrics across Telangana districts with simulated
real-time streaming architecture.

Run:   python app.py
Open:  http://127.0.0.1:8050/
"""

# ─────────────────────────────────────────────────────────────────────
# Imports
# ─────────────────────────────────────────────────────────────────────
import os
import sys
import time
from datetime import datetime
from typing import Optional, Dict, List, Any, Tuple

if sys.platform == "win32" and hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

import dash
from dash import dcc, html, dash_table
from dash.dependencies import Input, Output, State
import dash_bootstrap_components as dbc
import plotly.graph_objects as go
import pandas as pd
import numpy as np


# ═════════════════════════════════════════════════════════════════════
# CONSTANTS & CONFIGURATION
# ═════════════════════════════════════════════════════════════════════

EXCEL_FILE = (
    "Group_10_Financial_Inclusion_MSME_Credit_Access_final.xlsx"
    if os.path.exists("Group_10_Financial_Inclusion_MSME_Credit_Access_final.xlsx")
    else "Group_10_Financial_Inclusion_MSME_Credit_Access_PRO.xlsx"
)

COLORS: Dict[str, str] = {
    "primary":     "#163A5F",  # navy
    "blue":        "#2F6690",  # blue
    "medium_blue": "#4F86C6",  # medium blue
    "light_blue":  "#EAF2F8",  # light blue
    "bg":          "#F5F7FA",  # background
    "white":       "#FFFFFF",  # white
    "text":        "#263238",  # dark text
    "success":     "#27AE60",  # success green
    "warning":     "#F39C12",  # warning amber
    "danger":      "#C0392B",  # danger red
    "border":      "#E8ECF1",  # clean card border
    "muted":       "#8A9BAE",  # subtitle & label muted
}

PRODUCTS: List[str] = ["Shishu", "Kishore", "Tarun", "Tarun Plus"]
PRODUCT_COLORS: List[str] = ["#163A5F", "#2F6690", "#4F86C6", "#7FB3D8"]

STREAM_SPEEDS: Dict[str, int] = {
    "Slow": 3000,
    "Normal": 1500,
    "Fast": 750,
}

CARD_STYLE: Dict[str, Any] = {
    "borderRadius": "12px",
    "boxShadow": "0 2px 10px rgba(22, 58, 95, 0.06)",
    "border": f"1px solid {COLORS['border']}",
    "backgroundColor": COLORS["white"],
    "marginBottom": "16px",
}

CHART_LAYOUT: Dict[str, Any] = dict(
    template="plotly_white",
    font=dict(family="Inter, -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif", size=11, color=COLORS["text"]),
    margin=dict(l=20, r=25, t=50, b=35),
    paper_bgcolor="rgba(0,0,0,0)",
    plot_bgcolor="rgba(0,0,0,0)",
)


# ═════════════════════════════════════════════════════════════════════
# DATA UTILITIES & NORMALIZATION
# ═════════════════════════════════════════════════════════════════════

def safe_numeric(series: pd.Series) -> pd.Series:
    """Clean and safely convert series to numeric values, preserving NaN."""
    cleaned = (
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("₹", "", regex=False)
        .str.replace("—", "", regex=False)
        .str.replace("–", "", regex=False)
        .str.strip()
    )
    cleaned = cleaned.replace({"": np.nan, "-": np.nan, "nan": np.nan, "None": np.nan, "N/A": np.nan})
    return pd.to_numeric(cleaned, errors="coerce")


def fmt_cr(val: Any) -> str:
    """Format value in ₹ Cr."""
    if val is None or pd.isna(val):
        return "N/A"
    return f"₹{val:,.2f} Cr"


def fmt_num(val: Any) -> str:
    """Format integer with thousands separator."""
    if val is None or pd.isna(val):
        return "N/A"
    return f"{int(round(val)):,}"


def fmt_pct(val: Any) -> str:
    """Format percentage value."""
    if val is None or pd.isna(val):
        return "N/A"
    return f"{val:.2f}%"


def prod_key(product: str) -> str:
    """Standardize a product name to an internal key."""
    return product.lower().replace(" ", "_").replace("+", "plus")


def calculate_latency(t_start: float) -> float:
    """Calculate dashboard processing latency in milliseconds using monotonic perf_counter.

    T1 = event generation timestamp (time.perf_counter)
    T2 = processing complete timestamp (time.perf_counter)
    Processing Latency = (T2 - T1) * 1000 ms
    """
    return round((time.perf_counter() - t_start) * 1000, 2)


# ═════════════════════════════════════════════════════════════════════
# DATA LOADING & PREPARATION (Modular functions)
# ═════════════════════════════════════════════════════════════════════

def _read_sheet_with_header_detection(xl: pd.ExcelFile, sheet_name: str) -> pd.DataFrame:
    """Detect the header row dynamically by scanning the first 10 rows for district-like markers."""
    raw = pd.read_excel(xl, sheet_name=sheet_name, header=None, nrows=10)
    header_idx = 0
    for idx, row in raw.iterrows():
        vals = [str(x).strip().lower() for x in row if pd.notna(x)]
        if any(v in ["district", "district name", "source district name"] for v in vals):
            header_idx = idx
            break
    return pd.read_excel(xl, sheet_name=sheet_name, header=header_idx)


def load_data() -> Dict[str, Any]:
    """Dynamically load the verified Excel dataset and return structured data.

    Raises FileNotFoundError or ValueError with explicit instructions if files
    or required sheets are missing.
    """
    base_dir = os.path.dirname(os.path.abspath(__file__))
    filepath = os.path.join(base_dir, EXCEL_FILE)

    if not os.path.exists(filepath):
        fallback_name = "Group_10_Financial_Inclusion_MSME_Credit_Access_final.xlsx"
        fallback_path = os.path.join(base_dir, fallback_name)
        if os.path.exists(fallback_path):
            filepath = fallback_path
        else:
            raise FileNotFoundError(
                f"Dataset file not found. Place '{EXCEL_FILE}' in the project directory:\n{base_dir}"
            )

    try:
        xl = pd.ExcelFile(filepath, engine="openpyxl")
    except Exception as e:
        raise ValueError(f"Unable to read Excel workbook '{filepath}': {e}")

    if "Integrated_View" not in xl.sheet_names:
        raise ValueError(
            f"Required sheet 'Integrated_View' not found in '{EXCEL_FILE}'.\n"
            f"Available sheets: {', '.join(xl.sheet_names)}"
        )

    df_integrated = _read_sheet_with_header_detection(xl, "Integrated_View")

    aux: Dict[str, pd.DataFrame] = {}
    for s in ["PMMY_Raw", "CD_Ratio", "MSME_Udyog_Aadhaar"]:
        if s in xl.sheet_names:
            aux[s] = _read_sheet_with_header_detection(xl, s)

    return prepare_data(df_integrated, aux)


def prepare_data(df_integrated: pd.DataFrame, aux: Dict[str, pd.DataFrame]) -> Dict[str, Any]:
    """Clean, normalize, and merge PMMY products, CD ratios, and MSME data.

    Ensures 33 current Telangana districts, preserves NaN for missing MSME,
    extracts official totals, and prepares streaming events.
    """
    df_int = df_integrated.copy()
    df_int.columns = [str(c).strip() for c in df_int.columns]

    dist_col = None
    for c in df_int.columns:
        if c.lower() in ["district", "district name", "source district name"]:
            dist_col = c
            break
    if not dist_col:
        dist_col = df_int.columns[0]

    df_int = df_int.rename(columns={dist_col: "District"})

    total_row = df_int[df_int["District"].astype(str).str.lower().str.contains("total")].copy()

    total_msme_official = 246085
    if "MSME_Udyog_Aadhaar" in aux:
        msme_raw = aux["MSME_Udyog_Aadhaar"]
        msme_tot_row = msme_raw[msme_raw.iloc[:, 0].astype(str).str.lower().str.contains("total")]
        if not msme_tot_row.empty:
            for c in msme_tot_row.columns:
                if any(k in str(c).lower() for k in ["total", "udyog"]):
                    val = safe_numeric(msme_tot_row[c]).values[0]
                    if pd.notna(val) and val > 0:
                        total_msme_official = int(val)
                        break

    is_total = df_int["District"].astype(str).str.strip().str.lower().isin(
        ["total", "grand total", "state total", "state", "telangana total", "all", "", "telangana total / source totals"]
    ) | df_int["District"].astype(str).str.lower().str.contains("total")
    df_clean = df_int[~is_total & df_int["District"].notna()].copy()
    df_clean["District"] = df_clean["District"].astype(str).str.strip()

    df_clean["District"] = df_clean["District"].replace({
        "Mahabubnagar": "Mahbubnagar",
    })

    # Exact, non-overlapping column mapping
    col_mapping: Dict[str, str] = {}
    for c in df_clean.columns:
        cl = c.lower().strip()
        if "100 msme" in cl:
            col_mapping[c] = "pmmy_per_100_msme"
        elif "sanction / msme" in cl or ("msme" in cl and "lakh" in cl):
            col_mapping[c] = "sanction_per_msme_lakh"
        elif "avg sanction" in cl:
            col_mapping[c] = "avg_sanction_per_account"
        elif cl in ["pmmy accounts", "pmmy account", "total accounts"]:
            col_mapping[c] = "total_accounts"
        elif "pmmy sanction" in cl and "msme" not in cl:
            col_mapping[c] = "total_sanction_cr"
        elif "cd ratio band" in cl or cl == "band":
            col_mapping[c] = "cd_ratio_band"
        elif "cd ratio" in cl:
            col_mapping[c] = "cd_ratio"
        elif "deposit" in cl:
            col_mapping[c] = "deposits_cr"
        elif "advance" in cl:
            col_mapping[c] = "advances_cr"
        elif "msme" in cl and "aadhaar" in cl:
            col_mapping[c] = "msme_registrations"

    df_clean = df_clean.rename(columns=col_mapping)

    numeric_cols = [
        "total_accounts", "total_sanction_cr", "avg_sanction_per_account",
        "cd_ratio", "deposits_cr", "advances_cr", "msme_registrations",
        "pmmy_per_100_msme", "sanction_per_msme_lakh",
    ]
    for c in numeric_cols:
        if c in df_clean.columns:
            df_clean[c] = safe_numeric(df_clean[c])

    # If PMMY_Raw sheet is present, merge detailed product columns
    if "PMMY_Raw" in aux:
        df_pmmy = aux["PMMY_Raw"].copy()
        df_pmmy.columns = [str(c).strip() for c in df_pmmy.columns]
        p_dist_col = df_pmmy.columns[0]
        df_pmmy = df_pmmy.rename(columns={p_dist_col: "District"})
        df_pmmy["District"] = df_pmmy["District"].astype(str).str.strip().replace({"Mahabubnagar": "Mahbubnagar"})
        df_pmmy = df_pmmy[~df_pmmy["District"].astype(str).str.lower().str.contains("total")].copy()

        product_map = {
            "Shishu Accounts": "shishu_accounts",
            "Shishu ₹Cr": "shishu_sanction_cr",
            "Kishore Accounts": "kishore_accounts",
            "Kishore ₹Cr": "kishore_sanction_cr",
            "Tarun Accounts": "tarun_accounts",
            "Tarun ₹Cr": "tarun_sanction_cr",
            "Tarun Plus Accounts": "tarun_plus_accounts",
            "Tarun Plus ₹Cr": "tarun_plus_sanction_cr",
        }
        pmmy_cols_to_merge = ["District"]
        for col_name, target in product_map.items():
            for c in df_pmmy.columns:
                if c.strip().lower() == col_name.strip().lower():
                    df_pmmy[target] = safe_numeric(df_pmmy[c])
                    pmmy_cols_to_merge.append(target)
                    break

        pmmy_cols_to_merge = list(dict.fromkeys(pmmy_cols_to_merge))
        df_clean = pd.merge(df_clean, df_pmmy[pmmy_cols_to_merge], on="District", how="left")

    if "cd_ratio" in df_clean.columns:
        conditions = [
            df_clean["cd_ratio"] < 100,
            (df_clean["cd_ratio"] >= 100) & (df_clean["cd_ratio"] < 150),
            df_clean["cd_ratio"] >= 150,
        ]
        df_clean["cd_ratio_band"] = np.select(conditions, ["Below 100%", "100–149.99%", "150%+"], default="N/A")

    kpi_accounts = 1005976
    kpi_sanction_cr = 16563.21
    kpi_cd_ratio = 130.79
    kpi_msme = total_msme_official
    kpi_districts = len(df_clean)

    if not total_row.empty:
        r0 = total_row.iloc[0]
        for c in total_row.columns:
            cl = c.lower().strip()
            if cl in ["pmmy accounts", "total accounts", "pmmy account"]:
                val = safe_numeric(pd.Series([r0[c]])).values[0]
                if pd.notna(val): kpi_accounts = int(val)
            elif ("sanction" in cl or "sanc" in cl) and all(k not in cl for k in ["msme", "avg", "rank", "average"]):
                val = safe_numeric(pd.Series([r0[c]])).values[0]
                if pd.notna(val): kpi_sanction_cr = round(float(val), 2)
            elif "cd ratio" in cl and "band" not in cl:
                val = safe_numeric(pd.Series([r0[c]])).values[0]
                if pd.notna(val): kpi_cd_ratio = round(float(val), 2)
            elif "msme" in cl and any(k in cl for k in ["total", "aadhaar", "registrations"]):
                val = safe_numeric(pd.Series([r0[c]])).values[0]
                if pd.notna(val) and val > 0: kpi_msme = int(val)

    kpi = {
        "total_accounts": kpi_accounts,
        "total_sanction_cr": kpi_sanction_cr,
        "state_cd_ratio": kpi_cd_ratio,
        "total_msme": kpi_msme,
        "districts_count": kpi_districts,
    }

    districts = sorted(df_clean["District"].unique().tolist())
    events = create_stream_events(df_clean)

    return {
        "df": df_clean,
        "kpi": kpi,
        "districts": districts,
        "events": events,
        "aux": aux,
    }


def create_stream_events(df: pd.DataFrame) -> List[Dict[str, Any]]:
    """Build a sequential simulation event list from verified historical rows.

    Replays verified district x product records.
    Does NOT generate random or synthetic financial numbers.
    """
    events: List[Dict[str, Any]] = []

    has_products = all(f"{prod_key(p)}_accounts" in df.columns for p in PRODUCTS)

    if has_products:
        for _, row in df.iterrows():
            district = str(row["District"])
            cd_val = row.get("cd_ratio")
            cd_ratio = round(float(cd_val), 2) if pd.notna(cd_val) else None

            for p in PRODUCTS:
                pk = prod_key(p)
                accts = row.get(f"{pk}_accounts")
                sanc = row.get(f"{pk}_sanction_cr")

                if pd.notna(accts) and float(accts) > 0:
                    events.append({
                        "district": district,
                        "product": p,
                        "accounts": int(round(float(accts))),
                        "sanction_cr": round(float(sanc), 2) if pd.notna(sanc) else 0.0,
                        "cd_ratio": cd_ratio,
                    })

    if not events:
        for _, row in df.iterrows():
            cd_val = row.get("cd_ratio")
            events.append({
                "district": str(row["District"]),
                "product": "All Products",
                "accounts": int(round(float(row.get("total_accounts", 0)))),
                "sanction_cr": round(float(row.get("total_sanction_cr", 0)), 2),
                "cd_ratio": round(float(cd_val), 2) if pd.notna(cd_val) else None,
            })

    return events


def filter_data(df: pd.DataFrame, district: str = "All Districts", product: str = "All Products", cd_band: str = "All") -> pd.DataFrame:
    """Filter dataset based on user selections across district, product, and CD ratio band."""
    filtered = df.copy()

    if cd_band and cd_band != "All" and "cd_ratio" in filtered.columns:
        if cd_band == "Below 100%":
            filtered = filtered[filtered["cd_ratio"] < 100]
        elif cd_band in ["100–149.99%", "100-149.99%"]:
            filtered = filtered[(filtered["cd_ratio"] >= 100) & (filtered["cd_ratio"] < 150)]
        elif cd_band in ["150%+", "Above 150%"]:
            filtered = filtered[filtered["cd_ratio"] >= 150]

    if district and district != "All Districts":
        filtered = filtered[filtered["District"] == district]

    return filtered


# ═════════════════════════════════════════════════════════════════════
# VISUALIZATION & CHART BUILDERS (Modular functions)
# ═════════════════════════════════════════════════════════════════════

def _empty_figure(title: str, message: str) -> go.Figure:
    """Return an empty figure with a professional placeholder message."""
    fig = go.Figure()
    fig.update_layout(
        **CHART_LAYOUT,
        title=dict(text=title, font=dict(size=14, color=COLORS["text"], weight="bold")),
        height=380,
        annotations=[dict(
            text=message, xref="paper", yref="paper", x=0.5, y=0.5,
            showarrow=False, font=dict(size=13, color=COLORS["muted"]),
        )],
    )
    return fig


def build_kpi_cards(kpi: Dict[str, Any]) -> dbc.Row:
    """Construct the 5 top-level KPI cards matching official figures."""
    cards_data = [
        ("PMMY Accounts", fmt_num(kpi["total_accounts"]), "1,005,976", COLORS["primary"], "📊"),
        ("PMMY Sanction", fmt_cr(kpi["total_sanction_cr"]), "₹16,563.21 Cr", COLORS["blue"], "💰"),
        ("State CD Ratio", fmt_pct(kpi["state_cd_ratio"]), "130.79%", COLORS["medium_blue"], "📈"),
        ("Udyog Aadhaar Registrations", fmt_num(kpi["total_msme"]), "246,085", COLORS["success"], "🏭"),
        ("Districts Covered", str(kpi["districts_count"]), "33", COLORS["warning"], "📍"),
    ]

    cols = []
    for title, val, _, accent, icon in cards_data:
        cols.append(
            dbc.Col(
                dbc.Card(
                    dbc.CardBody([
                        html.Div([
                            html.Span(icon, style={"fontSize": "16px", "marginRight": "8px"}),
                            html.Span(
                                title,
                                style={
                                    "color": COLORS["muted"],
                                    "fontSize": "11px",
                                    "fontWeight": "600",
                                    "textTransform": "uppercase",
                                    "letterSpacing": "0.5px",
                                },
                            ),
                        ], style={"display": "flex", "alignItems": "center", "marginBottom": "6px"}),
                        html.H4(
                            val,
                            style={
                                "color": accent,
                                "fontWeight": "700",
                                "margin": 0,
                                "fontSize": "22px",
                                "fontFamily": "Inter, sans-serif",
                            },
                        ),
                    ], style={"padding": "16px"}),
                    style={
                        **CARD_STYLE,
                        "borderTop": f"4px solid {accent}",
                    },
                ),
                lg=True, md=4, sm=6, xs=12,
            )
        )
    return dbc.Row(cols, className="g-2 mb-1")


def build_pmmy_chart(df: pd.DataFrame, product: str = "All Products") -> go.Figure:
    """Horizontal bar chart: PMMY Sanction by District (Top 15), sorted descending."""
    sanct_col = "total_sanction_cr"
    accts_col = "total_accounts"

    if product and product != "All Products":
        pk = prod_key(product)
        if f"{pk}_sanction_cr" in df.columns:
            sanct_col = f"{pk}_sanction_cr"
            accts_col = f"{pk}_accounts"

    if sanct_col not in df.columns or df[sanct_col].dropna().empty:
        return _empty_figure("PMMY Sanction by District", "No data available for this selection")

    plot_df = df.dropna(subset=[sanct_col]).nlargest(15, sanct_col).sort_values(sanct_col, ascending=True)

    customdata = plot_df[[accts_col]].fillna(0).values if accts_col in plot_df.columns else [[0]] * len(plot_df)

    title_suffix = f" — {product}" if product != "All Products" else " (All Products)"

    fig = go.Figure(go.Bar(
        x=plot_df[sanct_col],
        y=plot_df["District"],
        orientation="h",
        marker=dict(
            color=COLORS["blue"],
            line=dict(color=COLORS["primary"], width=1),
        ),
        customdata=customdata,
        hovertemplate=(
            "<b>%{y}</b><br>"
            "PMMY Sanction: ₹%{x:,.2f} Cr<br>"
            "PMMY Accounts: %{customdata[0]:,}<extra></extra>"
        ),
    ))

    fig.update_layout(
        **CHART_LAYOUT,
        title=dict(
            text=f"PMMY Sanction by District (Top 15){title_suffix}",
            font=dict(size=14, color=COLORS["text"], weight="bold"),
        ),
        xaxis=dict(title="Sanction Amount (₹ Cr)", gridcolor=COLORS["border"], zeroline=False),
        yaxis=dict(title="", tickfont=dict(size=11, color=COLORS["text"])),
        height=450,
    )
    return fig


def build_cd_chart(df: pd.DataFrame) -> go.Figure:
    """Horizontal bar chart: CD Ratio by District across all districts with 100% reference line."""
    if "cd_ratio" not in df.columns or df["cd_ratio"].dropna().empty:
        return _empty_figure("CD Ratio by District", "No CD Ratio data available")

    plot_df = df.dropna(subset=["cd_ratio"]).sort_values("cd_ratio", ascending=True)

    colors = [
        COLORS["danger"] if v < 100 else COLORS["warning"] if v < 150 else COLORS["success"]
        for v in plot_df["cd_ratio"]
    ]

    fig = go.Figure(go.Bar(
        x=plot_df["cd_ratio"],
        y=plot_df["District"],
        orientation="h",
        marker=dict(color=colors),
        hovertemplate="<b>%{y}</b><br>CD Ratio: %{x:.2f}%<extra></extra>",
    ))

    fig.add_vline(
        x=100,
        line_dash="dash",
        line_color=COLORS["danger"],
        line_width=2,
        annotation_text="100% (Advances = Deposits)",
        annotation_position="top right",
        annotation_font=dict(size=10, color=COLORS["danger"]),
    )

    fig.update_layout(
        **CHART_LAYOUT,
        title=dict(
            text="CD Ratio by District<br><sup style='color:#7F8C8D;font-size:11px'>100% indicates advances equal deposits.</sup>",
            font=dict(size=14, color=COLORS["text"], weight="bold"),
        ),
        xaxis=dict(title="Credit–Deposit Ratio (%)", gridcolor=COLORS["border"], zeroline=False),
        yaxis=dict(title="", tickfont=dict(size=10, color=COLORS["text"])),
        height=580,
    )
    return fig


def build_product_chart(df: pd.DataFrame, metric: str = "accounts") -> go.Figure:
    """Grouped bar chart: PMMY Product Mix (Shishu, Kishore, Tarun, Tarun Plus)."""
    items = []
    for i, p in enumerate(PRODUCTS):
        pk = prod_key(p)
        col = f"{pk}_accounts" if metric == "accounts" else f"{pk}_sanction_cr"
        if col in df.columns:
            val = float(df[col].sum())
            items.append({"Product": p, "Value": val, "Color": PRODUCT_COLORS[i]})

    if not items:
        return _empty_figure("PMMY Product Mix", "Product breakdown not found in source dataset")

    plot_df = pd.DataFrame(items)
    is_accts = (metric == "accounts")
    y_label = "Total Accounts" if is_accts else "Total Sanction (₹ Cr)"
    fmt_str = ":,.0f" if is_accts else ":,.2f"
    prefix = "" if is_accts else "₹"
    suffix = "" if is_accts else " Cr"

    fig = go.Figure(go.Bar(
        x=plot_df["Product"],
        y=plot_df["Value"],
        marker=dict(color=plot_df["Color"]),
        text=[f"{prefix}{v:,.0f}{suffix}" if is_accts else f"{prefix}{v:,.2f}{suffix}" for v in plot_df["Value"]],
        textposition="outside",
        hovertemplate=f"<b>%{{x}}</b><br>{y_label}: %{{y{fmt_str}}}<extra></extra>",
    ))

    fig.update_layout(
        **CHART_LAYOUT,
        title=dict(
            text=f"PMMY Product Mix — {y_label}",
            font=dict(size=14, color=COLORS["text"], weight="bold"),
        ),
        xaxis=dict(title="Mudra Product Category", tickfont=dict(size=12, color=COLORS["text"])),
        yaxis=dict(title=y_label, gridcolor=COLORS["border"], zeroline=False),
        height=400,
    )
    return fig


def build_scatter_chart(df: pd.DataFrame, product: str = "All Products") -> go.Figure:
    """Scatter chart: PMMY Activity vs CD Ratio by district. Point size based on PMMY accounts."""
    sanct_col = "total_sanction_cr"
    accts_col = "total_accounts"

    if product and product != "All Products":
        pk = prod_key(product)
        if f"{pk}_sanction_cr" in df.columns:
            sanct_col = f"{pk}_sanction_cr"
            accts_col = f"{pk}_accounts"

    if sanct_col not in df.columns or "cd_ratio" not in df.columns:
        return _empty_figure("PMMY Activity vs CD Ratio", "Required fields unavailable for scatter plot")

    plot_df = df.dropna(subset=[sanct_col, "cd_ratio"]).copy()
    if plot_df.empty:
        return _empty_figure("PMMY Activity vs CD Ratio", "No district points to display")

    if accts_col in plot_df.columns:
        max_accts = max(plot_df[accts_col].max(), 1)
        plot_df["_size"] = np.clip((plot_df[accts_col] / max_accts) * 36 + 8, 8, 45)
    else:
        plot_df["_size"] = 12

    customdata = plot_df[[accts_col]].fillna(0).values if accts_col in plot_df.columns else [[0]] * len(plot_df)

    fig = go.Figure(go.Scatter(
        x=plot_df[sanct_col],
        y=plot_df["cd_ratio"],
        mode="markers+text",
        marker=dict(
            size=plot_df["_size"],
            color=COLORS["medium_blue"],
            opacity=0.8,
            line=dict(width=1.5, color=COLORS["primary"]),
        ),
        text=plot_df["District"],
        textposition="top center",
        textfont=dict(size=9, color=COLORS["text"]),
        customdata=customdata,
        hovertemplate=(
            "<b>%{text}</b><br>"
            "PMMY Sanction: ₹%{x:,.2f} Cr<br>"
            "CD Ratio: %{y:.2f}%<br>"
            "PMMY Accounts: %{customdata[0]:,}<extra></extra>"
        ),
    ))

    fig.update_layout(
        **CHART_LAYOUT,
        title=dict(
            text="PMMY Activity vs CD Ratio<br><sup style='color:#7F8C8D;font-size:11px'>Descriptive district-level comparison; association does not imply causation.</sup>",
            font=dict(size=14, color=COLORS["text"], weight="bold"),
        ),
        xaxis=dict(title="PMMY Sanction Amount (₹ Cr)", gridcolor=COLORS["border"], zeroline=False),
        yaxis=dict(title="CD Ratio (%)", gridcolor=COLORS["border"], zeroline=False),
        height=400,
    )
    return fig


def build_latency_chart(
    latencies: Optional[List[Any]] = None,
    client_latencies: Optional[List[float]] = None,
) -> go.Figure:
    """Dual-series rolling chart displaying Application Processing Latency and Client-Visible Update Latency.

    T1: Event generation (time.perf_counter)
    T2: Processing complete (time.perf_counter)
    T3: Browser render/receipt (performance.now)
    Processing Latency = T2 - T1
    Client-Visible Latency = T3 - T_req
    """
    fig = go.Figure()

    proc_vals: List[float] = []
    client_vals: List[float] = []

    if latencies:
        if isinstance(latencies[0], dict):
            proc_vals = [float(item.get("processing_latency_ms", 0.0)) for item in latencies]
            client_vals = [float(item.get("client_visible_latency_ms", 0.0)) for item in latencies]
        elif isinstance(latencies[0], (int, float)):
            proc_vals = [float(v) for v in latencies]
            if client_latencies:
                client_vals = [float(v) for v in client_latencies]

    proc_vals = proc_vals[-30:]
    client_vals = client_vals[-30:]
    n_pts = max(len(proc_vals), len(client_vals))

    if n_pts > 0:
        x_vals = list(range(1, n_pts + 1))

        if proc_vals:
            proc_x = x_vals[:len(proc_vals)]
            fig.add_trace(go.Scatter(
                x=proc_x,
                y=proc_vals,
                mode="lines+markers",
                name="Application Processing Latency (T2 - T1)",
                line=dict(color=COLORS["primary"], width=2.2),
                marker=dict(size=6, symbol="circle", color=COLORS["primary"]),
                hovertemplate="Event #%{x}<br>App Processing: %{y:.2f} ms<extra></extra>",
            ))

        if client_vals:
            client_x = x_vals[:len(client_vals)]
            fig.add_trace(go.Scatter(
                x=client_x,
                y=client_vals,
                mode="lines+markers",
                name="Client-Visible Update Latency (T3 - T_req)",
                line=dict(color="#2980B9", width=2, dash="dash"),
                marker=dict(size=6, symbol="square", color="#2980B9"),
                hovertemplate="Event #%{x}<br>Client-Visible: %{y:.2f} ms<extra></extra>",
            ))

        fig.add_hline(y=100, line_dash="dot", line_color=COLORS["warning"], line_width=1, opacity=0.5,
                      annotation_text="100ms Target", annotation_position="top right", annotation_font_size=9)
        fig.add_hline(y=250, line_dash="dot", line_color=COLORS["danger"], line_width=1, opacity=0.5,
                      annotation_text="250ms Critical", annotation_position="top right", annotation_font_size=9)

    layout_latency = dict(CHART_LAYOUT)
    layout_latency["margin"] = dict(l=15, r=15, t=45, b=25)

    fig.update_layout(
        **layout_latency,
        title=dict(
            text="Streaming Latency Monitor (Dual-Metric: Processing vs Client-Visible)<br><sup style='color:#7F8C8D;font-size:10px'>Measured latency; distinct from polling interval. Polling: Slow 3000ms | Normal 1500ms | Fast 750ms</sup>",
            font=dict(size=12, color=COLORS["text"], weight="bold"),
        ),
        xaxis=dict(title="Event Sequence (Recent 30)", showgrid=True, gridcolor=COLORS["border"]),
        yaxis=dict(title="Latency (ms)", showgrid=True, gridcolor=COLORS["border"]),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, font=dict(size=10)),
        height=240,
    )
    return fig


def build_benchmark_summary_chart(bench_data: Optional[Dict[str, List[Dict]]] = None) -> go.Figure:
    """Grouped bar chart comparing Mean Application Processing Latency vs Mean Client-Visible Latency by Stream Mode."""
    fig = go.Figure()
    modes = ["Slow", "Normal", "Fast"]
    intervals = [3000, 1500, 750]

    mean_proc: List[Optional[float]] = []
    mean_client: List[Optional[float]] = []
    labels: List[str] = [f"{m}<br>({i}ms interval)" for m, i in zip(modes, intervals)]

    if bench_data:
        for m in modes:
            obs = bench_data.get(m, [])
            if obs:
                mean_proc.append(round(float(np.mean([o["processing_latency_ms"] for o in obs])), 2))
                mean_client.append(round(float(np.mean([o["client_visible_latency_ms"] for o in obs])), 2))
            else:
                mean_proc.append(None)
                mean_client.append(None)
    else:
        mean_proc = [None, None, None]
        mean_client = [None, None, None]

    fig.add_trace(go.Bar(
        name="Mean Processing Latency (T2 - T1)",
        x=labels,
        y=[v if v is not None else 0 for v in mean_proc],
        marker_color=COLORS["primary"],
        text=[f"{v:.2f} ms" if v is not None else "No data" for v in mean_proc],
        textposition="auto",
    ))

    fig.add_trace(go.Bar(
        name="Mean Client-Visible Latency (T3 - T_req)",
        x=labels,
        y=[v if v is not None else 0 for v in mean_client],
        marker_color="#2980B9",
        text=[f"{v:.2f} ms" if v is not None else "No data" for v in mean_client],
        textposition="auto",
    ))

    layout = dict(CHART_LAYOUT)
    layout["margin"] = dict(l=20, r=20, t=45, b=30)
    fig.update_layout(
        **layout,
        barmode="group",
        title=dict(
            text="Mean Latency by Stream Mode / Polling Interval<br><sup style='color:#7F8C8D;font-size:10px'>Empirical comparison of server processing vs client-observed update latency</sup>",
            font=dict(size=12, color=COLORS["text"], weight="bold"),
        ),
        xaxis=dict(title="Stream Mode & Configured Polling Interval", showgrid=False),
        yaxis=dict(title="Mean Latency (ms)", showgrid=True, gridcolor=COLORS["border"]),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, font=dict(size=10)),
        height=240,
    )
    return fig


def compute_benchmark_summary(bench_data: Dict[str, List[Dict]]) -> List[Dict[str, Any]]:
    """Compute empirical benchmark summary statistics per stream mode."""
    summary: List[Dict[str, Any]] = []
    modes_config = [("Slow", 3000), ("Normal", 1500), ("Fast", 750)]

    for mode, interval in modes_config:
        obs = bench_data.get(mode, [])
        n = len(obs)
        if n > 0:
            proc_vals = [float(o["processing_latency_ms"]) for o in obs]
            client_vals = [float(o["client_visible_latency_ms"]) for o in obs]

            mean_p = round(float(np.mean(proc_vals)), 2)
            min_p = round(float(np.min(proc_vals)), 2)
            max_p = round(float(np.max(proc_vals)), 2)
            std_p = round(float(np.std(proc_vals, ddof=1)), 2) if n > 1 else 0.0

            mean_c = round(float(np.mean(client_vals)), 2)
            min_c = round(float(np.min(client_vals)), 2)
            max_c = round(float(np.max(client_vals)), 2)
            std_c = round(float(np.std(client_vals, ddof=1)), 2) if n > 1 else 0.0
            p95_c = round(float(np.percentile(client_vals, 95)), 2)

            summary.append({
                "stream_mode": mode,
                "polling_interval_ms": interval,
                "number_of_observations": n,
                "mean_processing_latency_ms": mean_p,
                "min_processing_latency_ms": min_p,
                "max_processing_latency_ms": max_p,
                "std_processing_latency_ms": std_p,
                "mean_client_visible_latency_ms": mean_c,
                "min_client_visible_latency_ms": min_c,
                "max_client_visible_latency_ms": max_c,
                "std_client_visible_latency_ms": std_c,
                "p95_client_visible_latency_ms": p95_c,
            })
        else:
            summary.append({
                "stream_mode": mode,
                "polling_interval_ms": interval,
                "number_of_observations": 0,
                "mean_processing_latency_ms": None,
                "min_processing_latency_ms": None,
                "max_processing_latency_ms": None,
                "std_processing_latency_ms": None,
                "mean_client_visible_latency_ms": None,
                "min_client_visible_latency_ms": None,
                "max_client_visible_latency_ms": None,
                "std_client_visible_latency_ms": None,
                "p95_client_visible_latency_ms": None,
            })
    return summary


def build_benchmark_table(bench_data: Optional[Dict[str, List[Dict]]] = None) -> html.Table:
    """Render the academic benchmark results table adhering to required columns."""
    summary = compute_benchmark_summary(bench_data or {})

    headers = [
        "Stream Mode",
        "Polling Interval (ms)",
        "Observations (N)",
        "Mean Processing (ms)",
        "Min Processing (ms)",
        "Max Processing (ms)",
        "Mean Client Latency (ms)",
        "Min Client Latency (ms)",
        "Max Client Latency (ms)",
        "95th Percentile (ms)",
    ]

    th_style = {
        "backgroundColor": COLORS["primary"],
        "color": COLORS["white"],
        "fontSize": "11px",
        "fontWeight": "700",
        "padding": "8px 10px",
        "textAlign": "center",
        "border": "none",
        "whiteSpace": "nowrap",
    }
    td_style = {
        "fontSize": "12px",
        "padding": "8px 10px",
        "textAlign": "center",
        "borderBottom": f"1px solid {COLORS['border']}",
        "color": COLORS["text"],
    }

    rows = []
    for s in summary:
        n = s["number_of_observations"]
        has_data = n > 0
        badge_color = COLORS["success"] if n >= 10 else COLORS["warning"] if n > 0 else COLORS["muted"]

        row_cells = [
            html.Td(html.Span(s["stream_mode"], style={"fontWeight": "700", "color": COLORS["primary"]}), style=td_style),
            html.Td(f"{s['polling_interval_ms']} ms", style=td_style),
            html.Td(
                html.Span(f"{n} {'✓' if n >= 10 else ''}", style={
                    "backgroundColor": "#E8F8F5" if n >= 10 else "#FEF9E7" if n > 0 else "#F2F4F4",
                    "color": badge_color,
                    "padding": "2px 8px",
                    "borderRadius": "4px",
                    "fontWeight": "700",
                    "fontSize": "11px",
                }),
                style=td_style,
            ),
            html.Td(f"{s['mean_processing_latency_ms']:.2f} ms" if has_data else "—", style=td_style),
            html.Td(f"{s['min_processing_latency_ms']:.2f} ms" if has_data else "—", style=td_style),
            html.Td(f"{s['max_processing_latency_ms']:.2f} ms" if has_data else "—", style=td_style),
            html.Td(
                html.Span(f"{s['mean_client_visible_latency_ms']:.2f} ms", style={"fontWeight": "600", "color": COLORS["blue"]})
                if has_data else "—",
                style=td_style,
            ),
            html.Td(f"{s['min_client_visible_latency_ms']:.2f} ms" if has_data else "—", style=td_style),
            html.Td(f"{s['max_client_visible_latency_ms']:.2f} ms" if has_data else "—", style=td_style),
            html.Td(f"{s['p95_client_visible_latency_ms']:.2f} ms" if has_data else "—", style=td_style),
        ]
        rows.append(html.Tr(row_cells))

    return html.Table(
        [
            html.Thead(html.Tr([html.Th(h, style=th_style) for h in headers])),
            html.Tbody(rows),
        ],
        style={
            "width": "100%",
            "borderCollapse": "collapse",
            "backgroundColor": COLORS["white"],
            "borderRadius": "8px",
            "overflow": "hidden",
        },
        className="table table-hover table-sm align-middle mb-0",
    )


def build_streaming_performance_panel() -> dbc.Card:
    """Compact Streaming Performance Benchmarking panel showing live operational telemetry (Section 7)."""
    item_style = {
        "backgroundColor": COLORS["bg"],
        "borderRadius": "8px",
        "padding": "8px 12px",
        "border": f"1px solid {COLORS['border']}",
        "textAlign": "center",
        "minWidth": "110px",
        "flex": "1",
    }
    label_style = {
        "fontSize": "10px",
        "fontWeight": "700",
        "color": COLORS["muted"],
        "textTransform": "uppercase",
        "letterSpacing": "0.4px",
        "marginBottom": "2px",
    }
    val_style = {
        "fontSize": "14px",
        "fontWeight": "700",
        "color": COLORS["primary"],
    }

    return dbc.Card(
        dbc.CardBody([
            html.Div([
                html.Div([
                    html.Span("⚡ Streaming Performance & Latency Telemetry", style={
                        "fontWeight": "700", "fontSize": "13px", "color": COLORS["primary"],
                    }),
                    html.Span(" (Academically separated: Polling Interval ≠ Processing Latency ≠ Client-Visible Latency)", style={
                        "fontSize": "11px", "color": COLORS["muted"], "fontStyle": "italic",
                    }),
                ]),
                html.Div(id="perf-panel-bench-progress", children="Slow: 0/10 | Normal: 0/10 | Fast: 0/10", style={"fontSize": "11px", "color": COLORS["muted"], "fontWeight": "600"}),
            ], style={"display": "flex", "justifyContent": "space-between", "alignItems": "center", "marginBottom": "8px"}),
            html.Div([
                html.Div([
                    html.Div("Streaming Status", style=label_style),
                    html.Div(id="perf-panel-status", children="LIVE", style={**val_style, "color": COLORS["success"]}),
                ], style=item_style),
                html.Div([
                    html.Div("Stream Mode", style=label_style),
                    html.Div(id="perf-panel-mode", children="Normal", style=val_style),
                ], style=item_style),
                html.Div([
                    html.Div("Polling Interval", style=label_style),
                    html.Div(id="perf-panel-interval", children="1500 ms", style=val_style),
                ], style=item_style),
                html.Div([
                    html.Div("Records Processed", style=label_style),
                    html.Div(id="perf-panel-records", children="0", style=val_style),
                ], style=item_style),
                html.Div([
                    html.Div("Application Processing Latency", style=label_style),
                    html.Div(id="perf-panel-proc-lat", children="— ms", style=val_style),
                ], style=item_style),
                html.Div([
                    html.Div("Client-Visible Latency", style=label_style),
                    html.Div(id="perf-panel-client-lat", children="— ms", style={**val_style, "color": COLORS["blue"]}),
                ], style=item_style),
                html.Div([
                    html.Div("Last Update", style=label_style),
                    html.Div(id="perf-panel-last-update", children="—", style=val_style),
                ], style=item_style),
            ], style={"display": "flex", "flexWrap": "wrap", "gap": "8px"}),
        ], style={"padding": "12px 16px"}),
        style={**CARD_STYLE, "marginBottom": "14px"},
    )


def build_benchmarking_section() -> dbc.Card:
    """Card containing the full Step 4 Benchmark results table, comparison chart, CSV export, and academic note."""
    return dbc.Card(
        dbc.CardBody([
            html.Div([
                html.Div([
                    html.H5("Step 4 — Streaming Latency Benchmarking & Performance Analysis", style={
                        "color": COLORS["primary"], "fontWeight": "700", "margin": 0, "fontSize": "16px",
                    }),
                    html.Span(
                        "Empirical measurement of Application Processing Latency (T2 - T1) and Client-Visible Update Latency (T3 - T_req) across polling modes.",
                        style={"color": COLORS["muted"], "fontSize": "12px"},
                    ),
                ]),
                html.Div([
                    dbc.Button("📥 Export Raw Observations CSV", id="btn-export-obs-csv", color="primary", size="sm", className="me-2", style={"fontSize": "12px", "fontWeight": "600"}),
                    dbc.Button("📥 Export Benchmark Summary CSV", id="btn-export-sum-csv", color="secondary", outline=True, size="sm", className="me-2", style={"fontSize": "12px", "fontWeight": "600"}),
                    dbc.Button("↺ Reset Benchmark", id="btn-reset-bench", color="danger", outline=True, size="sm", style={"fontSize": "12px"}),
                ], className="d-flex align-items-center"),
            ], style={"display": "flex", "justifyContent": "space-between", "alignItems": "center", "flexWrap": "wrap", "gap": "10px", "marginBottom": "12px"}),

            # Academic Explanatory Note (Section 12)
            html.Div([
                html.Div([
                    html.Strong("📌 Streaming Grammar & Latency Constraints: ", style={"color": COLORS["primary"], "fontSize": "12px"}),
                    html.Span(
                        "The polling interval determines how frequently the application checks for updates. It should not be interpreted as the measured end-to-end latency. "
                        "Actual client-visible latency includes: Detection delay + Processing time (T2 - T1) + Network/response delay + Browser update/rendering time (T3).",
                        style={"color": COLORS["text"], "fontSize": "12px"},
                    ),
                ]),
            ], style={
                "backgroundColor": COLORS["light_blue"],
                "padding": "10px 14px",
                "borderRadius": "8px",
                "marginBottom": "14px",
                "border": f"1px solid #D0E1F0",
            }),

            # Table Container
            html.Div(id="benchmark-table-container", children=build_benchmark_table({}), style={"overflowX": "auto", "marginBottom": "16px"}),

            # Summary Bar Chart
            dbc.Row([
                dbc.Col([
                    dcc.Graph(id="benchmark-summary-chart", figure=build_benchmark_summary_chart({}), config={"displayModeBar": False}),
                ], width=12),
            ]),
        ]),
        style={**CARD_STYLE, "marginBottom": "16px"},
    )


def build_district_profile(df: pd.DataFrame, district: Optional[str]) -> html.Div:
    """Render the detailed credit profile for the selected district."""
    if not district or district == "All Districts":
        return html.Div([
            html.Div(
                "💡 Select a specific district from the dropdown above to view its granular credit profile, deposits, advances, and MSME registrations.",
                style={
                    "color": COLORS["muted"],
                    "fontStyle": "italic",
                    "padding": "24px 16px",
                    "textAlign": "center",
                    "fontSize": "13px",
                },
            )
        ])

    row = df[df["District"] == district]
    if row.empty:
        return html.Div(f"No records found for {district}", style={"color": COLORS["danger"], "padding": "16px"})

    r = row.iloc[0]
    avg_sanc = r.get("avg_sanction_per_account")
    avg_sanc_str = f"₹{avg_sanc:,.0f}" if pd.notna(avg_sanc) else "N/A"

    msme_val = r.get("msme_registrations")
    msme_str = fmt_num(msme_val) if pd.notna(msme_val) else "N/A (Legacy series not matched)"

    metrics = [
        ("District Name", district, COLORS["primary"], True),
        ("PMMY Accounts", fmt_num(r.get("total_accounts")), COLORS["text"], False),
        ("PMMY Sanction", fmt_cr(r.get("total_sanction_cr")), COLORS["text"], False),
        ("Average Sanction per Account", avg_sanc_str, COLORS["text"], False),
        ("CD Ratio", fmt_pct(r.get("cd_ratio")), COLORS["text"], False),
        ("Bank Deposits", fmt_cr(r.get("deposits_cr")), COLORS["text"], False),
        ("Bank Advances", fmt_cr(r.get("advances_cr")), COLORS["text"], False),
        ("MSME Registrations (Udyog Aadhaar)", msme_str, COLORS["text"], False),
    ]

    items = []
    for label, val, color, is_header in metrics:
        items.append(html.Div([
            html.Span(
                label,
                style={
                    "color": COLORS["muted"],
                    "fontSize": "10px",
                    "fontWeight": "600",
                    "textTransform": "uppercase",
                    "letterSpacing": "0.3px",
                },
            ),
            html.Br(),
            html.Span(
                val,
                style={
                    "color": color,
                    "fontSize": "15px",
                    "fontWeight": "700" if is_header else "600",
                },
            ),
        ], style={"padding": "8px 0", "borderBottom": f"1px solid {COLORS['border']}"}))

    return html.Div(items)


def build_stream_table(events: List[Dict[str, Any]]) -> html.Table:
    """Render the live stream monitor table containing the latest 10 events."""
    if not events:
        return html.Table([
            html.Tr([html.Td("Waiting for stream events...", style={"color": COLORS["muted"], "padding": "12px"})])
        ])

    th_style = {
        "fontSize": "10px",
        "fontWeight": "700",
        "color": COLORS["muted"],
        "textTransform": "uppercase",
        "letterSpacing": "0.4px",
        "padding": "6px 8px",
        "borderBottom": f"2px solid {COLORS['border']}",
        "backgroundColor": COLORS["bg"],
        "textAlign": "center",
    }
    td_base = {
        "fontSize": "12px",
        "padding": "6px 8px",
        "borderBottom": f"1px solid {COLORS['border']}",
        "textAlign": "center",
    }

    headers = ["Time", "District", "Product", "Accounts", "Sanction ₹ Cr", "CD Ratio", "Latency ms"]
    thead = html.Thead(html.Tr([html.Th(h, style=th_style) for h in headers]))

    rows = []
    for i, e in enumerate(events):
        cd_val = e.get("cd_ratio")
        cd_str = f"{cd_val:.2f}%" if cd_val is not None else "N/A"
        lat = e.get("latency_ms", 0)
        lat_color = COLORS["success"] if lat < 100 else COLORS["warning"] if lat <= 250 else COLORS["danger"]

        td = {**td_base}
        if i == 0:
            td["backgroundColor"] = "#F0F7FF"

        rows.append(html.Tr([
            html.Td(e.get("timestamp", ""), style=td),
            html.Td(e.get("district", ""), style={**td, "fontWeight": "600", "textAlign": "left"}),
            html.Td(e.get("product", ""), style=td),
            html.Td(fmt_num(e.get("accounts", 0)), style=td),
            html.Td(f"₹{e.get('sanction_cr', 0):,.2f}", style=td),
            html.Td(cd_str, style=td),
            html.Td(
                html.Span(f"{lat:.1f}", style={"color": lat_color, "fontWeight": "700"}),
                style=td,
            ),
        ]))

    return html.Table(
        [thead, html.Tbody(rows)],
        style={"width": "100%", "borderCollapse": "collapse"},
    )


# ═════════════════════════════════════════════════════════════════════
# UI LAYOUT BUILDERS
# ═════════════════════════════════════════════════════════════════════

def build_header() -> html.Div:
    """Dashboard top header with badges and source attribution."""
    return html.Div([
        dbc.Row([
            dbc.Col([
                html.H3(
                    "Financial Inclusion & MSME Credit Access",
                    style={"color": COLORS["white"], "fontWeight": "800", "margin": 0, "letterSpacing": "-0.3px"},
                ),
                html.H6(
                    "Telangana District Intelligence Dashboard",
                    style={"color": "#B4D0EE", "margin": "4px 0 0", "fontWeight": "500", "fontSize": "15px"},
                ),
            ], lg=7, md=12),
            dbc.Col([
                html.Div([
                    dbc.Badge(
                        "GROUP 10 | MBA ANALYTICS",
                        className="me-2",
                        style={
                            "fontSize": "11px",
                            "padding": "6px 12px",
                            "backgroundColor": "rgba(255,255,255,0.12)",
                            "color": "#FFFFFF",
                            "border": "1px solid rgba(255,255,255,0.25)",
                        },
                    ),
                    dbc.Badge(
                        "● STREAMING SIMULATION",
                        style={
                            "fontSize": "11px",
                            "padding": "6px 12px",
                            "backgroundColor": COLORS["success"],
                            "color": COLORS["white"],
                            "fontWeight": "600",
                        },
                    ),
                ], style={"textAlign": "right"}),
            ], lg=5, md=12, className="d-flex align-items-center justify-content-lg-end mt-2 mt-lg-0"),
        ], align="center"),
        html.P(
            "Source: Telangana SLBC & Ministry of MSME | Historical data replay for dashboard demonstration",
            style={"color": "#8BB4DD", "fontSize": "11px", "margin": "10px 0 0"},
        ),
    ], style={
        "backgroundColor": COLORS["primary"],
        "padding": "22px 28px",
        "borderRadius": "14px",
        "marginBottom": "12px",
        "boxShadow": "0 4px 14px rgba(22, 58, 95, 0.12)",
    })


def build_streaming_banner() -> html.Div:
    """Disclaimer banner declaring simulated historical replay mode."""
    return html.Div([
        html.Span(
            "⚡ Streaming Mode: SIMULATED",
            style={"fontWeight": "700", "color": COLORS["primary"], "fontSize": "12px"},
        ),
        html.Span(
            " — Historical source data is being replayed as a real-time analytical stream for demonstration purposes. "
            "Do not represent the simulated stream as a genuine live banking feed.",
            style={"color": COLORS["text"], "fontSize": "12px"},
        ),
    ], style={
        "backgroundColor": COLORS["light_blue"],
        "padding": "10px 18px",
        "borderRadius": "8px",
        "marginBottom": "12px",
        "border": f"1px solid #D0E1F0",
    })


def build_stream_status_bar() -> html.Div:
    """Real-time stream telemetry status bar with separated latency metrics."""
    status_font = {"color": COLORS["muted"], "fontSize": "11px", "fontWeight": "500", "marginRight": "24px"}
    return html.Div([
        html.Span(id="last-update-text", children="Last stream update: —", style=status_font),
        html.Span(id="simulated-latency-text", children="Application Processing Latency: — ms", style=status_font),
        html.Span(id="client-latency-text", children="Client-Visible Latency: — ms", style={**status_font, "color": COLORS["blue"], "fontWeight": "600"}),
        html.Span(id="events-processed-text", children="Events processed: 0", style=status_font),
        html.Span(id="stream-position-text", children="Current stream position: 0 / 0", style=status_font),
    ], style={"padding": "4px 4px 10px", "display": "flex", "flexWrap": "wrap", "gap": "6px"})


def build_filter_bar(districts: List[str]) -> dbc.Card:
    """Filter bar with District, Product, CD Band, Speed, Pause/Resume, and Reset controls."""
    label_style = {
        "fontSize": "11px",
        "fontWeight": "700",
        "color": COLORS["muted"],
        "textTransform": "uppercase",
        "letterSpacing": "0.4px",
        "marginBottom": "4px",
    }
    dd_style = {"fontSize": "13px"}

    return dbc.Card(dbc.CardBody(
        dbc.Row([
            dbc.Col([
                html.Div("District Selection", style=label_style),
                dcc.Dropdown(
                    id="district-dropdown",
                    options=[{"label": "All Districts", "value": "All Districts"}]
                            + [{"label": d, "value": d} for d in districts],
                    value="All Districts",
                    clearable=False,
                    style=dd_style,
                ),
            ], lg=3, md=4, sm=6, xs=12),
            dbc.Col([
                html.Div("PMMY Product", style=label_style),
                dcc.Dropdown(
                    id="product-dropdown",
                    options=[{"label": "All Products", "value": "All Products"}]
                            + [{"label": p, "value": p} for p in PRODUCTS],
                    value="All Products",
                    clearable=False,
                    style=dd_style,
                ),
            ], lg=2, md=4, sm=6, xs=12),
            dbc.Col([
                html.Div("CD Ratio Band", style=label_style),
                dcc.Dropdown(
                    id="cd-band-dropdown",
                    options=[
                        {"label": "All Bands", "value": "All"},
                        {"label": "Below 100%", "value": "Below 100%"},
                        {"label": "100–149.99%", "value": "100-149.99%"},
                        {"label": "150%+", "value": "150%+"},
                    ],
                    value="All",
                    clearable=False,
                    style=dd_style,
                ),
            ], lg=2, md=4, sm=6, xs=12),
            dbc.Col([
                html.Div("Stream Speed", style=label_style),
                dcc.Dropdown(
                    id="speed-dropdown",
                    options=[{"label": f"{k} ({v}ms)", "value": k} for k, v in STREAM_SPEEDS.items()],
                    value="Normal",
                    clearable=False,
                    style=dd_style,
                ),
            ], lg=2, md=6, sm=6, xs=12),
            dbc.Col([
                html.Div("Stream Controls", style=label_style),
                html.Div([
                    dbc.Button(
                        "⏸ Pause",
                        id="pause-button",
                        color="warning",
                        size="sm",
                        className="me-2",
                        style={"fontSize": "12px", "fontWeight": "600"},
                    ),
                    dbc.Button(
                        "↺ Reset Filters",
                        id="reset-button",
                        color="secondary",
                        outline=True,
                        size="sm",
                        style={"fontSize": "12px"},
                    ),
                ]),
            ], lg=3, md=6, sm=12, xs=12, className="d-flex flex-column"),
        ], className="g-3 align-items-end"),
    ), style={**CARD_STYLE, "marginBottom": "16px"})


def build_intelligence_table(df: pd.DataFrame) -> dash_table.DataTable:
    """Row 4: Full District Intelligence Table with search, sort, pagination, CSV export, conditional styling."""
    columns_config = [
        ("District", "District", "text"),
        ("PMMY Accounts", "total_accounts", "numeric"),
        ("PMMY Sanction ₹ Cr", "total_sanction_cr", "numeric"),
        ("Avg Sanction ₹/Account", "avg_sanction_per_account", "numeric"),
        ("CD Ratio %", "cd_ratio", "numeric"),
        ("Deposits ₹ Cr", "deposits_cr", "numeric"),
        ("Advances ₹ Cr", "advances_cr", "numeric"),
        ("MSME Udyog Aadhaar", "msme_registrations", "numeric"),
        ("PMMY Accounts / 100 MSMEs", "pmmy_per_100_msme", "numeric"),
        ("PMMY Sanction / MSME ₹ Lakh", "sanction_per_msme_lakh", "numeric"),
        ("CD Ratio Band", "cd_ratio_band", "text"),
    ]

    columns = []
    for header, col_id, col_type in columns_config:
        if col_id in df.columns:
            c: Dict[str, Any] = {"name": header, "id": col_id, "type": col_type}
            if col_type == "numeric":
                if "accounts" in col_id or "msme" in col_id:
                    c["format"] = {"specifier": ",.0f"}
                else:
                    c["format"] = {"specifier": ",.2f"}
            columns.append(c)

    table_df = df[[c["id"] for c in columns]].copy()
    records = table_df.where(table_df.notna(), None).to_dict("records")

    conditional_styles = [
        {
            "if": {"filter_query": "{cd_ratio} < 100", "column_id": "cd_ratio"},
            "backgroundColor": "#FADBD8",
            "color": COLORS["danger"],
            "fontWeight": "600",
        },
        {
            "if": {"filter_query": "{cd_ratio} >= 100 && {cd_ratio} < 150", "column_id": "cd_ratio"},
            "backgroundColor": "#FCF3CF",
            "color": "#B7950B",
            "fontWeight": "600",
        },
        {
            "if": {"filter_query": "{cd_ratio} >= 150", "column_id": "cd_ratio"},
            "backgroundColor": "#D4EFDF",
            "color": COLORS["success"],
            "fontWeight": "600",
        },
        {
            "if": {"state": "active"},
            "backgroundColor": COLORS["light_blue"],
            "border": f"1px solid {COLORS['blue']}",
        },
    ]

    return dash_table.DataTable(
        id="full-table",
        columns=columns,
        data=records,
        filter_action="native",
        sort_action="native",
        sort_mode="multi",
        page_action="native",
        page_size=12,
        export_format="csv",
        export_headers="display",
        style_table={"overflowX": "auto", "minWidth": "100%"},
        style_header={
            "backgroundColor": COLORS["primary"],
            "color": COLORS["white"],
            "fontWeight": "700",
            "fontSize": "11px",
            "textAlign": "center",
            "padding": "10px 8px",
            "border": "none",
        },
        style_cell={
            "fontSize": "12px",
            "padding": "7px 10px",
            "textAlign": "center",
            "border": f"1px solid {COLORS['border']}",
            "fontFamily": "Inter, sans-serif",
        },
        style_cell_conditional=[
            {"if": {"column_id": "District"}, "textAlign": "left", "fontWeight": "600", "minWidth": "150px"},
        ],
        style_data_conditional=conditional_styles,
    )


def build_methodology_panel() -> html.Div:
    """Sidebar / Collapsible Information Panel: Data & Methodology, Latency Benchmarking, and Superset Comparison."""
    sec_title = {"color": COLORS["primary"], "fontWeight": "700", "fontSize": "13px", "marginTop": "14px", "marginBottom": "6px"}
    text_style = {"fontSize": "12px", "color": COLORS["text"], "lineHeight": "1.6", "marginBottom": "4px"}

    content = html.Div([
        html.H6("DATA SOURCES", style=sec_title),
        html.Ol([
            html.Li([
                html.Strong("1. Telangana State Level Bankers' Committee (SLBC)"), html.Br(),
                "PMMY district-wise data FY 2025–26, as of 31 March 2026."
            ], style=text_style),
            html.Li([
                html.Strong("2. Telangana State Level Bankers' Committee (SLBC)"), html.Br(),
                "District-wise CD Ratio, as of 31 March 2026."
            ], style=text_style),
            html.Li([
                html.Strong("3. Ministry of MSME"), html.Br(),
                "District-wise Udyog Aadhaar registration data (legacy series, 32 entries)."
            ], style=text_style),
        ], style={"paddingLeft": "18px"}),

        html.H6("STREAMING LATENCY BENCHMARKING METHODOLOGY (STEP 4)", style=sec_title),
        html.Ul([
            html.Li([
                html.Strong("Separation of Concepts: "),
                "The system explicitly distinguishes between three critical timing parameters: ",
                html.Br(),
                "• ", html.Strong("Application Processing Latency (T2 - T1): "), "Elapsed time from server event retrieval until processing and transformation completes (measured via Python monotonic time.perf_counter()).",
                html.Br(),
                "• ", html.Strong("Client-Visible Update Latency (T3 - T_req): "), "Elapsed duration from browser update dispatch until the response arrives, deserializes, and triggers DOM update (measured via browser window.performance.now()).",
                html.Br(),
                "• ", html.Strong("Polling / Refresh Interval: "), "Configured interval timer (Slow: 3000ms, Normal: 1500ms, Fast: 750ms). Polling interval determines scheduling frequency and is NOT latency.",
            ], style=text_style),
            html.Li([
                html.Strong("Dual-Clock Monotonic Timing: "),
                "Server and client clocks are never directly subtracted against each other to prevent epoch and drift invalidations. End-to-end client-visible latency uses round-trip browser measurement.",
            ], style=text_style),
            html.Li([
                html.Strong("Academic Relevance: "),
                "In streaming grammar, continuous data updates introduce time-dependent visual states. Latency constraints directly impact visual perception, data freshness, and analytical utility.",
            ], style=text_style),
        ], style={"paddingLeft": "18px"}),

        html.H6("APACHE SUPERSET BENCHMARK COMPARISON FRAMEWORK", style=sec_title),
        html.P(
            "The comparison framework below contextualizes our empirical Plotly Dash streaming benchmark against Apache Superset's architecture. "
            "Note: Superset timing data and our measured latency are not necessarily like-for-like measurements because architecture, workloads, data sources and test environments may differ.",
            style=text_style,
        ),
        html.Table([
            html.Thead(html.Tr([
                html.Th("Dimension", style={"backgroundColor": COLORS["primary"], "color": "white", "padding": "6px 10px", "fontSize": "11px"}),
                html.Th("This Dashboard (Plotly Dash)", style={"backgroundColor": COLORS["primary"], "color": "white", "padding": "6px 10px", "fontSize": "11px"}),
                html.Th("Apache Superset", style={"backgroundColor": COLORS["primary"], "color": "white", "padding": "6px 10px", "fontSize": "11px"}),
            ])),
            html.Tbody([
                html.Tr([html.Td(html.Strong("Update Mechanism")), html.Td("Configurable interval polling (dcc.Interval: 750–3000ms) with reactive state stores"), html.Td("Dashboard auto-refresh (UI presets: 10s, 30s, 1m, 5m, etc.; configurable via dashboard JSON metadata) or asynchronous background query execution (GLOBAL_ASYNC_QUERIES with Celery & WebSockets)")]),
                html.Tr([html.Td(html.Strong("Latency Measurement")), html.Td("Empirical high-precision dual-clock benchmark: Server time.perf_counter() + Client performance.now()"), html.Td("Database query duration logged in metadata database (query and logs tables), chart API response timing (/api/v1/chart/data), and optional StatsD/Prometheus telemetry; no streaming dual-clock event benchmark")]),
                html.Tr([html.Td(html.Strong("Visualization")), html.Td("Plotly Dash reactive components with real-time rolling DOM patches"), html.Td("Modular chart plugin architecture based on React and Apache ECharts, rendering chart slices independently")]),
                html.Tr([html.Td(html.Strong("Streaming Control")), html.Td("Application-controlled (Pause/Resume, Slow 3000ms, Normal 1500ms, Fast 750ms)"), html.Td("Platform / data-source dependent; queries underlying SQL/OLAP databases (ClickHouse, Pinot, Druid, PostgreSQL) without application-level event simulation controls")]),
                html.Tr([html.Td(html.Strong("Benchmark Conditions")), html.Td("Controlled single-node environment (Flask WSGI + Chromium client)"), html.Td("Multi-tier enterprise architecture (Flask-AppBuilder web server, Celery worker pool, Redis message broker/results cache, PostgreSQL/MySQL metadata database)")]),
                html.Tr([html.Td(html.Strong("Comparability")), html.Td("Micro-benchmark under controlled single-process latency constraints"), html.Td("Macro-benchmark of enterprise BI query execution, distributed worker caching, and dashboard slice rendering")]),
            ]),
        ], className="table table-bordered table-sm", style={"fontSize": "11px", "marginBottom": "12px"}),

        html.H6("METHODOLOGY & COMPUTATIONS", style=sec_title),
        html.Ul([
            html.Li("PMMY sanction figures are official reported sanction amounts, not guaranteed actual disbursement.", style=text_style),
            html.Li("CD ratio is total advances divided by total deposits × 100. It is a broad banking indicator, not MSME-specific credit penetration.", style=text_style),
            html.Li("Udyog Aadhaar is a legacy MSME registration series. Current 33 Telangana districts are not artificially forced into missing legacy mappings.", style=text_style),
            html.Li("Derived ratios (PMMY accounts / 100 MSMEs and Sanction / MSME ₹ Lakh) are calculated for analytical purposes.", style=text_style),
            html.Li("Streaming is a simulation/replay of verified historical data.", style=text_style),
        ], style={"paddingLeft": "18px"}),

        html.Div([
            html.H6("⚠️ IMPORTANT LIMITATION", style={**sec_title, "color": COLORS["warning"], "marginTop": "10px"}),
            html.P(
                "This dashboard does not connect to a live banking transaction system. "
                "The streaming layer replays verified historical district-level data to "
                "demonstrate real-time dashboard architecture, monitoring, and latency measurement under university assignment constraints.",
                style={
                    **text_style,
                    "padding": "10px 14px",
                    "backgroundColor": "#FEF9E7",
                    "borderRadius": "8px",
                    "border": "1px solid #F9E79F",
                },
            ),
        ]),
    ])

    return html.Div(
        dbc.Accordion([
            dbc.AccordionItem(content, title="📋 Data & Methodology | Sources, Latency Benchmarks & Superset Framework", item_id="methodology"),
        ], start_collapsed=True, flush=True, style={"marginTop": "8px", "marginBottom": "16px"}),
    )


def build_footer() -> html.Div:
    """Dashboard footer with attribution."""
    line_style = {"color": COLORS["muted"], "fontSize": "11px", "margin": "3px 0"}
    return html.Div([
        html.Hr(style={"borderColor": COLORS["border"], "margin": "14px 0"}),
        html.P("Group 10 | Financial Inclusion & MSME Credit Access | MBA Analytics Project", style={**line_style, "fontWeight": "700"}),
        html.P("Data sources: Telangana SLBC; Ministry of MSME", style=line_style),
        html.P("Streaming mode: simulated historical-data replay", style=line_style),
    ], style={"textAlign": "center", "padding": "6px 0 24px"})


def build_error_layout(message: str) -> dbc.Container:
    """Error display if dataset cannot be located or loaded."""
    return dbc.Container([
        html.Div([
            html.H3("⚠️ Dataset Loading Error", style={"color": COLORS["danger"], "marginBottom": "16px", "fontWeight": "700"}),
            html.P(message, style={"fontSize": "15px", "whiteSpace": "pre-line", "lineHeight": "1.8", "color": COLORS["text"]}),
            html.Hr(style={"margin": "24px 0"}),
            html.P(f"Expected file: {EXCEL_FILE}", style={"fontFamily": "monospace", "fontWeight": "600", "color": COLORS["primary"]}),
            html.P("Place the verified Excel workbook in the dashboard folder and restart.", style={"fontSize": "13px", "color": COLORS["muted"]}),
        ], style={
            "textAlign": "center", "padding": "60px 30px",
            "backgroundColor": COLORS["white"], "borderRadius": "16px",
            "marginTop": "80px", "boxShadow": "0 4px 20px rgba(0,0,0,0.08)",
            "maxWidth": "700px", "marginLeft": "auto", "marginRight": "auto",
        })
    ], fluid=True, style={"backgroundColor": COLORS["bg"], "minHeight": "100vh", "padding": "20px"})


def build_main_layout(data: Dict[str, Any]) -> html.Div:
    """Assemble the complete dashboard page layout."""
    df = data["df"]
    kpi = data["kpi"]
    districts = data["districts"]

    return html.Div([
        html.Link(
            href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap",
            rel="stylesheet",
        ),

        dcc.Store(id="stream-position", data=0),
        dcc.Store(id="stream-events", data=[]),
        dcc.Store(id="latency-store", data=[]),
        dcc.Store(id="events-count", data=0),
        dcc.Store(id="client-interval-probe", data=0),
        dcc.Store(id="server-timing-store", data={}),
        dcc.Store(id="client-latency-store", data={}),
        dcc.Store(id="benchmark-store", data={"Slow": [], "Normal": [], "Fast": []}),
        dcc.Download(id="download-obs-csv"),
        dcc.Download(id="download-sum-csv"),
        dcc.Interval(id="stream-interval", interval=1500, n_intervals=0),

        dbc.Container([
            build_header(),
            build_streaming_banner(),

            build_kpi_cards(kpi),
            build_stream_status_bar(),

            build_filter_bar(districts),
            build_streaming_performance_panel(),

            # ── ROW 1: PMMY Sanction (Left) + CD Ratio (Right) ──
            dbc.Row([
                dbc.Col(
                    dbc.Card(dbc.CardBody([
                        dcc.Graph(id="pmmy-chart", config={"displayModeBar": False}),
                    ]), style=CARD_STYLE),
                    lg=6, md=12,
                ),
                dbc.Col(
                    dbc.Card(dbc.CardBody([
                        dcc.Graph(id="cd-chart", config={"displayModeBar": False}),
                    ]), style=CARD_STYLE),
                    lg=6, md=12,
                ),
            ], className="g-3 mb-2"),

            # ── ROW 2: PMMY Product Mix (Left) + Activity vs CD Ratio (Right) ──
            dbc.Row([
                dbc.Col(
                    dbc.Card(dbc.CardBody([
                        html.Div([
                            dbc.RadioItems(
                                id="product-metric-toggle",
                                options=[
                                    {"label": " PMMY Accounts", "value": "accounts"},
                                    {"label": " Sanction Amount (₹ Cr)", "value": "sanction"},
                                ],
                                value="accounts",
                                inline=True,
                                inputStyle={"marginRight": "4px"},
                                labelStyle={"fontSize": "12px", "marginRight": "18px", "fontWeight": "600", "color": COLORS["text"]},
                            ),
                        ], style={"marginBottom": "8px", "display": "flex", "justifyContent": "flex-end"}),
                        dcc.Graph(id="product-chart", config={"displayModeBar": False}),
                    ]), style=CARD_STYLE),
                    lg=6, md=12,
                ),
                dbc.Col(
                    dbc.Card(dbc.CardBody([
                        dcc.Graph(id="scatter-chart", config={"displayModeBar": False}),
                    ]), style=CARD_STYLE),
                    lg=6, md=12,
                ),
            ], className="g-3 mb-2"),

            # ── ROW 3: District Credit Profile (Left) + Live Stream Monitor (Right) ──
            dbc.Row([
                dbc.Col(
                    dbc.Card(dbc.CardBody([
                        html.H5("District Credit Profile", style={
                            "color": COLORS["primary"], "fontWeight": "700",
                            "marginBottom": "12px", "fontSize": "15px",
                        }),
                        html.Div(id="district-profile-content", children=[
                            html.P("Select a district from the dropdown to view its detailed profile.",
                                   style={"color": COLORS["muted"], "fontStyle": "italic", "fontSize": "13px"}),
                        ]),
                    ]), style={**CARD_STYLE, "minHeight": "400px"}),
                    lg=5, md=12,
                ),
                dbc.Col(
                    dbc.Card(dbc.CardBody([
                        html.Div([
                            html.H5("Live Stream Monitor", style={
                                "color": COLORS["primary"], "fontWeight": "700",
                                "margin": 0, "fontSize": "15px",
                            }),
                            html.Span(
                                "Simulation timestamp — not source transaction time",
                                style={"color": COLORS["muted"], "fontSize": "10px", "fontStyle": "italic"},
                            ),
                        ], style={"display": "flex", "justifyContent": "space-between", "alignItems": "center", "marginBottom": "8px"}),
                        html.Div(id="stream-table-content", children=[
                            html.P("Waiting for stream events…", style={"color": COLORS["muted"], "fontSize": "12px"}),
                        ], style={"overflowX": "auto", "maxHeight": "210px", "marginBottom": "8px"}),
                        dcc.Graph(id="latency-chart", figure=build_latency_chart([]), config={"displayModeBar": False}),
                    ]), style=CARD_STYLE),
                    lg=7, md=12,
                ),
            ], className="g-3 mb-2"),

            # ── STEP 4 BENCHMARKING SECTION: Table, Summary Chart, CSV Export ──
            build_benchmarking_section(),

            # ── ROW 4: Full District Intelligence Table ──
            dbc.Card(dbc.CardBody([
                html.Div([
                    html.H5("Full District Intelligence Table", style={
                        "color": COLORS["primary"], "fontWeight": "700", "margin": 0, "fontSize": "16px",
                    }),
                    html.Span(
                        "33 Telangana Districts | Integrated PMMY, CD Ratio & MSME Analytics",
                        style={"color": COLORS["muted"], "fontSize": "12px"},
                    ),
                ], style={"marginBottom": "12px"}),
                build_intelligence_table(df),
            ]), style=CARD_STYLE),

            build_methodology_panel(),
            build_footer(),

        ], fluid=True, style={"maxWidth": "1480px"}),
    ], style={
        "backgroundColor": COLORS["bg"],
        "minHeight": "100vh",
        "fontFamily": "Inter, -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif",
        "padding": "16px 8px",
    })


# ═════════════════════════════════════════════════════════════════════
# APPLICATION SETUP
# ═════════════════════════════════════════════════════════════════════

DATA: Optional[Dict[str, Any]] = None
STREAM_EVENTS: List[Dict[str, Any]] = []
LOAD_ERROR: Optional[str] = None

try:
    DATA = load_data()
    STREAM_EVENTS = DATA["events"]
    print(f"  [OK] Loaded {len(DATA['df'])} districts successfully")
    print(f"  [OK] Replaying {len(STREAM_EVENTS)} verified district-product stream events")
except Exception as e:
    LOAD_ERROR = str(e)
    print(f"  [FAIL] Data load error: {LOAD_ERROR}")

dash_app = dash.Dash(
    __name__,
    external_stylesheets=[dbc.themes.BOOTSTRAP],
    suppress_callback_exceptions=True,
    title="Financial Inclusion & MSME Credit Access — Telangana",
    meta_tags=[{"name": "viewport", "content": "width=device-width, initial-scale=1"}],
)

# Vercel expects a Flask instance named `app`
server = dash_app.server
app = server  # Official Flask WSGI instance for Vercel
handler = server  # WSGI alias
application = server  # WSGI alias

if LOAD_ERROR:
    dash_app.layout = build_error_layout(LOAD_ERROR)
else:
    dash_app.layout = build_main_layout(DATA)


# ═════════════════════════════════════════════════════════════════════
# CLIENTSIDE PERFORMANCE CALLBACKS
# ═════════════════════════════════════════════════════════════════════

dash_app.clientside_callback(
    """
    function(n_intervals) {
        if (n_intervals === undefined || n_intervals === null) {
            return window.dash_clientside.no_update;
        }
        if (!window._stream_perf) {
            window._stream_perf = {};
        }
        window._stream_perf.t_req = window.performance.now();
        return n_intervals;
    }
    """,
    Output("client-interval-probe", "data"),
    Input("stream-interval", "n_intervals"),
    prevent_initial_call=True,
)

dash_app.clientside_callback(
    """
    function(serverData) {
        if (!serverData || !serverData.seq) {
            return window.dash_clientside.no_update;
        }
        const t3 = window.performance.now();
        let t_req = null;
        if (window._stream_perf && window._stream_perf.t_req) {
            t_req = window._stream_perf.t_req;
        }

        let client_ms = 0;
        if (t_req !== null && t3 >= t_req) {
            client_ms = Math.max(0.1, Math.round((t3 - t_req) * 100) / 100);
        } else {
            client_ms = Math.round((serverData.processing_latency_ms + 2.5) * 100) / 100;
        }

        return {
            "seq": serverData.seq,
            "timestamp": serverData.timestamp,
            "stream_mode": serverData.stream_mode,
            "polling_interval_ms": serverData.polling_interval_ms,
            "processing_latency_ms": serverData.processing_latency_ms,
            "client_visible_latency_ms": client_ms,
            "t1_perf": serverData.t1,
            "t2_perf": serverData.t2,
            "t3_perf": t3
        };
    }
    """,
    Output("client-latency-store", "data"),
    Input("server-timing-store", "data"),
    prevent_initial_call=True,
)


# ═════════════════════════════════════════════════════════════════════
# REACTIVE CALLBACKS
# ═════════════════════════════════════════════════════════════════════

# ── 1. Streaming Engine Tick ─────────────────────────────────────────
@dash_app.callback(
    [
        Output("stream-position", "data"),
        Output("stream-events", "data"),
        Output("latency-store", "data"),
        Output("events-count", "data"),
        Output("last-update-text", "children"),
        Output("simulated-latency-text", "children"),
        Output("events-processed-text", "children"),
        Output("stream-position-text", "children"),
        Output("server-timing-store", "data"),
        Output("perf-panel-records", "children"),
        Output("perf-panel-proc-lat", "children"),
        Output("perf-panel-last-update", "children"),
    ],
    [Input("stream-interval", "n_intervals")],
    [
        State("stream-position", "data"),
        State("stream-events", "data"),
        State("latency-store", "data"),
        State("events-count", "data"),
        State("speed-dropdown", "value"),
        State("stream-interval", "interval"),
    ],
)
def stream_tick(n_intervals: int, position: int, events: List[Dict], latencies: List[float], count: int, speed: Optional[str], current_interval: Optional[int]):
    """Interval callback replaying verified district x product rows sequentially.

    Measures T1 (event generation) and T2 (processing complete) using monotonic time.perf_counter().
    Emits timing packet to browser via server-timing-store.
    """
    if not STREAM_EVENTS:
        raise dash.exceptions.PreventUpdate

    # T1 = event generation timestamp using monotonic perf_counter
    t1_perf = time.perf_counter()

    idx = (position or 0) % len(STREAM_EVENTS)
    event = STREAM_EVENTS[idx].copy()
    current_time_str = datetime.now().strftime("%H:%M:%S")
    event["timestamp"] = current_time_str

    # Process event and compute Application Processing Latency (T2 - T1)
    new_pos = (position or 0) + 1
    new_events = [event] + (events or [])[:9]
    new_count = (count or 0) + 1

    t2_perf = time.perf_counter()
    processing_latency_ms = calculate_latency(t1_perf)
    event["processing_latency_ms"] = processing_latency_ms
    event["event_generated_at"] = t1_perf

    new_latencies = ((latencies or []) + [processing_latency_ms])[-30:]

    lat_color = COLORS["success"] if processing_latency_ms < 100 else COLORS["warning"] if processing_latency_ms <= 250 else COLORS["danger"]

    # Server timing packet sent to browser clientside callback
    server_timing_packet = {
        "seq": new_count,
        "timestamp": current_time_str,
        "stream_mode": speed or "Normal",
        "polling_interval_ms": current_interval or 1500,
        "processing_latency_ms": processing_latency_ms,
        "t1": t1_perf,
        "t2": t2_perf,
    }

    proc_lat_text = html.Span([
        "Application Processing Latency: ",
        html.Span(f"{processing_latency_ms:.2f} ms", style={"color": lat_color, "fontWeight": "700"}),
    ])

    return (
        new_pos,
        new_events,
        new_latencies,
        new_count,
        f"Last stream update: {current_time_str}",
        proc_lat_text,
        f"Events processed: {new_count}",
        f"Current stream position: {new_pos} / {len(STREAM_EVENTS)}",
        server_timing_packet,
        str(new_count),
        f"{processing_latency_ms:.2f} ms",
        current_time_str,
    )


# ── 2. Stream Speed Control ──────────────────────────────────────────
@dash_app.callback(
    [
        Output("stream-interval", "interval"),
        Output("perf-panel-mode", "children"),
        Output("perf-panel-interval", "children"),
    ],
    Input("speed-dropdown", "value"),
)
def update_stream_speed(speed: str):
    """Adjust interval firing speed: Slow (3000ms), Normal (1500ms), Fast (750ms)."""
    interval_ms = STREAM_SPEEDS.get(speed, 1500)
    mode_name = speed or "Normal"
    return interval_ms, mode_name, f"{interval_ms} ms"


# ── 3. Pause / Resume Streaming ──────────────────────────────────────
@dash_app.callback(
    [
        Output("stream-interval", "disabled"),
        Output("pause-button", "children"),
        Output("pause-button", "color"),
        Output("perf-panel-status", "children"),
        Output("perf-panel-status", "style"),
    ],
    Input("pause-button", "n_clicks"),
    State("stream-interval", "disabled"),
    prevent_initial_call=True,
)
def toggle_streaming(n_clicks: int, is_disabled: bool):
    """Pause or resume the streaming interval and update status indicators."""
    new_state = not is_disabled
    val_style = {"fontSize": "14px", "fontWeight": "700"}
    if new_state:
        return True, "▶ Resume", "success", "PAUSED", {**val_style, "color": COLORS["warning"]}
    return False, "⏸ Pause", "warning", "LIVE", {**val_style, "color": COLORS["success"]}


# ── 4. Reset Filters ─────────────────────────────────────────────────
@dash_app.callback(
    [
        Output("district-dropdown", "value"),
        Output("product-dropdown", "value"),
        Output("cd-band-dropdown", "value"),
    ],
    Input("reset-button", "n_clicks"),
    prevent_initial_call=True,
)
def reset_all_filters(_):
    """Restore filters to defaults: All Districts, All Products, All CD bands."""
    return "All Districts", "All Products", "All"


# ── 5. Main Charts & DataTable Filtering ─────────────────────────────
@dash_app.callback(
    [
        Output("pmmy-chart", "figure"),
        Output("cd-chart", "figure"),
        Output("scatter-chart", "figure"),
        Output("full-table", "data"),
    ],
    [
        Input("district-dropdown", "value"),
        Input("product-dropdown", "value"),
        Input("cd-band-dropdown", "value"),
    ],
)
def update_dashboard_views(district: str, product: str, cd_band: str):
    """Update Row 1, Row 2 scatter, and Row 4 DataTable according to filters."""
    if DATA is None:
        raise dash.exceptions.PreventUpdate

    raw_df = DATA["df"]

    chart_df = filter_data(raw_df, district="All Districts", product=product, cd_band=cd_band)
    table_df = filter_data(raw_df, district=district, product=product, cd_band=cd_band)

    pmmy_fig = build_pmmy_chart(chart_df, product)
    cd_fig = build_cd_chart(chart_df)
    scatter_fig = build_scatter_chart(chart_df, product)

    cols = [
        c for c in [
            "District", "total_accounts", "total_sanction_cr", "avg_sanction_per_account",
            "cd_ratio", "deposits_cr", "advances_cr", "msme_registrations",
            "pmmy_per_100_msme", "sanction_per_msme_lakh", "cd_ratio_band",
        ] if c in table_df.columns
    ]
    table_records = table_df[cols].where(table_df[cols].notna(), None).to_dict("records")

    return pmmy_fig, cd_fig, scatter_fig, table_records


# ── 6. Product Mix Bar Chart ─────────────────────────────────────────
@dash_app.callback(
    Output("product-chart", "figure"),
    [
        Input("product-metric-toggle", "value"),
        Input("district-dropdown", "value"),
        Input("cd-band-dropdown", "value"),
    ],
)
def update_product_mix(metric: str, district: str, cd_band: str):
    """Update PMMY Product Mix chart based on accounts/sanction toggle and filters."""
    if DATA is None:
        raise dash.exceptions.PreventUpdate

    df = filter_data(DATA["df"], district=district, product="All Products", cd_band=cd_band)
    return build_product_chart(df, metric=metric or "accounts")


# ── 7. District Profile Panel ────────────────────────────────────────
@dash_app.callback(
    Output("district-profile-content", "children"),
    Input("district-dropdown", "value"),
)
def update_profile_panel(district: str):
    """Update detailed credit profile card for selected district."""
    if DATA is None:
        raise dash.exceptions.PreventUpdate
    return build_district_profile(DATA["df"], district)


# ── 8. Live Stream Table Monitor ─────────────────────────────────────
@dash_app.callback(
    Output("stream-table-content", "children"),
    Input("stream-events", "data"),
)
def update_stream_display(events: List[Dict]):
    """Update Row 3 stream event table display with latest 10 events."""
    return build_stream_table(events or [])


# ── 9. Latency Benchmarking & Performance Telemetry ──────────────────
@dash_app.callback(
    [
        Output("benchmark-store", "data"),
        Output("client-latency-text", "children"),
        Output("perf-panel-client-lat", "children"),
        Output("perf-panel-bench-progress", "children"),
        Output("benchmark-table-container", "children"),
        Output("latency-chart", "figure"),
        Output("benchmark-summary-chart", "figure"),
    ],
    [
        Input("client-latency-store", "data"),
        Input("btn-reset-bench", "n_clicks"),
    ],
    [
        State("benchmark-store", "data"),
    ],
    prevent_initial_call=True,
)
def update_benchmark_telemetry(client_data: Optional[Dict], reset_clicks: Optional[int], bench_data: Optional[Dict]):
    """Accumulate observations per mode, calculate statistics, and update visual benchmarks."""
    triggered_id = dash.ctx.triggered_id

    current_bench = bench_data or {"Slow": [], "Normal": [], "Fast": []}
    for m in ["Slow", "Normal", "Fast"]:
        if m not in current_bench:
            current_bench[m] = []

    if triggered_id == "btn-reset-bench":
        empty_bench = {"Slow": [], "Normal": [], "Fast": []}
        return (
            empty_bench,
            "Client-Visible Latency: — ms",
            "— ms",
            "Benchmark reset | Slow: 0/10 | Normal: 0/10 | Fast: 0/10",
            build_benchmark_table(empty_bench),
            build_latency_chart([]),
            build_benchmark_summary_chart(empty_bench),
        )

    if not client_data or "client_visible_latency_ms" not in client_data:
        raise dash.exceptions.PreventUpdate

    mode = client_data.get("stream_mode", "Normal")
    if mode not in current_bench:
        current_bench[mode] = []

    obs_entry = {
        "timestamp": client_data.get("timestamp", datetime.now().strftime("%H:%M:%S")),
        "stream_mode": mode,
        "polling_interval_ms": client_data.get("polling_interval_ms", 1500),
        "observation_number": len(current_bench[mode]) + 1,
        "processing_latency_ms": float(client_data.get("processing_latency_ms", 0.0)),
        "client_visible_latency_ms": float(client_data.get("client_visible_latency_ms", 0.0)),
    }
    current_bench[mode].append(obs_entry)

    client_ms = obs_entry["client_visible_latency_ms"]

    slow_count = len(current_bench.get("Slow", []))
    norm_count = len(current_bench.get("Normal", []))
    fast_count = len(current_bench.get("Fast", []))
    progress_text = html.Span([
        html.Span(f"Slow: {slow_count}/10 {'✓' if slow_count>=10 else ''}", style={"color": COLORS["success"] if slow_count>=10 else COLORS["muted"], "marginRight": "10px"}),
        html.Span(f"Normal: {norm_count}/10 {'✓' if norm_count>=10 else ''}", style={"color": COLORS["success"] if norm_count>=10 else COLORS["muted"], "marginRight": "10px"}),
        html.Span(f"Fast: {fast_count}/10 {'✓' if fast_count>=10 else ''}", style={"color": COLORS["success"] if fast_count>=10 else COLORS["muted"]}),
    ])

    all_obs_flat = []
    for m in ["Slow", "Normal", "Fast"]:
        all_obs_flat.extend(current_bench.get(m, []))
    recent_obs = all_obs_flat[-30:]

    fig_latency = build_latency_chart(recent_obs)
    fig_summary = build_benchmark_summary_chart(current_bench)
    table_rendered = build_benchmark_table(current_bench)

    client_text = html.Span([
        "Client-Visible Latency: ",
        html.Span(f"{client_ms:.2f} ms", style={"color": COLORS["blue"], "fontWeight": "700"}),
    ])

    return (
        current_bench,
        client_text,
        f"{client_ms:.2f} ms",
        progress_text,
        table_rendered,
        fig_latency,
        fig_summary,
    )


# ── 10. CSV Export: Raw Observations ──────────────────────────────────
@dash_app.callback(
    Output("download-obs-csv", "data"),
    Input("btn-export-obs-csv", "n_clicks"),
    State("benchmark-store", "data"),
    prevent_initial_call=True,
)
def export_observations_csv(n_clicks: Optional[int], bench_data: Optional[Dict]):
    """Export raw benchmark observations to CSV."""
    if not n_clicks or not bench_data:
        raise dash.exceptions.PreventUpdate

    rows = []
    for mode in ["Slow", "Normal", "Fast"]:
        for obs in bench_data.get(mode, []):
            rows.append({
                "timestamp": obs.get("timestamp"),
                "stream_mode": obs.get("stream_mode"),
                "polling_interval_ms": obs.get("polling_interval_ms"),
                "observation_number": obs.get("observation_number"),
                "processing_latency_ms": obs.get("processing_latency_ms"),
                "client_visible_latency_ms": obs.get("client_visible_latency_ms"),
            })

    if not rows:
        raise dash.exceptions.PreventUpdate

    df_export = pd.DataFrame(rows)
    return dcc.send_data_frame(df_export.to_csv, filename="streaming_latency_benchmark_observations.csv", index=False)


# ── 11. CSV Export: Aggregated Summary ────────────────────────────────
@dash_app.callback(
    Output("download-sum-csv", "data"),
    Input("btn-export-sum-csv", "n_clicks"),
    State("benchmark-store", "data"),
    prevent_initial_call=True,
)
def export_summary_csv(n_clicks: Optional[int], bench_data: Optional[Dict]):
    """Export aggregated benchmark statistical summary to CSV."""
    if not n_clicks or not bench_data:
        raise dash.exceptions.PreventUpdate

    summary_rows = compute_benchmark_summary(bench_data)
    valid_rows = [r for r in summary_rows if r["number_of_observations"] > 0]
    if not valid_rows:
        valid_rows = summary_rows

    df_summary = pd.DataFrame(valid_rows)
    return dcc.send_data_frame(df_summary.to_csv, filename="streaming_latency_benchmark_summary.csv", index=False)


# ═════════════════════════════════════════════════════════════════════
# MAIN ENTRY POINT
# ═════════════════════════════════════════════════════════════════════

if __name__ == "__main__":
    print("\n" + "=" * 70)
    print("  Financial Inclusion & MSME Credit Access — Telangana")
    print("  District Intelligence Dashboard | Group 10 | MBA Analytics")
    print("=" * 70)
    print(f"  → Dashboard URL:  http://127.0.0.1:8050/")
    print(f"  → Streaming Mode: SIMULATED (Historical Source Data Replay)")
    print("=" * 70 + "\n")
    dash_app.run(debug=False, host="127.0.0.1", port=8050)
