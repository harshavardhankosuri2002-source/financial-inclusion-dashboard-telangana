# Financial Inclusion & MSME Credit Access — Telangana

**District Intelligence Dashboard**
Group 10 | MBA Analytics Project

---

## Overview

A production-quality interactive Plotly Dash dashboard for analyzing financial inclusion metrics — PMMY (Pradhan Mantri Mudra Yojana) lending, Credit-Deposit Ratios, and MSME registrations — across Telangana's 33 districts. The dashboard features a simulated real-time streaming engine that replays verified historical data to demonstrate streaming architecture, latency monitoring, and interactive decision intelligence.

> **Important:** This dashboard does not connect to a live banking transaction system. The streaming layer replays verified historical district-level data to demonstrate real-time dashboard architecture, monitoring, and latency measurement.

---

## Features

| Feature | Description |
|---|---|
| **Real-Time Streaming Simulation** | Replays verified district × product events sequentially via `dcc.Interval` |
| **Step 4 Streaming Latency Benchmarking** | Academic dual-clock benchmarking distinguishing Application Processing Latency ($T_2 - T_1$), Client-Visible Latency ($T_3 - T_{\text{req}}$), and Polling Interval ($\tau_{\text{poll}}$) |
| **Benchmark Summary Table & Statistics** | Live empirical computation of Mean, Min, Max, Std Dev, and 95th Percentile across Slow (3000ms), Normal (1500ms), and Fast (750ms) stream modes |
| **Dual-Series Latency Monitor** | Visualizes Application Processing Latency vs Client-Visible Update Latency simultaneously against 100ms and 250ms thresholds |
| **CSV Experiment Export** | Direct export of raw observations and aggregated benchmark summary CSVs for university assignment reporting |
| **Apache Superset Comparison** | Structured architectural comparison framework contextualizing streaming visualization trade-offs |
| **Interactive Filtering** | District, PMMY product, CD Ratio band, and stream speed controls |
| **5 KPI Cards** | PMMY Accounts, Sanction Amount, State CD Ratio, Udyog Aadhaar, Districts |
| **6 Analytical Charts** | PMMY bar chart, CD Ratio bar chart, Product mix, Scatter plot, Latency monitor, Stream table |
| **District Credit Profile** | Drill-down view for any selected district |
| **Full Intelligence Table** | Searchable, sortable, paginated DataTable with CSV export and conditional formatting |
| **Data Provenance** | Transparent methodology panel with source attribution, limitations, and link to [LATENCY_BENCHMARK_METHODOLOGY.md](LATENCY_BENCHMARK_METHODOLOGY.md) |

---

## Installation

### Prerequisites

- Python 3.11 or higher
- pip (Python package manager)

### Steps

```bash
# 1. Navigate to the project directory
cd financial_inclusion_dashboard

# 2. (Recommended) Create a virtual environment
python -m venv venv
venv\Scripts\activate       # Windows
# source venv/bin/activate  # macOS / Linux

# 3. Install dependencies
pip install -r requirements.txt

# 4. Place the dataset file in the project directory
#    File: Group_10_Financial_Inclusion_MSME_Credit_Access_PRO.xlsx

# 5. Run the dashboard
python app.py
```

### Open in Browser

```
http://127.0.0.1:8050/
```

---

## Deploying to Vercel

The project is pre-configured with `vercel.json` and `.vercelignore` for zero-configuration Vercel deployment.

### Method 1: Deploy with Vercel CLI (Quickest)

1. Open PowerShell or Terminal and navigate to the project directory:
   ```bash
   cd C:\Users\aaone\.gemini\antigravity\scratch\financial_inclusion_dashboard
   ```

2. Authenticate with your Vercel account:
   ```bash
   npx vercel login
   ```
   *(This opens your web browser to confirm sign-in with GitHub, GitLab, or Email)*

3. Deploy to production:
   ```bash
   npx vercel --prod
   ```
   - When asked *"Set up and deploy?"*, press **Y**
   - Accept default project name and settings by pressing Enter
   - Vercel will bundle the code and Excel dataset, install dependencies from `requirements.txt`, and give you a live production URL:
   ```
   ✅ Production: https://financial-inclusion-dashboard-xxx.vercel.app
   ```

---

### Method 2: Deploy via GitHub + Vercel Web Dashboard

1. Upload or push this folder to a GitHub repository (e.g. `telangana-financial-inclusion`).
2. Go to **[vercel.com/new](https://vercel.com/new)**.
3. Select your GitHub repository.
4. Click **Deploy** (no build settings need to be modified—`vercel.json` automatically handles `@vercel/python` and entrypoint routing).
5. Your dashboard will be live within ~1–2 minutes!

---

## Dataset Placement

The Excel workbook must be placed in the **same directory** as `app.py`:

```
financial_inclusion_dashboard/
├── app.py
├── requirements.txt
├── README.md
└── Group_10_Financial_Inclusion_MSME_Credit_Access_PRO.xlsx   ← place here
```

### Required Sheets

| Sheet Name | Description |
|---|---|
| `Integrated_View` | **Required.** Combined district-level view with PMMY products, CD Ratio, and MSME data |
| `PMMY_Raw` | Optional. Raw PMMY district-wise data |
| `CD_Ratio` | Optional. District-wise Credit-Deposit Ratio |
| `MSME_Udyog_Aadhaar` | Optional. Legacy MSME registration data |

If the file is missing or the required sheet is absent, the dashboard displays a clear error page with instructions.

---

## Dashboard Sections

### Header
- Project title and subtitle
- Group badge and streaming simulation status indicator
- Source attribution line

### KPI Summary Cards
- **PMMY Accounts** — Total accounts across all products and districts
- **PMMY Sanction** — Total sanctioned amount in ₹ Crore
- **State CD Ratio** — Aggregate advances / deposits × 100
- **Udyog Aadhaar Registrations** — Total legacy MSME registrations
- **Districts Covered** — Number of districts in the dataset

### Filter Bar
- **District** — Filter by individual district or view all
- **PMMY Product** — All Products / Shishu / Kishore / Tarun / Tarun Plus
- **CD Ratio Band** — All / Below 100% / 100–149.99% / 150%+
- **Stream Speed** — Slow (3s) / Normal (1.5s) / Fast (0.75s)
- **Pause / Resume** — Toggle streaming on/off
- **Reset** — Restore all filters to defaults

### Analytical Charts

1. **PMMY Sanction by District (Top 15)** — Horizontal bar chart, descending
2. **CD Ratio by District** — Horizontal bar chart with 100% reference line
3. **PMMY Product Mix** — Grouped bar chart with accounts/sanction toggle
4. **PMMY Activity vs CD Ratio** — Bubble scatter (size = accounts)
5. **District Credit Profile** — Detailed card for selected district
6. **Live Stream Monitor** — Latest 10 simulated events
7. **Processing Latency** — Line chart of last 30 latency measurements

### Full District Intelligence Table
- All districts with computed metrics
- Search, multi-column sort, pagination (15 rows/page)
- CSV export
- Conditional formatting on CD Ratio (red < 100%, amber 100–150%, green > 150%)

---

## Streaming Simulation

### How It Works

The streaming engine uses Dash's `dcc.Interval` component to replay verified historical data as sequential events. Each event contains:

```json
{
    "timestamp": "15:30:45",
    "district": "Hyderabad",
    "product": "Shishu",
    "accounts": 47704,
    "sanction_cr": 210.06,
    "cd_ratio": 113.19
}
```

- The **timestamp** is the dashboard simulation time, **not** the original source transaction date
- Events cycle through all district × product combinations from the verified dataset
- No synthetic or random data is generated — every value comes from the source Excel

### Stream Controls

| Control | Effect |
|---|---|
| Speed: Slow | 3000 ms interval |
| Speed: Normal | 1500 ms interval |
| Speed: Fast | 750 ms interval |
| Pause | Disables the interval timer |
| Resume | Re-enables the interval timer |

---

## Latency Methodology

The dashboard measures **application processing latency** — the time taken by the Python callback to process one stream event:

```python
latency_ms = (time_after_processing - time_before_processing) × 1000
```

This measures:
- Event retrieval from the in-memory dataset
- Data formatting and serialisation

**This is NOT:**
- Network latency to a banking system
- Database query latency
- End-to-end browser rendering time
- An official performance benchmark

### Colour Coding

| Colour | Threshold |
|---|---|
| 🟢 Green | < 100 ms |
| 🟡 Amber | 100–250 ms |
| 🔴 Red | > 250 ms |

---

## Data Sources

1. **Telangana State Level Bankers' Committee (SLBC)**
   - PMMY district-wise data for FY 2025–26, as of 31 March 2026
   - District-wise CD Ratio as of 31 March 2026

2. **Ministry of MSME**
   - District-wise Udyog Aadhaar (legacy MSME registration) data

### Data Notes

- PMMY values are **sanction amounts**, not guaranteed actual disbursement
- CD Ratio is a **broad banking indicator**, not MSME-specific credit penetration
- The PMMY dataset covers **33 current Telangana districts**
- The Udyog Aadhaar dataset covers **32 legacy district entries**
- Missing MSME values are preserved as `N/A` — never replaced with zero
- District name inconsistencies (e.g. Mahbubnagar / Mahabubnagar) are displayed as-is from source

---

## Limitations

- **No live banking data** — all streaming is simulated historical replay
- **No individual-level data** — no personal transactions, gender, caste, or income data
- **No causal claims** — scatter plots show descriptive associations only
- **Sanction ≠ Disbursement** — PMMY figures represent sanctioned amounts
- **CD Ratio ≠ MSME credit** — CD Ratio measures total banking, not MSME-specific lending
- **Legacy MSME data** — Udyog Aadhaar has been superseded by Udyam registration

---

## Academic Interpretation

This dashboard demonstrates:

1. **Real-time/streaming architecture** — `dcc.Interval`-based event replay
2. **Interactive filtering** — Multi-dimensional district/product/band selection
3. **Decision intelligence** — Derived ratios, cross-metric comparison, conditional formatting
4. **Visual analytics** — Six chart types with professional Plotly rendering
5. **Latency monitoring** — Application processing measurement and visualisation
6. **Data provenance** — Transparent source attribution and methodology documentation
7. **Transparent limitations** — Clear labelling of simulation vs. live data

## Apache Superset Contextual Comparison Framework

The comparison framework below contextualizes our empirical Plotly Dash streaming benchmark against Apache Superset's documented architecture:

| Dimension | This Dashboard (Plotly Dash) | Apache Superset |
| :--- | :--- | :--- |
| **Primary Focus** | Micro-level streaming visualization & high-precision latency measurement | Enterprise BI platform, federated ad-hoc SQL querying, and slice caching |
| **Update Mechanism** | Configurable client polling (`dcc.Interval` 750–3000ms) with reactive state stores | Dashboard auto-refresh (UI presets: 10s, 30s, 1m, 5m, etc.; configurable via dashboard JSON metadata with `stagger_refresh`) or asynchronous background query execution (`GLOBAL_ASYNC_QUERIES` via Celery & WebSockets) |
| **Latency Measurement** | Empirical Dual-Clock benchmark: Server `time.perf_counter()` + Client `performance.now()` | Database query duration logged in metadata database (`query` and `logs` tables), chart API response metadata (`/api/v1/chart/data`), and optional StatsD/Prometheus telemetry; no streaming dual-clock event benchmark |
| **Visualization Layer** | Plotly Dash reactive components with real-time rolling DOM patches | Modular chart plugin architecture powered by React and Apache ECharts, with chart slices executing independent query and rendering lifecycles |
| **Streaming Control** | Application-controlled (Pause/Resume, Slow 3000ms, Normal 1500ms, Fast 750ms) | Platform / data-source dependent; queries underlying SQL/OLAP databases (ClickHouse, Pinot, Druid, PostgreSQL) without application-level event simulation controls |
| **Benchmark Conditions** | Controlled single-node environment (Flask WSGI + Chromium client) | Multi-tier enterprise architecture (Flask-AppBuilder web server, Celery worker pool, Redis message broker/results cache, PostgreSQL/MySQL metadata database) |
| **Comparability** | Micro-benchmark under controlled single-process latency constraints | Macro-benchmark of enterprise BI query execution, distributed worker caching, and dashboard slice rendering |

> **Academic Note on Comparability**:
> *Superset timing data and our measured latency are not necessarily like-for-like measurements because architecture, workloads, data sources, and test environments differ substantially. No claim of platform superiority is asserted or implied.*

---

## Tech Stack

| Component | Technology |
|---|---|
| Framework | Dash (Plotly) |
| UI Components | Dash Bootstrap Components |
| Charting | Plotly Graph Objects |
| Data Processing | Pandas, NumPy |
| Excel Reading | openpyxl |
| Backend | Flask (via Dash) |
| Language | Python 3.11+ |

---

## Project Structure

```
financial_inclusion_dashboard/
│
├── app.py                 # Main dashboard application
├── requirements.txt       # Python dependencies
├── README.md              # This file
└── Group_10_Financial_Inclusion_MSME_Credit_Access_PRO.xlsx
                           # Source dataset (user-provided)
```

---

## Troubleshooting

| Issue | Solution |
|---|---|
| "Dataset file not found" error page | Place the Excel file in the same directory as `app.py` |
| "Required sheet not found" error | Ensure the workbook contains an `Integrated_View` sheet |
| Charts show "Data not available" | Check that column names in the Excel match expected patterns (see app.py `find_col` function) |
| Port 8050 already in use | Kill any existing Dash process, or change the port in `app.run()` |
| Streaming appears frozen | Check that the Pause button hasn't been activated |

---

*Group 10 | Financial Inclusion & MSME Credit Access | MBA Analytics Project*
*Streaming mode: simulated historical-data replay*
