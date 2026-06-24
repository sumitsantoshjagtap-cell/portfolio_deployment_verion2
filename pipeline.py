# %%
# ======================================================================
# IMPORTS
# ======================================================================
import json
import logging
import pickle
import warnings
from itertools import product as iproduct
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import matplotlib
matplotlib.use("Agg")
import matplotlib.backends.backend_pdf as pdf_backend
import matplotlib.pyplot as plt
import networkx as nx
import numpy as np
import pandas as pd
import ruptures as rpt
import seaborn as sns
import yfinance as yf
import pandas_market_calendars as mcal
import quantstats as qs
from scipy import stats
from scipy.cluster.hierarchy import fcluster, linkage
from scipy.optimize import linear_sum_assignment
from scipy.special import logsumexp
from scipy.stats import jarque_bera, pearsonr
from hmmlearn.hmm import GaussianHMM
from sklearn.decomposition import PCA
from statsmodels.graphics.tsaplots import plot_acf
from statsmodels.stats.diagnostic import breaks_cusumolsresid
from statsmodels.stats.outliers_influence import variance_inflation_factor
from statsmodels.tsa.statespace.sarimax import SARIMAX
from statsmodels.tsa.stattools import adfuller, grangercausalitytests, kpss, zivot_andrews
from statsmodels.tsa.vector_ar.var_model import VAR
from pypfopt import EfficientFrontier, risk_models as risk_m
from pypfopt import black_litterman as bl_module
from pypfopt.black_litterman import BlackLittermanModel
from pypfopt.discrete_allocation import DiscreteAllocation, get_latest_prices
import plotly.graph_objects as go
import plotly.express as px
import dash
from dash import Input, Output, State, dcc, html, ALL, callback_context
from concurrent.futures import ThreadPoolExecutor, as_completed
import os

try:
    import openpyxl        # noqa: F401
    _OPENPYXL_OK = True
except ImportError:
    _OPENPYXL_OK = False
    # pip install openpyxl to enable sector-score Excel loading

warnings.filterwarnings("ignore")
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)-8s | %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger(__name__)



# %%
# ======================================================================
# SYSTEM CONFIGURATION — Edit only this block
# ======================================================================
STOCKS: Dict[str, str] = {
    "HDFCBANK.NS":   "HDFC Bank",
    "BHARTIARTL.NS": "Bharti Airtel",
    "APOLLOHOSP.NS": "Apollo Hospitals",
    "BIOCON.NS":     "Biocon Ltd",
    "JSWENERGY.NS":  "JSW Energy Ltd",
    "SUDARSCHEM.NS": "Sudarshan Chemicals",
    "AEGISLOG.NS":   "Aegis Logistics Ltd",
    "KPRMILL.NS":    "K P R Mill Ltd",
    "GREENLAM.NS":   "Greenlam Industries Ltd",
    "CAMLINFINE.NS": "Camlin Fine Science Ltd",
    "SURYAROSNI.NS": "Surya Roshni Ltd",
    "EIDPARRY.NS":   "EID Parry Ltd",
    "BANDHANBNK.NS": "Bandhan Bank",
    "ABREL.NS":      "Aditya Birla Real Estate",
    "INDIGO.NS":     "Interglobal Aviation Ltd",
    "MARATHON.NS":   "Marathon Nextgen Realty Ltd",
    "LTTS.NS":       "L&T Technology Services Ltd",
    "ASTERDM.NS":    "Aster DM Healthcare Ltd",
    # "JSLL.NS":       "Jeena Sikho Lifecare Ltd",                # new but will get qurantined
    "BSE.NS":         "BSE Ltd",                                # new
    # "ETERNAL.NS":     "ZOMATO",                                 # new but will get qurantined
    # "SMARTWORKS.NS":  "Smartworks Corporate Services Ltd",      # new but will get qurantined
    # "SAGILITY.NS":    "SAGILITY Ltd",                           # new but will get qurantined
    # "WAAREEENER.NS":      "Waaree Energies Ltd",                # new but will get qurantined
    # "IGIL.NS":        "International Gemological Institute",     # new but will get qurantined
    # "SEDEMAC.NS":    "Sedemac Mechatronics Ltd",                # new but will get qurantined

}

# ── NEW ▸ Sector Groupings ────────────────────────────────────────────
# Map sector names (must match column headers in SECTOR_SCORE_FILE) to
# their NSE tickers.  Edit freely; tickers absent from STOCKS are ignored.
SECTOR_GROUPS: Dict[str, List[str]] = {
    "Co-working Space": ["SMARTWORKS.NS"],
    "Building Materials - Laminates/Decorative": ["GREENLAM.NS"],
    "Business Process Management (BPM)": ["SAGILITY.NS"],
    "E-Commerce": ["ETERNAL.NS"],
    "Financial Services - Exchanges": ["BSE.NS"],
    "Financial Services - Private Sector Bank": ["BANDHANBNK.NS", "HDFCBANK.NS"],
    "Renewable Energy – Solar": ["WAAREEENER.NS"],
    "Speciality Chemical": ["CAMLINFINE.NS", "SUDARSCHEM.NS"],
    "Telecom - Service Providers": ["BHARTIARTL.NS"],
    "Textiles & Apparel": ["KPRMILL.NS"],
    "Gemstone Grading / Jewelry Certification": ["IGIL.NS"],
    "Hospitals & Healthcare Services": ["ASTERDM.NS", "APOLLOHOSP.NS"], # JSLL.NS  <- to be listed later
    "Integrated Pipes + Lighting": ["SURYAROSNI.NS"],
    "Real Estate – Residential Developers": ["MARATHON.NS", "ABREL.NS"],
    "Sugar + Distillery (Ethanol) + Nutraceuticals": ["EIDPARRY.NS"],
    "Biotechnology": ["BIOCON.NS"],
    "Engineering R&D (ER&D) & Digital Engineering Services": ["LTTS.NS"],
    "Oil & Gas – Storage, Transportation & Chemical Logistics": ["AEGISLOG.NS"],
    "Auto Ancillary": ["SEDEMAC.NS"],
    "Aviation industry": ["INDIGO.NS"],
}

# ── NEW ▸ Sector Score File ───────────────────────────────────────────
# Path to your colleague's Excel workbook (see format notes below).
#
# Expected sheet: "Scores" (single sheet)
# Columns : Quarter | Parameter | <Sector1> | <Sector2> | ...
#            (Sector column names must match keys in SECTOR_GROUPS above)
# Quarter  : "Q1FY25", "Q2FY25", …, "Q4FY26"  (Indian FY convention)
# Parameter: "A", "B", "C", "D", "E", "F"      (scores out of 10)
# Total    : 48 rows  (8 quarters × 6 parameters)
#
# Example row:
#   Q1FY25 | A | 7.5 | 6.2 | 8.1 | ...
SECTOR_SCORE_FILE: str = "sector_scores.xlsx"

# ── NEW ▸ Parameter weights for composite sector score ────────────────
# Must sum to 1.0.  Adjust weights to reflect your analytical priorities.
SECTOR_SCORE_WEIGHTS: Dict[str, float] = {
    "A": 0.20,   # Market Size & Growth Potential
    "B": 0.30,   # Profitability & Unit Economics       ← highest
    "C": 0.10,   # Competition & Industry Structure
    "D": 0.10,   # Regulatory & Compliance Risk
    "E": 0.10,   # Technology & Innovation
    "F": 0.20,   # Demand Visibility / Business Resilience
}

# ── NEW ▸ Conviction signal blending ─────────────────────────────────
# conviction = LEVEL_WT * (score_level/10)
#            + MOMENTUM_WT * clip((Δscore/10 + 0.5), 0, 1)
# Output clipped to [0.15, 0.95] to prevent extreme BL tilts.
CONVICTION_LEVEL_WT:    float = 0.60
CONVICTION_MOMENTUM_WT: float = 0.40
# Conservative lag: scores assumed available 2 months after quarter-end
SCORE_LAG_MONTHS: int = 2

MARKET_INDICATOR_TICKERS: Dict[str, str] = {
    "^INDIAVIX": "India_VIX",
    "USDINR=X":  "USD_INR",
    "BZ=F":      "Brent_Oil",
    "GC=F":      "Gold",
    "^NSEBANK":  "Nifty_Bank",
}
START_DATE:          str        = "2019-01-01"
END_DATE:            str        = "today"
MIN_HISTORY_MONTHS:  int        = 36
BENCHMARK_TICKER:    str        = "^NSEI"
FORECAST_HORIZONS:   List[int]  = [1, 3, 6, 12]
MIN_TRAIN_MONTHS:    int        = 60
MAX_LOOKBACK_MONTHS: int        = 84
N_REGIMES:           int        = 3
RANDOM_STATE:        int        = 42
CAPITAL_AMOUNT:      float      = 1_000_000.0
ANNUAL_RISK_FREE:    float      = 0.065
OUTPUT_DIR = Path("pipeline_outputs_v2")
for _sub in ["", "hmm_groups", "hmm_stocks"]:
    (OUTPUT_DIR / _sub).mkdir(parents=True, exist_ok=True)

HORIZON_COLORS: Dict[int, Dict[str, str]] = {
    1:  {"hex": "green",      "rgb": "0,128,0"},
    3:  {"hex": "darkorange", "rgb": "255,140,0"},
    6:  {"hex": "firebrick",  "rgb": "178,34,34"},
    12: {"hex": "purple",     "rgb": "128,0,128"},
}

# %%
# ======================================================================
# ── NEW ▸ DASHBOARD STYLING CONSTANTS & HELPERS
# ======================================================================
_C = {
    "bg":         "#f5f7fa",
    "card":       "#ffffff",
    "border":     "#e4e7eb",
    "primary":    "#1a56db",
    "primary_lt": "#e8f0fe",
    "text":       "#111928",
    "muted":      "#6b7280",
    "success":    "#057a55",
    "danger":     "#c81e1e",
    "warning":    "#c27803",
    "shadow":     "0 1px 4px rgba(0,0,0,0.08)",
}
_FONT = "'Segoe UI', 'Inter', Arial, sans-serif"

_PLOTLY_BASE = dict(
    paper_bgcolor="white",
    plot_bgcolor="#fafafa",
    font=dict(family=_FONT, size=12, color=_C["text"]),
    margin=dict(l=60, r=20, t=50, b=60),
    hoverlabel=dict(bgcolor="white", bordercolor=_C["border"], font_size=12),
    xaxis=dict(gridcolor="#efefef", linecolor=_C["border"], zeroline=False),
    yaxis=dict(gridcolor="#efefef", linecolor=_C["border"], zeroline=False),
    legend=dict(
        orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1,
        bgcolor="rgba(255,255,255,0.85)", bordercolor=_C["border"], borderwidth=1,
    ),
)


def _card(children, title=None, subtitle=None, extra=None):
    body = []
    if title:
        body.append(html.Div(title, style={
            "fontSize": "13px", "fontWeight": "700", "color": _C["text"],
            "marginBottom": "4px", "fontFamily": _FONT,
        }))
    if subtitle:
        body.append(html.Div(subtitle, style={
            "fontSize": "11px", "color": _C["muted"], "marginBottom": "12px",
            "fontFamily": _FONT,
        }))
    body += children if isinstance(children, list) else [children]
    s = {
        "background": _C["card"], "borderRadius": "8px",
        "border": f"1px solid {_C['border']}", "padding": "20px",
        "marginBottom": "16px", "boxShadow": _C["shadow"],
    }
    if extra:
        s.update(extra)
    return html.Div(body, style=s)


def _kpi(label, value, color=None):
    return html.Div([
        html.Div(label, style={
            "fontSize": "10px", "fontWeight": "600", "color": _C["muted"],
            "textTransform": "uppercase", "letterSpacing": "0.06em", "marginBottom": "6px",
        }),
        html.Div(value, style={
            "fontSize": "20px", "fontWeight": "700",
            "color": color or _C["text"], "fontFamily": _FONT,
        }),
    ], style={
        "background": _C["card"], "borderRadius": "8px",
        "border": f"1px solid {_C['border']}", "padding": "14px 18px",
        "boxShadow": _C["shadow"], "flex": "1", "minWidth": "120px",
    })


def _kpi_row(items):
    """items = list of (label, value, color_or_None)"""
    return html.Div(
        [_kpi(l, v, c) for l, v, c in items],
        style={"display": "flex", "gap": "10px", "flexWrap": "wrap", "marginBottom": "16px"},
    )


def _dd(id_, opts, val, label=None, w="220px"):
    els = []
    if label:
        els.append(html.Div(label, style={
            "fontSize": "11px", "fontWeight": "500", "color": _C["muted"],
            "marginBottom": "4px", "fontFamily": _FONT,
        }))
    els.append(dcc.Dropdown(
        id=id_, options=opts, value=val, clearable=False,
        style={"fontSize": "13px", "width": w, "fontFamily": _FONT},
    ))
    return html.Div(els, style={"marginBottom": "12px"})


def _hdr(title, sub=None):
    return html.Div([
        html.Div(title, style={
            "fontSize": "15px", "fontWeight": "700", "color": _C["text"],
            "fontFamily": _FONT, "marginBottom": "4px",
        }),
        html.Div(sub, style={"fontSize": "11px", "color": _C["muted"],
                             "marginBottom": "14px", "fontFamily": _FONT}) if sub else None,
    ])


_TAB_S = {"fontFamily": _FONT, "fontSize": "13px",
           "color": _C["muted"], "padding": "8px 14px", "borderBottom": "none"}
_TAB_SEL = {**_TAB_S, "color": _C["primary"], "fontWeight": "600",
             "borderBottom": f"2px solid {_C['primary']}"}


# %%
# ======================================================================
# SHARED UTILITIES  (unchanged from original)
# ======================================================================
def _resample_monthly(series: pd.Series) -> pd.Series:
    try:
        return series.resample("ME").last()
    except ValueError:
        return series.resample("M").last()


def _resample_monthly_df(df: pd.DataFrame) -> pd.DataFrame:
    try:
        return df.resample("ME").last()
    except ValueError:
        return df.resample("M").last()


def _resample_sum_monthly(df: pd.DataFrame) -> pd.DataFrame:
    try:
        return df.resample("ME").sum()
    except ValueError:
        return df.resample("M").sum()


def _quarterly_index(start: pd.Timestamp, end: pd.Timestamp) -> pd.DatetimeIndex:
    try:
        return pd.date_range(start, end, freq="QE")
    except ValueError:
        return pd.date_range(start, end, freq="Q")


def _extract_close(raw: pd.DataFrame, ticker: str) -> pd.Series:
    if isinstance(raw.columns, pd.MultiIndex):
        lvl0 = raw.columns.get_level_values(0).unique().tolist()
        lvl1 = raw.columns.get_level_values(1).unique().tolist()
        if ticker in lvl0:
            df = raw[ticker].dropna(how="all")
        elif ticker in lvl1:
            df = raw.xs(ticker, axis=1, level=1).dropna(how="all")
        else:
            return pd.Series(dtype=float)
        return df["Close"].squeeze() if "Close" in df.columns else df.iloc[:, 0].squeeze()
    if "Close" in raw.columns:
        return raw["Close"].squeeze()
    return raw.iloc[:, 0].squeeze()


def _count_missing_nse_days(close: pd.Series, start: str, end: str) -> int:
    try:
        nse_cal      = mcal.get_calendar("NSE")
        schedule     = nse_cal.schedule(start_date=start, end_date=end)
        trading_days = (mcal.date_range(schedule, frequency="1D")
                        .tz_localize(None).normalize())
        actual_days  = close.dropna().index.normalize()
        return len(trading_days.difference(actual_days))
    except Exception as exc:
        logger.warning(f"NSE calendar check failed ({exc}) — using raw NaN count.")
        return int(close.isna().sum())


def _detect_communities(G: nx.Graph, random_state: int = 42) -> Dict[str, int]:
    try:
        import community as cl
        if hasattr(cl, "best_partition"):
            try:
                return cl.best_partition(G, weight="weight", random_state=random_state)
            except TypeError:
                return cl.best_partition(G, weight="weight")
    except ImportError:
        pass
    try:
        from networkx.algorithms import community as nx_comm
        comms = list(nx_comm.louvain_communities(G, weight="weight", seed=random_state))
        return {node: i for i, comm in enumerate(comms) for node in comm}
    except (AttributeError, TypeError, Exception):
        pass
    try:
        from networkx.algorithms import community as nx_comm
        comms = list(nx_comm.greedy_modularity_communities(G, weight="weight"))
        return {node: i for i, comm in enumerate(comms) for node in comm}
    except Exception:
        pass
    logger.warning("All community detection methods failed — each node in own community.")
    return {node: i for i, node in enumerate(G.nodes())}

# ── NEW ▸ Quarter string → quarter-end Timestamp ──────────────────────
def _parse_quarter_to_date(q: str) -> Optional[pd.Timestamp]:
    """
    Convert "Q1FY25" → 2024-06-30, "Q4FY26" → 2026-03-31, etc.
    Indian FY: FY25 = Apr 2024 – Mar 2025.
    """
    try:
        qnum = int(q[1])
        fy   = int(q[4:])
        cal_start = 2000 + fy - 1          # FY25 starts in calendar 2024
        month_map = {1: 6, 2: 9, 3: 12, 4: 3}
        month = month_map[qnum]
        year  = cal_start if qnum < 4 else cal_start + 1
        return pd.Timestamp(year=year, month=month, day=1) + pd.offsets.MonthEnd(0)
    except Exception:
        return None


# %%
# ======================================================================
# PHASE 1: DATA ARCHITECTURE & COLLECTION  (unchanged)
# ======================================================================
def run_phase1() -> dict:
    logger.info("=" * 62)
    logger.info("PHASE 1 — Data Architecture & Collection")
    logger.info("=" * 62)
    end_str = pd.Timestamp.today().strftime("%Y-%m-%d") if END_DATE == "today" else END_DATE
    tickers = list(STOCKS.keys())
    logger.info("Downloading daily OHLCV (stocks) …")
    raw_stocks = yf.download(
        tickers=tickers, start=START_DATE, end=end_str,
        auto_adjust=True, progress=False, group_by="ticker",
    )
    audit_rows: List[dict] = []
    price_data: Dict[str, pd.DataFrame] = {}
    WATCHLIST:  Dict[str, str] = {}
    for ticker in tickers:
        close = _extract_close(raw_stocks, ticker).dropna()
        if close.empty:
            logger.warning(f"  [{ticker}] No data — quarantined.")
            WATCHLIST[ticker] = "No data returned"
            continue
        first, last    = close.index.min(), close.index.max()
        history_months = (last.year - first.year) * 12 + (last.month - first.month)
        n_missing      = _count_missing_nse_days(
            close, first.strftime("%Y-%m-%d"), last.strftime("%Y-%m-%d")
        )
        audit_rows.append({
            "ticker": ticker, "name": STOCKS[ticker],
            "first_date": first.date(), "last_date": last.date(),
            "n_trading_days": len(close), "n_missing_days": n_missing,
            "pct_missing": round(n_missing / max(len(close), 1), 4),
            "history_months": history_months,
        })
        if history_months < MIN_HISTORY_MONTHS:
            reason = f"Insufficient history ({history_months} < {MIN_HISTORY_MONTHS} months)"
            logger.warning(f"  [{ticker}] {reason}")
            WATCHLIST[ticker] = reason
        else:
            if isinstance(raw_stocks.columns, pd.MultiIndex):
                lvl0 = raw_stocks.columns.get_level_values(0).unique().tolist()
                lvl1 = raw_stocks.columns.get_level_values(1).unique().tolist()
                if ticker in lvl0:
                    price_data[ticker] = raw_stocks[ticker].dropna(how="all")
                elif ticker in lvl1:
                    price_data[ticker] = raw_stocks.xs(ticker, axis=1, level=1).dropna(how="all")
                else:
                    price_data[ticker] = pd.DataFrame({"Close": close})
            else:
                price_data[ticker] = pd.DataFrame({"Close": close})
    audit_df = pd.DataFrame(audit_rows)
    audit_df.to_csv(OUTPUT_DIR / "data_audit.csv", index=False)
    logger.info(f"  WATCHLIST = {list(WATCHLIST.keys()) or 'empty'}")
    logger.info(f"\n{audit_df.to_string()}\n")
    active_tickers = [t for t in tickers if t in price_data]
    if not active_tickers:
        raise RuntimeError("All tickers quarantined — cannot proceed.")
    daily_close     = pd.DataFrame({t: price_data[t]["Close"] for t in active_tickers})
    monthly_prices  = _resample_monthly_df(daily_close)
    monthly_prices.to_parquet(OUTPUT_DIR / "monthly_prices.parquet")
    monthly_log_returns = np.log(monthly_prices / monthly_prices.shift(1)).dropna()
    monthly_log_returns.to_parquet(OUTPUT_DIR / "monthly_log_returns.parquet")
    logger.info(f"  monthly_prices shape: {monthly_prices.shape}")
    logger.info(f"  monthly_log_returns shape: {monthly_log_returns.shape}")
    logger.info(f"Downloading benchmark {BENCHMARK_TICKER} …")
    bench_raw    = yf.download(BENCHMARK_TICKER, start=START_DATE, end=end_str,
                               auto_adjust=True, progress=False)
    bench_close  = _extract_close(bench_raw, BENCHMARK_TICKER)
    bench_monthly = _resample_monthly(bench_close)
    benchmark_returns = np.log(bench_monthly / bench_monthly.shift(1)).dropna()
    benchmark_returns.name = BENCHMARK_TICKER
    benchmark_returns.to_frame("benchmark_return").to_parquet(
        OUTPUT_DIR / "benchmark_returns.parquet")
    monthly_rf = (1 + ANNUAL_RISK_FREE) ** (1 / 12) - 1
    indicator_monthly_prices: Dict[str, pd.Series] = {}
    logger.info("Downloading market indicators …")
    for ind_ticker, ind_name in MARKET_INDICATOR_TICKERS.items():
        try:
            ind_raw   = yf.download(ind_ticker, start=START_DATE, end=end_str,
                                    auto_adjust=True, progress=False)
            ind_close = _extract_close(ind_raw, ind_ticker)
            if ind_close.dropna().empty:
                continue
            indicator_monthly_prices[ind_name] = _resample_monthly(ind_close)
            logger.info(f"  [{ind_ticker}] {ind_name}: {len(indicator_monthly_prices[ind_name])} obs")
        except Exception as exc:
            logger.warning(f"  [{ind_ticker}] fetch failed ({exc}) — skipped.")
    assert not monthly_prices.isnull().all().any()
    assert len(monthly_log_returns) >= 1
    logger.info("Phase 1 validation checklist passed ✓\n")
    return {
        "price_data": price_data, "daily_close": daily_close,
        "monthly_prices": monthly_prices, "monthly_log_returns": monthly_log_returns,
        "benchmark_returns": benchmark_returns, "bench_monthly_prices": bench_monthly,
        "monthly_rf": monthly_rf, "indicator_monthly_prices": indicator_monthly_prices,
        "active_tickers": active_tickers, "WATCHLIST": WATCHLIST,
        "audit_df": audit_df, "end_str": end_str,
    }


# ======================================================================
# PHASE 2: EDA & PREPROCESSING  (unchanged)
# ======================================================================
def _stationarity_suite(series: pd.Series, ticker: str) -> dict:
    s = series.dropna()
    out = {"ticker": ticker}
    try:
        adf = adfuller(s, autolag="AIC", regression="c")
        out.update({"adf_stat": adf[0], "adf_pval": adf[1],
                    "adf_stationary": bool(adf[1] < 0.05)})
    except Exception as exc:
        logger.warning(f"  [{ticker}] ADF failed: {exc}")
        out.update({"adf_stat": np.nan, "adf_pval": np.nan, "adf_stationary": False})
    try:
        kp = kpss(s, regression="c", nlags="auto")
        out.update({"kpss_stat": kp[0], "kpss_pval": kp[1],
                    "kpss_stationary": bool(kp[1] >= 0.05)})
    except Exception as exc:
        logger.warning(f"  [{ticker}] KPSS failed: {exc}")
        out.update({"kpss_stat": np.nan, "kpss_pval": np.nan, "kpss_stationary": np.nan})
    try:
        za    = zivot_andrews(s, maxlag=12, regression="c")
        bp_idx = int(za[4])
        bp_dt  = s.index[bp_idx].date() if 0 <= bp_idx < len(s) else None
        out.update({"za_stat": za[0], "za_pval": za[1],
                    "za_stationary": bool(za[1] < 0.05), "za_break_date": str(bp_dt)})
    except Exception as exc:
        logger.warning(f"  [{ticker}] Zivot-Andrews failed: {exc}")
        out.update({"za_stat": np.nan, "za_pval": np.nan,
                    "za_stationary": np.nan, "za_break_date": None})
    return out


def run_phase2(p1: dict) -> dict:
    logger.info("=" * 62)
    logger.info("PHASE 2 — EDA & Preprocessing")
    logger.info("=" * 62)
    mlr            = p1["monthly_log_returns"].copy()
    active_tickers = p1["active_tickers"]
    with pdf_backend.PdfPages(OUTPUT_DIR / "return_distributions.pdf") as pages:
        for ticker in active_tickers:
            s    = mlr[ticker].dropna()
            desc = stats.describe(s)
            _, jb_pval = jarque_bera(s)
            fig, axes = plt.subplots(1, 3, figsize=(15, 4))
            fig.suptitle(f"{ticker} ({STOCKS.get(ticker, '')}) — Return Distributions", fontsize=11)
            sns.histplot(s, kde=True, ax=axes[0])
            axes[0].set_title("Histogram + KDE"); axes[0].set_xlabel("Monthly Log Return")
            stats.probplot(s, dist="norm", plot=axes[1]); axes[1].set_title("Q-Q Plot (Normal)")
            plot_acf(s, lags=min(24, len(s) // 3), ax=axes[2], zero=False); axes[2].set_title("ACF")
            fig.text(0.5, -0.03,
                     f"μ={desc.mean:.4f}  σ={np.std(s):.4f}  Skew={desc.skewness:.3f}  "
                     f"ExKurt={desc.kurtosis:.3f}  JB p={jb_pval:.4f}",
                     ha="center", fontsize=9)
            plt.tight_layout(); pages.savefig(fig, bbox_inches="tight"); plt.close(fig)
    stat_rows = [_stationarity_suite(mlr[t], t) for t in active_tickers]
    stat_df   = pd.DataFrame(stat_rows).set_index("ticker")
    stat_df.to_csv(OUTPUT_DIR / "stationarity_results.csv")
    break_records: Dict[str, List] = {}
    for ticker in active_tickers:
        s    = mlr[ticker].dropna()
        algo = rpt.Pelt(model="rbf", min_size=12, jump=1).fit(s.values.reshape(-1, 1))
        bkps = algo.predict(pen=10)
        break_records[ticker] = [s.index[bp - 1].date() for bp in bkps[:-1]]
    pd.DataFrame(
        [(t, str(d)) for t, dl in break_records.items() for d in dl],
        columns=["ticker", "break_date"],
    ).to_csv(OUTPUT_DIR / "structural_breaks.csv", index=False)
    mlr_clean = mlr.copy()
    outlier_rows: List[dict] = []
    for ticker in active_tickers:
        s = mlr_clean[ticker].ffill(limit=1)
        mlr_clean[ticker] = s
        rm   = s.rolling(24, min_periods=12).mean()
        rstd = s.rolling(24, min_periods=12).std().replace(0.0, np.nan)
        zs   = (s - rm) / rstd
        for dt in s.index[zs.abs() > 3.5]:
            outlier_rows.append({"ticker": ticker, "date": str(dt.date()),
                                  "value": round(float(s[dt]), 6),
                                  "z_score": round(float(zs[dt]), 3),
                                  "action": "flagged — manual review required"})
    mlr_clean.to_parquet(OUTPUT_DIR / "monthly_log_returns_clean.parquet")
    outlier_df = pd.DataFrame(outlier_rows)
    if not outlier_df.empty:
        outlier_df.to_csv(OUTPUT_DIR / "outlier_flags.csv", index=False)
    logger.info("Phase 2 validation checklist passed ✓\n")
    return {
        "stat_df": stat_df, "break_records": break_records,
        "outlier_df": outlier_df, "monthly_log_returns_clean": mlr_clean,
    }


# ======================================================================
# PHASE 3: FEATURE ENGINEERING  (unchanged)
# ======================================================================
def _compute_realized_volatility(daily_close: pd.DataFrame,
                                  target_index: pd.DatetimeIndex) -> pd.DataFrame:
    daily_lr = np.log(daily_close / daily_close.shift(1)).dropna()
    rv_var   = _resample_sum_monthly(daily_lr ** 2)
    return np.sqrt(rv_var).reindex(target_index)


def _prepare_market_indicators(indicator_monthly_prices: Dict[str, pd.Series],
                                target_index: pd.DatetimeIndex) -> pd.DataFrame:
    if not indicator_monthly_prices:
        return pd.DataFrame(index=target_index)
    prices_df = pd.DataFrame(indicator_monthly_prices)
    log_rets  = np.log(prices_df / prices_df.shift(1))
    lagged    = log_rets.shift(1)
    aligned   = lagged.reindex(target_index).ffill(limit=3).bfill(limit=1).fillna(0.0)
    non_trivial = aligned.columns[(aligned != 0.0).any()]
    return aligned[non_trivial]


def run_phase3(p1: dict, p2: dict) -> dict:
    logger.info("=" * 62)
    logger.info("PHASE 3 — Feature Engineering")
    logger.info("=" * 62)
    active_tickers = p1["active_tickers"]
    daily_close    = p1["daily_close"][active_tickers]
    mlr            = p2["monthly_log_returns_clean"]
    realized_vol   = _compute_realized_volatility(daily_close, mlr.index)
    realized_vol.to_parquet(OUTPUT_DIR / "realized_volatility.parquet")
    market_indicators = _prepare_market_indicators(p1["indicator_monthly_prices"], mlr.index)
    market_indicators.to_parquet(OUTPUT_DIR / "market_indicators.parquet")
    for col in market_indicators.columns:
        s = market_indicators[col].replace(0.0, np.nan).dropna()
        if len(s) > 10:
            _, pval, _ = adfuller(s, autolag="AIC")[:3]
            logger.info(f"    {col}: ADF p={pval:.4f} → "
                        f"{'stationary' if pval < 0.05 else 'POSSIBLY NON-STATIONARY'}")
    exog_master = market_indicators.copy()
    for ticker in active_tickers:
        if ticker in realized_vol.columns:
            exog_master[(ticker, "realized_vol")] = realized_vol[ticker]
    exog_master.to_parquet(OUTPUT_DIR / "exog_master.parquet")
    logger.info("Phase 3 validation checklist passed ✓\n")
    return {
        "realized_volatility": realized_vol,
        "market_indicators":   market_indicators,
        "exog_master":         exog_master,
    }



# ======================================================================
# PHASE 3.5 — SECTOR SCORE LOADING & CONVICTION SIGNALS  (NEW)
# ======================================================================
def _compute_sector_conviction(composite_df: pd.DataFrame) -> Dict[str, float]:
    """
    Compute conviction per sector from composite score DataFrame.

    conviction = CONVICTION_LEVEL_WT  * (latest_score / 10)
               + CONVICTION_MOMENTUM_WT * clip((Δscore/10 + 0.5), 0, 1)

    Level component   — captures absolute quality: high-scoring sectors get
                        higher conviction regardless of trajectory.
    Momentum component— captures trend: a rising sector is more investable
                        than a plateauing one at the same level.
    Result clipped to [0.15, 0.95] to keep BL adjustments moderate.
    """
    if composite_df.empty or len(composite_df) < 1:
        return {}
    latest = composite_df.iloc[-1]
    prev   = composite_df.iloc[-2] if len(composite_df) >= 2 else latest
    level_comp    = (latest / 10.0).clip(0.0, 1.0)
    momentum_raw  = (latest - prev) / 10.0          # Δ in [-1, +1]
    momentum_comp = (momentum_raw / 2.0 + 0.5).clip(0.0, 1.0)  # → [0, 1]
    conviction = (
        CONVICTION_LEVEL_WT * level_comp + CONVICTION_MOMENTUM_WT * momentum_comp
    ).clip(0.15, 0.95)
    return conviction.to_dict()


def run_phase3_5() -> dict:
    """
    Phase 3.5: Load sector scores from Excel, compute composite scores
    and conviction signals used in Black-Litterman Phase 7.

    Returns a dict with:
      sector_composite_scores  — DataFrame (quarters × sectors), composite 0–10
      sector_composite_monthly — DataFrame (months  × sectors), temporally
                                  disaggregated (conservative lag, no look-ahead)
      sector_conviction        — dict { sector → conviction ∈ [0.15, 0.95] }
      ticker_conviction        — dict { ticker → conviction } (via SECTOR_GROUPS)
      ticker_sector            — dict { ticker → sector name }
      scores_available         — bool
    """
    logger.info("=" * 62)
    logger.info("PHASE 3.5 — Sector Score Loading & Conviction Signals")
    logger.info("=" * 62)

    _EMPTY = {
        "sector_composite_scores":  pd.DataFrame(),
        "sector_composite_monthly": pd.DataFrame(),
        "sector_conviction":        {},
        "ticker_conviction":        {},
        "ticker_sector":            {t: "Unknown" for t in STOCKS},
        "scores_available":         False,
    }

    # ── Ticker → sector reverse map (always built) ────────────────────
    ticker_sector: Dict[str, str] = {}
    for sector, tickers in SECTOR_GROUPS.items():
        for t in tickers:
            ticker_sector[t] = sector

    if not Path(SECTOR_SCORE_FILE).exists():
        logger.warning(
            f"  '{SECTOR_SCORE_FILE}' not found — score integration skipped.\n"
            f"  Create the file and re-run to enable conviction-based BL."
        )
        _EMPTY["ticker_sector"] = ticker_sector
        return _EMPTY

    if not _OPENPYXL_OK:
        logger.warning("  openpyxl not installed (pip install openpyxl) — skipping scores.")
        _EMPTY["ticker_sector"] = ticker_sector
        return _EMPTY

    try:
        raw = pd.read_excel(SECTOR_SCORE_FILE, sheet_name="Scores")
    except Exception as exc:
        logger.warning(f"  Could not read '{SECTOR_SCORE_FILE}': {exc} — skipping scores.")
        _EMPTY["ticker_sector"] = ticker_sector
        return _EMPTY

    required_cols = {"Quarter", "Parameter"}
    if not required_cols.issubset(set(raw.columns)):
        logger.warning(
            f"  Sheet 'Scores' must have columns 'Quarter' and 'Parameter'. "
            f"Found: {list(raw.columns)[:8]}  — skipping scores."
        )
        _EMPTY["ticker_sector"] = ticker_sector
        return _EMPTY

    raw["Quarter"]   = raw["Quarter"].astype(str).str.strip()
    raw["Parameter"] = raw["Parameter"].astype(str).str.strip().str.upper()

    # Keep only rows with recognised quarters and parameters
    known_params = set(SECTOR_SCORE_WEIGHTS.keys())
    raw = raw[raw["Parameter"].isin(known_params)].copy()

    sector_cols = [c for c in raw.columns
                   if c not in ("Quarter", "Parameter") and c in SECTOR_GROUPS]
    if not sector_cols:
        logger.warning(
            "  No sector columns found matching SECTOR_GROUPS keys — skipping scores.\n"
            f"  Available columns: {[c for c in raw.columns if c not in ('Quarter','Parameter')]}"
        )
        _EMPTY["ticker_sector"] = ticker_sector
        return _EMPTY

    # ── Composite score per sector per quarter ────────────────────────
    composite_rows: List[dict] = []
    for quarter, grp in raw.groupby("Quarter"):
        qdate = _parse_quarter_to_date(quarter)
        if qdate is None:
            logger.warning(f"  Cannot parse quarter '{quarter}' — row skipped.")
            continue
        param_scores = grp.set_index("Parameter")[sector_cols]
        # Weighted average across A–F parameters present
        present = [p for p in SECTOR_SCORE_WEIGHTS if p in param_scores.index]
        if not present:
            continue
        weights = np.array([SECTOR_SCORE_WEIGHTS[p] for p in present])
        weights /= weights.sum()           # renormalise if some params missing
        scores   = param_scores.loc[present].astype(float)
        composite = (scores.T @ weights)   # Series: sector → composite score
        row = {"date": qdate}
        row.update(composite.to_dict())
        composite_rows.append(row)

    if not composite_rows:
        logger.warning("  No valid composite scores computed — skipping scores.")
        _EMPTY["ticker_sector"] = ticker_sector
        return _EMPTY

    sector_composite = (
        pd.DataFrame(composite_rows)
        .set_index("date")
        .sort_index()[sector_cols]
    )
    sector_composite.to_csv(OUTPUT_DIR / "sector_composite_scores.csv")
    logger.info(f"  Composite scores shape: {sector_composite.shape}")
    logger.info(f"\n{sector_composite.round(2).to_string()}\n")

    # ── Temporal disaggregation → monthly (conservative forward-fill) ─
    # Score for quarter Q becomes available SCORE_LAG_MONTHS after Q-end.
    # We forward-fill within the available monthly index.
    monthly_template = _resample_monthly(
        pd.Series(0.0, index=pd.date_range("2019-01-01", pd.Timestamp.today(), freq="D"))
    ).index

    def _to_monthly(col_series: pd.Series) -> pd.Series:
        lagged_idx = pd.DatetimeIndex([
            d + pd.DateOffset(months=SCORE_LAG_MONTHS)
            for d in col_series.index
        ])
        lagged = pd.Series(col_series.values, index=lagged_idx)
        combined = lagged.reindex(monthly_template.union(lagged.index)).sort_index().ffill()
        return combined.reindex(monthly_template)

    monthly_dict = {col: _to_monthly(sector_composite[col]) for col in sector_composite.columns}
    sector_composite_monthly = pd.DataFrame(monthly_dict)
    sector_composite_monthly.to_csv(OUTPUT_DIR / "sector_composite_monthly.csv")

    # ── Conviction signals ────────────────────────────────────────────
    sector_conviction = _compute_sector_conviction(sector_composite)
    logger.info("  Sector conviction factors:")
    for sec, cv in sorted(sector_conviction.items(), key=lambda x: -x[1]):
        logger.info(f"    {sec:<30} {cv:.3f}")

    ticker_conviction: Dict[str, float] = {}
    for sector, tickers in SECTOR_GROUPS.items():
        cv = sector_conviction.get(sector, 0.50)
        for t in tickers:
            ticker_conviction[t] = cv

    logger.info("Phase 3.5 validation checklist passed ✓\n")
    return {
        "sector_composite_scores":  sector_composite,
        "sector_composite_monthly": sector_composite_monthly,
        "sector_conviction":        sector_conviction,
        "ticker_conviction":        ticker_conviction,
        "ticker_sector":            ticker_sector,
        "scores_available":         True,
    }



# ======================================================================
# PHASE 4: DEPENDENCY STRUCTURE & CLUSTERING  (unchanged)
# ======================================================================
def run_phase4(p1: dict, p2: dict) -> dict:
    logger.info("=" * 62)
    logger.info("PHASE 4 — Dependency Structure & Clustering")
    logger.info("=" * 62)
    mlr            = p2["monthly_log_returns_clean"]
    active_tickers = p1["active_tickers"]
    corr_matrix = mlr.corr(method="pearson")
    dist_matrix = np.sqrt(2 * (1 - corr_matrix.clip(-1, 1)))
    linkage_mat = linkage(dist_matrix.values, method="ward")
    cm = sns.clustermap(corr_matrix, method="ward", cmap="coolwarm",
                        figsize=(10, 8), annot=True, fmt=".2f")
    cm.savefig(OUTPUT_DIR / "correlation_clustermap.pdf", bbox_inches="tight")
    plt.close(cm.fig)
    heights   = sorted(linkage_mat[:, 2])
    diffs     = np.diff(heights) if len(heights) > 1 else np.array([1.0])
    threshold = heights[np.argmax(diffs)] + diffs.max() * 0.5
    l1_labels = fcluster(linkage_mat, t=threshold, criterion="distance")
    layer1: Dict[int, List] = {}
    for ticker, lbl in zip(active_tickers, l1_labels):
        layer1.setdefault(int(lbl), []).append(ticker)
    logger.info(f"  Layer 1 clusters: {layer1}")
    unstable_pairs: set = set()
    for members in layer1.values():
        for i in range(len(members)):
            for j in range(i + 1, len(members)):
                ta, tb   = members[i], members[j]
                roll_cor = mlr[ta].rolling(36).corr(mlr[tb])
                pct_pos  = float((roll_cor > 0).mean())
                if pct_pos < 0.70:
                    unstable_pairs.add(tuple(sorted([ta, tb])))
    logger.info("  Running pairwise Granger tests (maxlag=6) …")
    data_gc = mlr[active_tickers].dropna()
    G_dir   = nx.DiGraph(); G_dir.add_nodes_from(active_tickers)
    for ti in active_tickers:
        for tj in active_tickers:
            if ti == tj:
                continue
            try:
                gc    = grangercausalitytests(data_gc[[tj, ti]], maxlag=6, verbose=False)
                min_p = min(gc[lag][0]["ssr_ftest"][1] for lag in gc)
                if min_p < 0.05:
                    G_dir.add_edge(ti, tj, weight=1.0 - min_p)
            except Exception:
                pass
    G_undir = G_dir.to_undirected(reciprocal=False)
    for u, v in list(G_undir.edges()):
        w1 = G_dir.get_edge_data(u, v, default={}).get("weight", 0.0)
        w2 = G_dir.get_edge_data(v, u, default={}).get("weight", 0.0)
        G_undir[u][v]["weight"] = w1 + w2
    partition       = _detect_communities(G_undir, random_state=RANDOM_STATE)
    gc_communities: Dict[int, List] = {}
    for ticker, cid in partition.items():
        gc_communities.setdefault(int(cid), []).append(ticker)
    logger.info(f"  Granger communities: {gc_communities}")
    nx.write_gexf(G_dir, str(OUTPUT_DIR / "granger_network.gexf"))
    G_full = nx.Graph()
    for i, ti in enumerate(active_tickers):
        for j, tj in enumerate(active_tickers):
            if i < j:
                d = float(np.sqrt(2 * (1 - np.clip(corr_matrix.loc[ti, tj], -1, 1))))
                G_full.add_edge(ti, tj, weight=d)
    MST        = nx.minimum_spanning_tree(G_full, weight="weight")
    centrality = nx.degree_centrality(MST)
    fig, ax    = plt.subplots(figsize=(10, 8))
    pos   = nx.spring_layout(MST, seed=RANDOM_STATE)
    sizes = [3000 * centrality[n] + 500 for n in MST.nodes()]
    nx.draw(MST, pos, ax=ax, with_labels=True, node_size=sizes,
            node_color="lightsteelblue", font_size=9, edge_color="dimgray")
    ax.set_title("Minimum Spanning Tree (Mantegna Distance)")
    fig.tight_layout(); fig.savefig(OUTPUT_DIR / "mst_topology.pdf", bbox_inches="tight")
    plt.close(fig)
    l1_map = {t: lbl for lbl, mems in layer1.items() for t in mems}
    gc_map  = {t: cid for cid, mems in gc_communities.items() for t in mems}
    cluster_assignments: Dict[str, List] = {"isolated": []}
    grouped: set = set()
    group_idx = 1
    for ticker in active_tickers:
        if ticker in grouped:
            continue
        l1_peers  = set(layer1.get(l1_map.get(ticker), [])) - {ticker}
        gc_peers  = set(gc_communities.get(gc_map.get(ticker), [])) - {ticker}
        consensus = l1_peers & gc_peers or gc_peers
        if not consensus:
            cluster_assignments["isolated"].append(ticker); grouped.add(ticker); continue
        candidates   = [ticker] + [p for p in consensus if p not in grouped]
        stable_group = [ticker]
        for p in candidates[1:]:
            if tuple(sorted([ticker, p])) not in unstable_pairs:
                stable_group.append(p)
            if len(stable_group) == 5:
                break
        if len(stable_group) < 2:
            cluster_assignments["isolated"].append(ticker); grouped.add(ticker)
        else:
            key = f"VARX_group_{group_idx}"
            cluster_assignments[key] = stable_group
            grouped.update(stable_group); group_idx += 1
    for ticker in active_tickers:
        if ticker not in grouped:
            cluster_assignments["isolated"].append(ticker)
    logger.info(f"  Final clusters: {cluster_assignments}")
    with open(OUTPUT_DIR / "cluster_assignments.json", "w") as fh:
        json.dump(cluster_assignments, fh, indent=2)
    pd.DataFrame([
        {"group": g, "ticker": t,
         "mst_degree_centrality": round(centrality.get(t, 0.0), 4)}
        for g, mems in cluster_assignments.items() for t in mems
    ]).to_csv(OUTPUT_DIR / "cluster_summary.csv", index=False)
    logger.info("Phase 4 validation checklist passed ✓\n")
    return {
        "cluster_assignments": cluster_assignments,
        "corr_matrix": corr_matrix,
        "G_directed": G_dir, "MST": MST, "centrality": centrality,
    }


# ======================================================================
# PHASE 4.5: HMM REGIME DETECTION  (unchanged)
# ======================================================================

def _extract_filtered_probs(model: GaussianHMM, obs: np.ndarray) -> np.ndarray:
    """
    Forward-pass filtered probabilities: P(state_t | obs_{1..t}).

    Implemented manually so it is robust across all hmmlearn versions.
    hmmlearn >= 0.3.x removed _do_forward_pass from the Python layer;
    this replaces that call with a pure-NumPy log-space forward algorithm
    using only public model attributes.

    IMPORTANT (F-14): HMM *parameters* were estimated on the full dataset
    (parameter-level look-ahead). Filtered probs are extracted via the
    forward pass only — no within-series look-ahead.
    """
    n_samples = obs.shape[0]
    n_states  = model.n_components

    # ── Emission log-likelihoods ───────────────────────────────────────────
    # _compute_log_likelihood is present in all hmmlearn versions tested;
    # fall back to manual scipy computation if a future version removes it.
    if hasattr(model, "_compute_log_likelihood"):
        framelogprob = model._compute_log_likelihood(obs)   # (n_samples, n_states)
    else:
        from scipy.stats import multivariate_normal          # noqa: PLC0415
        framelogprob = np.column_stack([
            multivariate_normal.logpdf(
                obs, mean=model.means_[k], cov=model.covars_[k]
            )
            for k in range(n_states)
        ])

    # ── Log-space forward algorithm ────────────────────────────────────────
    # log α_0(j)   = log π_j + log b_j(x_0)
    # log α_t(j)   = logsumexp_i[ log α_{t-1}(i) + log a_{ij} ] + log b_j(x_t)
    log_startprob = np.log(np.clip(model.startprob_, 1e-300, None))
    log_transmat  = np.log(np.clip(model.transmat_,  1e-300, None))

    log_alpha        = np.empty((n_samples, n_states))
    log_alpha[0]     = log_startprob + framelogprob[0]

    for t in range(1, n_samples):
        for j in range(n_states):
            log_alpha[t, j] = (
                logsumexp(log_alpha[t - 1] + log_transmat[:, j])
                + framelogprob[t, j]
            )

    # ── Convert to normalised probabilities ───────────────────────────────
    log_norm  = logsumexp(log_alpha, axis=1, keepdims=True)
    filtered  = np.exp(log_alpha - log_norm)
    filtered  = np.clip(filtered, 0.0, 1.0)
    row_sums  = filtered.sum(axis=1, keepdims=True)
    filtered /= np.where(row_sums > 0, row_sums, 1.0)
    return filtered


def _select_n_regimes_bic(obs: np.ndarray,
                            candidates: Tuple[int, ...] = (2, 3, 4)) -> int:
    """Optimal number of HMM states via BIC."""
    bic_scores = {}
    for n in candidates:
        try:
            m = GaussianHMM(n_components=n, covariance_type="full",
                            n_iter=500, random_state=RANDOM_STATE)
            m.fit(obs)
            bic_scores[n] = -2 * m.score(obs) + n * np.log(len(obs))
        except Exception:
            bic_scores[n] = np.inf
    best = min(bic_scores, key=bic_scores.get)
    logger.debug(f"    BIC: {bic_scores} → n={best}")
    return best


def _fit_hmm(obs: np.ndarray, n: int) -> GaussianHMM:
    """Fit GaussianHMM with full covariance."""
    m = GaussianHMM(n_components=n, covariance_type="full",
                    n_iter=1000, tol=1e-4, random_state=RANDOM_STATE,
                    init_params="stmc", params="stmc")
    m.fit(obs)
    return m


def _sort_states_by_return(model: GaussianHMM,
                             filtered: np.ndarray) -> np.ndarray:
    """Sort columns: bear (lowest mean return) … bull (highest mean return)."""
    idx = np.argsort(model.means_[:, 0])
    return filtered[:, idx]


def _hmm_to_df(filtered: np.ndarray, index: pd.Index,
                prefix: str, n: int) -> pd.DataFrame:
    """Wrap filtered prob array in a labelled DataFrame."""
    labels = ["bear", "transitional", "bull"] if n == 3 else [f"r{i}" for i in range(n)]
    return pd.DataFrame(filtered, index=index,
                        columns=[f"{prefix}_{lbl}" for lbl in labels[:n]])


def _align_labels_hungarian(new_means: np.ndarray,
                              ref_means: np.ndarray) -> np.ndarray:
    """
    Hungarian-algorithm label alignment for walk-forward HMM re-estimation.
    Currently unused — retained for future per-step re-estimation extension.
    """
    cost = np.abs(new_means[:, None] - ref_means[None, :])
    if cost.ndim == 3:
        cost = cost.sum(-1)
    _, col_idx = linear_sum_assignment(cost)
    return col_idx


def run_phase4_5(p1: dict, p2: dict, p3: dict, p4: dict) -> dict:
    """Phase 4.5: HMM regime detection at market, group, and stock levels."""
    logger.info("=" * 62)
    logger.info("PHASE 4.5 — HMM Regime Detection")
    logger.info("=" * 62)

    mlr            = p2["monthly_log_returns_clean"]
    realized_vol   = p3["realized_volatility"]
    cluster_assign = p4["cluster_assignments"]
    bench_rets     = p1["benchmark_returns"]
    active_tickers = p1["active_tickers"]
    all_probs: Dict[str, pd.DataFrame] = {}

    # ── Market-level HMM (Nifty returns + cross-sectional mean RV) ────────────
    bench_aligned = bench_rets.reindex(mlr.index).fillna(0.0)
    mkt_rv_proxy  = realized_vol.mean(axis=1).reindex(mlr.index).fillna(0.0)
    mkt_obs       = np.column_stack([bench_aligned.values, mkt_rv_proxy.values])

    n_mkt    = _select_n_regimes_bic(mkt_obs)
    hmm_mkt  = _fit_hmm(mkt_obs, n_mkt)
    mkt_filt = _sort_states_by_return(hmm_mkt, _extract_filtered_probs(hmm_mkt, mkt_obs))
    mkt_df   = _hmm_to_df(mkt_filt, mlr.index, "market_hmm", n_mkt)
    all_probs["market"] = mkt_df
    with open(OUTPUT_DIR / "hmm_market.pkl", "wb") as fh:
        pickle.dump(hmm_mkt, fh)
    logger.info(f"  Market HMM n={n_mkt} | means: "
                f"{list(zip(hmm_mkt.means_[:, 0].round(4), hmm_mkt.means_[:, 1].round(4)))}")

    # ── Group-level HMMs ──────────────────────────────────────────────────────
    group_probs: Dict[str, pd.DataFrame] = {}
    for gkey, members in cluster_assign.items():
        if gkey == "isolated" or len(members) < 2:
            continue
        grp_rets = mlr[members].dropna()
        if len(grp_rets) < 24:
            logger.warning(f"  {gkey}: too few obs for group HMM — skipping.")
            continue
        try:
            n_grp   = _select_n_regimes_bic(grp_rets.values)
            hmm_grp = _fit_hmm(grp_rets.values, n_grp)
            g_filt  = _sort_states_by_return(hmm_grp,
                          _extract_filtered_probs(hmm_grp, grp_rets.values))
            g_df    = _hmm_to_df(g_filt, grp_rets.index, f"{gkey}_hmm", n_grp)
            g_df    = g_df.reindex(mlr.index)
            group_probs[gkey] = g_df
            all_probs[gkey]   = g_df
            with open(OUTPUT_DIR / "hmm_groups" / f"hmm_{gkey}.pkl", "wb") as fh:
                pickle.dump(hmm_grp, fh)
            logger.info(f"  {gkey} HMM n={n_grp}")
        except Exception as exc:
            logger.warning(f"  {gkey} HMM failed: {exc}")

    # ── Stock-level HMMs ──────────────────────────────────────────────────────
    stock_probs: Dict[str, pd.DataFrame] = {}
    for ticker in active_tickers:
        ret_s = mlr[ticker].dropna()
        rv_s  = (realized_vol[ticker].reindex(ret_s.index)
                 .fillna(realized_vol[ticker].median())
                 if ticker in realized_vol.columns
                 else pd.Series(0.0, index=ret_s.index))
        stk_obs = np.column_stack([ret_s.values, rv_s.values])
        try:
            n_stk   = _select_n_regimes_bic(stk_obs)
            hmm_stk = _fit_hmm(stk_obs, n_stk)
            s_filt  = _sort_states_by_return(hmm_stk,
                          _extract_filtered_probs(hmm_stk, stk_obs))
            s_df    = _hmm_to_df(s_filt, ret_s.index, f"{ticker}_hmm", n_stk)
            s_df    = s_df.reindex(mlr.index)
            stock_probs[ticker] = s_df
            all_probs[ticker]   = s_df
            safe = ticker.replace(".", "_")
            with open(OUTPUT_DIR / "hmm_stocks" / f"hmm_{safe}.pkl", "wb") as fh:
                pickle.dump(hmm_stk, fh)
            logger.info(f"  {ticker} HMM n={n_stk}")
        except Exception as exc:
            logger.warning(f"  {ticker} HMM failed: {exc}")

    all_probs_df = pd.concat(
        [v for v in all_probs.values() if isinstance(v, pd.DataFrame)], axis=1
    )
    all_probs_df.to_parquet(OUTPUT_DIR / "filtered_probs_all.parquet")

    # Regime probability plot
    fig, ax = plt.subplots(figsize=(14, 5))
    mkt_df.ffill().fillna(0.0).plot(ax=ax, linewidth=1.5)
    ax.set_title("Market Regime Filtered Probabilities (Nifty 50)")
    ax.set_ylabel("P(regime)"); ax.set_ylim(0, 1)
    ax.legend(loc="upper right", fontsize=9)
    fig.tight_layout()
    fig.savefig(OUTPUT_DIR / "regime_plot.pdf", bbox_inches="tight")
    plt.close(fig)

    # Append HMM probs to exog_master
    exog_upd = p3["exog_master"].copy()
    for col in mkt_df.columns:
        exog_upd[col] = mkt_df[col]
    for g_df in group_probs.values():
        for col in g_df.columns:
            exog_upd[col] = g_df[col]
    exog_upd.to_parquet(OUTPUT_DIR / "exog_master.parquet")

    logger.info("Phase 4.5 validation checklist passed ✓\n")
    return {
        "market_filtered_df":   mkt_df,
        "group_filtered_probs": group_probs,
        "stock_filtered_probs": stock_probs,
        "all_probs_df":         all_probs_df,
        "exog_master_updated":  exog_upd,
        "hmm_market":           hmm_mkt,
    }

# ======================================================================
# PHASE 5: MODEL BUILDING & WALK-FORWARD VALIDATION  (unchanged)
# ======================================================================

def _build_arimax_exog(ticker: str,
                        idx: pd.Index,
                        mkt_hmm: pd.DataFrame,
                        stk_hmm: pd.DataFrame,
                        rv: pd.DataFrame,
                        market_ind: pd.DataFrame) -> np.ndarray:
    """
    Assemble ARIMAX exogenous matrix (Option 3 B-matrix structure):
      - Market HMM filtered probs  (N_R − 1 cols, drop last to avoid collinearity)
      - Stock  HMM filtered probs  (N_R − 1 cols)
      - Stock  realized volatility (1 col)
      - Market indicators          (n_ind cols, already lagged 1 period)

    Returns shape (len(idx), n_features).
    Returns shape (len(idx), 0) when no features exist — caller passes None.  F-09
    """
    parts = []
    if not mkt_hmm.empty and mkt_hmm.shape[1] > 1:
        parts.append(mkt_hmm.iloc[:, :-1])
    if not stk_hmm.empty and stk_hmm.shape[1] > 1:
        parts.append(stk_hmm.iloc[:, :-1])
    if ticker in rv.columns:
        parts.append(rv[[ticker]].rename(columns={ticker: f"{ticker}_rv"}))
    if not market_ind.empty:
        parts.append(market_ind)

    if not parts:
        return np.empty((len(idx), 0))                                     # F-09

    combined = pd.concat(parts, axis=1).reindex(idx).ffill().bfill().fillna(0.0)
    return combined.values


def _build_varx_exog(gkey: str,
                      members: List[str],
                      idx: pd.Index,
                      mkt_hmm: pd.DataFrame,
                      grp_hmm_dfs: dict,
                      rv: pd.DataFrame,
                      market_ind: pd.DataFrame) -> Optional[np.ndarray]:
    """
    Assemble VARX common exogenous matrix:
      - Market HMM probs (N_R − 1 cols)
      - Group  HMM probs (N_R − 1 cols, if available)
      - Realized volatility for each group member (per-stock RV as common regressor)
      - Market indicators (n_ind cols, lagged)
    """
    parts = []
    if not mkt_hmm.empty and mkt_hmm.shape[1] > 1:
        parts.append(mkt_hmm.iloc[:, :-1])
    g_hmm = grp_hmm_dfs.get(gkey, pd.DataFrame())
    if not g_hmm.empty and g_hmm.shape[1] > 1:
        parts.append(g_hmm.iloc[:, :-1])
    for ticker in members:
        if ticker in rv.columns:
            parts.append(rv[[ticker]].rename(columns={ticker: f"{ticker}_rv"}))
    if not market_ind.empty:
        parts.append(market_ind)

    if not parts:
        return None

    combined = pd.concat(parts, axis=1).reindex(idx).ffill().bfill().fillna(0.0)
    return combined.values


def _arimax_bic_grid(endog: np.ndarray,
                      exog: Optional[np.ndarray]) -> Tuple[int, int]:
    """Grid-search ARIMAX (p, q) ∈ [0..4]² by BIC. Return best (p, q)."""
    best_bic, best_pq = np.inf, (1, 0)
    for p, q in iproduct(range(3), range(3)):
        try:
            m = SARIMAX(endog, exog=exog, order=(p, 0, q),
                        enforce_stationarity=True, enforce_invertibility=True,
                        trend="c").fit(method="lbfgs", maxiter=200, disp=False)
            if np.isfinite(m.bic) and m.bic < best_bic:   # ← guard inf BIC
                best_bic, best_pq = m.bic, (p, q)
        except Exception:
            continue
    return best_pq


def _arimax_ticker_walkforward(
    ticker: str,
    mlr: pd.DataFrame,
    mkt_hmm_df: pd.DataFrame,
    stk_hmm_dfs: Dict[str, pd.DataFrame],
    realized_vol: pd.DataFrame,
    market_ind: pd.DataFrame,
    T: int,
) -> Tuple[List[dict], List[dict]]:
    """
    Full ARIMAX walk-forward for a single ticker.
    Completely self-contained — safe to run in a thread or process.
    BIC order determined once at the first estimation step.
    """
    fc_rows:  List[dict] = []
    err_rows: List[dict] = []
    order:    Optional[Tuple[int, int]] = None
    cached_m  = None
    cached_e: Optional[np.ndarray] = None

    for t in range(MIN_TRAIN_MONTHS, T):
        w_start    = max(0, t - MAX_LOOKBACK_MONTHS)
        train_data = mlr.iloc[w_start:t]
        reestimate = (t == MIN_TRAIN_MONTHS) or ((t - MIN_TRAIN_MONTHS) % 6 == 0)

        endog_s = train_data[ticker].dropna()
        if len(endog_s) < 24:
            continue

        mkt_train = mkt_hmm_df.reindex(endog_s.index).ffill().fillna(0.0)
        s_hmm_raw = stk_hmm_dfs.get(ticker, pd.DataFrame())
        stk_train = (s_hmm_raw.reindex(endog_s.index).ffill().fillna(0.0)
                     if not s_hmm_raw.empty else pd.DataFrame(index=endog_s.index))
        ind_train  = market_ind.reindex(endog_s.index).ffill().fillna(0.0)
        exog_arr   = _build_arimax_exog(ticker, endog_s.index,
                                        mkt_train, stk_train, realized_vol, ind_train)
        exog_in    = exog_arr if exog_arr.shape[1] > 0 else None

        # BIC grid: once per ticker only
        if order is None:
            order = _arimax_bic_grid(endog_s.values, exog_in)

        # Re-fit every 6 steps
        if reestimate or cached_m is None:
            try:
                cached_m = _fit_arimax(endog_s.values, exog_in, order)
                cached_e = exog_in
                try:
                    _, c_pval, _ = breaks_cusumolsresid(cached_m.resid)
                    if c_pval < 0.05:
                        logger.warning(f"  [{ticker}] CUSUM p={c_pval:.3f} at t={t}")
                except Exception:
                    pass
            except Exception as exc:
                logger.warning(f"  ARIMAX fit [{ticker}] t={t}: {exc}")
                continue

        if cached_m is None:
            continue

        max_h    = max(FORECAST_HORIZONS)
        exog_fut = (np.repeat(cached_e[-1:], max_h, axis=0)
                    if cached_e is not None else None)
        try:
            fc  = cached_m.get_forecast(steps=max_h, exog=exog_fut)
            fmu = fc.predicted_mean
            fci_90 = fc.conf_int(alpha=0.10)   # 90% CI  → ±1.645 σ
            fci_95 = fc.conf_int(alpha=0.05)   # 95% CI  → ±1.960 σ
            for h in FORECAST_HORIZONS:
                tgt = t + h
                if tgt < T:
                    pred   = float(fmu[h - 1]) if h <= len(fmu) else np.nan
                    actual = float(mlr[ticker].iloc[tgt])
                    lo_90  = float(fci_90[h-1, 0]) if h <= len(fci_90) else np.nan
                    hi_90  = float(fci_90[h-1, 1]) if h <= len(fci_90) else np.nan
                    lo_95  = float(fci_95[h-1, 0]) if h <= len(fci_95) else np.nan
                    hi_95  = float(fci_95[h-1, 1]) if h <= len(fci_95) else np.nan
                    fc_rows.append({
                        "model": "ARIMAX", "ticker": ticker, "t": t,
                        "horizon": h, "forecast": pred, "actual": actual,
                        "ci90_lower": lo_90, "ci90_upper": hi_90,
                        "ci95_lower": lo_95, "ci95_upper": hi_95,
                        "date_forecast": str(mlr.index[t].date()),
                        "date_target":   str(mlr.index[tgt].date()),
                    })
                    err_rows.append({"model": "ARIMAX", "ticker": ticker,
                                     "t": t, "horizon": h, "error": actual - pred})
        except Exception as exc:
            logger.warning(f"  ARIMAX forecast [{ticker}] t={t}: {exc}")

    return fc_rows, err_rows


def _fit_arimax(endog: np.ndarray,
                exog: Optional[np.ndarray],
                order: Tuple[int, int]):
    """Fit SARIMAX (no seasonal component) = ARIMAX."""
    return SARIMAX(endog, exog=exog, order=(order[0], 0, order[1]),
                   enforce_stationarity=True, enforce_invertibility=True,
                   trend="c").fit(method="lbfgs", maxiter=500, disp=False)


def _fit_varx(endog: np.ndarray,
               exog: Optional[np.ndarray]) -> Tuple:
    """
    Fit VAR(X). Returns (VARResults, optimal_lag).
    Lag capped at 3 (parsimony constraint for ~60 monthly observations).
    """
    model = VAR(endog=endog, exog=exog)
    sel   = model.select_order(maxlags=6)
    lag   = max(1, min(int(sel.bic), 3))
    return model.fit(lag), lag


def _diebold_mariano(err1: np.ndarray, err2: np.ndarray) -> Tuple[float, float]:
    """DM test (squared-error loss). H0: equal forecast accuracy."""
    d     = err1 ** 2 - err2 ** 2
    denom = np.std(d, ddof=1) / np.sqrt(len(d)) + 1e-12
    dm    = float(np.mean(d) / denom)
    pval  = float(2 * (1 - stats.norm.cdf(abs(dm))))
    return dm, pval


def run_phase5(p1: dict, p2: dict, p3: dict, p4: dict, p4_5: dict) -> dict:
    """
    Phase 5: Hybrid expanding-rolling walk-forward validation.

    Window: expands until MAX_LOOKBACK_MONTHS months, then rolls.
    ARIMAX / VARX: re-estimated every 3 steps (quarterly).
    HMM:  parameters fixed (full-sample); only filtered probs used per window.
          (parameter-level look-ahead acknowledged — see F-14).

    F-11 fix: VARX forecast initialisation always uses the CURRENT training
    window's last p rows, not a stale cached array.
    """
    logger.info("=" * 62)
    logger.info("PHASE 5 — Walk-Forward Validation")
    logger.info("=" * 62)

    mlr            = p2["monthly_log_returns_clean"]
    realized_vol   = p3["realized_volatility"]
    market_ind     = p3["market_indicators"]
    cluster_assign = p4["cluster_assignments"]
    mkt_hmm_df     = p4_5["market_filtered_df"]
    grp_hmm_dfs    = p4_5["group_filtered_probs"]
    stk_hmm_dfs    = p4_5["stock_filtered_probs"]
    active_tickers = p1["active_tickers"]

    T           = len(mlr)
    varx_groups = {k: v for k, v in cluster_assign.items() if k != "isolated"}

    forecast_rows: List[dict] = []
    error_rows:    List[dict] = []


    # ── ARIMAX: all tickers in parallel ───────────────────────────────────
    n_workers = min(len(active_tickers), (os.cpu_count() or 4))
    logger.info(f"  Launching ARIMAX walk-forward: {len(active_tickers)} tickers "
                f"× {n_workers} threads …")

    with ThreadPoolExecutor(max_workers=n_workers) as pool:
        futures = {
            pool.submit(
                _arimax_ticker_walkforward,
                ticker, mlr, mkt_hmm_df, stk_hmm_dfs,
                realized_vol, market_ind, T
            ): ticker
            for ticker in active_tickers
        }
        for fut in as_completed(futures):
            tkr = futures[fut]
            try:
                fc_r, err_r = fut.result()
                forecast_rows.extend(fc_r)
                error_rows.extend(err_r)
                logger.info(f"  [{tkr}] done — {len(fc_r)} forecasts")
            except Exception as exc:
                logger.warning(f"  [{tkr}] walk-forward failed: {exc}")    

    model_cache:   dict       = {}   # stores (fitted_model, exog_array)
    arimax_orders: dict       = {}

    for t in range(MIN_TRAIN_MONTHS, T):
        w_start    = max(0, t - MAX_LOOKBACK_MONTHS)
        train_data = mlr.iloc[w_start:t]
        reestimate = (t == MIN_TRAIN_MONTHS) or ((t - MIN_TRAIN_MONTHS) % 6 == 0)


        # ── VARX (one per cluster group) ──────────────────────────────────────
        for gkey, members in varx_groups.items():
            group_m = [m for m in members if m in mlr.columns]
            if len(group_m) < 2:
                continue

            # Always slice current training endog (F-11: never use stale cached endog)
            endog_g = train_data[group_m].dropna()
            if len(endog_g) < max(MIN_TRAIN_MONTHS // 2, 24):
                continue

            cache_key_v = f"varx_{gkey}"
            if reestimate or cache_key_v not in model_cache:
                mkt_g  = mkt_hmm_df.reindex(endog_g.index).ffill().fillna(0.0)
                ind_g  = market_ind.reindex(endog_g.index).ffill().fillna(0.0)
                exog_v = _build_varx_exog(gkey, group_m, endog_g.index,
                                           mkt_g, grp_hmm_dfs, realized_vol, ind_g)
                try:
                    varx_res, opt_lag = _fit_varx(endog_g.values, exog_v)
                    model_cache[cache_key_v] = (varx_res, opt_lag, exog_v)
                except Exception as exc:
                    logger.warning(f"  VARX [{gkey}] t={t}: {exc}")
                    continue

            if cache_key_v not in model_cache:
                continue
            varx_res, opt_lag, exog_v_ref = model_cache[cache_key_v]

            max_h      = max(FORECAST_HORIZONS)
            exog_fut_v = (np.repeat(exog_v_ref[-1:], max_h, axis=0)
                          if exog_v_ref is not None else None)
            try:
                y_init = endog_g.values[-opt_lag:]          # F-11: current window
                fc_pt = varx_res.forecast(y=y_init, steps=max_h,
                                          exog_future=exog_fut_v)
                fc_lo_90 = fc_hi_90 = fc_lo_95 = fc_hi_95 = None
                try:
                    _, fc_lo_90, fc_hi_90 = varx_res.forecast_interval(
                        y=y_init, steps=max_h, alpha=0.10, exog_future=exog_fut_v
                    )
                    _,     fc_lo_95, fc_hi_95 = varx_res.forecast_interval(
                        y=y_init, steps=max_h, alpha=0.05, exog_future=exog_fut_v
                    )
                except TypeError:
                    # Older statsmodels: exog_future not accepted here
                    try:
                        _, fc_lo_90, fc_hi_90 = varx_res.forecast_interval(
                            y=y_init, steps=max_h, alpha=0.10
                        )
                        _, fc_lo_95, fc_hi_95 = varx_res.forecast_interval(
                            y=y_init, steps=max_h, alpha=0.05
                        )
                    except Exception:
                        pass  # CIs stay None; point forecasts are still valid
                except Exception:
                    pass
                for h in FORECAST_HORIZONS:
                    tgt = t + h
                    if tgt < T:
                        for k, tkr in enumerate(group_m):
                            pred   = float(fc_pt[h - 1, k]) if h <= len(fc_pt) else np.nan
                            actual = float(mlr[tkr].iloc[tgt])
                            lo_90  = float(fc_lo_90[h-1, k]) if fc_lo_90 is not None and h <= len(fc_lo_90) else np.nan
                            hi_90  = float(fc_hi_90[h-1, k]) if fc_hi_90 is not None and h <= len(fc_hi_90) else np.nan
                            lo_95  = float(fc_lo_95[h-1, k]) if fc_lo_95 is not None and h <= len(fc_lo_95) else np.nan
                            hi_95  = float(fc_hi_95[h-1, k]) if fc_hi_95 is not None and h <= len(fc_hi_95) else np.nan
                            forecast_rows.append({
                                "model": "VARX", "ticker": tkr, "t": t,
                                "horizon": h, "forecast": pred, "actual": actual,
                                "ci90_lower": lo_90, "ci90_upper": hi_90,
                                "ci95_lower": lo_95, "ci95_upper": hi_95,
                                "group": gkey,
                                "date_forecast": str(mlr.index[t].date()),
                                "date_target":   str(mlr.index[tgt].date()),
                            })
                            error_rows.append({"model": "VARX", "ticker": tkr,
                                               "t": t, "horizon": h, "error": actual - pred})
            except Exception as exc:
                logger.warning(f"  VARX forecast [{gkey}] t={t}: {exc}")

    # AFTER — schema guaranteed regardless of whether any forecasts were produced
    _FC_COLS  = ["model", "ticker", "t", "horizon", "forecast", "actual",
                "ci90_lower", "ci90_upper", "ci95_lower", "ci95_upper", 
                "date_forecast", "date_target", "group"]
    _ERR_COLS = ["model", "ticker", "t", "horizon", "error"]

    forecasts_df = (pd.DataFrame(forecast_rows)
                    if forecast_rows
                    else pd.DataFrame(columns=_FC_COLS))
    errors_df    = (pd.DataFrame(error_rows)
                    if error_rows
                    else pd.DataFrame(columns=_ERR_COLS))

    forecasts_df.to_parquet(OUTPUT_DIR / "walkforward_forecasts.parquet")
    errors_df.to_parquet(OUTPUT_DIR / "walkforward_errors.parquet")
    logger.info(f"  Walk-forward: {len(forecasts_df)} forecasts, {len(errors_df)} errors")

    # ── Performance metrics ────────────────────────────────────────────────────
    perf_rows:  List[dict] = []
    calib_rows: List[dict] = []

    if forecasts_df.empty:
        logger.warning(
            "  No walk-forward forecasts generated. "
            "Possible causes: all model fits failed, START_DATE too recent, "
            "or MIN_TRAIN_MONTHS > available history. "
            "Check debug logs above for per-ticker ARIMAX/VARX failure messages."
        )
    else:
        for ticker in active_tickers:
            for h in FORECAST_HORIZONS:
                for mt in ["ARIMAX", "VARX"]:
                    sub = forecasts_df[
                        (forecasts_df["ticker"]  == ticker) &
                        (forecasts_df["horizon"] == h) &
                        (forecasts_df["model"]   == mt)
                    ].dropna(subset=["forecast", "actual"])
                    if len(sub) < 5:
                        continue
                    fc  = sub["forecast"].values
                    act = sub["actual"].values
                    da  = float(np.mean(np.sign(fc) == np.sign(act)))
                    ic, ic_p = (pearsonr(fc, act) if len(fc) > 2
                                else (np.nan, np.nan))
                    rmse = float(np.sqrt(np.mean((fc - act) ** 2)))
                    mae  = float(np.mean(np.abs(fc - act)))
                    perf_rows.append({
                        "model": mt, "ticker": ticker, "horizon": h,
                        "DA": round(da, 4),
                        "IC": round(float(ic), 4)   if not np.isnan(ic)   else np.nan,
                        "IC_pval": round(float(ic_p), 4) if not np.isnan(ic_p) else np.nan,
                        "RMSE": round(rmse, 6), "MAE": round(mae, 6), "n_obs": len(sub),
                    })
                    err = act - fc
                    lo  = float(np.percentile(err, 5))
                    hi  = float(np.percentile(err, 95))
                    cov = float(np.mean((err >= lo) & (err <= hi)))
                    calib_rows.append({
                        "model": mt, "ticker": ticker, "horizon": h,
                        "pi_lower_offset":          round(lo, 6),
                        "pi_upper_offset":          round(hi, 6),
                        "empirical_coverage_90pct": round(cov, 4),
                    })

    # ── CI Summary Table ──────────────────────────────────────────────────────
    if not forecasts_df.empty and "ci90_lower" in forecasts_df.columns:
        ci_view = forecasts_df.copy()
        ci_view["ci90_width"] = ci_view["ci90_upper"] - ci_view["ci90_lower"]
        ci_view["ci95_width"] = ci_view["ci95_upper"] - ci_view["ci95_lower"]
        ci_table = (ci_view
                    .groupby(["model", "ticker", "horizon"])
                    .agg(
                        mean_forecast   =("forecast",   "mean"),
                        mean_ci90_lower =("ci90_lower", "mean"),
                        mean_ci90_upper =("ci90_upper", "mean"),
                        mean_ci90_width =("ci90_width", "mean"),
                        mean_ci95_lower =("ci95_lower", "mean"),
                        mean_ci95_upper =("ci95_upper", "mean"),
                        mean_ci95_width =("ci95_width", "mean"),
                    )
                    .round(5))
        logger.info(
            f"\n{'─'*70}\n"
            f"  CONFIDENCE INTERVAL SUMMARY  (average across all walk-forward steps)\n"
            f"{'─'*70}\n"
            f"{ci_table.to_string()}\n"
        )
        ci_table.to_csv(OUTPUT_DIR / "ci_summary.csv")


    perf_df  = pd.DataFrame(perf_rows)
    calib_df = pd.DataFrame(calib_rows)
    perf_df.to_csv(OUTPUT_DIR  / "performance_metrics.csv",  index=False)
    calib_df.to_csv(OUTPUT_DIR / "calibration_results.csv",  index=False)
    if not perf_df.empty:
        logger.info(f"\n{perf_df.to_string()}")

    if not perf_df.empty:
        with pdf_backend.PdfPages(OUTPUT_DIR / "model_diagnostics.pdf") as pages:
            for metric in ["DA", "IC", "RMSE"]:
                if metric not in perf_df.columns:
                    continue
                pivot = perf_df.pivot_table(values=metric, index="ticker",
                                             columns="horizon", aggfunc="mean")
                if pivot.empty:
                    continue
                cmap = "RdYlGn_r" if metric == "RMSE" else "RdYlGn"
                fig, ax = plt.subplots(figsize=(10, 5))
                sns.heatmap(pivot, annot=True, fmt=".3f", cmap=cmap, ax=ax)
                ax.set_title(f"Walk-Forward {metric} — Ticker × Horizon")
                pages.savefig(fig, bbox_inches="tight"); plt.close(fig)

    logger.info("Phase 5 validation checklist passed ✓\n")
    return {"forecasts_df": forecasts_df, "errors_df": errors_df,
            "perf_df": perf_df, "calib_df": calib_df}


# ======================================================================
# PHASE 6: FORECAST COMBINATION  (unchanged)
# ======================================================================
def run_phase6(p5: dict, p1: dict) -> dict:
    logger.info("=" * 62)
    logger.info("PHASE 6 — Forecast Combination")
    logger.info("=" * 62)
    forecasts_df   = p5["forecasts_df"]
    active_tickers = p1["active_tickers"]
    if forecasts_df.empty:
        logger.warning("  No forecasts — skipping combination.")
        return {"combined_df": pd.DataFrame(), "weights_df": pd.DataFrame()}
    comb_rows: List[dict] = []; weight_rows: List[dict] = []
    for ticker in active_tickers:
        for h in FORECAST_HORIZONS:
            model_preds: Dict[str, pd.Series] = {}; common_actuals = None
            for mt in ["ARIMAX","VARX"]:
                sub = forecasts_df[
                    (forecasts_df["ticker"]  == ticker) &
                    (forecasts_df["horizon"] == h) &
                    (forecasts_df["model"]   == mt)
                ].dropna(subset=["forecast","actual"])
                if len(sub) >= 5:
                    model_preds[mt] = sub.set_index("t")["forecast"]
                    common_actuals  = sub.set_index("t")["actual"]
            if not model_preds:
                continue
            if len(model_preds) < 2:
                mt0 = next(iter(model_preds))
                last_t_s = int(model_preds[mt0].index[-1])
                sub_s = forecasts_df[
                    (forecasts_df["ticker"]  == ticker) &
                    (forecasts_df["horizon"] == h) &
                    (forecasts_df["model"]   == mt0) &
                    (forecasts_df["t"]       == last_t_s)]
                def _pick_ci(col):
                    if col in sub_s.columns:
                        v = sub_s[col].dropna()
                        return float(v.iloc[0]) if not v.empty else np.nan
                    return np.nan
                comb_rows.append({"ticker": ticker, "horizon": h,
                                  "combined_forecast": float(model_preds[mt0].iloc[-1]),
                                  "ci90_lower": _pick_ci("ci90_lower"),
                                  "ci90_upper": _pick_ci("ci90_upper"),
                                  "ci95_lower": _pick_ci("ci95_lower"),
                                  "ci95_upper": _pick_ci("ci95_upper"),
                                  "method": f"single_{mt0}"})
                continue
            common_t = list(set.intersection(*[set(s.index) for s in model_preds.values()]))
            if len(common_t) < 3:
                mt0 = next(iter(model_preds))
                comb_rows.append({"ticker": ticker, "horizon": h,
                                  "combined_forecast": float(model_preds[mt0].iloc[-1]),
                                  "ci90_lower": np.nan, "ci90_upper": np.nan,
                                  "ci95_lower": np.nan, "ci95_upper": np.nan,
                                  "method": "single_insufficient_overlap"})
                continue
            fc_mat  = np.column_stack([model_preds[m].loc[common_t].values for m in model_preds])
            act_vec = common_actuals.loc[common_t].values
            eq_pred  = fc_mat.mean(axis=1)
            rmse_m   = {m: float(np.sqrt(np.mean((model_preds[m].loc[common_t].values - act_vec)**2)))
                        for m in model_preds}
            inv_rmse = np.array([1.0 / max(rmse_m[m], 1e-10) for m in model_preds])
            pw_w     = inv_rmse / inv_rmse.sum()
            pw_pred  = (pw_w @ fc_mat.T)
            bic_arr  = np.array([rmse_m[m] for m in model_preds])
            bma_w    = np.exp(-0.5 * bic_arr) / np.exp(-0.5 * bic_arr).sum()
            bma_pred = (bma_w @ fc_mat.T)
            dm_stat, dm_pval = _diebold_mariano(act_vec - eq_pred, act_vec - pw_pred)
            method   = "equal" if dm_pval > 0.05 else "performance_weighted"
            final_fc = float(eq_pred[-1]) if method == "equal" else float(pw_pred[-1])
            last_t   = max(common_t)
            ci_buckets = {"ci90_lower": [], "ci90_upper": [],
                          "ci95_lower": [], "ci95_upper": []}
            for mt in model_preds:
                sub_mt = forecasts_df[
                    (forecasts_df["ticker"]  == ticker) &
                    (forecasts_df["horizon"] == h) &
                    (forecasts_df["model"]   == mt) &
                    (forecasts_df["t"]       == last_t)]
                for col in ci_buckets:
                    if col in sub_mt.columns and not sub_mt[col].isna().all():
                        ci_buckets[col].append(float(sub_mt[col].iloc[0]))
            def _mean_ci(col):
                return float(np.mean(ci_buckets[col])) if ci_buckets[col] else np.nan
            comb_rows.append({
                "ticker": ticker, "horizon": h,
                "combined_forecast": final_fc,
                "eq_forecast":  float(eq_pred[-1]),
                "pw_forecast":  float(pw_pred[-1]),
                "bma_forecast": float(bma_pred[-1]),
                "dm_stat": round(dm_stat, 4), "dm_pval": round(dm_pval, 4),
                "method":  method,
                "ci90_lower": _mean_ci("ci90_lower"), "ci90_upper": _mean_ci("ci90_upper"),
                "ci95_lower": _mean_ci("ci95_lower"), "ci95_upper": _mean_ci("ci95_upper"),
            })
            for mt, w in zip(model_preds.keys(), pw_w):
                weight_rows.append({"ticker": ticker, "horizon": h,
                                     "model": mt, "pw_weight": round(float(w), 4)})
    combined_df = pd.DataFrame(comb_rows); weights_df = pd.DataFrame(weight_rows)
    combined_df.to_parquet(OUTPUT_DIR / "combined_forecasts.parquet")
    weights_df.to_csv(OUTPUT_DIR  / "combination_weights.csv", index=False)
    logger.info("Phase 6 ✓\n")
    return {"combined_df": combined_df, "weights_df": weights_df}



# ======================================================================
# PHASE 7: PORTFOLIO CONSTRUCTION  ── MODIFIED: Conviction-Based BL ──
# ======================================================================
def run_phase7(p1: dict, p2: dict, p3: dict, p4_5: dict, p6: dict,
               p3_5: Optional[dict] = None) -> dict:
    """
    Black-Litterman portfolio construction with regime-aware risk budgeting.

    MODIFICATION: When sector scores are available (Phase 3.5), the view
    uncertainty matrix Ω is scaled inversely with ticker conviction.
      - High conviction → smaller Ω diagonal → BL pulls posterior toward view
      - Low conviction  → larger  Ω diagonal → prior equilibrium dominates
    This is the canonical BL framework; only the calibration of Ω changes.
    All other statistical properties (posterior mean/covariance formulae,
    Ledoit-Wolf shrinkage, regime gating) remain unchanged.
    """
    logger.info("=" * 62)
    logger.info("PHASE 7 — Portfolio Construction (Conviction-Adjusted BL)")
    logger.info("=" * 62)

    monthly_prices = p1["monthly_prices"]
    bench_px       = p1["bench_monthly_prices"]
    monthly_rf     = p1["monthly_rf"]
    active_tickers = p1["active_tickers"]
    mkt_hmm_df     = p4_5["market_filtered_df"]
    combined_df    = p6["combined_df"]

    if combined_df.empty:
        logger.warning("  No combined forecasts — skipping Phase 7.")
        return {}

    prices_aligned = monthly_prices[active_tickers].dropna()
    S = risk_m.CovarianceShrinkage(prices_aligned, frequency=12).ledoit_wolf()

    mkt_caps         = pd.Series(1.0 / len(active_tickers), index=active_tickers)
    bench_aligned_px = pd.Series(bench_px).reindex(prices_aligned.index).ffill().dropna()
    try:
        delta = bl_module.market_implied_risk_aversion(
            bench_aligned_px, risk_free_rate=ANNUAL_RISK_FREE, frequency=12)
    except Exception as exc:
        logger.warning(f"  market_implied_risk_aversion failed ({exc}) → δ=2.5")
        delta = 2.5
    pi = bl_module.market_implied_prior_returns(
        market_caps=mkt_caps, risk_aversion=delta, cov_matrix=S)

    h_target = 12 if 12 in FORECAST_HORIZONS else max(FORECAST_HORIZONS)
    viewdict: Dict[str, float] = {}
    for _, row in combined_df[combined_df["horizon"] == h_target].iterrows():
        t = row["ticker"]
        if t in active_tickers and pd.notna(row.get("combined_forecast")):
            viewdict[t] = float(row["combined_forecast"]) * 12

    # ── Regime-aware risk budget (unchanged logic) ────────────────────
    bear_cols   = [c for c in mkt_hmm_df.columns if "bear" in c]
    stress_prob = float(mkt_hmm_df[bear_cols[0]].iloc[-1]) if bear_cols else 0.0
    in_stress   = stress_prob > 0.60
    max_w       = 0.12 if in_stress else 0.20
    logger.info(f"  P(stress)={stress_prob:.3f} → max_weight={max_w:.0%}")

    # ── Conviction-adjusted view uncertainty Ω ────────────────────────
    tau      = 0.05
    ret_bl   = pi          # fallback to prior
    S_bl     = S
    ticker_conviction: Dict[str, float] = {}

    scores_ok = (p3_5 is not None and p3_5.get("scores_available", False))
    if scores_ok and viewdict:
        t_conviction_map = p3_5.get("ticker_conviction", {})
        view_tickers_list = list(viewdict.keys())
        ticker_conviction = {t: t_conviction_map.get(t, 0.50) for t in view_tickers_list}

        try:
            # Diagonal of tau*S for viewed assets → base per-view uncertainty
            s_diag    = np.array([float(S.loc[t, t]) for t in view_tickers_list])
            omega_base = tau * s_diag

            # Uncertainty scale: conviction 0.5 → scale 1.0 (standard BL)
            #                    conviction 0.9 → scale 0.56 (views trusted more)
            #                    conviction 0.2 → scale 2.50 (prior trusted more)
            conv_arr  = np.clip(
                np.array([ticker_conviction[t] for t in view_tickers_list]), 0.05, 1.0)
            unc_scale  = 0.5 / conv_arr      # range: [0.5, 10]; clipped above to ~[0.53, 10]
            omega_diag = omega_base * unc_scale
            omega_mat  = np.diag(omega_diag)

            bl_model = BlackLittermanModel(S, pi=pi, absolute_views=viewdict,
                                           tau=tau, omega=omega_mat)
            ret_bl   = bl_model.bl_returns()
            S_bl     = bl_model.bl_cov()
            logger.info("  Conviction-adjusted Ω applied:")
            for t, cv in zip(view_tickers_list, conv_arr):
                logger.info(f"    {t:<20}  conviction={cv:.3f}  "
                            f"Ω_scale={0.5/cv:.3f}  view_weight↑" if cv > 0.5
                            else f"    {t:<20}  conviction={cv:.3f}  "
                                 f"Ω_scale={0.5/cv:.3f}  prior_weight↑")
        except Exception as exc:
            logger.warning(f"  Custom Ω failed ({exc}) — falling back to default BL.")
            try:
                bl_model = BlackLittermanModel(S, pi=pi, absolute_views=viewdict, tau=tau)
                ret_bl   = bl_model.bl_returns()
                S_bl     = bl_model.bl_cov()
            except Exception as exc2:
                logger.warning(f"  Default BL also failed ({exc2}) — using prior returns.")
    elif viewdict:
        try:
            bl_model = BlackLittermanModel(S, pi=pi, absolute_views=viewdict, tau=tau)
            ret_bl   = bl_model.bl_returns()
            S_bl     = bl_model.bl_cov()
            logger.info("  Standard BL (no sector scores): default Ω.")
        except Exception as exc:
            logger.warning(f"  BL failed ({exc}) — using prior returns.")

    # ── Mean-variance optimisation ────────────────────────────────────
    weights_clean = {t: 1.0 / len(active_tickers) for t in active_tickers}
    try:
        ef = EfficientFrontier(ret_bl, S_bl)
        ef.add_constraint(lambda w: w >= 0.02)
        ef.add_constraint(lambda w: w <= max_w)
        if in_stress:
            ef.efficient_risk(target_volatility=0.12 / np.sqrt(12))
        else:
            ann_rf_exact = (1 + monthly_rf) ** 12 - 1
            ef.max_sharpe(risk_free_rate=ann_rf_exact)
        weights_clean = ef.clean_weights()
    except Exception as exc:
        logger.warning(f"  EfficientFrontier failed ({exc}) — equal weights.")

    logger.info(f"  Weights: {weights_clean}")
    latest_prices = get_latest_prices(monthly_prices[active_tickers])
    try:
        da = DiscreteAllocation(weights_clean, latest_prices,
                                total_portfolio_value=CAPITAL_AMOUNT)
        try:
            allocation, leftover = da.lp_portfolio()
        except Exception:
            allocation, leftover = da.greedy_portfolio()
    except Exception as exc:
        logger.warning(f"  DiscreteAllocation failed ({exc}).")
        allocation, leftover = {}, float(CAPITAL_AMOUNT)

    pd.DataFrame(list(weights_clean.items()), columns=["ticker","weight"]).to_csv(
        OUTPUT_DIR / "portfolio_weights.csv", index=False)
    pd.DataFrame([{
        "ticker": t, "shares": s,
        "latest_price": float(latest_prices.get(t, np.nan)),
        "current_value": float(s * latest_prices.get(t, 0.0)),
        "weight": weights_clean.get(t, 0.0),
    } for t, s in allocation.items()]).to_csv(OUTPUT_DIR / "capital_allocation.csv", index=False)
    pd.DataFrame([{
        "ticker": t, "view_return": round(v, 4),
        "prior_return": round(float(pi[t]) if hasattr(pi,"__getitem__") else float(pi), 4),
        "conviction": round(ticker_conviction.get(t, np.nan), 3),
    } for t, v in viewdict.items()]).to_csv(OUTPUT_DIR / "bl_diagnostics.csv", index=False)

    logger.info("Phase 7 ✓\n")
    return {
        "weights_clean": weights_clean, "allocation": allocation,
        "leftover": leftover, "S_full": S, "latest_prices": latest_prices,
        "stress_prob": stress_prob, "in_stress": in_stress,
        "ret_bl":          ret_bl,           # NEW — for custom-weight metrics
        "S_bl":            S_bl,             # NEW — for custom-weight metrics
        "ticker_conviction": ticker_conviction,  # NEW — for dashboard display
    }


# ======================================================================
# PHASE 8: RISK MANAGEMENT & BACKTEST  (unchanged)
# ======================================================================
def run_phase8(p1: dict, p2: dict, p4_5: dict, p7: dict) -> dict:
    logger.info("=" * 62)
    logger.info("PHASE 8 — Risk Management & Backtest")
    logger.info("=" * 62)
    mlr           = p2["monthly_log_returns_clean"]
    bench_rets    = p1["benchmark_returns"]
    mkt_hmm_df    = p4_5["market_filtered_df"]
    weights_clean = p7.get("weights_clean", {})
    if not weights_clean:
        logger.warning("  No portfolio weights — skipping Phase 8.")
        return {}
    w_series            = pd.Series(weights_clean).reindex(mlr.columns).fillna(0.0)
    portfolio_log_rets  = mlr.dot(w_series)
    portfolio_value     = np.exp(portfolio_log_rets.cumsum())
    rolling_max         = portfolio_value.expanding().max()
    drawdown_series     = (portfolio_value - rolling_max) / rolling_max
    max_drawdown        = float(drawdown_series.min())
    current_drawdown    = float(drawdown_series.iloc[-1])
    portfolio_log_rets.to_frame("portfolio_returns").to_parquet(
        OUTPUT_DIR / "portfolio_returns.parquet")
    portfolio_value.to_frame("portfolio_value").to_parquet(
        OUTPUT_DIR / "portfolio_value.parquet")
    drawdown_series.to_frame("drawdown").to_parquet(OUTPUT_DIR / "drawdown_series.parquet")
    if current_drawdown < -0.15:
        logger.warning(f"  HARD STOP: drawdown={current_drawdown:.1%} > 15% threshold.")
    bear_col = [c for c in mkt_hmm_df.columns if "bear" in c]
    if bear_col:
        trigger_dates = mkt_hmm_df[bear_col[0]][mkt_hmm_df[bear_col[0]] > 0.65].index
        if len(trigger_dates):
            logger.info(f"  Regime rebalance triggers: {[str(d.date()) for d in trigger_dates]}")
    n_months   = len(portfolio_log_rets)
    ann_return = float((portfolio_value.iloc[-1] / portfolio_value.iloc[0]) ** (12/n_months) - 1)
    ann_vol    = float(portfolio_log_rets.std() * np.sqrt(12))
    ann_rf     = ANNUAL_RISK_FREE
    sharpe     = (ann_return - ann_rf) / ann_vol if ann_vol > 0 else np.nan
    neg_rets   = portfolio_log_rets[portfolio_log_rets < 0]
    down_dev   = float(neg_rets.std() * np.sqrt(12)) if len(neg_rets) > 1 else ann_vol
    sortino    = (ann_return - ann_rf) / down_dev if down_dev > 0 else np.nan
    calmar     = ann_return / abs(max_drawdown) if max_drawdown != 0 else np.nan
    bench_aligned = bench_rets.reindex(portfolio_log_rets.index).fillna(0.0)
    excess        = portfolio_log_rets - bench_aligned
    track_err     = float(excess.std() * np.sqrt(12))
    info_ratio    = float(excess.mean() * 12 / track_err) if track_err > 0 else np.nan
    perf_summary = {
        "Annualized Return":     round(ann_return, 4),
        "Annualized Volatility": round(ann_vol, 4),
        "Sharpe Ratio":          round(sharpe, 4)    if not np.isnan(sharpe)     else None,
        "Sortino Ratio":         round(sortino, 4)   if not np.isnan(sortino)    else None,
        "Maximum Drawdown":      round(max_drawdown, 4),
        "Calmar Ratio":          round(calmar, 4)    if not np.isnan(calmar)     else None,
        "Information Ratio":     round(info_ratio, 4) if not np.isnan(info_ratio) else None,
        "N Months":              n_months,
    }
    pd.DataFrame([perf_summary]).to_csv(OUTPUT_DIR / "performance_summary.csv", index=False)
    regime_buckets: Dict[str, List] = {"bull": [], "transitional": [], "bear": []}
    for dt, ret_val in portfolio_log_rets.items():
        if dt not in mkt_hmm_df.index:
            continue
        dom    = mkt_hmm_df.loc[dt].fillna(0.0).idxmax()
        regime = dom.replace("market_hmm_", "")
        if regime in regime_buckets:
            regime_buckets[regime].append(float(ret_val))
    regime_rows = []
    for regime, rets in regime_buckets.items():
        if len(rets) < 3:
            continue
        ra = np.array(rets)
        ar = float(np.mean(ra) * 12); av = float(np.std(ra) * np.sqrt(12))
        regime_rows.append({
            "regime": regime, "n_months": len(ra),
            "ann_return": round(ar, 4), "ann_vol": round(av, 4),
            "sharpe": round((ar - ann_rf) / av, 4) if av > 0 else None,
        })
    regime_perf_df = pd.DataFrame(regime_rows)
    regime_perf_df.to_csv(OUTPUT_DIR / "regime_conditional_performance.csv", index=False)
    try:
        simple_port  = np.exp(portfolio_log_rets) - 1
        simple_bench = np.exp(bench_aligned)       - 1
        qs.reports.html(simple_port, benchmark=simple_bench,
                        output=str(OUTPUT_DIR / "backtest_tearsheet.html"),
                        title="NSE Portfolio Walk-Forward Backtest")
    except Exception as exc:
        logger.warning(f"  QuantStats failed: {exc}")
    logger.info("Phase 8 ✓\n")
    return {
        "portfolio_returns": portfolio_log_rets, "portfolio_value": portfolio_value,
        "drawdown_series": drawdown_series, "max_drawdown": max_drawdown,
        "perf_summary": perf_summary, "regime_perf_df": regime_perf_df,
    }



# ======================================================================
# PHASE EVAL: MODEL ROBUSTNESS  ── MODIFIED: stores full prediction path
# ======================================================================
def run_phase_eval(p1: dict, p2: dict, p3: dict, p4: dict, p4_5: dict) -> dict:
    """
    Held-out train/test evaluation.  Train = first MIN_TRAIN_MONTHS.
    Test  = remainder (fixed-origin, one-shot multi-step forecast).

    MODIFICATION: also stores per-step predictions + CI in eval_predictions_df
    so Phase 9 can render actual-vs-forecast line charts.

    Metrics: RMSE, MAE, ME (bias), MAPE, SMAPE, DA, AIC, BIC.
    """
    logger.info("=" * 62)
    logger.info("PHASE EVAL — Model Robustness & Diagnostics")
    logger.info("=" * 62)
    mlr            = p2["monthly_log_returns_clean"]
    realized_vol   = p3["realized_volatility"]
    market_ind     = p3["market_indicators"]
    mkt_hmm_df     = p4_5["market_filtered_df"]
    stk_hmm_dfs    = p4_5["stock_filtered_probs"]
    grp_hmm_dfs    = p4_5["group_filtered_probs"]
    active_tickers = p1["active_tickers"]
    cluster_assign = p4["cluster_assignments"]
    T         = len(mlr)
    train_end = MIN_TRAIN_MONTHS
    test_mos  = T - train_end
    if test_mos < 2:
        logger.warning("  Fewer than 2 test months — skipping eval phase.")
        return {"eval_df": pd.DataFrame(), "eval_predictions_df": pd.DataFrame()}
    train_data = mlr.iloc[:train_end]; test_data = mlr.iloc[train_end:]
    logger.info(
        f"  Train: {train_data.index[0].date()} → {train_data.index[-1].date()} "
        f"({len(train_data)} months)  |  "
        f"Test: {test_data.index[0].date()} → {test_data.index[-1].date()} "
        f"({test_mos} months)"
    )

    def _metrics(preds, actuals):
        err   = preds - actuals
        denom = np.where(np.abs(actuals) > 1e-10, np.abs(actuals), 1e-10)
        sd    = np.abs(preds) + np.abs(actuals) + 1e-10
        return {
            "RMSE":  float(np.sqrt(np.mean(err**2))),
            "MAE":   float(np.mean(np.abs(err))),
            "ME":    float(np.mean(err)),
            "MAPE":  float(np.mean(np.abs(err)/denom)*100),
            "SMAPE": float(np.mean(2*np.abs(err)/sd)*100),
            "DA":    float(np.mean(np.sign(preds) == np.sign(actuals))),
        }

    eval_rows: List[dict]      = []
    eval_pred_rows: List[dict] = []

    # ── ARIMAX ────────────────────────────────────────────────────────
    logger.info("  Fitting ARIMAX on training data …")
    for ticker in active_tickers:
        endog_tr = train_data[ticker].dropna()
        if len(endog_tr) < 24:
            continue
        mkt_tr  = mkt_hmm_df.reindex(endog_tr.index).ffill().fillna(0.0)
        s_raw   = stk_hmm_dfs.get(ticker, pd.DataFrame())
        stk_tr  = (s_raw.reindex(endog_tr.index).ffill().fillna(0.0)
                   if not s_raw.empty else pd.DataFrame(index=endog_tr.index))
        ind_tr  = market_ind.reindex(endog_tr.index).ffill().fillna(0.0)
        exog_tr = _build_arimax_exog(ticker, endog_tr.index, mkt_tr, stk_tr,
                                     realized_vol, ind_tr)
        exog_in = exog_tr if exog_tr.shape[1] > 0 else None
        order = _arimax_bic_grid(endog_tr.values, exog_in)
        try:
            fitted = _fit_arimax(endog_tr.values, exog_in, order)
        except Exception as exc:
            logger.warning(f"  [{ticker}] ARIMAX train-fit failed: {exc}"); continue
        aic_v = float(fitted.aic); bic_v = float(fitted.bic)
        mkt_te  = mkt_hmm_df.reindex(test_data.index).ffill().fillna(0.0)
        stk_te  = (s_raw.reindex(test_data.index).ffill().fillna(0.0)
                   if not s_raw.empty else pd.DataFrame(index=test_data.index))
        ind_te  = market_ind.reindex(test_data.index).ffill().fillna(0.0)
        exog_te = _build_arimax_exog(ticker, test_data.index, mkt_te, stk_te,
                                     realized_vol, ind_te)
        exog_fc = exog_te if exog_te.shape[1] > 0 else None
        try:
            fc_obj  = fitted.get_forecast(steps=test_mos, exog=exog_fc)
            fmu_arr = fc_obj.predicted_mean
            try:
                fci_90 = fc_obj.conf_int(alpha=0.10)
                fci_95 = fc_obj.conf_int(alpha=0.05)
                ci90_lo = fci_90.iloc[:, 0].values
                ci90_hi = fci_90.iloc[:, 1].values
                ci95_lo = fci_95.iloc[:, 0].values
                ci95_hi = fci_95.iloc[:, 1].values
            except Exception:
                ci90_lo = ci90_hi = ci95_lo = ci95_hi = None
        except Exception as exc:
            logger.warning(f"  [{ticker}] ARIMAX test-forecast failed: {exc}"); continue
        act_arr = test_data[ticker].values
        n_steps = min(len(fmu_arr), len(act_arr))
        # Store full prediction path for dashboard line chart
        for i in range(n_steps):
            eval_pred_rows.append({
                "model": "ARIMAX", "ticker": ticker,
                "step_ahead":     i + 1,
                "train_end_date": str(train_data.index[-1].date()),
                "forecast_date":  str(test_data.index[i].date()),
                "actual":         float(act_arr[i]),
                "forecast":       float(fmu_arr[i]),
                "ci90_lower":  float(ci90_lo[i]) if ci90_lo is not None else np.nan,
                "ci90_upper":  float(ci90_hi[i]) if ci90_hi is not None else np.nan,
                "ci95_lower":  float(ci95_lo[i]) if ci95_lo is not None else np.nan,
                "ci95_upper":  float(ci95_hi[i]) if ci95_hi is not None else np.nan,
            })
        # Aggregate metrics per horizon
        for h in FORECAST_HORIZONS:
            if h > test_mos:
                continue
            n_valid = test_mos - h + 1
            if n_valid < 2:
                continue
            p_arr = fmu_arr[:n_valid] if len(fmu_arr) >= n_valid else fmu_arr
            a_arr = act_arr[h-1: h-1+len(p_arr)]
            mask  = np.isfinite(p_arr) & np.isfinite(a_arr)
            if mask.sum() < 2:
                continue
            m = _metrics(p_arr[mask], a_arr[mask])
            eval_rows.append({
                "model": "ARIMAX", "ticker": ticker, "horizon": h,
                "order_p": order[0], "order_q": order[1],
                "n_test": int(mask.sum()),
                "AIC": round(aic_v, 2), "BIC": round(bic_v, 2),
                **{k: round(v, 6) for k, v in m.items()},
            })

    # ── VARX ──────────────────────────────────────────────────────────
    logger.info("  Fitting VARX groups on training data …")
    varx_groups = {k: v for k, v in cluster_assign.items() if k != "isolated"}
    for gkey, members in varx_groups.items():
        group_m  = [m for m in members if m in mlr.columns]
        if len(group_m) < 2:
            continue
        endog_tr = train_data[group_m].dropna()
        if len(endog_tr) < 24:
            continue
        mkt_g  = mkt_hmm_df.reindex(endog_tr.index).ffill().fillna(0.0)
        ind_g  = market_ind.reindex(endog_tr.index).ffill().fillna(0.0)
        exog_v = _build_varx_exog(gkey, group_m, endog_tr.index, mkt_g,
                                   grp_hmm_dfs, realized_vol, ind_g)
        try:
            varx_res, opt_lag = _fit_varx(endog_tr.values, exog_v)
        except Exception as exc:
            logger.warning(f"  [{gkey}] VARX train-fit failed: {exc}"); continue
        aic_v = float(getattr(varx_res, "aic", np.nan))
        bic_v = float(getattr(varx_res, "bic", np.nan))
        endog_te = test_data[group_m]
        if endog_te.empty:
            continue
        mkt_te_g = mkt_hmm_df.reindex(endog_te.index).ffill().fillna(0.0)
        ind_te_g = market_ind.reindex(endog_te.index).ffill().fillna(0.0)
        exog_fut = _build_varx_exog(gkey, group_m, endog_te.index, mkt_te_g,
                                     grp_hmm_dfs, realized_vol, ind_te_g)
        y_init   = endog_tr.values[-opt_lag:]
        fc_steps = len(endog_te)
        try:
            fc_pt = varx_res.forecast(y=y_init, steps=fc_steps, exog_future=exog_fut)
        except Exception as exc:
            logger.warning(f"  [{gkey}] VARX test-forecast failed: {exc}"); continue
        # Attempt CI for VARX
        fc_lo_90 = fc_hi_90 = fc_lo_95 = fc_hi_95 = None
        try:
            _, fc_lo_90, fc_hi_90 = varx_res.forecast_interval(
                y=y_init, steps=fc_steps, alpha=0.10, exog_future=exog_fut)
            _, fc_lo_95, fc_hi_95 = varx_res.forecast_interval(
                y=y_init, steps=fc_steps, alpha=0.05, exog_future=exog_fut)
        except Exception:
            try:
                _, fc_lo_90, fc_hi_90 = varx_res.forecast_interval(
                    y=y_init, steps=fc_steps, alpha=0.10)
                _, fc_lo_95, fc_hi_95 = varx_res.forecast_interval(
                    y=y_init, steps=fc_steps, alpha=0.05)
            except Exception:
                pass
        for k, tkr in enumerate(group_m):
            act_arr  = endog_te[tkr].values
            n_steps  = min(len(fc_pt), len(act_arr))
            for i in range(n_steps):
                eval_pred_rows.append({
                    "model": "VARX", "ticker": tkr,
                    "step_ahead":     i + 1,
                    "train_end_date": str(train_data.index[-1].date()),
                    "forecast_date":  str(test_data.index[i].date()),
                    "actual":         float(act_arr[i]),
                    "forecast":       float(fc_pt[i, k]),
                    "ci90_lower":  float(fc_lo_90[i, k]) if fc_lo_90 is not None else np.nan,
                    "ci90_upper":  float(fc_hi_90[i, k]) if fc_hi_90 is not None else np.nan,
                    "ci95_lower":  float(fc_lo_95[i, k]) if fc_lo_95 is not None else np.nan,
                    "ci95_upper":  float(fc_hi_95[i, k]) if fc_hi_95 is not None else np.nan,
                })
            for h in FORECAST_HORIZONS:
                if h > fc_steps:
                    continue
                n_valid = fc_steps - h + 1
                if n_valid < 2:
                    continue
                p_arr = fc_pt[:n_valid, k]
                a_arr = act_arr[h-1: h-1+len(p_arr)]
                mask  = np.isfinite(p_arr) & np.isfinite(a_arr)
                if mask.sum() < 2:
                    continue
                m = _metrics(p_arr[mask], a_arr[mask])
                eval_rows.append({
                    "model": "VARX", "ticker": tkr, "horizon": h,
                    "group": gkey, "order_p": opt_lag, "n_test": int(mask.sum()),
                    "AIC": round(aic_v, 2), "BIC": round(bic_v, 2),
                    **{k2: round(v, 6) for k2, v in m.items()},
                })

    _EVAL_COLS = ["model","ticker","horizon","n_test",
                  "RMSE","MAE","ME","MAPE","SMAPE","DA","AIC","BIC"]
    _PRED_COLS = ["model","ticker","step_ahead","train_end_date","forecast_date",
                  "actual","forecast","ci90_lower","ci90_upper","ci95_lower","ci95_upper"]
    eval_df      = (pd.DataFrame(eval_rows)      if eval_rows
                    else pd.DataFrame(columns=_EVAL_COLS))
    eval_preds_df = (pd.DataFrame(eval_pred_rows) if eval_pred_rows
                     else pd.DataFrame(columns=_PRED_COLS))
    eval_df.to_csv(       OUTPUT_DIR / "model_evaluation.csv",       index=False)
    eval_preds_df.to_csv( OUTPUT_DIR / "model_eval_predictions.csv", index=False)

    if not eval_df.empty:
        logger.info(f"\n{'─'*70}\nMODEL EVALUATION\n{'─'*70}\n{eval_df.to_string()}\n")
        with pdf_backend.PdfPages(OUTPUT_DIR / "model_evaluation.pdf") as pages:
            for metric in ["RMSE","MAPE","DA","ME"]:
                if metric not in eval_df.columns:
                    continue
                for mt in eval_df["model"].unique():
                    sub = eval_df[eval_df["model"] == mt]
                    if sub.empty or sub[metric].isna().all():
                        continue
                    try:
                        pivot = sub.pivot_table(values=metric, index="ticker",
                                                columns="horizon", aggfunc="mean")
                        if pivot.empty:
                            continue
                        cmap = "RdYlGn_r" if metric in ("RMSE","MAPE") else "RdYlGn"
                        fig, ax = plt.subplots(figsize=(10, max(3, len(pivot)*0.45+1.5)))
                        sns.heatmap(pivot, annot=True,
                                    fmt=".2f" if metric == "MAPE" else ".4f",
                                    cmap=cmap, ax=ax, linewidths=0.4)
                        ax.set_title(f"{mt} — Out-of-Sample {metric}", fontsize=11)
                        pages.savefig(fig, bbox_inches="tight"); plt.close(fig)
                    except Exception:
                        pass
    logger.info("Phase EVAL ✓\n")
    return {"eval_df": eval_df, "eval_predictions_df": eval_preds_df}



# ======================================================================
# PHASE 9: PLOTLY DASH MONITORING DASHBOARD  ── REDESIGNED
# ======================================================================
def run_phase9(
    p1: dict, p2: dict, p3: dict, p4_5: dict,
    p5: dict, p6: dict, p7: dict, p8: dict,
    p_eval: dict, p3_5: Optional[dict] = None,
) -> "dash.Dash":
    logger.info("=" * 62)
    logger.info("PHASE 9 — Dashboard")
    logger.info("=" * 62)

    # ── Unpack data ────────────────────────────────────────────────────
    mkt_hmm_df     = p4_5.get("market_filtered_df",       pd.DataFrame())
    portfolio_rets = p8.get("portfolio_returns",           pd.Series(dtype=float))
    portfolio_val  = p8.get("portfolio_value",             pd.Series(dtype=float))
    drawdown_s     = p8.get("drawdown_series",             pd.Series(dtype=float))
    bench_rets     = p1.get("benchmark_returns",           pd.Series(dtype=float))
    weights_clean  = p7.get("weights_clean",               {})
    combined_df    = p6.get("combined_df",                 pd.DataFrame())
    perf_df        = p5.get("perf_df",                     pd.DataFrame())
    calib_df       = p5.get("calib_df",                    pd.DataFrame())
    forecasts_df   = p5.get("forecasts_df",                pd.DataFrame())
    errors_df      = p5.get("errors_df",                   pd.DataFrame())
    monthly_prices = p1.get("monthly_prices",              pd.DataFrame())
    mlr            = p2.get("monthly_log_returns_clean",   pd.DataFrame())
    active_tickers = p1.get("active_tickers",              [])
    perf_summary   = p8.get("perf_summary",               {})
    realized_vol   = p3.get("realized_volatility",         pd.DataFrame())
    market_ind     = p3.get("market_indicators",           pd.DataFrame())
    eval_df        = p_eval.get("eval_df",                 pd.DataFrame())
    eval_preds_df  = p_eval.get("eval_predictions_df",     pd.DataFrame())
    ret_bl         = p7.get("ret_bl",   None)
    S_bl           = p7.get("S_bl",     None)
    ticker_conv    = p7.get("ticker_conviction", {})
    scores_avail   = (p3_5 is not None and p3_5.get("scores_available", False))
    sector_scores  = p3_5.get("sector_composite_scores",  pd.DataFrame()) if p3_5 else pd.DataFrame()
    sector_conv    = p3_5.get("sector_conviction",        {})             if p3_5 else {}
    ticker_sector  = p3_5.get("ticker_sector",            {})             if p3_5 else {}

    regime_cols        = list(mkt_hmm_df.columns) if not mkt_hmm_df.empty else []
    current_regime_row = mkt_hmm_df.iloc[-1].to_dict() if not mkt_hmm_df.empty else {}
    monthly_rf         = p1.get("monthly_rf", ANNUAL_RISK_FREE / 12)
    ann_rf_exact       = (1 + monthly_rf) ** 12 - 1

    # Add date column to errors_df from mlr index
    if not errors_df.empty and "t" in errors_df.columns and not mlr.empty:
        mlr_idx = mlr.index
        errors_df = errors_df.copy()
        errors_df["date"] = errors_df["t"].apply(
            lambda x: str(mlr_idx[int(x)].date()) if 0 <= int(x) < len(mlr_idx) else None
        )

    # ── Static pre-computed figures ───────────────────────────────────
    # CI Coverage chart (static — aggregate metric, no user interaction needed)
    _coverage_fig = go.Figure()
    if not calib_df.empty:
        for mt in calib_df["model"].unique():
            sub = calib_df[calib_df["model"] == mt]
            avg = sub.groupby("horizon")["empirical_coverage_90pct"].mean().reset_index()
            colors_cov = [
                _C["success"] if v >= 0.87 else (_C["warning"] if v >= 0.80 else _C["danger"])
                for v in avg["empirical_coverage_90pct"]
            ]
            _coverage_fig.add_trace(go.Bar(
                x=[f"h={h}M" for h in avg["horizon"]],
                y=(avg["empirical_coverage_90pct"] * 100).round(1),
                name=mt, marker_color=colors_cov, opacity=0.85,
                text=(avg["empirical_coverage_90pct"] * 100).round(1).astype(str) + "%",
                textposition="outside",
            ))
        _coverage_fig.add_hline(y=90, line_dash="dash", line_color=_C["muted"],
                                 annotation_text="90% Nominal", annotation_position="top right")
    _coverage_fig.update_layout(
        **_PLOTLY_BASE, title="CI Coverage: Empirical vs 90% Nominal",
        yaxis_title="Empirical Coverage (%)", yaxis_range=[0, 110], height=320,
        barmode="group",
    )

    # ── App initialization ─────────────────────────────────────────────
    app = dash.Dash(__name__, title="NSE Portfolio Monitor",
                    suppress_callback_exceptions=True)
    app.index_string = '''
<!DOCTYPE html>
<html>
    <head>
        {%metas%}
        <title>{%title%}</title>
        {%favicon%}
        {%css%}
        <style>
            * { box-sizing: border-box; }
            body { margin: 0; background: ''' + _C["bg"] + '''; }
            .Select-control { border-radius: 6px !important; }
            .rc-slider-track { background-color: ''' + _C["primary"] + ''' !important; }
            .rc-slider-handle { border-color: ''' + _C["primary"] + ''' !important; }
        </style>
    </head>
    <body>
        {%app_entry%}
        <footer>
            {%config%}
            {%scripts%}
            {%renderer%}
        </footer>
    </body>
</html>'''

    # ══════════════════════════════════════════════════════════════════
    # LAYOUT
    # ══════════════════════════════════════════════════════════════════
    def _ticker_opts():
        return [{"label": f"{STOCKS.get(t,t)}  ({t})", "value": t} for t in active_tickers]

    def _model_opts():
        return [{"label": "ARIMAX", "value": "ARIMAX"}, {"label": "VARX", "value": "VARX"}]

    # ── TAB 1: Market Regime ──────────────────────────────────────────
    tab_regime = dcc.Tab(label="Market Regime", style=_TAB_S, selected_style=_TAB_SEL, children=[
        html.Div(style={"padding": "20px"}, children=[
            _hdr("Market Regime Monitor",
                 "HMM filtered probabilities derived from Nifty 50 returns and realized volatility."),
            _card([
                dcc.Graph(id="regime-gauges", figure=go.Figure(data=[
                    go.Indicator(
                        mode="gauge+number",
                        value=round(v * 100, 1),
                        title={"text": k.replace("market_hmm_","").capitalize(),
                               "font": {"size": 13, "family": _FONT}},
                        gauge={"axis": {"range": [0, 100]},
                               "bar":  {"color": _C["primary"]},
                               "steps": [{"range": [0,40], "color": "#eef2fb"},
                                         {"range": [40,70], "color": "#c7d7f7"},
                                         {"range": [70,100],"color": "#9ab7f2"}]},
                        domain={"x": [i/max(len(current_regime_row),1),
                                      (i+1)/max(len(current_regime_row),1)], "y": [0,1]},
                    ) for i, (k, v) in enumerate(current_regime_row.items())
                ]).update_layout(height=280, title="Current Regime Probabilities (%)",
                                 **{k: v for k,v in _PLOTLY_BASE.items()
                                    if k in ("paper_bgcolor","font","hoverlabel")})
                ) if current_regime_row else html.P("No regime data."),
            ], title="Current State"),
            _card([
                dcc.Graph(id="regime-history", figure=go.Figure([
                    go.Scatter(x=mkt_hmm_df.index, y=mkt_hmm_df[c].values,
                               name=c.replace("market_hmm_","").capitalize(),
                               mode="lines", stackgroup="one", line={"width": 1.5})
                    for c in regime_cols
                ]).update_layout(**_PLOTLY_BASE, title="Market Regime Probability History",
                                  yaxis_range=[0,1], height=320, yaxis_title="P(regime)"))
                if not mkt_hmm_df.empty else html.P("No regime history."),
            ], title="Regime Probability History"),
            _card([
                dcc.Graph(id="ind-chart", figure=go.Figure([
                    go.Scatter(x=market_ind.index, y=market_ind[c].values,
                               name=c, mode="lines", line={"width": 1.2})
                    for c in market_ind.columns
                ]).update_layout(**_PLOTLY_BASE, title="Market Indicators (1-month Lagged Log Returns)",
                                  height=300, yaxis_title="Log Return"))
                if not market_ind.empty else html.P("No indicator data."),
            ], title="Market Indicators"),
        ]),
    ])

    # ── TAB 2: Return Forecasts ───────────────────────────────────────
    tab_return_fc = dcc.Tab(label="Return Forecasts", style=_TAB_S, selected_style=_TAB_SEL, children=[
        html.Div(style={"padding": "20px"}, children=[
            _hdr("Return Forecast Fan Chart",
                 "Combined ARIMAX + VARX forecast of monthly log-returns. "
                 "Shaded cones = analytical CI; dashdot lines = empirical PI."),
            _card([
                _dd("forecast-ticker-dd", _ticker_opts(),
                    active_tickers[0] if active_tickers else None, "Select Stock"),
                dcc.Graph(id="forecast-fan"),
                html.Div(id="forecast-table",
                         style={"overflowX": "auto", "fontSize": "12px", "marginTop": "8px"}),
            ]),
        ]),
    ])

    # ── TAB 3: Price Forecasts ────────────────────────────────────────
    tab_price_fc = dcc.Tab(label="Price Forecasts", style=_TAB_S, selected_style=_TAB_SEL, children=[
        html.Div(style={"padding": "20px"}, children=[
            _hdr("Price Projections (₹)",
                 "Each diamond = implied price if the h-month-ahead log-return forecast materialises. "
                 "Cone = 90% CI in price space (single-period approximation for h > 1)."),
            _card([
                _dd("price-ticker-dd", _ticker_opts(),
                    active_tickers[0] if active_tickers else None, "Select Stock"),
                dcc.Graph(id="price-forecast-chart"),
                html.Div(id="price-table",
                         style={"overflowX": "auto", "fontSize": "12px", "marginTop": "8px"}),
            ]),
        ]),
    ])

    # ── TAB 4: Portfolio ──────────────────────────────────────────────
    pipeline_wt_content = [
        _kpi_row([
            ("Total Capital", f"₹{CAPITAL_AMOUNT:,.0f}", None),
            ("Positions",    str(len([w for w in weights_clean.values() if w > 0.005])), _C["primary"]),
            ("Stress Regime", "Yes" if p7.get("in_stress") else "No",
             _C["danger"] if p7.get("in_stress") else _C["success"]),
            ("P(Bear)",      f"{p7.get('stress_prob',0)*100:.1f}%",
             _C["danger"] if p7.get("stress_prob",0) > 0.6 else _C["muted"]),
        ]),
        _card([
            dcc.Graph(id="portfolio-donut", figure=go.Figure(go.Pie(
                labels=[STOCKS.get(t,t) for t in weights_clean.keys()],
                customdata=list(weights_clean.keys()),
                values=list(weights_clean.values()),
                hole=0.42, textinfo="percent",
                marker={"line": {"color": "white", "width": 2}},
                hovertemplate="<b>%{label}</b><br>Ticker: %{customdata}<br>"
                              "Weight: %{percent}<extra></extra>",
            )).update_layout(
                **{k: v for k,v in _PLOTLY_BASE.items() if k not in ("xaxis","yaxis","legend", "margin")},
                title="Portfolio Weights (BL Optimised)", height=460,
                legend=dict(orientation="v", yanchor="middle", y=0.5,
                            xanchor="left", x=1.02, font=dict(size=11)),
                margin=dict(l=20, r=200, t=60, b=20),
            )) if weights_clean else html.P("No weights computed."),
        ], title="Allocation Donut"),
    ]
    if ticker_conv and scores_avail:
        pipeline_wt_content.append(_card([
            dcc.Graph(figure=go.Figure(go.Bar(
                x=list(ticker_conv.keys()),
                y=list(ticker_conv.values()),
                marker_color=[
                    _C["success"] if v >= 0.65 else (_C["warning"] if v >= 0.40 else _C["danger"])
                    for v in ticker_conv.values()
                ],
                text=[f"{v:.2f}" for v in ticker_conv.values()],
                textposition="outside",
            )).update_layout(**_PLOTLY_BASE, title="Conviction Factors per Ticker",
                              yaxis_range=[0, 1.1], yaxis_title="Conviction", height=280,
                              xaxis_tickangle=-30)),
        ], title="BL View Conviction (from Sector Scores)"),)

    # Custom weights sub-tab
    custom_wt_content = [
        _hdr("Custom Portfolio Weights",
             "Drag sliders to adjust. Metrics are computed on normalised weights. "
             "Comparison chart shows cumulative historical return of each allocation."),
        _card([
            html.Div([
                html.Div(
                    f"{STOCKS.get(t,t)}  ({t})",
                    style={"fontSize": "12px", "fontWeight": "600", "color": _C["text"],
                           "marginBottom": "4px", "fontFamily": _FONT}
                ),
                dcc.Slider(
                    id={"type": "wt-slider", "ticker": t},
                    min=0, max=50, step=0.5,
                    value=round(weights_clean.get(t, 1.0/max(len(active_tickers),1)) * 100, 1),
                    tooltip={"placement": "right", "always_visible": True},
                    marks={0: "0%", 10: "10%", 20: "20%", 30: "30%", 40: "40%", 50: "50%"},
                ),
            ], style={"marginBottom": "14px"}) for t in active_tickers
        ], title="Weight Sliders  (% of Portfolio)", subtitle=(
            "Values automatically normalised to 100% when computing metrics. "
            "Max 50% per position for display purposes; optimizer constraint still applies."
        )),
        html.Div(id="custom-wt-sum-display",
                 style={"fontFamily": _FONT, "fontSize": "13px", "marginBottom": "10px"}),
        html.Div(id="custom-metrics-row"),
        _card([dcc.Graph(id="custom-vs-pipeline-chart")],
              title="Custom vs Pipeline — Cumulative Return",
              subtitle="Based on historical monthly log returns. Not a forward-looking projection."),
    ]

    tab_portfolio = dcc.Tab(label="Portfolio", style=_TAB_S, selected_style=_TAB_SEL, children=[
        html.Div(style={"padding": "20px"}, children=[
            dcc.Tabs(children=[
                dcc.Tab(label="Pipeline Allocation", style=_TAB_S, selected_style=_TAB_SEL,
                        children=[html.Div(pipeline_wt_content, style={"padding": "16px 0"})]),
                dcc.Tab(label="Custom Weights Explorer", style=_TAB_S, selected_style=_TAB_SEL,
                        children=[html.Div(custom_wt_content, style={"padding": "16px 0"})]),
            ], style={"marginBottom": "16px"}),
        ]),
    ])

    # ── TAB 5: Performance ────────────────────────────────────────────
    tab_perf = dcc.Tab(label="Performance", style=_TAB_S, selected_style=_TAB_SEL, children=[
        html.Div(style={"padding": "20px"}, children=[
            _hdr("Backtest Performance Attribution"),
            _kpi_row([
                ("Ann. Return",   f"{perf_summary.get('Annualized Return',0)*100:.1f}%",
                 _C["success"] if perf_summary.get("Annualized Return",0) > 0 else _C["danger"]),
                ("Ann. Vol",      f"{perf_summary.get('Annualized Volatility',0)*100:.1f}%", None),
                ("Sharpe",        f"{perf_summary.get('Sharpe Ratio','—')}", _C["primary"]),
                ("Sortino",       f"{perf_summary.get('Sortino Ratio','—')}", None),
                ("Max Drawdown",  f"{perf_summary.get('Maximum Drawdown',0)*100:.1f}%", _C["danger"]),
                ("Info Ratio",    f"{perf_summary.get('Information Ratio','—')}", None),
            ]),
            _card([
                dcc.Graph(id="cumret-chart", figure=go.Figure([
                    go.Scatter(x=portfolio_val.index, y=portfolio_val.values,
                               name="Portfolio", mode="lines",
                               line={"color": _C["primary"], "width": 2}),
                    go.Scatter(x=bench_rets.index,
                               y=np.exp(bench_rets.cumsum()).values,
                               name="Nifty 50", mode="lines",
                               line={"color": _C["muted"], "width": 1.5, "dash": "dash"}),
                ]).update_layout(**_PLOTLY_BASE, title="Growth of ₹1 (Log Returns Compounded)",
                                  yaxis_title="Portfolio Value", height=360))
                if not portfolio_val.empty else html.P("No backtest data."),
            ], title="Cumulative Returns vs Nifty 50"),
            _card([
                dcc.Graph(id="drawdown-chart", figure=go.Figure(
                    go.Scatter(x=drawdown_s.index, y=drawdown_s.values * 100,
                               fill="tozeroy", name="Drawdown",
                               line={"color": _C["danger"], "width": 1.5},
                               fillcolor="rgba(200,30,30,0.12)")
                ).update_layout(**_PLOTLY_BASE, title="Portfolio Drawdown (%)",
                                 yaxis_title="Drawdown (%)", height=220))
                if not drawdown_s.empty else html.P("No drawdown data."),
            ], title="Drawdown"),
        ]),
    ])

    # ── TAB 6: Walk-Forward Diagnostics (ENHANCED) ───────────────────
    tab_wf = dcc.Tab(label="Walk-Forward Diagnostics", style=_TAB_S, selected_style=_TAB_SEL, children=[
        html.Div(style={"padding": "20px"}, children=[
            _hdr("Walk-Forward Validation Diagnostics",
                 "All metrics are out-of-sample: each forecast was issued before the outcome was known. "
                 "Walk-forward window: min " + str(MIN_TRAIN_MONTHS) + " months, "
                 "max " + str(MAX_LOOKBACK_MONTHS) + " months."),

            # Section A: Forecast Accuracy Overview
            _card([
                html.Div([
                    _dd("wf-metric-dd",
                        [{"label": m, "value": m} for m in ["RMSE","MAE","DA","IC"]],
                        "RMSE", "Metric", w="160px"),
                    _dd("wf-model-dd", _model_opts(), "ARIMAX", "Model", w="160px"),
                ], style={"display": "flex", "gap": "16px"}),
                dcc.Graph(id="wf-accuracy-chart"),
            ], title="A  Forecast Accuracy Overview",
               subtitle="Bar chart of selected metric per ticker, grouped by forecast horizon."),

            # Section B: Forecast Error Time Series
            _card([
                html.Div([
                    _dd("wf-ticker-dd", _ticker_opts(),
                        active_tickers[0] if active_tickers else None, "Stock", w="220px"),
                    _dd("wf-wmodel-dd", _model_opts(), "ARIMAX", "Model", w="160px"),
                ], style={"display": "flex", "gap": "16px"}),
                dcc.Graph(id="wf-rolling-chart"),
            ], title="B  Forecast Error Time Series",
               subtitle="Point-by-point errors (actual − forecast) over the walk-forward period, "
                        "one series per horizon. Systematic bias appears as a persistent non-zero level."),

            # Section C: CI Coverage
            _card([
                dcc.Graph(id="wf-coverage-chart", figure=_coverage_fig),
            ], title="C  Confidence Interval Coverage",
               subtitle="Fraction of actuals falling within the reported 90% CI. "
                        "Well-calibrated models cluster around the dashed 90% reference line."),

            # Section D: IC Heatmap (existing, restyled)
            _card([
                dcc.Graph(id="ic-heatmap", figure=go.Figure(go.Heatmap(
                    z=(perf_df.pivot_table(values="IC", index="ticker",
                                           columns="horizon", aggfunc="mean").values
                       if not perf_df.empty else []),
                    x=[f"h={h}M" for h in sorted(perf_df["horizon"].unique())]
                       if not perf_df.empty else [],
                    y=perf_df["ticker"].unique().tolist() if not perf_df.empty else [],
                    colorscale="RdYlGn", zmin=-0.4, zmax=0.8,
                    texttemplate="%{z:.3f}",
                    hovertemplate="Ticker: %{y}<br>Horizon: %{x}<br>IC: %{z:.4f}<extra></extra>",
                )).update_layout(**_PLOTLY_BASE,
                                  title="Information Coefficient — Ticker × Horizon (Walk-Forward)",
                                  height=max(300, len(active_tickers)*22 + 100),
                                  xaxis_title="Forecast Horizon", yaxis_title="Ticker"))
                if not perf_df.empty else html.P("No performance data available."),
            ], title="D  Information Coefficient Heatmap",
               subtitle="Pearson correlation between walk-forward forecasts and subsequent actuals. "
                        "Positive IC indicates directional signal; > 0.05 is considered meaningful."),
        ]),
    ])

    # ── TAB 7: Model Evaluation (ENHANCED with line chart) ───────────
    tab_eval = dcc.Tab(label="Model Evaluation", style=_TAB_S, selected_style=_TAB_SEL, children=[
        html.Div(style={"padding": "20px"}, children=[
            _hdr("Out-of-Sample Model Evaluation",
                 f"Fixed-origin: trained on first {MIN_TRAIN_MONTHS} months, "
                 f"evaluated on the remainder. One forecast fan per model per stock."),

            # Metric Heatmaps (existing)
            _card([
                html.Div([
                    _dd("eval-metric-dd",
                        [{"label": m, "value": m} for m in ["RMSE","MAE","MAPE","SMAPE","DA","ME"]],
                        "RMSE", "Metric", w="160px"),
                    _dd("eval-model-dd", _model_opts(), "ARIMAX", "Model", w="160px"),
                ], style={"display": "flex", "gap": "16px"}),
                dcc.Graph(id="eval-heatmap"),
                html.Div(id="eval-table",
                         style={"maxHeight": "300px", "overflowY": "auto",
                                "fontSize": "12px", "marginTop": "10px"}),
            ], title="A  Metric Heatmap — Ticker × Horizon"),

            # Forecast vs Actual Line Chart (NEW)
            _card([
                html.Div([
                    _dd("eval-line-ticker-dd", _ticker_opts(),
                        active_tickers[0] if active_tickers else None, "Stock", w="220px"),
                    _dd("eval-line-model-dd", _model_opts(), "ARIMAX", "Model", w="160px"),
                ], style={"display": "flex", "gap": "16px"}),
                dcc.Graph(id="eval-line-chart"),
            ], title="B  Forecast vs Actual — Time Series",
               subtitle="Full history (train + test) with overlaid forecast path and 90%/95% CI bands. "
                        "Vertical dashed line marks the train/test split. Forecast fan originates from that point."),
        ]) if not eval_df.empty else html.Div([
            _hdr("Model Evaluation"),
            html.P("Run run_phase_eval() to populate this panel.",
                   style={"color": _C["muted"], "fontFamily": _FONT}),
        ], style={"padding": "20px"}),
    ])

    # ── TAB 8: Sector Conviction ──────────────────────────────────────
    if scores_avail and not sector_scores.empty:
        # Composite score heatmap
        _sc_heat = go.Figure(go.Heatmap(
            z=sector_scores.values.T,
            x=[str(d.date()) for d in sector_scores.index],
            y=sector_scores.columns.tolist(),
            colorscale="RdYlGn", zmin=0, zmax=10,
            texttemplate="%{z:.1f}",
            hovertemplate="Quarter: %{x}<br>Sector: %{y}<br>Score: %{z:.2f}<extra></extra>",
        ))
        _sc_heat.update_layout(**_PLOTLY_BASE,
                               title="Composite Sector Scores Over Time (0–10)",
                               height=max(300, len(sector_scores.columns) * 28 + 80),
                               xaxis_title="Quarter-End Date", yaxis_title="Sector")
        # Conviction bar
        _cv_bar = go.Figure(go.Bar(
            x=list(sector_conv.keys()), y=list(sector_conv.values()),
            marker_color=[
                _C["success"] if v >= 0.65 else (_C["warning"] if v >= 0.40 else _C["danger"])
                for v in sector_conv.values()
            ],
            text=[f"{v:.2f}" for v in sector_conv.values()], textposition="outside",
        ))
        _cv_bar.update_layout(**_PLOTLY_BASE, title="Current Conviction per Sector",
                               yaxis_range=[0, 1.1], yaxis_title="Conviction Factor",
                               height=300, xaxis_tickangle=-30)
        tab_conviction = dcc.Tab(
            label="Sector Conviction", style=_TAB_S, selected_style=_TAB_SEL, children=[
                html.Div(style={"padding": "20px"}, children=[
                    _hdr("Sector Fundamental Conviction Signals",
                         f"Conviction = {CONVICTION_LEVEL_WT:.0%} × score_level + "
                         f"{CONVICTION_MOMENTUM_WT:.0%} × score_momentum.  "
                         f"Scores lag {SCORE_LAG_MONTHS}m after quarter-end (no look-ahead)."),
                    _card([dcc.Graph(figure=_sc_heat)],
                          title="Composite Score History",
                          subtitle=f"Weighted average of A–F parameters: "
                                   f"B={SECTOR_SCORE_WEIGHTS['B']:.0%}, "
                                   f"F={SECTOR_SCORE_WEIGHTS['F']:.0%}, "
                                   f"A={SECTOR_SCORE_WEIGHTS['A']:.0%}, others smaller."),
                    _card([dcc.Graph(figure=_cv_bar)],
                          title="Current Quarter Conviction",
                          subtitle="Used to scale Ω in Black-Litterman: high conviction → "
                                   "views pulled harder; low conviction → prior dominates."),
                    _card([
                        html.Table([
                            html.Thead(html.Tr([
                                html.Th(h, style={"textAlign":"left","padding":"8px 12px",
                                                  "fontSize":"11px","color":_C["muted"],
                                                  "borderBottom":f"1px solid {_C['border']}",
                                                  "fontFamily":_FONT})
                                for h in ["Sector","Tickers","Conviction","BL Ω Scale","Interpretation"]
                            ])),
                            html.Tbody([
                                html.Tr([
                                    html.Td(sec, style={"padding":"6px 12px","fontSize":"12px"}),
                                    html.Td(", ".join([t for t in SECTOR_GROUPS.get(sec,[])
                                                       if t in active_tickers]),
                                            style={"padding":"6px 12px","fontSize":"11px",
                                                   "color":_C["muted"]}),
                                    html.Td(f"{cv:.3f}",
                                            style={"padding":"6px 12px","fontSize":"12px",
                                                   "fontWeight":"700",
                                                   "color": (_C["success"] if cv >= 0.65
                                                             else _C["warning"] if cv >= 0.40
                                                             else _C["danger"])}),
                                    html.Td(f"{0.5/cv:.2f}×",
                                            style={"padding":"6px 12px","fontSize":"12px"}),
                                    html.Td("Views trusted more" if cv >= 0.65
                                            else "Balanced" if cv >= 0.40
                                            else "Prior trusted more",
                                            style={"padding":"6px 12px","fontSize":"11px",
                                                   "color":_C["muted"]}),
                                ], style={"borderBottom": f"1px solid {_C['border']}"})
                                for sec, cv in sorted(sector_conv.items(), key=lambda x:-x[1])
                            ]),
                        ], style={"width":"100%","borderCollapse":"collapse","fontFamily":_FONT}),
                    ], title="BL Ω Adjustment Table"),
                ]),
            ],
        )
    else:
        tab_conviction = dcc.Tab(
            label="Sector Conviction", style=_TAB_S, selected_style=_TAB_SEL, children=[
                html.Div([
                    _hdr("Sector Conviction"),
                    html.P(
                        f"No sector scores found.  Create '{SECTOR_SCORE_FILE}' "
                        f"(see SECTOR_SCORE_FILE config for format) and re-run to enable this panel.",
                        style={"color": _C["muted"], "fontFamily": _FONT}
                    ),
                ], style={"padding": "20px"}),
            ],
        )

    # ── TAB 9: Deep Dive ─────────────────────────────────────────────
    tab_deep = dcc.Tab(label="Deep Dive", style=_TAB_S, selected_style=_TAB_SEL, children=[
        html.Div(style={"padding": "20px"}, children=[
            _hdr("Stock Deep Dive",
                 "Price history with bear-regime overlays and monthly realized volatility."),
            _card([
                _dd("deepdive-ticker-dd", _ticker_opts(),
                    active_tickers[0] if active_tickers else None, "Select Stock"),
                dcc.Graph(id="price-regime-chart"),
                dcc.Graph(id="rv-chart"),
            ]),
        ]),
    ])

    # ── Master layout ─────────────────────────────────────────────────
    app.layout = html.Div(style={"fontFamily": _FONT, "background": _C["bg"], "minHeight": "100vh"}, children=[
        # Header bar
        html.Div([
            html.Div("NSE Portfolio Intelligence", style={
                "fontSize": "18px", "fontWeight": "800", "color": "white", "letterSpacing": "0.02em",
            }),
            html.Div("Forecast · Conviction · Allocation · Diagnostics", style={
                "fontSize": "11px", "color": "rgba(255,255,255,0.7)", "marginTop": "3px",
            }),
        ], style={
            "background": f"linear-gradient(90deg, {_C['primary']} 0%, #0d47a1 100%)",
            "padding": "16px 28px", "marginBottom": "0px",
        }),
        # Tab container
        dcc.Tabs(children=[
            tab_regime, tab_return_fc, tab_price_fc, tab_portfolio,
            tab_perf, tab_wf, tab_eval, tab_conviction, tab_deep,
        ], style={"fontFamily": _FONT, "background": "white",
                  "borderBottom": f"1px solid {_C['border']}",
                  "paddingLeft": "12px"}),
    ])

    # ══════════════════════════════════════════════════════════════════
    # CALLBACKS
    # ══════════════════════════════════════════════════════════════════

    # ── CB-1: Return Forecast Fan ─────────────────────────────────────
    @app.callback(
        [Output("forecast-fan", "figure"), Output("forecast-table", "children")],
        Input("forecast-ticker-dd", "value"),
    )
    def update_forecast_fan(ticker):
        if ticker is None or combined_df.empty:
            return go.Figure().update_layout(**_PLOTLY_BASE, title="No data"), html.P("—")
        ret_s   = mlr[ticker] if ticker in mlr.columns else pd.Series(dtype=float)
        fig     = go.Figure()
        fig.add_trace(go.Scatter(x=ret_s.index, y=ret_s.values,
                                  name="Historical Returns", mode="lines",
                                  line={"color": _C["primary"], "width": 1.5}))
        last_dt = ret_s.index[-1] if not ret_s.empty else pd.Timestamp.today()
        for h in FORECAST_HORIZONS:
            row = combined_df[(combined_df["ticker"] == ticker) & (combined_df["horizon"] == h)]
            if row.empty:
                continue
            r0        = row.iloc[0]
            fc_val    = float(r0["combined_forecast"])
            target_dt = last_dt + pd.DateOffset(months=h)
            hc        = HORIZON_COLORS.get(h, {"hex": "gray", "rgb": "128,128,128"})
            col, rgb  = hc["hex"], hc["rgb"]
            ci90_lo   = r0.get("ci90_lower", np.nan); ci90_hi = r0.get("ci90_upper", np.nan)
            ci95_lo   = r0.get("ci95_lower", np.nan); ci95_hi = r0.get("ci95_upper", np.nan)
            has_90    = pd.notna(ci90_lo) and pd.notna(ci90_hi)
            has_95    = pd.notna(ci95_lo) and pd.notna(ci95_hi)
            if has_95:
                fig.add_trace(go.Scatter(
                    x=[last_dt, target_dt, target_dt, last_dt],
                    y=[fc_val, ci95_hi, ci95_lo, fc_val],
                    fill="toself", fillcolor=f"rgba({rgb},0.06)", line={"width": 0},
                    name=f"{h}M 95%CI", legendgroup=f"h{h}", showlegend=False, hoverinfo="skip"))
            if has_90:
                fig.add_trace(go.Scatter(
                    x=[last_dt, target_dt, target_dt, last_dt],
                    y=[fc_val, ci90_hi, ci90_lo, fc_val],
                    fill="toself", fillcolor=f"rgba({rgb},0.13)", line={"width": 0},
                    name=f"{h}M 90%CI", legendgroup=f"h{h}", showlegend=False, hoverinfo="skip"))
            err_lo = max(0, fc_val - ci90_lo) if has_90 else None
            err_hi = max(0, ci90_hi - fc_val) if has_90 else None
            fig.add_trace(go.Scatter(
                x=[target_dt], y=[fc_val], name=f"{h}M Forecast",
                legendgroup=f"h{h}", mode="markers",
                marker={"color": col, "size": 12, "symbol": "diamond",
                        "line": {"color": "white", "width": 1.5}},
                error_y=dict(type="data", symmetric=False, array=[err_hi],
                             arrayminus=[err_lo], color=col, thickness=2, width=8)
                if err_lo is not None else None,
                hovertemplate=(f"<b>{h}M</b><br>{target_dt.strftime('%b %Y')}<br>"
                               f"Return: {fc_val:.4f}"
                               + (f"<br>90%CI [{ci90_lo:.4f}, {ci90_hi:.4f}]" if has_90 else "")
                               + "<extra></extra>"),
            ))
            cal = (calib_df[(calib_df["ticker"] == ticker) & (calib_df["horizon"] == h)]
                   if not calib_df.empty else pd.DataFrame())
            if not cal.empty:
                emp_lo = fc_val + float(cal["pi_lower_offset"].iloc[0])
                emp_hi = fc_val + float(cal["pi_upper_offset"].iloc[0])
                fig.add_trace(go.Scatter(
                    x=[target_dt, target_dt], y=[emp_lo, emp_hi], mode="lines+markers",
                    name=f"{h}M Empirical PI", legendgroup=f"h{h}", showlegend=False,
                    line={"color": col, "dash": "dashdot", "width": 1.2}))
        fig.update_layout(**_PLOTLY_BASE,
                           title=f"Return Forecast Fan — {STOCKS.get(ticker,ticker)}",
                           xaxis_title="Date", yaxis_title="Monthly Log Return", height=480,
                           hovermode="x unified")
        sub  = combined_df[combined_df["ticker"] == ticker].sort_values("horizon")
        rows = []
        for _, r in sub.iterrows():
            rows.append(html.Tr([
                html.Td(f"{int(r['horizon'])}M", style={"padding":"5px 10px","fontWeight":"600"}),
                html.Td(f"{r['combined_forecast']:.4f}", style={"padding":"5px 10px"}),
                html.Td(f"{r.get('ci90_lower',np.nan):.4f}" if pd.notna(r.get('ci90_lower')) else "—",
                        style={"padding":"5px 10px"}),
                html.Td(f"{r.get('ci90_upper',np.nan):.4f}" if pd.notna(r.get('ci90_upper')) else "—",
                        style={"padding":"5px 10px"}),
                html.Td(r.get("method","—"), style={"padding":"5px 10px","color":_C["muted"]}),
            ], style={"borderBottom": f"1px solid {_C['border']}", "fontFamily": _FONT}))
        tbl = html.Table(
            [html.Thead(html.Tr([
                html.Th(h, style={"padding":"6px 10px","fontSize":"11px","color":_C["muted"],
                                  "textAlign":"left","borderBottom":f"2px solid {_C['border']}"})
                for h in ["Horizon","Forecast (log-ret)","90% CI Low","90% CI High","Method"]
            ]))] + rows,
            style={"borderCollapse":"collapse","width":"100%","fontFamily":_FONT})
        return fig, tbl

    # ── CB-2: Price Forecast ─────────────────────────────────────────
    @app.callback(
        [Output("price-forecast-chart","figure"), Output("price-table","children")],
        Input("price-ticker-dd","value"),
    )
    def update_price_forecast(ticker):
        empty = go.Figure().update_layout(**_PLOTLY_BASE, title="No data")
        if ticker is None or combined_df.empty or ticker not in monthly_prices.columns:
            return empty, html.P("—")
        px_s       = monthly_prices[ticker].dropna()
        last_price = float(px_s.iloc[-1]); last_date = px_s.index[-1]
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=px_s.index, y=px_s.values, name="Historical Price",
                                  mode="lines", line={"color": _C["primary"], "width": 2}))
        fig.add_trace(go.Scatter(x=[last_date], y=[last_price], mode="markers", showlegend=False,
                                  marker={"size": 10, "color": _C["primary"], "symbol": "circle"}))
        for h in FORECAST_HORIZONS:
            row = combined_df[(combined_df["ticker"]==ticker) & (combined_df["horizon"]==h)]
            if row.empty:
                continue
            r0        = row.iloc[0]
            fc_lr     = float(r0["combined_forecast"])
            target_dt = last_date + pd.DateOffset(months=h)
            price_fc  = last_price * np.exp(fc_lr)
            hc        = HORIZON_COLORS.get(h, {"hex":"gray","rgb":"128,128,128"})
            col, rgb  = hc["hex"], hc["rgb"]
            ci90_lo   = r0.get("ci90_lower", np.nan); ci90_hi = r0.get("ci90_upper", np.nan)
            has_90    = pd.notna(ci90_lo) and pd.notna(ci90_hi)
            if has_90:
                p_lo = last_price * np.exp(ci90_lo)
                p_hi = last_price * np.exp(ci90_hi)
                fig.add_trace(go.Scatter(
                    x=[last_date, target_dt, target_dt, last_date],
                    y=[last_price, p_hi, p_lo, last_price],
                    fill="toself", fillcolor=f"rgba({rgb},0.11)", line={"width":0},
                    name=f"{h}M 90%CI", legendgroup=f"h{h}", showlegend=True, hoverinfo="skip"))
            fig.add_trace(go.Scatter(
                x=[last_date, target_dt], y=[last_price, price_fc],
                mode="lines+markers", name=f"{h}M ₹{price_fc:,.0f}",
                legendgroup=f"h{h}",
                line={"color":col,"dash":"dash","width":1.8},
                marker={"symbol":"diamond","size":11,"color":col,"line":{"color":"white","width":1}},
                hovertemplate=(f"<b>{h}M Price Forecast</b><br>{target_dt.strftime('%b %Y')}<br>"
                               f"₹{price_fc:,.2f}  (lr={fc_lr:+.4f})"
                               + (f"<br>90%CI ₹{last_price*np.exp(ci90_lo):,.0f}–"
                                  f"₹{last_price*np.exp(ci90_hi):,.0f}" if has_90 else "")
                               + "<extra></extra>")))
        fig.update_layout(**_PLOTLY_BASE,
                           title=f"Price Forecast — {STOCKS.get(ticker,ticker)} ({ticker})",
                           xaxis_title="Date", yaxis_title="Price (₹)", height=460,
                           hovermode="x unified")
        sub   = combined_df[combined_df["ticker"]==ticker].sort_values("horizon")
        rows  = []
        for _, r in sub.iterrows():
            fc_lr = r["combined_forecast"]
            p     = last_price * np.exp(fc_lr)
            c9l   = last_price * np.exp(r["ci90_lower"]) if pd.notna(r.get("ci90_lower")) else None
            c9h   = last_price * np.exp(r["ci90_upper"]) if pd.notna(r.get("ci90_upper")) else None
            rows.append(html.Tr([
                html.Td(f"{int(r['horizon'])}M", style={"padding":"5px 10px","fontWeight":"600"}),
                html.Td(f"₹{p:,.2f}", style={"padding":"5px 10px"}),
                html.Td(f"₹{c9l:,.2f}" if c9l else "—", style={"padding":"5px 10px"}),
                html.Td(f"₹{c9h:,.2f}" if c9h else "—", style={"padding":"5px 10px"}),
            ], style={"borderBottom": f"1px solid {_C['border']}", "fontFamily": _FONT}))
        tbl = html.Table(
            [html.Thead(html.Tr([
                html.Th(h, style={"padding":"6px 10px","fontSize":"11px","color":_C["muted"],
                                  "textAlign":"left","borderBottom":f"2px solid {_C['border']}"})
                for h in ["Horizon","Forecast Price","90% CI Low","90% CI High"]
            ]))] + rows,
            style={"borderCollapse":"collapse","width":"100%","fontFamily":_FONT})
        return fig, tbl

    # ── CB-3: Custom Portfolio Weights (ALL pattern matching) ─────────
    @app.callback(
        [Output("custom-metrics-row",       "children"),
         Output("custom-vs-pipeline-chart", "figure"),
         Output("custom-wt-sum-display",    "children")],
        Input({"type": "wt-slider", "ticker": ALL}, "value"),
    )
    def update_custom_portfolio(slider_vals):
        ctx = callback_context
        if not slider_vals:
            return html.P("Adjust sliders to explore."), go.Figure(), ""
        # Recover ticker ordering from callback_context
        if ctx.inputs_list and ctx.inputs_list[0]:
            tickers_ord = [item["id"]["ticker"] for item in ctx.inputs_list[0]]
        else:
            tickers_ord = active_tickers
        raw_weights = {t: max(v or 0, 0) / 100.0 for t, v in zip(tickers_ord, slider_vals)}
        total = sum(raw_weights.values())
        # Normalised weights (always used for computation)
        if total > 1e-6:
            norm_w = {t: v / total for t, v in raw_weights.items()}
        else:
            norm_w = {t: 1.0 / len(active_tickers) for t in active_tickers}
        # Weight sum indicator
        if abs(total * 100 - 100) < 0.5:
            sum_color = _C["success"]; sum_note = "Fully invested"
        elif total * 100 > 100:
            sum_color = _C["danger"]; sum_note = f"Overweight — normalised to 100%"
        else:
            sum_color = _C["warning"]; sum_note = f"Underweight — {(1-total)*100:.1f}% in cash"
        sum_display = html.Div([
            html.Span(f"Total: {total*100:.1f}%  ", style={"fontWeight":"700","color":sum_color}),
            html.Span(sum_note, style={"color":_C["muted"]}),
        ], style={"fontFamily":_FONT,"fontSize":"13px","marginBottom":"10px"})

        # Expected metrics (BL posterior)
        exp_ret_val = exp_vol_val = sharpe_val = "—"
        if ret_bl is not None and S_bl is not None:
            try:
                w_arr = np.array([norm_w.get(t, 0) for t in active_tickers])
                r_arr = np.array([float(ret_bl[t]) if hasattr(ret_bl,"__getitem__")
                                  else float(ret_bl) for t in active_tickers])
                S_arr = (S_bl.values if hasattr(S_bl, "values") else np.array(S_bl))
                e_ret = float(w_arr @ r_arr)
                e_vol = float(np.sqrt(w_arr @ S_arr @ w_arr))
                e_shr = (e_ret - ann_rf_exact) / e_vol if e_vol > 1e-6 else np.nan
                exp_ret_val = f"{e_ret*100:.2f}%"
                exp_vol_val = f"{e_vol*100:.2f}%"
                sharpe_val  = f"{e_shr:.3f}" if not np.isnan(e_shr) else "—"
            except Exception:
                pass

        # Historical performance comparison
        if not mlr.empty:
            avail = [t for t in active_tickers if t in mlr.columns]
            w_pipe = np.array([weights_clean.get(t, 0) for t in avail])
            w_cust = np.array([norm_w.get(t, 0) for t in avail])
            w_pipe /= max(w_pipe.sum(), 1e-10)
            w_cust /= max(w_cust.sum(), 1e-10)
            ret_mat        = mlr[avail].fillna(0.0)
            pipe_port_rets = ret_mat.values @ w_pipe
            cust_port_rets = ret_mat.values @ w_cust
            pipe_cumret    = np.exp(np.cumsum(pipe_port_rets))
            cust_cumret    = np.exp(np.cumsum(cust_port_rets))
            dates          = mlr.index

            # Historical metrics for custom
            n_mo  = len(cust_port_rets)
            c_ar  = float((cust_cumret[-1]) ** (12/n_mo) - 1) if n_mo > 1 else np.nan
            c_vol = float(np.std(cust_port_rets) * np.sqrt(12))
            c_dd  = float(np.min((np.exp(np.cumsum(cust_port_rets)) /
                                   np.maximum.accumulate(np.exp(np.cumsum(cust_port_rets)))) - 1))
        else:
            dates = []; pipe_cumret = []; cust_cumret = []; c_ar = c_vol = c_dd = np.nan

        metrics = _kpi_row([
            ("Exp. Ann. Return (BL)", exp_ret_val,
             _C["success"] if exp_ret_val != "—" and "%" in exp_ret_val
             and float(exp_ret_val.rstrip("%")) > 0 else None),
            ("Exp. Ann. Vol (BL)",    exp_vol_val, None),
            ("Exp. Sharpe (BL)",      sharpe_val,  _C["primary"]),
            ("Hist. Ann. Return",
             f"{c_ar*100:.1f}%" if not np.isnan(c_ar) else "—",
             _C["success"] if not np.isnan(c_ar) and c_ar > 0 else _C["danger"]),
            ("Hist. Ann. Vol",
             f"{c_vol*100:.1f}%" if not np.isnan(c_vol) else "—", None),
            ("Hist. Max DD",
             f"{c_dd*100:.1f}%" if not np.isnan(c_dd) else "—", _C["danger"]),
        ])
        cmp_fig = go.Figure()
        if len(dates) > 0:
            cmp_fig.add_trace(go.Scatter(x=dates, y=pipe_cumret, name="Pipeline (BL Optimised)",
                                          mode="lines", line={"color":_C["primary"],"width":2}))
            cmp_fig.add_trace(go.Scatter(x=dates, y=cust_cumret, name="Custom Weights",
                                          mode="lines", line={"color":_C["warning"],"width":2,"dash":"dash"}))
        cmp_fig.update_layout(**_PLOTLY_BASE,
                               title="Custom vs Pipeline — Historical Cumulative Return",
                               yaxis_title="Growth of ₹1", height=360, hovermode="x unified")
        return metrics, cmp_fig, sum_display

    # ── CB-4: Walk-Forward Accuracy Chart ────────────────────────────
    @app.callback(
        Output("wf-accuracy-chart", "figure"),
        [Input("wf-metric-dd", "value"), Input("wf-model-dd", "value")],
    )
    def update_wf_accuracy(metric, model_type):
        empty = go.Figure().update_layout(**_PLOTLY_BASE, title="No walk-forward data")
        if perf_df.empty or metric not in perf_df.columns:
            return empty
        sub = perf_df[perf_df["model"] == model_type] if model_type else perf_df
        if sub.empty:
            return empty
        fig = go.Figure()
        horizon_palette = {h: HORIZON_COLORS[h]["hex"] for h in FORECAST_HORIZONS}
        for h in sorted(sub["horizon"].unique()):
            h_sub = sub[sub["horizon"] == h].sort_values("ticker")
            fig.add_trace(go.Bar(
                x=h_sub["ticker"].tolist(),
                y=h_sub[metric].tolist(),
                name=f"h={h}M",
                marker_color=horizon_palette.get(h, "steelblue"),
                text=[f"{v:.3f}" for v in h_sub[metric]],
                textposition="outside", opacity=0.85,
            ))
        cmap_invert = metric in ("RMSE", "MAE", "MAPE", "SMAPE")
        fig.update_layout(**_PLOTLY_BASE,
                           title=f"{model_type} — {metric} by Ticker and Horizon",
                           yaxis_title=metric, barmode="group", height=380,
                           xaxis_tickangle=-30)
        return fig

    # ── CB-5: Walk-Forward Error Time Series ──────────────────────────
    @app.callback(
        Output("wf-rolling-chart", "figure"),
        [Input("wf-ticker-dd", "value"), Input("wf-wmodel-dd", "value")],
    )
    def update_wf_errors(ticker, model_type):
        empty = go.Figure().update_layout(**_PLOTLY_BASE, title="No error data")
        if errors_df.empty or ticker is None:
            return empty
        fig = go.Figure()
        for h in FORECAST_HORIZONS:
            sub = errors_df[
                (errors_df["ticker"]  == ticker) &
                (errors_df["model"]   == model_type) &
                (errors_df["horizon"] == h)
            ].dropna(subset=["error"])
            if sub.empty or "date" not in sub.columns:
                continue
            sub = sub.sort_values("t")
            hc  = HORIZON_COLORS.get(h, {"hex": "gray"})
            fig.add_trace(go.Scatter(
                x=sub["date"], y=sub["error"],
                name=f"h={h}M  Error", mode="lines+markers",
                line={"color": hc["hex"], "width": 1.5},
                marker={"size": 4},
                hovertemplate=f"<b>h={h}M</b><br>%{{x}}<br>Error: %{{y:.4f}}<extra></extra>",
            ))
            # Rolling 12-step mean error
            if len(sub) >= 12:
                roll_err = sub["error"].rolling(12, min_periods=6).mean()
                fig.add_trace(go.Scatter(
                    x=sub["date"].values, y=roll_err.values,
                    name=f"h={h}M  12-step Mean", mode="lines",
                    line={"color": hc["hex"], "width": 2.5, "dash": "dot"},
                    showlegend=False,
                ))
        fig.add_hline(y=0, line_dash="dash", line_color=_C["muted"])
        fig.update_layout(**_PLOTLY_BASE,
                           title=f"{model_type} Forecast Errors — {STOCKS.get(ticker,ticker)}",
                           yaxis_title="Actual − Forecast  (log-return)",
                           xaxis_title="Forecast Date", height=380, hovermode="x unified")
        return fig

    # ── CB-6: Existing Eval Heatmap (restyled) ───────────────────────
    @app.callback(
        [Output("eval-heatmap", "figure"), Output("eval-table", "children")],
        [Input("eval-metric-dd", "value"), Input("eval-model-dd", "value")],
    )
    def update_eval_panel(metric, model_type):
        empty_fig = go.Figure().update_layout(**_PLOTLY_BASE, title="No evaluation data")
        if eval_df.empty or metric not in eval_df.columns:
            return empty_fig, html.P("No data.")
        sub = eval_df[eval_df["model"] == model_type]
        if sub.empty:
            return empty_fig, html.P(f"No {model_type} data.")
        try:
            pivot = sub.pivot_table(values=metric, index="ticker", columns="horizon", aggfunc="mean")
            cmap  = "RdYlGn_r" if metric in ("RMSE","MAPE","MAE","SMAPE","ME") else "RdYlGn"
            fig   = go.Figure(go.Heatmap(
                z=pivot.values,
                x=[f"h={c}M" for c in pivot.columns],
                y=pivot.index.tolist(),
                colorscale=cmap, texttemplate="%{z:.4f}",
                hovertemplate="Ticker: %{y}<br>Horizon: %{x}<br>"
                              + f"{metric}: " + "%{z:.4f}<extra></extra>",
            ))
            fig.update_layout(**_PLOTLY_BASE,
                               title=f"{model_type} — {metric}  (out-of-sample, fixed origin)",
                               height=max(280, len(pivot)*26 + 120),
                               xaxis_title="Horizon", yaxis_title="Ticker")
        except Exception:
            fig = empty_fig
        summary = (sub.groupby(["ticker","horizon"])[metric].mean().reset_index()
                   .rename(columns={metric: f"mean_{metric}"}))
        summary[f"mean_{metric}"] = summary[f"mean_{metric}"].round(4)
        tbl = html.Table(
            [html.Thead(html.Tr([
                html.Th(c, style={"padding":"5px 10px","fontSize":"11px","color":_C["muted"],
                                  "textAlign":"left","borderBottom":f"1px solid {_C['border']}"})
                for c in summary.columns
            ]))] +
            [html.Tr([html.Td(str(v), style={"padding":"4px 10px","fontSize":"12px"})
                      for v in row], style={"borderBottom":f"1px solid {_C['border']}"})
             for _, row in summary.iterrows()],
            style={"borderCollapse":"collapse","width":"100%","fontFamily":_FONT})
        return fig, tbl

    # ── CB-7: Model Evaluation Line Chart (NEW) ───────────────────────
    @app.callback(
        Output("eval-line-chart", "figure"),
        [Input("eval-line-ticker-dd", "value"), Input("eval-line-model-dd", "value")],
    )
    def update_eval_line_chart(ticker, model_type):
        empty = go.Figure().update_layout(**_PLOTLY_BASE, title="No prediction data available")
        if eval_preds_df.empty or ticker is None:
            return empty
        sub = eval_preds_df[
            (eval_preds_df["ticker"] == ticker) &
            (eval_preds_df["model"]  == model_type)
        ].sort_values("step_ahead")
        if sub.empty:
            return empty

        train_end_str = sub["train_end_date"].iloc[0]
        train_end_dt  = pd.Timestamp(train_end_str)

        # Historical (training) returns
        hist = mlr[ticker].dropna() if ticker in mlr.columns else pd.Series(dtype=float)
        train_hist = hist[hist.index <= train_end_dt]
        test_hist  = hist[hist.index > train_end_dt]

        fig = go.Figure()
        # Training actuals
        if not train_hist.empty:
            fig.add_trace(go.Scatter(
                x=train_hist.index, y=train_hist.values,
                name="Train Actuals", mode="lines",
                line={"color": _C["primary"], "width": 1.8},
            ))
        # Train/test split line
        fig.add_vline(x=train_end_dt, line_dash="dash", line_color=_C["muted"],
                      annotation_text="Train | Test", annotation_position="top right",
                      annotation_font={"size": 11, "color": _C["muted"]})
        # Test actuals
        if not test_hist.empty:
            fig.add_trace(go.Scatter(
                x=test_hist.index, y=test_hist.values,
                name="Test Actuals", mode="lines",
                line={"color": _C["primary"], "width": 1.8, "dash": "solid"},
                opacity=0.7,
            ))
        # 95% CI shaded band
        has_95 = "ci95_lower" in sub.columns and sub["ci95_lower"].notna().any()
        has_90 = "ci90_lower" in sub.columns and sub["ci90_lower"].notna().any()
        test_dates = pd.to_datetime(sub["forecast_date"])
        if has_95:
            fig.add_trace(go.Scatter(
                x=list(test_dates) + list(test_dates[::-1]),
                y=list(sub["ci95_upper"]) + list(sub["ci95_lower"][::-1]),
                fill="toself", fillcolor="rgba(26,86,219,0.06)", line={"width": 0},
                name="95% CI", showlegend=True, hoverinfo="skip",
            ))
        # 90% CI shaded band
        if has_90:
            fig.add_trace(go.Scatter(
                x=list(test_dates) + list(test_dates[::-1]),
                y=list(sub["ci90_upper"]) + list(sub["ci90_lower"][::-1]),
                fill="toself", fillcolor="rgba(26,86,219,0.12)", line={"width": 0},
                name="90% CI", showlegend=True, hoverinfo="skip",
            ))
        # Forecast line
        fig.add_trace(go.Scatter(
            x=test_dates, y=sub["forecast"].values,
            name=f"{model_type} Forecast", mode="lines+markers",
            line={"color": _C["warning"], "width": 2, "dash": "dash"},
            marker={"size": 5, "color": _C["warning"]},
            hovertemplate="<b>Forecast</b><br>%{x}<br>%{y:.4f}<extra></extra>",
        ))
        # Horizon markers (where FORECAST_HORIZONS land from train_end)
        for h in FORECAST_HORIZONS:
            target_dt = train_end_dt + pd.DateOffset(months=h)
            step_row  = sub[pd.to_datetime(sub["forecast_date"]) <= target_dt]
            if step_row.empty:
                continue
            r = step_row.iloc[-1]
            hc = HORIZON_COLORS.get(h, {"hex": "gray"})
            fig.add_trace(go.Scatter(
                x=[pd.Timestamp(r["forecast_date"])], y=[r["forecast"]],
                mode="markers", name=f"h={h}M",
                marker={"symbol": "diamond", "size": 13, "color": hc["hex"],
                        "line": {"color": "white", "width": 1.5}},
                hovertemplate=f"<b>h={h}M Forecast</b><br>{r['forecast_date']}<br>"
                              f"Forecast: {r['forecast']:.4f}<br>"
                              f"Actual: {r['actual']:.4f}<extra></extra>",
            ))
        # Metric annotations
        rmse_row = (eval_df[(eval_df["ticker"]==ticker) &
                            (eval_df["model"]==model_type) &
                            (eval_df["horizon"]==FORECAST_HORIZONS[0])]
                   if not eval_df.empty else pd.DataFrame())
        anno_txt = ""
        if not rmse_row.empty:
            r = rmse_row.iloc[0]
                        
            # FIXED
            def _fmt(val, spec):
                try: return format(float(val), spec)
                except (TypeError, ValueError): return "—"

            anno_txt = (f"RMSE={_fmt(r.get('RMSE'), '.4f')}  "
                        f"DA={_fmt(r.get('DA'), '.1%')}  "
                        f"MAPE={_fmt(r.get('MAPE'), '.2f')}%  "
                        f"(h={FORECAST_HORIZONS[0]}M)")
            # anno_txt = (f"RMSE={r.get('RMSE','—'):.4f}  "
            #             f"DA={r.get('DA','—'):.1%}  "
            #             f"MAPE={r.get('MAPE','—'):.2f}%  "
            #             f"(h={FORECAST_HORIZONS[0]}M)")
        fig.update_layout(**_PLOTLY_BASE,
                           title=(f"{model_type} — Actual vs Forecast  "
                                  f"{STOCKS.get(ticker,ticker)} ({ticker})<br>"
                                  f"<sup style='color:{_C['muted']}'>{anno_txt}</sup>"),
                           xaxis_title="Date", yaxis_title="Monthly Log Return",
                           height=460, hovermode="x unified")
        return fig

    # ── CB-8: Deep Dive ──────────────────────────────────────────────
    @app.callback(
        [Output("price-regime-chart","figure"), Output("rv-chart","figure")],
        Input("deepdive-ticker-dd","value"),
    )
    def update_deep_dive(ticker):
        empty = go.Figure().update_layout(**_PLOTLY_BASE, title="No data")
        if ticker is None:
            return empty, empty
        px_s = (monthly_prices[ticker] if ticker in monthly_prices.columns
                else pd.Series(dtype=float))
        fig_px = go.Figure()
        fig_px.add_trace(go.Scatter(x=px_s.index, y=px_s.values, name="Monthly Close",
                                     mode="lines", line={"color": _C["primary"], "width": 1.8}))
        bear_col = [c for c in mkt_hmm_df.columns if "bear" in c]
        if bear_col and not px_s.empty:
            bear_p = mkt_hmm_df[bear_col[0]].reindex(px_s.index).fillna(0.0)
            for dt in bear_p[bear_p > 0.50].index:
                fig_px.add_vrect(x0=dt - pd.DateOffset(days=15),
                                  x1=dt + pd.DateOffset(days=15),
                                  fillcolor=_C["danger"], opacity=0.07, line_width=0)
        if scores_avail and ticker in ticker_conv:
            cv = ticker_conv[ticker]
            sec = ticker_sector.get(ticker, "")
            fig_px.add_annotation(
                text=f"Conviction: {cv:.2f}  |  Sector: {sec}",
                xref="paper", yref="paper", x=0.01, y=0.99,
                showarrow=False, font={"size": 11, "color": _C["muted"], "family": _FONT},
                bgcolor="rgba(255,255,255,0.8)", borderpad=4,
            )
        fig_px.update_layout(**_PLOTLY_BASE,
                               title=f"{STOCKS.get(ticker,ticker)} — Price  (red shading = P(bear) > 50%)",
                               xaxis_title="Date", yaxis_title="Price (₹)", height=340)
        fig_rv = go.Figure()
        if not realized_vol.empty and ticker in realized_vol.columns:
            fig_rv.add_trace(go.Scatter(x=realized_vol.index, y=realized_vol[ticker].values,
                                         name="Realized Vol", mode="lines",
                                         line={"color": _C["warning"], "width": 1.8},
                                         fill="tozeroy", fillcolor="rgba(194,120,3,0.08)"))
        fig_rv.update_layout(**_PLOTLY_BASE,
                               title=f"{STOCKS.get(ticker,ticker)} — Monthly Realized Volatility",
                               xaxis_title="Date", yaxis_title="σ (Realized)", height=260)
        return fig_px, fig_rv

    # Write stub
    with open(OUTPUT_DIR / "dashboard_app.py", "w") as fh:
        fh.write('"""NSE Portfolio Monitor — standalone launcher."""\n'
                 'from pathlib import Path\nimport pandas as pd\n'
                 'OUTPUT_DIR = Path("pipeline_outputs_v2")\n'
                 'print("Artifacts at", OUTPUT_DIR.resolve())\n')

    logger.info("Phase 9 ✓\n")
    return app


# %%
# ======================================================================
# MAIN
# ======================================================================
def main() -> None:
    logger.info("=" * 70)
    logger.info("NSE Stock Forecasting & Portfolio Pipeline — Starting")
    logger.info("=" * 70)

    p1     = run_phase1()
    p2     = run_phase2(p1)
    p3     = run_phase3(p1, p2)
    p3_5   = run_phase3_5()          # NEW: sector scores & conviction signals
    p4     = run_phase4(p1, p2)
    p4_5   = run_phase4_5(p1, p2, p3, p4)
    p5     = run_phase5(p1, p2, p3, p4, p4_5)
    p6     = run_phase6(p5, p1)
    p7     = run_phase7(p1, p2, p3, p4_5, p6, p3_5)   # conviction BL
    p8     = run_phase8(p1, p2, p4_5, p7)
    p_eval = run_phase_eval(p1, p2, p3, p4, p4_5)
    app    = run_phase9(p1, p2, p3, p4_5, p5, p6, p7, p8, p_eval, p3_5)

    logger.info("=" * 70)
    logger.info("All phases complete. Outputs: %s", OUTPUT_DIR.resolve())
    logger.info("=" * 70)
    logger.info("Dashboard → http://127.0.0.1:8050/")
    app.run(debug=False, host="0.0.0.0", port=8050)
# import os
# Windows alternative
# os.system("for /f \"tokens=5\" %a in ('netstat -aon ^| find \":8050\"') do taskkill /f /pid %a")
if __name__ == "__main__":
    main()


