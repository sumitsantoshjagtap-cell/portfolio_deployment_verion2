"""
Render‑ready Dash server for the NSE Portfolio Pipeline.
Requires that the full pipeline has been run locally first (produces pipeline_outputs_v2/).
"""

import os
import pandas as pd
import numpy as np
from pathlib import Path

# ----------------------------------------------------------------------
#  Import the dashboard builder and constants from the main pipeline script
#  (rename your original script to pipeline.py before pushing to Render)
# ----------------------------------------------------------------------
from pipeline import (
    run_phase9,
    STOCKS,
    FORECAST_HORIZONS,
    HORIZON_COLORS,
    MIN_TRAIN_MONTHS,
    SECTOR_GROUPS,
    SECTOR_SCORE_WEIGHTS,
    SCORE_LAG_MONTHS,
    SECTOR_SCORE_FILE,
    _C,
    _FONT,
    _PLOTLY_BASE,
    # … other constants that run_phase9 may need (it uses _C, _FONT etc)
)

OUTPUT_DIR = Path("pipeline_outputs_v2")
ANNUAL_RISK_FREE = 0.065

# ----------------------------------------------------------------------
#  Utility: compute sector conviction from composite scores
# ----------------------------------------------------------------------
def compute_sector_conviction(composite_df: pd.DataFrame) -> dict:
    """Replicate the conviction logic from Phase 3.5."""
    if composite_df.empty:
        return {}
    latest = composite_df.iloc[-1]
    prev = composite_df.iloc[-2] if len(composite_df) >= 2 else latest
    level_comp = (latest / 10.0).clip(0.0, 1.0)
    momentum_raw = (latest - prev) / 10.0
    momentum_comp = (momentum_raw / 2.0 + 0.5).clip(0.0, 1.0)
    # Weights from the original config (CONVICTION_LEVEL_WT, CONVICTION_MOMENTUM_WT)
    CONVICTION_LEVEL_WT = 0.6
    CONVICTION_MOMENTUM_WT = 0.4
    conviction = (
        CONVICTION_LEVEL_WT * level_comp + CONVICTION_MOMENTUM_WT * momentum_comp
    ).clip(0.15, 0.95)
    return conviction.to_dict()


# ----------------------------------------------------------------------
#  Load all pre‑computed outputs
# ----------------------------------------------------------------------
monthly_prices = pd.read_parquet(OUTPUT_DIR / "monthly_prices.parquet")
mlr_clean = pd.read_parquet(OUTPUT_DIR / "monthly_log_returns_clean.parquet")
bench_returns = pd.read_parquet(OUTPUT_DIR / "benchmark_returns.parquet").squeeze()
active_tickers = [t for t in STOCKS if t in monthly_prices.columns]
monthly_rf = (1 + ANNUAL_RISK_FREE) ** (1 / 12) - 1

realized_vol = pd.read_parquet(OUTPUT_DIR / "realized_volatility.parquet")
market_indicators = pd.read_parquet(OUTPUT_DIR / "market_indicators.parquet")

# HMM filtered probabilities – market level
all_probs = pd.read_parquet(OUTPUT_DIR / "filtered_probs_all.parquet")
market_filtered_df = all_probs.filter(regex="^market_hmm_")
# (group & stock HMM probs not used in dashboard, provide empty dicts)
group_filtered_probs = {}
stock_filtered_probs = {}

# Walk‑forward forecasts & performance
forecasts_df = pd.read_parquet(OUTPUT_DIR / "walkforward_forecasts.parquet")
errors_df = pd.read_parquet(OUTPUT_DIR / "walkforward_errors.parquet")
perf_df = pd.read_csv(OUTPUT_DIR / "performance_metrics.csv")
calib_df = pd.read_csv(OUTPUT_DIR / "calibration_results.csv")

# Combined forecasts
combined_df = pd.read_parquet(OUTPUT_DIR / "combined_forecasts.parquet")

# Portfolio weights
weights_df = pd.read_csv(OUTPUT_DIR / "portfolio_weights.csv")
weights_clean = dict(zip(weights_df["ticker"], weights_df["weight"]))

# Backtest performance
portfolio_rets = pd.read_parquet(OUTPUT_DIR / "portfolio_returns.parquet").squeeze()
portfolio_val = pd.read_parquet(OUTPUT_DIR / "portfolio_value.parquet").squeeze()
drawdown_s = pd.read_parquet(OUTPUT_DIR / "drawdown_series.parquet").squeeze()
perf_summary_df = pd.read_csv(OUTPUT_DIR / "performance_summary.csv")
perf_summary = perf_summary_df.iloc[0].to_dict() if not perf_summary_df.empty else {}

# Model evaluation
eval_df = pd.read_csv(OUTPUT_DIR / "model_evaluation.csv")
eval_preds_df = pd.read_csv(OUTPUT_DIR / "model_eval_predictions.csv")

# Load Black‑Litterman posterior (for custom weights explorer)
bl_returns_path = OUTPUT_DIR / "bl_posterior_returns.pkl"
bl_cov_path = OUTPUT_DIR / "bl_posterior_cov.pkl"
ret_bl = None
S_bl   = None
if bl_returns_path.exists() and bl_cov_path.exists():
    ret_bl = pd.read_pickle(bl_returns_path)
    S_bl   = pd.read_pickle(bl_cov_path)

# Sector scores (optional)
sector_composite_scores = pd.DataFrame()
sector_conviction = {}
ticker_conviction = {}
ticker_sector = {t: "Unknown" for t in STOCKS}
scores_available = False
sector_file = OUTPUT_DIR / "sector_composite_scores.csv"
if sector_file.exists():
    sector_composite_scores = pd.read_csv(sector_file, index_col=0, parse_dates=True)
    sector_conviction = compute_sector_conviction(sector_composite_scores)
    # map to tickers
    for sector, tickers in SECTOR_GROUPS.items():
        cv = sector_conviction.get(sector, 0.5)
        for t in tickers:
            ticker_conviction[t] = cv
    # build ticker -> sector map
    ticker_sector = {}
    for sector, tickers in SECTOR_GROUPS.items():
        for t in tickers:
            ticker_sector[t] = sector
    scores_available = True

# ----------------------------------------------------------------------
#  Assemble the dictionaries that run_phase9 expects
# ----------------------------------------------------------------------
p1 = {
    "monthly_prices": monthly_prices,
    "benchmark_returns": bench_returns,
    "active_tickers": active_tickers,
    "monthly_rf": monthly_rf,
    # the following are not used by the dashboard but ensure no KeyError
    "bench_monthly_prices": pd.Series(dtype=float),
    "daily_close": pd.DataFrame(),
    "price_data": {},
    "indicator_monthly_prices": {},
    "WATCHLIST": {},
    "audit_df": pd.DataFrame(),
    "end_str": "",
}

p2 = {"monthly_log_returns_clean": mlr_clean}

p3 = {
    "realized_volatility": realized_vol,
    "market_indicators": market_indicators,
    "exog_master": pd.DataFrame(),  # not used
}

p4_5 = {
    "market_filtered_df": market_filtered_df,
    "group_filtered_probs": group_filtered_probs,
    "stock_filtered_probs": stock_filtered_probs,
    "all_probs_df": all_probs,
    "exog_master_updated": pd.DataFrame(),
    "hmm_market": None,
}

p5 = {
    "forecasts_df": forecasts_df,
    "errors_df": errors_df,
    "perf_df": perf_df,
    "calib_df": calib_df,
}

p6 = {"combined_df": combined_df}

# For Black‑Litterman ret_bl / S_bl – not stored, skip or compute on‑the‑fly.
# The dashboard’s custom weights explorer will still work without them.
p7 = {
    "weights_clean": weights_clean,
    "allocation": {},
    "leftover": 0.0,
    "S_full": pd.DataFrame(),            # not used
    "latest_prices": pd.Series(dtype=float),
    "stress_prob": 0.0,
    "in_stress": False,
    "ret_bl": ret_bl,
    "S_bl": S_bl,
    "ticker_conviction": ticker_conviction,
}

p8 = {
    "portfolio_returns": portfolio_rets,
    "portfolio_value": portfolio_val,
    "drawdown_series": drawdown_s,
    "max_drawdown": perf_summary.get("Maximum Drawdown", 0),
    "perf_summary": perf_summary,
    "regime_perf_df": pd.DataFrame(),
}

p_eval = {
    "eval_df": eval_df,
    "eval_predictions_df": eval_preds_df,
}

p3_5 = {
    "sector_composite_scores": sector_composite_scores,
    "sector_composite_monthly": pd.DataFrame(),  # not used
    "sector_conviction": sector_conviction,
    "ticker_conviction": ticker_conviction,
    "ticker_sector": ticker_sector,
    "scores_available": scores_available,
}

# ----------------------------------------------------------------------
#  Build the Dash application and expose the Flask server
# ----------------------------------------------------------------------
app = run_phase9(p1, p2, p3, p4_5, p5, p6, p7, p8, p_eval, p3_5)
server = app.server

if __name__ == "__main__":
    # For local testing: python app.py
    port = int(os.environ.get("PORT", 8050))
    app.run(host="0.0.0.0", port=port, debug=False)
