"""
Nova Leisure Group — Predictive Demand Model (Python equivalent)
Mirrors demand_forecast.R for users without R installed.

Usage:  python3 demand_forecast.py
Outputs in Predictive_Demand_Model/:
  - forecast_output.csv
  - model_metrics.txt
  - forecast_plot.pdf
"""
import os, sys, warnings
import pandas as pd
import numpy as np
from datetime import timedelta

warnings.filterwarnings("ignore")

ROOT     = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
INPUT    = os.path.join(ROOT, "Daily_Site_Footfall.csv")
OUTDIR   = os.path.join(ROOT, "Predictive_Demand_Model")
os.makedirs(OUTDIR, exist_ok=True)
OUT_CSV  = os.path.join(OUTDIR, "forecast_output.csv")
OUT_METR = os.path.join(OUTDIR, "model_metrics.txt")
OUT_PLOT = os.path.join(OUTDIR, "forecast_plot.pdf")

HORIZON  = 90
HOLDOUT  = 28
np.random.seed(42)

from statsmodels.tsa.holtwinters import ExponentialSmoothing
from statsmodels.tsa.arima.model import ARIMA
from statsmodels.stats.diagnostic import acorr_ljungbox
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages

df = pd.read_csv(INPUT, parse_dates=["Date"]).sort_values(["Site_Code","Date"])
print(f"Loaded {len(df):,} daily rows across {df['Site_Code'].nunique()} sites.")

def metrics(y, yhat):
    err = np.array(y) - np.array(yhat)
    rmse = float(np.sqrt(np.mean(err**2)))
    mae  = float(np.mean(np.abs(err)))
    mape = float(np.mean(np.abs(err / np.where(np.array(y)==0, np.nan, np.array(y)))) * 100)
    return {"RMSE": rmse, "MAE": mae, "MAPE": mape}

def fit_arima(train):
    """Simple seasonal ARIMA(1,1,1)(1,1,1,7) — robust default that auto-arima typically selects."""
    try:
        m = ARIMA(train, order=(1,1,1), seasonal_order=(1,1,1,7),
                  enforce_stationarity=False, enforce_invertibility=False).fit()
        return m
    except Exception:
        m = ARIMA(train, order=(1,1,1)).fit()
        return m

def fit_hw(train):
    return ExponentialSmoothing(train, trend="add", seasonal="add", seasonal_periods=7,
                                initialization_method="estimated").fit(optimized=True)

def snaive(train, h, period=7):
    last_period = list(train[-period:])
    return [last_period[i % period] for i in range(h)]

all_fc = []
metrics_rows = []
diag_rows = []
plots = []

sites = df[["Site_Code","Site_Name","Region"]].drop_duplicates().to_records(index=False)

for site_code, site_name, region in sites:
    sdf = df[df["Site_Code"] == site_code].copy().reset_index(drop=True)
    y = sdf["Visitors"].astype(float).values
    n = len(y)
    train, test = y[:-HOLDOUT], y[-HOLDOUT:]

    # Holt-Winters
    try:
        hw_fit  = fit_hw(train)
        hw_pred = hw_fit.forecast(HOLDOUT)
    except Exception as e:
        hw_pred = np.repeat(np.mean(train), HOLDOUT)
        hw_fit  = None

    # ARIMA
    try:
        ar_fit  = fit_arima(train)
        ar_pred = ar_fit.forecast(HOLDOUT)
    except Exception:
        ar_pred = np.repeat(np.mean(train), HOLDOUT)
        ar_fit  = None

    # Seasonal Naive
    sn_pred = snaive(train, HOLDOUT)

    hw_m = metrics(test, hw_pred)
    ar_m = metrics(test, ar_pred)
    sn_m = metrics(test, sn_pred)

    for model, m in [("HoltWinters", hw_m), ("ARIMA", ar_m), ("SeasonalNaive", sn_m)]:
        metrics_rows.append({"Site_Code":site_code, "Site_Name":site_name,
                             "Model":model, **m})

    cands = {"HoltWinters": hw_m["MAPE"], "ARIMA": ar_m["MAPE"], "SeasonalNaive": sn_m["MAPE"]}
    best  = min(cands, key=cands.get)
    best_mape = cands[best]

    # Ljung-Box
    lb_p = None
    if best == "HoltWinters" and hw_fit is not None:
        resid = hw_fit.resid
    elif best == "ARIMA" and ar_fit is not None:
        resid = ar_fit.resid
    else:
        resid = None
    if resid is not None:
        try:
            lb = acorr_ljungbox(resid.dropna() if hasattr(resid,'dropna') else resid, lags=[14], return_df=True)
            lb_p = float(lb["lb_pvalue"].iloc[0])
        except Exception:
            lb_p = None

    diag_rows.append({"Site_Code":site_code, "Best_Model":best,
                      "LjungBox_p_lag14": lb_p})

    # Refit on full data → forecast HORIZON ahead with intervals (HW or ARIMA)
    if best == "ARIMA":
        m = fit_arima(y)
        fc_res = m.get_forecast(HORIZON)
        mean   = np.asarray(fc_res.predicted_mean)
        ci80   = np.asarray(fc_res.conf_int(alpha=0.20))
        ci95   = np.asarray(fc_res.conf_int(alpha=0.05))
        lo80, hi80 = ci80[:,0], ci80[:,1]
        lo95, hi95 = ci95[:,0], ci95[:,1]
    elif best == "HoltWinters":
        m = fit_hw(y)
        mean = np.array(m.forecast(HORIZON))
        # HW doesn't ship native PIs in statsmodels — bootstrap residuals
        res = m.resid.dropna().values if hasattr(m.resid,'dropna') else np.asarray(m.resid)
        sd  = np.std(res)
        lo80 = mean - 1.282 * sd; hi80 = mean + 1.282 * sd
        lo95 = mean - 1.960 * sd; hi95 = mean + 1.960 * sd
    else:
        mean = np.array(snaive(y, HORIZON))
        res  = np.array(y[-HOLDOUT:]) - np.array(snaive(y[:-HOLDOUT], HOLDOUT))
        sd   = np.std(res)
        lo80 = mean - 1.282 * sd; hi80 = mean + 1.282 * sd
        lo95 = mean - 1.960 * sd; hi95 = mean + 1.960 * sd

    last_date = sdf["Date"].iloc[-1]
    fc_dates  = [last_date + timedelta(days=i+1) for i in range(HORIZON)]

    for i in range(HORIZON):
        all_fc.append({
            "Date": fc_dates[i].date().isoformat(),
            "Site_Code": site_code,
            "Site_Name": site_name,
            "Region":    region,
            "Model":     best,
            "Forecast":  int(round(max(0, mean[i]))),
            "Lower_80":  int(round(max(0, lo80[i]))),
            "Upper_80":  int(round(max(0, hi80[i]))),
            "Lower_95":  int(round(max(0, lo95[i]))),
            "Upper_95":  int(round(max(0, hi95[i]))),
        })

    plots.append((site_code, site_name, sdf["Date"].values, y, fc_dates, mean, lo95, hi95, best))
    print(f"  {site_code} ({site_name}): best={best:13s}  MAPE={best_mape:5.2f}%  "
          f"LjungBox p={lb_p if lb_p is not None else float('nan'):.3f}")

# Save CSV
fc_df = pd.DataFrame(all_fc)
fc_df.to_csv(OUT_CSV, index=False)

# Save metrics text
mdf = pd.DataFrame(metrics_rows)
ddf = pd.DataFrame(diag_rows)
grp = mdf.groupby("Model")[["RMSE","MAE","MAPE"]].mean().sort_values("MAPE")

with open(OUT_METR, "w") as f:
    f.write("Nova Leisure Group — Demand Model Accuracy (28-day holdout)\n")
    f.write("=" * 70 + "\n\n")
    f.write("Per-site accuracy:\n")
    f.write(mdf.round(2).to_string(index=False))
    f.write("\n\nBest model per site & Ljung-Box residual whiteness (p > 0.05 = white noise):\n")
    f.write(ddf.round(3).to_string(index=False))
    f.write("\n\nGroup-level average accuracy (ranked):\n")
    f.write(grp.round(2).to_string())
    f.write("\n")

# Plot
with PdfPages(OUT_PLOT) as pdf:
    for site_code, site_name, hist_dates, hist_y, fc_dates, mean, lo95, hi95, best in plots:
        fig, ax = plt.subplots(figsize=(11, 6))
        ax.plot(hist_dates, hist_y, color="grey", linewidth=0.8, label="Historical")
        ax.fill_between(fc_dates, lo95, hi95, alpha=0.18, color="navy", label="95% PI")
        ax.plot(fc_dates, mean, color="navy", linewidth=2, label="Forecast")
        ax.axvline(hist_dates[-1], color="red", linestyle="--", linewidth=0.8)
        ax.set_title(f"{site_name} — {best} forecast (90d)")
        ax.set_xlabel("Date"); ax.set_ylabel("Visitors")
        ax.legend(loc="upper left", frameon=False)
        ax.grid(alpha=0.2)
        plt.tight_layout()
        pdf.savefig(fig); plt.close(fig)

print(f"\nWrote: {OUT_CSV}  ({len(fc_df):,} rows)")
print(f"Wrote: {OUT_METR}")
print(f"Wrote: {OUT_PLOT}")
