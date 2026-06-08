# FP&A Commercial Variance Engine

### A whitepaper-grade rolling forecast & EBITDA bridge for a UK multi-site leisure operator

**Author** &nbsp; Amogh H.H. — MSc Business Analytics (Essex) · MBA Finance & Marketing
**Sector** &nbsp; UK Retail / Leisure (multi-site)
**Stack** &nbsp; Excel (FAST) · Python (openpyxl, statsmodels) · R (forecast, tidyverse) · Git
**Latest refresh** &nbsp; 26 May 2026

---

## 1. Executive summary

Most month-end packs answer *what happened*. They rarely answer *why* or *what next*. This repository ships an end-to-end FP&A engine that closes that gap for **Nova Leisure Group**, a synthetic UK multi-site operator of leisure venues. The engine ingests a deliberately messy general-ledger export, maps it against a FAST-standard budget, isolates the price-vs-volume drivers of the variance, and overlays a probabilistic site-level demand forecast trained on 17 months of daily footfall.

The result is a single workbook that lets a finance partner walk into a CFO meeting and answer three questions on one screen:

1. **Where is the EBITDA leak?** &nbsp; A waterfall bridge decomposes the £393k FY25 EBITDA miss into revenue, COGS, wages and OPEX components.
2. **Is it price or volume?** &nbsp; A revenue bridge isolates a small adverse volume effect (footfall) from a stable price effect (avg spend per head).
3. **What's the forward setup?** &nbsp; A scenario-driven 12-month rolling forecast (Base / Bull / Bear) sensitised to the current UK macro pack (CPI 2.8%, footfall -10.7% YoY).

A predictive demand layer trained in `Predictive_Demand_Model/` adds a 90-day site-level visitor forecast with 80% / 95% prediction intervals. ARIMA edges Holt-Winters on group-level accuracy (MAPE **10.95%** vs **11.04%**), both materially beating the seasonal-naive benchmark (13.58%).

---

## 2. Why this matters commercially

A 1% improvement in forecast accuracy on a £18m revenue base equates to ~£180k of working-capital optimisation per annum (inventory, rostering, marketing spend). In multi-site leisure, where the cost stack is roughly 50% labour and 22% COGS, the asymmetry is more pronounced: an over-forecast burns wage hours that cannot be recovered, while an under-forecast destroys covers and reputation. This engine prices that asymmetry and lets the CFO trade off coverage against cost.

---

## 3. Repository layout

```
FPA_Commercial_Engine_UK/
├── README.md                          ← this whitepaper
├── .gitignore
├── Macro_Assumptions.txt              ← live UK CPI + BRC footfall pack (Apr-26)
├── FY25_Approved_Budget.csv           ← monthly P&L budget, 20 lines × 12 months
├── YTD_Actuals_Ledger.csv             ← 6,860-row transactional ledger (intentionally dirty)
├── Daily_Site_Footfall.csv            ← 2,550 daily site-level operational rows
├── Nova_FPA_Master_Model.xlsx         ← 8-tab FP&A engine, 724 live formulas, 0 errors
│
├── Predictive_Demand_Model/
│   ├── demand_forecast.R              ← R version (forecast::auto.arima + HoltWinters)
│   ├── demand_forecast.py             ← Python equivalent (statsmodels)
│   ├── forecast_output.csv            ← 450-row site × day forecast (90d horizon)
│   ├── model_metrics.txt              ← per-site RMSE / MAE / MAPE + Ljung-Box
│   └── forecast_plot.pdf              ← 5 site charts w/ 80% & 95% PI bands
│
└── _generators/                       ← reproducibility scripts
    ├── generate_datasets.py
    └── build_fpa_model.py
```

---

## 4. Methodology — variance analysis

### 4.1 Data scrubbing

The ledger replicates a real Xero / POS export with three categories of dirt that always appear in practice:

| Defect                              | Volume      | Treatment                                                       |
|-------------------------------------|-------------|-----------------------------------------------------------------|
| Miscoded OPEX (non-existent codes)  | ~3% of rows | Caught by `Actuals_Clean!UNMAPPED` line; flagged red on dashboard |
| Missing departmental tags           | ~6% of rows | Tolerated for P&L roll-up; surfaced as data-quality KPI         |
| Whitespace / case noise / dupes     | random      | Account-code level SUMIFS makes these idempotent                |
| Fully blank rows                    | 30 rows     | Stripped by SUMIFS' numeric criterion                           |

The mapping table is the contract between the ledger and the budget. Any account code that doesn't exist in `Mapping_Table` falls into the UNMAPPED bucket — a deliberate fail-loud design choice, because silent miscodes are how mid-cap finance teams lose tens of thousands per quarter to the wrong department.

### 4.2 EBITDA bridge

The waterfall on `BvA_Dashboard` decomposes the £393k EBITDA miss into five buckets — Revenue, COGS, Wages, Variable OPEX, Fixed OPEX — plus an Unmapped catcher. Each bucket is a `SUMIFS()` differential between actual and budget at the FY25 total level, so the bridge ties to the pence.

### 4.3 Price vs Volume bridge

Revenue variance is decomposed via the standard finance partition:

```
Volume Effect = (Actual Visitors  − Budget Visitors) × Budget Price
Price Effect  = (Actual Price     − Budget Price)    × Actual Visitors
Total Var     = Volume Effect + Price Effect
```

This isolates whether the revenue miss is footfall-driven (a marketing / catchment problem) or basket-driven (a pricing / merchandising problem). For Nova FY25 the bridge points to a roughly even split — actionable signal for both the Commercial Director and the Operations Director.

---

## 5. Methodology — predictive demand

### 5.1 Models evaluated

| Model               | Order / spec                                  | Rationale                                                    |
|---------------------|-----------------------------------------------|--------------------------------------------------------------|
| **Holt-Winters**    | Additive, weekly seasonality (period = 7)     | Robust against shocks, weekly seasonality is dominant signal |
| **ARIMA**           | (1,1,1)(1,1,1,7) — auto-selected in R version | Captures residual autocorrelation                            |
| **Seasonal Naive**  | y_t = y_{t-7}                                 | Benchmark; any production model must beat this               |

### 5.2 Validation

A 28-day out-of-sample holdout was used to score each model per site, then the lowest-MAPE model was refit on full history for the 90-day forward forecast. Residuals were checked for whiteness via the Ljung-Box test at lag 14.

### 5.3 Results (group-level, ranked by MAPE)

| Model           | RMSE   | MAE    | MAPE   |
|-----------------|--------|--------|--------|
| **ARIMA**       | 455.15 | 391.72 | **10.95%** |
| Holt-Winters    | 454.53 | 391.08 | 11.04% |
| Seasonal Naive  | 561.82 | 481.19 | 13.58% |

ARIMA wins on three of five sites; Holt-Winters wins on London and Edinburgh, both of which have stronger weekly cycles.

### 5.4 Honest model diagnostics

Ljung-Box p-values are all < 0.05, meaning the residuals retain some structure (likely weather and event effects not yet in the feature set). Logical next steps would be (i) an exogenous regression layer on weather and bank holidays via `xreg = ...` in `auto.arima`, or (ii) a Prophet model with explicit holiday regressors. The current models are deployed as the production benchmark, not the ceiling.

---

## 6. How the predictive layer feeds back into the P&L

`Predictive_Demand_Model/forecast_output.csv` produces a site × day visitor forecast for the next 90 days. Multiplied by the trailing 90-day avg spend per head (from `Daily_Site_Footfall.csv`) it produces a bottom-up revenue estimate that can replace or stress-test the top-down forecast on the `Rolling_Forecast` tab. The Base case in the macro pack is roughly consistent with the ARIMA central forecast; the Bear case (-6% footfall) approximates the lower bound of the 95% prediction interval over a 90-day window.

---

## 7. Macro pack

| Driver                      | Print          | Source                                  |
|-----------------------------|----------------|-----------------------------------------|
| Headline UK CPI (Apr-26)    | 2.8% YoY       | ONS Consumer Price Inflation Bulletin   |
| Core CPI                    | 2.5% YoY       | ONS                                     |
| UK retail footfall (Apr-26) | -10.7% YoY     | BRC-Sensormatic Footfall Monitor        |
| National Living Wage (21+)  | £12.21 / hr    | gov.uk (effective 1 Apr-26)             |
| BoE Bank Rate consensus     | 3.75% by Q4-26 | BoE Monetary Policy Report, May-26      |

The Assumptions tab is fully data-validation toggled — switching between Bear / Base / Bull recalculates the rolling forecast in place.

```

## 8. Skills demonstrated

| Skill                                | Evidence in this repo                                    |
|--------------------------------------|----------------------------------------------------------|
| FAST-standard financial modelling    | `Nova_FPA_Master_Model.xlsx` — colour coding, audit trail|
| Variance analysis & EBITDA bridging  | `BvA_Dashboard` waterfall + price/volume split           |
| Scenario & sensitivity analysis      | `Assumptions` tab, INDEX/MATCH-driven scenario toggle    |
| Data engineering on dirty ledgers    | SUMIFS + INDEX/MATCH mapping, UNMAPPED fail-loud catcher |
| Time-series forecasting (ARIMA / HW) | R + Python implementations, holdout validation, PIs      |
| Statistical diagnostics              | Ljung-Box residual whiteness test                        |
| Macro literacy (UK)                  | Live ONS / BRC / BoE pack                                |
| Reproducibility                      | Seed-fixed generators, single-command rebuild            |
| Version control                      | Git, `.gitignore`, semantic commit                       |

---

## 9. Disclaimer

Nova Leisure Group is a synthetic entity. All financial figures are illustrative. Macro figures are sourced from the references listed above as of the refresh date.
