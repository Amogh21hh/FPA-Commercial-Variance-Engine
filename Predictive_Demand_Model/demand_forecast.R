# =================================================================
# Nova Leisure Group — Predictive Demand Model (R)
# Author: Amogh H.H. — Commercial Finance / Business Analytics
#
# Purpose
#   Ingest Daily_Site_Footfall.csv and produce a 90-day site-level
#   demand forecast that feeds back into the Excel revenue line.
#
# Models
#   1. Holt-Winters triple exponential smoothing (additive, weekly
#      seasonality of 7).
#   2. ARIMA via auto.arima — model is auto-selected per site
#      based on AICc.
#   3. Naive Seasonal benchmark (sNaive) — used as accuracy baseline.
#
# Outputs
#   - forecast_output.csv   : site × date × point + 80/95% PIs
#   - model_metrics.txt     : per-model RMSE, MAE, MAPE on holdout
#   - residual_diagnostics  : Ljung-Box p-values (printed)
# =================================================================

suppressPackageStartupMessages({
  pkgs <- c("forecast", "dplyr", "lubridate", "readr", "tidyr")
  to_install <- setdiff(pkgs, rownames(installed.packages()))
  if (length(to_install)) install.packages(to_install, repos = "https://cloud.r-project.org")
  lapply(pkgs, library, character.only = TRUE)
})

# --- Config -------------------------------------------------------
HORIZON_DAYS  <- 90
HOLDOUT_DAYS  <- 28
SEED          <- 42
set.seed(SEED)

ROOT     <- dirname(dirname(normalizePath(sys.frame(1)$ofile, mustWork = FALSE)))
if (!nzchar(ROOT) || ROOT == "/") {
  ROOT <- "."
}
INPUT    <- file.path(ROOT, "Daily_Site_Footfall.csv")
OUTDIR   <- file.path(ROOT, "Predictive_Demand_Model")
if (!dir.exists(OUTDIR)) dir.create(OUTDIR, recursive = TRUE)

OUT_CSV  <- file.path(OUTDIR, "forecast_output.csv")
OUT_METR <- file.path(OUTDIR, "model_metrics.txt")
OUT_PLOT <- file.path(OUTDIR, "forecast_plot.pdf")

# --- Load ---------------------------------------------------------
df <- read_csv(INPUT, show_col_types = FALSE) %>%
  mutate(Date = as.Date(Date))

cat(sprintf("Loaded %d daily rows across %d sites.\n",
            nrow(df), length(unique(df$Site_Code))))

sites <- unique(df$Site_Code)

# --- Helpers ------------------------------------------------------
accuracy_metrics <- function(actual, pred) {
  err  <- actual - pred
  rmse <- sqrt(mean(err^2, na.rm = TRUE))
  mae  <- mean(abs(err), na.rm = TRUE)
  mape <- mean(abs(err / actual) * 100, na.rm = TRUE)
  c(RMSE = rmse, MAE = mae, MAPE = mape)
}

all_forecasts <- list()
metrics_rows  <- list()
diag_rows     <- list()

for (s in sites) {
  site_df <- df %>%
    filter(Site_Code == s) %>%
    arrange(Date)

  site_name <- unique(site_df$Site_Name)[1]
  region    <- unique(site_df$Region)[1]

  y <- ts(site_df$Visitors, frequency = 7)

  # Train / holdout
  n <- length(y)
  train <- window(y, end = c(floor((n - HOLDOUT_DAYS) / 7), ((n - HOLDOUT_DAYS) %% 7) + 1))
  # Simpler: by index
  train <- ts(head(site_df$Visitors, n - HOLDOUT_DAYS), frequency = 7)
  test  <- tail(site_df$Visitors, HOLDOUT_DAYS)

  # --- Model 1: Holt-Winters (additive, with damped trend off)
  hw_fit  <- HoltWinters(train, seasonal = "additive")
  hw_pred <- forecast(hw_fit, h = HOLDOUT_DAYS)$mean

  # --- Model 2: ARIMA (auto)
  ar_fit  <- auto.arima(train, seasonal = TRUE, stepwise = TRUE, approximation = FALSE)
  ar_pred <- forecast(ar_fit, h = HOLDOUT_DAYS)$mean

  # --- Model 3: Seasonal Naive benchmark
  sn_pred <- snaive(train, h = HOLDOUT_DAYS)$mean

  # --- Accuracy
  hw_m <- accuracy_metrics(test, hw_pred)
  ar_m <- accuracy_metrics(test, ar_pred)
  sn_m <- accuracy_metrics(test, sn_pred)

  metrics_rows[[length(metrics_rows) + 1]] <- data.frame(
    Site_Code = s, Site_Name = site_name,
    Model = c("HoltWinters", "ARIMA", "SeasonalNaive"),
    RMSE  = c(hw_m["RMSE"], ar_m["RMSE"], sn_m["RMSE"]),
    MAE   = c(hw_m["MAE"],  ar_m["MAE"],  sn_m["MAE"]),
    MAPE  = c(hw_m["MAPE"], ar_m["MAPE"], sn_m["MAPE"])
  )

  # --- Pick best model by MAPE
  best_idx <- which.min(c(hw_m["MAPE"], ar_m["MAPE"], sn_m["MAPE"]))
  best     <- c("HoltWinters", "ARIMA", "SeasonalNaive")[best_idx]

  # --- Ljung-Box on residuals
  resid <- switch(best,
                  HoltWinters  = residuals(hw_fit),
                  ARIMA        = residuals(ar_fit),
                  SeasonalNaive= rep(NA, length(train)))
  lb_p <- tryCatch(Box.test(resid, lag = 14, type = "Ljung-Box")$p.value,
                   error = function(e) NA)

  diag_rows[[length(diag_rows) + 1]] <- data.frame(
    Site_Code = s, Best_Model = best, LjungBox_p = lb_p
  )

  # --- Refit on FULL history and forecast forward 90 days
  y_full <- ts(site_df$Visitors, frequency = 7)
  final_fit <- switch(best,
                      HoltWinters = HoltWinters(y_full, seasonal = "additive"),
                      ARIMA       = auto.arima(y_full, seasonal = TRUE),
                      SeasonalNaive = NULL)
  if (best == "SeasonalNaive") {
    fc <- snaive(y_full, h = HORIZON_DAYS)
  } else {
    fc <- forecast(final_fit, h = HORIZON_DAYS, level = c(80, 95))
  }

  last_date <- max(site_df$Date)
  out_dates <- seq(last_date + 1, by = "day", length.out = HORIZON_DAYS)

  all_forecasts[[length(all_forecasts) + 1]] <- data.frame(
    Date          = out_dates,
    Site_Code     = s,
    Site_Name     = site_name,
    Region        = region,
    Model         = best,
    Forecast      = round(as.numeric(fc$mean)),
    Lower_80      = round(as.numeric(if (!is.null(fc$lower)) fc$lower[, 1] else fc$mean)),
    Upper_80      = round(as.numeric(if (!is.null(fc$upper)) fc$upper[, 1] else fc$mean)),
    Lower_95      = round(as.numeric(if (!is.null(fc$lower)) fc$lower[, 2] else fc$mean)),
    Upper_95      = round(as.numeric(if (!is.null(fc$upper)) fc$upper[, 2] else fc$mean))
  )

  cat(sprintf("  %s (%s): best=%s  MAPE=%.2f%%  LjungBox p=%.3f\n",
              s, site_name, best,
              min(c(hw_m["MAPE"], ar_m["MAPE"], sn_m["MAPE"])),
              lb_p))
}

# --- Write outputs -----------------------------------------------
fc_out <- bind_rows(all_forecasts)
metrics <- bind_rows(metrics_rows)
diag    <- bind_rows(diag_rows)

write_csv(fc_out, OUT_CSV)

sink(OUT_METR)
cat("Nova Leisure Group — Demand Model Accuracy (28-day holdout)\n")
cat(sprintf("Generated: %s\n", Sys.time()))
cat(paste(rep("=", 70), collapse = ""), "\n\n")
print(metrics, row.names = FALSE)
cat("\n\nBest model per site & residual whiteness (Ljung-Box p > 0.05 = white noise)\n")
cat(paste(rep("-", 70), collapse = ""), "\n")
print(diag, row.names = FALSE)
cat("\n\nGroup-level accuracy (visitor-weighted)\n")
cat(paste(rep("-", 70), collapse = ""), "\n")
grp <- metrics %>%
  group_by(Model) %>%
  summarise(Avg_RMSE = mean(RMSE), Avg_MAE = mean(MAE), Avg_MAPE = mean(MAPE)) %>%
  arrange(Avg_MAPE)
print(grp, row.names = FALSE)
sink()

# --- Plot --------------------------------------------------------
pdf(OUT_PLOT, width = 11, height = 7)
for (s in sites) {
  hist <- df %>% filter(Site_Code == s) %>% arrange(Date)
  fcs  <- fc_out %>% filter(Site_Code == s)
  plot(hist$Date, hist$Visitors, type = "l", col = "grey40",
       xlim = range(c(hist$Date, fcs$Date)),
       ylim = range(c(hist$Visitors, fcs$Lower_95, fcs$Upper_95), na.rm = TRUE),
       xlab = "Date", ylab = "Visitors",
       main = paste0(unique(hist$Site_Name), " — ", unique(fcs$Model), " forecast (90d)"))
  polygon(c(fcs$Date, rev(fcs$Date)),
          c(fcs$Lower_95, rev(fcs$Upper_95)),
          col = rgb(0.2, 0.4, 0.8, 0.15), border = NA)
  polygon(c(fcs$Date, rev(fcs$Date)),
          c(fcs$Lower_80, rev(fcs$Upper_80)),
          col = rgb(0.2, 0.4, 0.8, 0.25), border = NA)
  lines(fcs$Date, fcs$Forecast, col = "navy", lwd = 2)
  abline(v = max(hist$Date), col = "red", lty = 2)
  legend("topleft",
         legend = c("Historical", "Forecast", "80% PI", "95% PI"),
         col = c("grey40", "navy", rgb(0.2, 0.4, 0.8, 0.4), rgb(0.2, 0.4, 0.8, 0.2)),
         lwd = c(1, 2, 8, 8), bty = "n", cex = 0.8)
}
dev.off()

cat("\n--- WRITE COMPLETE ---\n")
cat(sprintf("Forecast CSV : %s  (%d rows)\n", OUT_CSV, nrow(fc_out)))
cat(sprintf("Metrics file : %s\n", OUT_METR))
cat(sprintf("Plot PDF     : %s\n", OUT_PLOT))
