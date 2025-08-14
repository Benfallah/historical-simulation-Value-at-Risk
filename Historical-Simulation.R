###############################################################################
# Historical Simulation VaR (medium-level R)
# - Reads prices from: "Historical Simulation Calcs.xlsx"
# - Indices: DJIA, FTSE-500, Nikkei, CAC-40  (name-matching is flexible)
# - Outputs: VaR table (95%, 99%) for 1, 5, 10-day horizons + 1-day volatility
# - Plot: histogram of 1-day returns with VaR cutoffs
###############################################################################

# ---- Packages (install once if needed) ----
# install.packages(c("readxl", "dplyr", "tidyr", "tibble", "ggplot2"))

library(readxl)
library(dplyr)
library(tidyr)
library(tibble)
library(ggplot2)

# ---- Config ----
excel_path        <- "Historical Simulation Calcs.xlsx"  # keep next to this script
sheet_name        <- 1                                   
confidence_levels <- c(0.95, 0.99)
horizons          <- c(1, 5, 10)                         # 1d, 5d, 10d VaR
indices_prefer    <- c("DJIA", "FTSE", "Nikkei", "CAC")  # display order

# ---- Small helpers ----

# Pick a column by a pattern (case-insensitive). 
# If multiple match, take the first (and warn). If none, stop with a clear message.
pick_col <- function(df, pattern) {
  idx <- grep(pattern, names(df), ignore.case = TRUE)
  if (length(idx) == 0) stop(paste("Could not find a column matching:", pattern))
  if (length(idx) > 1) warning(paste("Multiple columns matched", pattern, "- using the first."))
  df[[idx[1]]]
}

# Compute overlapping k-day log returns from a vector of prices.
kday_log_returns <- function(price_vec, k = 1) {
  lp <- log(price_vec)
  n  <- length(lp)
  if (n <= k) return(numeric(0))
  lp[(k + 1):n] - lp[1:(n - k)]
}

# Historical Simulation VaR: left-tail quantile of returns (as log-return).
# Returns a single number (log-return, e.g., -0.0245 for -2.45%).
hs_var <- function(returns, cl = 0.95) {
  if (length(returns) == 0) return(NA_real_)
  unname(quantile(returns, probs = 1 - cl, na.rm = TRUE))
}

# ---- Load data ----
message("Reading Excel file: ", excel_path)
raw <- read_excel(excel_path, sheet = sheet_name)

# If you have a Date column, sort by date (uncomment if needed)
# if ("Date" %in% names(raw)) raw <- raw[order(raw$Date), ]

# Map likely column names; this is flexible to variations like "FTSE-500" or "CAC-40"
DJIA_px   <- pick_col(raw, "DJIA")
FTSE_px   <- pick_col(raw, "FTSE")        # will match FTSE-500
Nikkei_px <- pick_col(raw, "Nikkei")
CAC_px    <- pick_col(raw, "CAC")         # will match CAC-40

prices_tbl <- tibble(
  DJIA   = as.numeric(DJIA_px),
  FTSE   = as.numeric(FTSE_px),
  Nikkei = as.numeric(Nikkei_px),
  CAC    = as.numeric(CAC_px)
) %>% drop_na()

# Basic sanity check
stopifnot(nrow(prices_tbl) > max(horizons))

message("Loaded ", nrow(prices_tbl), " price rows after cleaning.")

# ---- 1-day log returns for each index (useful for volatility and plots) ----
one_day_lr <- prices_tbl %>%
  transmute(
    Date_row = row_number(),                           # row index if no dates
    DJIA     = c(NA, diff(log(DJIA))),
    FTSE     = c(NA, diff(log(FTSE))),
    Nikkei   = c(NA, diff(log(Nikkei))),
    CAC      = c(NA, diff(log(CAC)))
  ) %>% drop_na()

# ---- Build VaR table across indices, horizons, and confidence levels ----
indices    <- c("DJIA", "FTSE", "Nikkei", "CAC")
price_list <- list(
  DJIA   = prices_tbl$DJIA,
  FTSE   = prices_tbl$FTSE,
  Nikkei = prices_tbl$Nikkei,
  CAC    = prices_tbl$CAC
)

results <- list()
for (idx in indices) {
  for (h in horizons) {
    # Use true historical k-day log returns (no sqrt(T) scaling)
    r_k <- kday_log_returns(price_list[[idx]], k = h)
    row <- list(Index = idx, Horizon_Days = h)
    for (cl in confidence_levels) {
      v_log <- hs_var(r_k, cl = cl)     # log-return
      # Convert log-return to percentage for human readability
      row[[paste0("VaR_", as.integer(cl * 100), "%")]] <- 100 * v_log
    }
    results[[length(results) + 1]] <- row
  }
}

var_table <- bind_rows(results) %>%
  mutate(
    Index      = factor(Index, levels = indices_prefer),
    `VaR_95%`  = sprintf("%.2f%%", `VaR_95%`),
    `VaR_99%`  = sprintf("%.2f%%", `VaR_99%`)
  ) %>%
  arrange(Index, Horizon_Days)

# ---- 1-day volatility (sample SD of 1-day log returns) ----
vol_1d <- sapply(one_day_lr[, indices], sd, na.rm = TRUE)
vol_table <- tibble(
  Index = factor(names(vol_1d), levels = indices_prefer),
  `Volatility_1D` = sprintf("%.2f%%", 100 * vol_1d)
) %>% arrange(Index)

# ---- Print results (human-friendly) ----
cat("\n==================== RESULTS ====================\n")
cat("Historical Simulation VaR (log-returns -> shown as %)\n")
print(var_table)
cat("\n1-Day Volatility (log-returns, sample SD)\n")
print(vol_table)
cat("=================================================\n\n")

# ---- Optional: Save tables to CSV (good for GitHub artifacts) ----
# write.csv(var_table, "hs_var_table.csv", row.names = FALSE)
# write.csv(vol_table, "volatility_1d.csv", row.names = FALSE)

# ---- Plot: histogram of 1-day returns with VaR lines (95%, 99%) for DJIA ----
# Pick one index for a quick visual; you can loop if you like.
plot_index <- "DJIA"
r1 <- one_day_lr[[plot_index]]

# Compute VaR cutoffs for the plot (1-day returns)
v95 <- hs_var(r1, cl = 0.95)
v99 <- hs_var(r1, cl = 0.99)

p <- ggplot(data.frame(r = r1), aes(x = r)) +
  geom_histogram(bins = 40) +
  geom_vline(xintercept = v95, linetype = "dashed") +
  geom_vline(xintercept = v99, linetype = "dotted") +
  labs(
    title = paste0(plot_index, ": 1-Day Return Distribution with VaR Cutoffs"),
    subtitle = paste0("95% VaR = ", sprintf("%.2f%%", 100 * v95),
                      " | 99% VaR = ", sprintf("%.2f%%", 100 * v99)),
    x = "1-Day Log Return", y = "Frequency"
  )

print(p)

message("Done. If you want CSV outputs, un-comment the write.csv lines above.")