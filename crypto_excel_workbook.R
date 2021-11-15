# R TIPS ----
# TIP 005: Automate Excel Workbooks with R ----
#
# 👉 For Weekly R-Tips, Sign Up Here: https://mailchi.mp/business-science/r-tips-newsletter

# 1.0 LIBRARIES ----

library(openxlsx)
library(tidyquant)
library(tidyverse)
library(timetk)

# 2.0 GET DATA ----

crypto_data_tbl <- c("BTC-USD", "ETH-USD", "DOGE-USD", "LTC-USD") %>%
    tq_get(from = "2010-01-01", to = "2021-11-14")

crypto_pivot_table <- crypto_data_tbl %>%
    pivot_table(
        .rows = ~ YEAR(date),
        .columns = ~ symbol,
        .values = ~ PCT_CHANGE_FIRSTLAST(adjusted)
    ) %>%
    rename(year = 1)

crypto_plot <- crypto_data_tbl %>%
    group_by(symbol) %>%
    plot_time_series(date, adjusted, .color_var = symbol, .facet_ncol = 2, .interactive = FALSE)

# 3.0 CREATE WORKBOOK ----

# * Initialize a workbook ----
wb <- createWorkbook()

# * Add a Worksheet ----
addWorksheet(wb, sheetName = "crypto_analysis")

# * Add Plot ----


print(crypto_plot)

wb %>% insertPlot(sheet = "crypto_analysis", startCol = "G", startRow = 3)

# * Add Data ----

writeDataTable(wb, sheet = "crypto_analysis", x = crypto_pivot_table)

# * Save Workbook ----
saveWorkbook(wb, "R_T_E/crypto_analysis.xlsx", overwrite = FALSE)

# * Open the Workbook ----
openXL("R_T_E/crypto_analysis.xlsx")
