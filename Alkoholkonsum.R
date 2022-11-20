
library(tidyverse)
library(tidyxl)
library(unpivotr)
library(here)
library(readxl)
library(tidyverse)

path <- here("Data", "Alkoholkonsum.xlsx")





cells <-
  xlsx_cells(path) %>%
  dplyr::filter(!is_blank) %>%
  select(row, col, data_type, character, numeric)

x <-
  cells %>%
  behead("left", "place") %>%
  behead("up-left", "category") %>%
  behead("up-left", "metric-cell-1") %>% # Treat every cell in every row as a header
  behead("up-left", "metric-cell-2") %>%
  behead("up-left", "metric-cell-3") %>%
  behead("up-left", "metric-cell-4") %>%
  behead("up-left", "metric-cell-5")
glimpse(x)
