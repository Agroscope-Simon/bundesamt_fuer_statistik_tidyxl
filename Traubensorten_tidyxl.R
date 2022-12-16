library(tidyverse)
library(tidyxl)
library(unpivotr)
library(here)
library(readxl)
library(highcharter)
library(htmlwidgets)
library(RColorBrewer)
library(echarts4r)

path <- here("Data", "Traubensorten.xlsx")
formats <- xlsx_formats(path)



# Loop --------------------------------------------------------------------



sheets <- as.character(c(1999:2020))
final_list <- list()

for (i in sheets) {
  
  cells <-
    xlsx_cells(path, sheets = i) %>%
    filter(!row %in% c(1:5,7:8,10:12, 52:60)) %>% 
    dplyr::filter( !is_blank) %>%
    behead_if(local_format_id == 25, direction =  "left-up", name = "region") %>%
    behead("up-left", "fläche") %>% # Treat every cell in every row as a header
    behead("up", "weinsorte") |> 
    behead("left", "kanton") %>% 
    select(row, col, data_type, character, numeric, sheet, style_format, local_format_id, region, fläche, weinsorte, kanton)
  
  final_list[[i]] <- cells
}


df <- bind_rows(final_list)
