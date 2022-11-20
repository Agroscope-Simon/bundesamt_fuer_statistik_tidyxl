
library(tidyverse)
library(tidyxl)
library(unpivotr)
library(here)
library(readxl)

path <- here("Data", "Traubenernte.xlsx")
formats <- xlsx_formats(path)



# Loop --------------------------------------------------------------------



sheets <- as.character(c(1998:2021))
final_list <- list()

for (i in sheets) {

cells <-
  xlsx_cells(path, sheets = i) %>%
  dplyr::filter(!row == 1 , !is_blank) %>%
  behead_if(local_format_id == 34, direction =  "left-up", name = "area") %>%
  behead("up-left", "hektoliter") %>% # Treat every cell in every row as a header
  behead("up-left", "sorten") %>%
  behead("up-left", "farbe") |> 
  select(row, col, data_type, character, numeric, sheet, style_format, local_format_id, area, hektoliter, sorten, farbe)

final_list[[i]] <- cells
}


df <- bind_rows(final_list)


# Graph -------------------------------------------------------------------

df |> 
  count(hektoliter)


df |> 
  filter(hektoliter == "Weinmost in hl ") |> 
  filter(area == "Genferseeregion") |> 
  ggplot(aes(sheet, numeric, color = area)) +
  geom_col()



