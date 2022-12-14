library(tidyverse)
library(tidyxl)
library(unpivotr)
library(here)
library(readxl)
library(highcharter)
library(htmlwidgets)
library(RColorBrewer)
library(echarts4r)

path <- here("Data", "Traubenernte.xlsx")
formats <- xlsx_formats(path)



# Loop --------------------------------------------------------------------



sheets <- as.character(c(1998:2021))
final_list <- list()

for (i in sheets) {
  
  cells <-
    xlsx_cells(path, sheets = i) %>%
    filter(!row %in% c(1,7,10, 52:60)) %>% 
    dplyr::filter( !is_blank) %>%
    behead_if(local_format_id == 34, direction =  "left-up", name = "region") %>%
    behead("up-left", "weinmost") %>% # Treat every cell in every row as a header
    behead("up-left", "gruppe") %>%
    behead("up", "farbe") |> 
    behead("left", "canton") %>% 
    select(row, col, data_type, character, numeric, sheet, style_format, local_format_id, region, weinmost, gruppe, farbe, canton)
  
  final_list[[i]] <- cells
}


df <- bind_rows(final_list)

# Delete redundant information

df <- df %>% 
  mutate(canton = if_else(region == "Zürich", "Zürich", canton)) %>% 
  mutate(canton = if_else(grepl("Tessin", region), "Tessin", canton)) 

df <- df %>%  
  filter(!weinmost == "Tafeltrauben in q") %>% 
  filter(!(row ==51 & region == "Tessin 5 "))




# Rename Categories

df <- df %>% 
  mutate(canton = case_when(grepl("Uri", canton) ~ "Uri/Obwalden/Nidwalden", 
                            grepl("Obwalden / Nidwalden 3", canton) ~ "Uri/Obwalden/Nidwalden", 
                            grepl("Obwalden", canton) ~ "Uri/Obwalden/Nidwalden", 
                            grepl("Nidwalden", canton) ~ "Uri/Obwalden/Nidwalden", 
                            grepl("Uri / Obwalden / Nidwalden 4", canton) ~ "Uri/Obwalden/Nidwalden", 
                            grepl("Genf 2", canton) ~ "Genf", 
                            grepl("Genf 3", canton) ~ "Genf", 
                            grepl("Appenzell A. Rh.", canton) ~ "Appenzell A. Rh. / I. Rh.", 
                            grepl("Appenzell A. Rh. / I. Rh. 4", canton) ~ "Appenzell A. Rh. / I. Rh.", 
                            grepl("Appenzell A. Rh. / I. Rh.3", canton) ~ "Appenzell A. Rh. / I. Rh.", 
                            grepl("Appenzell I. Rh.", canton) ~ "Appenzell A. Rh. / I. Rh.", 
                            TRUE ~ as.character(canton))) %>% 
  mutate(region = case_when(grepl("Total 2", region) ~ "Total_CH", 
                          grepl("Tessin 4", region) ~ "Tessin", 
                          grepl("Tessin 5", region) ~ "Tessin", 
                          TRUE ~ as.character(region))) %>% 
  mutate(gruppe = case_when(grepl("europäische Reben", gruppe) ~ "Europäische Rebsorten", 
                            grepl("Gesamt-", gruppe) ~ "Gesamtmittel",
                            grepl("Interspezifische", gruppe) ~ "Interspezifische Sorten",
                            grepl("Total 1", gruppe) ~ "Total__Region_Kanton",
                            TRUE ~ as.character(gruppe))) %>% 
  mutate(farbe = case_when(grepl("rote", farbe) ~ "Rote Trauben", 
                           grepl("Rote", farbe) ~ "Rote Trauben", 
                           grepl("rote ", farbe) ~ "Rote Trauben", 
                           grepl("weisse", farbe) ~ "Weisse Trauben", 
                           grepl("Weisse", farbe) ~ "Weisse Trauben", 
                           TRUE ~ as.character(farbe))) %>% 
  
  df |> 
  mutate(region = ifelse(, "Total_CH", region))

  rename(., Jahre = sheet) %>% 
  rename(., Traubenernte = numeric) 


df |> 
  filter(region == "Total_CH") |> 
  filter(weinmost == "Weinmost in hl ") |> view()
