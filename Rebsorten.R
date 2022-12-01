library(tidyverse)
library(readxl)
library(janitor)
library(here)
library(streamgraph)


# Import data -------------------------------------------------------------
# path to file
path = here("Data/Rebsorten_Schweiz_1999-2022.xlsx")

# Loop over all sheets


sheets <- excel_sheets(path)
file <- data.frame()
file_list <- list()

for (i in sheets){
  file <- read_excel(path = path, sheet = i)
  file |> 
    remove_empty("rows") |> 
    select(!starts_with("Total")) |> 
    replace("-", NA) |> 
    mutate(across(2:last_col(), as.numeric)) %>%
    select(-last_col()) |> 
    clean_names() |> 
    mutate(Jahr = i) -> file
 
  
  ncol(file)-1 -> col_number
    
 file <-  pivot_longer(file, cols = 2:col_number,
                 names_to = "Sorte",
                 values_to = "Anbauflaeache")
                 
 file <- file |>  filter(!Anbauflaeache == "NA")
    
 
 file ->  file_list[[i]]
} 

df <- bind_rows(file_list)


# Add column Kantone und Region

df<- df |> 
  mutate("kanton" = regionen) |> 
  coiu |> 
  mutate(regionen = case_when(grepl("Appenzell", regionen) ~ "Appenzell I.& A.", 
                            grepl("Aargau", regionen) ~ "Aargau", 
                            grepl("Luzern", regionen) ~ "Luzern", 
                            grepl("Uri", regionen) ~ "Uri/Obwalden/Nidwalden",
                            grepl("Nidwalden", regionen) ~ "Uri/Obwalden/Nidwalden",
                            grepl("Obwalden", regionen) ~ "Uri/Obwalden/Nidwalden",
                            grepl("Tessin", regionen) ~ "Tessin",
                            grepl("Wallis", regionen) ~ "Wallis",
                            grepl("Zug", regionen) ~ "Zug",
                            TRUE ~ as.character(regionen)))

df <- df |>
 mutate(regionen = case_when(regionen == "Waadt"  ~ "Genferseeregion",
                            regionen == "Genf"  ~ "Genferseeregion",
                            regionen == "Wallis"  ~ "Genferseeregion",
                            regionen == "Bern"  ~ "Espace Mittelland",
                            regionen == "Solothturn"  ~ "Espace Mittelland",
                            regionen == "Jura"  ~ "Espace Mittelland",
                            regionen == "Freiburg"  ~ "Espace Mittelland",
                            regionen == "Neuenburg"  ~ "Espace Mittelland",
                            regionen == "Basel-Stadt" ~ "Nordwestschweiz", 
                            regionen == "Basel-Landschaft" ~ "Nordwestschweiz", 
                            regionen == "Aargau" ~ "Nordwestschweiz", 
                            regionen == "Zürich" ~ "Zürich", 
                            regionen == "St. Gallen" ~ "Ostschweiz", 
                            regionen == "Glarus" ~ "Ostschweiz", 
                            regionen == "Schaffhausen" ~ "Ostschweiz", 
                            regionen == "Graubünden" ~ "Ostschweiz", 
                            regionen == "Thurgau" ~ "Ostschweiz", 
                            regionen == "Appenzell I.& A." ~ "Ostschweiz", 
                            regionen == "Zug" ~ "Zentralschweiz", 
                            regionen == "Luzern" ~ "Zentralschweiz", 
                            regionen == "Uri/Obwalden/Nidwalden" ~ "Zentralschweiz", 
                            regionen == "Schwyz" ~ "Zentralschweiz", 
                            regionen == "Tessin" ~ "Tessin", 
                            
                            TRUE~ as.character(regionen))) 



df <- df |> 
  filter(!grepl('Zentralschweiz|Tessin|Ostschweiz|Nordwestschweiz|Espace Mittelland|Genferseeregion|Zürich', kanton))


# Plot data ---------------------------------------------------------------


df %>%
  filter(kanton == "Wallis") |> 
  mutate(Anbauflaeache = as.integer(Anbauflaeache)) |> 
  group_by(Jahr, Sorte) %>%
  streamgraph(Sorte, Anbauflaeache, Jahr) |> 
  sg_axis_x(22, 1,  tick_format = NULL) 
  


  


 







