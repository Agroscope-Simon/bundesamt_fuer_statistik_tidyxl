
library(tidyverse)
library(tidyxl)
library(unpivotr)
library(here)
library(readxl)
library(tidyverse)

# input files

path <- here("Data", "Weinverbrauch.xlsx")
formats <- xlsx_formats(path)

# input cells
cells <-
  xlsx_cells(path) %>%
  filter(!is_blank, !row %in% c(1, 5,10, 111-121)) 

# find out local format ID of the year fields in excel
cells %>% 
  filter(!is.na(character)) %>% 
  filter(!character %in% c("Weiss", "Rot  ", "Rot", "Rot ")) %>%   
  view()

#  local ID = 22 and 28, define new colnames (wein, gruppe, untergruppe, jahr)
cells <-
  cells %>% 
  behead_if(local_format_id %in% c(22,28), direction =  "left-up", name = "jahr") %>% 
  behead("left", wein) %>% 
  behead("up-left", gruppe) %>% 
  behead("up", untergruppe) %>%
  select(row, col, data_type, character, numeric, jahr, wein, gruppe, untergruppe)


# print to console for future wrangling
cells %>% count(untergruppe)
cells %>% count(gruppe)

# rename 
cells <- cells %>% 
mutate(gruppe = case_when(grepl("Ausfuhr 2", gruppe) ~ "Ausfuhr", 
                          grepl("Inland-", gruppe) ~ "Inlandproduktion", 
                          grepl("Vorrat am Ende des Jahres 3", gruppe) ~ "Vorrat_Ende_Jahr", 
                          TRUE ~ as.character(gruppe))) %>% 
  mutate(untergruppe = case_when(grepl("ausländischer", untergruppe) ~ "wein_ausland", 
                            grepl("davon inländischer Wein", untergruppe) ~ "verbrauch_inland", 
                            grepl("inländischer", untergruppe) ~ "wein_inland", 
                            grepl("Total", untergruppe) ~ "verbrauch_total", 
                            grepl("Trinkwein", untergruppe) ~ "trinkwein", 
                            grepl("Wein zur Essig-", untergruppe) ~ "weinessig", 
                            TRUE ~ as.character(untergruppe))) %>% 
  mutate(jahr = str_sub(jahr, 1,4)) %>% 
  mutate(jahr = as.numeric(jahr)) %>% 
  mutate(jahr = ifelse(jahr < 2004, jahr+1, jahr)) %>% 
  rename(wein_hl = numeric) %>% 
  mutate(wein = case_when(grepl("Rot ", wein) ~ "Rot", 
         TRUE ~ as.character(wein)))



# delete redundant information

cells <- cells %>% 
  mutate(across(where(is_character),as_factor)) %>% 
  mutate(jahr = as.factor(jahr))


# filter(gruppe == "Inlandproduktion") %>% 
cells %>% 
  filter(!is.na(wein)) %>% 
  mutate(jahr = as.factor(jahr)) %>% 
  ggplot(aes(x=jahr, y=wein_hl, fill = wein))+
  geom_col() +
  facet_wrap(~ gruppe, scales = "free_y")
