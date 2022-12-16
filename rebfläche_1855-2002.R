library(tidyverse)
library(tidyxl)
library(unpivotr)
library(here)
library(readxl)
library(tidyverse)

# input files

path <- here("Data", "Rebfläche_nach_Kantonen_1855-2002.xlsx")
cells <- read_excel("Data/Rebfläche_nach_Kantonen_1855-2002.xlsx", 
                    skip = 3)



# clean data fix cantons
df <- cells %>% 
  select(-36, -29) %>% 
  mutate(across(where(is_character), as.numeric)) %>% 
  mutate(Jura = replace_na(Jura, 0)) %>% 
  mutate(Bern = replace_na(Bern, 0)) %>% 
  mutate(Bern = Bern_Jura - Jura) %>% 
  mutate(Appenzell = Appenzell_A - Appenzell_I) %>% 
  rename(Schweiz = CH) %>% 
  slice(-c(1:4, 32:67)) %>% 
  select(-c("Bern_Jura", "Appenzell_A", "Appenzell_I"))
  
df  <- df %>% 
  mutate(across(everything(), ~ replace_na(.,0)))



# pivot

df <- df %>% 
  pivot_longer(c("Zürich":"Jura", "Appenzell", "Schweiz"), names_to = "Kanton", values_to = "Fläche") 

Zentralschweiz <- c("Uri", "Zug", "Obwalden", "Nidwalden", "Schwyz", "Luzern")
Tessin <- "Tessin"
Nordwestschweiz <- c("Basel-Stadt", "Basel-Land", "Aargau")
Zürich <- "Zürich"
Ostschweiz <- 
Genferseeregion <- c("Genf", "Waadt", "Wallis")



# Graphs

  


