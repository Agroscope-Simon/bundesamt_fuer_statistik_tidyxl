
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
  behead_if(local_format_id == 34, direction =  "left-up", name = "area") %>%
  behead("up-left", "hektoliter") %>% # Treat every cell in every row as a header
  behead("up-left", "sorten") %>%
  behead("up", "farbe") |> 
  behead("left", "Kanton") %>% 
  select(row, col, data_type, character, numeric, sheet, style_format, local_format_id, area, hektoliter, sorten, farbe, Kanton)

final_list[[i]] <- cells
}


df <- bind_rows(final_list)

# Delete redundant information

df <- df %>% 
  mutate(Kanton = if_else(area == "Zürich", "Zürich", Kanton)) %>% 
  mutate(Kanton = if_else(grepl("Tessin", area), "Tessin", Kanton)) 

 df <- df %>%  
  filter(!hektoliter == "Tafeltrauben in q") %>% 
  filter(!(row ==51 & area == "Tessin 5 ")) %>% 
  filter(!is.na(Kanton)) 
  
  df %>% filter(area == "Tessin")
 # Rename Categories
 
 df <- df %>% 
   mutate(area = case_when(grepl("Total 2", area) ~ "Total", 
                             grepl("Tessin 4", area) ~ "Tessin", 
                             grepl("Tessin 5", area) ~ "Tessin", 
                             TRUE ~ as.character(area))) %>% 
   mutate(sorten = case_when(grepl("europäische Reben", sorten) ~ "Europäische Rebsorten", 
                             grepl("Gesamt-", sorten) ~ "Gesamtmittel",
                             grepl("Interspezifische", sorten) ~ "Interspezifische Sorten",
                             grepl("Total 1", sorten) ~ "Total",
                             TRUE ~ as.character(sorten))) %>% 
   mutate(hektoliter = case_when(grepl("Weinmost in hl pro ha", hektoliter) ~ "Weinmost in l pro m²", 
                             grepl("Gesamt-", sorten) ~ "Gesamtmittel",
                             TRUE ~ as.character(hektoliter))) %>% 
   mutate(farbe = case_when(grepl("rote", farbe) ~ "Rote Trauben", 
                            grepl("Rote", farbe) ~ "Rote Trauben", 
                            grepl("rote ", farbe) ~ "Rote Trauben", 
                            grepl("weisse", farbe) ~ "Weisse Trauben", 
                            grepl("Weisse", farbe) ~ "Weisse Trauben", 
                                 TRUE ~ as.character(farbe))) %>% 
   mutate(Kanton = case_when(grepl("Uri", Kanton) ~ "Uri/Obwalden/Nidwalden", 
                            grepl("Obwalden / Nidwalden 3", Kanton) ~ "Uri/Obwalden/Nidwalden", 
                            grepl("Obwalden", Kanton) ~ "Uri/Obwalden/Nidwalden", 
                            grepl("Nidwalden", Kanton) ~ "Uri/Obwalden/Nidwalden", 
                            grepl("Uri / Obwalden / Nidwalden 4", Kanton) ~ "Uri/Obwalden/Nidwalden", 
                            grepl("Genf 2", Kanton) ~ "Genf", 
                            grepl("Genf 3", Kanton) ~ "Genf", 
                            grepl("Appenzell A. Rh.", Kanton) ~ "Appenzell A. Rh. / I. Rh.", 
                            grepl("Appenzell A. Rh. / I. Rh. 4", Kanton) ~ "Appenzell A. Rh. / I. Rh.", 
                            grepl("Appenzell A. Rh. / I. Rh.3", Kanton) ~ "Appenzell A. Rh. / I. Rh.", 
                            grepl("Appenzell I. Rh.", Kanton) ~ "Appenzell A. Rh. / I. Rh.", 
                            TRUE ~ as.character(Kanton))) %>% 
   rename(., Jahre = sheet) %>% 
   rename(., Region = area) %>% 
   rename(., Rebsortenfarbe = farbe) %>% 
   rename(., Rebsorte = sorten) %>% 
   rename(., Weinmost = hektoliter) %>% 
   rename(., Traubenernte = numeric) 
   
   
   # Calculate sums for Appenzell and Unterwalden
 
 df_sum <- df %>% 
   group_by(Jahre, Region, Weinmost, Rebsorte, Rebsortenfarbe, Kanton ) %>% 
   summarise(Result = sum(Traubenernte, na.rm = T)) %>% 
   ungroup()
 
 
 total_region_hl <- df_sum %>%
   filter(Weinmost == "Weinmost in hl ") %>% 
   filter(Rebsorte != "Total") %>% 
   group_by(Jahre, Region) %>% 
   summarise(Total_Kanton = sum(Result, na.rm = T)) 
 


 total_kanton_hl <- df_sum %>% 
   filter(Weinmost == "Weinmost in hl ") %>% 
   filter(Rebsorte != "Total") %>% 
   group_by(Jahre, Region, Kanton) %>% 
   summarise(Total_Kanton = sum(Result, na.rm = T)) 
 
 total_jahre_hl <- total_region_hl %>% 
      group_by(Jahre) %>% 
      summarise(Total_Jahre = sum(Total_Kanton, na.rm = T)) 
 
barplot_Stacked <- df_sum %>% 
  filter(Weinmost == "Weinmost in hl ") %>% 
  filter(Rebsorte != "Total") %>% 
  select(Jahre, Region, Kanton, Rebsortenfarbe, Result) 

# Graph -------------------------------------------------------------------


 
 cols <- brewer.pal(7, "Set1")

 total_region_hl %>%   
 hchart("streamgraph", hcaes(x = Jahre, y = Total_Kanton, group = Region)) %>%      # basic definition
   hc_colors(cols) %>%                                                        # COLOR
   # hc_xAxis(title = list(text="")) %>%                                    # x-axis
   hc_yAxis(title = list(text="Weinmost in hl"))  %>%                       # y-axis
   hc_chart(style = list(fontFamily = "Georgia",                  
                         fontWeight = "bold")) %>%                               # FONT
   hc_plotOptions(series = list(marker = list(symbol = "circle"))) %>%           # SYMBOLS        
   hc_legend(align = "right",                                         
             verticalAlign = "top") %>%                                       # LEGEND
   hc_tooltip(shared = F,                    
              borderColor = "black",
              pointFormat = "{point.Region}: {point.Total_Kanton:.2f}<br>") %>% 
   hc_title(
     text = "Traubenernte 1998-2021",
     margin = 20,
     align = "center",
     style = list( useHTML = TRUE)
   ) %>% 
   hc_subtitle(
     text = "Bundesamt für Statistik Schweiz",
     align = "center",
     style = list(fontWeight = "bold")
   ) 
   
   # TOOLTIP

 
 

 # Echarts4R
 
 total_region_hl |> 
   group_by(Region) |> 
   e_charts(x = Jahre) |> # initialise and set x
   e_river(serie = Total_Kanton) |> 
   e_title("Traubenernte 1998-2021", "Use the brush") |> # title
   e_theme("royal") |> # add a theme
   e_brush() |> # add the brush
   e_tooltip(trigger = "axis") |> 
   e_toolbox_feature(feature = "saveAsImage") |> # hit the download button!# Add tooltips
   e_toolbox_feature(feature = "dataZoom") |> 
   e_toolbox_feature(feature = "dataView") |> 
   e_image_g(
     right = 20,
     top = 20,
     z = -999,
     style = list(
       image = "https://thumbs.dreamstime.com/t/rotweinflasche-und-weinglas-68897210.jpg",
       width = 100,
       height = 100,
       opacity = .9
     )
   ) 
 