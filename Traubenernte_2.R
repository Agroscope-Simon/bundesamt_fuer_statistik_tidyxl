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
                            grepl("Appenzell A. Rh.", canton) ~ "Appenzell", 
                            grepl("Appenzell A. Rh. / I. Rh. 4", canton) ~ "Appenzell", 
                            grepl("Appenzell A. Rh. / I. Rh.3", canton) ~ "Appenzell", 
                            grepl("Appenzell I. Rh.", canton) ~ "Appenzell", 
                            TRUE ~ as.character(canton))) %>% 
  mutate(region = case_when(grepl("Total 2", region) ~ "Total_CH", 
                          grepl("Tessin 4", region) ~ "Tessin", 
                          grepl("Tessin 5", region) ~ "Tessin", 
                          TRUE ~ as.character(region))) %>% 
  mutate(gruppe = case_when(grepl("europäische Reben", gruppe) ~ "Europäische_Rebsorten", 
                            grepl("Gesamt-", gruppe) ~ "Gesamtmittel",
                            grepl("Interspezifische", gruppe) ~ "Interspezifische_Sorten",
                            grepl("Total 1", gruppe) ~ "Total_Region_Kanton",
                            TRUE ~ as.character(gruppe))) %>% 
  mutate(farbe = case_when(grepl("rote", farbe) ~ "Rote_Trauben", 
                           grepl("Rote", farbe) ~ "Rote_Trauben", 
                           grepl("rote ", farbe) ~ "Rote_Trauben", 
                           grepl("weisse", farbe) ~ "Weisse_Trauben", 
                           grepl("Weisse", farbe) ~ "Weisse_Trauben", 
                           TRUE ~ as.character(farbe))) %>% 
  rename(jahre = sheet) %>%
  mutate(jahre = as.factor(jahre)) %>% 
  rename(Traubenernte = numeric) 

# Calculate sums for Appenzell and Unterwalden

df_sum <- df %>% 
  group_by(jahre, region, weinmost,gruppe, farbe, canton ) %>% 
  summarise(Traubenernte_in_hl = sum(Traubenernte, na.rm = T)) %>% 
  ungroup()


# Check out groups

df %>% 
  count(region)

df %>% 
  count(weinmost)

df %>% count(gruppe)

df %>% count(farbe)

# dataframes for charts

total_region_hl <- df_sum %>%
  filter(weinmost == "Weinmost in hl ") %>% 
  filter(gruppe == "Total_Region_Kanton") %>%
  filter(!region %in% c("Total", "Total_CH")) %>% 
group_by(jahre, region) %>%   
  summarise(Total_Kanton = sum(Traubenernte_in_hl, na.rm = T))
  
total_kanton_hl <- df_sum %>% 
  filter(weinmost == "Weinmost in hl ") %>% 
  filter(gruppe == "Total_Region_Kanton") %>% 
  group_by(jahre, region, canton) %>% 
  summarise(Total_Kanton = sum(Traubenernte_in_hl, na.rm = T)) 

total_jahre_hl <- total_region_hl %>% 
  group_by(jahre) %>% 
  summarise(Total_Jahre = sum(Total_Kanton, na.rm = T)) 

barplot_Stacked <- df_sum %>% 
  filter(weinmost == "Weinmost in hl ") %>% 
  filter(gruppe != "Total_Region_Kanton") %>% 
  select(jahre, region, canton, farbe, Traubenernte_in_hl) 

# Graph -------------------------------------------------------------------


# graph für region
col <- brewer.pal(7, "Set1")

total_region_hl %>%   
  hchart("streamgraph", hcaes(x = jahre, y = Total_Kanton, group = region)) %>%      # basic definition
  hc_colors(col) %>%                                                        # COLOR
  # hc_xAxis(title = list(text="")) %>%                                    # x-axis
  hc_yAxis(title = list(text="Weinmost in hl"))  %>%                       # y-axis
  hc_chart(style = list(fontFamily = "Georgia",                  
                        fontWeight = "bold")) %>%                               # FONT
  hc_plotOptions(series = list(marker = list(symbol = "circle"))) %>%           # SYMBOLS        
  hc_legend(align = "right",                                         
            verticalAlign = "top") %>%                                       # LEGEND
  hc_tooltip(shared = F,                    
             borderColor = "black",
             pointFormat = "{point.region}: {point.Total_Kanton:.2f}<br>") %>% 
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

# graph für kanton
col <- colorRampPalette(brewer.pal(9, "Spectral"))

total_kanton_hl %>%   
  hchart("streamgraph", hcaes(x = Jahre, y = Total_Kanton, group = Kanton)) %>%      # basic definition
  hc_colors(col(23)) %>%                                                        # COLOR
  # hc_xAxis(title = list(text="")) %>%                                    # x-axis
  hc_yAxis(title = list(text="Weinmost in hl"))  %>%                       # y-axis
  hc_chart(style = list(fontFamily = "Georgia",                  
                        fontWeight = "bold")) %>%                               # FONT
  hc_plotOptions(series = list(marker = list(symbol = "circle"))) %>%           # SYMBOLS        
  hc_legend(align = "right",                                         
            verticalAlign = "top") %>%                                       # LEGEND
  hc_tooltip(shared = F,                    
             borderColor = "black",
             pointFormat = "{point.Kanton}: {point.Total_Kanton:.2f}<br>") %>% 
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


df %>% 
  filter(weinmost == "Weinmost in hl ") %>% 
  filter(gruppe == "Total_Region_Kanton") %>% 
  ggplot(aes(Jahre,Traubenernte, fill = region)) + 
  geom_col() +
  facet_grid(~farbe)


df |> 
  mutate(region = ifelse(, "Total_CH", region))

  rename(., Jahre = sheet) %>% 
  rename(., Traubenernte = numeric) 


df |> 
  filter(region == "Total_CH") |> 
  filter(weinmost == "Weinmost in hl ") |> view()
