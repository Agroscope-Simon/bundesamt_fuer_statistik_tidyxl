library(tidyverse)
library(tidyxl)
library(unpivotr)
library(here)
library(readxl)
library(highcharter)
library(htmlwidgets)
library(RColorBrewer)
library(echarts4r)
library(streamgraph)
library(janitor)
library(gganimate)
library(gifski)
library(png)
library(tweenr)
library(transformr)

path <- here("Data", "Traubensorten.xlsx")
formats <- xlsx_formats(path)



# Loop --------------------------------------------------------------------



sheets <- as.character(c(1999:2020))
final_list <- list()

for (i in sheets) {
  
  cells <-
    xlsx_cells(path, sheets = i) %>%
    filter(!row %in% c(1:5,7:8,10:12, 52:60)) %>% 
    filter( !is_blank) %>% 
    behead_if(local_format_id %in% c(25,50), direction =  "left-up", name = "region") %>% 
    behead("up-left", "fläche") %>% # Treat every cell in every row as a header
    behead("up", "weinsorte") |>
    behead("left", "kanton") %>%
    select(row, col, data_type, character, numeric, sheet, style_format, local_format_id, region, fläche, weinsorte, kanton)
  
  final_list[[i]] <- cells
}

   
df <- bind_rows(final_list)

# check variables

df %>% count(sheet)

df %>% count(fläche)

df %>% count(weinsorte) %>% view()

df %>% count(kanton) %>% view() 

df %>% count(region) %>% view()

# tidy

df <- df %>% 
  mutate(kanton = if_else(region == "Zürich", "Zürich", kanton)) %>% 
  mutate(kanton = if_else(grepl("Tessin", region), "Tessin", kanton)) 

df <- df %>% 
  mutate(kanton = case_when(grepl("Uri", kanton) ~ "Uri/Obwalden/Nidwalden", 
                            grepl("Obwalden / Nidwalden 3", kanton) ~ "Uri/Obwalden/Nidwalden", 
                            grepl("Obwalden", kanton) ~ "Uri/Obwalden/Nidwalden", 
                            grepl("Nidwalden", kanton) ~ "Uri/Obwalden/Nidwalden", 
                            grepl("Uri / Obwalden / Nidwalden 4", kanton) ~ "Uri/Obwalden/Nidwalden", 
                            grepl("Genf 2", kanton) ~ "Genf", 
                            grepl("Genf 3", kanton) ~ "Genf", 
                            grepl("Appenzell A. Rh.", kanton) ~ "Appenzell", 
                            grepl("Appenzell A. Rh. ", kanton) ~ "Appenzell", 
                            grepl("Appenzell A. Rh. /  I. Rh. 1", kanton) ~ "Appenzell", 
                            grepl("Appenzell A. Rh. / I. Rh. 1", kanton) ~ "Appenzell", 
                            grepl("Appenzell A. Rh. 1", kanton) ~ "Appenzell", 
                            grepl("Appenzell I. Rh.", kanton) ~ "Appenzell", 
                            grepl("Appenzell I. Rh. 1", kanton) ~ "Appenzell", 
                            grepl("Aargau ", kanton) ~ "Aargau",
                            grepl("Aargau 1", kanton) ~ "Aargau",
                            grepl("Luzern 1)", kanton) ~ "Luzern",
                            grepl("Nidwalden 1)", kanton) ~ "Uri/Obwalden/Nidwalden",
                            grepl("Nidwalden", kanton) ~ "Uri/Obwalden/Nidwalden",
                            grepl("Obwalden 1)", kanton) ~ "Uri/Obwalden/Nidwalden",
                            grepl("Obwalden", kanton) ~ "Uri/Obwalden/Nidwalden",
                            grepl("Uri", kanton) ~ "Uri/Obwalden/Nidwalden",
                            grepl("Uri / Obwalden/ Nidwalden 2", kanton) ~ "Uri/Obwalden/Nidwalden",
                            grepl("Wallis 1", kanton) ~ "Wallis",
                            grepl("Zug 1", kanton) ~ "Zug",
                            TRUE ~ as.character(kanton)))

df <- df %>% 
  mutate(weinsorte = case_when(grepl("Cabernet-", weinsorte) ~ "Cabernet-Sauvignon", 
                            grepl("Char-", weinsorte) ~ "Chardonnay", 
                            grepl("Chardon-", weinsorte) ~ "Chardonnay", 
                            grepl("Gewürz-", weinsorte) ~ "Gewürztraminer", 
                            grepl("Gutedel /", weinsorte) ~ "Gutudel/Chasselas", 
                            grepl("Interspezifische", weinsorte) ~ "Interspezifische_rote_sorten", 
                            grepl("Müller-  ", weinsorte) ~ "Müller-Thurgau", 
                            grepl("Pinot ", weinsorte) ~ "Pinot gris", 
                            grepl("Pinot", weinsorte) ~ "Pinot noir", 
                            grepl("Räusch-", weinsorte) ~ "Räuschling", 
                            grepl("Sylvaner ", weinsorte) ~ "Sylvaner", 
                            grepl("Übrige ", weinsorte) ~ "Übrige_weisse_Sorten", 
                            grepl("Übrige europäische", weinsorte) ~ "Übrige_europ_rote_sorten",
                            grepl("Übrige rote", weinsorte) ~ "Übrige_rote_Sorten",
                            grepl("Blauburgunder", weinsorte) ~ "Pinot noir", 
                            TRUE ~ as.character(weinsorte)))

df <- df %>% 
  mutate(region = case_when(grepl("Tessin 2", region) ~ "Tessin",
                            TRUE ~ as.character(region)))

df <- df %>% 
  mutate(fläche = case_when(grepl("Rebfläche", fläche) ~ "Total Rebsorten",
                            TRUE ~ as.character(fläche)))



df <- df %>% 
  rename(jahre = sheet) %>%
  # mutate(jahre = as.factor(jahre)) %>% 
  rename(Rebfläche_ha = numeric) 

# Calculate sums for Appenzell and Unterwalden

df_sum <- df %>% 
  group_by(jahre, region, kanton, fläche, weinsorte) %>% 
  summarise(Rebfläche_ha = sum(Rebfläche_ha, na.rm = T)) %>% 
  ungroup()


# cumsum kanton

df_cumsum <- df_sum %>% 
  filter(region == "Total") %>% 
  filter(!fläche == "Total Rebsorten") %>% 
  filter(!weinsorte == "Total") 


df_cumsum_weinsorte <- df_cumsum %>% 
  group_by(kanton,weinsorte) %>% 
  dplyr::mutate(csum = cumsum(Rebfläche_ha))

write_csv(df_cumsum_weinsorte, here("output", "df_cumsum_weinsorte"))

# cumsum total weinsorten

df_cumsum <- df_sum %>% 
  filter(region == "Total") %>% 
  filter(!fläche == "Total Rebsorten") %>% 
  filter(!weinsorte == "Total") %>% 
  filter(is.na(kanton))

df_cumsum_total <- df_cumsum %>% 
  filter(!weinsorte == "Übrige_weisse_Sorten") %>% 
  group_by(weinsorte) %>% 
  mutate(csum = cumsum(Rebfläche_ha))

df_cumsum_total <- df_cumsum_total %>% 
  pivot_wider(id_cols = weinsorte, 
              names_from = jahre, 
              values_from = Rebfläche_ha) %>% 
  mutate(across(everything(), ~ replace_na(.,0))) %>% 
  mutate(across(everything(), ~ round(.,0)))

write_csv(df_cumsum_total, here("output", "df_cumsum_weinsorte"))

#  datasets

total_rebsorten <- df_sum %>% 
  filter(region == "Genferseeregion") %>% 
  # filter(fläche == "Rote Rebsorten") %>% 
  filter(is.na(kanton)) %>% 
  filter(!weinsorte == "Total") %>% 
  mutate(jahre = as.factor(jahre)) %>% 
  mutate(jahre = fct_inorder(jahre))
    
    
   
         # graph für region
col <- colorRampPalette(brewer.pal(11, "Spectral"))
         
         total_rebsorten %>%   
           hchart("streamgraph", hcaes(x = jahre, y = Rebfläche_ha, group = weinsorte)) %>%      # basic definition
           hc_colors(col(18)) %>%                                                        # COLOR
           # hc_xAxis(title = list(text="")) %>%                                    # x-axis
           hc_yAxis(title = list(text="Rebfläche in Aren"))  %>%                       # y-axis
           hc_chart(style = list(fontFamily = "Georgia",                  
                                 fontWeight = "bold")) %>%                               # FONT
           hc_plotOptions(series = list(marker = list(symbol = "circle"))) %>%           # SYMBOLS        
           hc_legend(align = "right",                                         
                     verticalAlign = "top") %>%                                       # LEGEND
           hc_tooltip(shared = F,                    
                      borderColor = "black",
                      pointFormat = "{point.weinsorte}: {point.Rebfläche_ha:.2f}<br>") %>% 
           hc_title(
             text = "Rebflächen und Sorten 1999-2020",
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
         
         
    
    


# bar race -----------------------------------------------------------------


    
         
         df_barrace <- df_cumsum %>%
           filter(!weinsorte %in% c("Interspezifische_rote_sorten","Übrige_weisse_Sorten")) %>% 
           group_by(jahre) %>%
           # The * 1 makes it possible to have non-integer ranks while sliding
           mutate(rank = rank(-Rebfläche_ha),
                  Value_rel = Rebfläche_ha/Rebfläche_ha[rank==1],
                  Value_lbl = paste0(" ",round(Rebfläche_ha, 0))) %>%
           group_by(weinsorte) %>% 
           filter(rank <=15) %>%
           ungroup()
         
         # Animation
         mypalette<-brewer.pal(7,"Greens")
         
         anim <- ggplot(df_barrace, aes(rank, group = weinsorte, 
                                           fill = as.factor(weinsorte), color = as.factor(weinsorte))) +
           geom_tile(aes(y = Rebfläche_ha/2,
                         height = Rebfläche_ha,
                         width = 0.9), alpha = 0.8, color = NA) +
           geom_text(aes(y = 0, label = paste(weinsorte, " ")), vjust = 0.2, hjust = 1) +
           geom_text(aes(y=Rebfläche_ha,label = Value_lbl, hjust=0)) +
           coord_flip(clip = "off", expand = FALSE) +
           scale_y_continuous(labels = scales::comma) +
           scale_x_reverse() +
           guides(color = FALSE, fill = FALSE) +
           theme(axis.line=element_blank(),
                 axis.text.x=element_blank(),
                 axis.text.y=element_blank(),
                 axis.ticks=element_blank(),
                 axis.title.x=element_blank(),
                 axis.title.y=element_blank(),
                 legend.position="none",
                 panel.background=element_blank(),
                 panel.border=element_blank(),
                 panel.grid.major=element_blank(),
                 panel.grid.minor=element_blank(),
                 panel.grid.major.x = element_line( size=.1, color="grey" ),
                 panel.grid.minor.x = element_line( size=.1, color="grey" ),
                 plot.title=element_text(size=20, hjust=0.5, face="bold", colour="black", vjust=1),
                 plot.subtitle=element_text(size=14, hjust=0.5, face="italic", color="grey"),
                 plot.caption =element_text(size=8, hjust=0.5, face="italic", color="black"),
                 plot.background=element_blank(),
                 plot.margin = margin(2,2, 2, 4, "cm")) +
           transition_states(jahre, transition_length = 4, state_length = 1, wrap = FALSE) +
           view_follow(fixed_x = TRUE)  +
           labs(title = 'Rebflächen und Sorten 1999-2020: {closest_state}',  
                subtitle  =  "Top 15 Rot- und Weissweine",
                caption  = "Rebfläche in Aren | Data Source: Bundesamt für Statistik Schweiz") 
         
         # run animation in viewer
         anim
         
         # For GIF
         
         animate(anim, 200, fps = 20,  width = 1200, height = 1000, 
                 renderer = gifski_renderer("gganim.gif"), end_pause = 15, start_pause =  15) 
    
    
