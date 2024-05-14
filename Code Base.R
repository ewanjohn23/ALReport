## Import Packages

library(tidyverse)
library(readxl)
library(factoextra)

## Import Data and Data Engineering

df_general <- read_excel('imported_data.xlsx', col_names = TRUE, 1, .name_repair = "minimal")
df_underwriting <- read_excel('imported_data.xlsx', col_names = TRUE, 2, .name_repair = "minimal")

col_years_general = rep(unlist(unname(df_general[1,-1])), 325)
df_general_pivoted <- df_general[-1,] %>%
  set_names(replace(names(df_general), names(df_general) == "", "Firm Number")) %>%
  pivot_longer(cols = -`Firm Number`, names_to = "Metric", values_to = "Value") %>%
  add_column(col_years_general) %>%
  rename( "Year" = "col_years_general" ) %>%
  pivot_wider(names_from = `Metric`, values_from = `Value`) %>%
  mutate("Year" = str_sub(`Year`, 1, str_length(`Year`)-2)) %>% 
  mutate_at(vars(-c(`Firm Number`, `Year`)), as.numeric) %>%
  mutate("NWP/GWP" = `NWP (£m)`/`GWP (£m)`) %>%
  mutate("GWP Bucket" = case_when(`GWP (£m)` > 5000 ~ "Large",
                                  `GWP (£m)` > 500 ~ "Mid",
                                  `GWP (£m)` > 0 ~ "Small",
                                  `GWP (£m)` <= 0 ~ "None"))

col_years_underwriting = rep(unlist(unname(df_underwriting[1,-1])), 456)
df_underwriting_pivoted <- df_underwriting[-1,] %>%
  set_names(replace(names(df_underwriting), names(df_underwriting) == "", "Firm Number")) %>%
  pivot_longer(cols = -`Firm Number`, names_to = "Metric", values_to = "Value") %>%
  add_column(col_years_underwriting) %>%
  rename( "Year" = "col_years_underwriting" ) %>%
  pivot_wider(names_from = `Metric`, values_from = `Value`) %>%
  mutate_at(vars(-c(`Firm Number`, `Year`)), as.numeric) %>%
  mutate("Year" = str_sub(`Year`, 1, str_length(`Year`)-2))

df_joined <- left_join(df_general_pivoted, df_underwriting_pivoted, by = c("Firm Number", "Year"))

df_outliers_remove <- df_joined %>%
  filter((`Net combined ratio` > 100) | 
         (`Net combined ratio` < -100) |
         (`NWP/GWP` > 100) | 
         (`NWP/GWP` < -100) |
         (`SCR coverage ratio` > 100) |
         (`SCR coverage ratio` < -100))

df_joined_curated <- df_joined %>%
  filter(!(`Firm Number` %in% df_outliers_remove$`Firm Number`))

df_volatility <- df_joined_curated %>%
  group_by(`Firm Number`) %>%
  summarise(sd = sd(`Gross claims incurred (£m)`), count = sum(`Gross claims incurred (£m)` > 0), active = sum(((`Year` == "2020") & (`Gross claims incurred (£m)` != 0)))) %>%
  arrange(desc(sd)) %>%
  filter(`active` == 1) %>%
  head(10)

df_violation_5 <- df_joined_curated %>%
  group_by(`Firm Number`) %>%
  summarise(count = sum((`SCR coverage ratio` < 1) & (`SCR coverage ratio` != 0)), active = sum(((`Year` == "2020") & ((`SCR coverage ratio` != 0) & ( !is.null(`SCR coverage ratio`)))))) %>%
  arrange(desc(count)) %>%
  filter(`active` == 1) %>%
  filter(`count` == 5)

df_violation_4 <- df_joined_curated %>%
  group_by(`Firm Number`) %>%
  summarise(count = sum((`SCR coverage ratio` < 1) & (`SCR coverage ratio` != 0)), active = sum(((`Year` == "2020") & ((`SCR coverage ratio` != 0) & ( !is.null(`SCR coverage ratio`)))))) %>%
  arrange(desc(count)) %>%
  filter(`active` == 1) %>%
  filter(`count` == 4)

## GWP vs NWP

df_joined_curated %>%
  filter(`Year` == '2020') %>%
  ggplot(aes(x = `GWP (£m)`, y = `NWP (£m)`)) +
  geom_point(size = 3)  +
  theme_light()  +
  theme(axis.text = element_text(size = 18),
        axis.title = element_text(size = 18)) 
ggsave('EDA.png')

## Firm Size Table

df_joined_curated %>%
  filter(`Year` == 2020) %>%
  filter(`GWP Bucket` == "Large" | `GWP Bucket` == "Mid" | `GWP Bucket` == "Small") %>%
  group_by(`GWP Bucket`) %>%
  summarize(count = n(), 
            "SCR Coverage Ratio Mean" = mean(`SCR coverage ratio`),
            "NWP/GWP Mean" = mean(`NWP/GWP`),
            "Net Combined Ratio Mean" = mean(`Net combined ratio`)) 

## Volatile Firms

df_joined_curated %>%
  filter(`Firm Number` %in% df_volatility$`Firm Number`) %>%
  group_by(`Firm Number`) %>%
  ggplot(aes(x = `Year`, y = `Net combined ratio`, group = `Firm Number`, col = `Firm Number`)) +
  geom_line()  +
  scale_colour_brewer(palette = "Paired") + geom_line(size = 1.2) +
  theme_light() + 
  geom_hline(yintercept = 1, col = 'red', size = 2, linetype='dashed') +
  theme(axis.text.y = element_text(size = 18),
        axis.text.x = element_text(size =18),
        axis.title = element_text(size = 18),
        legend.title = element_text(size = 18),
        legend.text = element_text(size = 18))
ggsave('Profitability.png')

## Violation Firms

df_joined_curated %>%
  filter(`Firm Number` %in% df_violation_5$`Firm Number`) %>%
  group_by(`Firm Number`) %>%
  ggplot(aes(x = `Year`, y = `SCR coverage ratio`, group = `Firm Number`, col = `Firm Number`)) +
  geom_line() +
  scale_colour_brewer(palette = "Paired") + geom_line(size = 1.2) +
  theme_light() + 
  geom_hline(yintercept = 1, col = 'red', size = 2, linetype='dashed') + 
  theme(axis.text.y = element_text(size = 18),
        axis.text.x = element_text(size =18),
        axis.title = element_text(size = 18),
        legend.title = element_text(size = 18),
        legend.text = element_text(size = 18))
ggsave('Violation5.png')

df_joined_curated %>%
  filter(`Firm Number` %in% df_violation_4$`Firm Number`) %>%
  group_by(`Firm Number`) %>%
  ggplot(aes(x = `Year`, y = `SCR coverage ratio`, group = `Firm Number`, col = `Firm Number`)) +
  geom_line() +
  scale_colour_brewer(palette = "Paired") + geom_line(size = 1.2) +
  theme_light() + 
  geom_hline(yintercept = 1, col = 'red', size = 2, linetype='dashed') + 
  theme(axis.text.y = element_text(size = 18),
        axis.text.x = element_text(size =18),
        axis.title = element_text(size = 18),
        legend.title = element_text(size = 18),
        legend.text = element_text(size = 18))
ggsave('Violation4.png')

## Outliers Analysis

df_summary <- df_joined_curated %>%
  filter(`Year` == "2020") %>%
  summarise(across(.cols = is.numeric, 
                   .fns = list(Mean = mean, SD = sd), na.rm = TRUE, 
                   .names = "{col}_{fn}"))

df_outliers <- df_joined_curated %>%
  filter(`Year` == "2020") %>%
  filter((`NWP/GWP` > df_summary$`NWP/GWP_Mean`+2*df_summary$`NWP/GWP_SD`) |
         (`NWP/GWP` < df_summary$`NWP/GWP_Mean`-2*df_summary$`NWP/GWP_SD`) |
         (`SCR coverage ratio` > df_summary$`SCR coverage ratio_Mean`+2*df_summary$`SCR coverage ratio_SD`) |
         (`SCR coverage ratio` < df_summary$`SCR coverage ratio_Mean`-2*df_summary$`SCR coverage ratio_SD`)  |
         (`Net combined ratio` > df_summary$`Net combined ratio_Mean`+2*df_summary$`Net combined ratio_SD`) |
         (`Net combined ratio` < df_summary$`Net combined ratio_Mean`-2*df_summary$`Net combined ratio_SD`))

df_outliers_curated <- df_joined_curated %>%
  filter(`Year` == "2020") %>%
  filter((`NWP/GWP` > df_summary$`NWP/GWP_Mean`+2*df_summary$`NWP/GWP_SD`) |
           (`NWP/GWP` < df_summary$`NWP/GWP_Mean`-2*df_summary$`NWP/GWP_SD`) |
           (`SCR coverage ratio` < df_summary$`SCR coverage ratio_Mean`-2*df_summary$`SCR coverage ratio_SD`)  |
           (`Net combined ratio` > df_summary$`Net combined ratio_Mean`+2*df_summary$`Net combined ratio_SD`))


## Machine Learning Appendix

df_pca <- df_joined_curated %>% filter(`Year` == "2020")
fviz_nbclust(select(df_pca, -c("Firm Number", "Year", "NWP/GWP", "GWP Bucket")), method = 'wss', FUNcluster = kmeans, k.max = 15)

clus <- kmeans(select(df_pca, -c("Firm Number", "Year", "NWP/GWP", "GWP Bucket")), 5, iter.max = 10, nstart = 25)
fviz_cluster(clus, select(df_pca, -c("Firm Number", "Year", "NWP/GWP", "GWP Bucket")), stand = TRUE, pointsize = 2, 
             geom = c('point'), show.clust.cent = TRUE, ellipse.level = 0.9,
             ggtheme = theme_light(), main = '5 Cluster K-Means graph of values')  + 
  theme(axis.text.y = element_text(size = 18),
        axis.text.x = element_text(size =18),
        axis.title = element_text(size = 18),
        legend.title = element_text(size = 18),
        legend.text = element_text(size = 18))
ggsave('Clusters.png')
df_pca <- df_pca %>%
  cbind(clus$cluster)

colnames(df_pca)[22] <- 'Cluster'

df_pca %>%
  group_by(Cluster) %>%
  ggplot() +
  geom_bar(aes(fill = factor(`GWP Bucket`), x = factor(Cluster)), 
           stat = 'count', position = position_dodge(preserve = 'single'))  +
  theme_light()  +
  xlab('Cluster') +
  theme(legend.position="bottom") +
  scale_fill_discrete(name = 'Size of Firm') +
  ylab('Count of Firms')   + 
  theme(axis.text.y = element_text(size = 18),
        axis.text.x = element_text(size =18),
        axis.title = element_text(size = 18),
        legend.title = element_text(size = 18),
        legend.text = element_text(size = 18))
ggsave('ClustersSize.png')


df_pca %>%
  group_by(Cluster, `GWP Bucket`) %>%
  summarize(group_mean = mean(`GWP (£m)`, na.rm = TRUE)) %>%
  ggplot() +
  geom_bar(aes(fill = factor(`GWP Bucket`), x = factor(Cluster), y = group_mean), 
           stat = 'identity', position = position_dodge(preserve = 'single'))  +
  theme_light()  +
  xlab('Cluster') +
  theme(legend.position="bottom") +
  scale_fill_discrete(name = 'Size of Firm') +
  ylab('Mean of GWP (£m)')  + 
  theme(axis.text.y = element_text(size = 18),
        axis.text.x = element_text(size =18),
        axis.title = element_text(size = 18),
        legend.title = element_text(size = 18),
        legend.text = element_text(size = 18))
ggsave('ClustersGWP.png')

df_pca %>%
  group_by(Cluster, `GWP Bucket`) %>%
  summarize(group_mean = mean(`Gross claims incurred (£m)`, na.rm = TRUE)) %>%
  ggplot() +
  geom_bar(aes(fill = factor(`GWP Bucket`), x = factor(Cluster), y = group_mean), 
           stat = 'identity', position = position_dodge(preserve = 'single'))  +
  theme_light()  +
  xlab('Cluster') +
  theme(legend.position="bottom") +
  scale_fill_discrete(name = 'Size of Firm') +
  ylab('Mean of Gross Claims Incurred (£m)') + 
  theme(axis.text.y = element_text(size = 18),
        axis.text.x = element_text(size =18),
        axis.title = element_text(size = 18),
        legend.title = element_text(size = 18),
        legend.text = element_text(size = 18))
ggsave('ClustersClaims.png')

df_pca %>%
  group_by(Cluster, `GWP Bucket`) %>%
  summarize(group_mean = mean(`SCR coverage ratio`, na.rm = TRUE)) %>%
  ggplot() +
  geom_bar(aes(fill = factor(`GWP Bucket`), x = factor(Cluster), y = group_mean), 
           stat = 'identity', position = position_dodge(preserve = 'single'))  +
  theme_light()  +
  xlab('Cluster') +
  theme(legend.position="bottom") +
  scale_fill_discrete(name = 'Size of Firm') +
  ylab('Mean of SCR Coverage Ratio')  + 
  theme(axis.text.y = element_text(size = 18),
        axis.text.x = element_text(size =18),
        axis.title = element_text(size = 18),
        legend.title = element_text(size = 18),
        legend.text = element_text(size = 18))
ggsave('ClustersSCR.png')

df_pca %>%
  group_by(Cluster, `GWP Bucket`) %>%
  summarize(group_mean = mean(`Net combined ratio`, na.rm = TRUE)) %>%
  ggplot() +
  geom_bar(aes(fill = factor(`GWP Bucket`), x = factor(Cluster), y = group_mean), 
           stat = 'identity', position = position_dodge(preserve = 'single'))  +
  theme_light()  +
  xlab('Cluster') +
  theme(legend.position="bottom") +
  scale_fill_discrete(name = 'Size of Firm') +
  ylab('Mean of Net Combined Ratio')  + 
  theme(axis.text.y = element_text(size = 18),
        axis.text.x = element_text(size =18),
        axis.title = element_text(size = 18),
        legend.title = element_text(size = 18),
        legend.text = element_text(size = 18))
ggsave('ClustersNCR.png')