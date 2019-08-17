library(data.table)
library(tidyverse)
library(ggrepel)
library(ggdark)
library(readxl)
library(ggthemes)

# Create a list of all the excel files
my_dir <- "data/"
file_list <- paste(my_dir, list.files(path = my_dir, pattern = "*.xlsx"), sep = "/")

# Read the files in one list
my_list <- lapply(file_list, read_excel)

# Generate a vector of the names of each dataframe in the generated list
# Most of the names are displayed in the first row & first column of the dataframe


df_names <- unlist(sapply(my_list, function(elem) elem[1, 1])) %>% tbl_df()

# OPEC Members' values of petroleum exports (m $)
df_values <- my_list[[8]]
setDT(df_values)
noms <- as.character(df_values[2])
noms[1] <- 'country'
names(df_values) <- noms
df_values <- df_values[3:16]
df_values <- melt.data.table(df_values, id.vars = "country", measure.vars = 2:ncol(df_values),
                            variable.name = "year", value.name = "petroleum_exp", value.factor = FALSE)
df_values <- df_values[, c("year", "petroleum_exp") := 
                         .(as.integer(as.character(year)), round(as.numeric(petroleum_exp), 2))]

df_values[country %like% "Algeria"] %>% 
  ggplot() +
  aes(year, petroleum_exp/1000, color = country) +
  geom_line(size = 1, color = "red") +
  geom_point(color = "black", alpha = .5) +
  geom_line(data = df_values[, median(petroleum_exp/1000, na.rm = TRUE), by = year],
            mapping = aes(x = year, V1), color = "black", alpha = .5, size = 1, linetype = "dashed") +
  geom_label_repel(data = df_values[year == 2017 & country %like% "Algeria"], 
                   mapping =  aes(x = 2018, y = petroleum_exp/1000, label = "Algeria"),
                   hjust = -.5, color = "red", face = "bold") +
  geom_label_repel(data = df_values[year == 2017, .("opec_median" = median(petroleum_exp/1000, na.rm = TRUE))], 
                   mapping =  aes(x = 2018, y = opec_median, label = "OPEC_median"),
                   hjust = -.5, color = "gray50", face = "bold") +
  scale_x_continuous(limits = c(1990, 2018), breaks = seq(1990, 2018, 3)) +
  scale_y_continuous(limits = c(0, 80), breaks = seq(0, 80, 10)) +
  labs(
    title = "Values of petroleum exports (billion U.S. dollars)",
    x = "",
    y = "",
    caption = "Source : © 2019 Organization of the Petroleum Exporting Countries"
  ) +
  theme_economist() +
  theme(legend.position = "none")

ggsave(filename = "values_of_petroleum_exports.png", path = "figs/", width = 12, height = 6, dpi = 300)

#  Active rigs by country
df_rigs <- as.data.table(my_list[[12]])
noms <- as.character(df_rigs[2, ])
names(df_rigs) <- noms

df_rigs <- melt.data.table(df_rigs[3:52], id.vars = c("country", "region"), measure.vars = 3:38,
                           variable.name = "year", value.name = "number_rigs", variable.factor = FALSE)
df_rigs <- df_rigs[, c("year", "number_rigs") := .(as.integer(year), as.integer(number_rigs))]

zoom_country <- c("United States", "China1", "Canada", "Russia", "Saudi Arabia")
df_rigs[country %in% zoom_country] %>% 
  ggplot() +
  aes(year, number_rigs, color = country) +
  geom_line(size = 1) +
  scale_x_continuous(limits = c(1990, 2018), breaks = seq(1990, 2018, 3)) +
  scale_y_continuous(limits = c(0, 2100), breaks = seq(0, 2100, 300)) +
  # scale_y_sqrt() +
  geom_label_repel(data = df_rigs[year == 2017 & country %in% zoom_country], 
                   mapping =  aes(x = 2018, y = number_rigs, label = country),
                   hjust = -.5, face = "bold") +
  labs(
    title = "Active rigs by country",
    x = "",
    y = "",
    caption = "Source : © 2019 Organization of the Petroleum Exporting Countries"
  ) +
  theme(plot.title = element_text(size = 20, face = "bold"),
        panel.border = element_blank(),
        panel.background = element_blank(),
        plot.background = element_blank(),
        panel.grid.major.x= element_line(size=0.2,linetype="dotted", color="#6D7C83"),
        panel.grid.major.y= element_line(size=0.2,linetype="dotted", color="#6D7C83"),
        panel.grid.minor = element_blank(),
        axis.line.x = element_blank(),
        axis.line.y = element_blank(),
        axis.ticks.y = element_blank(),
        axis.ticks.x = element_blank(),
        legend.position = "none")

ggsave(filename = "active_rigs.png", path = "figs/", width = 12, height = 6, dpi = 300)
