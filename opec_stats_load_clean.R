library(data.table)
library(tidyverse)
library(ggrepel)
library(ggdark)
library(readxl)
library(ggthemes)
library(janitor)

# Create a list of all the excel files
my_dir <- "data"
file_list <- paste(my_dir, list.files(path = my_dir, pattern = "*.xlsx"), sep = "/")

# Read the files in one list
my_list <- lapply(file_list, read_excel)

# Generate a vector of the names of each dataframe in the generated list
# Most of the names are displayed in the first row & first column of the dataframe

df_names <- unlist(lapply(my_list, function(elem) elem[1, 1])) %>% tbl_df()

names(my_list) <- df_names$value

# Spot prices ($/b)
# Load spot prices dt 
dt_prices <- as.data.table(my_list[[1]])
# Remove empty columns (6 & 7)
dt_prices <- dt_prices[, !c(6:7)]
# Extract variables name
prices_nom <- as.character(dt_prices[1, ])
prices_nom[1] <- 'year'
# Set variables name
names(dt_prices) <- prices_nom %>% make_clean_names()
# remove rows 1 to 4
dt_prices <- dt_prices[5:51
                       ][, 'year' := as.integer(year) 
                       ][, lapply(.SD, as.numeric), .SDcols = 2:5, by = year
                         ][, lapply(.SD, round, 2), .SDcols = 2:5, by = year]

# Table 1.1: OPEC Members' facts and figures, 2017
dt_facts <- as.data.table(my_list[[2]])
# Extract variables name
noms <- as.character(dt_facts[2])
noms[1] <- 'variables' 
names(dt_facts) <- noms
# removing empty rows
dt_facts <- dt_facts[3:23]
# Melting the country variable
dt_facts <- melt.data.table(dt_facts, id.vars = 'variables', measure.vars = 2:ncol(dt_facts), variable.name = 'country')
# dcast the variable column
dt_facts <- dcast.data.table(dt_facts, country ~ variables) 
# Clean dt variables name
names(dt_facts) <- names(dt_facts) %>% make_clean_names()
dt_facts <- dt_facts[, lapply(.SD, as.numeric), .SDcols=2:ncol(dt_facts), by = country]



# Table 1.2: OPEC Members' crude oil production allocations (1,000 b/d)
dt_oil_prod <- as.data.table(my_list[[3]])

# Table 2.1: OPEC Members' population (million inhabitants)
dt_population <- as.data.table(my_list[[4]])

# Table 2.2: OPEC Members' GDP at current market prices (m $)
dgp_current_price <- as.data.table(my_list[[5]])

# Table 2.3: OPEC Members real GDP growth rates PPP based weights (%)
gdp_growth <- as.data.table(my_list[[6]])

# Table 2.4: OPEC Members' values of exports (m $)
dt_values_exp <- as.data.table(my_list[[7]])

# Table 2.5:OPEC Members' values of petroleum exports (m $)
dt_values <- as.data.table(my_list[[8]])
# Extract variables name
noms <- as.character(dt_values[2])
noms[1] <- 'country'
names(dt_values) <- noms
# Remove empty rows
dt_values <- dt_values[3:16]
# Melt the year variable
dt_values <- melt.data.table(dt_values, id.vars = "country", measure.vars = 2:ncol(dt_values),
                            variable.name = "year", value.name = "petroleum_exp", value.factor = FALSE)
dt_values <- dt_values[, c("year", "petroleum_exp") := 
                         .(as.integer(as.character(year)), round(as.numeric(petroleum_exp), 2))]

dt_values[country %like% "Algeria"] %>% 
  ggplot() +
  aes(year, petroleum_exp/1000, color = country) +
  geom_line(size = 1, color = "red") +
  geom_point(color = "black", alpha = .5) +
  geom_line(data = dt_values[, median(petroleum_exp/1000, na.rm = TRUE), by = year],
            mapping = aes(x = year, V1), color = "black", alpha = .5, size = 1, linetype = "dashed") +
  geom_line(data = dt_prices[, .(year, brent)],
            mapping = aes(x = year, y = brent), color = "blue", alpha = .5, size = 1) +
  geom_label_repel(data = dt_values[year == 2017 & country %like% "Algeria"], 
                   mapping =  aes(x = 2018, y = petroleum_exp/1000, label = "Algeria"),
                   hjust = -.5, color = "red", face = "bold") +
  geom_label_repel(data = dt_values[year == 2017, .("opec_median" = median(petroleum_exp/1000, na.rm = TRUE))], 
                   mapping =  aes(x = 2018, y = opec_median, label = "OPEC_median"),
                   hjust = -.5, color = "gray50", face = "bold") +
  scale_x_continuous(limits = c(1990, 2018), breaks = seq(1990, 2018, 3)) +
  scale_y_continuous(limits = c(0, 120), breaks = seq(0, 120, 10)) +
  labs(
    title = "Values of petroleum exports (billion U.S. dollars)",
    subtitle = 'Evolution of the brent price in blue line',
    x = "",
    y = "",
    caption = "Source : © 2019 Organization of the Petroleum Exporting Countries"
  ) +
  theme_economist() +
  theme(legend.position = "none")

ggsave(filename = "values_of_petroleum_exports.png", path = "figs/", width = 12, height = 6, dpi = 300)

#  Active rigs by country
dt_rigs <- as.data.table(my_list[[12]])
noms <- as.character(dt_rigs[2, ])
names(dt_rigs) <- noms

dt_rigs <- melt.data.table(dt_rigs[3:52], id.vars = c("country", "region"), measure.vars = 3:38,
                           variable.name = "year", value.name = "number_rigs", variable.factor = FALSE)
dt_rigs <- dt_rigs[, c("year", "number_rigs") := .(as.integer(year), as.integer(number_rigs))]

zoom_country <- c("United States", "China1", "Canada", "Russia", "Saudi Arabia")
dt_rigs[country %in% zoom_country] %>% 
  ggplot() +
  aes(year, number_rigs, color = country) +
  geom_line(size = 1) +
  scale_x_continuous(limits = c(1990, 2018), breaks = seq(1990, 2018, 3)) +
  scale_y_continuous(limits = c(0, 2100), breaks = seq(0, 2100, 300)) +
  # scale_y_sqrt() +
  geom_label_repel(data = dt_rigs[year == 2017 & country %in% zoom_country], 
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

