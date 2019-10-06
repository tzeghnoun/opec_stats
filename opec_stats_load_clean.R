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

# 1 Table : Spot prices ($/b)
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

# 2  Table 1.1: OPEC Members' facts and figures, 2017
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



# 3  Table 1.2: OPEC Members' crude oil production allocations (1,000 b/d)
dt_oil_prod <- as.data.table(my_list[[3]])

# 4  Table 2.1: OPEC Members' population (million inhabitants)
dt_population <- as.data.table(my_list[[4]])

# 5  Table 2.2: OPEC Members' GDP at current market prices (m $)
dgp_current_price <- as.data.table(my_list[[5]])

# 6  Table 2.3: OPEC Members real GDP growth rates PPP based weights (%)
gdp_growth <- as.data.table(my_list[[6]])

# 7  Table 2.4: OPEC Members' values of exports (m $)
dt_values_exp <- as.data.table(my_list[[7]])

# 8  Table 2.5:OPEC Members' values of petroleum exports (m $)
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

# 9  Table 2.6: OPEC Members' values of imports (m $)
dt_values_imp <- as.data.table(my_list[[9]])

# 10  Table 2.7: Current account balances in OPEC Members (m $)
dt_current_balance <- as.data.table(my_list[[10]])

# 11  Table 3.1: World proven crude oil reserves by country (m b)
dt_world_oil_reserves <- as.data.table(my_list[[11]])

#  12  Table 3.2: Active rigs by country
dt_rigs <- as.data.table(my_list[[12]])
noms <- as.character(dt_rigs[2, ])
names(dt_rigs) <- noms

dt_rigs <- melt.data.table(dt_rigs[3:52], id.vars = c("country", "region"), measure.vars = 3:38,
                           variable.name = "year", value.name = "number_rigs", variable.factor = FALSE)
dt_rigs <- dt_rigs[, c("year", "number_rigs") := .(as.integer(year), as.integer(number_rigs))]

# 13  Table 3.3: Wells completed  in OPEC Members
dt_wells_completed <- as.data.table(my_list[[13]])

# 14  Table 3.4: Producing wells in OPEC Members
dt_producing_wells <- as.data.table(my_list[[14]])

# 15  Table 3.5: Daily and cumulative crude oil production in OPEC Members (1,000 b)
dt_daily_cumul_oil_prod <- as.data.table(my_list[[15]])

# 16  Table 3.6: World crude oil production by country (1,000 b/d)
dt_world_oil_prod <- as.data.table(my_list[[16]])

# 17	Table 3.7: Non-OPEC oil supply and OPEC NGLs (1,000 b/d)
dt_nonopec_opecngls <- as.data.table(my_list[[17]])

# 18	Table 4.1: Refinery capacity in OPEC Members by company and location (1,000 b/cd)
dt_opec_ref_cap <- as.data.table(my_list[[18]])

# 19	Table 4.2: Charge refinery capacity in OPEC Members, 2017 (1,000 b/cd)
dt_opec_ref_charge <- as.data.table(my_list[[19]])

# 20	Table 4.3: World refinery capacity by country (1,000 b/cd)
dt_world_ref_cap <- as.data.table(my_list[[20]])

# 21	Table 4.4: World refinery throughput by country (1,000 b/d)
dt_world_ref_thr <- as.data.table(my_list[[21]])

# 22	Table 4.5: Output of petroleum products in OPEC Members (1,000 b/d)
dt_opec_output_products <- as.data.table(my_list[[22]])

# 23	Table 4.6: World output of petroleum products by country (1,000 b/d)
dt_world_output_products <- as.data.table(my_list[[23]])

# 24	Table 4.7: Oil demand by main petroleum product in OPEC Members (1,000 b/d)
dt_opec_oil_dm_product <- as.data.table(my_list[[24]])

# 25	Table 4.8: World oil demand by country (1,000 b/d)
dt_world_oil_dm_country <- as.data.table(my_list[[25]])

# 26	Table 4.9: World oil demand by main petroleum product and region (1,000 b/d)
dt_world_dm_product_region <- as.data.table(my_list[[26]])

# 27	Table 5.1: OPEC Members' crude oil exports by destination (1,000 b/d)
dt_opec_oil_exp_destination <- as.data.table(my_list[[27]])

# 28	Table 5.10: World imports of crude oil and petroleum products by country (1,000 b/d)
dt_world_imp_products_country <- as.data.table(my_list[[28]])

# 29	Table 5.2: OPEC Members' petroleum product exports by destination (1,000 b/d)
dt_opec_product_exp_destination <- as.data.table(my_list[[29]])

# 30	Table 5.3: World crude oil exports by country (1,000 b/d)
dt_world_oil_exp_country <- as.data.table(my_list[[30]])

# 31	Table 5.4: World exports of petroleum products by country (1,000 b/d)
dt_world_product_exp_country <- as.data.table(my_list[[31]])

# 32	Table 5.5: World exports of petroleum products by main petroleum product and region (1,000 b/d)
dt_world_product_exp_product_region <- as.data.table(my_list[[32]])

# 33	Table 5.6: World exports of crude oil and petroleum products by country (1,000 b/d)
dt_world_exp_oil_prod_country <- as.data.table(my_list[[33]])

# 34	Table 5.7: World imports of crude oil by country (1,000 b/d)
dt_world_imp_oil_country <- as.data.table(my_list[[34]]) 

# 35	Table 5.8: World imports of petroleum products by country (1,000 b/d)
dt_world_imp_products_country <- as.data.table(my_list[[35]])

# 36	Table 5.9: World imports of petroleum products by main petroleum product and region (1,000 b/d)
dt_world_product_imp_region <- as.data.table(my_list[[36]])

# 37	Table 6.1: World tanker fleet by year of build and categories (1,000 dwt)
dt_world_tanker_year_category <- as.data.table(my_list[[37]])

# 38	Table 6.2: World LPG carrier fleet by size (1,000 cu m)
dt_world_lpg_fleet <- as.data.table(my_list[[38]])

# 39	Table 6.3: World combined carrier fleet by size (1,000 dwt)
dt_world_fleet_size <- as.data.table(my_list[[39]])

# 40	Table 6.4: Average spot freight rates by vessel category (% of Worldscale)
dt_freight_category <- as.data.table(my_list[[40]])

# 41	Table 6.5: Dirty tanker spot freight rates (% of Worldscale and $/t)
dt_dirty_tanker_freight_rate <- as.data.table(my_list[[41]])

# 42	Table 6.6: Clean tanker spot freight rates (% of Worldscale and $/t)
dt_clean_tanker_freight_rate <- as.data.table(my_list[[42]])

# 43	Table 7.1: OPEC Reference Basket (ORB) and corresponding components spot prices ($/b)
dt_opec_price_co <- as.data.table(my_list[[43]])

# 44	Table 7.2: Selected spot crude oil prices ($/b)
dt_selected_spot_prices <- as.data.table(my_list[[44]])

# 45	Table 7.3: ICE Brent, NYMEX WTI and DME Oman annual average of the  1st, 6th and 12th forward months ($/b)
dt_brent_wti <- as.data.table(my_list[[45]])

# 46	Table 7.4: OPEC Reference Basket in nominal and real terms ($/b)
dt_opec_price_nominal_real <- as.data.table(my_list[[46]])

# 47	Table 7.5: Annual average of premium factors for selected OPEC Reference Basket components ($/b)
dt_annual_avg_selected_opec_price <- as.data.table(my_list[[47]])

# 48	Table 7.6: Spot prices of petroleum products in major markets  ($/b)
dt_srices_products <- as.data.table(my_list[[48]])

# 49	Table 7.7: Retail prices of petroleum products in OPEC Members  (units of national currency/b)
dt_retail_price_product_opec <- as.data.table(my_list[[49]])

# 50	Table 7.8: Crack spread in major markets ($/b)
dt_crack_spread <- as.data.table(my_list[[50]])

# 51	Table 8.1: Composite barrel and its components in major OECD oil consuming countries ($/b)
dt_composite_barrel_oecd <- as.data.table(my_list[[51]])

# 52	Table 8.2: Tax to CIF crude oil price ratio in major OECD oil consuming countries (ratio)
dt_tax_cif_oecd <- as.data.table(my_list[[52]])

# 53	Table 8.3: Euro Big 4 household energy prices, 2017 (USD/toe NCV)
dt_euro_big4_household <- as.data.table(my_list[[53]])

# 54	Table 9.1: World proven natural gas reserves by country (bn standard cu m)
dt_proven_ngas_reserves <- as.data.table(my_list[[54]])

# 55	Table 9.2: Natural gas marketed production in OPEC Members (million standard cu m)
dt_ngas_marketed_prod_opec <- as.data.table(my_list[[55]])

# 56	Table 9.3: World marketed production of natural gas by country (million standard cu m)
dt_world_ngas_marketed_prod <- as.data.table(my_list[[56]])

# 57	Table 9.4: World natural gas exports by country (m standard cu m)
dt_world_ngas_exports_country <- as.data.table(my_list[[57]])

# 58	Table 9.5: World natural gas imports by country (m standard cu m)
dt_world_ngas_imports_country <- as.data.table(my_list[[58]])

# 59	Table 9.6: World  natural gas demand by country (m standard cu m)
dt_world_ngas_demand_country <- as.data.table(my_list[[59]])

# 60	Table 9.7: World LNG carrier fleet by size and type (1,000 cu m)
dt_world_lng_fleet_size_type <- as.data.table(my_list[[60]])



# Plotting
# Plot 1
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

# Plot 2
dt_rigs[country %in% c('USA', 'Russia')] %>% 
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

