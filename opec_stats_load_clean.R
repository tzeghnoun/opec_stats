library(data.table)
library(tidyverse)
library(ggrepel)
library(ggdark)
library(readxl)

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

df_values %>% 
  ggplot() +
  aes(year, petroleum_exp/1000, color = country) +
  geom_line() +
  scale_x_continuous(limits = c(1960, 2018), breaks = seq(1960, 2018, 3)) +
  labs(
    title = "Values of petroleum exports (10^9 $)",
    x = "",
    y = ""
  ) +
  dark_mode()
