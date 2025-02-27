# Event Study: Analysis of Abnormal Returns on Green Bond Announcements
# by US Corporates from 2018-2022

# 1. LOADING PACKAGES
required_packages <- c("tidyverse", "readxl", "ggplot2", "lmtest", "sandwich", "car", "boot", "zoo")
new_packages <- required_packages[!(required_packages %in% installed.packages()[,"Package"])]
if(length(new_packages)) install.packages(new_packages)

library(tidyverse)
library(readxl)
library(ggplot2)
library(lmtest)
library(sandwich)
library(car)
library(boot)
library(zoo)

# 2. READ FILES
# Define file paths
events_path <- "C:/Users/DELL/Downloads/1. Events_in_ET Kopie.xlsx"
index_path <- "C:/Users/DELL/Downloads/2. Index_Data_in_ET Kopie.xlsx"
stock_path <- "C:/Users/DELL/Downloads/3. Stock_Data_in_ET Kopie.xlsx"
regression_path <- "C:/Users/DELL/Downloads/4. Regression_Data Kopie.xlsx"

# Read Regression Data
industry_data <- read_excel(regression_path, sheet = 1, col_names = FALSE) %>%
  rename(ISIN = "...1", Industry = "...2")

total_assets_data <- read_excel(regression_path, sheet = 2, col_names = FALSE) %>%
  rename(ISIN = "...1", `Total Assets` = "...2")

market_value_data <- read_excel(regression_path, sheet = 3, col_names = FALSE) %>%
  rename(ISIN = "...1", `Market Value` = "...2")

# Merge Regression Data
regression_data <- industry_data %>%
  left_join(total_assets_data, by = "ISIN") %>%
  left_join(market_value_data, by = "ISIN")

# 3. CLEAN AND PREPARE THE DATA
events <- read_excel(events_path) %>%
  select(`Unternehmens ISIN`, `Bond ISIN`, `Announc. in ET`, `ESG Bond Type`, `Issued Amount (USD)`, `Yield`, `Maturity Date`, `Issue Date`) %>%
  mutate(
    `Announc. in ET` = as.numeric(`Announc. in ET`),
    `Unternehmens ISIN` = trimws(`Unternehmens ISIN`)
  ) %>%
  drop_na(`Unternehmens ISIN`, `Announc. in ET`) %>%
  arrange(`Unternehmens ISIN`, `Bond ISIN`, `Announc. in ET`)

index_data <- read_excel(index_path) %>%
  select(Eventtime, `IndexS&P500`) %>%
  mutate(
    Eventtime = as.numeric(Eventtime),
    `IndexS&P500` = as.numeric(`IndexS&P500`)
  ) %>%
  arrange(Eventtime) %>%
  mutate(Market_Returns = log(`IndexS&P500` / lag(`IndexS&P500`))) %>%
  mutate(Market_Returns = na.locf(Market_Returns, na.rm = FALSE)) %>%
  mutate(Market_Returns = ifelse(is.na(Market_Returns), 0, Market_Returns))

# 4. CALCULATE LOG RETURNS
stock_prices <- read_excel(stock_path) %>%
  mutate(`Isin/Eventtime` = as.numeric(`Isin/Eventtime`))

event_isins <- events %>% pull(`Unternehmens ISIN`) %>% unique()
valid_isins <- intersect(event_isins, colnames(stock_prices))

stock_returns <- stock_prices %>%
  select("Isin/Eventtime", all_of(valid_isins)) %>%
  pivot_longer(cols = -`Isin/Eventtime`, names_to = "ISIN", values_to = "Price") %>%
  mutate(
    Price = as.numeric(Price),
    Returns = log(Price / lag(Price))
  ) %>%
  drop_na() %>%
  arrange(ISIN, `Isin/Eventtime`)

# 5. ESTIMATION WINDOW (-220 to -20) & MARKET MODEL
estimation_data <- stock_returns %>%
  inner_join(events, by = c("ISIN" = "Unternehmens ISIN")) %>%
  rename(`Unternehmens ISIN` = ISIN) %>%  
  filter(`Isin/Eventtime` >= (`Announc. in ET` - 220) & `Isin/Eventtime` <= (`Announc. in ET` - 20)) %>%
  left_join(index_data, by = c("Isin/Eventtime" = "Eventtime")) %>%
  arrange(`Unternehmens ISIN`, `Bond ISIN`, `Announc. in ET`, `Isin/Eventtime`)

market_model_params <- estimation_data %>%
  group_by(`Unternehmens ISIN`) %>%
  summarise(
    alpha = coef(lm(Returns ~ Market_Returns, data = pick(everything())))[1],
    beta = coef(lm(Returns ~ Market_Returns, data = pick(everything())))[2],
    .groups = "drop"
  ) %>%
  rename(ISIN = `Unternehmens ISIN`)

# 6. CALCULATE ABNORMAL RETURNS & CAR
event_stock_data <- stock_returns %>%
  inner_join(events, by = c("ISIN" = "Unternehmens ISIN")) %>%
  left_join(market_model_params, by = "ISIN") %>%
  left_join(index_data, by = c("Isin/Eventtime" = "Eventtime")) %>%
  mutate(
    Expected_Return = alpha + beta * Market_Returns,
    Abnormal_Return = Returns - Expected_Return
  ) %>%
  arrange(`Unternehmens ISIN`, `Bond ISIN`, `Announc. in ET`, `Isin/Eventtime`)

CAR_data <- event_stock_data %>%
  filter(`Isin/Eventtime` >= (`Announc. in ET` - 10) & `Isin/Eventtime` <= (`Announc. in ET` + 10)) %>%
  group_by(`Unternehmens ISIN`, `Bond ISIN`, `Announc. in ET`) %>%
  summarise(
    CAR_Pre = sum(Abnormal_Return[(`Isin/Eventtime` >= (`Announc. in ET` - 10)) & (`Isin/Eventtime` <= (`Announc. in ET` - 4))], na.rm = TRUE),
    CAR_Event = sum(Abnormal_Return[(`Isin/Eventtime` >= (`Announc. in ET` - 3)) & (`Isin/Eventtime` <= (`Announc. in ET` + 3))], na.rm = TRUE),
    CAR_Post = sum(Abnormal_Return[(`Isin/Eventtime` >= (`Announc. in ET` + 4)) & (`Isin/Eventtime` <= (`Announc. in ET` + 10))], na.rm = TRUE),
    .groups = "drop"
  )

print(CAR_data)
