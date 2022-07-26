library(tidyverse)
library(magrittr)
library(openxlsx)
library(readxl)
library(writexl)
library(reshape2)
library(skimr)


# reading

# Main data
data <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Mustard Constrain/For R/Mustards - Commercial Team_07252022.xlsx")
colnames(data)[16] <- "comp_no"

data %>% 
  dplyr::filter(comp_no %in% c(6866, 6868)) -> data

# Super Customer
super_cust <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Mustard Constrain/For R/Super Customer.xlsx")
names(super_cust) <- stringr::str_replace_all(names(super_cust), c(" " = "_"))

# Sales Manager
sales_manager <- read_excel("C:/Users/slee/OneDrive - Ventura Foods/Ventura Work/SCE/Project/FY 23/Mustard Constrain/For R/Sales Manager.xlsx")
names(sales_manager) <- stringr::str_replace_all(names(sales_manager), c(" " = "_"))


# merge Super Customer
super_cust %>% 
  dplyr::select(ref, Super_Customer_No, Super_Customer_Name) -> super_cust_2

dplyr::left_join(data, super_cust_2, by = "ref") -> main_data

# merge Sales Manager
sales_manager %>%
  dplyr::select(ref, Sales_Manager_No, Sales_Manager_Name, Sales_Territory_No, Sales_Territory_Name) -> sales_manager_2

dplyr::left_join(main_data, sales_manager_2, by = "ref") -> main_data


# back to original col names





# export to excel
writexl::write_xlsx(main_data, "mustard_data.xlsx")
