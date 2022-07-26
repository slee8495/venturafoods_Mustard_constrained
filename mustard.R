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
colnames(main_data)[1]	<- "ref"
colnames(main_data)[2]	<- "comp ref"
colnames(main_data)[3]	<- "SKU Status"
colnames(main_data)[4]	<- "Category"
colnames(main_data)[5]	<- "Label"
colnames(main_data)[6]	<- "where used count per loc(Active only)"
colnames(main_data)[7]	<- "where used all loc (Active only)"
colnames(main_data)[8]	<- "Business Unit"
colnames(main_data)[9]	<- "Level"
colnames(main_data)[10]	<- "Parent Item Number"
colnames(main_data)[11]	<- "Parent Description"
colnames(main_data)[12]	<- "UoM"
colnames(main_data)[13]	<- "Net lbs."
colnames(main_data)[14]	<- "FG On hand"
colnames(main_data)[15]	<- "FG WOH"
colnames(main_data)[16]	<- "Comp#/labor code"
colnames(main_data)[17]	<- "Comp description"
colnames(main_data)[18]	<- "Commodity Class"
colnames(main_data)[19]	<- "Um"
colnames(main_data)[20]	<- "Lead time"
colnames(main_data)[21]	<- "RM on hand"
colnames(main_data)[22]	<- "RM overall weeks on hand"
colnames(main_data)[23]	<- "Stocking Type"
colnames(main_data)[24]	<- "Percent Scrap"
colnames(main_data)[25]	<- "Quantity Per"
colnames(main_data)[26]	<- "Quantity w/ Scrap"
colnames(main_data)[27]	<- "Standard cost"
colnames(main_data)[28]	<- "next 28 days open order"
colnames(main_data)[29]	<- "Jul fcst"
colnames(main_data)[30]	<- "Aug fcst"
colnames(main_data)[31]	<- "Sep fcst"
colnames(main_data)[32]	<- "Oct fcst"
colnames(main_data)[33]	<- "Nov fcst"
colnames(main_data)[34]	<- "Dec fcst"
colnames(main_data)[35]	<- "RM Jul 2022"
colnames(main_data)[36]	<- "RM Aug 2022"
colnames(main_data)[37]	<- "RM Sep 2022"
colnames(main_data)[38]	<- "RM Oct 2022"
colnames(main_data)[39]	<- "RM Nov 2022"
colnames(main_data)[40]	<- "RM Dec 2022"






# export to excel
writexl::write_xlsx(main_data, "mustard_data.xlsx")
