## Make changes log


# Load packages
pacman::p_load(dplyr, readxl, openxlsx, tidyverse)





# Step 1: 

# 










# Source function 
source(paste0("C:/Users/HD95LP/OneDrive - Aalborg Universitet/visual_studio_code/2023_update_MIDAS_field_guide/scripts/select_rows_and_create_excel_with_changes.R"))

WD = "C:/Users/HD95LP/OneDrive - Aalborg Universitet/visual_studio_code/2023_update_MIDAS_field_guide"
# Update december: 

setwd(WD)

"data/changes_tab_file/2024_02_27_get_data_from_old_tax_names.xlsx"

# Keep old decriptions when updating to MIDAS 5.2
make_new_excel(
  input_excel_file = paste0(WD, "/data/excel_from_Marta/20240223_microbe_details.xlsx"),
  input_change_file = paste0(WD, "/data/changes_tab_file/2024_02_27_get_data_from_old_tax_names_and_Opimibacter.xlsx"),
  output_excel_file <-  paste0(WD, "/output/edits_in_excel/", format(Sys.Date(), "%Y%m%d"), "_microbe_details_update_1.xlsx")
)


make_new_excel(
  input_excel_file = paste0(WD, "/data/excel_from_Marta/microbe_details 14_11-2023_add_names.xlsx"),
  input_change_file = paste0(WD, "/data/changes_tab_file/2023_12_changes.xlsx"),
  output_excel_file <-  paste0(WD, "/output/edits_in_excel/", format(Sys.Date(), "%Y%m%d"), "_microbe_details_update_v4.xlsx")
)


