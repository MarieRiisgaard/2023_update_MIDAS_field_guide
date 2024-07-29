##########################################
### Update data in MIDAS field guide  ####
##########################################

make_new_excel <- function(
  input_excel_file = paste0(WD, "/data/excel_from_Marta/microbe_details 14_11-2023.xlsx"),
  input_change_file = paste0(WD, "/data/changes_tab_file/2023_12_changes.xlsx") , 
  output_excel_file = paste0(WD, "/output/edits_in_excel/", format(Sys.Date(), "%Y%m%d"), "_microbe_details_update.xlsx")
  ){

# Specify the path to the output Excel file
sheet_names <- c( "refs","basic","text fields","metabolism","cell properties")
change_file <- read_excel(input_change_file)
genus_names <- unique(change_file$genus_name)


# Create a new Excel workbook
output_wb <- createWorkbook()

# Loop through each sheet, filter rows, and add to the new workbook
for (i in seq_along(sheet_names)) {
  sheet_name <- sheet_names[i]
  
  # Read the sheet into a data frame
  sheet_data <- read_excel(input_excel_file, sheet = sheet_name)
  
  # Filter rows based on the "name" column for sheets 2 to 5
  if (i >= 2 && i <= 5) {
    ## Filter rows based on genus name
    filtered_data <- sheet_data %>%
      filter(name %in% genus_names)
  
    for (i in seq(nrow(change_file))) {
      name_to_update <- change_file$genus_name[i]
      colname_to_update <- change_file$column_name_for_change[i]
      new_value <- change_file$new_value[i]
      
      # Check if the name exists in 'filtered_data' and the column exists in 'filtered_data'
      if (name_to_update %in% filtered_data$name && colname_to_update %in% colnames(filtered_data)) {
        filtered_data[filtered_data$name == name_to_update, colname_to_update] <- new_value
      }
    }
    
  } else {
    # For other sheets (1 and 5 in this example), keep all rows
    filtered_data <- sheet_data
  }
  
  
  
  # Add a worksheet to the new Excel workbook
  addWorksheet(output_wb, sheetName = sheet_name)
  
  # Write the filtered data to the new worksheet
  writeData(output_wb, sheet = sheet_name, filtered_data)
}

# Save the new Excel workbook
saveWorkbook(output_wb, output_excel_file)

}






