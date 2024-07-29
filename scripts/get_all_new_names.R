# Fix the missing names problem 


missing_names <- function(
    WD = "C:/Users/HD95LP/OneDrive - Aalborg Universitet/visual_studio_code/2023_update_MIDAS_field_guide",
    # Input:
    ## Newest excel from Marta
    input_excel_file = paste0(WD, "/data/excel_from_Marta/20240223_microbe_details.xlsx"),
    ## Newest version Midas database 
    midas_db = paste0(WD,"/../data_microbial_groups/MiDAS/MIDAS_52_taxa.txt"), 
    # Output
    ## Excel with new names  
    output_excel_file = paste0(WD, "/output/include_all_names/MIDAS_52_excel_web_format.xlsx")
    ){


# Get all Genus and Species in the database
MIDAS_db <- 
  read_csv2(file = midas_db, 
              col_names = F) %>% 
  select(X6, X7) %>% 
  rename("Species" = 2,
         "Genus" = 1) %>% 
  mutate(
    Species = str_remove(Species, "s__"),
    Species = if_else(str_detect(Species, "midas"), Species, str_replace_all(Species, "_", " ")),
    #Species = if_else(str_detect(Species, "Ca "), str_replace_all(Species, "Ca", "Candidatus"), Species),
    Genus = str_remove(Genus, "g__"),
    Genus = if_else(str_detect(Genus, "midas"), Genus, str_replace_all(Genus, "_", " ")),
  ) %>% 
  pivot_longer(cols = c(Genus, Species), names_to = "tax_level", values_to = "name") %>% 
  #filter(tax_level == "Genus") %>% 
  distinct(name)
  

# Read all sheets into a list of data frames
all_sheets <- lapply(excel_sheets(paste0(input_excel_file)), function(sheet) {
  read_excel(input_excel_file, sheet = sheet)
})

process_sheet <- function(
    sheet_index = 1, 
    all_sheets = all_sheets) {

  # Find colnames per sheet
  tbl_colnames <- colnames(all_sheets[[sheet_index]])
  

  # Make tibble for per sheet
  if (sheet_index %in% c(1)) {
    tbl <- all_sheets[[sheet_index]]
  } else if (sheet_index %in% c(2, 3)) {
    # Find missing names per sheet
    name <- anti_join(MIDAS_db, all_sheets[[sheet_index]] %>% select(name))
    # Add all names as new row and fill other coloums with NA/empty (for sheet 2 and 3)
    tbl <- 
      as_tibble(matrix(nrow = 0, ncol = length(tbl_colnames)), 
                .name_repair = ~ tbl_colnames) %>%
      mutate(name = as.character()) %>% 
      add_row(name = unlist(name))
    # RBIND all the data from previosly
    tbl <- rbind(all_sheets[[sheet_index]], tbl)
    
    ## Fill In situ=na||Other=na (for sheet 4 and 5) 
  } else if (sheet_index %in% c(4, 5)) {
    # Find missing names per sheet
    name <- anti_join(MIDAS_db, all_sheets[[sheet_index]] %>% select(name))
    tbl <- 
      as_tibble(matrix(nrow = 0, ncol = length(tbl_colnames)), 
                .name_repair = ~ tbl_colnames) %>%
      mutate(name = as.character()) %>% 
      add_row(name = unlist(name))
    tbl[, -1] <- "In situ=na||Other=na"
    # RBIND all the data from previosly
    tbl <- rbind(all_sheets[[sheet_index]], tbl)
  } else {
    print("mistake")
  }

# Combine the two dataframes
return(tbl)
}

# Specify the path to the output Excel file
sheet_names <- c( "refs","basic","text fields","metabolism","cell properties")

# Create a new Excel workbook
output_wb <- createWorkbook()

# Process sheets and add to new Excel workbook
for (i in 1:5) {
  S <- process_sheet(i, all_sheets)
  addWorksheet(output_wb, sheetName = sheet_names[i])
  writeData(output_wb, sheet = sheet_names[i], S)
}

# Save the new Excel workbook
saveWorkbook(output_wb, output_excel_file)


}




# # Apply the function to all sheets
# S1 <- all_sheets[[1]]
# S2 <- process_sheet(2, all_sheets)
# S3 <- process_sheet(3, all_sheets)
# S4 <-process_sheet(4, all_sheets)
# S5 <-process_sheet(5, all_sheets)
# 
# # Add a worksheet to the new Excel workbook
# addWorksheet(output_wb, sheetName = sheet_names[1])
# addWorksheet(output_wb, sheetName = sheet_names[2])
# addWorksheet(output_wb, sheetName = sheet_names[3])
# addWorksheet(output_wb, sheetName = sheet_names[4])
# addWorksheet(output_wb, sheetName = sheet_names[5])
# 
# 
# # Write the filtered data to the new worksheet
# writeData(output_wb, sheet = sheet_names[1], S1)
# writeData(output_wb, sheet = sheet_names[2], S2)
# 










          