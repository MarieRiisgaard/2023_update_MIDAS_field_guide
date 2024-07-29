## Import change log and format data in format for change file



get_old_info <- function(
    WD = WD, 
    # Input: Excel file WITH ALL names from new database
    input_excel_file = paste0(WD, "/output/include_all_names/MIDAS_52_excel_web_format.xlsx"),
    # Input: Change log with new names
    input_MIDAS_change_log = paste0(WD, "/data/MIDAS_change_log/MIDAS_changelog_51.txt"),
    # Output: Tab changes file
    output_excel_file = paste0(WD, "/output/changes_tab_file/20240227_MIDASv52_data_from_old_tax_names_test1.xlsx"),
    old_midas_version = "MiDAS 5.1"
  ){

# Output: changes file  
  # Specify the path to your Excel file
  #excel_file <- paste0(WD, "/data/excel_from_Marta/microbe_details 14_11-2023.xlsx")
  #rm(excel_file)
  
  
################################

d = vroom::vroom(input_MIDAS_change_log)
  
# Seperate the change log into speices and genera  
  ## New SPECIES NAMES <- Get names from before with ID that link to new name
  new_species_names <- d %>% 
    filter(str_detect(Before, "s:")) %>% 
    mutate(Before = str_remove(Before, ".*s:"), 
           After = str_remove(After, ".*s:"), 
           Before = if_else(!str_detect(Before, "midas"), str_replace_all(Before, "_", " "), Before),
           After = if_else(!str_detect(After, "midas"), str_replace_all(After, "_", " "), After),
           Before = str_remove(Before, ";"),
           After = str_remove(After, ";"),
           Before = str_replace(Before, "Ca ", "Candidatus "),
           After = str_replace(After, "Ca ", "Candidatus "),
    ) %>% 
    rename("name" = "After")

  # New Genus NAMES <- Get names from before with ID that link to new name
  new_genus_names <- d %>% 
    filter(str_detect(Before, "g:")) %>% 
    mutate(Before = str_replace(Before, ",.*$", ""), 
           After = str_replace(After, ",.*$", ""), 
           Before = if_else(!str_detect(Before, "midas"), str_replace_all(Before, "_", " "), Before),
           After = if_else(!str_detect(After, "midas"), str_replace_all(After, "_", " "), After),
           Before = str_remove(Before, "g:"),
           After = str_remove(After, "g:"),
    ) %>% 
    rename("name" = "After") 
  
new_tax_names <- rbind(new_species_names, new_genus_names) %>% 
  mutate(Reference = str_replace(Reference, "=", ":"), 
         Reference = paste0("[", Reference, "]")
         )



# Read all sheets into a list of data frames
all_sheets <- lapply(excel_sheets(paste0(input_excel_file)), function(sheet) {
  read_excel(input_excel_file, sheet = sheet)
})


# First step is to get the info from the old names
  # NOTE: for some old names no info is availible

# Define the vector of IDs you want to filter (filter all the old names with info)
old_names <- unlist(new_tax_names$Before) 

# This function returns the sheets with new names and all the old info. 
process_sheet <- function(sheet_index = 2, id_vector, all_sheets) {
  
  exist <- all_sheets[[sheet_index]] %>% 
    filter(name %in% old_names) %>% #print(n = 100)
    rename("Before" = "name") %>%  # RENAME the name column containing old names 
    left_join(new_tax_names, by = "Before")  
  
  # Only for the sheet containing the alternative names column: 
  if(colnames(exist)[2] == "alternative names"){
    exist <-
      exist %>% 
      group_by(name) %>% 
      mutate("alternative names" = 
               case_when(
                 !is.na(`alternative names`) & is.na(Reference) ~ 
                   paste0("Previously ", paste0(Before, collapse = "/")," (",old_midas_version,")", ", ",`alternative names`),
                 is.na(`alternative names`) & is.na(Reference) ~ 
                   paste0("Previously ", paste0(Before, collapse = "/")," (",old_midas_version,")"),
                 !is.na(`alternative names`) & !is.na(Reference) ~ 
                   paste0("Previously ", paste0(Before, collapse = "/")," (",old_midas_version,")", Reference, ", ",`alternative names`),
                 is.na(`alternative names`) & !is.na(Reference) ~ 
                   paste0("Previously ", paste0(Before, collapse = "/")," (",old_midas_version,")", Reference),
                 .default = `alternative names`), 
             "alternative names" = str_replace_all(`alternative names`, " Ca ", " *Ca* ")
      )
  }
  
  
  exist <- 
    exist %>% 
    select(-Before, -Change, -Reference) %>%
    distinct() %>% 
    relocate(name, .before = 1) 
  
  return(exist)
  
}

# Apply the function to all sheets
S2 <- process_sheet(2, old_names, all_sheets)
S3 <- process_sheet(3, old_names, all_sheets)
S4 <-process_sheet(4, old_names, all_sheets)
S5 <-process_sheet(5, old_names, all_sheets)

old_info_sheets <- list(
  NA, 
  S2, S3, S4, S5
)

# Second step is to track the name changes of those not where the OLD name not had any info

process_sheet2 <- function(sheet_index = 2, id_vector, all_sheets) {
  
  exist <- all_sheets[[sheet_index]] %>% 
    filter(!name %in% unique(old_info_sheets[[sheet_index]]$name)) %>% #REMOVE the names with old info
    filter(name %in% unlist(unique(new_tax_names$name))) %>% 
    #rename("Before" = "name") %>%  # RENAME the name column containing old names 
    left_join(new_tax_names)  
  
  # Only for the sheet containing the alternative names column: 
  if(colnames(exist)[2] == "alternative names"){
    exist <-
      exist %>%
      group_by(name) %>% 
      mutate("alternative names" = 
               case_when(
                 !is.na(`alternative names`) & is.na(Reference) ~ 
                   paste0("Previously ", paste0(Before, collapse = "/")," (",old_midas_version,")", ", ",`alternative names`),
                 is.na(`alternative names`) & is.na(Reference) ~ 
                   paste0("Previously ", paste0(Before, collapse = "/")," (",old_midas_version,")"),
                 !is.na(`alternative names`) & !is.na(Reference) ~ 
                   paste0("Previously ", paste0(Before, collapse = "/")," (",old_midas_version,")", Reference, ", ",`alternative names`),
                 is.na(`alternative names`) & !is.na(Reference) ~ 
                   paste0("Previously ", paste0(Before, collapse = "/")," (",old_midas_version,")", Reference),
                 .default = `alternative names`), 
             "alternative names" = str_replace_all(`alternative names`, " Ca ", " *Ca* ")
      ) 
  }
  
  
  exist <- 
    exist %>% 
    select(-Before, -Change, -Reference) %>%
    distinct() %>% 
    relocate(name, .before = 1) %>% 
    rbind(old_info_sheets[[sheet_index]], .) %>% 
    distinct() %>% 
    rename("genus_name" = "name") %>%
    pivot_longer(cols = -genus_name, names_to = "column_name_for_change", values_to = "new_value") %>% 
    filter(!is.na(new_value))
  
  return(exist)
  
}


# Apply the function to all sheets
S2 <- process_sheet2(2, old_names, all_sheets)
S3 <- process_sheet2(3, old_names, all_sheets)
S4 <-process_sheet2(4, old_names, all_sheets)
S5 <-process_sheet2(5, old_names, all_sheets)

final <- rbind(S2,S3,S4,S5) %>% 
  filter(!new_value == "In situ=na||Other=na")


# Create a new Excel workbook
output_wb <- createWorkbook()

# Add a worksheet to the new Excel workbook
addWorksheet(output_wb, sheetName = "change_file")

# Write the filtered data to the new worksheet
writeData(output_wb, x = final, sheet = "change_file")

openxlsx::saveWorkbook(output_wb, output_excel_file)

print(paste0("Output: New change log file. Path: ", output_excel_file))
}

############################################################################



# 
# # Output is old `names<-.POSIXlt`()# Output is old names in the name column and all 
# 
# 
# # Change to new name 
# change_to_new_name <- function(sheet = s, names_df = new_tax_names){
#   x <- 
#     sheet %>% 
#     rename("Before" = "name") %>%
#     left_join({{names_df}} %>% 
#                 select(Before, name), by = "Before")
#   if(colnames(x)[2] == "alternative names"){
#     x <- x %>% 
#       mutate("alternative names" = 
#                if_else(!is.na(`alternative names`), 
#                        paste0("Previously ", Before," (",old_midas_version,")", ", ",`alternative names`), 
#                        `alternative names`), 
#              "alternative names" = str_replace_all(`alternative names`, " Ca ", " *Ca* "))
#   }
#   x <- x %>% 
#     select(-Before) %>%
#     relocate(name, .before = 1) %>% 
#     rename("genus_name" = "name") %>%
#     pivot_longer(cols = -genus_name, names_to = "column_name_for_change", values_to = "new_value")
#   return(x)
# }
# 
# 
# 
# 
# # Change to new name version 2
# change_to_new_name <- function(sheet = S2, 
#                                names_df = new_tax_names){
#   x <- 
#     sheet %>% 
#     left_join({{names_df}} %>% 
#                 select(Before, name), by = "name") %>% 
#     mutate("alternative names" = 
#                case_when(
#                  !is.na(`alternative names`) ~ paste0("Previously ", Before," (MiDAS 5.1)", ", ",`alternative names`),
#                  is.na(`alternative names`) ~ paste0("Previously ", Before," (MiDAS 5.1)"),
#                  .default = `alternative names`), 
#              "alternative names" = str_replace_all(`alternative names`, " Ca ", " *Ca* ")
#              ) %>% 
#     select(-Before) %>%
#     #relocate(name, .before = 1) %>% 
#     rename("genus_name" = "name") %>%
#     pivot_longer(cols = -genus_name, names_to = "column_name_for_change", values_to = "new_value")
#   return(x)
# }
# 
# change_file_new_names <- rbind(
# change_to_new_name(sheet = S2, names_df = new_tax_names)#,
# #change_to_new_name(sheet = S3, names_df = new_tax_names),
# #change_to_new_name(sheet = S4, names_df = new_tax_names),
# #change_to_new_name(sheet = S5, names_df = new_tax_names)
# )
# 
# 
# 
# # Create a new Excel workbook
# output_wb <- createWorkbook()
# 
# # Add a worksheet to the new Excel workbook
# addWorksheet(output_wb, sheetName = "change_file")
# 
# # Write the filtered data to the new worksheet
# writeData(output_wb, x = change_file_new_names, sheet = "change_file")
# 
# openxlsx::saveWorkbook(output_wb, output_excel_file)
# 
# print(paste0("Output: New change log file. Path: ", output_excel_file))
# }
# 
# 
# 
# 
# 
# 
# 
# 
# 
# 
# 
# 
