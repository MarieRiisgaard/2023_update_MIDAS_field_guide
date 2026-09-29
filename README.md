# Update details in MiDAS field guide

This repo contains info about how to upload changes to the descriptions of genera and speices in the MiDAS field guide


## Description of scripts


### Update the Field Guide 

[`update_midas_database_simplified.Rmd`] prepares an Excel workbook containing manual updates to the MiDAS Field Guide for upload to the website.

**Inputs and setup**

Set the working directory (`WD`) and specify:
- `input_excel_file_from_Marta`: the latest database export in the website’s Excel format.
- `input_change_file_manual`: an Excel workbook containing the manual changes. The file should contain 3 columns: `name`, `column_name_for_change`, and `new_value`. 
- `output_excel_file`: the path for the generated update workbook that then can uploaded to the website.

**Generate and check updates**

The script calls `make_new_excel()` from `scripts/select_rows_and_create_excel_with_changes.R` to apply the changes. It then reports:
- Invalid formatting in the **cell properties** and **metabolism** sheets, where values should follow `In situ=...||Other=...`, using `na`, `pos`, `neg`, or `var` and optional reference tags.
- Malformed references in the **general information** and **descriptions** sheets, including incorrectly formatted `[PMID:12345678]` or `[REPLACE1]` tags, references outside brackets, and unbalanced square brackets.

Review and correct the reported issues before uploading the workbook; the checks do not automatically fix them.

**Extract existing information for editing**

A separate workflow reads the current database export and converts information from sheets 2–5 into the columns `name`, `column_name_for_change`, and `new_value`. 
- These can be exported to one Excel workbook or separate workbooks for selected taxa. Adjust the filters to select the taxa and fields needed; the supplied code currently extracts **FISH probes**, with the taxon filter commented out.
- The common practice is to supply these files to people that then will update the information (see PPP in teams for guidelines) and then all these will be merged to a single file `input_change_file_manual` that can be reformatted in the above workflow. 



### Update the Field Guide for a new MiDAS taxonomy version

[`update_midas_database.Rmd`] This extended workflow adds support for new taxa and taxonomic name changes to the manual-update workflow described above.

**Additional inputs**

- `midas_db_tax_file`: the taxonomy file for the new MiDAS version.
- `input_MIDAS_change_log_file`: a tab-separated name-change log with the columns `Before`, `After`, `Change`, and `Reference (if any)`.

**Additional functions**

1. **Add missing taxa:** `missing_names()` from `get_all_new_names.R` compares the database export against the taxonomy file and generates a website-formatted workbook containing the additional names.
2. **Transfer existing information to new names:** `get_old_info()` from `update_names_but_retain_old_data.R` uses the change log to generate a change file carrying information from old taxonomic names to their replacements. Set `old_midas_version` to the relevant previous version.
3. **Merge transferred information with manual edits:** Combines the two change files, prioritising non-missing manual values over transferred values for matching entries, and saves the merged changes as an intermediate workbook.

The resulting workbook and change file are then passed to `make_new_excel()` as in the simplified workflow. The name-addition and information-transfer steps can be skipped when unnecessary.

**Before running:** Select the appropriate input assignments in the final update block. Keep `output_excel_file_all_names` pointing to the expanded workbook when adding taxa, and set `input_change_file_x` to either the manual or merged change file. The supplied code currently overwrites the expanded-workbook path with the original export and selects the merged change file.



