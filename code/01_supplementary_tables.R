#########################################################################
# Name of file - 01_supplementary_tables.R
#
# Type - Reproducible Analytical Pipeline (RAP)
# Written/run on - RStudio Desktop
# Version of R - 4.2.2
#
# Description - Imports the needed files, cleans the data, creates 
# the supplementary tables and exports them to an Excel file.

#########################################################################


### 0 - Setup ----

# Run setup script which loads all required packages and functions and 
# executes the config.R script.

source(here::here("code", "00_setup.R"))

### 1 - Import QA file ----

# import QA worksheet outputted from SAS
cleaned_qa_list <- get_qa_data(qa.path)

# check QA workbook includes all required variables
# if output is not character(0), check if 00_setup script includes 
# all required variables
setdiff(names(cleaned_qa_list), lookup_df$vname)

### 2 - Import CI files ----

# get paths of CI Excel files outputted from SAS
files <- list.files(path = ci.path,
                    pattern = paste0("Subgroup CIs\\.xlsx$"),
                    full.names = TRUE,
                    recursive = TRUE,
                    ignore.case = TRUE)
files

# view files that have CI sheets but are not included in lookup_df
# update lookup_df where appropriate
files[!grepl(paste(lookup_df$vname,
                            collapse = "|"), 
                      tolower(files))]

# select files that include the data needed for final output 
# (i.e., file names that include variables mentioned in the 
# lookup table)
files <- files[grepl(paste(lookup_df$vname,
                           collapse = "|"), 
                     tolower(files))]

# check all required files are present in CI folder
length(files) == length(lookup_df$vname)

# import CI files and clean for export
ci_tables_list <- get_ci_data(files[30:31])

# reorder list of CI data
ci_tables_list <- ci_tables_list[lookup_df$vname]

### 3 - Export ----

# export as Excel file
# A few warning messages are displayed when executing the following code
# (e.g., tab_titles have been cleaned automatically +
# there are no notes in your table)
# This is nothing to be concerned about as the notes are displayed right 
# after the word they refer to rather than in a separate column
export_tables(ci_tables_list)
