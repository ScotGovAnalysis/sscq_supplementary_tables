# SSCQ Supplementary Tables
This repository contains the Reproducible Analytical Pipeline (RAP) for the Scottish Surveys Core Questions (SSCQ) supplementary tables.

## Updating Required Information

The only two files that need to be updated before running the sscq_supplementary_tables RAP are the 00_setup.R and the config.R file. 

In the config.R file, values need to be assigned to the following objects:
1. `ci.path`: Path of the 'CI Sheets' folder to which SAS outputs the Subgroup CI files.
2. `qa.path`: File path of the 'SSCQ Subgroup Tables QA' file outputted from SAS.
3. `export.path`: Path of the folder to which the file with the supplementary tables will be saved.

The config.R file should be saved in the code folder.

## Running the RAP

To run the RAP, execute the 01_supplementary_tables.R file which automatically loads the config.R file, the 00_setup.R file and all functions in the functions folder.

## Licence

This repository is available under the [Open Government Licence v3.0](https://www.nationalarchives.gov.uk/doc/open-government-licence/version/3/).
