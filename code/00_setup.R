#########################################################################
# Name of file - 00_setup.R
#
# Type - Reproducible Analytical Pipeline (RAP)
# Written/run on - RStudio Desktop
# Version of R - 4.2.2
#
# Description - Sets up environment required for running the 
# SSCQ_supplementary_tables RAP.

#########################################################################

### 1 - Dates - TO UPDATE ----

# Publication date
pub_date <- "26 March 2024"

# Year the published data is from
sscq_year <- "2022"

### 2 - File name of supplementary tables - TO UPDATE ----

# name of final file
fname <- "supplementary_tables.xlsx"

### 3 - Cover sheet - TO UPDATE ----

# titles and accompanying texts for cover sheet
cover_list <-  list(
  "Date of Publication" = paste0("Date published: ", pub_date),
  "Overview" = c("These tables provide the latest results from the Scottish Surveys Core Questions dataset, covering the collection period for 2022.",
                     "They consist of a full analysis of each topic across all possible social and geographic breakdowns."),
  "Grouping of Variables" = c("Certain ethnic groups, religions and countries of birth were grouped to sufficient numbers of responses to enable statistical analysis.",
                                  "More information on this can be found in the technical report."),
  "Confidence Intervals" = c("Also included in these tables are the 95% confidence intervals on each estimate.",
                                 "Where confidence intervals do not overlap, users may assume that there is a statistically significant difference between the two groups."),
  "Table Structure" = c("Most information is transposed in tables across different sections, providing different options for comparisons.",
                            "All tables break down percentages in rows.",
                            "‘Refused’ and ‘don’t know’ responses are excluded, so row totals may not add to 100%, and numbers of adults and sample may not add to the Scotland total for each cross-variable."),
  "Weighting" = c("Percentage estimates are based on weighted analysis of the SSCQ data.",
                      "In a dataset with full reponse, individual respondents would have a weight of 1, but due to the weighting procedures to account for non-response and sampling, individual respondents can have any value positive weight.",
                      "It is therefore not possible to calculate individual sample numbers in each respondent grouping by combining weighted estimates with unweighted sample size (N)."),
  "Disclosure Control" = c("All estimates based on five or fewer respondents and displayed in the supplementary tables have been denoted with ‘*’ to safeguard the confidentiality of respondents with rare characteristics.",
                           "Cells with true zero counts are denoted with ‘.’ or 0 throughout, unless denoted ‘*’ as part of disclosure control."),
  "Further Information" = c("[SSCQ website](https://www.gov.scot/collections/scottish-surveys-core-questions/)"),
  "Contact" = c("[sscq@gov.scot](mailto:sscq@gov.scot)")
)

### 4 - Notes - TO UPDATE ----

# notes and their explanation
notes_lookup <- tibble::tribble(
  ~short, ~text,
  "Single adult", "Single adult: 1 adult (not pension aged), no children",
  "Small adult", "Small adult: 2 adults (not pension aged), no children",
  "Large adult", "Large adult: 3 or more adults/pensioners, no children",
  "Single parent", "Single parent: 1 adult/pensioner, 1 or more children",
  "Small family", "Small family: 2 adults/pensioners, 1 or 2 children",
  "Large family", "Large family: 2 adults/pensioners + 3 or more children, 3 or more adults/pensioners + 1 or more children",
  "Single pensioner", "Single pensioner: One adult pension aged, no children",
  "Older couple", "Older couple: 1 or more adults pension aged or 1 adult + 1 pensioner, no children",
  "Scotland", "Scotland: Respondents who specifically list 'Scotland' as their country of birth",
  "Rest of UK", "Rest of UK: England, Northern Ireland, Wales, Great Britain/United Kingdom (Not Otherwise Specified)",
  "EU27", "EU27: Austria, Belgium, Bulgaria, Croatia, Cyprus (European Union), Czech Republic, Denmark, Estonia, Finland, France, Germany, Greece, Hungary, Ireland, Italy, Latvia, Lithuania, Luxembourg, Malta, Netherlands, Poland, Portugal, Romania, Slovakia, Slovenia, Spain, Sweden",
  "Rest of World", "Rest of World: All other responses (excluding refusals)",
  "White: Other", "'White: Other' includes ‘White: Irish’, ‘White: Gypsy/Traveller’ and ‘White: Other White Ethnic Group’",
  "Asian", "'Asian' includes the categories Asian, Asian Scottish or Asian British",
  "All other ethnic groups", "'All other ethnic groups' includes categories within the 'Mixed or Multiple Ethnic Group', ‘African’, ‘Caribbean or Black’, 'Roma', 'Showman/Showwoman' and ‘Other Ethnic Group’ sections",
  "Other", "The 'Other' group includes Hindu, Buddhist, Pagan, Jewish, Sikh, and 'Another religion' responses",
  "No", "'No' includes responses 'No' and those currently serving",
  "Yes", "'Yes' includes responses 'Yes, previously served in regular Armed Forces' and 'Yes, previously served in reserve Armed Forces'"
)

notes_lookup$number <- paste0("[note ", 1:length(notes_lookup$short), "]")


### 5 - Variable lookups - TO UPDATE ----

# variable name, sheet title and tab name of variables with
# their own sheet 
# (i.e., all variables which have their own table and are 
# listed in the contents sheet)
lookup_df <- data.frame(rbind(
  c("genhealth_F", "Self-assessed General Health (All Categories)", "General Health"),
  c("genhealth", "Self-assessed General Health (Grouped)", "General Health Grouped"),
  c("LTCondition", "Limiting Long-term Physical or Mental Health Condition", "Long-term Conditions"),
  c("smoking", "Currently Smokes Cigarettes", "Smoking"),
  c("IndCare", "Provides Care", "Care"),
  c("swemwbs", "Average Shortened Warwick-Edinburgh Mental Wellbeing (SWEMWBS) Score",
    "Mental Wellbeing"),
  c("CrimeArea_F", "Perceptions of Local Crime Rate (All Categories)", "Crime in Area"),
  c("CrimeArea", "Perceptions of Local Crime Rate (Grouped)", "Crime in Area Grouped"),
  c("PolConA_F", "Confidence in Police to Prevent Crime (All Categories)", "Police Confidence A"),
  c("PolConA", "Confidence in Police to Prevent Crime (Grouped)", "Police Confidence A Grouped"),
  c("PolConB_F", "Confidence in Police to Respond Quickly to Appropriate Calls and Information from the Public (All Categories)", "Police Confidence B"),
  c("PolConB", "Confidence in Police to Respond Quickly to Appropriate Calls and Information from the Public (Grouped)", "Police Confidence B Grouped"),
  c("PolConC_F", "Confidence in Police to Deal with Incidents as they Occur (All Categories)", "Police Confidence C"),
  c("PolConC", "Confidence in Police to Deal with Incidents as they Occur (Grouped)", "Police Confidence C Grouped"),
  c("PolConD_F", "Confidence in Police to Investigate Incidents after they Occur (All Categories)", "Police Confidence D"),
  c("PolConD", "Confidence in Police to Investigate Incidents after they Occur (Grouped)", "Police Confidence D Grouped"),
  c("PolConE_F", "Confidence in Police to Solve Crimes (All Categories)", "Police Confidence E"),
  c("PolConE", "Confidence in Police to Solve Crimes (Grouped)", "Police Confidence E Grouped"),
  c("PolConF_F", "Confidence in Police to Catch Criminals (All Categories)", "Police Confidence F"),
  c("PolConF", "Confidence in Police to Catch Criminals (Grouped)", "Police Confidence F Grouped"),
  c("htype2a", "Household Type", "Household Type"),
  c("outten", "Detailed Tenure", "Tenure"),
  c("CarAccess", "Car Access", "Car Access"),
  c("cobeu17", "Country of Birth", "Country of Birth"),
  c("ethSuperGroup", "Ethnic Group", "Ethnic Group"),
  c("religionB", "Religion", "Religion"),
  c("sexIDg", "Sexual Orientation", "Sexual Orientation"),
  c("asg", "Respondent Age and Sex", "Age and Sex"),
  c("ageG", "Respondent Age", "Age"),
  c("marStatB", "Marital Status", "Marital Status"),
  c("sex", "Respondent Sex", "Sex"),
  c("ILOEmp", "Respondent Economic Activity", "Economic Activity"),
  c("TopQual", "Highest Qualification Held", "Highest Qualification"),
  c("SIMD20Q", "Scottish Index of Multiple Deprivation - Quintiles", "SIMD Quintiles"),
  c("UrbRur20Code", "Urban/Rural Classification", "Urban Rural"),
  c("Vets", "Veteran Status", "Veterans")
))
names(lookup_df) <- c("vname", "title", "tabname")
lookup_df$vname <- tolower(lookup_df$vname)
lookup_df$title <- paste0("Table ", 
                          1:length(lookup_df$title), 
                          ": ", 
                          lookup_df$title)
head(lookup_df)

### 6 - Formatting - TO UPDATE ----

# Labels for variables included in CI sheets.
# Provided labels must match the labels used in SAS.
f_trans_factor <- function(x) {
  dplyr::case_when(
    x == "ageG1"       ~ "16-24",
    x == "ageG2"       ~ "25-34",
    x == "ageG3"       ~ "35-44",
    x == "ageG4"       ~ "45-54",
    x == "ageG5"       ~ "55-64",
    x == "ageG6"       ~ "65-74",
    x == "ageG7"      ~ "75+",
    x == "asg1" ~ "Female 16-24",
    x == "asg2" ~ "Female 25-34",
    x == "asg3" ~ "Female 35-44",
    x == "asg4" ~ "Female 45-54",
    x == "asg5" ~ "Female 55-64",
    x == "asg6" ~ "Female 65-74",
    x == "asg7" ~ "Female 75+",
    x == "asg8" ~ "Male 16-24",
    x == "asg9" ~ "Male 25-34",
    x == "asg10" ~ "Male 35-44",
    x == "asg11" ~ "Male 45-54",
    x == "asg12" ~ "Male 55-64",
    x == "asg13" ~ "Male 65-74",
    x == "asg14" ~ "Male 75+",
    x == "LTCondition1" ~ "Limiting condition",
    x == "LTCondition2" ~ "No limiting condition",
    x == "smoking1" ~ "Yes",
    x == "smoking2" ~ "No",
    x == "IndCare1" ~ "Provides unpaid care",
    x == "IndCare2" ~ "No care",
    x == "CrimeArea_F1" ~ "A lot more",
    x == "CrimeArea_F2" ~ "A little more",
    x == "CrimeArea_F3" ~ "About the same",
    x == "CrimeArea_F4" ~ "A little less",
    x == "CrimeArea_F5" ~ "A lot less",
    x == "CrimeArea1" ~ "About the same/A little/A lot more",
    x == "CrimeArea2" ~ "A little/A lot less",
    x == "PolConA_F1" ~ "Very confident",
    x == "PolConA_F2" ~ "Fairly confident",
    x == "PolConA_F3" ~ "Not very confident",
    x == "PolConA_F4" ~ "Not at all confident",
    x == "PolConA1" ~ "Very/fairly confident",
    x == "PolConA2" ~ "Not very/not at all confident",
    x == "PolConB_F1" ~ "Very confident",
    x == "PolConB_F2" ~ "Fairly confident",
    x == "PolConB_F3" ~ "Not very confident",
    x == "PolConB_F4" ~ "Not at all confident",
    x == "PolConB1" ~ "Very/fairly confident",
    x == "PolConB2" ~ "Not very/not at all confident",
    x == "PolConC_F1" ~ "Very confident",
    x == "PolConC_F2" ~ "Fairly confident",
    x == "PolConC_F3" ~ "Not very confident",
    x == "PolConC_F4" ~ "Not at all confident",
    x == "PolConC1" ~ "Very/fairly confident",
    x == "PolConC2" ~ "Not very/not at all confident",
    x == "PolConD_F1" ~ "Very confident",
    x == "PolConD_F2" ~ "Fairly confident",
    x == "PolConD_F3" ~ "Not very confident",
    x == "PolConD_F4" ~ "Not at all confident",
    x == "PolConD1" ~ "Very/fairly confident",
    x == "PolConD2" ~ "Not very/not at all confident",
    x == "PolConE_F1" ~ "Very confident",
    x == "PolConE_F2" ~ "Fairly confident",
    x == "PolConE_F3" ~ "Not very confident",
    x == "PolConE_F4" ~ "Not at all confident",
    x == "PolConE1" ~ "Very/fairly confident",
    x == "PolConE2" ~ "Not very/not at all confident",
    x == "PolConF_F1" ~ "Very confident",
    x == "PolConF_F2" ~ "Fairly confident",
    x == "PolConF_F3" ~ "Not very confident",
    x == "PolConF_F4" ~ "Not at all confident",
    x == "PolConF1" ~ "Very/fairly confident",
    x == "PolConF2" ~ "Not very/not at all confident",
    x == "htype2a1" ~ "Single adult",
    x == "htype2a2" ~ "Small adult",
    x == "htype2a3" ~ "Large adult",
    x == "htype2a4" ~ "Single parent",
    x == "htype2a5" ~ "Small family",
    x == "htype2a6" ~ "Large family",
    x == "htype2a7" ~ "Single pensioner",
    x == "htype2a8" ~ "Older couple",
    x == "outten1" ~ "Owned outright",
    x == "outten2" ~ "Mortgaged",
    x == "outten3" ~ "Social rented",
    x == "outten4" ~ "Private rented",
    x == "outten5" ~ "Unknown rented",
    x == "CarAccess1" ~ "1 car",
    x == "CarAccess2" ~ "2 cars",
    x == "CarAccess3" ~ "3 cars", 
    x == "CarAccess4" ~ "No cars",
    x == "cobeu171" ~ "Scotland",
    x == "cobeu172" ~ "Rest of UK",
    x == "cobeu173" ~ "EU27",
    x == "cobeu174" ~ "Rest of World",
    x == "ethSuperGroup1" ~ "White: Scottish",
    x == "ethSuperGroup2" ~ "White: Other British",
    x == "ethSuperGroup3" ~ "White: Polish",
    x == "ethSuperGroup4" ~ "White: Other",
    x == "ethSuperGroup5" ~ "Asian",
    x == "ethSuperGroup6" ~ "All other ethnic groups",
    x == "religionB1" ~ "None", 
    x == "religionB2" ~ "Church of Scotland", 
    x == "religionB3" ~ "Roman Catholic", 
    x == "religionB4" ~ "Other Christian", 
    x == "religionB5" ~ "Muslim", 
    x == "religionB6" ~ "Other", 
    x == "sexIDg1" ~ "Heterosexual",
    x == "sexIDg2" ~ "LGB & other",
    x == "marStatB1" ~ "Never married - single",
    x == "marStatB2" ~ "Married/Civil partnership",
    x == "marStatB3" ~ "Seperated",
    x == "marStatB4" ~ "Divorced/Dissolved civil partnership",
    x == "marStatB5" ~ "Widowed/Bereaved civil partner",
    x == "sex1" ~ "Male",
    x == "sex2" ~ "Female",
    x == "ILOEmp1" ~ "In employment",
    x == "ILOEmp2" ~ "Unemployed",
    x == "ILOEmp3" ~ "Inactive",
    x == "TopQual1" ~ "Level 1 - O Grade, Standard Grade or equiv (SVQ level 1 or 2)",
    x == "TopQual2" ~ "Level 2 - Higher, A level or equivalent (SVQ Level 3)",
    x == "TopQual3" ~ "Level 3 - HNC/HND or equivalent (SVQ Level 4)",
    x == "TopQual4" ~ "Level 4 - Degree, Professional qualification (Above SVQ Level 4)",
    x == "TopQual5" ~ "Other qualification",
    x == "TopQual6" ~ "No qualifications",
    x == "SIMD20Q1" ~ "1",
    x == "SIMD20Q2" ~ "2",
    x == "SIMD20Q3" ~ "3",
    x == "SIMD20Q4" ~ "4",
    x == "SIMD20Q5" ~ "5",
    x == "UrbRur20Code1" ~ "Large Urban Area",
    x == "UrbRur20Code2" ~ "Other Urban Area",
    x == "UrbRur20Code3" ~ "Accessible Small Town",
    x == "UrbRur20Code4" ~ "Remote Small Town",
    x == "UrbRur20Code5" ~ "Accessible Rural",
    x == "UrbRur20Code6" ~ "Remote Rural",
    x == "Vets0" ~ "No",
    x == "Vets1" ~ "Yes",
    x == "genhealth_F1" ~ "Very good",
    x == "genhealth_F2" ~ "Good",
    x == "genhealth_F3" ~ "Fair",
    x == "genhealth_F4" ~ "Bad",
    x == "genhealth_F5" ~ "Very bad",
    x == "genhealth1" ~ "Good/Very good",
    x == "genhealth2" ~ "Fair",
    x == "genhealth3" ~ "Bad/Very bad",
    x == "topQual1" ~ "Level 1 - O Grade, Standard Grade or equiv (SVQ level 1 or 2)",
    x == "topQual2" ~ "Level 2 - Higher, A level or equivalent (SVQ Level 3)",
    x == "topQual3" ~ "Level 3 - HNC/HND or equivalent (SVQ Level 4)",
    x == "topQual4" ~ "Level 4 - Degree, Professional qualification (Above SVQ Level 4)",
    x == "topQual5" ~ "Other qualification",
    x == "topQual6" ~ "No qualifications",
    x == "swemwbs" ~ "Average Wellbeing (/35)"
  )
}
### 7 - Required variables - TO UPDATE ----

# Variable names (=levels) and labels of variables to be 
# included in each sheet (as rows).
# The labels must match the label provided by SAS as the QA and CI data
# won't match otherwise.
# Variable names should be in lower case.
# The order of the variables determines the order in the final output.
reqvar <- tibble::tribble(
  ~levels, ~labels,
  "all", "All", 
  "simd20q", "Scottish Index of Multiple Deprivation - Quintiles", 
  "urbrur20code", "Urban/Rural Classification",
  "la", "Local Authority", 
  "healthboard", "Health Board",
  "htype2a", "Household Type",
  "outten", "Detailed Tenure",
  "caraccess", "Car Access", 
  "cobeu17", "Country of Birth", 
  "ethsupergroup", "Ethnic Group", 
  "religionb", "Religion",
  "sexidg", "Sexual Orientation", 
  "asg", "Respondent Age and Sex",
  "ageg","Respondent Age Group",
  "marstatb", "Marital Status", 
  "topqual", "Highest Qualification Held",
  "iloemp", "Respondent Economic Activity",
  "vets", "Veterans",
  "ltcondition", "Limiting Long-term Physical or Mental Health Condition",
  "smoking", "Currently Smokes Cigarettes",
  "indcare", "Provides unpaid care",
  "psd", "Police Scotland Division"
)

### 8 - Load packages ----

library(tidyverse)
library(here)
library(readxl)
library(sdcTable)
library(a11ytables)
library(data.table)
library(odsconvertr)

### 9 - Load functions from functions folder of Rproject ----

walk(
  list.files(here("functions"), pattern = "\\.R$", full.names = TRUE),
  source
)

### 10 - Load config file from code folder of RProject ----

# The config.R script is the only file which needs to be updated before
# the RAP can be run.

source(here::here("code", "config.R"))
