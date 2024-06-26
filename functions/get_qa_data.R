#' @title Import QA data and prepares it for further steps
#'
#' @description This function reads in the QA Excel file, 
#' loops through the different sheets and cleans the data.
#'
#' @param filepath Path to folder containing the QA sheet 
#' outputted by SAS.
#'
#' @returns Cleaned list of tibbles containing QA data.
#'
#' @examples
#' get_qa_data(qa.path)

get_qa_data <- function(filepath) {
  
  # add message to inform user about progress
  message("Importing QA data")
  
  path <- filepath
  
  # get names of Excel sheets
  sheetnames.qa <- excel_sheets(path)
  
  # import all sheets and add to list
  qa.list <- pblapply(excel_sheets(path), read_excel, 
                    path = path, skip = 3, col_names = TRUE,
                    .name_repair = "unique_quiet")
  
  # clean list names
  names(qa.list) <- gsub("^.*\\- ","", sheetnames.qa)
  
  # remove swemwbs sheet as this variable is continuous
  qa.list[["swemwbs"]] <- NULL
  
  # loop through the list and select relevant variables 
  # (i.e., Weighted N and number of observations in each category)
  cleaned.qa.list <- list()
  for (c in seq_along(qa.list)) {
    
    # add message to inform user about progress
    message(paste0("Preparing data for ", 
                   names(qa.list)[c]))
    
    # get item c in list
    cleaned_df <- qa.list[[c]]
    
    cleaned_df <- cleaned_df %>% 
      
      # create new column with variable name
      mutate(varname = ifelse(is.na(.[[2]]) == TRUE, .[[1]], NA)) %>%
      
      # reorder columns: move new variable to beginning
      select(length(cleaned_df)+1, 1:length(cleaned_df)) %>%
      
      # replace NAs in new variable with existing values (filled downwards)
      fill(varname) %>% 
      
      # remove rows with NA in third column
      filter(!is.na(.[[3]])) %>%
      
      # select relevant columns
      select(1:2, Sum:N) %>%
      
      # rename 2nd and 3rd columns
      rename_with(~ c('category', "Weighted N"), c(2, 3)) %>%
      
      # remove unwanted * (see below for explanation)
      mutate(category = gsub("\\*+","", category)) %>%
      
      # remove columns that include the word 'refused'
      select(-contains("refused", ignore.case = TRUE)) %>%
      
      # remove double spaces
      mutate(varname = str_squish(varname),
             category = str_squish(category))
    
    # remove unwanted * in variables 
    # (* was previously used for footnotes but is no longer needed)
    names(cleaned_df) <- gsub("\\.[0-9]*","", names(cleaned_df))
    names(cleaned_df) <- gsub("\\*+","", names(cleaned_df))
    
    # Assign 'All' to 1st and 2nd column of 1st row
    cleaned_df[1,1:2] <- "All"
    
    cleaned_df <- cleaned_df %>%
      
      # remove unnecessary line breaks
      mutate(across(everything(), ~ gsub("[\r\n]", " ", .)),
             
             # remove "1 =", "2 = " in case SAS didn't properly label data
             across(everything(), ~ gsub("[[:digit:]] = ", "", .)),
             
             # rename variables if SAS hasn't labelled them properly
             varname = ifelse(varname == "cobeu17" , "Country of Birth", varname))
    
    # add df to list
    cleaned.qa.list <- c(cleaned.qa.list, list(cleaned_df))
    
    # set variable name as item name in list
    names(cleaned.qa.list)[c] <- names(qa.list[c]) %>% tolower()
  }
  
  return(cleaned.qa.list)
  
}
