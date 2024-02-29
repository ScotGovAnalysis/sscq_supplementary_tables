#' @title Import QA data and prepares it for further steps
#'
#' @description This function reads in the QA Excel file, 
#' loops through the different sheets and cleans the data.
#'
#' @param path Path to folder containing the QA sheet 
#' outputted by SAS.
#'
#' @returns Cleaned list of tibbles containing QA data.
#'
#' @examples
#' get_qa_data(qa.path)

get_qa_data <- function(path) {
  
  # get names of Excel sheets
  sheetnames.qa <- excel_sheets(path)
  
  # import all sheets and add to list
  qa.list <- lapply(excel_sheets(path), read_excel, 
                    path = path, skip = 3, col_names = TRUE)
  
  # clean list names
  names(qa.list) <- gsub("^.*\\- ","", sheetnames.qa)
  
  # loop through the list and select relevant variables 
  # (i.e., Weighted N and number of observations in each category)
  cleaned.qa.list <- list()
  for (c in seq_along(qa.list)) {
    
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
      mutate(category = gsub("\\*+","", category))
    
    # remove unwanted * in variables 
    # (* was previously used for footnotes but is no longer needed)
    names(cleaned_df) <- gsub("\\.[0-9]*","", names(cleaned_df))
    names(cleaned_df) <- gsub("\\*+","", names(cleaned_df))
    
    # Assign 'All' to 1st and 2nd column of 1st row
    cleaned_df[1,1:2] <- "All"
    
    # add df to list
    cleaned.qa.list <- c(cleaned.qa.list, list(cleaned_df))
    
    # set variable name as item name in list
    names(cleaned.qa.list)[c] <- names(qa.list[c]) %>% tolower()
  }
  
  return(cleaned.qa.list)
  
}
