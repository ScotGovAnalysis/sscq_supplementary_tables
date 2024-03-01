#' @title Import CI data and create supplementary tables
#'
#' @description This function reads in the data, loops through 
#' the different Excel files, loops through the different sheets 
#' in each Excel file and cleans the data to create the 
#' supplementary tables.
#'
#' @param paths File paths to CI sheets outputted by SAS.
#'
#' @returns Cleaned list of data frames ready for export.
#'
#' @examples
#' get_ci_data(ci.path)

get_ci_data <- function(paths){
  
  # loop through file names
  ci_tables_list <- list()
  for (a in seq_along(paths)) {
    
    # get path of one file
    path <- paths[a]
    
    # get sheet names
    sheetnames <- excel_sheets(path)[-1]
    
    # import all but the first sheet of filename a and add to list
    mylist <- lapply(excel_sheets(path)[-1], read_excel, 
                     path = path, skip = 10, col_names = FALSE)
    
    # add sheet names as list names
    names(mylist) <- sheetnames
    
    # loop through each list item
    ci_list <- list()
    col <- list()
    for (i in seq_along(mylist)){
      
      # split list item i into separate tables
      # (SAS unhelpfully outputs multiple tables to a 
      # single Excel sheet)
      split_table <- mylist[[i]] %>%
        split_df(complexity = 1)
      
      # remove tibbles with only one row
      split_table <- split_table[purrr::map_lgl(split_table, ~ nrow(.) != 1)]
      
      # assign name to list item, remove first two rows and
      # select relevant columns
      split_table <- lapply(split_table, function(w){
        names(w) <- w[2,]
        w <- w[-c(1:2),]
        w <- w[,c(1:4, 6:7)]
        return(w)
      }
      )
      
      split_table[[1]] <- split_table[[1]] %>%
        `colnames<-` (c(colnames(split_table[[2]])[2:6], "All")) %>%
        mutate(All = "All",
               `NA` = as.double(`NA`)) %>%
        select(c(6, 1:5))
      
      # assign name to second column in first list item
      colnames(split_table[[1]])[2] <- "Variable"
      
      # assign name to each list element
      names(split_table) <- names(sapply(split_table, "[", 1))
      names(split_table)[1] <- "all"
      
      # rename first column in each list item
      split_table <- lapply(split_table, function(k){
        names(k)[1] <- "Category"
        return(k)
      })
      
      # combine all list items into one data frame
      df <- bind_rows(split_table, .id = "category")
      
      # remove trailing digits from value in first row
      # second column
      var <- gsub("[[:digit:]]*$", "", df[1,3]) %>% tolower()
      var <- ifelse(var == "cobeu", "cobeu17", var)
      
      # remove * from values in Category column
      # (notes were previously indicated with * but this is no
      # longer the case)
      df$Category <- gsub("\\*+","", df$Category)
      
      # assign 'All' two first two columns in first row
      df[1,1:2] <- "All"
      
      # clean the data frame for export
      df <- df %>% 
        
        # rename columns
        `colnames<-` (c("Variable", "Category", "group", 
                        "Unweighted N", "%", "95% CI\nlower limit",
                        "95% CI\nupper limit")) %>%
        
        # convert tibble to data frame to facilitate later steps
        as.data.frame(.) %>%
        
        # convert decimals to %, replace true 0 with . and 
        # round to 1 decimal
        mutate(`%` = as.numeric(`%`)*100,
               `%` = ifelse(`%` == 0, ".", round(`%`, 1)),
               `95% CI\nlower limit` = as.numeric(`95% CI\nlower limit`)*100,
               `95% CI\nlower limit` = ifelse(`95% CI\nlower limit` == 0, 
                                              ".", 
                                              round(`95% CI\nlower limit`, 1)),
               `95% CI\nupper limit` = as.numeric(`95% CI\nupper limit`)*100,
               `95% CI\nupper limit` = ifelse(`95% CI\nupper limit` == 0, 
                                              ".", 
                                              round(`95% CI\nupper limit`, 1)),
               `Unweighted N` = as.numeric(`Unweighted N`),
               
               # transform 'Variable' to ordered factor
               Variable =  factor(tolower(Variable),
                                  ordered = TRUE, 
                                  levels = reqvar$levels,
                                  labels = reqvar$labels),
               
               # format 'group' column
               across(group, .fns = f_trans_factor)) %>%
        
        # sort by 'Variable'
        arrange(Variable)
      
      # add notes to 'group' column
      for (b in 1:length(notes_lookup$short)) {
        df <- df %>% mutate(group = ifelse(group == notes_lookup$short[b],
                                           paste0(group, 
                                                  "\n",
                                                  notes_lookup$number[b]),
                                           group))
      }
      
      # change column names of 4-7th columns
      # ('group' value is added as prefix to column name)
      df <- df %>%
        rename(!!paste0((.)[3][5, ], "\n(", colnames(.)[5], ")") := 5,
               !!paste0((.)[3][6, ], "\n(", colnames(.)[6], ")") := 6,
               !!paste0((.)[3][7, ], "\n(", colnames(.)[7], ")") := 7) %>%
        mutate(Variable = as.character(Variable))
      
      # merge data frame with qa data to add 'Weighted N' column
      df <- df %>%
        left_join(y = cleaned_qa_list[[tolower(var)]][, 
                                                      c(1:3, 3+i,
                                                        length(cleaned_qa_list[[tolower(var)]]))], 
                  by = c("Variable" = "varname", "Category" = "category"))
        
      # transform . to 100.1 to avoid the introduction of NAs in next step
      df[[9]] <- ifelse(df[[9]] == ".", 100.1, df[[9]])
      
      # transform column with number of observations to numeric
      df[[9]] <- as.numeric(df[[9]])
      
      df <- df %>%
        
        # suppress values which are based on 0 or 1 observation
        mutate(across(c(5:7), \(x) ifelse(.[[9]] <= 1, "*", 
                                          # re-transform 101.1 to .
                                          ifelse(.[[9]] == 101.1, ".", x)))) %>%
        
        # delete response categories which include 'refused'
        filter(!str_detect(Category, regex('refused', ignore_case = T))) %>%
        
        # select relevant columns
        select(c(1:3, 8, 4, 5:7))
      
      # add notes to column names
      for (w in 1:length(notes_lookup$short)) {
        df <- df %>% mutate(Category = ifelse(Category == notes_lookup$short[w],
                                              paste0(Category, 
                                                     "\n",
                                                     notes_lookup$number[w]),
                                              Category))
      }
      
      # get columns that are named the same for all list items i
      col <- df %>% select(-group) %>% select(c(1:4))
      
      # add unique columns of list item i to list
      # (%, upper CI and lower CI)
      ci_list <- c(ci_list, list(df[, 6:8]))
      
    }
    
    # merge common columns and unique columns to data frame
    ci_tables <- do.call("cbind", list(col, ci_list))
    
    # reorder columns
    ci_tables <- ci_tables %>% 
      relocate(c(`Weighted N`, `Unweighted N`), .after = last_col()) 
    
    # add data frame to list
    ci_tables_list <- c(ci_tables_list, list(ci_tables))
    
    # assign variable name as list item name
    names(ci_tables_list)[a] <- var
  }
  
  return(ci_tables_list)
}
