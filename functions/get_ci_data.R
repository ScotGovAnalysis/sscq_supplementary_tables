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
    
    # add message to inform user about progress
    message(paste0("Importing data for ", 
                   word(excel_sheets(path), 1)[1]))
    
    if(all(excel_sheets(path) != "swemwbs")){
    # get sheet names
    sheetnames <- excel_sheets(path)[-1]
    
    # import all but the first sheet of filename a and add to list
    mylist <- pblapply(excel_sheets(path)[-1], read_excel, 
                     path = path, col_names = FALSE,
                     .name_repair = "unique_quiet")
    }
    
    # different code for continuous variables
    if(all(excel_sheets(path) == "swemwbs")){
      
      # get sheet names
      sheetnames <- "swemwbs"
      
      # import sheet of filename a and add to list
      mylist <- pblapply(excel_sheets(path), read_excel, 
                       path = path, col_names = FALSE,
                       .name_repair = "unique_quiet")
      
    }
    
    # add sheet names as list names
    names(mylist) <- sheetnames
    
    # loop through each list item
    ci_list <- list()
    col <- list()
    n_list <- list()
    for (i in seq_along(mylist)){
      
      # add message to inform user about progress
      message(paste0("  Preparing data for ", 
                     names(mylist)[i]))
      
      # split list item i into separate tables
      # (SAS unhelpfully outputs multiple tables to a 
      # single Excel sheet)
      split_table <- mylist[[i]] %>%
        split_df(complexity = 1)
      
      # remove tibbles with only one row
      split_table <- split_table[purrr::map_lgl(split_table, ~ nrow(.) != 1)]
      
      # remove first list item (table with data summary)
      split_table <- split_table[-1]
      
      # remove labels column for continuous variables
      if(split_table[[1]][2,2] == "Label"){
        split_table[[1]] <- split_table[[1]] %>%
          select(c(1, 3:8))
        
        split_table[-1] <- lapply(split_table[-1], function(x) { x["...3"] <- NULL; x })
        
        split_table <- lapply(split_table, function(x) { colnames(x) <- c(paste0("...", 1:7)); x })
      }
      
      split_table
      
      # reorder columns of first table
      split_table[[1]] <- split_table[[1]] %>%
        select(c(7, 1:6)) %>%
        `colnames<-` (c("All", colnames(split_table[[2]])[2:7])) %>%
        mutate(All = as.character(All),
               All = "All",
               `...7` = as.double(`...7`))
      
      # assign name to list item, remove first two rows and
      # select relevant columns
      split_table <- lapply(split_table, function(w){
        names(w) <- w[2,]
        w <- w[-c(1:2),]
        w <- w[,c(1:4, 6:7)]
        return(w)
      }
      )
      
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
        
        # transform % and CIs to numeric
        mutate(`%` = as.numeric(`%`),
              `95% CI\nlower limit` = as.numeric(`95% CI\nlower limit`),
              `95% CI\nupper limit` = as.numeric(`95% CI\nupper limit`))
      
      # multiply % and CIs by 100 if not continuous
      if(var != "swemwbs"){
        df <- df %>% 
          mutate(`%` = `%`*100,
                 `95% CI\nlower limit` = `95% CI\nlower limit`*100,
                 `95% CI\nupper limit` = `95% CI\nupper limit`*100)
        }
      
      
      df <- df %>%
        
        # convert decimals to %, replace true 0 with . and 
        # round to 1 decimal
        mutate(`%` = ifelse(`%` == 0, ".", janitor::round_half_up(`%`, 1)),
               
               # replace negative lower CI values with 0
               `95% CI\nlower limit` = ifelse(`95% CI\nlower limit` < 0, 
                                              0, 
                                              `95% CI\nlower limit`),
               
               # replace upper CI values > 100 with 100
               `95% CI\nupper limit` = ifelse(`95% CI\nupper limit` > 100, 
                                              100, 
                                              `95% CI\nupper limit`),
               
               # convert decimals to %, replace true 0 with . and 
               # round to 1 decimal
               `95% CI\nlower limit` = ifelse(`95% CI\nlower limit` == 0, 
                                              ".", 
                                              janitor::round_half_up(`95% CI\nlower limit`, 1)),
               `95% CI\nupper limit` = ifelse(`95% CI\nupper limit` == 0, 
                                              ".", 
                                              janitor::round_half_up(`95% CI\nupper limit`, 1)),
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
        df <- df %>% mutate(group = ifelse(var != "smoking" & group == notes_lookup$short[b],
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
      
      # delete response categories which include 'refused'
      df <- df %>%
        filter(!str_detect(Category, regex('refused', ignore_case = T)))
      
      if(var != "swemwbs"){
      df <- df %>%
        
        # merge data frame with qa data to add 'Weighted N' column
        left_join(y = cleaned_qa_list[[tolower(var)]][, 
                                                      c(1:3, 3+i,
                                                        length(cleaned_qa_list[[tolower(var)]]))], 
                  by = c("Variable" = "varname", "Category" = "category")) 
      
      # transform . to 999999.1 to avoid the introduction of NAs in next step
      df[[9]] <- ifelse(df[[9]] == ".", 999999.1, df[[9]])
      
      # transform column with number of observations to numeric
      df[[9]] <- as.numeric(df[[9]])
      
      # extract number of observations of item i
      df_n <- df %>% select(9)
      
      df <- df %>%
        
        # suppress values which are based 5 or fewer observations
        mutate(across(c(5:7), \(x) ifelse(.[[9]] <= 5, "*", 
                                          # re-transform 999999.1 to .
                                          ifelse(.[[9]] == 999999.1, ".", x)))) %>%
        
        # select relevant columns
        select(c(1:3, 8, 4, 5:7))
      }
      
      # add notes to column names
      for (w in 1:length(notes_lookup$short)) {
        df <- df %>% mutate(Category = ifelse(
          Variable != "Currently Smokes Cigarettes" & Category == notes_lookup$short[w],
          paste0(Category, 
                 "\n",
                 notes_lookup$number[w]),
          Category))
      }
      
      # Round Weighted and Unweighted N
      df <- df %>% 
        
        # Unweighted N are rounded to nearest 100
        mutate(`Unweighted N` = janitor::round_half_up(
          as.numeric(
            `Unweighted N`), -2)) %>%
        
        # If Weighted N exists, it is transformed to a numeric variable
        mutate_at(vars(contains(c("Weighted N"), ignore.case = F)), 
                  as.numeric) %>%
        
        # If Weighted N exists, it is rounded to nearest 1,000
        mutate_at(vars(contains(c("Weighted N"), ignore.case = F)), 
                  janitor::round_half_up, -3)
    
      
      # get columns that are named the same for all list items i
      col <- df %>% select(-group) %>% select(c("Variable":"Unweighted N"))
      
      # add unique columns of list item i to list
      # (%, upper CI and lower CI)
      ci_list <- c(ci_list, list(df[, (length(df)-2):length(df)]))
      
      if(var != "swemwbs"){
      # add number of observations of item i to list
      n_list <- c(n_list, list(round(df_n, 0)))
      }
    }
    
    # merge common columns and unique columns to data frame
    ci_tables <- do.call("cbind", list(col, ci_list))
    
    if(var != "swemwbs"){
    # bind list into one data frame
    n_tables <- bind_cols(n_list)
    }
    
    # reorder columns
    ci_tables <- ci_tables %>% 
      relocate(any_of(c("Weighted N", "Unweighted N")), .after = last_col()) 
    
    
    # Add disclosure control
    
    if(var != "swemwbs"){
      # identify rows and columns which are based on <= 5 observations
      to_be_suppressed <- which(n_tables <= 5, arr.ind = TRUE) %>% 
        as.data.frame() %>% 
        arrange(row)
      
      # loop through all rows with at least one n <= 5 to apply suppression
      for(row_n in to_be_suppressed[,1]){
        
        # order columns by number of observations
        # note: . are still coded 999999.1
        order <- t(apply(n_tables[row_n, ], 1, order)) %>% as.data.frame()
        
        # identify second lowest n
        s_lowest <- order[, 2]
        
        # reset counter
        number_obs <- 0
        
        # loop through all columns in row
        for(column in to_be_suppressed %>% 
            filter(row == row_n) %>% 
            pull()){
          
          # suppress column with n <= 5
          ci_tables[row_n, (2+3*column-2):(2+3*column)] <- "*"
          
          # suppress column with second lowest n
          ci_tables[row_n, (2+3*s_lowest-2):(2+3*s_lowest)] <- "*"
          
          # count n of suppressed columns
          number_obs <- number_obs + n_tables[row_n, column]
          
          # if only one observation is <= 5, add second lowest to counter
          if(to_be_suppressed %>% filter(row == row_n) %>% count(row) %>% select(n) == 1){
            number_obs <- number_obs + n_tables[row_n, s_lowest]
          }
          
        }
       
        # if counter is <= 5 (i.e., the total number of suppressed observations
        # is still <5), then suppress additional column
        if(number_obs <= 5){
          add <- order[, which(order == column) + 1]
          ci_tables[row_n, (2+3*add-2):(2+3*add)] <- "*"
        }
         
      }
    }

    # add data frame to list
    ci_tables_list <- c(ci_tables_list, list(ci_tables))
    
    # assign variable name as list item name
    names(ci_tables_list)[a] <- var
  }
  
  return(ci_tables_list)
}
