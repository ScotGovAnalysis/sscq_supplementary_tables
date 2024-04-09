#' @title Export supplementary tables as accessible Excel file
#'
#' @description This function looks up the sheet titles and 
#' tab names in the lookup data frame, creates a cover sheet,
#' a contents sheet, a notes sheet and exports them along with
#' the tables to an accessible Excel file.
#'
#' @param data List of all tables to be exported.
#'
#' @returns Accessible Excel file with all supplementary tables.
#'
#' @examples
#' export_tables(data)

export_tables <- function(data) {
  
  ### 1 - Sheet titles and tab names ----
  
  # get sheet title from lookup table
  sheet.titles <- data.table(var = names(data))
  sheet.titles[lookup_df, var := title, on = .(var = vname)]
  sheet.titles <- as.vector(sheet.titles)[[1]]
  
  # get tab names from lookup table
  tab.names <- data.table(var = names(data))
  tab.names[lookup_df, var := tabname, on = .(var = vname)]
  tab.names <- as.vector(tab.names)[[1]]
  tab.names.cleaned <- gsub(" ", "_", tab.names)
  tab.names.cleaned <- gsub("-", "", tab.names.cleaned)
  
  ### 3 - Contents sheet ----
  
  # create contents page
  contents_df <- data.frame(
    "Sheet name" = paste0("=HYPERLINK(\"#'", 
                          c("Notes", tab.names.cleaned), 
                          "'!A1\", \"", 
                          c("Notes", tab.names.cleaned), 
                          "\")"),
    "Sheet title" = c("Notes", sheet.titles),
    check.names = FALSE
  )
  class(contents_df$`Sheet name`) <- 'formula'
  
  ### 4 - Notes ----
  
  # create notes page
  notes_df <- data.frame(
    "Note number" = notes_lookup$number,
    "Note text" = notes_lookup$text,
    check.names = FALSE
  )
  
  ### 5 - Data ----
  
  # create table
  my_a11ytable <-
    a11ytables::create_a11ytable(
      tab_titles = c(
        "Cover",
        "Contents",
        "Notes",
        tab.names.cleaned
      ),
      sheet_types = c(
        "cover",
        "contents",
        "notes",
        rep("tables", length(data))
      ),
      sheet_titles = c(
        paste0("Scottish Surveys Core Questions (SSCQ) ",
               sscq_year, ": Supplementary Tables"),
        "Contents",
        "Notes",
        sheet.titles
      ),
      blank_cells = c(
        rep(NA_character_, 3+length(data))
      ),
      custom_rows = c(
        rep(list(NA_character_), 2),
        rep(list(" "), length(data)+1)
      ),
      sources = c(
        rep(NA_character_, 3),
        rep(paste0("Scottish Household Survey, ", 
                   "Scottish Health Survey ",
                   "and Scottish Crime and Justice Survey"), length(data))
      ),
      tables = c(
        list(
          cover_list,
          contents_df,
          notes_df),
        data
      )
    )
  
  
  ### 6 - Export to XLSX ----
  
  my_wb <- a11ytables::generate_workbook(my_a11ytable)
  
  # add link colour to content sheet
  linkstyle <- openxlsx::createStyle(fontColour = "#0000EE", 
                                     textDecoration = "underline")
  openxlsx::addStyle(wb = my_wb, 
                     sheet = "Contents", 
                     style = linkstyle, 
                     rows = 4:(4+length(data)), 
                     cols = 1, 
                     stack = TRUE)
  
  # add 'back to contents page' to each sheet
  for(i in 3:(3+length(data))){
  openxlsx::writeFormula(my_wb, sheet = i, 
                         x = "=HYPERLINK(\"#'Contents'!A1\", \"Back to Contents page\")", 
                         startCol = 1, startRow = 3)
  }
  
  # change width of first column in contents sheet
  openxlsx::setColWidths(my_wb, sheet = "Contents", cols = 1, widths = "auto")

  # open temp copy
  openxlsx::openXL(my_wb)
  
  # export to xlsx file
  openxlsx::saveWorkbook(my_wb, 
                         paste0(export.path,
                                       fname),
                         overwrite = TRUE)
  
  convert_to_ods(paste0(export.path,
                        fname))
}

