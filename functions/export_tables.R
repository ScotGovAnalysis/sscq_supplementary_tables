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
  
  ### 2 - Cover list ----
  
  # create spreadsheet cover
  cover_df <- data.frame(
    subsection_title = cover$title,
    subsection_content = cover$text,
    check.names = FALSE
  )
  
  ### 3 - Contents sheet ----
  
  # create contents page
  contents_df <- data.frame(
    "Sheet name" = c("Notes", tab.names),
    "Sheet title" = c("Notes", sheet.titles),
    check.names = FALSE
  )
  
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
        tab.names
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
      sources = c(
        rep(NA_character_, 3),
        rep(paste0("Scottish Household Survey, ", 
                   "Scottish Health Survey, ",
                   "and Scottish Crime and Justice Survey"), length(data))
      ),
      tables = c(
        list(
          cover_df,
          contents_df,
          notes_df),
        data
      )
    )
  
  ### 6 - Export to XLSX ----
  
  my_wb <- a11ytables::generate_workbook(my_a11ytable)
  
  # open temp copy
  openxlsx::openXL(my_wb)
  
  # export to xlsx file
  openxlsx::saveWorkbook(my_wb, 
                         paste0(export.path,
                                       fname),
                         overwrite = TRUE)
  
}
