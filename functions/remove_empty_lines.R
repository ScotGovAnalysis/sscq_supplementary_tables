#########################################################################
# Name of file - remove_empty_lines.R
#
# Type - Reproducible Analytical Pipeline (RAP)
# Written/run on - RStudio Desktop
# Version of R - 4.2.2
#
# Description - Overwrites internal functions of a11ytables package
# to remove empty lines caused by the .has_notes function.

#########################################################################

my.get_start_row_blanks_message <- function(has_notes, start_row = 3) {
  
  if (has_notes) {
    start_row <- start_row  #+ 1
  }
  
  return(start_row)
  
}

assignInNamespace(".get_start_row_blanks_message", 
                  my.get_start_row_blanks_message , 
                  ns = "a11ytables")


my.get_start_row_custom_rows <- function(
    has_notes,
    has_blanks_message,
    start_row = 3
) {
  
  if (has_notes) {
    start_row <- start_row #+ 1
  }
  
  if (has_blanks_message) {
    start_row <- start_row + 1
  }
  
  return(start_row)
  
}

assignInNamespace(".get_start_row_custom_rows", 
                  my.get_start_row_custom_rows , 
                  ns = "a11ytables")

my.get_start_row_source <- function(
    content,
    tab_title,
    has_notes,
    has_blanks_message,
    has_custom_rows,
    start_row = 3
) {
  
  if (has_notes) {
    start_row <- start_row #+ 1
  }
  
  if (has_blanks_message) {
    start_row <- start_row + 1
  }
  
  if (has_custom_rows) {
    custom_rows <- content[content$tab_title == tab_title, "custom_rows"][[1]]
    start_row <- start_row + length(custom_rows)
  }
  
  return(start_row)
  
}

assignInNamespace(".get_start_row_source", 
                  my.get_start_row_source , 
                  ns = "a11ytables")

my.get_start_row_table <- function(
    content,
    tab_title,
    has_notes,
    has_blanks_message,
    has_custom_rows,
    has_source,
    start_row = 3
) {
  
  if (has_notes) {
    start_row <- start_row #+ 1
  }
  
  if (has_blanks_message) {
    start_row <- start_row + 1
  }
  
  if (has_custom_rows) {
    custom_rows <- content[content$tab_title == tab_title, "custom_rows"][[1]]
    start_row <- start_row + length(custom_rows)
  }
  
  if (has_source) {
    start_row <- start_row + 1
  }
  
  return(start_row)
  
}

assignInNamespace(".get_start_row_table", 
                  my.get_start_row_table, 
                  ns = "a11ytables")




my.insert_notes_statement <- function(wb, content, tab_title) {
  
  has_notes <- .has_notes(content, tab_title)
  
  if (has_notes) {
    
   # text <-
   #   "This table contains notes, which can be found in the Notes worksheet."
    
   # openxlsx::writeData(
   #   wb = wb,
   #   sheet = tab_title,
   #   x = text,
   #   startCol = 1,
   #   startRow = 3  # notes will always go in row 3 if they exist
   # )
    
  }
  
  return(wb)
  
}

assignInNamespace(".insert_notes_statement", 
                  my.insert_notes_statement, 
                  ns = "a11ytables")


