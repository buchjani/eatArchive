#' Format the excel sheet
#'
#' @param wb
#' @param sheet_name
#' @param df
#' @param autocomplete_values
#'
#' @keywords internal

.add_sheet_with_style <- function(wb, sheet_name, df, autocomplete_values = NULL) {
  openxlsx::addWorksheet(wb, sheet_name)

  # Handle AIP column and optional autocomplete rows
  add_aip_col <- !is.null(autocomplete_values)  
  add_autocomplete_rows <- is.character(autocomplete_values)

  if (add_aip_col && !"AIP" %in% names(df)) {
    df$AIP <- NA
  }

  if (add_autocomplete_rows) {
    autocomplete_df <- as.data.frame(matrix(NA, nrow = length(autocomplete_values), ncol = ncol(df)))
    names(autocomplete_df) <- names(df)
    autocomplete_df$AIP <- autocomplete_values
    df <- rbind(autocomplete_df, df)
  }

  # Fix time format
  df$Last_Modified <- as.POSIXct(df$Last_Modified)

  # Add the data table to openxlsx object
  openxlsx::writeData(wb, sheet = sheet_name, x = df, startRow = 1, withFilter = TRUE)

  # Hide autocomplete rows, if any
  # Note: Using height = 1 (not 0) was neccessary for Excel filters/autocomplete to work
  if (is.character(autocomplete_values)) {
    autocomplete_rows <- 2:(1 + length(autocomplete_values))
    openxlsx::setRowHeights(wb, sheet = sheet_name, rows = autocomplete_rows, heights = 1)
  }

  # Header
  headerStyle <- openxlsx::createStyle(fontSize = 12, fontColour = "#FFFFFF", halign = "left",
                                       fgFill = "#821123", border = "TopBottom", borderColour = "#821123")
  openxlsx::addStyle(wb, sheet = sheet_name, cols = 1: ncol(df), rows = 1, style = headerStyle,
                     gridExpand = TRUE, stack = TRUE)
  openxlsx::freezePane(wb, sheet = sheet_name, firstRow = TRUE)

  # Set column widths and styles
  openxlsx::setColWidths(wb, sheet = sheet_name, cols = 1, widths = "auto")  # File name
  openxlsx::setColWidths(wb, sheet = sheet_name, cols = 2, widths = 10)      # File name
  openxlsx::setColWidths(wb, sheet = sheet_name, cols = 3, widths = 20)      # Last modified

  # Format column 'Last_Modified' as date/time
  dateStyle <- openxlsx::createStyle(numFmt = "yyyy-mm-dd hh:mm:ss")
  openxlsx::addStyle(wb, sheet = sheet_name, style = dateStyle,
                     cols = 3, rows = 2:(nrow(df) + 1), gridExpand = TRUE, stack = TRUE)

  # Remove default lines between rows
  openxlsx::addStyle(wb, sheet = sheet_name, cols = 1: ncol(df), rows = 2:nrow(df), gridExpand = TRUE, stack = TRUE,
                     style = openxlsx::createStyle(border = "TopBottom", borderColour = 'lightgrey'))

  # Draw vertical lines
  openxlsx::addStyle(wb, sheet = sheet_name, cols = 1: ncol(df), rows = 2:(nrow(df)+1), gridExpand = TRUE, stack = TRUE,
                     style = openxlsx::createStyle(border = "LeftRight", borderColour = "#821123"))

  # Draw line when folder changes
  dirs <- dirname(df$File_Name)
  dirs[is.na(dirs) | dirs == "."] <- ""
  change_rows <- which(dirs[-1] != dirs[-length(dirs)]) + 1
  borderStyle <- openxlsx::createStyle(border = "top", borderStyle = "thin", borderColour = "#821123")

  for (row in change_rows) {
    openxlsx::addStyle(wb, sheet = sheet_name, style = borderStyle, gridExpand = TRUE, stack = TRUE,
                       rows = row + 1, cols = 1:ncol(df))
  }

  # Draw line on bottom
  last_row <- nrow(df) + 1
  bottomStyle <- openxlsx::createStyle(border = "bottom", borderStyle = "thin", borderColour = "#821123")
  openxlsx::addStyle(wb, sheet = sheet_name, style = bottomStyle,
                     rows = last_row, cols = 1:ncol(df),
                     gridExpand = TRUE, stack = TRUE)
}
