#' Format an excel sheet
#'
#' @param wb An excel workbook.
#' @param sheet_name Name of the excel sheet.
#' @param df A data.frame.
#' @param autocomplete_values A character vector of values which should be offered as autocomplete.
#'
#' @keywords internal

.add_sheet_with_style <- function(wb, sheet_name, df, autocomplete_values = NULL) {
  openxlsx::addWorksheet(wb, sheet_name)

  # Handle "Archive" column and optional autocomplete rows
  add_archive_col <- !is.null(autocomplete_values)
  add_autocomplete_rows <- is.character(autocomplete_values)

  if (add_archive_col && !"Archive" %in% names(df)) {
    df$Archive <- NA
  }

  if (add_autocomplete_rows) {
    autocomplete_df <- as.data.frame(matrix(NA, nrow = length(autocomplete_values), ncol = ncol(df)))
    names(autocomplete_df) <- names(df)
    autocomplete_df$Archive <- autocomplete_values
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
                                       fgFill = "#B40036", border = "TopBottom", borderColour = "#B40036")
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
                     style = openxlsx::createStyle(border = "LeftRight", borderColour = "#B40036"))

  # Draw line when folder changes
  dirs <- dirname(df$File_Name)
  dirs[is.na(dirs) | dirs == "."] <- ""
  change_rows <- which(dirs[-1] != dirs[-length(dirs)]) + 1
  borderStyle <- openxlsx::createStyle(border = "top", borderStyle = "thin", borderColour = "#B40036")

  for (row in change_rows) {
    openxlsx::addStyle(wb, sheet = sheet_name, style = borderStyle, gridExpand = TRUE, stack = TRUE,
                       rows = row + 1, cols = 1:ncol(df))
  }

  # Draw line on bottom
  last_row <- nrow(df) + 1
  bottomStyle <- openxlsx::createStyle(border = "bottom", borderStyle = "thin", borderColour = "#B40036")
  openxlsx::addStyle(wb, sheet = sheet_name, style = bottomStyle,
                     rows = last_row, cols = 1:ncol(df),
                     gridExpand = TRUE, stack = TRUE)
}

.add_sheet_overview <- function(wb, sheet_name, df, autocomplete_values = NULL) {
  openxlsx::addWorksheet(wb, sheet_name)

  # Add the data table to openxlsx object
  openxlsx::writeData(wb, sheet = sheet_name, x = df, startRow = 1, withFilter = TRUE)

  # Header
  headerStyle <- openxlsx::createStyle(fontSize = 12, fontColour = "#FFFFFF", halign = "left",
                                       fgFill = "#B40036", border = "TopBottom", borderColour = "#B40036")
  openxlsx::addStyle(wb, sheet = sheet_name, cols = 1: ncol(df), rows = 1, style = headerStyle,
                     gridExpand = TRUE, stack = TRUE)
  openxlsx::freezePane(wb, sheet = sheet_name, firstRow = TRUE)

  # Format column 'Size_Bytes': comma-separated thousands
  commastyle <- openxlsx::createStyle(numFmt = "COMMA")
  openxlsx::addStyle(wb, sheet = sheet_name, style = commastyle,
                     cols = 2, rows = 2:(nrow(df) + 2), gridExpand = TRUE, stack = TRUE)

  # Set column widths and styles
  openxlsx::setColWidths(wb, sheet = sheet_name, cols = 1, widths = "auto")  # Folder name
  openxlsx::setColWidths(wb, sheet = sheet_name, cols = 2, widths = "auto")  # Total size

  # Draw vertical lines
  openxlsx::addStyle(wb, sheet = sheet_name, cols = 1: ncol(df), rows = 2:(nrow(df)+1), gridExpand = TRUE, stack = TRUE,
                     style = openxlsx::createStyle(border = "LeftRight", borderColour = "#B40036"))

  # Draw line on bottom
  last_row <- nrow(df) + 1
  bottomStyle <- openxlsx::createStyle(border = "bottom", borderStyle = "thin", borderColour = "#B40036")
  openxlsx::addStyle(wb, sheet = sheet_name, style = bottomStyle,
                     rows = last_row, cols = 1:ncol(df),
                     gridExpand = TRUE, stack = TRUE)
}

.add_sheet_for_size <- function(wb, sheet_name, dir, df, show_max, subtitle) {

  ## Make sheet
  openxlsx::addWorksheet(wb, sheet_name)

  ## Prepare data
  df$Last_Modified <- as.POSIXct(df$Last_Modified)
  show_max <- ifelse(nrow(df) < show_max, nrow(df), show_max)
  df_full <- df
  df_sel <- df[1:show_max, ]

  ## Write data

  # Header and meta info
  openxlsx::writeData(wb, sheet_name, x = "Selection based on file size", startRow = 1, startCol = 1)
  openxlsx::writeData(wb, sheet_name, x = subtitle, startRow = 2, startCol = 1)

  openxlsx::writeData(wb, sheet_name, x = dir, startRow = 4, startCol = 1)
  openxlsx::writeData(wb, sheet_name, x = " - Number of files in directory:", startRow = 5, startCol = 1)
  openxlsx::writeData(wb, sheet_name, x = " - Total size of files in directory (Bytes):", startRow = 6, startCol = 1)
  openxlsx::writeData(wb, sheet_name, x = " - Number of files to show in this report:", startRow = 7, startCol = 1)
  openxlsx::writeData(wb, sheet_name, x = " - Total size of files in this report (Bytes):", startRow = 8, startCol = 1)
  openxlsx::writeData(wb, sheet_name, x = " - Date of this report:", startRow = 9, startCol = 1)

  openxlsx::writeData(wb, sheet_name, x = nrow(df_full), startRow = 5, startCol = 2)
  openxlsx::writeData(wb, sheet_name, x = sum(df_full$Size_Bytes, na.rm = T), startRow = 6, startCol = 2)
  openxlsx::writeData(wb, sheet_name, x = nrow(df_sel), startRow = 7, startCol = 2)
  openxlsx::writeData(wb, sheet_name, x = Sys.time(), startRow = 8, startCol = 2)  # date (...)
  openxlsx::writeData(wb, sheet_name, x = sum(df_sel$Size_Bytes, na.rm = T), startRow = 8, startCol = 2)


  # Table with meta info
  headerStyle <- openxlsx::createStyle(
    fontSize = 12, fontColour = "#FFFFFF", fgFill = "#B40036", textDecoration = "Bold",
    halign = "left", valign = "CENTER", border = "TopBottom", borderColour = "#B40036")
  openxlsx::writeData(wb, sheet_name, df_sel, startRow = 11, startCol = 1, headerStyle = headerStyle)


  ## Define Styles
  titleStyle <- openxlsx::createStyle(fontSize = 16,  fontColour = "#B40036", textDecoration = "Bold", halign = "LEFT")
  subtitleStyle <- openxlsx::createStyle(fontSize = 12, halign = "LEFT", textDecoration = "Bold")
  dateStyle <- openxlsx::createStyle(numFmt = "yyyy-mm-dd hh:mm:ss")
  commastyle <- openxlsx::createStyle(numFmt = "COMMA")

  ## Apply styles to the data

  # Header, meta info
  openxlsx::addStyle(wb, sheet = sheet_name, cols = 1, rows = 1, style = titleStyle, stack = TRUE)
  openxlsx::addStyle(wb, sheet = sheet_name, cols = 1, rows = 2:9, style = subtitleStyle, stack = TRUE)
  openxlsx::addStyle(wb, sheet_name, style = commastyle, rows = c(5, 6, 8), cols = 2, gridExpand = TRUE)

  openxlsx::addStyle(wb, sheet_name, style = dateStyle, rows = 9, cols = 2, gridExpand = TRUE)

  # Set column widths and styles
  openxlsx::setColWidths(wb, sheet = sheet_name, cols = 1, widths = "auto")  # File name
  openxlsx::setColWidths(wb, sheet = sheet_name, cols = 2, widths = 20)      # Size
  openxlsx::setColWidths(wb, sheet = sheet_name, cols = 3, widths = 20)      # Last modified

  # Format column 'Size_Bytes': comma-separated thousands
  openxlsx::addStyle(wb, sheet = sheet_name, style = commastyle,
                     cols = 2, rows = 11:(nrow(df) + 11), gridExpand = TRUE, stack = TRUE)

  # Format column 'Last_Modified' as date/time
  openxlsx::addStyle(wb, sheet = sheet_name, style = dateStyle,
                     cols = 3, rows = 11:(nrow(df) + 11), gridExpand = TRUE, stack = TRUE)

  # Remove default lines between rows
  openxlsx::addStyle(wb, sheet = sheet_name, cols = 1: ncol(df), rows = 2:nrow(df), gridExpand = TRUE, stack = TRUE,
                     style = openxlsx::createStyle(border = "TopBottom", borderColour = 'lightgrey'))

  # Draw vertical lines
  openxlsx::addStyle(wb, sheet = sheet_name, cols = 1: ncol(df), rows = 11:(nrow(df_sel)+11), gridExpand = TRUE, stack = TRUE,
                     style = openxlsx::createStyle(border = "LeftRight", borderColour = "#B40036"))

  # Draw line on bottom
  last_row <- nrow(df_sel) + 11
  bottomStyle <- openxlsx::createStyle(border = "bottom", borderStyle = "thin", borderColour = "#B40036")
  openxlsx::addStyle(wb, sheet = sheet_name, style = bottomStyle,
                     rows = last_row, cols = 1:ncol(df),
                     gridExpand = TRUE, stack = TRUE)
}

.add_sheet_for_date <- function(wb, sheet_name, dir, df, show_max, subtitle) {

  ## Make sheet
  openxlsx::addWorksheet(wb, sheet_name)

  ## Prepare data
  df$Last_Modified <- as.POSIXct(df$Last_Modified)
  show_max <- ifelse(nrow(df) < show_max, nrow(df), show_max)
  df_full <- df
  df_sel <- df[1:show_max, ]

  ## Write data

  # Header and meta info
  openxlsx::writeData(wb, sheet_name, x = "Selection based on date (last modified)", startRow = 1, startCol = 1)
  openxlsx::writeData(wb, sheet_name, x = subtitle, startRow = 2, startCol = 1)

  openxlsx::writeData(wb, sheet_name, x = dir, startRow = 4, startCol = 1)
  openxlsx::writeData(wb, sheet_name, x = " - Number of files in directory:", startRow = 5, startCol = 1)
  openxlsx::writeData(wb, sheet_name, x = " - Newest file in directory:", startRow = 6, startCol = 1)
  openxlsx::writeData(wb, sheet_name, x = " - Oldest file in directory:", startRow = 7, startCol = 1)
  openxlsx::writeData(wb, sheet_name, x = " - Number of files in this report (Bytes):", startRow = 8, startCol = 1)
  openxlsx::writeData(wb, sheet_name, x = " - Date of this report:", startRow = 9, startCol = 1)

  openxlsx::writeData(wb, sheet_name, x = min(df_full$Last_Modified, na.rm = T), startRow = 6, startCol = 2)     # row 7 is date (...)
  openxlsx::writeData(wb, sheet_name, x = max(df_full$Last_Modified, na.rm = T), startRow = 5, startCol = 2)     # row 6 is date (...)
  openxlsx::writeData(wb, sheet_name, x = Sys.time(), startRow = 8, startCol = 2)                                # row 9 is date (...)
  openxlsx::writeData(wb, sheet_name, x = nrow(df_full), startRow = 5, startCol = 2)
  openxlsx::writeData(wb, sheet_name, x = nrow(df_sel), startRow = 8, startCol = 2)

  # Table with meta info
  headerStyle <- openxlsx::createStyle(
    fontSize = 12, fontColour = "#FFFFFF", fgFill = "#B40036", textDecoration = "Bold",
    halign = "left", valign = "CENTER", border = "TopBottom", borderColour = "#B40036")
  openxlsx::writeData(wb, sheet_name, df_sel, startRow = 11, startCol = 1, headerStyle = headerStyle)


  ## Define Styles

  titleStyle <- openxlsx::createStyle(fontSize = 16,  fontColour = "#B40036", textDecoration = "Bold", halign = "LEFT")
  subtitleStyle <- openxlsx::createStyle(fontSize = 12, halign = "LEFT", textDecoration = "Bold")
  dateStyle <- openxlsx::createStyle(numFmt = "yyyy-mm-dd hh:mm:ss")
  commastyle <- openxlsx::createStyle(numFmt = "COMMA")

  ## Apply styles to the data

  # Header, meta info
  openxlsx::addStyle(wb, sheet = sheet_name, cols = 1, rows = 1, style = titleStyle, stack = TRUE)
  openxlsx::addStyle(wb, sheet = sheet_name, cols = 1, rows = 2:9, style = subtitleStyle, stack = TRUE)
  openxlsx::addStyle(wb, sheet_name, style = dateStyle, rows = c(6, 7, 9), cols = 2, gridExpand = TRUE)
  openxlsx::addStyle(wb, sheet_name, style = commastyle, rows = c(5, 8), cols = 2, gridExpand = TRUE)

  # Set column widths and styles
  openxlsx::setColWidths(wb, sheet = sheet_name, cols = 1, widths = "auto")  # File name
  openxlsx::setColWidths(wb, sheet = sheet_name, cols = 2, widths = 20)      # Size
  openxlsx::setColWidths(wb, sheet = sheet_name, cols = 3, widths = 20)      # Last modified

  # Format column 'Last_Modified' as date/time
  openxlsx::addStyle(wb, sheet = sheet_name, style = dateStyle,
                     cols = 3, rows = 11:(nrow(df) + 11), gridExpand = TRUE, stack = TRUE)

  # Format column 'Size_Bytes': comma-separated thousands
  openxlsx::addStyle(wb, sheet = sheet_name, style = commastyle,
                     cols = 2, rows = 11:(nrow(df) + 11), gridExpand = TRUE, stack = TRUE)

  # Remove default lines between rows
  openxlsx::addStyle(wb, sheet = sheet_name, cols = 1: ncol(df), rows = 2:nrow(df), gridExpand = TRUE, stack = TRUE,
                     style = openxlsx::createStyle(border = "TopBottom", borderColour = 'lightgrey'))

  # Draw vertical lines
  openxlsx::addStyle(wb, sheet = sheet_name, cols = 1: ncol(df), rows = 11:(nrow(df_sel)+11), gridExpand = TRUE, stack = TRUE,
                     style = openxlsx::createStyle(border = "LeftRight", borderColour = "#B40036"))

  # Draw line on bottom
  last_row <- nrow(df_sel) + 11
  bottomStyle <- openxlsx::createStyle(border = "bottom", borderStyle = "thin", borderColour = "#B40036")
  openxlsx::addStyle(wb, sheet = sheet_name, style = bottomStyle,
                     rows = last_row, cols = 1:ncol(df),
                     gridExpand = TRUE, stack = TRUE)
}


