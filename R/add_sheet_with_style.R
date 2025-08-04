add_sheet_with_style <- function(wb, sheet_name, df, autofill_values) {

  openxlsx::addWorksheet(wb, sheet_name)

  # Add "Archive"-column
  df$Archive <- NA

  # Insert autofill values (if provided)
  if (!is.null(autofill_values)) {
    autofill_df <- data.frame(
      File_Name = rep(NA, length(autofill_values)),
      Size_Bytes = rep(NA, length(autofill_values)),
      Last_Modified = rep(NA, length(autofill_values)),
      Archive = autofill_values
    )
    df <- rbind(autofill_df, df)
  }
  df$Last_Modified <- as.POSIXct(df$Last_Modified)

  # Write table
  openxlsx::writeDataTable(wb, sheet = sheet_name, x = df, withFilter = TRUE)

  # Set column widths and row heights (hidden rows)
  openxlsx::setColWidths(wb, sheet = sheet_name, cols = 1, widths = 65)  # File name
  openxlsx::setColWidths(wb, sheet = sheet_name, cols = 3, widths = 25)  # Last modified
  openxlsx::setColWidths(wb, sheet = sheet_name, cols = 4, widths = 20)  # Archive
  openxlsx::setRowHeights(wb, sheet = sheet_name, rows = 2:(length(autofill_values)+1), heights = 0)

  #  Format column 'Last_Modified' as date/time
  dateStyle <- openxlsx::createStyle(numFmt = "yyyy-mm-dd hh:mm:ss")
  openxlsx::addStyle(wb, sheet = sheet_name, style = dateStyle,
                     cols = 3, rows = 2:(nrow(df) + 1), gridExpand = TRUE, stack = TRUE)

  # Header
  headerStyle <- openxlsx::createStyle(fontSize = 12, fontColour = "#FFFFFF", halign = "left",
                                       fgFill = "#821123", border = "TopBottom", borderColour = "#821123")
  openxlsx::addStyle(wb, sheet = sheet_name, cols = 1: ncol(df), rows = 1, style = headerStyle,
                     gridExpand = TRUE, stack = TRUE)
  openxlsx::freezePane(wb, sheet = sheet_name, firstRow = TRUE)

  # remove default lines between rows
  openxlsx::addStyle(wb, sheet = sheet_name, cols = 1: ncol(df), rows = 2:nrow(df), gridExpand = TRUE, stack = TRUE,
                     style = openxlsx::createStyle(border = "TopBottom", borderColour = 'lightgrey'))

  # Vertical lines
  openxlsx::addStyle(wb, sheet = sheet_name, cols = 1: ncol(df), rows = 2:nrow(df)+1, gridExpand = TRUE, stack = TRUE,
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
