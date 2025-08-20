#' Combine tables from multiple excel sheets in directory report
#'
#' @param path_to_directory_report Character. Path to excel file resulting from `write_directory_report()` containing column
#' "Archive", indicating whether and where to archive the file.
#'
#' @keywords internal

.combine_excel_sheets <- function(path_to_directory_report) {

  stopifnot(file.exists(path_to_directory_report))
  sheet_names <- openxlsx::getSheetNames(path_to_directory_report)

  all_data <- list()

  for (sheet in sheet_names) {
    df <- openxlsx::read.xlsx(path_to_directory_report, sheet = sheet, detectDates = TRUE)

    if (ncol(df)>1){
      df$Last_Modified <- as.POSIXct((df$Last_Modified - 25569) * 86400, origin = "1970-01-01", tz = "UTC")
      all_data[[sheet]] <- df
    } else {
      df <- NULL # ignore sheets for empty folders
    }
  }

  combined_df <- do.call(rbind, all_data)
  rownames(combined_df) <- NULL
  return(combined_df)
}

### check
# all_sheets <- .combine_excel_sheets(path_to_directory_report = "C:/R/files_report.xlsx")
