#' Create an Excel workbook with metadata for all subfolders of a directory
#'
#' Scans each immediate subfolder of the specified directory (excluding any listed in `exclude_folders`)
#' and writes an Excel file containing file-level metadata. Each subfolder gets its own sheet, listing
#' all files (recursively) with their name, size, and last modified timestamp.\cr \cr
#' Optionally, an additional column named "AIP" is included in each sheet. The values of "AIP"
#' will later be used by `create_archive()` to indicate the destination folder for copying or converting files.
#'
#' @param dir Character. Path to the directory whose immediate subfolders should be scanned.
#' @param path_to_directory_report Character. Path to the Excel file to be written (e.g., "C:/study/dir_report.xlsx").
#' @param exclude_folders Character vector. Names of subfolders to exclude from processing.
#' @param autocomplete_values Optional. Controls whether the "AIP" column is included and which (if any)
#' autocomplete values are to be provided:
#'   - If `NULL`, the "AIP" column is not included.
#'   - If `NA`, a column "AIP" is included but autocomplete and filtering in Excel will not be available.
#'   - If a character vector is provided, a column "AIP" is included and autocomplete and filtering in Excel is
#'     enabled, making it easier to assign target folders for file archiving.
#'
#' @details
#' The Excel workbook includes:
#' \itemize{
#'   \item One sheet per immediate subfolder of the specified directory
#'   \item Each sheet contains a recursive listing of all files in that subfolder (including sub-subfolders)
#'   \item Optional autocomplete rows inserted directly below the header (e.g., for manual annotation)
#'   \item Column widths and light styling for improved readability
#' }
#'
#' Files located in any subfolder whose name matches one of the entries in
#' \code{exclude_folders} will be excluded entirely (i.e., they are not scanned or listed).
#'
#' This function is intended as part of a file management workflow in which files may later be
#' reviewed, copied, or converted based on the generated metadata file.
#'
#' @return Invisibly returns the path to the written Excel file.
#'
#' @seealso \code{\link[openxlsx]{createWorkbook}}, \code{\link[fs]{dir_ls}}, \code{\link[base]{file.info}}
#'
#' @examples
#' \dontrun{
#' # Write metadata for all subfolders of "my_project/"
#' write_directory_report("my_project/", "metadata_summary.xlsx")
#'
#' # Include an autocomplete row with names of archival folders
#' write_directory_report("my_project/", "summary.xlsx", autocomplete_values = c("contracts", "data"))
#'
#' # Include all subfolders (do not exclude any)
#' write_directory_report("my_project/", "summary.xlsx", exclude_folders = NULL)
#' }
#'
#' @export

write_directory_report <- function(dir,
                           path_to_directory_report,
                           exclude_folders = "_Archive",
                           autocomplete_values = NA
                           ) {

  # Write message in console
  cat(paste0("Scanning base directory: ", "\n", dir, "\n", "\n"))

  # Create workbook
  wb <- openxlsx::createWorkbook()

  # Toplevel files (non-recursive)
  toplevel_df <- .get_file_info(dir = dir, recursive = FALSE)
  toplevel_name <- basename(normalizePath(dir))

  if (!is.null(toplevel_df) && nrow(toplevel_df) > 0) {
    .add_sheet_with_style(wb, sheet_name = toplevel_name, df = toplevel_df, autocomplete_values = autocomplete_values)
  } else if (is.null(toplevel_df) || nrow(toplevel_df) == 0) {
    toplevel_df <- data.frame("Info:" = paste0("Folder '", toplevel_name, "' contains no files."), stringsAsFactors = FALSE)
    openxlsx::addWorksheet(wb, toplevel_name)
    openxlsx::writeData(wb, sheet = toplevel_name, x = toplevel_df, startRow = 1, withFilter = F)
  }

  # Files in subfolders (one sheet per subfolder)
  subfolders <- fs::dir_ls(dir, type = "directory", recurse = FALSE)
  subfolders <- subfolders[!fs::path_file(subfolders) %in% exclude_folders]

  for (f in 1:length(subfolders)) {
    folder <- subfolders[f]
    folder_name <- basename(normalizePath(folder))

    # Print progress
    cat(paste0("... Folder [", f, "/", length(subfolders), "]: ", basename(folder), "\n"))

    df <- .get_file_info(dir = folder, recursive = TRUE)
    if (!is.null(df) && nrow(df) > 0) {
      sheet_name <- make.names(fs::path_file(folder))
      sheet_name <- ifelse(substr(sheet_name, 1,1)=="X" & substr(fs::path_file(folder), 1, 1) != "X", gsub('^X', '', sheet_name), sheet_name)
      sheet_name <- substr(sheet_name, 1, 31)
      .add_sheet_with_style(wb, sheet_name = sheet_name, df = df, autocomplete_values = autocomplete_values)
    } else if (is.null(df) || nrow(df) == 0) {
      df <- data.frame("Info" = paste0("Folder '", folder_name, "' contains no files."), stringsAsFactors = FALSE)
      openxlsx::addWorksheet(wb, basename(normalizePath(folder)))
      openxlsx::writeData(wb, sheet = folder_name, x = df, startRow = 1, withFilter = F)
    }
  }

  # Save the final workbook, write message
  cat("\nWriting directory report to:\n")
  cat(path_to_directory_report)
  openxlsx::saveWorkbook(wb, path_to_directory_report, overwrite = TRUE)
  invisible(path_to_directory_report)
}
