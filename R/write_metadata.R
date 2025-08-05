#' Create an Excel workbook with metadata for all subfolders of a directory
#'
#' Scans each immediate subfolder of the specified directory (excluding any listed in `exclude_folders`)
#' and writes an Excel file containing file-level metadata. Each subfolder gets its own sheet, listing
#' all files (recursively) with their name, size, and last modified timestamp.\cr \cr
#' Optionally, an additional column named "Archive" is included in each sheet. The values of "Archive"
#' will later be used by `create_archive()` to indicate the destination folder for copying or converting files.
#'
#' @param dir Character. Path to the directory whose immediate subfolders should be scanned.
#' @param output_file Character. Path to the Excel file to be written (e.g., "metadata.xlsx").
#' @param exclude_folders Character vector. Names of subfolders to exclude from processing.
#' @param autofill_values Optional. Controls whether and how the "Archive" column is included and pre-filled:
#'   - If `NULL`, the "Archive" column is not included.
#'   - If `NA`, a column "Archive" is included but autocomplete and filtering in Excel will not be available.
#'   - If a character vector is provided, the values are added to enable autocomplete and filtering in Excel,
#'     making it easier to assign target folders for file archiving.
#'
#' @details
#' The Excel workbook includes:
#' \itemize{
#'   \item One sheet per immediate subfolder of the specified directory
#'   \item Each sheet contains a recursive listing of all files in that subfolder (including sub-subfolders)
#'   \item Optional autofill rows inserted directly below the header (e.g., for manual annotation)
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
#' write_metadata("my_project/", "metadata_summary.xlsx")
#'
#' # Include an autofill row with values for later tagging
#' write_metadata("my_project/", "summary.xlsx", autofill_values = c("keep", "archive"))
#'
#' # Include all subfolders (do not exclude any)
#' write_metadata("my_project/", "summary.xlsx", exclude_folders = NULL)
#' }
#'
#' @export

write_metadata <- function(dir,
                           output_file,
                           exclude_folders = "_Archive",
                           autofill_values = NA
                           ) {

  # Write message in console
  cat(paste0("Scanning base directory: ", "\n", dir, "\n", "\n"))

  # Create workbook
  wb <- openxlsx::createWorkbook()

  # Toplevel files (non-recursive)
  toplevel_df <- get_file_info(dir = dir, recursive = FALSE)
  toplevel_name <- basename(normalizePath(dir))
  if (!is.null(toplevel_df)) {
    add_sheet_with_style(wb, sheet_name = toplevel_name, df = toplevel_df, autofill_values = autofill_values)
  }

  # Files in subfolders (one sheet per subfolder)
  subfolders <- fs::dir_ls(dir, type = "directory", recurse = FALSE)
  subfolders <- subfolders[!fs::path_file(subfolders) %in% exclude_folders]

  for (f in 1:length(subfolders)) {
    folder <- subfolders[f]
    df <- get_file_info(dir = folder, recursive = TRUE)
    if (!is.null(df)) {
      # Print progress
      cat(paste0("... Folder [", f, "/", length(subfolders), "]: ", basename(folder), " - ", "\n"))
      # Make sheet
      sheet_name <- make.names(fs::path_file(folder))
      sheet_name <- substr(sheet_name, 1, 31)
      add_sheet_with_style(wb, sheet_name = sheet_name, df = df, autofill_values = autofill_values)
    }
  }

  # Save the final workbook, write message
  cat("\nWriting metadata report to:\n")
  cat(output_file)
  openxlsx::saveWorkbook(wb, output_file, overwrite = TRUE)
}
