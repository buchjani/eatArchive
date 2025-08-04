#' Write metadata summary of a directory to an Excel file
#'
#' Recursively scans a directory and writes file metadata (name, size, last modification time)
#' into an Excel workbook. Each top-level folder becomes a separate sheet, and an optional
#' "toplevel" sheet lists files directly in the base directory.
#'
#' Optionally, the function can add custom autofill rows to each sheet (e.g., for later
#' manual input), and exclude files from specific subfolders (e.g., "_Archive") from the output.
#'
#' @param base_dir Character. Path to the base directory to be scanned.
#' @param output_file Character. Path to the Excel file to be created.
#' @param autofill_values Optional character vector. If provided, these values will be inserted
#'        into the "Archive" column as autofill rows beneath the header row in each sheet. Each
#'        value becomes one row.
#' @param exclude_folders Character vector. Names of folders to exclude (e.g., "_Archive").
#'        Defaults to \code{c("_Archive")}. Set to \code{NULL} or \code{character(0)} to include all folders.
#'
#' @return Invisibly returns the workbook object (from the \pkg{openxlsx} package),
#'         and creates the Excel file as a side effect.
#'
#' @details
#' The Excel output includes:
#' \itemize{
#'   \item One sheet per top-level subdirectory (recursively listing files within)
#'   \item An optional "toplevel" sheet with files directly in the base directory
#'   \item Column widths and light styling for readability
#'   \item Optional autofill rows below the header in each sheet (for later manual tagging or categorization)
#' }
#'
#' Files located in any folder (or subfolder) whose name matches one of the entries in
#' \code{exclude_folders} will be excluded.
#'
#' @seealso \code{\link[openxlsx]{createWorkbook}}, \code{\link[fs]{dir_ls}}, \code{\link[base]{file.info}}
#'
#' @examples
#' \dontrun{
#' write_metadata("my_project/", "metadata_summary.xlsx")
#' write_metadata("my_project/", "summary.xlsx", autofill_values = c("keep", "delete"))
#' write_metadata("my_project/", "summary.xlsx", exclude_folders = NULL)  # include all
#' }
#'
#' @export

write_metadata <- function(base_dir,
                           output_file,
                           autofill_values = NULL,
                           exclude_folders = c("_Archive")) {
  # ----------------
  # HELPER FUNCTIONS
  # ----------------

  # helper function to get file info from folder


  # function for formatting & styling


  # -----------
  #  WRITE FILE
  # -----------

  # Write message in console
  cat(paste0("Scanning base directory: ", "\n", base_dir, "\n", "\n"))

  # Show progress
  n_folders <- length(fs::dir_ls(base_dir, type = "directory", recurse = FALSE))
  n_folders <- n_folders - sum(exclude_folders %in% dir(base_dir))


  # Create workbook
  wb <- openxlsx::createWorkbook()

  # Dataframe of files in Top-level folder
  df_baseDir <- get_file_info(dir = base_dir, recursive = FALSE, exclude_folders = exclude_folders)

  # Dataframe of files for each subdirectory
  subDirs <- fs::dir_ls(base_dir, type = "directory", recurse = FALSE)
  for (f in 1:length(subDirs)) {
    folder <- subDirs[f]
    df <- get_file_info(folder, recursive = TRUE, exclude_folders = exclude_folders)
    if (!is.null(df)) {
      # Print progress
      progress_count <- f
      cat(paste0("... Folder [", progress_count, "/", n_folders, "]: ", basename(folder), "\n"))
      # Sheet
      sheet_name <- make.names(fs::path_file(folder))
      sheet_name <- substr(sheet_name, 1, 31)
      add_sheet_with_style(wb, sheet_name, df, autofill_values)
    }
  }

  # Save workbook, write message
  cat("\nWriting file report to:\n")
  openxlsx::saveWorkbook(wb, output_file, overwrite = TRUE)
  cat(output_file)

}
