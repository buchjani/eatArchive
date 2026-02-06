#' Search a directory for specific files (date, size)
#'
#' Scans each immediate subfolder of the specified directory (excluding any listed in `exclude_folders`)
#' and writes an Excel file containing file-level metadata for specific files. Files can be selected by
#' size and by date of last modification.
#'
#' @param dir Character. Path to the directory whose immediate subfolders should be scanned.
#' @param path_to_report Character. Path to the Excel file to be written (e.g., "C:/study/dir_report.xlsx").
#' @param show_what Character, indicating whehter to sort by size or date.
#' \itemize{
#'   \item `"largest"` for sorting by file size (large to small)
#'   \item `"smallest"` for sorting by file size (small to large)
#'   \item `"newest"` for sorting by date of last modification (new to old)
#'   \item `"oldest"` for sorting by date of last modification (old to new)
#' }
#' @param show_max Numeric, indicating the maximum number of files to include in the report.
#' @param exclude_folders Character vector. Names of subfolders to exclude from processing.
#' @param overwrite Logical, indicating whether to overwrite existing report.
#'
#' @details
#' The Excel workbook includes a recursive listing of all files meeting the search criteria (including sub-subfolders).
#'
#' Files located in any subfolder whose name matches one of the entries in
#' \code{exclude_folders} will be excluded entirely (i.e., they are not scanned or listed).
#'
#' @return Invisibly returns the path to the written Excel file.
#'
#' @export

scan_directory <- function(dir,
                           path_to_report,
                           exclude_folders = "",
                           show_what = c("largest", "smallest", "newest", "oldest"),
                           show_max = 1000,
                           overwrite = FALSE) {

  # Check if report file can be created
  if(dir.exists(path_to_report) & overwrite == FALSE){
    stop(paste0("\nThe file you indicated already exists:\n",
                path_to_report,
                "\nDelete this file or set 'overwrite = TRUE'"))
  }

  # Write message in console
  cat(paste0("Scanning base directory: ", "\n", dir, "\n", "\n"))

  # Create workbook
  wb <- openxlsx::createWorkbook()

  # Toplevel files (non-recursive)
  toplevel_df <- .get_file_info(dir = dir, recursive = FALSE)
  overview_df <- data.frame("Folder" = dir, "Size_Bytes" = sum(toplevel_df$Size_Bytes, na.rm = T))

  # Files in subfolders (one sheet per subfolder)
  subfolders <- fs::dir_ls(dir, type = "directory", recurse = FALSE)
  subfolders <- subfolders[!fs::path_file(subfolders) %in% exclude_folders]
  subfolders_df <- data.frame()

  for (f in 1:length(subfolders)) {
    folder <- subfolders[f]
    folder_name <- basename(normalizePath(folder))

    # Print progress
    cat(paste0("- Folder [", f, "/", length(subfolders), "]: ", basename(folder), "\n"))

    # collect data
    df <- .get_file_info(dir = folder, recursive = TRUE)

    # Only keep  subdirectories - but without those mentioned in exclude_folders (recursively!)
    if (!is.null(df) && nrow(df) > 0 && length(exclude_folders) > 0) {
      parts <- fs::path_split(df$File_Name)
      keep  <- !vapply(parts, function(x) any(x %in% exclude_folders), logical(1))
      df    <- df[keep, , drop = FALSE]
    }
    subfolders_df <- rbind(subfolders_df, df)
    overview_df <-  rbind(overview_df,
                          data.frame("Folder" = folder, "Size_Bytes" = sum(df$Size_Bytes, na.rm = T)))
    }

  df <- rbind(toplevel_df, subfolders_df)

  # write overview sheet
  .add_sheet_overview(wb = wb, sheet_name = "SUMMARY", df = overview_df)

  # write one sheet for each search criterion
  if("largest" %in% tolower(show_what)){
    df_sorted <- df[order(df$Size_Bytes, decreasing = TRUE),]
    .add_sheet_for_size(wb = wb, sheet_name = "Size (Largest)", dir = dir, df = df_sorted, show_max = show_max, subtitle = "(Large to small)")
  }
  if("smallest" %in% tolower(show_what)){
    df_sorted <- df[order(df$Size_Bytes, decreasing = FALSE),]
    .add_sheet_for_size(wb = wb, sheet_name = "Size (Smallest)", dir = dir, df = df_sorted, show_max = show_max, subtitle = "(Small to large)")
  }
  if("newest" %in% tolower(show_what)){
    df_sorted <- df[order(df$Last_Modified, decreasing = TRUE),]
    .add_sheet_for_date(wb = wb, sheet_name = "Date (Newest)", dir = dir, df = df_sorted, show_max = show_max, subtitle = "(New to old)")
  }
  if("oldest" %in% tolower(show_what)){
    df_sorted <- df[order(df$Last_Modified, decreasing = FALSE),]
    .add_sheet_for_date(wb = wb, sheet_name = "Date (Oldest)", dir = dir, df = df_sorted, show_max = show_max, subtitle = "(Old to new)")
  }

  # Save the final workbook, write message
  cat("\nWriting directory report to:\n")
  cat(path_to_report)
  openxlsx::saveWorkbook(wb, path_to_report, overwrite = TRUE)
  invisible(path_to_report)
}

# Check:
# scan_directory(
#   dir = "Q:/FDZ/Alle/20_Veranstaltungen",
#   path_to_report = "Q:/FDZ/Alle/20_Veranstaltungen/Ordner-Bericht-Auswahl.xlsx",
#   # exclude_folders = c("_Archive", "_Archiv"),
#   show_what = c("Largest", "SMALLEST", "neWest", "OLDEst"),
#   show_max = 100,
#   overwrite = TRUE
# )
