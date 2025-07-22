#' Title
#'
#' @param base_dir
#' @param output_file
#' @param autofill_values
#' @param exclude_folders
#'
#' @returns
#' @export
#'
#' @examples


writeDirectory <- function(base_dir,
                           output_file,
                           autofill_values = NULL,
                           exclude_folders = c("_Archive")) {
  # ----------------
  # HELPER FUNCTIONS
  # ----------------

  # helper function to get file info from folder
  get_file_info <- function(path,
                            relative_to = base_dir,
                            recursive = FALSE,
                            exclude_folders = c("_Archive")) {

    files <- fs::dir_ls(path, recurse = recursive, type = "file")

    # filter out files in any excluded folders
    if (!is.null(exclude_folders) && length(exclude_folders) > 0) {
      pattern <- paste0("/", exclude_folders, "/")
      combined_pattern <- paste(pattern, collapse = "|")
      files <- files[!grepl(combined_pattern, files)]
    }

    if (length(files) == 0) return(NULL)

    data.frame(
      File_Name = fs::path_rel(files, start = relative_to),
      Size_Bytes = fs::file_info(files)$size,
      Last_Modified = as.POSIXct(fs::file_info(files)$modification_time)
    )
  }

  # function for formatting & styling
  add_sheet_with_style <- function(wb, sheet_name, df, autofill_values) {

    openxlsx::addWorksheet(wb, sheet_name)

    # Add "Archive"-column
    df$Archive <- NA

    # Insert autofill values (if provided)
    if (!is.null(autofill_values)) {
      autofill_df <- tibble::tibble(
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

  # -----------
  #  WRITE FILE
  # -----------

  # Create workbook
  wb <- openxlsx::createWorkbook()

  # Dataframes of files to write
  df_baseDir <- get_file_info(base_dir, recursive = FALSE, exclude_folders = exclude_folders)
  subDirs    <- fs::dir_ls(base_dir, type = "directory", recurse = FALSE)

  # Write message in console
  cat(paste0("Scanning base directory: ", "\n", base_dir, "\n", "\n"))

  # Show progress
  n_folders <- length(fs::dir_ls(base_dir, type = "directory", recurse = FALSE))
  n_folders <- n_folders - sum(exclude_folders %in% dir(base_dir))

  # Loop through all subfolders ("sub-directories") to get file info
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
