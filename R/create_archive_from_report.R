#' Create an archival directory based on Excel file
#'
#' @param path_to_directory_report Character. Path to Excel file resulting from `write_directory_report()` containing column "Archive", indicating whether and where to archive the file.
#' @param path_to_archive_directory Character. Path to ....
#' @param convert Logical, indicating whether to convert file formats or not ....
#' @param overwrite Logical, indicating whether to overwrite existing files ....
#'
#' @returns Folder and documentation of all files that have been copied...
#'
#' @export
#'
#' @examples
#'

create_archive_from_report <- function(path_to_directory_report,
                                       path_to_archive_directory,
                                       convert = TRUE,
                                       overwrite = TRUE){

  # combine sheets in dataframe
  df <- .combine_excel_sheets(path_to_directory_report)

  # check if column Archive exists - if not , return message that function doesnt kknow what to move
  if("Archive" %in% colnames(df) == FALSE){
    stop("Column 'Archive' is missing. Please use function `write_file_report()` and specify argument `autofill_values`.")
  }

  # keep only rows of files to be archived
  df <- df[!is.na(df$Archive) & !is.na(df$File_Name),]
  if(nrow(df) == 0){
    stop("Column 'Archive' is empty across all sheets of your Excel file. Looks like there's nothing to be archived.") # ¯\\_(ツ)_/¯
  }

  # create archival folder and subfolders as indicated
  if(!dir.exists(path_to_archive_directory)){
    dir.create(path_to_archive_directory)
    message(paste("Info. New folder has been created:", path_to_archive_directory))
  }

  subdirs <- paste0(path_to_archive_directory, "/", unique(df$Archive))
  invisible(lapply(subdirs, dir.create, recursive = TRUE, showWarnings = FALSE))

  # move files to folder indicated in column "Archive"
  cat(paste(" Creating archive...\n",
            "- Copying from: ", .longest_common_path(df$File_Name), "\n",
            "- Copying to:   ", path_to_archive_directory, "\n",
            "- Files to copy:", nrow(df), "\n",
            "- Size (KB):    ", sum(df$Size_Bytes/1000), "\n"
            ))

  invisible(mapply(file.copy,
                   from = df$File_Name,
                   to = paste0(path_to_archive_directory, "/", df$Archive),
                   copy.date = TRUE,
                   overwrite = overwrite))


  # write a report of what has been moved to where
  report <- data.frame(File_Name = basename(df$File_Name),
                       Dir_Origin = df$File_Name,
                       Dir_Archive =  paste0(path_to_archive_directory, "/", df$Archive, "/", basename(df$File_Name)),
                       Converted = FALSE)

  ## Converting files
  # xlsx --> csv
  if(convert == TRUE){

    df_xlsx <- df[grep("\\.xlsx?$", df$File_Name, ignore.case = TRUE),]
    if(nrow(df_xlsx) > 0){

      for (i in 1:nrow(df_xlsx)){
        csv_names <- .convert_xlsx_to_csv(xlsx_path = df_xlsx$File_Name[i],
                             save_to = paste0(path_to_archive_directory, "/", df_xlsx$Archive[i]),
                             overwrite = overwrite)
        report <- rbind(report,
                        data.frame(File_Name = rep(basename(df_xlsx$File_Name[i]), times = length(csv_names)),
                                   Dir_Origin = rep(df_xlsx$File_Name[i], times = length(csv_names)),
                                   Dir_Archive =  csv_names,
                                   Converted = rep(TRUE, times = length(csv_names)))
        )
      }
    }
  }

  # write report
  report <- report[order(report$Dir_Archive),]
  row.names(report) <- NULL
  utils::write.csv(x = report,
                   file = paste0(path_to_archive_directory, "/_archive_documentation.csv"), row.names=FALSE, quote = FALSE)

  # print message
  cat(paste0(" Archive hast been created!\n",
             " Documentation: ", path_to_archive_directory, "/_archive_documentation.csv"))

  # invisibly return report
  invisible(report)
}

# ## Check
# a <- create_archive_from_report(path_to_directory_report = "C:/R/files_report.xlsx",
#                               path_to_archive_directory = "C:/R/meinArchiv",
#                               overwrite = TRUE,
#                               convert = TRUE)
# a
