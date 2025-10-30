#' Create an archival directory based on Excel file
#'
#' @param path_to_directory_report Character. Path to Excel file resulting from `write_directory_report()` containing column "Archive", indicating whether and where to archive the file.
#' @param path_to_archive_directory Character. Path indicating where to create the archival directory.
#' @param convert Logical, indicating whether to convert file formats or not. Current files for conversion are:
#'   - xlsx --> csv: each sheet of an excel file is converted to a csv file
#'   - docx --> txt: word files are converted to txt files
#'   - doc --> txt: word files are converted to txt files
#' @param overwrite Logical, indicating whether to overwrite existing files in archival directory.
#'
#' @returns Folder and documentation of all files that have been copied.
#'
#' @export
#'
#' @examples
#'

create_archive_from_report <- function(path_to_directory_report,
                                       path_to_archive_directory,
                                       convert = TRUE,
                                       overwrite = TRUE){

  stopifnot(file.exists(path_to_directory_report))

  # combine sheets in dataframe
  df <- .combine_excel_sheets(path_to_directory_report)

  # check if column Archive exists - if not, return message that function doesnt kknow what to move
  if("Archive" %in% colnames(df) == FALSE){
    stop("Column 'Archive' is missing. Please use function `write_directory_report()` and specify argument `autocomplete_values`.")
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

  # MOVING FILES ####

  # move files to folder indicated in column "Archive"
  cat(paste(" Creating Archive...\n",
            "- Copying from: ", .longest_common_path(df$File_Name), "\n",
            "- Copying to:   ", path_to_archive_directory, "\n",
            "- Files to copy:", nrow(df), "\n",
            "- Size (KB):    ", sum(df$Size_Bytes/1000), "\n\n"
  ))

  invisible(mapply(file.copy,
                   from = df$File_Name,
                   to = paste0(path_to_archive_directory, "/", df$Archive),
                   copy.date = TRUE,
                   overwrite = overwrite))


  # write a report of what has been moved to where
  report <- data.frame(File_Name = basename(df$File_Name),
                       Last_Modified = as.POSIXct(df$Last_Modified),
                       Size_Bytes = df$Size_Bytes,
                       Converted = FALSE,
                       Dir_Archive =  paste0(path_to_archive_directory, "/", df$Archive, "/", basename(df$File_Name)),
                       Dir_Origin = df$File_Name
  )

  # CONVERTING FILES ####
  if(convert == TRUE){

    cat("\n Converting...\n")

    # xlsx --> csv ----

    df_xlsx <- df[grep("\\.xlsx?$", df$File_Name, ignore.case = TRUE),]
    if(nrow(df_xlsx) > 0){

      for (i in 1:nrow(df_xlsx)){
        csv_names <- .convert_xlsx_to_csv(xlsx_path = df_xlsx$File_Name[i],
                                          save_to = paste0(path_to_archive_directory, "/", df_xlsx$Archive[i]),
                                          overwrite = overwrite)
        report <- rbind(report,
                        data.frame(
                          File_Name = basename(csv_names),
                          Last_Modified = as.POSIXct(df_xlsx$Last_Modified[i]),
                          Size_Bytes = df_xlsx$Size_Bytes[i],
                          Converted = TRUE,
                          Dir_Archive = csv_names,
                          Dir_Origin = rep(df_xlsx$File_Name[i], times = length(csv_names)))
        )
      }
    }

    # sav --> csv ----


    # eml --> txt ----


    # doc --> pdfa ----


    # doc(x) --> txt ----

    df_docx <- df[grep("\\.docx?$", df$File_Name, ignore.case = TRUE),]
    if(nrow(df_docx) > 0){

      for (i in 1:nrow(df_docx)){
        txt_name <- .convert_docx_to_txt(docx_path = df_docx$File_Name[i],
                                         save_to = paste0(path_to_archive_directory, "/", df_docx$Archive[i]),
                                         overwrite = overwrite
        )
        report <- rbind(report,
                        data.frame(
                          File_Name = basename(txt_name),
                          Last_Modified = as.POSIXct(df_docx$Last_Modified[i]),
                          Size_Bytes = df_docx$Size_Bytes[i],
                          Converted = TRUE,
                          Dir_Archive = txt_name,
                          Dir_Origin = rep(df_docx$File_Name[i], times = length(txt_name)))
        )
      }
    }


    # pdf --> pdfa ----

  }

  # write report
  report <- report[order(report$Dir_Archive),]
  row.names(report) <- NULL
  .write_csv_utf8_bom(df = report,
                      path = paste0(path_to_archive_directory, "/_archive_documentation.csv"), sep = ",", overwrite = overwrite)

  # print message
  cat(paste0("\n",
             " Archive hast been created!\n",
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
