#' Create an archival directory of all files in a directory
#'
#' @param path_to_working_directory Character. Path to existing folder whose files are to be copied.
#' @param path_to_archive_directory Character. Path to new archival directory.
#' @param convert Logical, indicating whether to convert file formats or not.
#' @param overwrite Logical, indicating whether to overwrite existing files or not.
#'
#' @export
#'
#' @examples
#' \dontrun{
#' create_archive_from_directory()
#' }
#'

create_archive_from_directory <- function(path_to_working_directory,
                                          path_to_archive_directory,
                                          convert = TRUE,
                                          overwrite = FALSE){

  # create archival folder
  if(!dir.exists(path_to_archive_directory)){
    dir.create(path_to_archive_directory)
    message(paste("Info. New folder has been created:", path_to_archive_directory))
  }

  # gather file info
  df <- .get_file_info(path_to_working_directory, recursive = TRUE)

  # move complete folder
  cat(paste(" Creating AIP...\n\n",
            "- Copying from: ", path_to_working_directory, "\n",
            "- Copying to:   ", path_to_archive_directory, "\n",
            "- Files to copy:", nrow(df), "\n",
            "- Size (KB):    ", sum(df$Size_Bytes/1000), "\n\n"
  ))

  invisible(mapply(file.copy,
                   from = df$File_Name,
                   to = path_to_archive_directory,
                   copy.date = TRUE,
                   overwrite = overwrite))


  # write a report of what has been moved to where
  report <- data.frame(File_Name = basename(df$File_Name),
                       Dir_Origin = df$File_Name,
                       Dir_Archive = gsub(path_to_working_directory, path_to_archive_directory, df$File_Name),
                       Converted = FALSE)

  ## Converting files

  if(convert == TRUE){

  # xlsx --> csv
    df_xlsx <- df[grep("\\.xlsx?$", df$File_Name, ignore.case = TRUE),]
    if(nrow(df_xlsx) > 0){

      for (i in 1:nrow(df_xlsx)){
        csv_names <- .convert_xlsx_to_csv(xlsx_path = df_xlsx$File_Name[i],
                                          save_to = paste0(path_to_archive_directory, "/", df_xlsx$AIP[i]),
                                          overwrite = overwrite)
        report <- rbind(report,
                        data.frame(File_Name = rep(basename(df_xlsx$File_Name[i]), times = length(csv_names)),
                                   Dir_Origin = rep(df_xlsx$File_Name[i], times = length(csv_names)),
                                   Dir_Archive =  csv_names,
                                   Converted = rep(TRUE, times = length(csv_names)))
        )
      }
    }

  # eml --> txt
  # PDF --> PDFA
  # docx --> PDFA
  # sav --> csv

  }

  # write report
  report <- report[order(report$Dir_Archive),]
  row.names(report) <- NULL
  .write_csv_utf8_bom(df = report,
                      path = paste0(path_to_archive_directory, "/_archive_documentation.csv"), sep = ",", overwrite = overwrite)

  # print message
  cat(paste0(" AIP hast been created!\n",
             " Documentation: ", path_to_archive_directory, "/_archive_documentation.csv"))

  # invisibly return report
  invisible(report)
}


