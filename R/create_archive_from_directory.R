#' Create an archival directory of all files in a directory
#'
#' @param path_to_working_directory Character. Path to working directory which is to be archived.
#' @param path_to_archive_directory Character. Path indicating where to create the archival directory.
#' @param convert Logical, indicating whether to convert file formats or not. Current files for conversion are:
#'   - xlsx --> csv: each sheet of an excel file is converted to a csv file
#'   - xlsm --> csv: each sheet of an excel macro file is converted to a csv file
#'   - sav  --> csv: SPSS data files are converted to two csv files each (data, metadata)
#'   - eml  --> txt: Thunderbird email objects (eml-files) are converted to txt-files
#'   - docx --> txt: word files are converted to txt files
#'   - doc --> txt: word files are converted to txt files
#' @param overwrite Logical, indicating whether to overwrite existing files in archival directory.
#' @param exclude_folders Character vector. Names of subfolders to exclude from processing.
#' @param csv Character specifying the csv type
#'  - "csv" uses "." for the decimal point and "," for the separator
#'  - "csv2" uses "," for the decimal point and ";" for the separator, the Excel convention for CSV files in some Western European locales.
#' @param pdf_flavor Character, indicating the particular flavor of PDF files to be created (e.g., "2b")
#'
#' @returns Folder and documentation of all files that have been copied.
#'
#' @export
#'
#' @examples

create_archive_from_directory <- function(path_to_working_directory,
                                          path_to_archive_directory,
                                          exclude_folders = "_Archive",
                                          convert = TRUE,
                                          overwrite = FALSE,
                                          csv = "csv",
                                          pdf_flavor = "2b"
                                          ){

  # INPUT VALIDATION ----

  # Check if directory to copy exists
  stopifnot(dir.exists(path_to_working_directory))

  # Check if archive folder can be created
  if(dir.exists(path_to_archive_directory) & overwrite == FALSE){
    stop(paste0("\nThe archive directory you indicated already exists:\n",
                path_to_archive_directory,
                "\nDelete this folder or set 'overwrite = TRUE'"))
  }

  # if PDF conversion is requested, check if veraPDF is installed
  verapdf_path <- .find_veraPDF()

  # Define csv type
  sep <- ifelse(csv == "csv2", ";", ",")
  dec <- ifelse(csv == "csv2", ",", ".")


  # SCAN WORKING DIRECTORY ----

  dir <- path_to_working_directory

  # Write message in console
  cat(paste0(" Scanning directory:", "\n ", dir, "\n", "\n"))

  # Toplevel files
  df <- .get_file_info(dir = dir, recursive = FALSE)

  # Files in subdirectories - but without those mentioned in exclude_folders (recursively!)
  subdirs <- fs::dir_ls(dir, type = "directory", recurse = TRUE, all = TRUE, fail = FALSE)
  parts <- fs::path_split(subdirs)
  keep  <- !vapply(parts, function(x) any(x %in% exclude_folders), logical(1))
  subdirs <- subdirs[keep]

  for (f in 1:length(subdirs)) {
    folder <- subdirs[f]
    folder_name <- gsub(paste0(path_to_working_directory, "/"), "", subdirs)[f]

    # Print progress
    cat(paste0(" - Folder [", f, "/", length(subdirs), "]: ", folder_name, "\n"))

    df_tmp <- .get_file_info(dir = folder, recursive = FALSE)
    df <- rbind(df, df_tmp)
  }
  cat("\n")

  # CREATE ARCHIVE DIRECTORY ----

  # create archival folder and subfolders as indicated
  if(!dir.exists(path_to_archive_directory)){
    dir.create(path_to_archive_directory)
    message(paste("Info. New folder has been created:", path_to_archive_directory))
  }

  subfolders <- gsub(path_to_working_directory, path_to_archive_directory, subdirs)

  # fix Umlaute in folder names - but only those that are to be copied. The rest needs to remain identical.
  subfolders_fixed <- .fix_umlaut(gsub(path_to_archive_directory, "", subfolders))
  subfolders <- paste0(path_to_archive_directory, subfolders_fixed)

  # create folders with fixed names (Umlaute)
  invisible(lapply(subfolders, dir.create, recursive = TRUE, showWarnings = FALSE))


  # COPYING FILES ----

  # print message
  cat(paste(" Creating Archive...\n",
            "- Copying from: ", path_to_working_directory, "\n",
            "- Copying to:   ", path_to_archive_directory, "\n",
            "- Files to copy:", nrow(df), "\n",
            "- Size (KB):    ", sum(df$Size_Bytes/1000), "\n"
  ))

  # move files to respective archive folder
  df$File_Name_Archive <- gsub(path_to_working_directory, path_to_archive_directory, .fix_umlaut(df$File_Name))
  invisible(mapply(file.copy,
                   from = df$File_Name[1],
                   to   = df$File_Name_Archive[1],
                   copy.date = TRUE,
                   overwrite = overwrite,
                   copy.mode = FALSE))


  # write a report of what has been moved to where
  report <- data.frame(File_Name = basename(df$File_Name),
                       Last_Modified = as.POSIXlt(df$Last_Modified),
                       Size_Bytes = df$Size_Bytes,
                       Status = "original",
                       Dir_Archive =  df$File_Name_Archive,
                       Dir_Origin = df$File_Name
  )


  # CONVERTING FILES ----

  if(convert == TRUE){

    cat("\n Converting...\n")


    # pdf --> pdfa ----

    # check #1: are there any PDF files in the first place?
    df_pdf_all <- df[grep("\\.pdf?$", df$File_Name, ignore.case = TRUE),]

    # check #2: is veraPDF installed?
    veraPDF_installed <- suppressMessages(ifelse(!is.na(.find_veraPDF()), TRUE, FALSE))

    # --> outcome #A (bad):
    # there are (any) PDF files, but veraPDF is not installed
    # action: write a message
    if(nrow(df_pdf_all) > 0 & veraPDF_installed == FALSE){
      cat(paste0(" - pdf  --> pdf/a (n = ", nrow(df_pdf), "): Conversion not possible."))

      message(
        "veraPDF was not found. PDF/A validation and conversion will be skipped.\n",
        "Install veraPDF from https://verapdf.org/software/ and try again.\n",
        "Already installed? Run `eatArchive::configure_veraPDF()` to select the folder."
      )
    }

    # --> outcome #B (good):
    # there are PDF files to convert, and veraPDF is installed
    # action: check pdf flavor to see if conversion is needed
    if(nrow(df_pdf_all) > 0 & veraPDF_installed == TRUE){
      for (p in 1:nrow(df_pdf_all)){
        df_pdf_all$flavor[p] <- .check_pdf_flavor(file = df_pdf_all$File_Name[p],
                                                  flavor = pdf_flavor,
                                                  veraPDF_path = .find_veraPDF())
      }
      # collect files with "false" flavour --> will be converted; those with "true" (i.e., correct) flavor are ignored
      df_pdf <- df_pdf_all[df_pdf_all$flavor == FALSE,]

      # check #3:
      # are there any files that do not conform the desired flavor (e.g., "2b")
      # action: convert those files

      if(nrow(df_pdf) > 0){
        cat(paste0(" - pdf  --> pdf/a (n = ", nrow(df_pdf), ")\n"))

        for (i in 1:nrow(df_pdf)){

          pdf_name <- .convert_pdf_to_pdfa(pdf_path = df_pdf$File_Name[i],
                                           save_to = paste0(path_to_archive_directory, "/", df_pdf$Archive[i]),
                                           pdf_flavor = pdf_flavor)

          # in case of failed conversion, pdf_name will be NA
          # --> add error message to report

          # in case of failed conversion:
          if(is.na(pdf_name)){
            report <- rbind(report,
                            data.frame(
                              File_Name = basename(df_pdf$File_Name[i]),
                              Last_Modified = as.POSIXct(df_pdf$Last_Modified[i]),
                              Size_Bytes = df_pdf$Size_Bytes[i],
                              Status = "requires_manual_conversion",
                              Dir_Archive = df_pdf$File_Name_Archive[i],
                              Dir_Origin = df_pdf$File_Name[i]))
            # in case of successful conversion:
          } else {
            report <- rbind(report,
                            data.frame(
                              File_Name = basename(pdf_name),
                              Last_Modified = as.POSIXct(df_pdf$Last_Modified[i]),
                              Size_Bytes = df_pdf$Size_Bytes[i],
                              Status = paste0("converted (PDF/A-", pdf_flavor, ")"),
                              Dir_Archive = pdf_name,
                              Dir_Origin = df_pdf$File_Name[i]))
          }
        }
      }

    }


    ## xls --> csv ----

    df_xls <- df[grep("\\.xls$", df$File_Name, ignore.case = TRUE),]
    if(nrow(df_xls) > 0){

      cat(paste0(" - xls  --> csv (n = ", nrow(df_xls), ")\n"))

      for (i in 1:nrow(df_xls)){
        csv_names <- tryCatch(
          .convert_xls_to_csv(xls_path = df_xls$File_Name[i],
                              save_to = paste0(path_to_archive_directory, "/", df_xls$Archive[i]),
                              csv = csv),
        error = function(e) character(0))

# in case of conversion problem, csv_names will be empty
# --> add error message to report

if (length(csv_names) == 1) {   # conversion successful
  report <- rbind(report,
                  data.frame(
                    File_Name = basename(csv_names),
                    Last_Modified = as.POSIXct(df_xls$Last_Modified[i]),
                    Size_Bytes = df_xls$Size_Bytes[i],
                    Status = "converted",
                    Dir_Archive = csv_names,
                    Dir_Origin = rep(df_xls$File_Name[i], times = length(csv_names))))
} else {                        # failed conversion
  report <- rbind(report,
                  data.frame(
                    File_Name = basename(df_xls$File_Name[i]),
                    Last_Modified = as.POSIXct(df_xls$Last_Modified[i]),
                    Size_Bytes = df_xls$Size_Bytes[i],
                    Status = "requires_manual_conversion",
                    Dir_Archive = paste0(path_to_archive_directory,"/",df_xls$Archive[i],"/",basename(df_xls$File_Name[i])),
                    Dir_Origin = df_xls$File_Name[i]))
}
    }
  }


    ## xlsx --> csv ----

    df_xlsx <- df[grep("\\.xlsx$", df$File_Name, ignore.case = TRUE),]
    if(nrow(df_xlsx) > 0){

      cat(paste0(" - xlsx --> csv (n = ", nrow(df_xlsx), ")\n"))

      for (i in 1:nrow(df_xlsx)){
        csv_names <- tryCatch(
          .convert_xlsx_to_csv(xlsx_path = df_xlsx$File_Name[i],
                               save_to = paste0(path_to_archive_directory, "/", df_xlsx$Archive[i]),
                               csv = csv),
          error = function(e) character(0))

        # in case of conversion problem, csv_names will be empty
        # --> add error message to report

        if (length(csv_names) == 1) {   # conversion successful
          report <- rbind(report,
                          data.frame(
                            File_Name = basename(csv_names),
                            Last_Modified = as.POSIXct(df_xlsx$Last_Modified[i]),
                            Size_Bytes = df_xlsx$Size_Bytes[i],
                            Status = "converted",
                            Dir_Archive = csv_names,
                            Dir_Origin = rep(df_xlsx$File_Name[i], times = length(csv_names))))
        } else {                        # failed conversion
          report <- rbind(report,
                          data.frame(
                            File_Name = basename(df_xlsx$File_Name[i]),
                            Last_Modified = as.POSIXct(df_xlsx$Last_Modified[i]),
                            Size_Bytes = df_xlsx$Size_Bytes[i],
                            Status = "requires_manual_conversion",
                            Dir_Archive = paste0(path_to_archive_directory,"/",df_xlsx$Archive[i],"/",basename(df_xlsx$File_Name[i])),
                            Dir_Origin = df_xlsx$File_Name[i]))
        }
      }
    }
    ## xlsm --> csv ----

    df_xlsm <- df[grep("\\.xlsm$", df$File_Name, ignore.case = TRUE),]
    if(nrow(df_xlsm) > 0){

      cat(paste0(" - xlsm --> csv (n = ", nrow(df_xlsm), ")\n"))

      for (i in 1:nrow(df_xlsm)){
        csv_names <- .convert_xlsm_to_csv(xlsm_path = df_xlsm$File_Name[i],
                                          save_to = dirname(df_xlsm$File_Name_Archive[i]),
                                          csv = csv)

        # in case of buggy (corrupted) xlsm-file, csv_names will be empty
        # --> add error message to report

        if(length(csv_names) == 0){
          report <- rbind(report,
                          data.frame(
                            File_Name = basename(df_xlsm$File_Name[i]),
                            Last_Modified = as.POSIXct(df_xlsm$Last_Modified[i]),
                            Size_Bytes = df_xlsm$Size_Bytes[i],
                            Status = "requires_manual_conversion",
                            Dir_Archive = df_xlsm$File_Name_Archive[i],
                            Dir_Origin = df_xlsm$File_Name[i]))
        } else {
          report <- rbind(report,
                          data.frame(
                            File_Name = basename(csv_names),
                            Last_Modified = as.POSIXct(df_xlsm$Last_Modified[i]),
                            Size_Bytes = df_xlsm$Size_Bytes[i],
                            Status = "converted",
                            Dir_Archive = csv_names,
                            Dir_Origin = rep(df_xlsm$File_Name[i], times = length(csv_names)))
          )
        }
      }
    }

    ## sav --> csv ----

    df_sav <- df[grep("\\.sav?$", df$File_Name, ignore.case = TRUE),]
    if(nrow(df_sav) > 0){

      cat(paste0(" - sav  --> csv (n = ", nrow(df_sav), ")\n"))

      for (i in 1:nrow(df_sav)) {

        csv_names <- tryCatch(
          .convert_sav_to_csv(sav_path = df_sav$File_Name[i],
                              save_to = paste0(path_to_archive_directory, "/", df_sav$Archive[i]),
                              csv = csv),
          error = function(e) character(0))

        # in case of conversion problem, csv_names will be empty
        # --> add error message to report

        if (length(csv_names) == 0) {
          report <- rbind(
            report,
            data.frame(File_Name = basename(df_sav$File_Name[i]),
                       Last_Modified = as.POSIXct(df_sav$Last_Modified[i]),
                       Size_Bytes = df_sav$Size_Bytes[i],
                       Status = "requires_manual_conversion",
                       Dir_Archive = paste0(path_to_archive_directory,"/",df_sav$Archive[i],"/",basename(df_sav$File_Name[i])),
                       Dir_Origin = df_sav$File_Name[i]))
        } else {
          report <- rbind(
            report,
            data.frame(File_Name = basename(csv_names),
                       Last_Modified = as.POSIXct(df_sav$Last_Modified[i]),
                       Size_Bytes = df_sav$Size_Bytes[i],
                       Status = "converted",
                       Dir_Archive = csv_names,
                       Dir_Origin = rep(df_sav$File_Name[i], times = length(csv_names)))
          )
        }
      }
    }

    ## eml --> txt ----

    df_eml <- df[grep("\\.eml?$", df$File_Name, ignore.case = TRUE),]
    if(nrow(df_eml) > 0){

      cat(paste0(" - eml  --> txt (n = ", nrow(df_eml), ")\n"))

      for (i in 1:nrow(df_eml)){
        txt_name <- tryCatch(
          .convert_eml_to_txt(eml_path = df_eml$File_Name[i],
                              save_to = paste0(path_to_archive_directory, "/", df_eml$Archive[i])),
          error = function(e) character(0))

        # in case of conversion problem, txt_name will be empty
        # --> add error message to report

        if (length(csv_names) == 1) {   # conversion successful
          report <- rbind(report,
                          data.frame(
                            File_Name = basename(txt_name),
                            Last_Modified = as.POSIXct(df_eml$Last_Modified[i]),
                            Size_Bytes = df_eml$Size_Bytes[i],
                            Status = "converted",
                            Dir_Archive = txt_name,
                            Dir_Origin = df_eml$File_Name[i]))
        } else {                        # failed conversion
          report <- rbind(report,
                          data.frame(
                            File_Name = basename(df_eml$File_Name[i]),
                            Last_Modified = as.POSIXct(df_eml$Last_Modified[i]),
                            Size_Bytes = df_eml$Size_Bytes[i],
                            Status = "requires_manual_conversion",
                            Dir_Archive = paste0(path_to_archive_directory,"/",df_eml$Archive[i],"/",basename(df_eml$File_Name[i])),
                            Dir_Origin = df_eml$File_Name[i]))
        }
      }
    }


    ## doc(x) --> txt ----

    df_docx <- df[grep("\\.docx?$", df$File_Name, ignore.case = TRUE),]
    if(nrow(df_docx) > 0){

      cat(paste0(" - docx --> txt (n = ", nrow(df_docx), ")\n"))

      for (i in 1:nrow(df_docx)){
        txt_name <- tryCatch(
          .convert_docx_to_txt(docx_path = df_docx$File_Name[i],
                               save_to = paste0(path_to_archive_directory, "/", df_docx$Archive[i])),
          error = function(e) character(0))

        # in case of conversion problem, txt_name will be empty
        # --> add error message to report

        if (length(txt_name) == 1) {    # conversion successful
          report <- rbind(report,
                          data.frame(
                            File_Name = basename(txt_name),
                            Last_Modified = as.POSIXct(df_docx$Last_Modified[i]),
                            Size_Bytes = df_docx$Size_Bytes[i],
                            Status = "converted",
                            Dir_Archive = txt_name,
                            Dir_Origin = df_docx$File_Name[i]))
        } else {                         # conversion problem
          report <- rbind(report,
                          data.frame(
                            File_Name = basename(df_docx$File_Name[i]),
                            Last_Modified = as.POSIXct(df_docx$Last_Modified[i]),
                            Size_Bytes = df_docx$Size_Bytes[i],
                            Status = "requires_manual_conversion",
                            Dir_Archive = txt_name,
                            Dir_Origin = df_docx$File_Name[i]))
        }
      }
    }

    ## docx --> zip ----

    df_docx <- df[grep("\\.docx$", df$File_Name, ignore.case = TRUE),]
    if(nrow(df_docx) > 0){

      cat(paste0(" - docx --> txt (n = ", nrow(df_docx), ")\n"))

      for (i in 1:nrow(df_docx)){
        zip_name <- tryCatch(
          .convert_docx_to_zip(docx_path = df_docx$File_Name[i],
                               save_to = paste0(path_to_archive_directory, "/", df_docx$Archive[i])),
          error = function(e) character(0))

        # in case of conversion problem, zip_name will be empty
        # --> add error message to report

        if (length(zip_name) == 1) {    # conversion successful
          report <- rbind(report,
                          data.frame(
                            File_Name = basename(zip_name),
                            Last_Modified = as.POSIXct(df_docx$Last_Modified[i]),
                            Size_Bytes = df_docx$Size_Bytes[i],
                            Status = "converted",
                            Dir_Archive = zip_name,
                            Dir_Origin = df_docx$File_Name[i]))
        } else {                         # conversion problem
          report <- rbind(report,
                          data.frame(
                            File_Name = basename(df_docx$File_Name[i]),
                            Last_Modified = as.POSIXct(df_docx$Last_Modified[i]),
                            Size_Bytes = df_docx$Size_Bytes[i],
                            Status = "requires_manual_conversion",
                            Dir_Archive = zip_name,
                            Dir_Origin = df_docx$File_Name[i]))
        }
      }
    }
  }

  # WRITE DOCUMENTATION ---------

  report <- report[order(report$Dir_Archive),]
  report$Last_Modified <-  format(report$Last_Modified, "%Y-%m-%d %H:%M:%S", tz = "Europe/Berlin")
  row.names(report) <- NULL
  .write_csv_utf8_bom(df = report,
                      path = paste0(path_to_archive_directory, "/_archive_documentation.csv"), sep = sep, dec = dec, overwrite = overwrite)

  # print message
  cat(paste0("\n",
             " Archive hast been created!\n",
             " Documentation: ", path_to_archive_directory, "/_archive_documentation.csv"))

  # invisibly return report
  invisible(report)
}

## Check:
# create_archive_from_directory(
#   path_to_working_directory = "Q:/FDZ/Alle/99_MitarbeiterInnen/JB/eatArchive/20251110_Demo/TVD",
#   path_to_archive_directory = "Q:/FDZ/Alle/99_MitarbeiterInnen/JB/eatArchive/20251110_Demo/TVD_AIP",
#   exclude_folders = c("_Archiv", "1a_Daten", "1b_Dokumentation", "3a_Vertrag"),
#   convert = TRUE,
#   overwrite = TRUE,
#   csv = "csv",
#   pdf_flavor = "2b"
# )
