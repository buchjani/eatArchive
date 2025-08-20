#' Convert all sheets of an Excel file to CSV
#'
#' @param xlsx_path Path to the Excel file.
#' @param save_to Directory where the CSV files should be written.
#' @param overwrite Logical, whether to overwrite existing CSV files. Default = FALSE.
#'
#' @return Invisibly returns a character vector of paths to the written CSV files.
#' @keywords internal

.convert_xlsx_to_csv <- function(xlsx_path,
                                 save_to,
                                 overwrite = FALSE
                                 #encoding = "UTF-8"
                                 ) {

  # get sheet names
  sheet_names <- openxlsx::getSheetNames(xlsx_path)

  # prepare output directory
  if (!dir.exists(save_to)) dir.create(save_to, recursive = TRUE)

  # build output paths: one CSV per sheet
  out_files <- file.path(save_to,
                         paste0(tools::file_path_sans_ext(basename(xlsx_path)), "_", sheet_names, ".csv"))

  # iterate over sheets
  for (i in seq_along(sheet_names)) {
    df <- openxlsx::read.xlsx(xlsx_path, sheet = sheet_names[i])
    .write_csv_utf8_bom(df, out_files[i], sep = ",", overwrite = overwrite)
  }

  invisible(out_files)
}

.write_csv_utf8_bom <- function(df, path, sep = ",", overwrite = FALSE) {
  if (file.exists(path) && !overwrite) {
    stop("File exists and overwrite=FALSE: ", path)
  }
  # open binary connection
  con <- file(path, open = "wb")
  # write BOM
  writeBin(as.raw(c(0xEF, 0xBB, 0xBF)), con)
  # now switch the connection to text mode with UTF-8
  close(con)
  con <- file(path, open = "a", encoding = "UTF-8")

  # write the table
  utils::write.table(
    df, file = con,
    sep = sep,
    row.names = FALSE, col.names = TRUE,
    qmethod = "double", na = ""
  )
  close(con)
}

# check
.convert_xlsx_to_csv(xlsx_path = "Q:/FDZ/Alle/01_Studien/TIMSS/TIMSS_2019/2_Pruefung/2a_Datenschutz/TIMSS_2019_v4_neueDatenschutzpruefung_mitNW.xlsx",
                     save_to = "C:/R/",
                     overwrite = TRUE)

