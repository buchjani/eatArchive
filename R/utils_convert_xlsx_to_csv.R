#' Convert all sheets of an Excel file to CSV
#'
#' @param xlsx_path Path to the Excel file.
#' @param save_to Directory where the CSV files should be written.
#' @param csv Character specifying the csv type
#'  - "csv" uses "." for the decimal point and "," for the separator
#'  - "csv2" uses "," for the decimal point and ";" for the separator, the Excel convention for CSV files in some Western European locales.
#'
#' @return Invisibly returns a character vector of paths to the written CSV files.
#' @keywords internal

.convert_xlsx_to_csv <- function(xlsx_path,
                                 save_to,
                                 csv = "csv"
                                 ) {

  sep <- ifelse(csv == "csv2", ";", ",")
  dec <- ifelse(csv == "csv2", ",", ".")

  stopifnot(is.character(xlsx_path), length(xlsx_path) == 1)
  if (!file.exists(xlsx_path)) {
    stop("File does not exist: ", xlsx_path, call. = FALSE)
  }

  # get sheet names
  # sheet_names <- openxlsx::getSheetNames(tolower(xlsx_path))  # fails when sheet names contain "()"
  wb <- suppressWarnings(openxlsx2::wb_load(xlsx_path))
  sheet_names <- openxlsx2::wb_get_sheet_names(wb)

  # prepare output directory
  if (!dir.exists(save_to)) dir.create(save_to, recursive = TRUE)

  # build output paths: one CSV per sheet
  out_names <- fs::path_sanitize(sheet_names)
  out_files <- file.path(save_to,
                         paste0(tools::file_path_sans_ext(basename(xlsx_path)), "_", out_names, ".csv"))

  # iterate over sheets
  for (i in seq_along(sheet_names)) {
    # df <- suppressWarnings(openxlsx::read.xlsx(tolower(xlsx_path), sheet = sheet_names[i])) # fails when sheet names contain "()"
    df <- suppressWarnings(openxlsx2::wb_to_df(xlsx_path, sheet = sheet_names[i]))
    .write_csv_utf8_bom(df, out_files[i], sep = sep, dec = dec)
  }

  invisible(out_files)
}

## check
# .convert_xlsx_to_csv(xlsx_path = "Q:/FDZ/Alle/99_MitarbeiterInnen/JB/eatArchive/20251110_Demo/TVD/Diverse Dateiformate/Excel-Datei mit Großbuchstaben.XLSX",
#                     save_to = "Q:/FDZ/Alle/99_MitarbeiterInnen/JB/eatArchive/20251110_Demo/TVD_AIP",
#                     csv = "csv2")

