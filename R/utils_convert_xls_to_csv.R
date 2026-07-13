#' Convert all sheets of an xls-file to CSV
#'
#' @param xls_path Path to the Excel file.
#' @param save_to Directory where the CSV files should be written.
#' @param csv Character specifying the csv type
#'  - "csv" uses "." for the decimal point and "," for the separator
#'  - "csv2" uses "," for the decimal point and ";" for the separator, the Excel convention for CSV files in some Western European locales.
#'
#' @return Invisibly returns a character vector of paths to the written CSV files.
#' @keywords internal

.convert_xls_to_csv <- function(xls_path,
                                save_to,
                                csv = "csv"
) {

  sep <- ifelse(csv == "csv2", ";", ",")
  dec <- ifelse(csv == "csv2", ",", ".")

  stopifnot(is.character(xls_path), length(xls_path) == 1)
  if (!file.exists(xls_path)) {
    stop("File does not exist: ", xls_path, call. = FALSE)
  }

  # get sheet names
  sheet_names <- readxl::excel_sheets(xls_path)

  # prepare output directory
  if (!dir.exists(save_to)) dir.create(save_to, recursive = TRUE)

  # build output paths: one CSV per sheet
  out_names <- fs::path_sanitize(sheet_names)
  out_files <- file.path(save_to,
                         paste0(tools::file_path_sans_ext(basename(xls_path)), "_", out_names, ".csv"))

  # iterate over sheets
  for (i in seq_along(sheet_names)) {
    # df <- suppressWarnings(openxlsx::read.xlsx(tolower(xlsx_path), sheet = sheet_names[i])) # fails when sheet names contain "()"
    # df <- suppressWarnings(openxlsx2::wb_to_df(xlsx_path, sheet = sheet_names[i]))
    df <- suppressMessages(suppressWarnings(readxl::read_xls(xls_path, sheet = sheet_names[i], col_names = FALSE)))
    .write_csv_utf8_bom(df, out_files[i], sep = sep, dec = dec, col.names = FALSE)
  }

  invisible(out_files)
}

## check
# .convert_xls_to_csv(xls_path = "Q:/FDZ/Alle/99_MitarbeiterInnen/JB/eatArchive/20251110_Demo/TVD/Diverse Dateiformate/eine alte XLS DATEI.xls",
#                     save_to = "Q:/FDZ/Alle/99_MitarbeiterInnen/JB/eatArchive/20251110_Demo/TVD_AIP",
#                     csv = "csv2")

