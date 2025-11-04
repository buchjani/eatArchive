#' Convert all sheets of an Excel macro file to CSV
#'
#' @param xlsm_path Path to the Excel macro file.
#' @param save_to Directory where the CSV files should be written.
#' @param csv Character specifying the csv type
#'  - "csv" uses "." for the decimal point and "," for the separator
#'  - "csv2" uses "," for the decimal point and ";" for the separator, the Excel convention for CSV files in some Western European locales.
#'
#' @return Invisibly returns a character vector of paths to the written CSV files.
#' @keywords internal

.convert_xlsm_to_csv <- function(xlsm_path,
                                 save_to,
                                 csv = "csv"
) {

  sep <- ifelse(csv == "csv2", ";", ",")
  dec <- ifelse(csv == "csv2", ",", ".")

  stopifnot(is.character(xlsm_path), length(xlsm_path) == 1)
  if (!file.exists(xlsm_path)) {
    stop("File does not exist: ", xlsm_path, call. = FALSE)
  }

  # get sheet names
  sheet_names <- openxlsx::getSheetNames(tolower(xlsm_path))

  # in case of buggy (corrupted) xlsm-file, sheet_names will be empty
  # --> stop with reading the file, add error message to report

  if (length(sheet_names > 0)){
    # prepare output directory
    if (!dir.exists(save_to)) dir.create(save_to, recursive = TRUE)

    # build output paths: one CSV per sheet
    out_files <- file.path(save_to,
                           paste0(tools::file_path_sans_ext(basename(xlsm_path)), "_", sheet_names, ".csv"))

    # iterate over sheets
    for (i in seq_along(sheet_names)) {
      df <- suppressWarnings(openxlsx::read.xlsx(tolower(xlsm_path), sheet = sheet_names[i]))
      .write_csv_utf8_bom(df, out_files[i], sep = sep, dec = dec)
    }
    invisible(out_files)
  }
}



# # check with correct file
# .convert_xlsm_to_csv(xlsm_path = "Q:/FDZ/Alle/99_MitarbeiterInnen/JB/eatArchive/20251110_Demo/TVD/Diverse Dateiformate/Funktionierende Makro-Datei.Xlsm",
#                      save_to   = "Q:/FDZ/Alle/99_MitarbeiterInnen/JB/eatArchive/20251110_Demo/TVD_AIP",
#                      csv = "csv2")
#
# # check with corrupted file
# .convert_xlsm_to_csv(xlsm_path = "Q:/FDZ/Alle/99_MitarbeiterInnen/JB/eatArchive/20251110_Demo/TVD/Diverse Dateiformate/Kaputte Makro-Datei.Xlsm",
#                      save_to   = "Q:/FDZ/Alle/99_MitarbeiterInnen/JB/eatArchive/20251110_Demo/TVD_AIP",
#                      csv = "csv2")
