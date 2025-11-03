#' Convert all sheets of an Excel macro file to CSV
#'
#' @param xlsm_path Path to the Excel macro file.
#' @param save_to Directory where the CSV files should be written.
#' @param overwrite Logical, whether to overwrite existing CSV files. Default = FALSE.
#'
#' @return Invisibly returns a character vector of paths to the written CSV files.
#' @keywords internal

.convert_xlsm_to_csv <- function(xlsm_path,
                                 save_to,
                                 overwrite = FALSE
                                 #encoding = "UTF-8"
) {

  xlsm_path <- tolower(xlsm_path)

  stopifnot(is.character(xlsm_path), length(xlsm_path) == 1)
  if (!file.exists(xlsm_path)) {
    stop("File does not exist: ", xlsm_path, call. = FALSE)
  }

  # get sheet names
  sheet_names <- openxlsx::getSheetNames(xlsm_path)

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
      df <- suppressWarnings(openxlsx::read.xlsx(xlsm_path, sheet = sheet_names[i]))
      .write_csv_utf8_bom(df, out_files[i], sep = ",", overwrite = overwrite)
    }
    invisible(out_files)
  }
}

.write_csv_utf8_bom <- function(df, path, sep = ",", dec = ".", overwrite = FALSE) {
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
    dec = dec,
    quote = FALSE,
    row.names = FALSE,
    col.names = TRUE,
    qmethod = "double",
    na = ""
  )
  close(con)
}

# # check with corrupted file
# .convert_xlsm_to_csv(xlsm_path = "Q:/FDZ/Alle/99_MitarbeiterInnen/JB/eatArchive/20251103_Demo/Beispielordner/Eval-VD-NRW-K3_MD_gesamt_EN.xlsm",
#                     save_to = "Q:/FDZ/Alle/99_MitarbeiterInnen/JB/eatArchive/20251103_Demo/Beispielordner_AIP/zentrale Dokumente/",
#                     overwrite = TRUE)
#
# # check with correct file
# .convert_xlsm_to_csv(xlsm_path = "Q:/FDZ/Alle/99_MitarbeiterInnen/JB/eatArchive/20251103_Demo/Beispielordner/LFB_Ausgabe.xlsm",
#                     save_to = "Q:/FDZ/Alle/99_MitarbeiterInnen/JB/eatArchive/20251103_Demo/Beispielordner_AIP/zentrale Dokumente",
#                     overwrite = TRUE)
