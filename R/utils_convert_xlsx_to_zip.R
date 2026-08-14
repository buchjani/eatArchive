#' Title
#'
#' @param xlsx_path
#'
#' @returns
#'
#' @examples
#' @keywords internal

.convert_xlsx_to_zip <- function(xlsx_path,
                                 save_to) {

  zip_path <- file.path(
    save_to,
    paste0(tools::file_path_sans_ext(.fix_umlaut(basename(xlsx_path))), ".zip")
  )

  file.copy(xlsx_path, zip_path, overwrite = TRUE)

  if (file.exists(zip_path)) {
    zip_path
  } else {
    NA_character_
  }
}

## Check:
# .convert_xlsx_to_zip("Q:/FDZ/Alle/99_MitarbeiterInnen/JB/eatArchive/Beispieldateien/xlsx_with formating & sheets.xlsx",
#                     save_to  = "C:/R/.tmpdesktop")
