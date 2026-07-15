#' Title
#'
#' @param docx_path
#'
#' @returns
#'
#' @examples
#' @keywords internal

.convert_docx_to_zip <- function(docx_path,
                                 save_to) {

  zip_path <- file.path(
    save_to,
    paste0(tools::file_path_sans_ext(basename(docx_path)), ".zip")
  )

  file.copy(docx_path, zip_path, overwrite = TRUE)

  if (file.exists(zip_path)) {
    zip_path
  } else {
    NA_character_
  }
}

## Check:
# .convert_docx_to_zip("Q:/FDZ/Alle/99_MitarbeiterInnen/JB/eatArchive-Spielwiese/AIP/1b_Dokumentation/Lehrer_Voranalysen.docx",
#                      save_to  = "C:/R/.tmpdesktop")
