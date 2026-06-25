#' Converts Word files (docx and doc) to txt-files
#'
#' @param docx_path
#' @param save_to
#'
#' @returns
#' @keywords internal

.convert_docx_to_txt <- function(docx_path,
                                 save_to = dirname(docx_path)) {

  is_docx <- grepl("\\.docx$", docx_path, ignore.case = TRUE)

  out <- file.path(
    save_to,
    paste0(tools::file_path_sans_ext(basename(docx_path)), ".txt")
  )

  if (is_docx) {
    if (nzchar(Sys.which("pandoc"))) {
      system2(
        "pandoc",
        c("-s", docx_path, "-t", "plain", "-o", out)
      )
    } else if (requireNamespace("readtext", quietly = TRUE)) {
      txt <- readtext::readtext(docx_path)$text
      writeLines(txt, out, useBytes = TRUE)
    } else {
      stop(
        "Converting .docx-files requires either 'pandoc' or the R package 'readtext'."
      )
    }
  } else {
    txt <- antiword::antiword(docx_path)
    txt <- iconv(
      txt,
      from = "WINDOWS-1252",
      to = "UTF-8"
    )
    writeLines(txt, out)
  }
  out
}

## Check:
# .convert_docx_to_txt(
#   docx_path = "Q:/FDZ/Alle/01_Studien/DigiFeed/Dokumentation/Infos_Studie/Anonymisierungsprotokoll.doc",
#   save_to  = "C:/R/.tmpdesktop")
