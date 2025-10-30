#' Converts Word files (docx and doc) to txt-files
#'
#' @param docx_path
#' @param save_to
#'
#' @returns
#' @keywords internal

.convert_docx_to_txt <- function(docx_path,
                                save_to = dirname(docx_path),
                                overwrite = TRUE) {

  f <- docx_path
  is_docx <- grepl("\\.docx$", f, ignore.case = TRUE)
  base    <- tools::file_path_sans_ext(basename(f))
  out     <- file.path(save_to, paste0(base, ".txt"))

  if (file.exists(out) & !overwrite) {
    stop("File already exists and overwrite = FALSE:\n", out, call. = FALSE)
  }

  has_bin <- function(cmd) nzchar(Sys.which(cmd))

  if (is_docx) {
    # 1) Preferred: pandoc
    if (has_bin("pandoc")) {
      system2("pandoc", c("-s", f, "-t", "plain", "-o", out))
    } else {
      # 2) Fallback: R-only
      if (requireNamespace("readtext", quietly = TRUE)) {
        txt <- readtext::readtext(f)$text
        writeLines(txt, out, useBytes = TRUE)
      } else if (requireNamespace("textreadr", quietly = TRUE)) {
        txt <- textreadr::read_docx(f)
        writeLines(txt, out, useBytes = TRUE)
      } else {
        stop("For converting .docx files, please install either 'pandoc' or one of these R packages: 'readtext', 'textreadr'.")
      }
    }
  } else {
    # .doc
    if (has_bin("antiword")) {
      system2("antiword", c("-w", "0", f), stdout = out)
    } else if (has_bin("soffice")) {
      system2("soffice", c("--headless", "--convert-to", "txt:Text", "--out", out, f))
    } else {
      stop("For converting .doc files, please install 'antiword' or LibreOffice ('soffice').")
    }
  }
  return(out)
}

## Check:
# .convert_docx_to_txt(
#   docx_path = "C:/R/.tmpdesktop/Tables2-4 (Automatisch wiederhergestellt).docx",
#   save_to  = "C:/R/.tmpdesktop")
