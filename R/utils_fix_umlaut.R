#' Umlaute in Dateinamen ersetzen (ä→ae, ö→oe, ü→ue, ß→ss; Großbuchstaben: Ae/Oe/Ue)
#'
#' @param string
#'
#' @returns
#' @keywords internal

.fix_umlaut <- function(string) {
  string <- gsub("ä", "ae", string, fixed = TRUE)
  string <- gsub("ö", "oe", string, fixed = TRUE)
  string <- gsub("ü", "ue", string, fixed = TRUE)
  string <- gsub("Ä", "Ae", string, fixed = TRUE)
  string <- gsub("Ö", "Oe", string, fixed = TRUE)
  string <- gsub("Ü", "Ue", string, fixed = TRUE)
  string <- gsub("ß", "ss", string, fixed = TRUE)
  return(string)
}

# # check:
# meinpfad <- "c:/mein/Pfäd/mit/Ümläutenß"
# .fix_umlaut(meinpfad)
