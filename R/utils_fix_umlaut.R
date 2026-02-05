#' Umlaute in Dateinamen ersetzen (ä→ae, ö→oe, ü→ue, ß→ss; Großbuchstaben: Ae/Oe/Ue)
#'
#' @param x
#'
#' @returns
#' @keywords internal

.fix_umlaut <- function(x) {
  x <- gsub("\u00C4", "Ae", x, fixed = TRUE) # Ä
  x <- gsub("\u00E4", "ae", x, fixed = TRUE) # ä
  x <- gsub("\u00D6", "Oe", x, fixed = TRUE) # Ö
  x <- gsub("\u00F6", "oe", x, fixed = TRUE) # ö
  x <- gsub("\u00DC", "Ue", x, fixed = TRUE) # Ü
  x <- gsub("\u00FC", "ue", x, fixed = TRUE) # ü
  x <- gsub("\u00DF", "ss", x, fixed = TRUE) # ß
  x <- gsub(" ", "_", x, fixed = TRUE)       # Leerzeichen
  return(x)
}

# # check:
# meinpfad <- "c:/mein/Pfäd/m i t/Ümläutenß"
# .fix_umlaut(meinpfad)
