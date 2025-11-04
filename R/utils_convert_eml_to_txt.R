#' Converts Thunderbird email objects (eml-files) to txt-files
#'
#' @param eml_path
#' @param save_to
#'
#' @returns
#' @keywords internal

.convert_eml_to_txt <- function(eml_path,
                                save_to = dirname(eml_path)) {

  base <- basename(eml_path)
  base_converted <-  gsub(".eml", ".txt", base)
  outfile <- file.path(save_to, base_converted)

  # save email as text file
  file.copy(from = eml_path,
            to = outfile,
            copy.date = FALSE)

  # return names of files
  return(outfile)
}

# # Check:
# .convert_eml_to_txt(
#   eml_path = "Q:/FDZ/Alle/99_MitarbeiterInnen/JB/eatArchive/20251110_Demo/TVD/Diverse Dateiformate/E-Mail-Datei.eml",
#   save_to  = "Q:/FDZ/Alle/99_MitarbeiterInnen/JB/eatArchive/20251110_Demo/TVD_AIP/Diverse Dateiformate")
