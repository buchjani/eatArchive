#' Converts sav files (GADS objects) to csv-files
#'
#' @param sav_path
#' @param save_to
#' @param csv
#'
#' @returns
#' @keywords internal

.convert_sav_to_csv <- function(sav_path,
                                save_to = dirname(sav_path),
                                csv = csv) {

  sep <- ifelse(csv == "csv2", ";", ",")
  dec <- ifelse(csv == "csv2", ",", ".")

  f <- sav_path
  base    <- tools::file_path_sans_ext(basename(f))
  df_GADS <- eatGADS::import_spss(f, labeledStrings = "keep")

  # files to write
  outfiles <- c(file.path(save_to, paste0(base, "_meta.csv")),
                file.path(save_to, paste0(base, "_data.csv"))
  )

  # write csv-files
  .write_csv_utf8_bom(df_GADS$labels, outfiles[1], sep = sep, dec = dec)
  .write_csv_utf8_bom(df_GADS$dat, outfiles[2], sep = sep, dec = dec)

  # return names of files
  return(outfiles)
}

# Check:
# .convert_sav_to_csv(
#   sav_path = "Q:/FDZ/Alle/99_MitarbeiterInnen/JB/eatArchive/20251110_Demo/TVD/Diverse Dateiformate/SPSS-Datei.sav",
#   save_to  = "Q:/FDZ/Alle/99_MitarbeiterInnen/JB/eatArchive/20251110_Demo/TVD_AIP/Diverse Dateiformate",
#   csv = "csv")
