#' Write a csv file with an encoding that preserves Umlaute (BOM)
#'
#' @param df
#' @param path
#' @param sep
#' @param dec
#' @param overwrite
#'
#' @returns
#'
#' @keywords internal

.write_csv_utf8_bom <- function(df, path, sep = ",", dec = ".", overwrite = TRUE) {

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
