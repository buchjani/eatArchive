#' Scan directory and write table of file names and info
#'
#' @param dir A directory.
#' @param recursive Logical.
#'
#' @keywords internal

.get_file_info <- function(dir, recursive = FALSE) {

  # get file info, write to data.frame

  paths <- fs::dir_ls(dir, recurse = recursive, type = "file")
  paths <- as.character(paths)

  info <- data.frame(
    File_Name = paths,
    Size_Bytes = file.size(paths),
    Last_Modified = file.mtime(paths),
    stringsAsFactors = FALSE
  )

  return(info)
}
