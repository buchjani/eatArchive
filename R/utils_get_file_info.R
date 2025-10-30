#' Scan directory and write table of file names and info
#'
#' @param dir A directory.
#' @param recursive Logical.
#'
#' @keywords internal

.get_file_info <- function(dir, recursive = FALSE) {
  # Get all files (not directories)
  paths <- fs::dir_ls(dir, recurse = recursive, type = "file")

  # Convert to character vector for easier handling
  paths <- as.character(paths)

  # Create a data.frame with desired metadata
  info <- data.frame(
    File_Name = paths,
    Size_Bytes = file.size(paths),
    Last_Modified = file.mtime(paths),
    stringsAsFactors = FALSE
  )

  return(info)
}
