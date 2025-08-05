get_file_info <- function(dir, recursive = FALSE) {
  # Get all files (not directories)
  paths <- fs::dir_ls(dir, recurse = recursive, type = "file")

  # Convert to character vector for easier handling
  paths <- as.character(paths)

  file_info <- file.info(paths)

  # Create a data.frame with desired metadata
  info <- data.frame(
    File_Name = paths,
    Size_Bytes = file_info$size,
    Last_Modified = file_info$mtime,
    stringsAsFactors = FALSE
  )

  return(info)
}

# check:
# get_file_info(dir = "Q:/FDZ/_Datensaetze/LZA/078_MEZ_v1/Begleitmaterialien LZA-konform", recursive = TRUE)
