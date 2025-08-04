get_file_info <- function(dir,
                          recursive = FALSE,
                          exclude_folders = NULL,
                          relative_to = dir) {

  files <- fs::dir_ls(dir, recurse = recursive, type = "file")

  # filter out files in any excluded folders
  if (!is.null(exclude_folders) && length(exclude_folders) > 0) {
    pattern <- paste0("/", exclude_folders, "/")
    combined_pattern <- paste(pattern, collapse = "|")
    files <- files[!grepl(combined_pattern, files)]
  }

  if (length(files) == 0) return(NULL)

  data.frame(
    File_Name = fs::path_rel(files, start = relative_to),
    Size_Bytes = fs::file_info(files)$size,
    Last_Modified = as.POSIXct(fs::file_info(files)$modification_time)
  )
}
