#' Configure the location of veraPDF installation
#'
#' Opens an RStudio directory selection dialog and searches the selected directory and its
#' subdirectories for the veraPDF executable. This is helpful when veraPDF is installed, but
#' 'eatArchive' fails to find it.
#'
#' @return Invisibly returns `TRUE` if veraPDF was configured successfully, otherwise `FALSE`.
#' @export

configure_verapdf <- function() {
  if (!requireNamespace("rstudioapi", quietly = TRUE)) {
    stop(
      "Package 'rstudioapi' is required to select the veraPDF folder.",
      call. = FALSE
    )
  }

  if (!rstudioapi::isAvailable()) {
    stop(
      "This function must be run from RStudio.",
      call. = FALSE
    )
  }

  selected_dir <- rstudioapi::selectDirectory(
    caption = "Select the folder containing veraPDF"
  )

  if (is.null(selected_dir) || !nzchar(selected_dir)) {
    message("No folder selected.")
    return(invisible(FALSE))
  }

  hits <- list.files(
    path.expand(selected_dir),
    pattern = "^verapdf(\\.bat|\\.exe|\\.sh)?$",
    recursive = TRUE,
    full.names = TRUE,
    ignore.case = TRUE
  )

  # Exclude directories that happen to be named "verapdf"
  hits <- hits[file.exists(hits) & !dir.exists(hits)]

  if (length(hits) == 0L) {
    message(
      "No veraPDF executable was found in the selected folder ",
      "or its subdirectories."
    )
    return(invisible(FALSE))
  }

  verapdf_path <- normalizePath(
    hits[[1L]],
    winslash = "/",
    mustWork = TRUE
  )

  config_dir <- tools::R_user_dir(
    package = "eatArchive",
    which = "config"
  )

  dir.create(
    config_dir,
    recursive = TRUE,
    showWarnings = FALSE
  )

  writeLines(
    verapdf_path,
    file.path(config_dir, "verapdf_path.txt"),
    useBytes = TRUE
  )

  message("veraPDF configured successfully:\n", verapdf_path)

  invisible(TRUE)
}
# set_verapdf()
