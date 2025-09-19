#' Find the longest common path in a vector of paths
#'
#' @param vec_paths A character vector of paths.
#'
#' @keywords internal
#'
#' @examples
#' \dontrun{
#' paths <- c(
#' "c:/a/b/c/d",
#' "c:/a/b/e/f",
#' "c:/a/b/g/h",
#' "c:/a/i/j/d"
#' )
#' .longest_common_path(paths)
#' }

.longest_common_path <- function(vec_paths) {
  parts <- strsplit(vec_paths, "/")
  max_len <- min(lengths(parts))
  prefix <- character()

  for (i in seq_len(max_len)) {
    ith <- sapply(parts, `[`, i)
    if (length(unique(ith)) == 1) {
      prefix <- c(prefix, ith[1])
    } else break
  }
  paste(prefix, collapse = "/")
}

### check
# paths <- c(
#   "c:/a/b/c/d",
#   "c:/a/b/e/f",
#   "c:/a/b/g/h",
#   "c:/a/i/j/d"
# )
# .longest_common_path(paths)

