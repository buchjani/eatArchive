#' Convert a PDF file to PDF/A file with a particular flavor (e.g. "2b")
#'
#' @param pdf_path
#' @param save_to
#' @param pdf_flavor
#'
#' @returns
#'
#' @examples
#' @keywords internal



# MAIN FUNCTION ---------------------------------------------------------------

.convert_pdf_to_pdfa <- function(pdf_path,
                                 save_to,
                                 pdf_flavor = "2b") {

  if (!file.exists(pdf_path)) {
    stop("Input PDF not found: ", pdf_path, call. = FALSE)
  }

  icc <- .find_icc()
  if (!nzchar(icc)) {
    stop(
      "Internal ICC profile not found in installed package.",
      call. = FALSE
    )
  }

  gs <- .find_ghostscript()
  if (!nzchar(gs)) {
    stop("Ghostscript not found.", call. = FALSE)
  }

  veraPDF_path <- .find_veraPDF()

  if (!dir.exists(save_to)) {
    dir.create(save_to, recursive = TRUE)
  }

  base <- basename(pdf_path)
  base_converted <- sub(
    "\\.pdf$",
    "_pdfa.pdf",
    base,
    ignore.case = TRUE
  )
  base_converted <- .fix_umlaut(base_converted)
  outfile <- file.path(save_to, base_converted)

  part <- as.integer(substr(pdf_flavor, 1L, 1L))
  conf <- substr(pdf_flavor, 2L, 2L)

  if (conf == "a") {
    warning(
      "Target flavor 'PDF/A-", part,
      "a' requires a tagged (structured) PDF; ",
      "Ghostscript usually does not create tags. ",
      "Validate output with veraPDF.",
      call. = FALSE
    )
  }

  if (file.exists(outfile)) {
    file.remove(outfile)
  }

  pdf_path_n <- normalizePath(
    pdf_path,
    winslash = "/",
    mustWork = TRUE
  )

  out_path_n <- normalizePath(
    outfile,
    winslash = "/",
    mustWork = FALSE
  )

  icc_n <- normalizePath(
    icc,
    winslash = "/",
    mustWork = TRUE
  )

  pdfa_def_tmp <- .write_pdfa_def_pdfmark(
    icc_n,
    part = part
  )

  on.exit(
    unlink(pdfa_def_tmp, force = TRUE),
    add = TRUE
  )

  pdfa_def_n <- normalizePath(
    pdfa_def_tmp,
    winslash = "/",
    mustWork = TRUE
  )

  args <- c(
    "-dSAFER",

    # PDFA_def.ps opens the ICC profile itself.
    # Therefore SAFER must explicitly allow reading this file.
    paste0(
      "--permit-file-read=",
      shQuote(icc_n)
    ),

    paste0("-dPDFA=", part),
    "-dPDFACompatibilityPolicy=1",
    "-sDEVICE=pdfwrite",
    "-dBATCH",
    "-dNOPAUSE",
    "-dEmbedAllFonts=true",
    "-dSubsetFonts=true",

    paste0(
      "-sOutputFile=",
      shQuote(out_path_n)
    ),

    "-f",
    shQuote(pdfa_def_n),
    shQuote(pdf_path_n)
  )

  gs_log <- suppressWarnings(
    system2(
      gs,
      args = args,
      stdout = TRUE,
      stderr = TRUE
    )
  )

  exit_code <- attr(gs_log, "status")
  if (is.null(exit_code)) {
    exit_code <- 0L
  }

  # Ghostscript may create a partial or blank PDF before failing.
  # Delete such output immediately.
  if (exit_code != 0L || !file.exists(outfile)) {
    if (file.exists(outfile)) {
      file.remove(outfile)
    }

    return(NA_character_)
  }

  ok <- isTRUE(
    .check_pdf_flavor(
      outfile,
      flavor = pdf_flavor,
      veraPDF_path = veraPDF_path
    )
  )

  if (!ok) {
    file.remove(outfile)
    return(NA_character_)
  }

  outfile
}


## Test:
# .convert_pdf_to_pdfa(
#   pdf_path = "C:/R/eatArchive_lokal/PDF/PISA_2022_Datenbereitstellungsvertrag.pdf",
#   save_to = "C:/R/eatArchive_lokal/PDF",
#   pdf_flavor = "1b")
# .check_pdf_flavor("C:/R/eatArchive_lokal/PDF/PISA_2022_Datenbereitstellungsvertrag_pdfa.pdf",
#                   "1b",
#                   .find_veraPDF())

## Test:
# outfile <- .convert_pdf_to_pdfa(
#   pdf_path = "Q:/FDZ/Alle/99_MitarbeiterInnen/JB/eatArchive/Beispieldateien/PDF_mit Formatierung.pdf",
#   save_to = "C:/R/PDF",
#   pdf_flavor = "2b")
# .check_pdf_flavor(outfile,
#                   "2b",
#                   .find_veraPDF())
#


# HELPER FUNCTIONS ---------------------------------------------------------------

.find_veraPDF <- function() {
  config_file <- file.path(
    tools::R_user_dir("eatArchive", "config"),
    "verapdf_path.txt"
  )

  # 1) Use the manually configured path
  if (file.exists(config_file)) {
    verapdf_path <- readLines(
      config_file,
      n = 1L,
      warn = FALSE
    )

    if (
      length(verapdf_path) == 1L &&
      nzchar(verapdf_path) &&
      file.exists(verapdf_path)
    ) {
      return(verapdf_path)
    }
  }

  # 2) Search the system PATH
  hits <- Sys.which(c(
    "verapdf",
    "verapdf.bat",
    "verapdf.sh"
  ))

  hits <- unname(hits[nzchar(hits)])

  if (length(hits) > 0L && file.exists(hits[[1L]])) {
    return(hits[[1L]])
  }

  # 3) veraPDF is unavailable --> return empty string
  NA_character_
}
# .find_veraPDF()


.find_icc <- function() {
  icc <- system.file("extdata", "sRGB2014.icc", package = "eatArchive")
  if (!nzchar(icc) || !file.exists(icc)) return("")
  icc
}
# .find_icc()

.find_ghostscript <- function() {
  candidates <- c(Sys.which("gswin64c.exe"), Sys.which("gs"))
  gs <- candidates[nzchar(candidates)][1]
  if (!nzchar(gs)) return("")
  gs
}
# .find_ghostscript()

# This is a function for creating a temporary ghostscript file, including the actual paths.
# A "permanent" .ps-file didn't (always) produce the correct outcome
.write_pdfa_def_pdfmark <- function(icc_path, part = 2L) {
  icc_path <- normalizePath(
    icc_path,
    winslash = "/",
    mustWork = TRUE
  )

  # Escape parentheses for PostScript strings
  icc_path <- gsub(
    "([()])",
    "\\\\\\1",
    icc_path
  )

  lines <- c(
    "%!",
    "% PDF/A pdfmark definition generated by R",
    sprintf("/ICCProfile (%s) def", icc_path),
    "",
    "% Define ICC profile stream",
    "[ /_objdef {icc_PDFA} /type /stream /OBJ pdfmark",
    "[ {icc_PDFA} <</N 3>> /PUT pdfmark",
    "[ {icc_PDFA} ICCProfile (r) file /PUT pdfmark",
    "",
    "% Define OutputIntent dictionary",
    "[ /_objdef {OutputIntent_PDFA} /type /dict /OBJ pdfmark",
    paste0(
      "[ {OutputIntent_PDFA} <<",
      "/Type /OutputIntent ",
      "/S /GTS_PDFA1 ",
      "/DestOutputProfile {icc_PDFA}"
    ),
    "  /OutputConditionIdentifier (sRGB)",
    "  /Info (sRGB IEC61966-2.1)",
    "  /RegistryName (http://www.color.org)",
    ">> /PUT pdfmark",
    "",
    "% Attach OutputIntent to the catalog",
    "[ {Catalog} <</OutputIntents [ {OutputIntent_PDFA} ]>> /PUT pdfmark"
  )

  file <- tempfile(
    pattern = "PDFA_def_",
    fileext = ".ps"
  )

  writeLines(
    lines,
    file,
    useBytes = TRUE
  )

  file
}


.check_pdf_flavor <- function(file,
                              flavor = "2b",
                              veraPDF_path) {

  if (!is.character(veraPDF_path) || !nzchar(veraPDF_path) || !file.exists(veraPDF_path)) {
    return(FALSE)
  }

  xml_lines <- suppressWarnings(system2(
    veraPDF_path,
    args   = c("--loglevel", "0", "--format", "xml", "-f", flavor, shQuote(file)),
    stdout = TRUE,
    stderr = FALSE
  ))

  line <- xml_lines[
    grepl("<validationReport\\b", xml_lines) &
      grepl("isCompliant=", xml_lines)
  ][1]

  !is.na(line) && grepl('isCompliant="true"', line)
}
# .check_pdf_flavor("C:/R/eatArchive_lokal/PDF/PISA_2022_Datenbereitstellungsvertrag_2b.pdf",
#                   "2b",
#                   "C:\\Users\\buchjani/verapdf/verapdf.bat")

