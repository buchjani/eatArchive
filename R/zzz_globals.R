.onAttach <- function(libname, pkgname) {
  if (interactive()) {
    vers <- as.character(utils::packageVersion(pkgname))
    packageStartupMessage(sprintf("%s %s", pkgname, vers))
  }
}

utils::globalVariables(c(
  ".add_sheet_with_style",
  ".combine_excel_sheets",
  ".get_file_info",
  ".longest_common_path",
  ".fix_umlaut",
  ".write_csv_utf8_bom",
  ".convert_xlsx_to_csv",
  ".convert_xlsm_to_csv",
  ".convert_docx_to_txt",
  ".convert_docx_to_zip",
  ".convert_pdf_to_pdfa"
))
