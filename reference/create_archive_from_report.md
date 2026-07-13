# Create an archival directory based on Excel file

Create an archival directory based on Excel file

## Usage

``` r
create_archive_from_report(
  path_to_directory_report,
  path_to_archive_directory,
  convert = TRUE,
  overwrite = TRUE,
  csv = "csv",
  pdf_flavor = "2b"
)
```

## Arguments

- path_to_directory_report:

  Character. Path to Excel file resulting from
  [`write_directory_report()`](https://buchjani.github.io/eatArchive/reference/write_directory_report.md)
  containing column "Archive", indicating whether and where to archive
  the file.

- path_to_archive_directory:

  Character. Path indicating where to create the archival directory.

- convert:

  Logical, indicating whether to convert file formats or not. Current
  files for conversion are:

  - xlsx –\> csv: each sheet of an excel file is converted to a csv file

  - xlsm –\> csv: each sheet of an excel macro file is converted to a
    csv file

  - sav –\> csv: SPSS data files are converted to two csv files each
    (data, metadata)

  - eml –\> txt: Thunderbird email objects (eml-files) are converted to
    txt-files

  - docx –\> txt: word files are converted to txt files

  - doc –\> txt: word files are converted to txt files

- overwrite:

  Logical, indicating whether to overwrite existing files in archival
  directory.

- csv:

  Character specifying the csv type

  - "csv" uses "." for the decimal point and "," for the separator

  - "csv2" uses "," for the decimal point and ";" for the separator, the
    Excel convention for CSV files in some Western European locales.

- pdf_flavor:

  Character, indicating the particular flavor of PDF files to be created
  (e.g., "2b")

## Value

Folder and documentation of all files that have been copied.
