# Create an archival directory of all files in a directory

Create an archival directory of all files in a directory

## Usage

``` r
create_archive_from_directory(
  path_to_working_directory,
  path_to_archive_directory,
  exclude_folders = "_Archive",
  convert = TRUE,
  overwrite = FALSE,
  csv = "csv"
)
```

## Arguments

- path_to_working_directory:

  Character. Path to working directory which is to be archived.

- path_to_archive_directory:

  Character. Path indicating where to create the archival directory.

- exclude_folders:

  Character vector. Names of subfolders to exclude from processing.

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

## Value

Folder and documentation of all files that have been copied.
