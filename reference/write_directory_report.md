# Create an Excel workbook with metadata for all subfolders of a directory

Scans each immediate subfolder of the specified directory (excluding any
listed in `exclude_folders`) and writes an Excel file containing
file-level metadata. Each subfolder gets its own sheet, listing all
files (recursively) with their name, size, and last modified
timestamp.  
  
Optionally, an additional column named "Archive" is included in each
sheet. The values of "Archive" will later be used by `create_archive()`
to indicate the destination folder for copying or converting files.

## Usage

``` r
write_directory_report(
  dir,
  path_to_directory_report,
  exclude_folders = "_Archive",
  autocomplete_values = NA,
  overwrite = FALSE
)
```

## Arguments

- dir:

  Character. Path to the directory whose immediate subfolders should be
  scanned.

- path_to_directory_report:

  Character. Path to the Excel file to be written (e.g.,
  "C:/study/dir_report.xlsx").

- exclude_folders:

  Character vector. Names of subfolders to exclude from processing.

- autocomplete_values:

  Optional. Controls whether the "Archive" column is included and which
  (if any) autocomplete values are to be provided:

  - If `NULL`, the "Archive" column is not included.

  - If `NA`, a column "Archive" is included but autocomplete and
    filtering in Excel will not be available.

  - If a character vector is provided, a column "Archive" is included
    and autocomplete and filtering in Excel is enabled, making it easier
    to assign target folders for file archiving.

- overwrite:

  Logical, indicating whether to overwrite existing report.

## Value

Invisibly returns the path to the written Excel file.

## Details

The Excel workbook includes:

- One sheet per immediate subfolder of the specified directory

- Each sheet contains a recursive listing of all files in that subfolder
  (including sub-subfolders)

- Optional autocomplete rows inserted directly below the header (e.g.,
  for manual annotation)

- Column widths and light styling for improved readability

Files located in any subfolder whose name matches one of the entries in
`exclude_folders` will be excluded entirely (i.e., they are not scanned
or listed).

This function is intended as part of a file management workflow in which
files may later be reviewed, copied, or converted based on the generated
metadata file.

## See also

[`createWorkbook`](https://rdrr.io/pkg/openxlsx/man/createWorkbook.html),
[`dir_ls`](https://fs.r-lib.org/reference/dir_ls.html),
[`file.info`](https://rdrr.io/r/base/file.info.html)

## Examples

``` r
if (FALSE) { # \dontrun{
# Write metadata for all subfolders of "my_project/"
write_directory_report("my_project/", "metadata_summary.xlsx")

# Include an autocomplete row with names of archival folders
write_directory_report("my_project/", "summary.xlsx", autocomplete_values = c("contracts", "data"))

# Include all subfolders (do not exclude any)
write_directory_report("my_project/", "summary.xlsx", exclude_folders = NULL)
} # }
```
