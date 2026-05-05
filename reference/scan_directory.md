# Search a directory for specific files (date, size)

Scans each immediate subfolder of the specified directory (excluding any
listed in `exclude_folders`) and writes an Excel file containing
file-level metadata for specific files. Files can be selected by size
and by date of last modification.

## Usage

``` r
scan_directory(
  dir,
  path_to_report,
  exclude_folders = "",
  show_what = c("largest", "smallest", "newest", "oldest"),
  show_max = 1000,
  overwrite = FALSE
)
```

## Arguments

- dir:

  Character. Path to the directory whose immediate subfolders should be
  scanned.

- path_to_report:

  Character. Path to the Excel file to be written (e.g.,
  "C:/study/dir_report.xlsx").

- exclude_folders:

  Character vector. Names of subfolders to exclude from processing.

- show_what:

  Character, indicating whehter to sort by size or date.

  - `"largest"` for sorting by file size (large to small)

  - `"smallest"` for sorting by file size (small to large)

  - `"newest"` for sorting by date of last modification (new to old)

  - `"oldest"` for sorting by date of last modification (old to new)

- show_max:

  Numeric, indicating the maximum number of files to include in the
  report.

- overwrite:

  Logical, indicating whether to overwrite existing report.

## Value

Invisibly returns the path to the written Excel file.

## Details

The Excel workbook includes a recursive listing of all files meeting the
search criteria (including sub-subfolders).

Files located in any subfolder whose name matches one of the entries in
`exclude_folders` will be excluded entirely (i.e., they are not scanned
or listed).
