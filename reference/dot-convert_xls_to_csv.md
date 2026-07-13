# Convert all sheets of an xls-file to CSV

Convert all sheets of an xls-file to CSV

## Usage

``` r
.convert_xls_to_csv(xls_path, save_to, csv = "csv")
```

## Arguments

- xls_path:

  Path to the Excel file.

- save_to:

  Directory where the CSV files should be written.

- csv:

  Character specifying the csv type

  - "csv" uses "." for the decimal point and "," for the separator

  - "csv2" uses "," for the decimal point and ";" for the separator, the
    Excel convention for CSV files in some Western European locales.

## Value

Invisibly returns a character vector of paths to the written CSV files.
