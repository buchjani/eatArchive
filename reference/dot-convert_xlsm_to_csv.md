# Convert all sheets of an Excel macro file to CSV

Convert all sheets of an Excel macro file to CSV

## Usage

``` r
.convert_xlsm_to_csv(xlsm_path, save_to, csv = "csv")
```

## Arguments

- xlsm_path:

  Path to the Excel macro file.

- save_to:

  Directory where the CSV files should be written.

- csv:

  Character specifying the csv type

  - "csv" uses "." for the decimal point and "," for the separator

  - "csv2" uses "," for the decimal point and ";" for the separator, the
    Excel convention for CSV files in some Western European locales.

## Value

Invisibly returns a character vector of paths to the written CSV files.
