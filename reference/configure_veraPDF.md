# Configure the location of veraPDF installation

Opens an RStudio directory selection dialog and searches the selected
directory and its subdirectories for the veraPDF executable. This is
helpful when veraPDF is installed, but 'eatArchive' fails to find it.

## Usage

``` r
configure_veraPDF()
```

## Value

Invisibly returns `TRUE` if veraPDF was configured successfully,
otherwise `FALSE`.
