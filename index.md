# eatArchive

eatArchive helps you automate the archiving of directory contents using
open, software-agnostic file formats. The package supports scanning
nested directories, copying files into a new folder structure, and
converting selected formats (e.g., XLSX to CSV, DOCX to PDF/A). Each
step is documented in a machine-readable CSV log that records source
paths, destination paths, and any format conversions applied.

## Installation

You can install the development version of eatArchive from Github with

``` r

remotes::install_github("buchjani/eatArchive")
```

## Use eatArchive

You can use eatArchive after loading it from your library. Make sure to
always refer to the current version when reporting
[issues](https://github.com/buchjani/eatArchive/issues)

``` r

library(eatArchive)
```

> **Version 0.5.1**
