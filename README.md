
<!-- README.md is generated from README.Rmd. Please edit that file -->

# eatArchive

<!-- badges: start -->

<!-- badges: end -->

The goal of eatArchive is to automate the archiving of directory
contents using open, software-agnostic file formats. The package
supports scanning nested directories, copying files into a new folder
structure, and converting selected formats (e.g., XLSX to CSV, DOCX to
PDF/A). Each step is documented in a machine-readable log (CSV),
containing the source paths, destination paths, and any format
conversions applied.

## Installation

You can install the development version of eatArchive like so:

``` r
# FILL THIS IN! HOW CAN PEOPLE INSTALL YOUR DEV PACKAGE?
```

## Workflow

This is a schematic representation of the workflow, consisting of three
major steps:

1.  R function `write_metadata()` for generating an excel file
    documenting the contents of a directory  
2.  selecting files for archive, by specifying their archival
    subdirectory in the excel file
3.  R function `create_archive()` for reading the excel file, copying
    and converting files from the original directory to the archival
    directory

<figure id="id">
<img src="./man/figures/README-workflow.png" class="class" width="900"
alt="eatArchive-workflow" />
<figcaption aria-hidden="true">eatArchive-workflow</figcaption>
</figure>

What is special about using `README.Rmd` instead of just `README.md`?
You can include R chunks like so:

``` r
summary(cars)
#>      speed           dist       
#>  Min.   : 4.0   Min.   :  2.00  
#>  1st Qu.:12.0   1st Qu.: 26.00  
#>  Median :15.0   Median : 36.00  
#>  Mean   :15.4   Mean   : 42.98  
#>  3rd Qu.:19.0   3rd Qu.: 56.00  
#>  Max.   :25.0   Max.   :120.00
```

You’ll still need to render `README.Rmd` regularly, to keep `README.md`
up-to-date. `devtools::build_readme()` is handy for this.
