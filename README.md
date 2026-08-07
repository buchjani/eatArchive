
<!-- README.md is generated from README.Rmd. Please edit that file -->

# eatArchive <a href="https://buchjani.github.io/eatArchive/"><img src="man/figures/logo.png" align="right" height="120" alt="eatArchive website" /></a>

<!-- badges: start -->

[![R-CMD-check](https://github.com/buchjani/eatArchive/actions/workflows/R-CMD-check.yaml/badge.svg)](https://github.com/buchjani/eatArchive/actions/workflows/R-CMD-check.yaml)
[![Lifecycle:
experimental](https://img.shields.io/badge/lifecycle-experimental-orange.svg)](https://lifecycle.r-lib.org/articles/stages.html#experimental)
<!-- [![Project Status: Active - The project has reached a stable, usable state and is being actively developed.](https://www.repostatus.org/badges/latest/active.svg)](https://www.repostatus.org/#active) -->
[![](https://img.shields.io/github/last-commit/buchjani/eatArchive.svg)](https://github.com/buchjani/eatArchive/commits/master)
[![](https://img.shields.io/github/languages/code-size/buchjani/eatArchive.svg)](https://github.com/buchjani/eatArchive)
[![](https://img.shields.io/badge/author%20experience-1st%20R%20package-green.svg)](https://www.iqb.hu-berlin.de/institut/staff/?pg=c163)
<!-- [![](http://cranlogs.r-pkg.org/badges/grand-total/eatArchive?color=green)](https://cran.r-project.org/package=eatArchive) -->
<!-- badges: end -->

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
