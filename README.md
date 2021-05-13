# jpmorgan-check-transform

This R script was developed to streamline the process of transforming data exported from Financial Edge (.xlsx) into a format prescribed by JPMorgan Chase.

## System Dependencies

### R
To run this script, R (https://cran.r-project.org/) must be installed on a Windows machine. Once installed, locate the "R" folder within C:\Program Files\, drill down to the "bin" folder and copy the folder path as text. The path to the bin folder should look similar to "C:\Program Files\R\R-3.6.1\bin", with possible variation of the R version installed. Add this R bin folder path to the Path environment variables (both user and system) on the computer (This PC --> Properties --> Advanced system settings --> Environment Variables).

To test that R has been installed and added to the Path environement variable successfully, open a cmd prompt window, type in "Rscript" and hit Enter. If set up correctly, the help menu for Rscript should be returned.

### R Packages
Two R packages must be installed on the computer in order for the script to successfully run: readxl and tcltk.

To install these packages, open RStudio and run the following commands:

- install.packages("readxl")
- install.packages("tcltk")

<br/>

## File Location
In order to successfully run the script, the files must be saved to a folder named "jpmorgan-check-transformer" on the computer Desktop. This folder should contain the following files downloaded from this GitHub repository:

- README.md
- RunTransformation.cmd

A subfolder, "script" should contain:

- Transform.r

<br/>

## Running the Script
Once all of the above steps have been followed, the script can be run with the following steps:

1. Export the Financial Edge query file (.xlsx) and save to the jpmorgan-check-transform folder as "SourceDate.xlsx".

2. Double click RunTransformation.cmd to call the transformation.

3. If any error messages pop up, read and then troubleshoot the indicated error.

4. If successfully run, the transformed file for JPMorgan (output.csv) will be created in the jpmorgan-check-transform folder. You may need to refresh the Windows File Explorer to see the new file.
