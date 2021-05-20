# jpmorgan-check-transform

This R script was developed to streamline the process of transforming data exported from Financial Edge (.xlsx) into a format prescribed by JPMorgan Chase.

## System Dependencies

### R
To run this script, R (https://cran.r-project.org/) must be installed on a Windows machine. Once installed, locate the "R" folder within C:\Program Files\, drill down to the "bin" folder and copy the folder path as text. The path to the bin folder should look similar to "C:\Program Files\R\R-3.6.1\bin", with possible variation of the R version installed. Add this R bin folder path to the Path environment variables (both user and system) on the computer (This PC --> Properties --> Advanced system settings --> Environment Variables).

To test that R has been installed and added to the Path environement variable successfully, open a cmd prompt window, type in "Rscript" and hit Enter. If set up correctly, the help menu for Rscript should be returned.

### R Packages
Two R packages must be installed on the computer in order for the script to successfully run: readxl and tcltk.

To install these packages, open RStudio or RGUI and run the following commands:

- install.packages("readxl")
- install.packages("tcltk")

Confirm successful installation of each package after running each command. Each package library will either be saved to the C:\ drive or in a personal library on successful install.

<br/>

## Directory Set-Up

This script has been developed to be run from any directory location, as long as the following requirements are met:

- Script file (Transform.r) is stored in the same folder as the command file (RunTransformation.cmd).

- Three sub-folders must exist in this directory:
    * "completed"
    * "input"
    * "output"
    
<br/>

## Running the Script
Once all of the above set-up steps have been followed, the script can be run as follows:

1. Export the Financial Edge query file(s) (.xlsx) that need to be transformed for JPMorgan check creation. Save the file(s) in the "input" folder.

2. Double click RunTransformation.cmd in the main directory to call the transformation.

3. If any error message or notifications pop up, read the message and troubleshoot as needed.

4. If successfully run, the transformed file(s) for JPMorgan will be created in the "output" folder. The output csv(s) will have the same filename(s) as the input file(s). Input files will be moved to the "completed" folder following a successful transformation. Any input files that raised an error during the process will stay in the "input" folder. You may need to refresh the Windows File Explorer to see the new file(s).
