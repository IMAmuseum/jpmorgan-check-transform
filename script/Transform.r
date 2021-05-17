## install packages
## install.packages("readxl")
## install.packages("tcltk")

## load package libraries
library("readxl")
library("tcltk")

## set variable user
user <- Sys.getenv("USERNAME")

######################
## Question: Where would this process live, ideally? Demo was from the Desktop, but we can change this to any
## location on the machine.
######################

## set variable wd as path to JPMorgan folder on desktop
wd <- paste("C:\\Users\\", user, "\\Desktop\\jpmorgan-check-transform", sep="")

## setwd
setwd(wd)

######################
## Question: Will this process be run one file at a time, and/or with batches of multiple source files?
## Current set-up assumes one single source file, with specific filename SourceData.xlsx. 
## I can adjust this filename for the source data if there is something else you'd prefer, or change the process
## to look for all .xlsx files, regardless of name, in the root folder, and output a csv of the same name.
######################

## check that SourceData.xlsx exists in wd
if (file.exists("SourceData.xlsx")) {

## read source .xlsx file into R
source_data <- read_excel("SourceData.xlsx")

## check that SourceData.xlsx is structured as expected
file_cols <- colnames(source_data)
expected_cols <- c("Check format","Payment date","Amount","Account number","Payment number","Payee name","Address line 1","Address line 2","City","State/Province","ZIP/Post code","Invoice number","Description","Invoice date","Invoice amount","Discount amount","Payment type")

if(identical(file_cols,expected_cols)) {

######################
## Question: Is there a specific filename you would like the output file to have? Currently coded as simply
## "output.csv". Filename can include dynamic info such as current date.
######################
  
## set variable filename
filename <- paste(wd, "\\output.csv", sep="")

## create header dataframe for output.csv
header <- data.frame("FILHDR","PWS","",Sys.Date(),format(Sys.time(), format = "%H:%M:%S"))

## write header to csv
write.table(header, filename, sep = ",", quote = TRUE, na = "", col.names = FALSE, row.names = FALSE)

## set variable linecount as 1
linecount <- 1

## create dataframe of unique payment numbers from source_data
payment_numbers <- data.frame(unique(source_data[,5]))

## loop through payment_numbers dataframe to process each payment into a new block in the csv
for (row in 1:nrow(payment_numbers)) {
  
  ## set variable paynum
  paynum <- payment_numbers[row,1]
  
  ## create dataframe payment_info with all rows of source_data that match paynum
  payment_info <- source_data[which(source_data$`Payment number`==paynum), ]
  
  ## populate static info for each payment into csv
  line2 <- data.frame("PMTHDR","USPS","AP6DFDL",payment_info[1,]$`Payment date`,payment_info[1,]$Amount,payment_info[1,]$`Account number`,payment_info[1,]$`Payment number`)
  write.table(line2, "output.csv", sep = ",", na = "", col.names = FALSE, row.names = FALSE, append = TRUE)
  
  line3 <- data.frame("PAYENM",payment_info[1,]$`Payee name`,"","VENDOR ID")
  write.table(line3, "output.csv", sep = ",", na = "", col.names = FALSE, row.names = FALSE, append = TRUE)
  
  line4 <- data.frame("PYEADD",payment_info[1,]$`Address line 1`,payment_info[1,]$`Address line 2`)
  write.table(line4, "output.csv", sep = ",", na = "", col.names = FALSE, row.names = FALSE, append = TRUE)
  
  line5 <- data.frame("ADDPYE","","")
  write.table(line5, "output.csv", sep = ",", na = "", col.names = FALSE, row.names = FALSE, append = TRUE)
  
######################
## Question: currently set Country to NULL if state is not populated, since Country doesn't pull to source
## data file. Would you prefer it be a holder text like "COUNTRY"?
######################
  
  ## set variable country
  if (is.null(payment_info[1,]$`State/Province`)) {
    country <- ''
  } else {
    country <- 'USA'
  }
  
  line6 <- data.frame("PYEPOS",payment_info[1,]$City,payment_info[1,]$`State/Province`,payment_info[1,]$`ZIP/Post code`,country)
  write.table(line6, "output.csv", sep = ",", na = "", col.names = FALSE, row.names = FALSE, append = TRUE)
  
  ## loop through payment_info to add one line to the csv per item
  for (i in 1:nrow(payment_info)) {
    
    ## set variable gross
    gross <- as.numeric(payment_info[i,]$`Invoice amount`)
    
    ## set variable discount
    if (is.na(payment_info[i,]$`Discount amount`)) {
      discount <- 0
    } else {
      discount <- as.numeric(payment_info[i,]$`Discount amount`)
    }
    
    ## set variable net
    net <- gross - discount
    
######################
## Question: Description is often longer than 30 characters in the source file. Should the script be set up
## to truncate to 30 characters? Notes doc indicated 30 character maximum.
## Question: Notes in example file said description should not have any commas, but with a comma delimited
## file, you can encase a string containing commas in quotations, so that it will be parsed correctly.
## Commas are allowed in Payee name (e.g., " , inc."), so is this note about no commas in Description correct?
######################

    ## set variable description
    description <- strtrim(gsub(',', '', payment_info[1,]$Description),30)
    
    ## create line item dataframe
    line <- data.frame("RMTDTL",payment_info[i,]$`Invoice number`,description,payment_info[i,]$`Invoice date`,net,gross,discount)
    write.table(line, "output.csv", sep = ",", na = "", col.names = FALSE, row.names = FALSE, append = TRUE)

  }
  
  ## update linecount
  linecount <- linecount + 5 + nrow(payment_info)
  
}

## update linecount to add footer row
linecount <- linecount + 1

## create footer dataframe and add line to output.csv
footer <- data.frame("FILTRL",linecount)
write.table(footer, "output.csv", sep = ",", na = "", col.names = FALSE, row.names = FALSE, append = TRUE)

######################
## Discussion: Error handling. Currently set the file to return an error if SourceData.xlsx is not found in
## the directory, and also if SourceData.xlsx does not have the expected column headers.
## Question: Any other errors that should be thrown?
######################

} else {
  error_message <- paste("SourcData.xlsx is not structured correctly.\n\nExpected column headers and column order:\n\n", paste(expected_cols, collapse = ', '), "\n\nPlease make sure that SourceData.xlsx has all columns, and in the correct order and then run RunTransformation.cmd again.", sep="")
  tcltk::tk_messageBox(caption = "Error", message = error_message, icon = "error", type = "ok")
}
} else {
  tcltk::tk_messageBox(caption = "Error", message = "SourcData.xlsx not found in jpmorgan-check-transform folder.\n\nPlease add SourceData.xlsx to the folder and then run RunTransformation.cmd again.", icon = "error", type = "ok")
}
