## install packages
## install.packages("readxl")
## install.packages("tcltk")

## load package libraries
library("readxl")
library("tcltk")

## set variable wd as path to current working directory (i.e., location of the file)
wd <- getwd()
## wd <- "Y:\\Finance\\Accounts Payable\\JP Morgan Check Printing\\"

## set variable xslx_files as array of .xslx files in input folder
input_dir <- paste(wd, "\\input\\", sep="")
file_list <- list.files(input_dir, "*.xlsx$")

## error handling to check that at least one input .xlsx was provided
if (length(file_list) != 0) {

## loop through input .xlsx files to output transformed .csv
for (x in 1:length(file_list)) {
  
input_filepath <- paste(input_dir, file_list[x], sep="")
input_filename <- file_list[x]
  
## read source .xlsx file into R
source_data <- read_excel(input_filepath)

## check that source_data is structured as expected
file_cols <- colnames(source_data)
expected_cols <- c("Check format","Payment date","Amount","Account number","Payment number","Payee name","Address line 1","Address line 2","City","State/Province","ZIP/Post code","Invoice number","Description","Invoice date","Invoice amount","Discount amount","Payment type")

if (identical(file_cols,expected_cols)) {
  
## check that there are no null address rows
na_addline1 <- source_data[which(is.na(source_data$`Address line 1`)), ]
na_city <- source_data[which(is.na(source_data$City)), ]
na_state <- source_data[which(is.na(source_data$`State/Province`)), ]
na_zip <- source_data[which(is.na(source_data$`ZIP/Post code`)), ]

if (nrow(na_addline1) == 0 && nrow(na_city) == 0 && nrow(na_state) == 0 && nrow(na_zip) == 0) {

## print transformation status message
message1 <- paste("No errors in ", input_filename, ". Transforming to csv for JPMorgan.", sep="")
print(message1)
  
## set variable filename
filename <- paste(wd, "\\output\\", gsub("xlsx", "csv", input_filename), sep="")

## create header dataframe for output file
header <- data.frame("FILHDR","PWS","",Sys.Date(),format(Sys.time(), format = "%H:%M:%S"))

## write header to csv
write.table(header, filename, sep = ",", quote = TRUE, na = "", col.names = FALSE, row.names = FALSE)

## set variable linecount as 1
linecount <- 1

## set variable no_country as 0
no_country <- 0

## create dataframe of unique payment numbers from source_data
payment_numbers <- data.frame(unique(source_data[,5]))

## loop through payment_numbers dataframe to process each payment into a new block in the csv
for (row in 1:nrow(payment_numbers)) {
    
  ## set variable paynum
  paynum <- payment_numbers[row,1]
  paynum_total <- nrow(payment_numbers)
  
  ## print message on payment processing status
  message2 <- paste("Processing payment ", row, " of ", paynum_total, ".", sep="")
  print(message2)
  
  ## create dataframe payment_info with all rows of source_data that match paynum
  payment_info <- source_data[which(source_data$`Payment number`== paynum), ]
  
  ## populate static info for each payment into csv
  subtotal <- format(payment_info[1,]$Amount, digits = 2, decimal.mark = ".", nsmall = 2)
  line2 <- data.frame("PMTHDR","USPS","AP6DFDL",payment_info[1,]$`Payment date`,subtotal,payment_info[1,]$`Account number`,payment_info[1,]$`Payment number`)
  write.table(line2, filename, sep = ",", na = "", col.names = FALSE, row.names = FALSE, append = TRUE)
  
  line3 <- data.frame("PAYENM",payment_info[1,]$`Payee name`,"","VNDRID")
  write.table(line3, filename, sep = ",", na = "", col.names = FALSE, row.names = FALSE, append = TRUE)
  
  line4 <- data.frame("PYEADD",payment_info[1,]$`Address line 1`,payment_info[1,]$`Address line 2`)
  write.table(line4, filename, sep = ",", na = "", col.names = FALSE, row.names = FALSE, append = TRUE)
  
  line5 <- data.frame("ADDPYE","","")
  write.table(line5, filename, sep = ",", na = "", col.names = FALSE, row.names = FALSE, append = TRUE)
  
  ## create vector state_codes
  state_codes <- c("AL","AK","AZ","AR","CA","CO","CT","DE","DC","FL","GA","HI","ID","IL","IN","IA","KS","KY","LA","ME","MD","MA","MI","MN","MS","MO","MT","NE","NV","NH","NJ","NM","NY","NC","ND","OH","OK","OR","PA","PR","RI","SC","SD","TN","TX","UT","VT","VA","VI","WA","WV","WI","WY")
  
  ## set variable country
  if (payment_info[1,]$`State/Province` %in% state_codes) {
    country <- 'USA'
  } else {
    country <- 'XXX'
    no_country <- no_country + 1
  }
  
  line6 <- data.frame("PYEPOS",payment_info[1,]$City,payment_info[1,]$`State/Province`,payment_info[1,]$`ZIP/Post code`,country)
  write.table(line6, filename, sep = ",", na = "", col.names = FALSE, row.names = FALSE, append = TRUE)
  
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
    
    ## format numbers
    gross_formatted <- format(gross, digits = 2, decimal.mark = ".", nsmall = 2)
    discount_formatted <- format(discount, digits = 2, decimal.mark = ".", nsmall = 2)
    net_formatted <- format(net, digits = 2, decimal.mark = ".", nsmall = 2)

    ## set variable description
    description <- strtrim(gsub(',', '', payment_info[i,]$Description),30)
    
    ## create line item dataframe
    line <- data.frame("RMTDTL",payment_info[i,]$`Invoice number`,description,payment_info[i,]$`Invoice date`,net_formatted,gross_formatted,discount_formatted)
    write.table(line, filename, sep = ",", na = "", col.names = FALSE, row.names = FALSE, append = TRUE)

    }
      
  ## update linecount
  linecount <- linecount + 5 + nrow(payment_info)

}

## update linecount to add footer row
linecount <- linecount + 1

## create footer dataframe and add line to output.csv
footer <- data.frame("FILTRL",linecount)
write.table(footer, filename, sep = ",", na = "", col.names = FALSE, row.names = FALSE, append = TRUE)

## move completed source file to completed folder
destination <- paste(wd, "\\completed\\", sep="")
file.copy(input_filepath, destination)
file.remove(input_filepath)

## print message of transformation complete
message3 <- paste("Transformation of ", input_filename, " complete!", sep="")
print(message3)

## notification about Country field needing review in the output file
if (no_country > 0) {
  notice <- paste("The output file ", filename, " has ", no_country, " payments that require a review of the Country field. The placeholder 'XXX' has been populated where review is needed.", sep="")
  tcltk::tk_messageBox(caption = "Notification", message = notice, icon = "info", type = "ok")
}

} else {
  na_addline1_paynums <- data.frame(unique(na_addline1[,5]))
  na_city_paynums <- data.frame(unique(na_city[,5]))
  na_state_paynums <- data.frame(unique(na_state[,5]))
  na_zip_paynums <- data.frame(unique(na_zip[,5]))
  na_add_paynums <- unique(c(na_addline1_paynums$Payment.number, na_city_paynums$Payment.number, na_state_paynums$Payment.number, na_zip_paynums$Payment.number))
  error_message <- paste(input_filename, " contains ", nrow(na_add_paynums), " payments that are missing address information in the Address Line 1, City, State/Provence, and/or Zip/Postal Code columns.\n\nThe payment number(s) with missing address information are:\n\n", paste(na_add_paynums, collapse = ', '), "\n\nPlease make sure that ", input_filename, " has address information populated for all payments in Address Line 1, City, State/Provence, and Zip/Postal Code fields, then run RunTransformation.cmd again.\n\nIf other input files were staged for this transformation, the process will continue and transform to csv those input files that are well-formed.", sep="")
  tcltk::tk_messageBox(caption = "Error", message = error_message, icon = "error", type = "ok")
} ## end of missing address information error handling

} else {
  error_message <- paste(input_filename, " is not structured correctly.\n\nExpected column headers and column order:\n\n", paste(expected_cols, collapse = ', '), "\n\nPlease make sure that ", input_filename, " has all columns, and in the correct order and then run RunTransformation.cmd again.\n\nIf other input files were staged for this transformation, the process will continue and transform to csv those input files that are well-formed.", sep="")
  tcltk::tk_messageBox(caption = "Error", message = error_message, icon = "error", type = "ok")
} ## end of incorrect file structure error handling

} } else {
  tcltk::tk_messageBox(caption = "Error", message = "Input .xslx file(s) not found in the 'input' folder.\n\nPlease add Financial Edge export .xlsx file(s) to the folder and then run RunTransformation.cmd again.", icon = "error", type = "ok")
} ## end of no input files error handling

  