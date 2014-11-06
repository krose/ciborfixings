

setwd("c:/users/kro/documents/github/ciborfixings/")

library(xlsx)
library(lubridate)
library(stringr)
library(dplyr)

# ### Build sequence of dates to try
# try_dates <- seq.Date(from = as.Date(ymd("2004-02-01")), to = as.Date(ymd("2014-11-05")), by = 1)
# 
# ### Define the base url
# base_url <- "http://www.finansraadet.dk/Historical%20Rates/cibor/"
# 
# 
# 
# #######################################################
# ## Download files
# ########################################################
# 
# ### Loop over try_dates and try to download a file
# for(i in seq_along(try_dates)){
#   
#   try_url_xlsx <- paste0(base_url, as.character(year(try_dates[i])), "/", as.character(try_dates[i]), ".xlsx")
#   try_url_xls <- paste0(base_url, as.character(year(try_dates[i])), "/", as.character(try_dates[i]), ".xls")
#   
#   try(download.file(url = try_url_xlsx, destfile = paste0("./data/", try_dates[i], ".xlsx"), mode = "wb"), silent = TRUE)
#   try(download.file(url = try_url_xls, destfile = paste0("./data/", try_dates[i], ".xls"), mode = "wb"), silent = TRUE)
#   
#   Sys.sleep(time = 0.5)
# }
# 
# ###########################################################
# ## Remove files with file size == 0
# ############################################################
# 
# xl_file_names <- dir(path = "./data/")
# 
# for(i in seq_along(xl_file_names)){
#   
#   if(file.info(paste0("./data/", xl_file_names[i])$size) == 0){
#     file.remove(paste0("./data/", xl_file_names[i]))
#   }
# }
# 

###########################################################
## Read first column to figure out which rows to download
############################################################


### only read up until 2011-05-07 as they move to a very old
### xls file format.
## minus 59 because of NAs
xl_file_names <- dir(path = "./data/")
xl_file_names <- xl_file_names[1:(1818-59)]

## Create list for data
first_col <- vector(mode = "list", length = length(xl_file_names))

## Bad names was used in the exploratory phase
# bad_names <- c("danmarks nationalbank", "handelsafdelingen", "ciborfixing den", "stiller", "nordic operations", "nasdaq omx", "mail",
#                "ciborfixing", "fax", "to:", "market operations", "NA", "yours sincerely",
#                "This sheet containing rates from the CIBOR reporting banks shall be regarded as an integrated part of the control procedure",
#                "of CIBOR. The rates do not represent the final CIBOR rates and are not to be published by any CIBOR reporting bank, nor",
#                "by Finansraadet (Danish Bankers Association) or by anyone else before the official CIBOR rates are available. The official")

## These bank names are used going forward.
good_names <- unique(tolower(c("Amtssparekassen Fyn", "Danske Bank", "Jyske Bank", "HSH Nordbank", 
                               "Nordea", "Nykredit Bank", "Spar Nord Bank", "Sydbank", "FIXING", 
                               "Fionia Bank", "ABN Amro Bank", "Barclays Capital", 
                               "Deutsche Bank", "Royal Bk of Scotland", "DANSKE BANK", "BARCLAYS", "Deutsche", "JYSKE BANK", "NORDEA", 
                               "NYKREDIT", "RBS FM", "SYDBANK")))

## Read the first column of every file into a list object
## The column is being sequenced to single out the rows we want
## and matched based on the good_names
## I also do a bit of cleaning up like setting character to lower
for(i in seq_along(xl_file_names)){
  
  ## Print the file name being read
  print(xl_file_names[i])
  
  ## Read file
  first_col[[i]] <- try(read.xlsx2(file = paste0("./data/", xl_file_names[i]),sheetIndex = 1, startRow = 1, colIndex = 1, endRow = 70, as.data.frame = TRUE, header = FALSE, colClasses = "character", stringsAsFactors = FALSE), silent = FALSE)
  
  ## Make sequence
  first_col[[i]]$row_number <- seq_along(first_col[[i]][,1])
  
  ## Add file name
  first_col[[i]]$file_name <- xl_file_names[i]
  
  ## set characters to lower
  first_col[[i]]$X1 <- tolower(str_trim(first_col[[i]]$X1))
  
  ## Test for good names
  test_good <- first_col[[i]]$X1 %in% good_names
  
  ## filter and keep only the good names
  first_col[[i]] <- first_col[[i]][test_good, ]
  
  ### Used in the exploratory phase
#   ## get length of each character string
#   strlen <- str_length(string = first_col[[i]][, 1])
#   
#   ## remove rows where length > 0 or bigger than 50 or is na
#   first_col[[i]] <- first_col[[i]][strlen > 0,]
#   first_col[[i]] <- first_col[[i]][strlen < 50,]
#   first_col[[i]] <- first_col[[i]][!is.na(first_col[[i]]$X1),]
#   
#   ### test for bad names
#   test_bad <- tolower(first_col[[i]]$X1) %in% bad_names
#   ## remove bad names
#   first_col[[i]] <- first_col[[i]][ !test_bad, ]
#   
#   test_bad <- str_detect(first_col[[i]]$X1, "[[:digit:]]")
#   first_col[[i]] <- first_col[[i]][!test_bad,]
  
}

## Take every list object and rowbind them all
test_ob <- do.call("rbind", first_col)

## Group_by file_name and find the highest and lowest row number
## we can use for reading the the actual fixing and discard the rest
test_ob <- test_ob %>% group_by(file_name) %>%
  summarise(min_row = min(row_number), max_row = max(row_number)) %>% ungroup

##############################################################
## Read the fixings
###########################################################

## create list object to save fixings into
curve_values <- list()

## Loop over each row in test_ob and read the file name using these variables.
for(i in seq_along(test_ob$file_name)){
  
  ## Print out where you are
  print(test_ob$file_name[i])
  
  ## Read the file
  curve_values[[i]] <- read.xlsx2(file = paste0("./data/", test_ob$file_name[i]), sheetIndex = 1, startRow = test_ob$min_row[i], endRow = test_ob$max_row[i], as.data.frame = TRUE, header = FALSE, stringsAsFactors = FALSE)
  
  ## Set first column tolower characters
  curve_values[[i]]$X1 <- tolower(curve_values[[i]]$X1)
}


##########################################################
## clean and format the fixings
##########################################################

## Remove empty columns by looping through each reading and its columns.
for(i in seq_along(curve_values)){
  
  for(n in length(names(curve_values[[i]])):1){
    
    if(is.character(curve_values[[i]][, n])){
      
      test <- sum(str_length(string = curve_values[[i]][, n])) < 1
      
      if(test){
        curve_values[[i]] <- curve_values[[i]][, -n]
      }
    }
  }
}

## Remove blank rows from the file 
for(i in seq_along(curve_values)){
  
  test <- str_length(curve_values[[i]]$X1) > 0
  
  curve_values[[i]] <- curve_values[[i]][test,]
}


## Look at the length of each column and keep only 
## the ones with length == 15
test_ob$col_len <- sapply(curve_values, FUN = function(x) length(names(x)),simplify = TRUE)

## Keep only file readings with col_len == 15
## THis means that data from 2005-04-01 to 2011-04-01
## is retained
curve_values <- curve_values[test_ob$col_len == 15]
test_ob <- test_ob[test_ob$col_len == 15, ]

## Set colnames for fixing period 
curve_col_names <- c("bank", "d7", "d14", "d30", "d60", "d90", "d120", "d150", "d180", 
                     "d210", "d240", "d270", "d300", "d330", "d360")

## Set colnames
for(i in seq_along(curve_values)){
  
  names(curve_values[[i]]) <- curve_col_names
}


## set fixings as numeric instead of character
for(i in seq_along(curve_values)){
  for(n in seq_along(names(curve_values[[i]]))){
    if(n == 1){
      next
    } else {
      curve_values[[i]][,n] <- as.numeric(curve_values[[i]][,n])
    }
  }
}


## Set date
for(i in seq_along(curve_values)){
  curve_values[[i]]$date <- str_replace(string = test_ob$file_name[i], pattern = ".xls", replacement = "")
  curve_values[[i]]$date <- str_replace(string = curve_values[[i]]$date, pattern = "x", replacement = "")
  curve_values[[i]]$date <- ymd(curve_values[[i]]$date)
}

## merge all the list objects
curve_values <- do.call(rbind, curve_values)

## set the rows in the "correct" order
curve_values <- curve_values %>% select(date, bank, d7:d360)

## write tidy data to the disk
write.csv2(curve_values, file = "cibor_csv2.csv")

