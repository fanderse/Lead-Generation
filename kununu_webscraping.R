################################################################################
# author: florian andersen
# last updated: 2021-Oct-25
# purpose: automatically extract (scrape) data from kununu,
# kununu scores and recommendation probability
################################################################################

library(tidyverse)
library(xlsx)
library(data.table)
library(rvest)
library(XML)
library(xml2)
library(writexl)
library(readxl)

################################################################################
# load list of company names

setwd("")


dataset <- read_excel("muenchen_company_names.xlsx", 
                                          col_names = FALSE)

# add columns for kununu score and recommendation probability
dataset$kununu <- NA
dataset$kununu <- as.character(dataset$kununu)
dataset$recomm <- NA
dataset$recomm <- as.character(dataset$recomm)
dataset$nr_scores <- NA
dataset$nr_scores <- as.character(dataset$nr_scores)

################################################################################
# specifying the url for desired website to be scraped

# general kununu url
url <- 'https://www.kununu.com/de/'

for (i in 1:nrow(dataset)) {
  # try catch disallows all error messages which interrupt the loop
  tryCatch({

  # use company names, add - between parts and combine with kununu url
string <- str_split(str_trim(dataset[i,1], side = "right"), pattern = " ")

string <- paste(string[[1]], collapse = "-")

final_url <- paste(url, string, sep = "")

# reading the HTML code from the website
download.file(final_url, destfile = "scrapedpage.html", quiet=TRUE)
content <- read_html("scrapedpage.html")

################################################################################
# kununu score
kununu_score <- html_nodes(content,'.h3-semibold-tablet')

dataset[i,2] <- html_text(kununu_score)[1]
dataset[i,3] <- html_text(kununu_score)[2]

score_count <- html_nodes(content, '.index__metric__3GbPP .text-dark-53')
dataset[i,4] <- html_text(score_count)[1]

################################################################################
}, error=function(e){})

  print(i)
}

################################################################################
# delete percentage sign from recommendation probability

dataset$recomm <- substr(dataset$recomm,1,nchar(dataset$recomm)-1)
dataset$kununu <- str_replace(dataset$kununu, ",", ".")

# delete word " Bewertungen" from scores count
dataset[dataset == "Eine Bewertung"] <- "1"
dataset$nr_scores <- str_replace(dataset$nr_scores, " Bewertungen", "")

# turn data into numerics

dataset$kununu <- as.numeric(dataset$kununu)
dataset$recomm <- as.numeric(dataset$recomm)
dataset$nr_scores <- as.integer(dataset$nr_scores)

################################################################################
# get rid of companies for which scraping was unsuccessful

dataset <- dataset[!is.na(dataset$kununu),]

################################################################################
mean <- mean(dataset$kununu, na.rm = T)

dataset_worstcompanies <- dataset[dataset$kununu <= mean,]

################################################################################
# save data

write_xlsx(dataset, "")
write_xlsx(dataset_worstcompanies, "")


#write.xlsx(dataset, "kununu_scores.xlsx", sheetName = "Main", append = FALSE)
#write.xlsx(dataset_worstcompanies, "kununu_scores.xlsx", sheetName = "Bad Companies", append = TRUE)

