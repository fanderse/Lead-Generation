# author: florian andersen
# last updated: 2021-Sep-29
# purpose: clean xlsx export from anzeigendaten.de. 
# The script separates those vacancies with AP and personal mail from those without. 
# It excludes all vacancies with irrelevant positions, i.e. junior or trainee positions.
# It also separates all companies which our consultants already are in contact with.
# After this script has run, manual data cleaning should be minimal.
########################################################

# load packages
library(tidyverse)
library(xlsx)
library(stringi)
library(data.table)
library(svDialogs)

# set working directory (i.e. where the xlsx is stored)
setwd("")

################################################################################
# load xlsx

# the script asks you which file you would like to load
source_name <- dlgInput("Name der xlsx von anzeigendaten (mit .xlsx am Ende bitte):", 
                        Sys.info()["user"])$res

# benchmarking
start_time <- Sys.time()

dataset <- read.xlsx2(source_name,
                      sheetIndex = 1, 
                      startRow = 5)

# load consultant target list 
target_list <- read.xlsx2("Target Liste Essen komplett.xlsx",
                          sheetIndex = 1, 
                          header = TRUE)

# empty cells are set to NA for easier handling in R
dataset <- dataset %>% mutate_all(na_if,"")

# initiate empty data frame for storage of entries
# w/o AP or personal mail address
dataset_ohneAP <- data.frame(Date=as.Date(character()),
                             File=character(), 
                             User=character(), 
                             stringsAsFactors=FALSE) 

# initiate empty data frame for storage of entries
# our consultants already are in contact with
dataset_targetlist <- data.frame(Date=as.Date(character()),
                             File=character(), 
                             User=character(), 
                             stringsAsFactors=FALSE) 

################################################################################
# remove duplicate APs
dataset <- distinct(dataset, Ansprechperson..AP....Firma, .keep_all = TRUE)

# search for intern, junior, and trainee positions and delete them
dataset <- dataset[!dataset$Position %ilike% "Junior", ]
dataset <- dataset[!dataset$Position %ilike% "Intern", ]
dataset <- dataset[!dataset$Position %ilike% "Trainee", ]
dataset <- dataset[!dataset$Position %ilike% "Prakti", ]

################################################################################
# move entries w/o AP or mail address to new dataframe and delete from original
dataset_ohneAP <- subset(dataset, is.na(dataset$Ansprechperson..AP....Firma))
dataset_ohneAP <- rbind(dataset_ohneAP, subset(dataset, 
                                               is.na(dataset$E.Mail...AP.Firma)))

dataset <- subset(dataset, !is.na(dataset$Ansprechperson..AP....Firma))
dataset <- subset(dataset, !is.na(dataset$E.Mail...AP.Firma))

# move entries with genereic mail addresses to separate dataframe
dataset_ohneAP <- rbind(dataset_ohneAP, dataset[dataset$E.Mail...AP.Firma %ilike% "recruit", ])
dataset_ohneAP <- rbind(dataset_ohneAP, dataset[dataset$E.Mail...AP.Firma %ilike% "info@", ])
dataset_ohneAP <- rbind(dataset_ohneAP, dataset[dataset$E.Mail...AP.Firma %ilike% "bewerb", ])
dataset_ohneAP <- rbind(dataset_ohneAP, dataset[dataset$E.Mail...AP.Firma %ilike% "career", ])
dataset_ohneAP <- rbind(dataset_ohneAP, dataset[dataset$E.Mail...AP.Firma %ilike% "karriere", ])
dataset_ohneAP <- rbind(dataset_ohneAP, dataset[dataset$E.Mail...AP.Firma %ilike% "ausbild", ])
dataset_ohneAP <- rbind(dataset_ohneAP, dataset[dataset$E.Mail...AP.Firma %ilike% "jobs", ])
dataset_ohneAP <- rbind(dataset_ohneAP, dataset[dataset$E.Mail...AP.Firma %ilike% "job", ])
dataset_ohneAP <- rbind(dataset_ohneAP, dataset[dataset$E.Mail...AP.Firma %ilike% "bewerb", ])
dataset_ohneAP <- rbind(dataset_ohneAP, dataset[dataset$E.Mail...AP.Firma %ilike% "personal", ])
dataset_ohneAP <- rbind(dataset_ohneAP, dataset[dataset$E.Mail...AP.Firma %ilike% "human", ])

# remove those same entries from original data frame
dataset <- dataset[!dataset$E.Mail...AP.Firma %ilike% "recruit", ]
dataset <- dataset[!dataset$E.Mail...AP.Firma %ilike% "info@", ]
dataset <- dataset[!dataset$E.Mail...AP.Firma %ilike% "bewerb", ]
dataset <- dataset[!dataset$E.Mail...AP.Firma %ilike% "career", ]
dataset <- dataset[!dataset$E.Mail...AP.Firma %ilike% "karriere", ]
dataset <- dataset[!dataset$E.Mail...AP.Firma %ilike% "ausbild", ]
dataset <- dataset[!dataset$E.Mail...AP.Firma %ilike% "jobs", ]
dataset <- dataset[!dataset$E.Mail...AP.Firma %ilike% "job", ]
dataset <- dataset[!dataset$E.Mail...AP.Firma %ilike% "bewerb", ]
dataset <- dataset[!dataset$E.Mail...AP.Firma %ilike% "personal", ]
dataset <- dataset[!dataset$E.Mail...AP.Firma %ilike% "human", ]

################################################################################
# compare all entries from both datasets with consultant target list

for (i in 1:nrow(dataset)) {
  for (j in 1:nrow(target_list)) {
    if (dataset[i, 3] %ilike% target_list[j, 1]) {
      dataset_targetlist <- rbind(dataset_targetlist, 
                                  cbind(dataset[i, ], target_list[j, ]))
      dataset <- dataset[-i, ]
      
    }
    if (dataset_ohneAP[i, 3] %ilike% target_list[j, 1]) {
      dataset_targetlist <- rbind(dataset_targetlist, 
                                  cbind(dataset_ohneAP[i, ], target_list[j, ]))
      dataset_ohneAP <- dataset_ohneAP[-i,]
    }
    
  }
}

# benchmarking
end_time <- Sys.time()
end_time - start_time

################################################################################
# write all three dataframes into xlsx sheets
# the file names need to be changed as desired

destination_name <- dlgInput("Name der Zieldatei wählen (mit .xlsx am Ende bitte):", 
                             Sys.info()["user"])$res

# main data (vacancies with AP and personal mail)
write.xlsx(dataset, destination_name, sheetName = "Main", append = FALSE)

# secondary data: vanacies w/o AP or personal mail
write.xlsx(dataset_ohneAP, destination_name, sheetName = "ohne AP or generic mail", 
           append = TRUE)

# tertiary data: vacancies w/ APs already known by our consultants
write.xlsx(dataset_targetlist, destination_name, sheetName = "in Target List", append = TRUE)

################################################################################
# create another .xlsx with one sheet for each consultant 
# (can be sent to consultants immediately imho)

targetlist_name <- dlgInput("Name der Zieldatei wählen (mit .xlsx am Ende bitte):", 
                            Sys.info()["user"])$res

consultants <- unique(dataset_targetlist$Owner)
for (c in consultants) {
write.xlsx(subset(dataset_targetlist, dataset_targetlist$Owner == c), 
           targetlist_name, sheetName = c, append = TRUE)
}

################################################################################




