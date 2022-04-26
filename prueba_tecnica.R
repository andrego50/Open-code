setwd("~/Downloads")

# import libraries
library("rjson")
library("dplyr")
library('tibble')
library('data.table')
library('stringi')
library('gridExtra')

## import data sets
json_data <- fromJSON(file='aux_data.json')

# importing required packages
library(readxl)    
multiplesheets <- function(fname) {
  # getting info about all excel sheets
  sheets <- readxl::excel_sheets(fname)
  tibble <- lapply(sheets, function(x) readxl::read_excel(fname, sheet = x, col_names = FALSE))
  data_frame <- lapply(tibble, as.data.frame)
  
  # assigning names to data frames
  names(data_frame) <- sheets
  
  # print data frame
  print(data_frame)
}

# specifying the path name
path <- "prueba.xlsx"
dataset <- multiplesheets(path)
dataset <- dataset[c(2:34)]

# convert NA to 0
for (week in names(dataset)) {
  dataset[[week]][is.na(dataset[[week]])] <- 0
}

# merge variables by score: sum numeric, intestinal as mean, categorical as paste
table <- NULL
for (semana in names(dataset)) {
  prueba <- dataset[[semana]]
  # 0-1
  test_2 <- NULL
  for (i in c(7:8,10:11)) {
    test_1 <- cbind(sum(abs(as.numeric(prueba[i,c(5:8)]))),
                    sum(abs(as.numeric(prueba[i,c(11:14)]))),
                    sum(abs(as.numeric(prueba[i,c(17:20)]))),
                    sum(abs(as.numeric(prueba[i,c(23:26)]))),
                    sum(abs(as.numeric(prueba[i,c(29:32)]))))
    test_2 <- rbind(test_2,test_1)
  }
  
  # score range 0-2
  test_3 <- cbind(sum(abs(as.numeric(prueba[9,c(6:8)]))),
                  sum(abs(as.numeric(prueba[9,c(12:14)]))),
                  sum(abs(as.numeric(prueba[9,c(18:20)]))),
                  sum(abs(as.numeric(prueba[9,c(24:26)]))),
                  sum(abs(as.numeric(prueba[9,c(30:32)]))))  
  
  # score range 0-3
  test_4 <- cbind(sum(abs(as.numeric(prueba[9,c(7:8)]))),
                  sum(abs(as.numeric(prueba[9,c(13:14)]))),
                  sum(abs(as.numeric(prueba[9,c(19:20)]))),
                  sum(abs(as.numeric(prueba[9,c(25:26)]))),
                  sum(abs(as.numeric(prueba[9,c(31:32)]))))  
  
  higado_sexo <- NULL
  for (i in c(36,38)) {
    test_higado_sexo <- cbind(paste(prueba[i,c(4:8)], collapse = '-'),
                              paste(prueba[i,c(10:14)], collapse = '-'),
                              paste(prueba[i,c(16:20)], collapse = '-'),
                              paste(prueba[i,c(22:26)], collapse = '-'),
                              paste(prueba[i,c(28:32)], collapse = '-'))
    higado_sexo <- rbind(higado_sexo,test_higado_sexo)
  }
  
  peso_intestino <- NULL
  for (i in c(37,39)) {
    test_peso_intestino <- cbind(mean(abs(as.numeric(prueba[i,c(4:8)]))),
                                 mean(abs(as.numeric(prueba[i,c(10:14)]))),
                                 mean(abs(as.numeric(prueba[i,c(16:20)]))),
                                 mean(abs(as.numeric(prueba[i,c(22:26)]))),
                                 mean(abs(as.numeric(prueba[i,c(28:32)]))))
    peso_intestino <- rbind(peso_intestino,test_peso_intestino)
  }
  
  test_5 <- rbind(test_2[1:2,1:5],
                  test_4,
                  test_2[3:4,1:5],
                  test_3,
                  higado_sexo[1,1:5],
                  peso_intestino[1,1:5],
                  higado_sexo[2,1:5],
                  peso_intestino[2,1:5])
  
  table[[semana]] <- setNames(data.frame(prueba[4,2],
                                         prueba[c(7:12,36:39),2:3],
                                         test_5),
                              c('Week','Disease','Score',prueba[1,4],
                                prueba[1,10],prueba[1,16],prueba[1,22],prueba[1,28]))
}

# from json file getting correct names of farms
granja <- data.frame(cbind(paste(names(json_data[[1]][["granja"]])), 
                           paste(json_data[[1]][["granja"]])))
granja_1 <- data.frame(X1=c('Santa Defina','Sta.delfina','Sta. Delfina','Deicias'),
                       X2=c('Santa Delfina','Santa Delfina','Santa Delfina','Las Delicias PS'))
granja_2 <- rbind(granja,granja_1)

# loop over column names to put correct names
tables <- NULL
for (week in names(table)) {
  test <- table[[week]]
  col_names <- names(test)
  gsub(granja_2$X1, granja_2$X2, col_names)
  col_names <- stri_replace_all_regex(col_names,
                                    pattern=c(paste0(granja_2$X1)),
                                    replacement=c(paste0(granja_2$X2)),
                                    vectorize=FALSE)
  names(test) <- col_names
  tables[[week]] <- test
}

# export table to pdf in working directory
pdf('weeks.pdf',
    width = 15.5, # The width of the plot in inches
    height = 6)
for (week in names(tables)) {
  rownames(tables[[week]]) <- NULL
  grid.arrange(tableGrob(tables[[week]][,2:ncol(tables[[week]])]), 
               top = paste('Diseases', week),
               heights=c(2,1))
}
dev.off()
