# WARNING: DO NOT OPEN THE OUTPUT FILE IN EXCEL AS EXCEL CONVERTS THE INFO IN THE NORMALIZATION COLUMN TO DATES (eg: 1-12 gets converted to Jan-12)

# function to generate input file for rppaExtractSampleListText.cc.py
createSlideTableInput <- function(inputPath = ".") {
  # Loading required packages and installing ones not present
  list.of.packages <- c("openxlsx")
  new.packages <- list.of.packages[!(list.of.packages %in% installed.packages()[,"Package"])]
  if(length(new.packages)>0) {install.packages(new.packages)} else {lapply(list.of.packages, require, character.only = TRUE)}
  
  inputFile <- list.files(path = inputPath,
                          pattern = "Total protein slide table.xlsx$|Total protein slide table histone marks.xlsx$",
                          ignore.case = T, recursive = F, full.names = T)
  inputFile <- inputFile[!grepl("[~$]",inputFile)]
  mydf <- openxlsx::read.xlsx(inputFile, sheet = 1, colNames = F, rowNames = F)
  
  # slideRow <- which(unname(apply(mydf,1,function(x){any(grepl("Slide#",x,ignore.case = T))})))
  slideIdx <- which(sapply(mydf, function(x){grepl("Slide#",x,ignore.case = T)}), arr.ind = T)
  
  if (slideIdx[1] > 1) {
    mydf <- mydf[-c(1:(slideIdx[1]-1)),]
  }
  
  if (slideIdx[2] > 1) {
    mydf <- mydf[,-c(1:(slideIdx[2]-1))]
  }
  
  # mydf[2:nrow(mydf),ncol(mydf)] <- paste0('"',mydf[2:nrow(mydf),ncol(mydf)],'"') # wrapping in quotes to prevent excel from reading as dates
  # tried using above code to prevent excel from interpreting the last column as dates (eg: 1-12 gets converted to Jan-12)
  # tried several other approaches as well but to no avail
  # solution: DO NOT OPEN THIS FILE IN EXCEL, ONLY OPEN IN A TEXT EDITOR
  
  outputFileName <- gsub(" ","_",basename(inputFile))
  myDate <- format(Sys.Date(), "%m%d%Y")
  outputFileName <- gsub("\\.xlsx$",paste0("_",myDate,".txt"),outputFileName)
  outputFile <- file.path(inputPath, outputFileName)
  
  write.table(mydf,outputFile,append = F,quote = F,sep = "\t",row.names = F,col.names = F)
  
  cat("##### OUTPUT FILE:",outputFile,"#####\n")
}

createSlideTableInput()

# WARNING: DO NOT OPEN THE OUTPUT FILE IN EXCEL AS EXCEL CONVERTS THE INFO IN THE NORMALIZATION COLUMN TO DATES (eg: 1-12 gets converted to Jan-12)