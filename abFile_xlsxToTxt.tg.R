# function to appropriately format the antibody list file for RPPA normalization
abFile_xlsxToTxt <- function(workDir = ".") {
  # Loading required packages and installing ones not present
  list.of.packages <- c("tools","openxlsx")
  new.packages <- list.of.packages[!(list.of.packages %in% installed.packages()[,"Package"])]
  if(length(new.packages)>0) {install.packages(new.packages)} else {lapply(list.of.packages, require, character.only = TRUE)}
  
  # finding file with antibody list (the file name must end in "antibody list.xlsx")
  ab_file <- list.files(workDir, "antibody list.xlsx",
                        ignore.case = T, 
                        recursive = F, 
                        all.files = F, 
                        full.names = T, 
                        include.dirs = T)
  ab_file <- ab_file[!grepl("[~$]",ab_file)]
  
  # read in antibody file
  mydf <- read.xlsx(ab_file,
                    colNames = F,
                    rowNames = F,
                    check.names = F)
  
  # change cell B1 to `PI_name` for the normalization code to work; current value is `Reported Antibody Name`
  mydf[1,2] <- "PI_name"
  
  # change cell H1 to `Gene_ID` for the normalization code to work; current value is `Gene_Symbol`
  mydf[1,8] <- "Gene_ID"
  
  # # remove leading or trailing whitespaces from AB IDs and gene symbols
  # mydf[,2] <- trimws(mydf[,1])
  # mydf[,3] <- trimws(mydf[,2])
  
  # save as tab delimited text file
  myDate <- format(Sys.Date(), "%m-%d-%Y")
  saveFile <- gsub(".xlsx",paste0("_",myDate,".txt"),ab_file)
  saveFile <- gsub(" ","_",saveFile)
  write.table(mydf,saveFile,append = F,quote = F,sep = "\t",row.names = F,col.names = F)
}

abFile_xlsxToTxt()
