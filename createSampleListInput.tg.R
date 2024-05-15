# DO NOT USE THIS SCRIPT
# THE SAMPLE LIST INPUT GENERATED HAS NUMERIC VALUES BEING TREATED AS TEXT BY EXCEL
# THIS CAUSES ISSUES WITH THE SAMPLE LIST GENERATION PYTHON SCRIPT
# CREATE THE SAMPLE LIST INPUT MANUALLY BY FOLLOWING INSTRUCTIONS FROM THE CHECKLIST

# function to remove formulas from .xlsx file and save the formula less new .xlsx file
removeFormulas <- function(inputFile,
                           sheetIdx = 1,
                           outputFile = inputFile) {
  # Loading required packages and installing ones not present
  list.of.packages <- c("openxlsx")
  new.packages <- list.of.packages[!(list.of.packages %in% installed.packages()[,"Package"])]
  if(length(new.packages)>0) {install.packages(new.packages)} else {lapply(list.of.packages, require, character.only = TRUE)}
  
  mydf <- openxlsx::read.xlsx(inputFile, sheet = sheetIdx, colNames = F, rowNames = F)
  openxlsx::write.xlsx(mydf, outputFile, colNames = F, rowNames = F, overwrite = T)
}

# function to generate input file for rppaExtractSampleListText.cc.py
createSampleListInput <- function(inputPath = ".") {
  inputFile <- list.files(path = inputPath,
                          pattern = "sample list.xlsx$|sample list histone marks.xlsx$",
                          ignore.case = T, recursive = F, full.names = T)
  inputFile <- inputFile[!grepl("[~$]",inputFile)]
  
  outputFileName <- gsub(" ","_",basename(inputFile))
  myDate <- format(Sys.Date(), "%m%d%Y")
  outputFileName <- gsub("\\.xlsx$",paste0("-noformulas-",myDate,".xlsx"),outputFileName)
  outputFile <- file.path(inputPath, outputFileName)
  
  removeFormulas(inputFile = inputFile,
                 sheetIdx = 1,
                 outputFile = outputFile)
  
  cat("##### OUTPUT FILE:",outputFile,"#####\n")
}

createSampleListInput()
