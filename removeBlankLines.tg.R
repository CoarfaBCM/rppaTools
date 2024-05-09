# function to remove blank lines from gpr files which failed consistency check due to differing number of lines

removeBlankLines <- function(consistencyFile) {
  # read in consistency check report
  consistency.df <- read.table(consistencyFile, header = T, sep = "\t")
  
  # picking out files which failed due to number of lines not matching
  keep <- consistency.df$Consistent == "FAIL (number of lines not matching)"
  failedFiles <- consistency.df$FilePath[keep]
  
  lapply(1:length(failedFiles), function(x){
    cat("####### PROCESSING FILE:", failedFiles[x],"#######\n")
    # read file
    myfile <- readLines(failedFiles[x])
    
    cat("####### Number of lines (failed file):", length(myfile),"#######\n")
    
    # remove empty lines
    myfile <- myfile[!myfile == ""]
    
    cat("####### Number of lines (new file):", length(myfile),"#######\n")
    
    # save file
    writeLines(myfile, failedFiles[x])
    cat("####### COMPLETED FILE:", failedFiles[x],"#######\n\n")
  })
}

consistencyFile = list.files(".", pattern = "rppaGprConsistencyCheck", full.names = T, recursive = T)

removeBlankLines(consistencyFile)
