# function to print out which PMTs to include in the final processing
usePMTs <- function(inputPath = ".") {
  alldirs <- basename(list.dirs(path = inputPath, recursive = F))
  gpr.dirs <- alldirs[grepl("_gpr$",alldirs)]
  
  # list all gpr files
  gpr.files <- list.files(gpr.dirs, pattern = ".gpr$", recursive = T, include.dirs = F)
  
  # extract unique pmts from all gprs files
  all.pmts <- unique(sapply(gpr.files, function(x){
    strsplit(strsplit(x[1],"/")[[1]][2], "_")[[1]][3]
  }))
  
  # extract total protein gpr files
  protein.gprs <- gpr.files[grepl("protein", gpr.files, ignore.case = T)]
  
  # extract unique pmts from all total protein gpr files
  protein.pmts <- unique(sapply(protein.gprs, function(x){
    strsplit(strsplit(x[1],"/")[[1]][2], "_")[[1]][3]
  }))
  
  # exclude PMTs only used for total proteins in the final processing
  # saving the PMTs to be included
  keep.pmts <- setdiff(all.pmts, protein.pmts)
  
  # printing out full list of PMTs
  cat("#### FULL LIST OF PMTS ####\n")
  for (i in seq_along(all.pmts)) {
    cat(all.pmts[i],"\n")
  }
  
  # printing out PMTs only used for total proteins, to be excluded in the final processing
  cat("#### PMTs ONLY USED FOR TOTAL PROTEIN ####\n")
  for (j in seq_along(protein.pmts)) {
    cat(protein.pmts[j],"\n")
  }
  
  # printing out PMTs to be used in the normalization
  cat("#### PMTs TO USE IN NORMALIZATION ####\n")
  for (k in seq_along(keep.pmts)) {
    cat(keep.pmts[k],"\n")
  }
}

usePMTs()
