rppaTool <- function(inputFile,
                     outputFile,
                     cv_cutoff = 0.25,
                     sampleIDRow = 2,
                     replacement = NULL) {
  library(readxl)
  library(openxlsx)
  all.sheets <- excel_sheets(inputFile)
  wb <- loadWorkbook(inputFile)
  
  mydf <- read.xlsx(inputFile, sheet = "Norm")
  
  mygroups <- data.frame(colnames(mydf)[-c(1:5)], t(mydf[1,-c(1:5)]))
  names(mygroups) <- c("ID", "Sample")
  
  mydf <- data.frame(mydf[-1,-c(1,3,5)], row.names = NULL, check.rows = F, check.names = F)
  mydf[,1] <- gsub("_R_V","",mydf[,1])
  mydf[,2] <- unname(sapply(mydf[,2], function(x){strsplit(x, ",")[[1]][1]}))
  
  mydf[,-c(1,2)] <- sapply(mydf[,-c(1,2)], as.numeric)
  mydf[is.na(mydf)] <- 1
  
  tempdf <- data.frame(Sample = mygroups$Sample, t(data.frame(mydf[,-2], row.names = 1, check.rows = F, check.names = F)), check.names = F, check.rows = F)
  newdf <- aggregate(. ~ Sample, data = tempdf, FUN = median)
  cvdf <- aggregate(. ~ Sample, data = tempdf, FUN = function(x){sd(x)/mean(x)})
  cvdf.1 <- cvdf[,-1]
  if (!(is.null(replacement))) {
    newdf.1 <- newdf[,-1]
    newdf.1[cvdf.1 > cv_cutoff] <- replacement
    newdf[,-1] <- newdf.1
  }
  
  if (sampleIDRow == 1) {
    uniqueNames <- data.frame(ID = unname(sapply(mygroups$ID, function(x){strsplit(x,"[.]")[[1]][1]})),
                              sample = mygroups$Sample)
    uniqueNames <- uniqueNames[!duplicated(uniqueNames$sample),]
    uniqueNames <- uniqueNames[match(newdf$Sample, uniqueNames$sample),]
    if (identical(newdf$Sample, uniqueNames$sample)) {
      newdf$Sample <- uniqueNames$ID 
    }
  }
  
  # write.xlsx(list(rppa = newdf), paste0(outdir, "/aggregate_data.xlsx"), rowNames = F) 
  newdf <- data.frame(newdf, row.names = 1, check.rows = F, check.names = F)
  rownames(cvdf.1) <- rownames(newdf)
  
  finaldf <- data.frame(GeneSymbol = mydf[,2], AB_name = colnames(newdf), t(newdf), check.rows = F, check.names = F)
  finaldf.cv <- data.frame(GeneSymbol = mydf[,2], AB_name = colnames(cvdf.1), t(cvdf.1), check.rows = F, check.names = F)
  # write.xlsx(list(rppa = finaldf), paste0(outdir, "/full_aggregate_data.xlsx"), rowNames = F)
  
  addWorksheet(wb,"Norm_Median")
  writeData(wb,"Norm_Median",finaldf,rowNames = FALSE)
  addStyle(wb = wb,
           sheet = "Norm_Median",
           style = createStyle(numFmt = "0.00"),
           cols = 3:ncol(finaldf.cv),
           rows = 2:nrow(finaldf.cv),
           gridExpand = T)
  
  addWorksheet(wb,"Norm_CV")
  writeData(wb,"Norm_CV",finaldf.cv,rowNames = FALSE)
  mystyle <- createStyle(fontColour = "#000000", bgFill = "#FFFF00")
  conditionalFormatting(wb,
                        "Norm_CV",
                        cols = 3:ncol(finaldf.cv),
                        rows = 2:nrow(finaldf.cv),
                        rule = paste0(">",cv_cutoff),
                        style = mystyle)
  addStyle(wb = wb,
           sheet = "Norm_CV",
           style = createStyle(numFmt = "PERCENTAGE"),
           cols = 3:ncol(finaldf.cv),
           rows = 2:nrow(finaldf.cv),
           gridExpand = T)
  if (any(grepl("mouse",all.sheets,ignore.case = T))) {
    sheetName <- "Mouse_Norm"
    mydf <- read.xlsx(inputFile, sheet = sheetName)
    
    mygroups <- data.frame(colnames(mydf)[-c(1:5)], t(mydf[1,-c(1:5)]))
    names(mygroups) <- c("ID", "Sample")
    
    mydf <- data.frame(mydf[-1,-c(1,3,5)], row.names = NULL, check.rows = F, check.names = F)
    mydf[,1] <- gsub("_R_V","",mydf[,1])
    mydf[,2] <- unname(sapply(mydf[,2], function(x){strsplit(x, ",")[[1]][1]}))
    
    mydf[,-c(1,2)] <- sapply(mydf[,-c(1,2)], as.numeric)
    mydf[is.na(mydf)] <- 1
    
    tempdf <- data.frame(Sample = mygroups$Sample, t(data.frame(mydf[,-2], row.names = 1, check.rows = F, check.names = F)), check.names = F, check.rows = F)
    newdf <- aggregate(. ~ Sample, data = tempdf, FUN = median)
    cvdf <- aggregate(. ~ Sample, data = tempdf, FUN = function(x){sd(x)/mean(x)})
    cvdf.1 <- cvdf[,-1]
    if (!(is.null(replacement))) {
      newdf.1 <- newdf[,-1]
      newdf.1[cvdf.1 > cv_cutoff] <- replacement
      newdf[,-1] <- newdf.1
    }
    
    if (sampleIDRow == 1) {
      uniqueNames <- data.frame(ID = unname(sapply(mygroups$ID, function(x){strsplit(x,"[.]")[[1]][1]})),
                                sample = mygroups$Sample)
      uniqueNames <- uniqueNames[!duplicated(uniqueNames$sample),]
      uniqueNames <- uniqueNames[match(newdf$Sample, uniqueNames$sample),]
      if (identical(newdf$Sample, uniqueNames$sample)) {
        newdf$Sample <- uniqueNames$ID 
      }
    }
    
    # write.xlsx(list(rppa = newdf), paste0(outdir, "/aggregate_data.xlsx"), rowNames = F) 
    newdf <- data.frame(newdf, row.names = 1, check.rows = F, check.names = F)
    rownames(cvdf.1) <- rownames(newdf)
    
    finaldf <- data.frame(GeneSymbol = mydf[,2], AB_name = colnames(newdf), t(newdf), check.rows = F, check.names = F)
    finaldf.cv <- data.frame(GeneSymbol = mydf[,2], AB_name = colnames(cvdf.1), t(cvdf.1), check.rows = F, check.names = F)
    # write.xlsx(list(rppa = finaldf), paste0(outdir, "/full_aggregate_data.xlsx"), rowNames = F)
    
    addWorksheet(wb,paste0(sheetName,"_Median"))
    writeData(wb,paste0(sheetName,"_Median"),finaldf,rowNames = FALSE)
    addStyle(wb = wb,
             sheet = paste0(sheetName,"_Median"),
             style = createStyle(numFmt = "0.00"),
             cols = 3:ncol(finaldf.cv),
             rows = 2:nrow(finaldf.cv),
             gridExpand = T)
    
    addWorksheet(wb,paste0(sheetName,"_CV"))
    writeData(wb,paste0(sheetName,"_CV"),finaldf.cv,rowNames = FALSE)
    mystyle <- createStyle(fontColour = "#000000", bgFill = "#FFFF00")
    conditionalFormatting(wb,
                          paste0(sheetName,"_CV"),
                          cols = 3:ncol(finaldf.cv),
                          rows = 2:nrow(finaldf.cv),
                          rule = paste0(">",cv_cutoff),
                          style = mystyle)
    addStyle(wb = wb,
             sheet = paste0(sheetName,"_CV"),
             style = createStyle(numFmt = "PERCENTAGE"),
             cols = 3:ncol(finaldf.cv),
             rows = 2:nrow(finaldf.cv),
             gridExpand = T)
  }
  
  saveWorkbook(wb,outputFile,overwrite = TRUE)
}

runRppa <- function(inputDir = ".") {
  allfiles <- list.files(path = inputDir,
                         pattern = "_final_report.xlsx",
                         full.names = T,
                         recursive = T,
                         include.dirs = T)
  for (i in seq_along(allfiles)) {
    inputFile <- allfiles[i]
    outputFile <- gsub("_final_report.xlsx","_final_report-complete.xlsx",inputFile)
    rppaTool(inputFile = inputFile,
             outputFile = outputFile)
  }
}

runRppa()