rppaTool <- function(inputFile,
                     outputFile,
                     cv_cutoff = 0.25,
                     sampleIDRow = 1,
                     replacement = NULL) {
  library(readxl)
  library(openxlsx)
  all.sheets <- excel_sheets(inputFile)
  wb <- loadWorkbook(inputFile)
  
  tempfunc <- function(inputFile,
                       wb,
                       sheetName,
                       cv_cutoff,
                       sampleIDRow,
                       replacement) {
    mydf <- read.xlsx(inputFile, sheet = sheetName)
    
    mygroups <- data.frame(ID=unname(sapply(colnames(mydf)[-c(1:5)], function(x){strsplit(x,"[.]")[[1]][1]})),
                           Sample=t(mydf[1,-c(1:5)]))
    
    mydf <- data.frame(mydf[-1,-c(3,5)], row.names = NULL, check.rows = F, check.names = F)
    
    # trimming any trailing whitespaces from antibody IDs, antibody names and gene symbols
    mydf[,c(1:3)] <- apply(mydf[,c(1:3)], 2, trimws)
    
    AB_species <- rep("",nrow(mydf))
    AB_species[grepl("_R_V", mydf$AB_name)] <- "rabbit"
    AB_species[grepl("_M_V", mydf$AB_name)] <- "mouse"
    
    mydf$AB_name <- gsub("_R_V","",mydf$AB_name)
    mydf$AB_name <- gsub("_M_V","",mydf$AB_name)
    mydf$Gene_ID <- unname(sapply(mydf$Gene_ID, function(x){strsplit(x, ",")[[1]][1]}))
    
    # Function to add ID to duplicate names
    makeUniqueNames <- function(IDs, names) {
      unique_names <- character(length(names))  # Initialize a vector to store unique names
      
      # Store the indices of duplicates
      duplicate_indices <- which(duplicated(names))
      
      # Iterate over each name
      for (i in seq_along(names)) {
        # Check if the name is a duplicate
        if (i %in% duplicate_indices) {
          # Find the index of the first instance of the current name
          first_instance_index <- min(which(names == names[i]))
          
          # Append the specific ID to both the first instance and the duplicates
          unique_names[first_instance_index] <- paste(names[first_instance_index], IDs[first_instance_index], sep = "_")
          unique_names[i] <- paste(names[i], IDs[i], sep = "_")
        } else {
          unique_names[i] <- names[i]  # Name is unique, no need to modify it
        }
      }
      
      return(unique_names)
    }
    
    mydf$AB_name <- makeUniqueNames(IDs = mydf$X1, names = mydf$AB_name)
    
    mydf[,-c(1:3)] <- sapply(mydf[,-c(1:3)], as.numeric)
    # mydf[is.na(mydf)] <- 1
    # mydf[,-c(1:3)][is.na(mydf[,-c(1:3)])] <- 1
    
    tempdf <- data.frame(Sample = mygroups[,sampleIDRow], t(data.frame(mydf[,-c(1,3)], row.names = 1, check.rows = F, check.names = F)), check.names = F, check.rows = F)
    
    # custom median and cv functions because the default R
    # functions drop a sample if it has all NAs for even one feature
    custom_median <- function(x) {
      if(all(is.na(x))) {
        NA
      } else {
        median(x, na.rm = T)
      }
    }
    
    custom_cv <- function(x) {
      if(all(is.na(x))) {
        NA
      } else {
        sd(x, na.rm = T)/mean(x, na.rm = T)
      }
    }
    
    newdf <- tempdf %>%
      group_by(Sample) %>%
      summarise(across(.cols = everything(), .fns = custom_median))
    cvdf <- tempdf %>%
      group_by(Sample) %>%
      summarise(across(.cols = everything(), .fns = custom_cv))
    
    # The commented out lines below drop a sample if all it's replicates have NAs for even just 1 feature
    # newdf <- aggregate(. ~ Sample, data = tempdf, FUN = function(x){median(x, na.rm = T)})
    # cvdf <- aggregate(. ~ Sample, data = tempdf, FUN = function(x){sd(x, na.rm = T)/mean(x, na.rm = T)})
    
    newdf <- data.frame(newdf, row.names = 1, check.rows = F, check.names = F)
    newdf <- newdf[unique(mygroups[,sampleIDRow]),]
    
    cvdf <- data.frame(cvdf, row.names = 1, check.rows = F, check.names = F)
    cvdf <- cvdf[unique(mygroups[,sampleIDRow]),]
    
    if (!(is.null(replacement)) & !(is.null(cv_cutoff))) {
      newdf[cvdf > cv_cutoff] <- replacement
    }
    
    finaldf <- data.frame(AB_ID = as.numeric(mydf[,1]), AB_name = colnames(newdf), AB_species = AB_species, GeneSymbol = mydf[,3], t(newdf), check.rows = F, check.names = F)
    finaldf.cv <- data.frame(AB_ID = as.numeric(mydf[,1]), AB_name = colnames(cvdf), AB_species = AB_species, GeneSymbol = mydf[,3], t(cvdf), check.rows = F, check.names = F)
    # write.xlsx(list(rppa = finaldf), paste0(outdir, "/full_aggregate_data.xlsx"), rowNames = F)
    
    addWorksheet(wb,paste0(sheetName,"_Median"))
    writeData(wb,paste0(sheetName,"_Median"),finaldf,rowNames = FALSE)
    addStyle(wb = wb,
             sheet = paste0(sheetName,"_Median"),
             style = createStyle(numFmt = "0.00"),
             cols = 5:ncol(finaldf.cv),
             rows = 2:(nrow(finaldf.cv)+1),
             gridExpand = T)
    
    addWorksheet(wb,paste0(sheetName,"_CV"))
    writeData(wb,paste0(sheetName,"_CV"),finaldf.cv,rowNames = FALSE)
    mystyle <- createStyle(fontColour = "#000000", bgFill = "#FFFF00")
    conditionalFormatting(wb,
                          paste0(sheetName,"_CV"),
                          cols = 5:ncol(finaldf.cv),
                          rows = 2:(nrow(finaldf.cv)+1),
                          rule = paste0(">",cv_cutoff),
                          style = mystyle)
    addStyle(wb = wb,
             sheet = paste0(sheetName,"_CV"),
             style = createStyle(numFmt = "PERCENTAGE"),
             cols = 5:ncol(finaldf.cv),
             rows = 2:(nrow(finaldf.cv)+1),
             gridExpand = T)
    
    return(wb)
  }
  
  wb1 <- tempfunc(inputFile = inputFile,
                  wb = wb,
                  sheetName = "Norm",
                  cv_cutoff = cv_cutoff,
                  sampleIDRow = sampleIDRow,
                  replacement = replacement)
  
  if (any(grepl("mouse",all.sheets,ignore.case = T))) {
    wb1 <- tempfunc(inputFile = inputFile,
                    wb = wb1,
                    sheetName = "Mouse_Norm",
                    cv_cutoff = cv_cutoff,
                    sampleIDRow = sampleIDRow,
                    replacement = replacement)
  }
  
  tempdf1 <- readWorkbook(wb1, sheet = "Norm_Median", rowNames = T)
  if (any(grepl("mouse",all.sheets,ignore.case = T))) {
    tempdf2 <- readWorkbook(wb1, sheet = "Mouse_Norm_Median", rowNames = T)
    tempdf <- rbind(tempdf1,tempdf2) 
  } else {
    tempdf <- tempdf1
  }
  
  ab_file <- "~/Box/gitrepos/rppaTool/data/RPPA0054_Antibody_list-clean.xlsx"
  new.names <- excel_sheets(ab_file)
  
  for (idx in seq_along(new.names)) {
    ab_list <- read.xlsx(ab_file, sheet = new.names[idx], rowNames = T)
    keep <- intersect(rownames(ab_list), rownames(tempdf))
    
    df.new <- tempdf[keep,]
    df.new <- data.frame(AB_ID=as.numeric(rownames(df.new)),
                         df.new,
                         check.rows = F,
                         check.names = F)
    
    sheetName <- paste0("Norm_Median_", new.names[idx])
    addWorksheet(wb1,sheetName)
    writeData(wb1,sheetName,df.new,rowNames = FALSE)
    addStyle(wb = wb1,
             sheet = sheetName,
             style = createStyle(numFmt = "0.00"),
             cols = 4:ncol(df.new),
             rows = 2:(nrow(df.new)+1),
             gridExpand = T) 
  }
  
  # editing the raw data sheet to remove extra negative control rows without data
  fixRawData <- function(wbObject, sheetName) {
    rawdf <- readWorkbook(wbObject, sheet = sheetName, rowNames = F, colNames = F)
    idx.AB_name <- which(rawdf[,2] == "AB_name")
    flag <- c()
    if (length(idx.AB_name) > 2) {
      for (i in 3:length(idx.AB_name)) {
        # if the ab_name row is the last row of the data, we drop it
        if (idx.AB_name[i] == nrow(rawdf)) {
          flag <-c(flag, idx.AB_name[i])
          next
        }
        
        # for the last ab_name index, we check if there is anything but negative controls
        # in the ab_id column between this row and the last row of the dataframe.
        # if there is nothing but negative controls in between, we drop this
        # last ab_name index and all rows below
        if (i == length(idx.AB_name)) {
          ab_ids <- rawdf[(idx.AB_name[i]+1):nrow(rawdf),1]
          if (all(grepl("^Ne",ab_ids))) {
            flag <- c(flag, idx.AB_name[i]:nrow(rawdf))
          }
          next
        }
        
        # if there are no rows of data between this ab_name
        # index and the next one, drop this ab_index
        if ((idx.AB_name[i]+1) == idx.AB_name[i+1]) {
          flag <- c(flag, idx.AB_name[i])
          next
        }
        
        # picking all entries from the 1st column (AB_ID) and between the
        # row after the AB_name row, since that contains the negative control,
        # and the row before the next AB_name row since that is where the antibody data ends
        ab_ids <- rawdf[(idx.AB_name[i]+1):(idx.AB_name[i+1]-1),1]
        
        # checking if there is anything but negative controls between
        # this antibody and the next antibody. if there is nothing but
        # negative controls, those rows will be dropped
        if (all(grepl("^Ne",ab_ids))) {
          flag <- c(flag, idx.AB_name[i]:(idx.AB_name[i+1]-1))
        }
      }
      
      if (!is.null(flag)) {
        rawdf <- rawdf[-flag,]
        removeWorksheet(wbObject,sheetName)
        addWorksheet(wbObject,sheetName)
        writeData(wbObject,sheetName,rawdf,rowNames = F, colNames = F)
        
        # bold the first column
        addStyle(wb = wbObject,
                 sheet = sheetName,
                 style = createStyle(textDecoration = "bold"),
                 cols = 1,
                 rows = 1:nrow(rawdf),
                 gridExpand = T)
        # bold every row containing "AB_name" in the 2nd column
        temp.idx <- which(grepl("AB_name",rawdf[,2]))
        addStyle(wb = wbObject,
                 sheet = sheetName,
                 style = createStyle(textDecoration = "bold"),
                 cols = 1:ncol(rawdf),
                 rows = temp.idx,
                 gridExpand = T)
        # bold and highlight rows with negative controls
        temp.idx <- which(grepl("^Ne",rawdf[,1]))
        addStyle(wb = wbObject,
                 sheet = sheetName,
                 style = createStyle(fgFill = "#FFFF00", fontColour = "#000000", textDecoration = "bold"),
                 cols = 1:ncol(rawdf),
                 rows = temp.idx,
                 gridExpand = T)
      }
    }
    return(wbObject)
  }
  wb1 <- fixRawData(wb1, "Raw")
  if (any(grepl("mouse",all.sheets,ignore.case = T))) {
    wb1 <- fixRawData(wb1, "Mouse_Raw")
    tempset1 <- c(1:3,which(names(wb1)=="Raw"),which(names(wb1)=="Mouse_Raw"))
    tempset2 <- setdiff(1:length(names(wb1)), tempset1)
    worksheetOrder(wb1) <- c(tempset1, tempset2)
  } else {
    tempset1 <- c(1:3,which(names(wb1)=="Raw"))
    tempset2 <- setdiff(1:length(names(wb1)), tempset1)
    worksheetOrder(wb1) <- c(tempset1, tempset2)
  }
  saveWorkbook(wb1,outputFile,overwrite = TRUE)
}

runRppa <- function(inputDir = ".") {
  allfiles <- list.files(path = inputDir,
                         pattern = "_final_report.xlsx",
                         full.names = T,
                         recursive = T,
                         include.dirs = T,
                         all.files = F)
  for (i in seq_along(allfiles)) {
    print(cat("##### START PROCESSING:", allfiles[i],"#####\n"))
    inputFile <- allfiles[i]
    outputFile <- gsub("_final_report.xlsx","_final_report-complete.xlsx",inputFile)
    rppaTool(inputFile = inputFile,
             outputFile = outputFile)
    print(cat("##### END PROCESSING:", allfiles[i],"#####\n\n"))
  }
}

runRppa()
