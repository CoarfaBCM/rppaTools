rppaTool <- function(inputFile,
                     outputFile,
                     abFile = NULL,
                     cv_cutoff = 0.25,
                     sampleIDRow = 1,
                     replace.cvcutoff = NULL,
                     replace.na = NULL) {
  library(readxl)
  library(openxlsx)
  library(dplyr)
  
  tempfunc <- function(inputFile,
                       wb,
                       sheetName,
                       cv_cutoff,
                       sampleIDRow,
                       replace.cvcutoff = NULL,
                       replace.na = NULL) {
    mydf <- read.xlsx(inputFile, sheet = sheetName)
    
    if (nrow(mydf) <= 1) { # if there is no normalized data
      addWorksheet(wb,paste0(sheetName,"_Median"))
      writeData(wb,paste0(sheetName,"_Median"),mydf,rowNames = FALSE)
      
      addWorksheet(wb,paste0(sheetName,"_CV"))
      writeData(wb,paste0(sheetName,"_CV"),mydf,rowNames = FALSE)
      
      return(wb)
    } else {
      if (!("AB_species" %in% names(mydf))) {
        rawdf <- mydf
        
        # trimming any trailing whitespaces from antibody IDs, antibody names and gene symbols
        rawdf[,c(1:5)] <- apply(rawdf[,c(1:5)], 2, trimws)
        
        # removing trailing characters other than antibody name
        rawdf$AB_name <- gsub("_V$|_$","", rawdf$AB_name)
        
        # adding antibody species column
        AB_species <- rep("",nrow(rawdf))
        AB_species[1] <- "AB_species"
        AB_species[grepl("_R$", rawdf$AB_name)] <- "rabbit"
        AB_species[grepl("_M$", rawdf$AB_name)] <- "mouse"
        
        rawdf$AB_name <- gsub("_R$","",rawdf$AB_name)
        rawdf$AB_name <- gsub("_M$","",rawdf$AB_name)
        # rawdf$Gene_ID <- unname(sapply(rawdf$Gene_ID, function(x){strsplit(x, ",")[[1]][1]}))
        
        # trimming any trailing whitespaces from antibody IDs, antibody names and gene symbols
        rawdf[,c(1:5)] <- apply(rawdf[,c(1:5)], 2, trimws)
        
        colnames(rawdf)[1] <- "AB_ID"
        rawdf$AB_species <- AB_species
        mycols1 <- c("AB_ID","AB_name","AB_species","Slide_file","Gene_ID", "Swiss_ID")
        mycols2 <- setdiff(names(rawdf), mycols1)
        rawdf <- rawdf[, c(mycols1, mycols2)]
        
        tempdf <- rawdf
        writeData(wb, sheetName, rbind(names(tempdf), tempdf[1,]), startRow = 1, rowNames = F, colNames = F)
        
        tempdf <- tempdf[-1,]
        tempdf[7:ncol(tempdf)] <- sapply(tempdf[7:ncol(tempdf)],as.numeric)
        writeData(wb, sheetName, tempdf, startRow = 3, rowNames = F, colNames = F)
        
        # formatting all numeric cells
        addStyle(wb = wb,
                 sheet = sheetName,
                 style = createStyle(numFmt = "0.00"),
                 cols = 7:ncol(rawdf),
                 rows = 3:(nrow(rawdf)+1),
                 gridExpand = T)
        
        # bold the first column
        addStyle(wb = wb,
                 sheet = sheetName,
                 style = createStyle(textDecoration = "bold"),
                 cols = 1,
                 rows = 1:nrow(rawdf),
                 gridExpand = T)
        
        # bold the first row
        addStyle(wb = wb,
                 sheet = sheetName,
                 style = createStyle(textDecoration = "bold"),
                 cols = 1:ncol(rawdf),
                 rows = 1,
                 gridExpand = T)
        
        mydf <- rawdf
      }
      
      mygroups <- data.frame(ID=unname(sapply(colnames(mydf)[-c(1:6)], function(x){strsplit(x,"[.]")[[1]][1]})),
                             Sample=t(mydf[1,-c(1:6)]))
      colnames(mygroups) <- c("ID","Sample")
      
      AB_species <- mydf$AB_species[-1]
      mydf <- data.frame(mydf[-1,-c(3,4,6)], row.names = NULL, check.rows = F, check.names = F)
      
      # cleaning up gene IDs
      mydf$Gene_ID <- unname(sapply(mydf$Gene_ID, function(x){strsplit(x, ",")[[1]][1]}))
      
      # trimming any trailing whitespaces from antibody IDs, antibody names and gene symbols
      mydf[,c(1:3)] <- apply(mydf[,c(1:3)], 2, trimws)
      
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
      
      mydf$AB_name <- makeUniqueNames(IDs = mydf$AB_ID, names = mydf$AB_name)
      
      mydf[,-c(1:3)] <- sapply(mydf[,-c(1:3)], as.numeric)
      
      if (!(is.null(replace.na))) {
        mydf[,-c(1:3)][is.na(mydf[,-c(1:3)])] <- replace.na
      }
      
      # tempdf <- data.frame(Sample = mygroups[,sampleIDRow], t(data.frame(mydf[,-c(1,3)], row.names = 1, check.rows = F, check.names = F)), check.names = F, check.rows = F)
      tempdf <- data.frame(mygroups[,sampleIDRow], t(mydf[,-c(1:3)]), check.names = F, check.rows = F)
      names(tempdf) <- c("Sample",mydf$AB_name)
      
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
      newdf <- newdf[unique(mygroups[,sampleIDRow]), , drop=F]
      
      cvdf <- data.frame(cvdf, row.names = 1, check.rows = F, check.names = F)
      cvdf <- cvdf[unique(mygroups[,sampleIDRow]), , drop=F]
      
      if (!(is.null(replace.cvcutoff)) & !(is.null(cv_cutoff))) {
        newdf[cvdf > cv_cutoff] <- replace.cvcutoff
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
      
      # bold the first column
      addStyle(wb = wb,
               sheet = paste0(sheetName,"_Median"),
               style = createStyle(textDecoration = "bold"),
               cols = 1,
               rows = 1:(nrow(finaldf.cv)+1),
               gridExpand = T)
      
      # bold the first row
      addStyle(wb = wb,
               sheet = paste0(sheetName,"_Median"),
               style = createStyle(textDecoration = "bold"),
               cols = 1:ncol(finaldf.cv),
               rows = 1,
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
      
      # bold the first column
      addStyle(wb = wb,
               sheet = paste0(sheetName,"_CV"),
               style = createStyle(textDecoration = "bold"),
               cols = 1,
               rows = 1:(nrow(finaldf.cv)+1),
               gridExpand = T)
      
      # bold the first row
      addStyle(wb = wb,
               sheet = paste0(sheetName,"_CV"),
               style = createStyle(textDecoration = "bold"),
               cols = 1:ncol(finaldf.cv),
               rows = 1,
               gridExpand = T)
      
      return(wb) 
    }
  }
  
  fixQIsheets <- function(wb, sheetName) {
    rawdf <- readWorkbook(wb, sheet = sheetName, colNames = T)
    
    if (nrow(rawdf) > 0) { # if there is at least one row of QI data
      # trimming any trailing whitespaces from antibody IDs, antibody names and gene symbols
      rawdf[,c(1:5)] <- apply(rawdf[,c(1:5)], 2, trimws)
      
      # removing trailing characters other than antibody name
      # adding antibody species column
      rawdf$AB_name <- gsub("_V$|_$","", rawdf$AB_name)
      
      AB_species <- rep("",nrow(rawdf))
      AB_species[grepl("_R$", rawdf$AB_name)] <- "rabbit"
      AB_species[grepl("_M$", rawdf$AB_name)] <- "mouse"
      
      rawdf$AB_name <- gsub("_R$","",rawdf$AB_name)
      rawdf$AB_name <- gsub("_M$","",rawdf$AB_name)
      
      # trimming any trailing whitespaces from antibody IDs, antibody names and gene symbols
      rawdf[,c(1:5)] <- apply(rawdf[,c(1:5)], 2, trimws)
      
      rawdf$AB_species <- AB_species
      rawdf <- rawdf[, c(1,2,ncol(rawdf),3:(ncol(rawdf)-1))]
      
      writeData(wb, sheetName, rawdf, rowNames = F, colNames = T)
      
      # formatting all numeric cells
      addStyle(wb = wb,
               sheet = sheetName,
               style = createStyle(numFmt = "0.00"),
               cols = 7:ncol(rawdf),
               rows = 2:(nrow(rawdf)),
               gridExpand = T)
      
      # bold the first column
      addStyle(wb = wb,
               sheet = sheetName,
               style = createStyle(textDecoration = "bold"),
               cols = 1,
               rows = 1:nrow(rawdf),
               gridExpand = T)
      
      # bold the first row
      addStyle(wb = wb,
               sheet = sheetName,
               style = createStyle(textDecoration = "bold"),
               cols = 1:ncol(rawdf),
               rows = 1,
               gridExpand = T)
      
      return(wb) 
    } else {
      return(wb)
    }
  }
  
  all.sheets <- excel_sheets(inputFile)
  wb <- loadWorkbook(inputFile)
  if(!("Norm_Median" %in% all.sheets)) {
    # adding AB_species column to QI and mouse_QI sheet
    wb <- fixQIsheets(wb = wb, sheetName = "QI")
    if (any(grepl("mouse",all.sheets,ignore.case = T))) {
      wb <- fixQIsheets(wb = wb, sheetName = "Mouse_QI")
    }
    
    wb1 <- tempfunc(inputFile = inputFile,
                    wb = wb,
                    sheetName = "Norm",
                    cv_cutoff = cv_cutoff,
                    sampleIDRow = sampleIDRow,
                    replace.cvcutoff = replace.cvcutoff,
                    replace.na = replace.na)
    
    if (any(grepl("mouse",all.sheets,ignore.case = T))) {
      wb1 <- tempfunc(inputFile = inputFile,
                      wb = wb1,
                      sheetName = "Mouse_Norm",
                      cv_cutoff = cv_cutoff,
                      sampleIDRow = sampleIDRow,
                      replace.cvcutoff = replace.cvcutoff,
                      replace.na = replace.na)
    }
    
    tempdf1 <- readWorkbook(wb1, sheet = "Norm_Median", rowNames = T)
    
    if (!(is.null(abFile)) & nrow(tempdf1) > 1) {
      if (any(grepl("mouse",all.sheets,ignore.case = T))) {
        tempdf2 <- readWorkbook(wb1, sheet = "Mouse_Norm_Median", rowNames = T)
        tempdf <- rbind(tempdf1,tempdf2)
      } else {
        tempdf <- tempdf1
      }
      
      new.names <- excel_sheets(abFile)
      
      for (idx in seq_along(new.names)) {
        ab_list <- read.xlsx(abFile, sheet = new.names[idx], rowNames = T)
        keep <- intersect(rownames(ab_list), rownames(tempdf))
        if (grepl("epi", new.names[idx], ignore.case = T)) {
          sheetName <- "Epi"
        } else {
          sheetName <- "Phos"
        }
        
        if (length(keep) == 0) {
          cat("### No ", sheetName," found ###\n")
          df.new <- NA
        } else {
          df.new <- tempdf[keep,]
          df.new <- data.frame(AB_ID=as.numeric(rownames(df.new)),
                               df.new,
                               check.rows = F,
                               check.names = F)
        }
        
        sheetName <- paste0("Norm_Median_", sheetName)
        addWorksheet(wb1,sheetName)
        writeData(wb1,sheetName,df.new,rowNames = FALSE)
        if (length(keep) > 0) {
          addStyle(wb = wb1,
                   sheet = sheetName,
                   style = createStyle(numFmt = "0.00"),
                   cols = 4:ncol(df.new),
                   rows = 2:(nrow(df.new)+1),
                   gridExpand = T)
          
          # bold the first column
          addStyle(wb = wb1,
                   sheet = sheetName,
                   style = createStyle(textDecoration = "bold"),
                   cols = 1,
                   rows = 1:(nrow(df.new)+1),
                   gridExpand = T)
          
          # bold the first row
          addStyle(wb = wb1,
                   sheet = sheetName,
                   style = createStyle(textDecoration = "bold"),
                   cols = 1:ncol(df.new),
                   rows = 1,
                   gridExpand = T)
        }
      }
    } else {
      df.new <- NA
      
      addWorksheet(wb1,"Norm_Median_Epi")
      writeData(wb1,"Norm_Median_Epi",df.new,rowNames = FALSE)
      
      addWorksheet(wb1,"Norm_Median_Phos")
      writeData(wb1,"Norm_Median_Phos",df.new,rowNames = FALSE)
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
        }
        
        # adding AB_species column
        # trimming any trailing whitespaces from antibody IDs, antibody names and gene symbols
        rawdf[,c(1:5)] <- apply(rawdf[,c(1:5)], 2, trimws)
        
        # removing trailing characters other than antibody name
        # adding antibody species column
        rawdf$X2 <- gsub("_V$|_$","", rawdf$X2)
        
        AB_species <- rep("",nrow(rawdf))
        AB_species[1:2] <- "AB_species"
        AB_species[grepl("_R$", rawdf$X2)] <- "rabbit"
        AB_species[grepl("_M$", rawdf$X2)] <- "mouse"
        
        rawdf$X2 <- gsub("_R$","",rawdf$X2)
        rawdf$X2 <- gsub("_M$","",rawdf$X2)
        # rawdf$Gene_ID <- unname(sapply(rawdf$Gene_ID, function(x){strsplit(x, ",")[[1]][1]}))
        
        # trimming any trailing whitespaces from antibody IDs, antibody names and gene symbols
        rawdf[,c(1:5)] <- apply(rawdf[,c(1:5)], 2, trimws)
        
        rawdf$AB_species <- AB_species
        rawdf <- rawdf[, c(1,2,ncol(rawdf),3:(ncol(rawdf)-1))]
        # done adding AB_species column
        
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
      return(wbObject)
    }
    
    wb1 <- fixRawData(wb1, "Raw")
    if (any(grepl("mouse",all.sheets,ignore.case = T))) {
      wb1 <- fixRawData(wb1, "Mouse_Raw")
      worksheetOrder(wb1) <- c(which(names(wb1)=="Legend"),
                               which(names(wb1)=="Norm"),
                               which(names(wb1)=="QI"),
                               which(names(wb1)=="Raw"),
                               which(names(wb1)=="Mouse_Norm"),
                               which(names(wb1)=="Mouse_QI"),
                               which(names(wb1)=="Mouse_Raw"),
                               which(names(wb1)=="Norm_Median"),
                               which(names(wb1)=="Norm_CV"),
                               which(names(wb1)=="Mouse_Norm_Median"),
                               which(names(wb1)=="Mouse_Norm_CV"),
                               which(names(wb1)=="Norm_Median_Phos"),
                               which(names(wb1)=="Norm_Median_Epi"))
    } else {
      worksheetOrder(wb1) <- c(which(names(wb1)=="Legend"),
                               which(names(wb1)=="Norm"),
                               which(names(wb1)=="QI"),
                               which(names(wb1)=="Raw"),
                               which(names(wb1)=="Norm_Median"),
                               which(names(wb1)=="Norm_CV"),
                               which(names(wb1)=="Norm_Median_Phos"),
                               which(names(wb1)=="Norm_Median_Epi"))
    }
    saveWorkbook(wb1,outputFile,overwrite = TRUE) 
  } else {
    saveWorkbook(wb,outputFile,overwrite = TRUE)
  }
}

runRppa <- function(inputDir = ".") {
  allfiles <- list.files(path = inputDir,
                         pattern = "_final_report.xlsx|_final_report_testing_antibodies.xlsx",
                         full.names = T,
                         recursive = T,
                         include.dirs = T,
                         all.files = F)
  allfiles <- allfiles[!grepl("[~$]",allfiles)]
  ABFile <- list.files(path= inputDir,
                       pattern = "-clean.xlsx$",
                       ignore.case = T,
                       full.names = T,
                       include.dirs = T,
                       recursive = F)
  if (length(ABFile) == 0) {ABFile <- NULL}
  
  for (i in seq_along(allfiles)) {
    inputFile <- allfiles[i]
    print(cat("##### INPUT:", inputFile,"#####\n"))
    outputFile <- gsub("_final_report.xlsx",
                       "_final_report-complete.xlsx",
                       inputFile)
    outputFile <- gsub("_final_report_testing_antibodies.xlsx",
                       "_final_report_testing_antibodies-complete.xlsx",
                       outputFile)
    rppaTool(inputFile = inputFile,
             outputFile = outputFile,
             abFile = ABFile)
    print(cat("##### OUTPUT:", outputFile,"#####\n\n"))
  }
}

runRppa()
