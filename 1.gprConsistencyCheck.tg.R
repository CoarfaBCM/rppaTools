# Function to check if text file is utf-8 encoded
check_utf8_encoding <- function(file_path) {
  tryCatch({
    readLines(file_path, encoding = "UTF-8")
    TRUE  # If readLines succeeds, it's UTF-8
  }, error = function(e) {
    FALSE  # If an error occurs, not UTF-8
  })
}

# Function to extract information from a line
extract_info <- function(line, search_pattern) {
  if (grepl("tif",search_pattern,ignore.case = T)) {
    lineFlag <- "ImageFiles="
  } else {
    lineFlag <- "JpegImage="
  }
  if (grepl(lineFlag, line)) {
    parts <- unlist(strsplit(line, "="))
    if (length(parts) == 2) {
      info_parts <- unlist(strsplit(parts[2], "\\\\"))
      extracted_info <- tail(info_parts, 1)
      extracted_info <- paste0(strsplit(extracted_info, search_pattern)[[1]][1], search_pattern)
      return(extracted_info)
    }
  }
  return(NULL)
}

# Function to read info from one gpr file at a time
readOneGpr <- function(input_file) {
  cat("##### PROCESSING:", input_file, "#####\n")
  
  # Initialize variables to store the extracted information
  tiff_file <- NULL
  jpeg_file <- NULL
  
  # Read the .gpr file line by line
  con <- file(input_file, "r")
  while (length(line <- readLines(con, n = 1)) > 0) {
    # Extract information for "ImageFiles="
    tiff_info <- extract_info(line, ".tif")
    if (!is.null(tiff_info)) {
      tiff_file <- tiff_info
    }
    
    # Extract information for "JpegImage="
    jpeg_info <- extract_info(line, ".jpg")
    if (!is.null(jpeg_info)) {
      jpeg_file <- jpeg_info
    }
    
    # Exit loop if both tiff and jpeg information have been found
    if (!is.null(tiff_file) && !is.null(jpeg_file)) {
      break
    }
  }
  
  # Close the file
  close(con)
  
  # Read in gpr file
  temp <- readLines(input_file, warn = FALSE) # this part also strips windows style line endings (\r\n) and leaves behind unix style line endings (\n)
  
  # Check for F532 field in total protein gpr files and F635 field in other antibody gpr files
  if (grepl("protein", input_file, ignore.case = T)) {
    F532_F635 <- any(grepl("F532",temp))
  } else {
    F532_F635 <- any(grepl("F635",temp))
  }
  
  # Count number of quotation marks to be removed
  quotesRemoved <- sum(grepl('"',temp))
  
  # Remove quotation marks
  temp.final <- gsub('"', '', temp)
  
  # Count number of empty blank lines to be removed
  blankLinesRemoved1 <- sum(temp.final == "")
  blankLines <- sapply(1:length(temp.final), function(x){gsub("\t","",temp.final[x]) == ""})
  blankLinesRemoved2 <- sum(blankLines)
  blankLinesRemoved <- blankLinesRemoved1 + blankLinesRemoved2
  
  # Remove blank lines
  temp.final <- temp.final[!blankLines]
  temp.final <- temp.final[!(temp.final == "")]
  
  # Save number of lines
  num_lines <- length(temp.final)
  
  # Save final gpr file after removing quotation marks and blank lines
  writeLines(temp.final, input_file) # windows may need edited command: writeLines(temp.final, input_file, eol = "\r\n")
  
  # Check if saved gpr file is utf-8 encoded or not
  checkEncoding <- check_utf8_encoding(input_file)
  
  full_info <- c(input_file, basename(input_file), tiff_file, jpeg_file, F532_F635, quotesRemoved, blankLinesRemoved, num_lines, checkEncoding)
  
  cat("##### COMPLETED:", input_file, "#####\n")
  
  return(full_info)
}

# Main function
gprConsistencyCheck <- function(input_path = ".",
                                out_suffix = NULL) {
  # Loading required packages and installing ones not present
  list.of.packages <- c("tools","openxlsx")
  new.packages <- list.of.packages[!(list.of.packages %in% installed.packages()[,"Package"])]
  if(length(new.packages)>0) {install.packages(new.packages)} else {lapply(list.of.packages, require, character.only = TRUE)}
  
  finaldf <- NULL
  for (myfile in list.files(input_path,pattern = "\\.gpr$",full.names = T,recursive = T)) {
    finaldf <- rbind(finaldf, readOneGpr(myfile))
  }
  finaldf <- as.data.frame(finaldf)
  finaldf$V8 <- as.numeric(finaldf$V8)
  
  # function to get statistical mode of the input
  getmode <- function(myvec) {
    uniqv <- unique(myvec)
    uniqv[which.max(tabulate(match(myvec, uniqv)))]
  }
  
  # get mode of number of lines in gpr files
  mode_num_lines <- getmode(finaldf[,8])
  
  finaldf$Consistency <- sapply(1:nrow(finaldf), function(x){
    input_file <- finaldf[x,1]
    tiff_file <- finaldf[x,3]
    jpeg_file <- finaldf[x,4]
    F532_F635 <- finaldf[x,5]
    num_lines <- finaldf[x,8]
    checkEncoding <- finaldf[x,9]
    if((file_path_sans_ext(basename(input_file)) == file_path_sans_ext(basename(tiff_file))) && (file_path_sans_ext(basename(input_file)) == file_path_sans_ext(basename(jpeg_file)))) {
      if (grepl("protein", input_file, ignore.case = T)) {
        temp.list <- unlist(strsplit(basename(input_file),"_"))
        temp.idx <- which(grepl("protein",temp.list,ignore.case = T))
        if (gsub("[^0-9]", "",temp.list[temp.idx-1]) == gsub("[^0-9]", "",temp.list[temp.idx+1])) {
          consistency <- "PASS"
        } else { 
          consistency <- "FAIL (protein name and slide number not consistent)"
        }
      } else {consistency <- "PASS"}
    } else { 
      consistency <- "FAIL (input, tiff and jpeg file names not matching)"
    }
    if (num_lines != mode_num_lines) {
      consistency <- "FAIL (number of lines not matching)"
    }
    if (!(as.logical(F532_F635))) {
      consistency <- "FAIL (F532/F635 field missing or mislabelled)"
    }
    if (!(as.logical(checkEncoding))) {
      consistency <- "FAIL (file is not UTF-8 encoded)"
    }
    return(consistency)
  })
  colnames(finaldf) <- c("FilePath","GPR File","TIFF File","JPEG File","F532/F635","QuoteMarksRemoved","BlankLinesRemoved","Num of Lines","UTF-8 Encoding","Consistency")
  if (all(finaldf$Consistency == "PASS")) {myFlag <- "PASS"} else {myFlag <- "FAIL"}
  myDate <- format(Sys.Date(), "%Y%m%d")
  outFileName <- paste0(input_path, "/", myDate,"_rppaGprConsistencyCheck-",myFlag,out_suffix,".xlsx")
  
  # write.table(finaldf, outFileName, quote = F, sep = "\t", row.names = F, col.names = T)
  openxlsx::write.xlsx(finaldf, outFileName, rowNames = F, colNames = T, overwrite = T)
  
  cat("##### CONSISTENCY CHECK COMPLETE #####\n")
  cat("##### OUTPUT FILE:",outFileName,"#####\n")
}

gprConsistencyCheck()
