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
  
  num_lines <- length(readLines(input_file))
  
  full_info <- c(input_file, basename(input_file), tiff_file, jpeg_file, num_lines)
  
  return(full_info)
}

# Main function
gprConsistencyCheck <- function(input_path = ".",
                                out_suffix = NULL) {
  # Loading required packages and installing ones not present
  list.of.packages <- c("tools")
  new.packages <- list.of.packages[!(list.of.packages %in% installed.packages()[,"Package"])]
  if(length(new.packages)>0) {install.packages(new.packages)} else {lapply(list.of.packages, require, character.only = TRUE)}
  
  finaldf <- NULL
  for (myfile in list.files(input_path,pattern = ".gpr",full.names = T,recursive = T)) {
    finaldf <- rbind(finaldf, readOneGpr(myfile))
  }
  finaldf <- as.data.frame(finaldf)
  finaldf$V5 <- as.numeric(finaldf$V5)
  
  # function to get statistical mode of the input
  getmode <- function(myvec) {
    uniqv <- unique(myvec)
    uniqv[which.max(tabulate(match(myvec, uniqv)))]
  }
  
  # get mode of number of lines in gpr files
  mode_num_lines <- getmode(finaldf[,5])
  
  finaldf$Consistent <- sapply(1:nrow(finaldf), function(x){
    input_file <- finaldf[x,1]
    tiff_file <- finaldf[x,3]
    jpeg_file <- finaldf[x,4]
    num_lines <- finaldf[x,5]
    if((file_path_sans_ext(basename(input_file)) == file_path_sans_ext(basename(tiff_file))) && (file_path_sans_ext(basename(input_file)) == file_path_sans_ext(basename(jpeg_file)))) {
      if (grepl("protein", input_file, ignore.case = T)) {
        temp.list <- unlist(strsplit(basename(input_file),"_"))
        temp.idx <- which(grepl("protein",temp.list,ignore.case = T))
        if (gsub("[^0-9]", "",temp.list[temp.idx-1]) == gsub("[^0-9]", "",temp.list[temp.idx+1])) {
          consistency <- "TRUE"
        } else { 
          consistency <- "FALSE (protein name and slide number not consistent)"
        }
      } else {consistency <- "TRUE"}
    } else { 
      consistency <- "FALSE (input, tiff and jpeg file names not matching)"
    }
    if (num_lines != mode_num_lines) {
      consistency <- "FALSE (number of lines not matching)"
    }
    return(consistency)
  })
  colnames(finaldf) <- c("FilePath","GPR File","TIFF File","JPEG File","Num of lines","Consistent")
  if (all(finaldf$Consistent == "TRUE")) {myFlag <- "PASS"} else {myFlag <- "FAIL"}
  myDate <- format(Sys.Date(), "%Y%m%d")
  outFileName <- paste0(input_path, "/", myDate,"_rppaGprConsistencyCheck-",myFlag,out_suffix,".xls")
  
  write.table(finaldf, outFileName, quote = F, sep = "\t", row.names = F, col.names = T)
}

gprConsistencyCheck()
