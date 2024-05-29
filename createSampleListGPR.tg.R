mainFunction <- function(sampleListFile, gprFile, outputFile) {
  # reading in gpr file
  # we want the data from row 33 and below and columns 1 through 5
  # since these gpr files have columns with no data, read.table is unable to read this data in
  # hence using readlines instead to read in the lines of data and combine them to form our needed dataframe
  rawgpr <- readLines(gprFile)
  split_lines <- lapply(rawgpr[33:length(rawgpr)], function(x) strsplit(x, "\t")[[1]][1:5])
  rawgprdf <- data.frame(do.call(rbind, split_lines), stringsAsFactors = F)
  colnames(rawgprdf) <- rawgprdf[1,]
  rawgprdf <- rawgprdf[-1,]
  rawgprdf[1:3] <- apply(rawgprdf[1:3], 2, as.numeric)
  rawgprdf$`Name for Sample group protein` <- ""
  
  # reading in sample list file
  df.samplelist <- openxlsx::read.xlsx(sampleListFile, sheet = 1, colNames = T, rowNames = F, startRow = 2)
  
  commonSamples <- intersect(rawgprdf$Name, df.samplelist$Name.in.print.map)
  
  if (length(commonSamples) == 0) {stop(cat("#### NO SAMPLE NAMES MATCHING BETWEEN SAMPLE LIST FILE AND GPR FILE ####\n#### CHECK SAMPLE LIST FILE:", sampleListFile,"####\n#### CHECK GPR FILE:", gprFile,"####\n"))}
  
  numSamples.sampleList <- sum(!(grepl("^Cal|^Ctrl", df.samplelist$Name.in.print.map)))
  numSamples.match <- sum(df.samplelist$Name.in.print.map %in% rawgprdf$Name)
  
  if (numSamples.sampleList > numSamples.match) {stop((cat("#### NUMBER OF SAMPLES FROM SAMPLE LIST FILE MISSING IN THE GPR FILE:", numSamples.sampleList - numSamples.match, "####\n")))}
  
  rawgprdf$`protein concentration for normalization` <- sapply(rawgprdf$Name,
                                                               function(x) {
                                                                 if (x %in% df.samplelist$Name.in.print.map) {
                                                                   df.samplelist$`Normalization.Conc.(ug/ul)`[df.samplelist$Name.in.print.map == x]
                                                                 } else {""}
                                                               })
  rawgprdf <- rawgprdf[,c(1:4,6:7,5)]
  
  myheader <- paste(names(rawgprdf), collapse = "\t")
  data_lines <- apply(rawgprdf, 1, paste, collapse = "\t")
  
  finaldata <- c(rawgpr[1:32], myheader, data_lines)
  
  writeLines(finaldata, outputFile)
  
  cat("##### OUTPUT FILE:",outputFile,"#####\n")
}

createSampleListGPR <- function(inputPath = ".") {
  # Loading required packages and installing ones not present
  list.of.packages <- c("openxlsx")
  new.packages <- list.of.packages[!(list.of.packages %in% installed.packages()[,"Package"])]
  if(length(new.packages)>0) {install.packages(new.packages)} else {lapply(list.of.packages, require, character.only = TRUE)}
  
  # finding gpr file
  alldirs <- basename(list.dirs(path = inputPath, recursive = F))
  gpr.dirs <- alldirs[grepl("_gpr$",alldirs)]
  mydir <- list.dirs(gpr.dirs, recursive = F, full.names = T)[2]
  gprFile <- list.files(path = mydir, pattern = ".gpr$", full.names = T)[1]
  
  # finding sample list table file
  sampleListFile <- list.files(path = inputPath,
                          pattern = "sample list.xlsx$|sample list histone marks.xlsx$",
                          ignore.case = T, recursive = F, full.names = T)
  sampleListFile <- sampleListFile[!grepl("[~$]",sampleListFile)]
  
  mylabel <- strsplit(gpr.dirs[1], "_")[[1]][1]
  myDate <- format(Sys.Date(), "%m%d%Y")
  outputFile <- file.path(inputPath,paste0(mylabel,"_sample_list_gpr.processed.", myDate,".txt"))
  
  mainFunction(sampleListFile = sampleListFile,
               gprFile = gprFile,
               outputFile = outputFile)
}

createSampleListGPR()
