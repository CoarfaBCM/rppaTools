mainFunction <- function(inputFile, rawConfigFile, outputFile) {
  # Loading required packages and installing ones not present
  list.of.packages <- c("openxlsx")
  new.packages <- list.of.packages[!(list.of.packages %in% installed.packages()[,"Package"])]
  if(length(new.packages)>0) {install.packages(new.packages)} else {lapply(list.of.packages, require, character.only = TRUE)}
  
  mydf.pilist <- read.xlsx(inputFile, sheet = 2)
  all.pis <- unname(unlist(head(na.omit(unique(mydf.pilist[4])), n = -1))) # extracting pi initials
  all.pis <- all.pis[order(all.pis, decreasing = F)] # sorting pi initials in ascending order
  
  raw.config <- readLines(rawConfigFile)
  
  prep.config <- raw.config
  prep.config[3] <- paste0("pi_list\t",paste(all.pis, collapse = "\t"),"\tCal\tCtrl")
  
  mydf.samplelist <- read.xlsx(inputFile, sheet = 1, startRow = 2, check.names = F)
  idx.to.include <- c(1,2,5)
  
  # if cal_2 samples are not present in sample list, create config file accordingly
  if (!(any(grepl("^Cal_2",mydf.samplelist$SampleName, ignore.case = T)))) {
    for (idx in 1:length(prep.config)) {
      if (idx %in% idx.to.include) {
        temp1 <- strsplit(prep.config[idx], "\t")[[1]]
        temp2 <- temp1[grep("cal_2", temp1, ignore.case = T, invert = T)]
        prep.config[idx] <- paste(temp2, collapse = "\t")
      }
    }
    prep.config[7] <- 'Cal_HP\tCal_1_HP'
    prep.config[8] <- 'Cal_JC\tCal_1_JC'
    prep.config[9] <- 'pi_cellmix\t"Cellmix2,Ctrl_Cellmix2_"'
    prep.config[10] <- 'pi_Cal_HP\tCal_1_HP'
    prep.config[11] <- 'pi_Cal_JC\tCal_1_JC'
  }
  
  # if cellmix2 samples not present in sample list, remove cellmix2 and other qc2 samples from the configuration file
  if (!(any(grepl("cellmix2",mydf.samplelist$SampleName, ignore.case = T)))) {
    for (idx in 1:length(prep.config)) {
      if (idx %in% idx.to.include) {
        temp1 <- strsplit(prep.config[idx], "\t")[[1]]
        temp2 <- temp1[grep("mix2|cal_2|buffer2", temp1, ignore.case = T, invert = T)]
        prep.config[idx] <- paste(temp2, collapse = "\t")
      }
    }
    prep.config[7] <- 'Cal_HP\tCal_1_HP'
    prep.config[8] <- 'Cal_JC\tCal_1_JC'
    prep.config[9] <- 'pi_cellmix\t"Cellmix1,Ctrl_Cellmix1_"'
    prep.config[10] <- 'pi_Cal_HP\tCal_1_HP'
    prep.config[11] <- 'pi_Cal_JC\tCal_1_JC'
  }
  
  keep <- mydf.samplelist$Sample.Type == "M"
  if (any(keep)) { # check if there is at least 1 mouse sample present
    mydf.samplelist <- mydf.samplelist[keep, , drop=F]
    mouse.pis <- unique(unname(sapply(mydf.samplelist[,2],function(x){strsplit(x,"_")[[1]][1]}))) # extract initials for pis with mouse samples
    if (!(all(grepl("ctrl",mouse.pis,ignore.case = T)))) { # check if it's not only control sample(s)
      mouse.pis <- mouse.pis[!(grepl("ctrl",mouse.pis,ignore.case = T))] # drop "Ctrl" from pi list
      mouse.pis <- mouse.pis[order(mouse.pis, decreasing = F)] # sorting pi initials in ascending order
      prep.config[4] <- paste0("mouse_pi_list\t",paste(mouse.pis, collapse = "\t"))
    } else {
      prep.config[4] <- paste0("mouse_pi_list")
    }
  } else {
    prep.config[4] <- paste0("mouse_pi_list")
  }
  
  writeLines(prep.config, outputFile)
}


createConfigFile <- function(inputPath = "."){
  inputFile <- list.files(path = inputPath,
                          pattern = "sample list.xlsx$|sample list histone marks.xlsx$",
                          ignore.case = T, recursive = F, full.names = T)
  inputFile <- inputFile[!grepl("[~$]",inputFile)]
  
  rawConfigFile <- list.files(path = inputPath,
                              pattern = "rawConfigFile.txt",
                              ignore.case = T, recursive = F, full.names = T)
  
  alldirs <- basename(list.dirs(path = inputPath, recursive = F))
  gpr.dirs <- alldirs[grepl("_gpr$",alldirs)]
  mylabel <- strsplit(gpr.dirs[1], "_")[[1]][1]
  myDate <- format(Sys.Date(), "%m%d%Y")
  outputFile <- file.path(inputPath,paste0(mylabel, "_configuration_file.", myDate,".txt"))
  mainFunction(inputFile = inputFile,
               rawConfigFile = rawConfigFile,
               outputFile = outputFile)
  
  cat("##### OUTPUT FILE:",outputFile,"#####\n")
}

createConfigFile()
