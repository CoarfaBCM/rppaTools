arrayLayout <- function(inputFile) {
  library(openxlsx)
  fullmap <- read.xlsx(inputFile, startRow = 2, rowNames = T, colNames = F, skipEmptyRows = F, skipEmptyCols = F)
  fullmap[is.na(fullmap)] <- 0
  mycols <- c(12,9,6,3)
  mygroup <- unlist(lapply(1:4,function(x){paste0(rep(x,2),c("A","B"))}))
  finaldf.layout <- c()
  finaldf.array <- c()
  
  for (k in seq_along(mygroup)) {
    layoutdf <- matrix(data = NA, nrow = 4, ncol = 10)
    arraydf <- matrix(data = NA, nrow = 40, ncol = 4, dimnames = list(NULL,c("Row","Col","ID","Dilution")))
    idx <- 1
    for (j in 1:ncol(layoutdf)) {
      for (i in 1:nrow(layoutdf)) {
        if (j==1) {myrow <- k} else {myrow <- 8*(j-1) + k}
        layoutdf[i,j] <- fullmap[myrow, mycols[i]]
        arraydf[idx,1] <- i + 4*(ceiling(k/2)-1)
        if (k %% 2 == 0) {arraydf[idx,2] <- j + 10} else {arraydf[idx,2] <- j}
        arraydf[idx,3] <- fullmap[myrow, mycols[i]]
        arraydf[idx,4] <- 0
        idx <- idx + 1
      }
    }
    if (k %% 2 == 0) {newrow <- 11:20} else {newrow <- 1:10}
    tempdf <- cbind(rep(mygroup[k],nrow(layoutdf)+1),rbind(newrow,layoutdf))
    finaldf.layout <- rbind(finaldf.layout, tempdf)
    finaldf.array <- rbind(finaldf.array, arraydf)
  }
  
  write.table(finaldf.layout,gsub(".xlsx$","-layout.csv",inputFile), sep = ",", row.names = F, col.names = F)
  write.table(finaldf.array,gsub(".xlsx$","-array.csv",inputFile), sep = ",", row.names = F, col.names = T)
}

inputFile <- Sys.glob("./*-map.xlsx")
arrayLayout(inputFile)
