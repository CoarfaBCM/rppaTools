# organize the gpr and the jpg files into two big folders ----
rppaOrganize <- function(input_path = ".") {
  alldirs <- basename(list.dirs(path = input_path, recursive = F))
  
  gpr.dirs <- alldirs[grepl("_gpr$",alldirs)]
  jpg.dirs <- alldirs[grepl("_image$",alldirs)]
  
  mylabel <- strsplit(gpr.dirs[1], "_")[[1]][1]
  
  if (!dir.exists(file.path(input_path, paste0(mylabel,"_gpr")))) {
    dir.create(file.path(input_path, paste0(mylabel,"_gpr")))
  }
  if (!dir.exists(file.path(input_path, paste0(mylabel,"_jpg")))) {
    dir.create(file.path(input_path, paste0(mylabel,"_jpg")))
  }
  
  file.rename(gpr.dirs, file.path(input_path, paste0(mylabel,"_gpr"), gpr.dirs))
  file.rename(jpg.dirs, file.path(input_path, paste0(mylabel,"_jpg"), jpg.dirs))
  
  cat("##### ORGANIZING FILES COMPLETE #####\n")
}

rppaOrganize()
