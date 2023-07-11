source("~/Desktop/rppaTools/rppaTool.R")
rppaTool(inputFile = "/Users/tanmaygandhi/Box/Tanmay_Projects/20230118_RPPA0040_RPPA0044/request/BM_RPPA0040_final_report1.xlsx",
         outputFile = "/Users/tanmaygandhi/Box/Tanmay_Projects/20230118_RPPA0040_RPPA0044/request/report_final-NULL.xlsx",
         cv_cutoff = 0.25,
         sampleIDRow = 2,
         replacement = NULL)

rppaTool(inputFile = "/Users/tanmaygandhi/Box/Tanmay_Projects/20230118_RPPA0040_RPPA0044/request/BM_RPPA0040_final_report1.xlsx",
         outputFile = "/Users/tanmaygandhi/Box/Tanmay_Projects/20230118_RPPA0040_RPPA0044/request/report_final-NA.xlsx",
         cv_cutoff = 0.25,
         sampleIDRow = 1,
         replacement = NA)
