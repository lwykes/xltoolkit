#UPDATED 301020
# Relative Importance - Drivers Analysis (Shapley Regression)

#error logging
log.path <- "C:/xl_toolbox/error.log"

tryCatch({
  # Code to attempt
  
  library(readxl)
  library(openxlsx)
  
  raw <- read_xlsx("C:/xl_toolbox/data_in.xlsx", sheet = "data", col_names = TRUE, na="")
      
      s<-data.frame(scale(raw))
      
      ols <- lm(
        as.formula(paste(colnames(s)[1], "~",
                         paste(colnames(s)[c(2:ncol(s))], collapse = "+"),
                         sep = ""
        )),
        data=s
      )
      
      library(relaimpo)
      r <- calc.relimp(ols, type = c("lmg"), rela = TRUE) 
      s <- as.data.frame(capture.output(summary(ols),r))
      
      
      
      scores <- cbind.data.frame(names(r$lmg),r$lmg)
      rownames(scores)<-NULL
      names(scores) <-c("var","relimp")
      result <- scores[order(-scores$relimp),]    
      r2 <- paste0("Variance explained ",round(r$R2,3)*100,"%")
      
      #save output
      options("openxlsx.numFmt" = "0.0%")
      style <- createStyle( fontSize = 14, fontName = "Arial", textDecoration = "bold", halign = "left")
      
      wb <- createWorkbook() 
      addWorksheet(wb, sheetName = "Results", gridLines = FALSE)
      addWorksheet(wb, sheetName = "Stats", gridLines = FALSE)
      
      writeData(wb, sheet = 1, "Relative Importance (Shapley)", startCol = 2, startRow=2)
      writeData(wb, sheet = 1, r2, startCol = 2, startRow=3)
      writeData(wb, sheet = 1, x = result, colNames = TRUE, rowNames = FALSE, startRow = 5, startCol = 2)
      addStyle(wb, sheet = 1, style, rows = 2, cols = 2)
      
      writeData(wb, sheet = 2, "Regression - Stats OLS & Shapley", startCol = 2, startRow=2)
      writeData(wb, sheet = 2, x = s, colNames = TRUE, rowNames = FALSE, startRow = 5, startCol = 2)
      addStyle(wb, sheet = 2, style, rows = 2, cols = 2)
      
      saveWorkbook(wb, "C:/xl_toolbox/data_out.xlsx", overwrite = TRUE)
      
      table_range <- paste0("B5:C",nrow(result)+5)
      
      #hand back to excel
      write(table_range,file="c:/xl_toolbox/finished.log")
      
      
}, error = function(err.msg){
  # Add error message to the error log file
  write(toString(err.msg), log.path, append=TRUE)
}
)
