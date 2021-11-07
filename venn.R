#UPDATED 301020

#error logging
log.path <- "C:/xl_toolbox/error.log"

tryCatch({
  # Code to attempt
  library(readxl)
  library(openxlsx)
  
  table <- read_xlsx("C:/xl_toolbox/data_in.xlsx", sheet = "data", col_names = TRUE, na="")
  
  #create venn if < 7
  library(venn)
  
  if (ncol(table) < 8 ) {
  options(warn=-1)
  win.metafile("c:/xl_toolbox/plot.wmf")
    venn(table,zcolor = "style")
  dev.off()
  options(warn=0)
  }
  
  #supplementary table
  result<- as.data.frame(table(table))
  result[] <- lapply(result, as.numeric)
  result[,1:(ncol(result)-1)]<- result[,1:(ncol(result)-1)]-1

  #save output
  options("openxlsx.numFmt" = "#,#0.00")
  style <- createStyle( fontSize = 14, fontName = "Arial", textDecoration = "bold", halign = "left")
  
  wb <- createWorkbook() 
    addWorksheet(wb, sheetName = "Results", gridLines = FALSE)
    writeData(wb, sheet = 1, "Pattern Analysis", startCol = 2, startRow=2)
    if (ncol(table) < 8 ) {
    insertImage(wb, 1, "c:/xl_toolbox/plot.wmf", startRow = 2, startCol = 8,height = 25,width=30,unit="cm")
      }
    addStyle(wb, sheet = 1, style, rows = 2, cols = 2)
    writeData(wb, sheet = 1, x = result, colNames = TRUE, rowNames = FALSE, startRow = 5, startCol = 2)
  saveWorkbook(wb, "C:/xl_toolbox/data_out.xlsx", overwrite = TRUE)
  
  #pass table range back to XL
  extend <- function(alphabet) function(i) {
    #stackoverflow.com/questions/25876755
    base10toA <- function(n, A) {
      stopifnot(n >= 0L)
      N <- length(A)
      j <- n %/% N 
      if (j == 0L) A[n + 1L] else paste0(Recall(j - 1L, A), A[n %% N + 1L])
    }   
    vapply(i-1L, base10toA, character(1L), alphabet)
  }
  MORELETTERS <- extend(LETTERS)
  table_range <- paste0("B5:",MORELETTERS(ncol(table)+1),nrow(table)+5)
  
  
  #hand back to excel
  write(table_range,file="c:/xl_toolbox/finished.log")
  
  
}, error = function(err.msg){
  # Add error message to the error log file
  write(toString(err.msg), log.path, append=TRUE)
}
)


