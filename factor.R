#UPDATED 301020

#error logging
log.path <- "C:/xl_toolbox/error.log"

tryCatch({
  # Code to attempt
  library(readxl)
  library(openxlsx)
  library(psych)
  
  raw <- read_xlsx("C:/xl_toolbox/data_in.xlsx", sheet = "data", col_names = TRUE, na="")
  
  #convert to matrix
  table<-raw[,2:ncol(raw)]
  table<-as.matrix(table)
  
  # Determine Number of Factors to Extract
  #install.packages("nFactors")
  library(nFactors)
  ev <- eigen(cor(table)) # get eigenvalues
  ap <- parallel(subject=nrow(table),var=ncol(table),
                 rep=100,cent=.05)
  nS <- nScree(x=ev$values, aparallel=ap$eigen$qevpea)
  plotnScree(nS)
  
  #number to extract
  ne<-as.numeric(nS$Components [4])
  
  #run factor
  pc<-principal(table, nfactors = ne, residuals = FALSE,rotate="varimax", covar=FALSE,
            scores=TRUE,missing=FALSE,impute="median",oblique.scores=TRUE,method="regression")
  
  s <- as.data.frame(capture.output(pc))
  
  result<-unclass(pc$loadings)
  rownames(result)<-raw[,1][[1]]
  result<-fa.sort(result)
  result<-as.data.frame(result)
  
  #save output
  options("openxlsx.numFmt" = "#,#0.00")
  style <- createStyle( fontSize = 14, fontName = "Arial", textDecoration = "bold", halign = "left")
  
  wb <- createWorkbook() 
    addWorksheet(wb, sheetName = "Results", gridLines = FALSE)
    addWorksheet(wb, sheetName = "Stats", gridLines = FALSE)
  
    writeData(wb, sheet = 1, "Factor Analysis", startCol = 2, startRow=2)
    writeData(wb, sheet = 1, "Rotated factor loadings (varimax)", startCol = 2, startRow=3)
    writeData(wb, sheet = 1, x = result, colNames = TRUE, rowNames = TRUE, startRow = 5, startCol = 2)
    addStyle(wb, sheet = 1, style, rows = 2, cols = 2)
    
    writeData(wb, sheet = 2, "Factor Analysis - Stats", startCol = 2, startRow=2)
    writeData(wb, sheet = 2, x = s, colNames = TRUE, rowNames = FALSE, startRow = 5, startCol = 2)
    addStyle(wb, sheet = 2, style, rows = 2, cols = 2)
    
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