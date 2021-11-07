#UPDATED 301020
# Stack data

#error logging
log.path <- "C:/xl_toolbox/error.log"

tryCatch({
  # Code to attempt
  
  library(readxl)
  library(openxlsx)
  
  raw <- read_xlsx("C:/xl_toolbox/data_in.xlsx", sheet = "data", col_names = TRUE, na="")
  
  nfixed <- noquote(names(raw)[1])
  nfixed <- as.numeric(strsplit(nfixed, ":")[[1]][2])
  
  vnames <- names(raw)[1:length(names(raw))]
  vid <- noquote(paste0("V",1:(length(vnames))))
  
  vnames<-cbind.data.frame(vid,vnames)
  names(vnames)<-c("variable","Group")
  
  idvars <-noquote(paste0("V",1:nfixed))
  
  #Stack data and merge on column names
  library(reshape2)
  d <- raw
  names(d)<-vid
  
  d<-melt(d,id.vars=idvars)
  l<-merge(x = d , y = vnames, by = "variable", all.x = TRUE)
      
  #reorder & rename
  l<-l[,c(noquote(paste0("V",1:nfixed)), "Group","value")]
  names(l)[1:nfixed]<-names(raw)[1:nfixed]
  names(l)[1]<-"Labels"
  
  #omit rows with no values
  l<-l[!is.na(l$value),]
      
  #export
  #save output
  options("openxlsx.numFmt" = "#,#0.00")
  style <- createStyle( fontSize = 14, fontName = "Arial", textDecoration = "bold", halign = "left")
  
  wb <- createWorkbook() 
  addWorksheet(wb, sheetName = "Results", gridLines = FALSE)
  
  writeData(wb, sheet = 1, "Stacked table", startCol = 2, startRow=2)
  writeData(wb, sheet = 1, x = l, colNames = TRUE, rowNames = FALSE, startRow = 5, startCol = 2)
  addStyle(wb, sheet = 1, style, rows = 2, cols = 2)
  
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
  table_range <- paste0("B5:",MORELETTERS(ncol(l)+1),nrow(l)+5)
  
  
#hand back to excel
  write(table_range,file="c:/xl_toolbox/finished.log")
      
}, error = function(err.msg){
  # Add error message to the error log file
  write(toString(err.msg), log.path, append=TRUE)
}
)



