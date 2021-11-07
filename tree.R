#UPDATED 301020

#error logging
log.path <- "C:/xl_toolbox/error.log"

tryCatch({
  # Code to attempt
  
  library(readxl)
  library(openxlsx)
  
  raw <- read_xlsx("C:/xl_toolbox/data_in.xlsx", sheet = "data", col_names = TRUE, na="")
  table <- raw[,2:ncol(raw)]
  
  # Ward Hierarchical Clustering
  d <- dist(table, method = "euclidean") # distance matrix
  fit <- hclust(d, method="ward.D") 
      
  #save groupings for all solutions
  datalist = list()
  for (i in 1:nrow(raw)){
    dat <- cutree(fit, k=i)
    datalist[[i]] <- dat
  }
  
  big_data = do.call(rbind, datalist)
  big_data <- t(big_data)
  big_data <- as.data.frame(big_data)
  
  colnames(big_data)<-paste0("clu",1:nrow(raw))
      
  table <- cbind(raw,big_data)
  table <- table[(fit$order),]
 
  #draw tree
  png("C:/xl_toolbox/plot.png", width = 1200)
  plot(fit)
  dev.off()

  #save output
  options("openxlsx.numFmt" = "#,#0.00")
  style <- createStyle( fontSize = 14, fontName = "Arial", textDecoration = "bold", halign = "left")
  
  wb <- createWorkbook() 
    addWorksheet(wb, sheetName = "Results")
    writeData(wb, sheet = 1, "Dendogram (Ward's method)", startCol = 2, startRow=2)
    addStyle(wb, sheet = 1, style, rows = 2, cols = 2)
    writeData(wb, sheet = 1, x = table, colNames = TRUE, rowNames = FALSE, startRow = 5, startCol = 2)
    insertImage(wb, 1, "C:/xl_toolbox/plot.png", startRow = 5, startCol = 10)
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
  table_range <- paste0("B4:",MORELETTERS(ncol(table)+1),nrow(table)+5)
  
  #hand back to excel
  write(table_range,file="c:/xl_toolbox/finished.log")
      
      
}, error = function(err.msg){
  # Add error message to the error log file
  write(toString(err.msg), log.path, append=TRUE)
}
)
