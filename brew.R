#UPDATED 301020
    
#error logging
log.path <- "C:/xl_toolbox/error.log"

tryCatch({
  # Code to attempt
  
  library(readxl)
  library(openxlsx)
  
  table <- read_xlsx("C:/xl_toolbox/data_in.xlsx", sheet = "data", col_names = FALSE, na="")
    
    f <- strsplit(as.character(table[,1]),"-")
    f <- lapply(f,function(d) {as.numeric(d)}/255)
    
    t <- strsplit(as.character(table[,2]),"-")
    t <- lapply(t,function(d) {as.numeric(d)}/255)
    
    n <- as.numeric(table[,3])
    
    from <- rgb(f[[1]][1],f[[1]][2],f[[1]][3])
    to   <- rgb(t[[1]][1],t[[1]][2],t[[1]][3])
    
    colfunc <- colorRampPalette(c(from,to))
    
    result <- colfunc(n)
    result <- sapply(result, function(x) {paste0("'",paste(as.vector(col2rgb(x)), collapse = ","))})
    result <- as.data.frame(result)
    names(result)[1]<-" "
    
    #export
    #save output
    options("openxlsx.numFmt" = "#,#0.00")
    style <- createStyle( fontSize = 14, fontName = "Arial", textDecoration = "bold", halign = "left")
    
    wb <- createWorkbook() 
    addWorksheet(wb, sheetName = "Results", gridLines = FALSE)
    
    writeData(wb, sheet = 1, "Color Ramp", startCol = 2, startRow=2)
    writeData(wb, sheet = 1, x = result, colNames = TRUE, rowNames = TRUE, startRow = 5, startCol = 2)
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
    table_range <- paste0("B5:",MORELETTERS(ncol(result)+1),nrow(result)+5)
    
    #hand back to excel
    write(table_range,file="c:/xl_toolbox/finished.log")
    
}, error = function(err.msg){
  # Add error message to the error log file
  write(toString(err.msg), log.path, append=TRUE)
}
)
    