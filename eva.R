#UPDATED 301020

#error logging
log.path <- "C:/xl_toolbox/error.log"
  
  tryCatch({
    # Code to attempt
    
      library(readxl)
      library(openxlsx)
      
      raw <- read_xlsx("C:/xl_toolbox/data_in.xlsx", sheet = "data", col_names = TRUE, na="")
      table <- raw
      
      x <- chisq.test(table[2:ncol(table)]) 
      result<- as.data.frame(x$observed-x$expected)
      
      table[2:ncol(table)] <- result
      names(table)[1]<-""
      
      #save output
      options("openxlsx.numFmt" = "#,#0.00")
      style <- createStyle( fontSize = 14, fontName = "Arial", textDecoration = "bold", halign = "left")
      
      wb <- createWorkbook() 
        addWorksheet(wb, sheetName = "Results")
        writeData(wb, sheet = 1, "EVA Results", startCol = 2, startRow=2)
        addStyle(wb, sheet = 1, style, rows = 2, cols = 2)
        writeData(wb, sheet = 1, x = table, colNames = TRUE, rowNames = FALSE, startRow = 5, startCol = 2)
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
      table_range <- paste0("B5",MORELETTERS(ncol(table)+1),nrow(table)+5)
      
#hand back to excel
write(table_range,file="c:/xl_toolbox/finished.log")
    
    
  }, error = function(err.msg){
    # Add error message to the error log file
    write(toString(err.msg), log.path, append=TRUE)
  }
  )

