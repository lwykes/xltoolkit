#UPDATED 301020

#error logging
log.path <- "C:/xl_toolbox/error.log"

tryCatch({
  # Code to attempt
  library(readxl)
  library(openxlsx)
  
  table <- read_xlsx("C:/xl_toolbox/data_in.xlsx", sheet = "data", col_names = TRUE, na="")
  
      #Spearman - more appropriate for ordinal scales, no assumption normality of distribution
      c <- cor(table, use="complete.obs", method="spearman")
      
      # reorder matrix
      reorder_cormat <- function(cormat){
        # Use correlation between variables as distance
        dd <- as.dist((1-cormat)/2)
        hc <- hclust(dd)
        cormat <-cormat[hc$order, hc$order]
      }
      
      # Get lower triangle of the correlation matrix
      get_lower_tri<-function(cormat){
        cormat[upper.tri(cormat)] <- NA
        return(cormat)
      }
      # Get upper triangle of the correlation matrix
      get_upper_tri <- function(cormat){
        cormat[lower.tri(cormat)]<- NA
        return(cormat)
      }

      c<- reorder_cormat(c)
      #c<- get_lower_tri(c)
      
      table<-c
      #save output
      options("openxlsx.numFmt" = "#,#0.00")
      style <- createStyle( fontSize = 14, fontName = "Arial", textDecoration = "bold", halign = "left")
      
      wb <- createWorkbook() 
      addWorksheet(wb, sheetName = "Results")
      writeData(wb, sheet = 1, "Spearman Correlation", startCol = 2, startRow=2)
      addStyle(wb, sheet = 1, style, rows = 2, cols = 2)
      writeData(wb, sheet = 1, x = table, colNames = TRUE, rowNames = TRUE, startRow = 5, startCol = 2)
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
      table_range
      
      #hand back to excel
      write(table_range,file="c:/xl_toolbox/finished.log")
      
      
}, error = function(err.msg){
  # Add error message to the error log file
  write(toString(err.msg), log.path, append=TRUE)
}
)