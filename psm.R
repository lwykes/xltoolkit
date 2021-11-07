#UPDATED 301020
# Van Westendorm PSM with NM Extension optimised for subgroups

#error logging
log.path <- "C:/xl_toolbox/error.log"

tryCatch({
  # Code to attempt
  
  library(readxl)
  library(openxlsx)
  
  raw <- read_xlsx("C:/xl_toolbox/data_in.xlsx", sheet = "data", col_names = TRUE, na="")

    require(reshape2)
    require(pricesensitivitymeter)
    
    #Work our how many sub groups
    attach(raw)
    d<-melt(raw,id=names(raw)[1:6])
    detach(raw)
    
    d$unique<-as.factor(paste0(d$variable,"-",d$value))
    
    #loop through the sub groups applying PMS
    datalist = list()
    for (i in 1: length(levels(d$unique))) {
      
      group <- print(levels(d$unique)[i])
      s <- d[d$unique==group,]
      
      toocheap      <-s[,1]
      cheap         <-s[,2]
      expensive     <-s[,3]
      tooexpensive  <-s[,4]
      pi_expensive  <-s[,5]
      pi_cheap      <-s[,6]
      
      psm<-psm_analysis(
        toocheap, cheap, expensive, tooexpensive,
        data = NA,
        validate = TRUE,
        interpolate = FALSE,
        intersection_method = "min",
        pi_cheap = pi_cheap, pi_expensive = pi_expensive,
        pi_scale = 5:1,
        pi_calibrated = c(0.7, 0.5, 0.3, 0.1, 0))
      
      plot(psm$data_nms$price,psm$data_nms$trial,type="l")
      
      out<- cbind(group,
                  psm$data_vanwestendorp,
                  psm$data_nms,
                  psm$pricerange_lower,
                  psm$pricerange_upper,
                  psm$idp,
                  psm$opp,
                  psm$price_optimal_trial,
                  psm$price_optimal_revenue,
                  psm$total_sample)
      
      datalist[[i]] <- out
    }
    
    #combined and output
    result<-do.call(rbind, datalist)
  
    
    #save output
    options("openxlsx.numFmt" = "#,#0.00")
    style <- createStyle( fontSize = 14, fontName = "Arial", textDecoration = "bold", halign = "left")
    
    wb <- createWorkbook() 
    addWorksheet(wb, sheetName = "Results", gridLines = FALSE)

    writeData(wb, sheet = 1, "Price Sensitivity Meter", startCol = 2, startRow=2)
    writeData(wb, sheet = 1, x = result, colNames = TRUE, rowNames = FALSE, startRow = 5, startCol = 2)
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
    