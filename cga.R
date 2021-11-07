#UPDATED 301020

#error logging
log.path <- "C:/xl_toolbox/error.log"
  
  tryCatch({
    # Code to attempt
    
    library(readxl)
    library(openxlsx)
    
    raw <- read_xlsx("C:/xl_toolbox/data_in.xlsx", sheet = "data", col_names = TRUE, na="")
    
    #convert to decimal percentage
    if(max(raw[-1])>1) {
      dt <- t(raw[-1]) / 100  
    } else {
      dt <- t(raw[-1])
    }
    
    # transform into log, then scale, then calc distances
    library(gtools)
    dtl <- logit(dt)
    dtls <- scale(dtl)
    m <- dist(dtls, method = "euclidean", diag = FALSE, upper = FALSE)
    m <- as.matrix(m)
    
    # Determine number of clusters
    wss <- (nrow(m)-1)*sum(apply(m,2,var))
    for (i in 2:ncol(m)-1) wss[i] <- sum(kmeans(m,centers=i)$withinss)
    
    # Ward Hierarchical Clustering
    d <- dist(m, method = "euclidean") # distance matrix
    fit <- hclust(d, method="ward.D") 
    
    library(png)
    png("c:/xl_toolbox/plot.png",width = 1300)
      par(mfrow=c(1,2))
      plot(1:length(wss), wss, type="b", xlab="Number of Clusters", ylab="Within groups sum of squares",main="Solution evaluation")
      plot(fit)
    dev.off()
    
    
    # Get a list of all solutions
    config = list()
    for (i in 1:length(fit$labels)) {
      dat <- cutree(fit, k=i)
      dat <- as.data.frame(dat)
      names(dat)<-paste0("clu",i)
      config[[i]] <- dat
    }
    
    config <- do.call(cbind, config)
    variable <-row.names(config)
    config <- cbind.data.frame(variable,config)
    rownames(config)<-NULL
    
    # Col Name Helpers
    extend <- function(alphabet) function(i) {
      base10toA <- function(n, A) {
        stopifnot(n >= 0L)
        N <- length(A)
        j <- n %/% N 
        if (j == 0L) A[n + 1L] else paste0(Recall(j - 1L, A), A[n %% N + 1L])
      }   
      vapply(i-1L, base10toA, character(1L), alphabet)
    }
    MORELETTERS <- extend(LETTERS)       
    
    # Calculate average profiles for each solution
    library(reshape2)
    scores<-melt(raw)
    sol<-merge(x = scores, y = config, by = "variable", all.x = TRUE)
    
    nclu <- ncol(sol)-3
    
    datalist = list()
    
    for (i in 1:nclu){
      tmp <- sol[,c(1:3,3+i)]
      fx <- formula(paste0(names(tmp)[2], "~", names(tmp)[4]))
      out <-dcast(tmp, fx, fun.aggregate = mean)
      nm<-paste0(MORELETTERS(i), names(out))
      names(out)<-nm
      names(out)[1]<-" "
      
      datalist[[i]] <- out
    }
    
    profiles <- do.call(cbind, datalist)
  
    #save output
    options("openxlsx.numFmt" = "#,#0.00")
    style <- createStyle( fontSize = 14, fontName = "Arial", textDecoration = "bold", halign = "left")
    
    wb <- createWorkbook() 
      addWorksheet(wb, sheetName = "Clusters", gridLines = FALSE)
      addWorksheet(wb, sheetName = "Profiles", gridLines = FALSE)
      addWorksheet(wb, sheetName = "Dendogram", gridLines = FALSE)
      addWorksheet(wb, sheetName = "Input Data", gridLines = FALSE)
      
      writeData(wb, sheet = 4, x = raw, colNames = TRUE, rowNames = FALSE, startRow = 5, startCol = 2)
      writeData(wb, sheet = 3, x = config, colNames = TRUE, rowNames = FALSE, startRow = 5, startCol = 2)
      writeData(wb, sheet = 2, x = profiles, colNames = TRUE, rowNames = FALSE, startRow = 5, startCol = 2)
      
      insertImage(wb, 1, "c:/xl_toolbox/plot.png", startRow = 3, startCol = 2, width=25, height=12, units="cm")
      
      writeData(wb, sheet = 1, "Stretched & Standardized Dendrogram", startCol = 2, startRow=2)
      addStyle(wb, sheet = 1, style, rows = 2, cols = 2)
  
      writeData(wb, sheet = 2, "Segment average profiles", startCol = 2, startRow=2)
      addStyle(wb, sheet = 2, style, rows = 2, cols = 2)
  
      writeData(wb, sheet = 3, "Segment tree", startCol = 2, startRow=2)
      addStyle(wb, sheet = 3, style, rows = 2, cols = 2)
  
      writeData(wb, sheet = 4, "Raw data", startCol = 2, startRow=2)
      addStyle(wb, sheet = 4, style, rows = 2, cols = 2)
      
    saveWorkbook(wb, "C:/xl_toolbox/data_out.xlsx", overwrite = TRUE) ## save to working directory
    
    table_range <- paste0("B5:",MORELETTERS(ncol(table)+1),nrow(table)+5)
    
#hand back to excel
write(table_range,file="c:/xl_toolbox/finished.log")
    
    
  }, error = function(err.msg){
    # Add error message to the error log file
    write(toString(err.msg), log.path, append=TRUE)
  }
  )





  
