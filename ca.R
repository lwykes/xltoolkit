#UPDATED 301020

#error logging
log.path <- "C:/xl_toolbox/error.log"

tryCatch({
  # Code to attempt
  
  library(readxl)
  library(openxlsx)
  
  raw <- read_xlsx("C:/xl_toolbox/data_in.xlsx", sheet = "data", col_names = TRUE, na="")
  table <- as.data.frame(raw[,2:ncol(raw)])
  
  #check for supplementary cols, hidden functionality
  passive<-noquote(names(raw)[1])
  passive<-strsplit(passive, ":")
    
    if (is.na(passive)) {
      include_sup = FALSE
    } else {
      if (passive[[1]][1]=="suprow") {
        
        passive=passive[[1]][2]:passive[[1]][3]
        include_sup = "row"
        
        } else {
          if (passive[[1]][1]=="supcol") {
            
            passive=passive[[1]][2]:passive[[1]][3]
            include_sup = "col"
            
            } else {
              if (passive[[1]][1]=="supboth") {
                
                passive_row=passive[[1]][4]:passive[[1]][5]
                passive_col=passive[[1]][2]:passive[[1]][3]
                
                include_sup = "both"
              } else {             
                  include_sup = FALSE
            }}}}
    
    #Run correspondence (turn warnings of to prevent creation of log)
    options(warn=-1)
    library(ca)
    options(warn=0)
    
    if (include_sup=="row")  {
        fit <- ca(table,suprow=passive)
        } else {
          if (include_sup=="col")  {
            fit <- ca(table,supcol=passive)
            } else {
              if (include_sup=="both")  {
                fit <- ca(table,supcol=passive_col, suprow=passive_row)
                } else {
                fit <- ca(table)
                }}}
    
    
    s<-summary(fit, scree = TRUE, rows=FALSE, columns=FALSE)
    passive_note <- paste0("Passive cols: ",length(fit$colsup),", Passive rows: ",length(fit$rowsup))
    
  #put names as positional
    fit$rownames<-paste0("R",1:nrow(table))
    fit$colnames<-paste0("C",2:ncol(table))
    
    #Format scores for easy export to XL
    cols<-cbind("Columns",as.data.frame(fit$colcoord))
    names(cols)[1]<-"role"
    
    rows<-cbind("Rows",as.data.frame(fit$rowcoord))
    names(rows)[1]<-"role"
    scores<-rbind(cols,rows)
    
    labels<-c(names(raw)[2:length(names(raw))],raw[,1][[1]])
    
    scores<-cbind(scores[,1],labels,scores[,2:ncol(scores)])
    rownames(scores)<-NULL
    names(scores)[1]<-"role"
    names(scores)[2]<-"item"
    
    #Append var explained to dimension names
    var<-round(s$scree[,3],1)
    dims<-names(scores)[3:ncol(scores)]
    new<-paste(dims," (",var,"%)",sep="")
    names(scores)[3:ncol(scores)]<-new
    
    table <- scores
    
    #save output
    options("openxlsx.numFmt" = "#,#0.00")
    style <- createStyle( fontSize = 14, fontName = "Arial", textDecoration = "bold", halign = "left")
    
    wb <- createWorkbook() 
    addWorksheet(wb, sheetName = "Results")
    writeData(wb, sheet = 1, "Correspondence Analysis", startCol = 2, startRow=2)
    writeData(wb, sheet = 1, passive_note, startCol = 2, startRow=3)
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
    table_range <- paste0("B5:",MORELETTERS(ncol(table)+1),nrow(table)+5)

    #hand back to excel
    write(table_range,file="c:/xl_toolbox/finished.log")
    
}, error = function(err.msg){
  # Add error message to the error log file
  write(toString(err.msg), log.path, append=TRUE)
}
)


