#UPDATED 301020

#error logging
log.path <- "C:/xl_toolbox/error.log"
  
  tryCatch({
    # Code to attempt
    rm(list=ls())
      library(readxl)
      library(openxlsx)
      
      data <- read_xlsx("C:/xl_toolbox/data_in.xlsx", sheet = "data", col_names = FALSE, na="")
      
      names(data)[1] <- "comments"
      
      library(wordcloud)
      library(jiebaR)
      library(stringr)
      
      #remove stopwords
      stopwords <- read.csv('C:/xl_toolbox/stopwords.r',encoding="UTF-8",stringsAsFactors = FALSE,header = FALSE)
      stopwords <- str_c(unlist(stopwords), collapse="|")
      comments <- lapply(data$comments,function(x) gsub(stopwords,"",as.character(x)))
      comments<-unlist(comments)
      
      #tag parts of speech
      tagger = worker("tag",bylines=T)
      
      tags <-segment(comments, tagger)
      pos <- lapply(tags,function(x) names(x))
      
      library(reshape2)
      tags_long <- melt(tags)
      pos_long <- melt(pos)
      
      tags_long <- cbind(tags_long, pos_long$value)
      data$L1 <- 1:nrow(data)
      
      tags_long <- merge(x = tags_long, y = data, by = "L1", all.x = TRUE)
      tags_long <- tags_long[,c(1,4,2,3)]
      names(tags_long) <- c("id","full_comment","token","part_of_speech")      

      #summary table token frequency    
      token_freq <- as.data.frame(table(tags_long$token))
      token_freq <- token_freq[order(-token_freq$Freq),]
      
      #token freq by POS
      a<-subset(as.data.frame(table(tags_long[,c(3,4)])), Freq != 0)
      a<-a[order(a$part_of_speech,-a$Freq),]
    
      #noun word clouds
      df<-a[a$part_of_speech=="n",]
      library(devEMF)
      emf(file = "plot.emf", width = 3, height = 3,  bg = "transparent", pointsize = 12, family = "Helvetica", coordDPI = 300)
      set.seed(1234) # for reproducibility 
      wordcloud(words = df$token, freq = df$Freq, min.freq = 2, max.words=200, 
                random.order=FALSE, rot.per=0, colors=brewer.pal(8, "Dark2"))
      dev.off()
      
      #save output
      options("openxlsx.numFmt" = "#,#0.00")
      style <- createStyle( fontSize = 14, fontName = "Arial", textDecoration = "bold", halign = "left")
      
      wb <- createWorkbook() 
        addWorksheet(wb, sheetName = "Results")
        
        showGridLines(wb, 1, showGridLines = FALSE)
        
        writeData(wb, sheet = 1, "Parts of speech (POS)", startCol = 2, startRow=2)
        addStyle(wb, sheet = 1, style, rows = 2, cols = 2)
        writeData(wb, sheet = 1, x = tags_long, colNames = TRUE, rowNames = FALSE, startRow = 5, startCol = 2)
        
        writeData(wb, sheet = 1, "Tag frequency by POS", startRow=2,startCol = 7)
        addStyle(wb, sheet = 1, style, rows = 2, cols = 7)
        writeData(wb, sheet = 1, x = a, colNames = TRUE, rowNames = FALSE, startRow = 5, startCol = 7)
        
        writeData(wb, sheet = 1, "Overall tag frequency", startRow=2,startCol = 12)
        addStyle(wb, sheet = 1, style, rows = 2, cols = 12)
        writeData(wb, sheet = 1, x = token_freq, colNames = TRUE, rowNames = FALSE, startRow = 5, startCol = 12)
        
        writeData(wb, sheet = 1, "Wordclound of nouns", startRow=2,startCol =16)
        addStyle(wb, sheet = 1, style, rows = 2, cols = 16)
        insertImage(wb, 1, "plot.emf", width = 15, height = 12, startRow = 3, startCol = 16, units = "cm")
              
        setColWidths(wb, sheet = 1, cols = c(2,5,8,9,11,4,7,10,3), widths = c(rep(5,5),rep(15,3),20))
        
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

