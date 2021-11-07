
#error logging
log.path <- "C:/xl_toolbox/error.log"

tryCatch({
  # Code to attempt

  #Get API key from INI file, VBA already validated code exists
  blank = "^\\s*$"
  header = "^\\[(.*)\\]$"
  key_value = "^.*=.*$"
  
  extract = function(regexp, x) regmatches(x, regexec(regexp, x))[[1]][2]
  
  read_ini = function(fn) {
    lines = readLines(fn)
    ini = list()
    for (l in lines) {
      if (grepl(blank, l)) next
      if (grepl(header, l)) {
        section = extract(header, l)
        ini[[section]] = list()
      }
      if (grepl(key_value, l)) {
        kv = strsplit(l, "\\s*=\\s*")[[1]]
        ini[[section]][[kv[1]]] = kv[2]
      }
    }
    return(ini)
  }
  
  ini<-read_ini("c:/xl_toolbox/settings.ini")
  apikey <- ini$deeplAPIkey[[1]]
  
  #Function making API requests
  translate <- function(text) {
    body <- paste0(
      "auth_key=", apikey,
      "&text=", utils::URLencode(text),
      "&target_lang=", "EN")
    
    response <- httr::POST(
      "https://api.deepl.com/v2/translate",
      body = body,
      httr::add_headers(
        "Content-Type" = "application/x-www-form-urlencoded",
        "Content-Length" = nchar(body),
        "User-Agent" = "xl_toolbox-Trans" 
      ))
      
      if (response$status_code == 200){
          return(httr::content(response))
        } else {
  
          error_list<- data.frame(error_flag=c(rep("api_error",9)),error_code= c(400,403,404,413,414,429,456,503,529),
          error_type = c("Bad request. Please check error message and your parameters.",
                          "Authorization failed. Please supply a valid auth_key parameter.",
                          "The requested resource could not be found.",
                          "The request size exceeds the limit.",
                          "The request URL is too long. You can avoid this error by using a POST request instead of a GET request, and sending the parameters in the HTTP body.",
                          "Too many requests. Please wait and resend your request.",
                          "Quota exceeded. The character limit has been reached.",
                          "Resource currently unavailable. Try again later.",
                          "Too many requests. Please wait and resend your request."))
          return(error_list[error_list$error_code==response$status_code,])
        }
  }  
  
  #GET DATA & RESHAPE table wide to long
      library(readxl)
      library(openxlsx)
      library(tidyverse)    
      
      raw <- read_xlsx("C:/xl_toolbox/data_in.xlsx", sheet = "data", col_names = FALSE, na="")
      raw$ID <-rownames(raw)
      long <- raw %>% 
      pivot_longer(-ID, names_to = "column", values_to = "entry")
      chunk<-long$entry
      
      #PREPARE chunks of concatenated text, break list into chunks of 50 for API
      library(stringr)
      itt<-ceiling((length(chunk))/50)
      
      chunk50 <- list()
      for (i in 1:itt) {
        from = ((i-1)*50)+1
        if (i == itt) {to=length(chunk)}  else {to = i * 50}
        text<-paste(c(chunk[from : to]),collapse='&text=')
        chunk50[[i]] <- text
      }
      
      #TRANSLATE loop calling translate for each chunk
      tlated <- list()
      
      for (i in 1:length(chunk50)){
        #translate chunks
            t<-translate(chunk50[[i]])
            if (t[1]=="api_error"){error_flag = TRUE; write(toString(paste(t[1],t[2],t[3]),sep=","), log.path, append=TRUE);break}
            else {
            error_flag =FALSE
            trans_text  <-  matrix(unlist(t), nrow=length(unlist(t[1]))/2,byrow=TRUE)
            tlated[[i]] <- trans_text
            }
      }
      
      if (error_flag == FALSE) {
          
        #REASSEMBLE in original structure
        entry <- do.call(rbind, tlated)[,2]
        long$entry <- entry
        result <- pivot_wider(long, names_from = column, values_from = entry)
        result[result=="NA"] <- ""
        
        #SAVE output
        options("openxlsx.numFmt" = "#,#0.00")
        style <- createStyle( fontSize = 14, fontName = "Arial", textDecoration = "bold", halign = "left")
              
        wb <- createWorkbook() 
          addWorksheet(wb, sheetName = "Results")
          writeData(wb, sheet = 1, "Translation", startCol = 2, startRow=2)
          addStyle(wb, sheet = 1, style, rows = 2, cols = 2)
          writeData(wb, sheet = 1, x = result, colNames = FALSE, rowNames = FALSE, startRow = 5, startCol = 2)
        saveWorkbook(wb, "C:/xl_toolbox/data_out.xlsx", overwrite = TRUE)
              
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
      }
        
  }, error = function(err.msg){
    # Add error message to the error log file
    write(toString(err.msg), log.path, append=TRUE)
  }
)
