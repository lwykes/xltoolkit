
#error logging
log.path <- "C:/xl_toolbox/error.log"
  
  tryCatch({
    # Code to attempt
    setwd("c:/xl_toolbox")
  
#import table
	library(readxl)

	raw <- read_xlsx("data_in.xlsx",col_names = FALSE)
	
	pic_dir <- as.character(raw[1,])
	url_list<-as.list(raw[2:nrow(raw),])
	url_list <- url_list[[1]]
	
	for (i in 1:length(url_list)){

	 url <- url_list[i]
	 ext <- strsplit(basename(url), split="\\.")[[1]][2]
	 
	 pad <- nchar(length(url_list))
	 pad <- sprintf(paste0("%0",pad,"d"), i)
	 
	 out_file <- paste0(pic_dir,"\\image-",pad,".",ext)
	 
	 getpic <- tryCatch(
	   download.file(url,out_file,quite=TRUE,mode="wb"),
	   error=function(e) e
	 )
	 
	 if(inherits(getpic, "error")) next 
	 
	 Sys.sleep(sample(1:3, 1))
	 
	}
	
	pad <- nchar(length(url_list))
	im_list <- paste0("image-",sprintf(paste0("%0",pad,"d"), seq(1:length(url_list))))
	im_list <- as.data.frame(cbind(url_list,im_list))
	
	library(openxlsx)
	wb <- createWorkbook()
	  addWorksheet(wb, sheetName = "names")
  	  writeDataTable(wb, sheet = 1, im_list, colNames = TRUE, rowNames = FALSE)
	saveWorkbook(wb, "C:/xl_toolbox/data_out.xlsx", overwrite = TRUE)

    #export
    write.table(NULL,"C:/xl_toolbox/data_out.csv", sep=",",col.names=FALSE)
      
#hand back to excel
write(paste0("Finished ",Sys.time()),file="c:/xl_toolbox/finished.log")
    
    
  }, error = function(err.msg){
    # Add error message to the error log file
    write(toString(err.msg), log.path, append=TRUE)
  }
  )

