
#error logging
log.path <- "C:/xl_toolbox/error.log"
  
  tryCatch({
    # Code to attempt

    # libraries
    library(packcircles)
    library(ggplot2)
    library(viridis)
    library(stringr)
    library(readxl)
    library(openxlsx)
    
    raw <- read_xlsx("C:/xl_toolbox/data_in.xlsx", sheet = "data", col_names = TRUE, na="")
    names(raw) <-c("item","pct","fill","border")
    
    # Generate the layout
    packing <- circleProgressiveLayout(raw$pct, sizetype='area')
    packing$radius=0.95*packing$radius
    data <- cbind(raw, packing)
    dat.gg <- circleLayoutVertices(packing, npoints=50)
    
    #set colours
    ntimes <- nrow(dat.gg)/nrow(data)
    
    
    #fill colour
    r<-as.numeric(sapply(strsplit(as.character(raw$fill), "\\."), `[`, 1))
    g<-as.numeric(sapply(strsplit(as.character(raw$fill), "\\."), `[`, 2))
    b<-as.numeric(sapply(strsplit(as.character(raw$fill), "\\."), `[`, 3))
    
    fc<-cbind(r,g,b)
    fc<-fc/255
    
    fc<-rgb(fc[,1],fc[,2],fc[,3])
    fc<-rep(fc, each = ntimes)
    
    #border colour
    r<-as.numeric(sapply(strsplit(as.character(raw$border), "\\."), `[`, 1))
    g<-as.numeric(sapply(strsplit(as.character(raw$border), "\\."), `[`, 2))
    b<-as.numeric(sapply(strsplit(as.character(raw$border), "\\."), `[`, 3))
    
    bc<-cbind(r,g,b)
    bc<-bc/255
    
    bc<-rgb(bc[,1],bc[,2],bc[,3])
    bc<-rep(bc, each = ntimes)
    
    txt<- rep(data$pct, each=ntimes)
    
    #create plot
      ggplot() + 
      geom_polygon(data = dat.gg, aes(x, y, group = id, fill=id), colour=bc , alpha = 2, size=2, fill=fc) +
      
      geom_text(data = data, aes(x, y, size=pct, label = str_wrap(item, width = 10)), color="black") +
      
      theme_void() + 
      theme(legend.position="none") + 
      coord_equal()
      
      ggsave("c:/xl_toolbox/plot.wmf", width = 15, height = 15, units = "cm")
      
      #save output
      options("openxlsx.numFmt" = "#,#0.00")
      style <- createStyle( fontSize = 14, fontName = "Arial", textDecoration = "bold", halign = "left")
      
      wb <- createWorkbook() 
      addWorksheet(wb, sheetName = "Results")
      
      showGridLines(wb, 1, showGridLines = FALSE)
      
      writeData(wb, sheet = 1, "Circle Plot", startRow=2,startCol =2)
      addStyle(wb, sheet = 1, style, rows = 2, cols = 2)
      insertImage(wb, 1, "c:/xl_toolbox/plot.wmf", width = 15, height = 15, startRow = 3, startCol = 2, units = "cm")
      
      saveWorkbook(wb, "C:/xl_toolbox/data_out.xlsx", overwrite = TRUE)
      
      #hand back to excel
write("no range",file="c:/xl_toolbox/finished.log")
      

  }, error = function(err.msg){
    # Add error message to the error log file
    write(toString(err.msg), log.path, append=TRUE)
  }
  )

