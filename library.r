#check required packages are available

pack_list <- c("ca","reshape2","psych","nFactors","googleVis","ggplot2","packcircles","viridis",
            "stringr","relaimpo","rvest","qgraph","RColorBrewer","pricesensitivitymeter",
            "cellranger","tidyverse","httr","readxl","openxlsx","devEMF","jiebaR","wordcloud",
            "methods","datasets","utils","grDevices","graphics","stats")

## Now load or install&load all
package.check <- lapply(
  pack_list,
  FUN = function(x) {
    if (!require(x, character.only = TRUE)) {
      install.packages(x, dependencies = TRUE)
      library(x, character.only = TRUE)
    }
  }
)






