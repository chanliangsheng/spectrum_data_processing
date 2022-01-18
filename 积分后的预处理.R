library("xlsx")
setwd("D:/研究生学习/mission/广谱质谱数据处理")
#设置工作路径

pre_dealing <- function(file , sheetName , outputName){
  rawdata <- read.xlsx(file = file,sheetName = sheetName)
  #??ȡԭʼ?ļ?

  raw_data_sample_name <- rawdata$Sample.Name
  end_rawdata <- cbind(raw_data_sample_name,rawdata$Component.Name)
  end_rawdata <- cbind(end_rawdata,rawdata$Area)
  end_rawdata <- cbind(end_rawdata,rawdata$Area.Ratio)
  end_rawdata <- cbind(end_rawdata,rawdata$Signal...Noise)

  colnames(end_rawdata) <- c("Sample.Name","Component.Name","Area","Area.Ratio","Signal...Noise")
  #??????
  end_rawdata <- as.data.frame(end_rawdata)
  #?Ӿ?????Ϊ???ݿ?

  nrow <- length(unique(end_rawdata$Component.Name)) + 1
  #??????��???ڴ?л?????Ƹ???????

  ncol <- length(end_rawdata$Component.Name) / length(unique(end_rawdata$Component.Name)) + 1
  #??????ô????

  store_matrix <- matrix(nrow = nrow,ncol = ncol)
  #???ɿվ??????ڴ洢

  sample_name <- c()
  #????????��???ڴ洢??Ʒ????

  for (i in 1:(ncol-1)) {
    sample_name <- c(sample_name,end_rawdata$Sample.Name[i*length(unique(end_rawdata$Component.Name))-1])
  }


  store_matrix[-1,1] <- end_rawdata$Component.Name[1:length(unique(end_rawdata$Component.Name))]
  store_matrix[1,-1] <- sample_name

  count <- 1

  for (i in 1:(length(store_matrix[1,]) - 1)) {
    store_matrix[-1,i + 1] <- end_rawdata$Area[count:(length(unique(end_rawdata$Component.Name)) + count - 1)]
    count <- count + length(unique(end_rawdata$Component.Name))
  }

  write.xlsx(store_matrix,file = outputName,sheetName = "Area",row.names = FALSE,col.names = FALSE)
  #д??excel Component.Name

  count <- 1

  for (i in 1:(length(store_matrix[1,]) - 1)) {
    store_matrix[-1,i + 1] <- end_rawdata$Area.Ratio[count:(length(unique(end_rawdata$Component.Name)) + count - 1)]
    count <- count + length(unique(end_rawdata$Component.Name))
  }
  write.xlsx(store_matrix,file = outputName,sheetName = "Area.Ratio",row.names = FALSE,col.names = FALSE,append = TRUE)
  #д??excel Area.Ratio

  count <- 1
  for(i in 1:(length(store_matrix[1,]) - 1)) {
    store_matrix[-1,i + 1] <- end_rawdata$Signal...Noise[count:(length(unique(end_rawdata$Component.Name)) + count - 1)]
    count <- count + length(unique(end_rawdata$Component.Name))
  }
  write.xlsx(store_matrix,file = outputName,sheetName = "Signal...Noise",row.names = FALSE,col.names = FALSE,append = TRUE)
  #д??excel Area.Ratio
}


pre_dealing(file = "20211015-YM-ZX-GFPT1.xlsx" , sheetName = "neg",outputName = "neg.xlsx")
