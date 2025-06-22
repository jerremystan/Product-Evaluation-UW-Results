library(readxl)
library(dplyr)
library(writexl)
library(openxlsx)

prodev_df <- read_excel("8. Rincian produk TMI.v.1.xlsx", 
                        sheet = "From ProDev")

max_column <- 8

base_df <- prodev_df[,1:24]

join_df <- base_df

for (i in 2:max_column) {
  assign(paste0("df_", i), prodev_df[, c(1:22, 22+(2*i-1), 23+(2*i-1))])
  
  df_temp <- get(paste0("df_", i))
  colnames(df_temp)[c(23,24)] <- c("Major Class Code", "Contract Type Code")
  df_temp <- df_temp[!is.na(df_temp$'Major Class Code'),]
  join_df <- union(join_df, df_temp)
  
}



join_df <- join_df[order(join_df$No.),]
join_df$`Manfaat yang diberikan` <- NA

write.xlsx(join_df, "ProdukList.xlsx", overwrite = TRUE)
