library(dplyr)
library(readxl)
library(writexl)
library(openxlsx)

mcct <- c("FAPT", "FCMI", "FEAQ", "FFDH", "FHOM", "FIAR", "FISP", "FLPR", "FOCT", "FPAR", "FPOV", "FRCC", "FSME", "FUMP", "FXOL", "LAMB", "LDNO", "LELI", "LGEN", "LPER", "LPIA", "LPIC", "LPIM", "LPRO", "LWCI", "MEXP", "MFFL", "MIMP", "MINL", "MINR", "MMXL", "MTOL", "PACC", "PALP", "PATH", "PIDB", "PIMA", "PIPA", "PMPK", "PPBI", "PPXL", "PTKA", "PTYR", "SADS", "TDTI", "TDTP", "TORC", "TORP", "TOTI", "TOTP", "TTPF", "TTPW", "TTXL", "TWRP", "TWTI", "TWTP", "VACL", "VBUG", "VCOM", "VCSI", "VCYB", "VFIG", "VGLF", "VHIO", "VMNY", "VMOV", "VPAC", "VTCS", "VTYR", "VVXL", "PAML", "VOTP", "FHOP", "LBLW", "PACT", "PATP", "PIGP", "PLIA", "TMTP", "VCRD", "VDTI", "VDTP", "VHOS", "VLIB", "VMTI", "VMTP", "VOFC", "VOTA", "VOTI", "VPAS", "VWTI", "VWTP")

df_gwp <- read_excel("Raw Files DC Total/GWP Total DC.xlsx")
df_gwp$MCCT <- paste0(df_gwp$`Major Class Code`,df_gwp$`Contract Type Code`)

df_os <- read_excel("Raw Files DC Total/OS Total DC.xlsx")
df_os$MCCT <- paste0(df_os$`Major Class Code`,df_os$`Contract Type Code`)

df_paid <- read_excel("Raw Files DC Total/Paid Total DC.xlsx")
df_paid$MCCT <- paste0(df_paid$`Major Class Code`,df_paid$`Contract Type Code`)

df_upr <- read_excel("Raw Files DC Total/UPR Total DC.xlsx")
df_upr$MCCT <- paste0(df_upr$`Major Class Code`,df_upr$`Contract Type Code`)

for (i in mcct){
  
  new_loc <- paste0("Breakdown Result/Raw Files DC ",i)
  dir.create(new_loc)
  
  filter_gwp <- df_gwp[df_gwp$MCCT == i,]
  sum_gwp <- filter_gwp %>%
    group_by(`D/Channel Code`,`GWP - Year`) %>%
    summarise(`GWP Aft Co-Out` = sum(`GWP Aft Co-Out`),
              `Net GWP` = sum(`Net GWP`),
              `Reinsurance GWP` = sum(`Reinsurance GWP`),
              `Comms, VAT, Fees Aft Co-Out` = sum(`Comms, VAT, Fees Aft Co-Out`),
              `Net Comms, VAT, Fees` = sum(`Net Comms, VAT, Fees`),
              `Reinsurance Commission` = sum(`Reinsurance Commission`))
  write.xlsx(sum_gwp,paste0(new_loc,"/GWP ",i,".xlsx"))
  
  filter_os <- df_os[df_os$MCCT == i,]
  sum_os <- filter_os %>%
    group_by(`D/Channel Code`,`RPT - Month (YYYYMM)`) %>%
    summarise(`Claims Outstanding After Co-Out AsAt` = sum(`Claims Outstanding After Co-Out AsAt`),
              `Net Claims Outstanding AsAt` = sum(`Net Claims Outstanding AsAt`))
  write.xlsx(sum_os,paste0(new_loc,"/OS ",i,".xlsx"))
  
  filter_paid <- df_paid[df_paid$MCCT == i,]
  sum_paid <- filter_paid %>%
    group_by(`D/Channel Code`,`App - Year`) %>%
    summarise(`Claims Paid After Co-Out` = sum(`Claims Paid After Co-Out`),
              `Net Claims Paid` = sum(`Net Claims Paid`),
              `Reinsurance Recovery` = sum(`Reinsurance Recovery`))
  write.xlsx(sum_paid,paste0(new_loc,"/Paid ",i,".xlsx"))
  
  filter_upr <- df_upr[df_upr$MCCT == i,]
  sum_upr <- filter_upr %>%
    group_by(`D/Channel Code`,`EP - Month (YYYYMM)`, `Long Term Policy Flag (Y/N)`) %>%
    summarise(`UPR Aft Co-Out` = sum(`UPR Aft Co-Out`),
              `Un-incurred Commission Aft Co-Out` = sum(`Un-incurred Commission Aft Co-Out`),
              `Net UPR` = sum(`Net UPR`),
              `Net Un-incurred Commission` = sum(`Net Un-incurred Commission`))
  write.xlsx(sum_upr,paste0(new_loc,"/UPR ",i,".xlsx"))
  
  
}
