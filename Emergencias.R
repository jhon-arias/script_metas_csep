#------------------------------------------------------------------------------------------------
library(easypackages)
pqt<- c("tidyverse","readxl", "openxlsx","lubridate","bizdays" ,"janitor", "knitr","formattable")
libraries(pqt)
#-----

setwd("C:/Users/jach_/OneDrive/Documentos/OEFA/Inputs/Data2021/")

#Cargamos las acciones
load("Diciembre/Acc_Diciembre_2021.RData")
load("Diciembre/Emergencias_Diciembre_2021.RData")

Evaluacion <- "2021-12-31" # Cambiar la fecha de corte
#--
Emergencias<- Emergencias_R %>%
  left_join( Acciones_R %>% select(TXCOORDINACION, TXCUC) %>% rename("TX_ACCSUP"=TXCUC)) %>%
  mutate(TXCOORDINACION = case_when(
    str_detect(TXCOORDINACION, "MINER") ~ "MINERIA",
    str_detect(TXCOORDINACION, "HIDROCAR") ~ "HIDROCARBUROS",
    str_detect(TXCOORDINACION, "ELECTRIC") ~ "ELECTRICIDAD",
    str_detect(TXCOORDINACION, "INDUSTR") ~ "INDUSTRIA",
    str_detect(TXCOORDINACION, "PES") ~ "PESCA",
    str_detect(TXCOORDINACION, "AGRIC") ~ "AGRICULTURA",
    str_detect(TXCOORDINACION, "CONSUL") ~ "CONSULTORAS AMBIENTALES",
    str_detect(TXCOORDINACION, "RESID") ~ "RESIDUOS SOLIDOS",
    str_detect(TXCOORDINACION, "OFIC") & str_detect(TX_SUB_SECTOR,"Hidro") ~ "OD_HID",
    str_detect(TXCOORDINACION, "OFIC") & str_detect(TX_SUB_SECTOR,"Resid") ~ "OD_RES")) %>%
  mutate_if(is.character, function(x){toupper(x)}) %>%
  mutate_at(vars(starts_with("FE_FECHA")), function(x){as.Date(x)}) %>%
  filter(!is.na(TX_ACCSUP),
         #FE_FECHA_FIN >= as.Date("2020-03-01"), 
         FE_FECHA_FIN <= as.Date(Evaluacion)) 

#Reporte de INAF todas---------------
Reporte_all<- Emergencias %>%
  filter(FE_FECHA_FIN >= as.Date("2021-01-01")) %>% #Solo 2021
  group_by(TXCOORDINACION) %>%
  summarise(Acciones = n_distinct(TX_ACCSUP),
            Emergencias = n_distinct(TX_CODIGO)) %>% adorn_totals(c("row"))
Reporte_all


#Reporte de INAF Inmediatas y de verificacion---------------
Reporte_inmediatas<- Emergencias %>%
  filter(FE_FECHA_FIN >= as.Date("2021-01-01"), #Solo 2021
         str_detect(TX_TIPO_ATENCION, "VERIF"),
         str_detect(TX_VERIFICACION,"^INMEDIATA")) %>%
  group_by(TXCOORDINACION) %>%
  summarise(Acciones = n_distinct(TX_ACCSUP),
            Emergencias = n_distinct(TX_CODIGO)) %>% adorn_totals(c("row"))
Reporte_inmediatas 

# #Reporte de Lo que esta en la BD Antigua (Inmediatas y Principales):-----
# Eme.antes<- read_excel("C:/Users/jach_/Downloads/Emergencias_Febrero2020.xlsx", sheet = "CONSOLIDADO", skip = 1) %>%
#   rename("TXCOORDINACION"= COORD_ATEND_EME) %>%
#   mutate(TXCOORDINACION =  case_when(
#     str_detect(TXCOORDINACION,"RESID") ~ "RESIDUOS SOLIDOS",
#     str_detect(TXCOORDINACION,"CODE") & str_detect(SUBSECTOR,"HIDR") ~ "OD_HID",
#     str_detect(TXCOORDINACION,"CODE") & str_detect(SUBSECTOR,"RESI") ~ "OD_RES",
#     TRUE ~ TXCOORDINACION )) %>%
#   # mutate(TXCOORDINACION = ifelse(str_detect(TXCOORDINACION,"RESIDUOS"), "RESIDUOS SOLIDOS", TXCOORDINACION),
#   #        TXCOORDINACION = ifelse(str_detect(TXCOORDINACION,"CODE"), "OD", TXCOORDINACION)) %>%
#   filter(!is.na(CUC), str_detect(PROIRIDAD,"^INMEDIATA"),FECHA_FIN_SUP >= as.POSIXct("2019-01-01")) %>% 
#   mutate(ANO_FIN_SUP= year(FECHA_FIN_SUP)) %>%
#   group_by(TXCOORDINACION#, ANO_FIN_SUP
#            ) %>%
#   summarise(Acciones = n_distinct(CUC),
#             Emergencias = n_distinct(COD_EME))
# Eme.antes  
# 
# 
# # Reporte_Antes<- tribble(
# #   ~TXCOORDINACION, ~ TX_ACCSUP, ~ TX_CODIGO,
# #   "MINERIA", 5, 5,
# #   "HIDROCARBUROS",23, 27,
# #   "ELECTRICIDAD", 5,5,
# #   "INDUSTRIA", 0,0,
# #   "PESCA", 1, 2,
# #   "AGRICULTURA", 0, 0,
# #   "CONSULTORAS AMBIENTALES", 0, 0,
# #   "RESIDUOS SOLIDOS", 1, 1,
# #   "OD", 10, 10) 
# 
# #Reporte de Emergencias total 2020--------
# 
# RptEmergencias<- Reporte %>% full_join(Eme.antes) %>%
#   replace(is.na(.),0) %>%
#   mutate(Acciones = Acciones1 + Acciones,
#          Emergencias = Emergencias1 + Emergencias ) %>%
#   rename("Subsector" = TXCOORDINACION) %>%
#   select(-c(Acciones1,Emergencias1)) %>%
#   adorn_totals()
# RptEmergencias
#   
#Exportamos----

#Guardamos el reporte
if(1==1){
  
  wb<- createWorkbook()
  #----------------------
  addWorksheet(wb,"BD", gridLines = FALSE)
  addWorksheet(wb,"REPORTE", gridLines = FALSE)
  #Tabla 1
  writeData(wb,"BD",as.data.frame(paste(c("1. BD DE EMERGENCIAS - INAF"),c(Sys.time()), sep = "-"))
            , startRow = 1 ,startCol = "A", colNames = FALSE)
  writeDataTable(wb,"BD",Emergencias,startRow = 2,startCol = "A")
  #Tabla 2
  writeData(wb,"REPORTE",as.data.frame(paste(c("1. REPORTE CONSOLIDADO EMERGENCIAS ATENDIDAS ENERO"),toupper(months(as.Date(Evaluacion))),"2021", sep = "-"))
            , startRow = 1 ,startCol = "A", colNames = FALSE)
  writeDataTable(wb,"REPORTE",Reporte_all,startRow = 2,startCol = "A")
  #Tabla 3
  writeData(wb,"REPORTE",as.data.frame(paste(c("2. REPORTE CONSOLIDADO EMERGENCIAS ATENDIDAS - VERIFICACION Y PRINCIPALES ENERO "),toupper(months(as.Date(Evaluacion))),"2021", sep = "-"))
            , startRow = 15 ,startCol = "A", colNames = FALSE)
  writeDataTable(wb,"REPORTE",Reporte_inmediatas,startRow = 16,startCol = "A")

  saveWorkbook(wb, "C:/Users/jach_/OneDrive/Documentos/OEFA/Outputs/Boletin/Emergencias_Diciembre2021.xlsx", overwrite = TRUE)
}


