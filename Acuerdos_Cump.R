#01--RUTA Y LIBRERIA--------
library(easypackages)
pqt<- c("googlesheets4","gargle", "tidyverse","lubridate", "janitor","knitr","formattable","bizdays" , "openxlsx","readxl", "stringr")
libraries(pqt)

#Creamos rutas para no exceder la ruta de los archivos
ruta_DataInput<- "C:/Users/jach_/OneDrive/Documentos/OEFA/Inputs/Data/2022/"
ruta_BolOutput<- "C:/Users/jach_/OneDrive/Documentos/OEFA/Outputs/Boletin/2022/"
ruta_DatOutput<- "C:/Users/jach_/OneDrive/Documentos/OEFA/Outputs/Boletin/2022/Data/"
ruta_descargas<- "C:/Users/jach_/Downloads/"

#Pegamos las direcciones cortas
load(paste0(ruta_DataInput,"Doc_Febrero_2022.RData"))
load(paste0(ruta_DataInput,"Acc_Febrero_2022.RData"))
load(paste0(ruta_DataInput,"Inf_Febrero_2022.RData"))


#Usuario
gs4_user()

#Configurar autorizaciones
gs4_auth(
  email = "jarias@oefa.gob.pe",
  token = "Gargle2.0"
)

#rutas actualizadas
id_acu<- c(
"1Oq4W2NbKpIbhYrf0OLu0KVbQggwbWzxIQpvJB3mXte4",
"1Kj-18Te32qbfOlZ0F8j17tr8kWh9-NyjvLqwQLahLk0",
"1tYKnzD1WPGPmQLCB0759e-afu9uoaZ2-v8oR0HEfEeI",
"1dXJcVfdk7e-z2SI0bhaYIKQaJKX19P2pBWCBKCYHGtY",
"11pOB4qu9xLWNYhhCPMgL4E9Nh9IpV6ktxmswqfLP2Ms",
"14lBSuuLcNsw039F3aurqOmOGDNd3a-xxqm6IOoupdd8",
"1GPqh-2aSUqxtUe2s6_wEvltBExNUQb_aL_QUHJc7N28",
"1UvcUViphvvXi2gybOxIaj2X5WxeO0C2IV5tDxrIjq0U")



#BUscar archivos con informacion 
archivos<- gs4_find("BD ACUERDO CUMPLIMIENTO -", n_max=500) %>% 
  filter(!str_detect(name,"@|Copia")) %>%
  filter(id %in% id_acu) %>% arrange(name)

#Confirmas permisos
1

#Creamos 2 listas para acuerdos y para reuniones
file.list1<- list("")
file.list2<- list("")
file.list3<- list("")

#Importar datos reuniones
for(i in 1:8){
  file.list1[[i]]<- as.data.frame(read_sheet(ss=archivos$id[i],
                                            sheet = "Reuniones",
                                            range = "A3:N2000",
                                            #skip = 2,
                                            col_types = "ccDcccccccccic")%>% 
                                    mutate(BD = archivos$name[i]))
} #Hasta CRES

#Importar datos acuerdos
for(i in 1:8){
  file.list2[[i]]<- as.data.frame(read_sheet(ss=archivos$id[i],
                                            sheet = "Base Acuerdos",
                                            range = "A3:S2000",
                                            #skip = 2,
                                            col_types = "cciccDccDccccccccDc")%>% 
                                    mutate(BD = archivos$name[i]))
} #Hasta CRES

#Importar datos acuerdos
for(i in 1:8){
  file.list3[[i]]<- as.data.frame(read_sheet(ss=archivos$id[i],
                                             sheet = "Datos adicionales Acuerdos",
                                             range = "A3:Z2000",
                                             #skip = 2,
                                             col_types = "ccccDDccccciiiiiiccccccccc")%>% 
                                    mutate(BD = archivos$name[i]))
} #Hasta CRES



#Finalizar conexion
gs4_deauth()


#Base de Reuniones
REUNIONES<- bind_rows(file.list1) %>%
  mutate_if(is.character,function(x){str_to_upper(x)}) %>%
  mutate_if(is.character, function(x){gsub("[[:cntrl:]]", " ",x)}) %>%
  mutate_if(is.POSIXct,function(x){as.Date(x, format=" %H:%M:%S")})

#Base de Acuerdos de cumplimiento
ACUERDOS_PREV<- bind_rows(file.list2) %>% 
  mutate_if(is.character, function(x){str_to_upper(x)}) %>%
  mutate_if(is.character, function(x){gsub("[[:cntrl:]]", " ",x)}) %>%
  #rename("NRO_ACUERDO" = NRO_ACUERDO...2) %>%
  #rename("COR_ACUERDO" = NRO_ACUERDO...3) %>%
  rename("NRO_EXPEDIENTE_VER" = NRO_EXPEDIENTE) %>%
  rename("COD_ACCION_VER" = COD_ACCION) %>%
  rename("COMENTARIOS_ACUERDOS" = COMENTARIOS)
  

#Unimos en una sola BD
ACUERDOS<- REUNIONES %>% right_join(ACUERDOS_PREV) %>%
  filter(!is.na(NRO_ACUERDO), !is.na(FEC_REUNION)) %>%
  mutate(SECTOR = case_when(
    str_detect(BD,"MIN") ~ "MINERIA",
    str_detect(BD,"HID") ~ "HIDROCARBUROS",
    str_detect(BD,"ELE") ~ "ELECTRICIDAD",
    str_detect(BD,"IND") ~ "INDUSTRIA",
    str_detect(BD,"PES") ~ "PESQUERIA",
    str_detect(BD,"AGR") ~ "AGRICULTURA",
    str_detect(BD,"CAM") ~ "CONSULTORAS AMBIENTALES",
    str_detect(BD,"RES") ~ "RESIDUOS SOLIDOS"),
    ANO = year(FEC_REUNION)) %>% select(-BD)
# Corrections of data


#Base de Datos adicionales Acuerdos
ADICIONALES<- bind_rows(file.list3) %>%
  mutate_if(is.character,function(x){str_to_upper(x)}) %>%
  mutate_if(is.character, function(x){gsub("[[:cntrl:]]", " ",x)}) %>%
  mutate_if(is.POSIXct,function(x){as.Date(x, format=" %H:%M:%S")})


#Guardamos la data
save(ACUERDOS, file = paste0(ruta_DatOutput,"Acu_Febrero_2020.RData"))
load(paste0(ruta_DatOutput,"Acu_Febrero_2020.RData"))

#Guardamos el reporte
if(1==1){
  
  wb<- createWorkbook()
  #----------------------
  addWorksheet(wb,"REUNIONES", gridLines = FALSE)
  addWorksheet(wb,"BD", gridLines = FALSE)
  addWorksheet(wb,"ADICIONAL", gridLines = FALSE)
  
  #Tabla 1
  writeData(wb,"REUNIONES",as.data.frame(paste(c("0. REPORTE CONSOLIDADO REUNIONES / "),c(Sys.time()), sep = "-"))
            , startRow = 1 ,startCol = "A", colNames = FALSE)
  writeDataTable(wb,"REUNIONES",(REUNIONES %>% filter(!is.na(COD_REUNION))) ,startRow = 2,startCol = "A")
  writeData(wb,"BD",as.data.frame(paste(c("1. REPORTE CONSOLIDADO REUNIONES CON ACUERDOS DE CUMPLIMIENTO / "),c(Sys.time()), sep = "-"))
            , startRow = 1 ,startCol = "A", colNames = FALSE)
  writeDataTable(wb,"BD",ACUERDOS,startRow = 2,startCol = "A")
  writeData(wb,"ADICIONAL",as.data.frame(paste(c("2. REPORTE CONSOLIDADO DE DATOS ADICIONALES SOBRE ACUERDOS DE CUMPLIMIENTO / "),c(Sys.time()), sep = "-"))
            , startRow = 1 ,startCol = "A", colNames = FALSE)
  writeDataTable(wb,"ADICIONAL",ADICIONALES,startRow = 2,startCol = "A")
  
  saveWorkbook(wb, paste0(ruta_BolOutput,"Acuerdos_Febrero_2022.xlsx"), overwrite = TRUE)
}

