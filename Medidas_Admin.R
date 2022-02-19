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
load(paste0(ruta_DataInput,"Doc_Enero_2022.RData"))
load(paste0(ruta_DataInput,"Acc_Enero_2022.RData"))
load(paste0(ruta_DataInput,"Inf_Enero_2022.RData"))

#Usuario
gs4_user()

#Configurar autorizaciones
gs4_auth(
  email = "jarias@oefa.gob.pe",
  token = "Gargle2.0"
)

#rutas actualizadas
id_med<- c(
"1nUnLGsjHh-Oc6Pgk6vLVlapOKy0GjWEMUF8zi0F7ATs",
"1WQUkNXEbUR0kzMAgiUX_qeRIA26O7OV94863UCsiguw",
"10bVhg7w_EtMs7G0Mylz2MnT8lA1cZYrBWSx99SgjdCU",
"1X4kjCq2kk69CSY2o-k3nMHckRjjekZfeQ0w1nk1rRgw",
"1wZkZejtnD3_-T2aPKDGMfkyhXLmNRYq0hzbHe0kn_iE",
"1UnuRf_PNqytlKPSPmIJZOl_jTa1e_rCM3WMuhWRJrac",
"1B3Q62Gp5Zq2szElp2wR_kgae7Mn7QSxcwQH7Qr2K2uA",
"1wvgQXrxNo-zPEHaJeEBPmk3nuHdKSqBhSVuxaCc-KTA")



#BUscar archivos con informacion 
archivos<- gs4_find("BD MEDIDAS ADMINISTRATIVAS", n_max=500) %>%
  filter(!str_detect(name,"@|Copia")) %>%
  filter(id %in% id_med) %>% arrange(name)

#Confirmas permisos
1

#Creamos una LIsta
file.list<- list("")

#Importar datos
for(i in 1:8){
  file.list[[i]]<- as.data.frame(read_sheet(ss=archivos$id[i],
                                            sheet = "Base Medidas",
                                            range = "A2:CG",
                                            #skip = 1,
                                            col_types = "icccccccDDcccDicccccccccciiiiiiicDccDDcDccDcccDccDcDcccDcDcDcDccDcccccccccccicDiicccD"))
} #Hasta CRES

#Finalizar conexion
gs4_deauth()

#Unimos en una sola BD

MEDIDAS_ADM<- bind_rows(file.list[1:8]) %>% 
  mutate_if(is.character, function(x){str_to_upper(x)}) %>%
  mutate_if(is.character, function(x){gsub("[[:cntrl:]]", " ",x)}) %>%
  rename(VENC_CONDICION= FECHA_VENC_CONDICION,
         PROY_VERIF= FECHA_PROY_VERIF)

# Corrections of data


#05--ADD VAR SUPERV-----
#Agregamos otra Ruta
setwd("C:/Users/jach_/OneDrive/Documentos/ShinyApp/Planefa2021/")
#CARGAMOS LOS FERIADOS 2019 MODIFICAR RUTA
source("Feriados.R")
create.calendar("Peru2021",holidays = feriados,weekdays = c("saturday","sunday"))

MEDIDAS_ADM<- MEDIDAS_ADM %>%
  mutate(MES_RESOL= toupper(months(FECHA_RESOL)),
         MES_RESOL = if_else(str_detect(MES_RESOL,"SEPTIEMBRE"),"SETIEMBRE",MES_RESOL)) %>%
  mutate(ANO= as.numeric(year(FECHA_RESOL))) %>%
  mutate(DIAS_INI_HECHOS_HASTA_RESOLUCION = (bizdays(FECHA_DETEC_INICIO ,FECHA_RESOL, "Peru2021")+1)) %>%
  mutate(DESCRIPCION_TIPO= case_when(
    TIPO_MEDIDA=="MP" ~ "MEDIDA PREVENTIVA",
    TIPO_MEDIDA=="MCP" ~ "MEDIDA DE CARACTER PARTICULAR",
    TIPO_MEDIDA=="RA" ~ "REQUERIMIENTO EN EL MARCO SEIA", TRUE ~ "OTROS"),
    N = case_when(
      nchar(ID)==1 ~ paste0("000",ID),
      nchar(ID)==2 ~ paste0("00",ID), TRUE ~ paste0("0",ID)),
    COD_MED   = "",
    COD_OEFA  = "",
    SUBSECTOR = "",
    DOCUMENTO = "",
    PUBLICAR  = "",
    FECHA_ACTUAL = today())

#08--NUEVA BD MERGE CON INAPS----
MEDIDAS_ADM2<- MEDIDAS_ADM %>% left_join(Acciones_R, by=c("CUC" = "TXCUC"))
#09--REMPLAZO VALORES----
MEDIDAS_ADM2<- MEDIDAS_ADM2 %>%
  mutate(EXPEDIENTE= if_else(!is.na(TXNUMEXP), TXNUMEXP,EXPEDIENTE)) %>%
  mutate(RUC= if_else(!is.na(TXNUMDOC_ADM) ,TXNUMDOC_ADM, RUC)) %>%
  mutate(ADMIN= if_else(!is.na(TXADMINISTRADO_ADM), TXADMINISTRADO_ADM, ADMIN)) %>%
  mutate(SUBSECTOR= SECTOR) %>%
  select(ID:FECHA_ACTUAL)

#10--GUARDAMOS LA BD FINAL----
save(MEDIDAS_ADM2, file = paste0(ruta_DatOutput,"Med_Enero_2022.RData"))
load(paste0(ruta_DatOutput,"Med_Enero_2022.RData"))

#Meses Evaluados ========
EVALUACION<- as.Date("2022-01-31")

#11--REPORTE CSEP-----
REP_CSEP<- MEDIDAS_ADM2 %>%
  filter(FECHA_RESOL<= EVALUACION, 
         FECHA_RESOL>= as.Date("2016-09-01")) %>%
  select(ID, EXPEDIENTE, CUC,RUC, MED_DIC_EN, ADMIN, UF, SECTOR,FECHA_DETEC_INICIO, FECHA_DETEC_FIN,INF_TECNICO, 
         ORIGEN_RESOL,RESOL, FECHA_RESOL,MES_RESOL,ANO, DIAS_INI_HECHOS_HASTA_RESOLUCION, CORR_MEDIDA,
         OBLIG_RESUMEN, TIPO_MEDIDA,DESCRIPCION_TIPO, RESOL_PARRAF_MULTA, REGION, PROVINCIA, 
         DISTRITO,TIPIFICACION,DESCR_MANDATO, SUST_MANDATO, HUMANOS, FLORA, FAUNA, AGUA, AIRE, SUELO, PLAZO_DIAS, CUENTA_NOTIF,
         FECHA_NOTIF, PRORROGA_AMPLIAC, RESOL_PRORROGA,FECHA_RESOL_PRORROGA_AMPLIAC, FECHA_NOTIF_PRORROGA_AMPLIAC,VENC_CONDICION,
         FECHA_VENCIMIENTO, PROY_VERIF, ESTADO_MEDIDA, FECHA_VERIF, CUC_VERIF, EXPEDIENTE_VERIF, INF_CARTA_VERIF, FECHA_INF_VERIF,
         ANALIS_LAB, SOLIC_VARIA, FECHA_SOLIC_VARIA,RESOL_VARIA,INTERP_RECURSO, FIRME, RECONS, FECHA_RECONS, RESOL_RECONS, APELA,
         FECHA_APELA, RESOL_TFA, REVOC_CONF_NUL, DETAL_COORD,CANTIDAD_N,EXTENSION_N,PELIGROSIDAD_N,MEDIO_POTENCIALMENTE_AFECTADO_N,CANTIDAD_H,
         EXTENSION_H,	PELIGROSIDAD_H,PERSONAS_POTENCIALMENTE_EXPUESTAS_H,PROBABILIDAD,AMERITA_MULTA_COERCITIVA,NUMERO_DE_MULTA,
         ULTIMA_RD_MULTA_COERCITIVA,FECHA_RD_MULTA_COERCITIVA,MONTO_SOLES_MULTA,MULTA_COERCITIVA_EN_UIT,ESTADO_MULTA_COERCITIVA,
         SE_REALIZO_EJECUCION_FORZOSA,ACTA_EJECUCION_FORZOSA,FECHA_ACTA_EJECUCION_FORZOSA,MES_RESOL,ANO,DIAS_INI_HECHOS_HASTA_RESOLUCION,
         DESCRIPCION_TIPO,N,COD_MED,COD_OEFA,SUBSECTOR,DOCUMENTO,PUBLICAR,FECHA_ACTUAL) %>%
    arrange(SECTOR,ID) 

#12--REPORTE CSIG----
PUBLIC<- MEDIDAS_ADM2 %>%
  group_by(SECTOR,RESOL,FIRME) %>%
  summarise(TOTAL=n()) %>%
  spread(FIRME,TOTAL) %>%
  mutate(PUB=if_else(SI>0 & is.na(NO),1,0)) %>%
  select(SECTOR,RESOL,PUB)
  
REP_CSIG<-
  MEDIDAS_ADM2 %>%
  filter(FECHA_RESOL <= EVALUACION, #Cambiar con información de  
         FECHA_RESOL >= as.Date("2016-09-01")) %>%
  left_join(PUBLIC,by =c("SECTOR","RESOL")) %>%
  mutate(PUBLICAR=PUB) %>%
  mutate(PUBLICAR = if_else(is.na(PUBLICAR),0,PUBLICAR)) %>%
  select(N,COD_MED, COD_OEFA, EXPEDIENTE, CUC, RUC, MED_DIC_EN, ADMIN, UF, SECTOR, SUBSECTOR,FECHA_DETEC_INICIO, FECHA_DETEC_FIN, INF_TECNICO, RESOL, FECHA_RESOL,
         DOCUMENTO, CORR_MEDIDA, OBLIG_RESUMEN, DESCRIPCION_TIPO, REGION, PROVINCIA, DISTRITO, DESCR_MANDATO, SUST_MANDATO, HUMANOS, FLORA,FAUNA,AGUA,AIRE,SUELO,
         PLAZO_DIAS,CUENTA_NOTIF, FECHA_NOTIF, PRORROGA_AMPLIAC,RESOL_PRORROGA, FECHA_RESOL_PRORROGA_AMPLIAC, FECHA_NOTIF_PRORROGA_AMPLIAC,
         FECHA_VENCIMIENTO, ESTADO_MEDIDA, INF_CARTA_VERIF, ANALIS_LAB,SOLIC_VARIA, FECHA_SOLIC_VARIA, RESOL_VARIA, INTERP_RECURSO,FIRME, RECONS, 
         FECHA_RECONS, RESOL_RECONS, APELA, FECHA_APELA, RESOL_TFA, REVOC_CONF_NUL, DETAL_COORD, PUBLICAR, FECHA_ACTUAL) %>%
  arrange(FECHA_RESOL,SECTOR,N)

colnames(REP_CSIG)<- c("NRO","COD_MED","COD_OEFA","N_EXPEDIENTE","CUC","RUC","MED_DIC_EN","ADMINISTRADO","UNIDAD_FISCALIZABLE","SECTOR","SUBSECTOR",
                       "FEC_DETECC_INIC","FEC_DETECC_FIN","INFOR_TEC","RESOLUCION","FEC_RESOL","DOCUMENTO","CORRELA_MEDIDA","OBLIGA_RESUMEN",
                       "TIPO_MEDIDA","REGION","PROVINCIA","DISTRITO","DESCRIP_MANDA","SUSTEN_MANDA","COMP_HUMANOS","COMP_FLORA","COMP_FAUNA",
                       "COMP_AGUA","COMP_AIRE","COMP_SUELO","PLAZO_DIAS","NOTIFICACION","FEC_NOTIFICACION","PRO_AMP","RESOL_PRO_AMP","FEC_RESOL_PRO_AMP",
                       "FEC_NOTI_PRO_AMP","FEC_PLAZ_VENC","ESTADO_MED","INFOR_SUP_CUMP","ANA_LAB_MED","SOLIC_VARIACION","FEC_SOLIC","RESOL_VARIACION",
                       "INTERP_RECURSO","MED_FIRME","RECONSIDERA","FEC_RECONSIDERA","RESOL_RECONSIDERA","APELACION","FEC_APELA","RESOL_TFA",
                       "REVOCA_CONFIR_NULA","OBS_COORD","PUBLICAR","FEC_ACTUAL")
#EXPORTAR PARA CSIG------
#write.xlsx(REP_CSIG,"C:/Users/jach_/OneDrive/Documentos/OEFA/Outputs/CSIG/BASE DE DATOS MEDIDAS ADMINISTRATIVAS - GENERAL 31.01.2022(CSIG).xlsx", startRow=1, colNames=TRUE)


#---REPORTES PARA MEDIDAS ADMIN----
if(3==3){
#Resoluciones Dictadas y Medidas Impuestas Dictadas
RESOL_MED<- REP_CSEP %>%
  mutate(REV= case_when(
    str_detect(ESTADO_MEDIDA,"NO AMERITA") ~ 1,
    str_detect(REVOC_CONF_NUL,"REVOCADA") ~ 1,
    TRUE~ 0 )) %>%
   filter(REV<1) %>%
  group_by(ANO,SECTOR,RESOL)%>% summarise(CANTIDAD_MEDIDAS=n()) %>%
  group_by(ANO, SECTOR) %>% 
  summarise(CANTIDAD_RESOL=n(),
            CANTIDAD_MEDIDAS = sum(CANTIDAD_MEDIDAS, na.rm = TRUE)) %>%
  adorn_totals("row")
RESOL_MED

#cantidad de Medidas segun estado de firmeza
MED_FIRM<- REP_CSEP %>%
  mutate(REV= case_when(
    str_detect(ESTADO_MEDIDA,"NO AMERITA") ~ 1,
    str_detect(REVOC_CONF_NUL,"REVOCADA") ~ 1,
    TRUE~ 0 )) %>%
  filter(REV<1) %>%
  group_by(ANO,SECTOR,FIRME)%>% summarise(CANTIDAD_MEDIDAS=n()) %>% spread(FIRME,CANTIDAD_MEDIDAS) %>%
  group_by(ANO,SECTOR)%>% summarise(SI_FIRME = sum(SI, na.rm = TRUE),NO_FIRME = sum(NO, na.rm = TRUE)) %>%
  mutate(TOTAL= SI_FIRME + NO_FIRME) %>%
  adorn_totals("row")
MED_FIRM

#Cantidad de Medidas segun tipo de medida
MED_TIPO<- REP_CSEP %>%
  mutate(REV= case_when(
    str_detect(ESTADO_MEDIDA,"NO AMERITA") ~ 1,
    str_detect(REVOC_CONF_NUL,"REVOCADA") ~ 1,
    TRUE~ 0 )) %>%
  filter(REV<1) %>%
  group_by(ANO,SECTOR,DESCRIPCION_TIPO)%>% summarise(CANTIDAD_MEDIDAS=n()) %>% spread(DESCRIPCION_TIPO,CANTIDAD_MEDIDAS) %>%
  group_by(ANO,SECTOR)%>% summarise(PREVENTIVA = sum(`MEDIDA PREVENTIVA`, na.rm = TRUE),
                                           CARACTER_PARTICULAR = sum(`MEDIDA DE CARACTER PARTICULAR`, na.rm = TRUE),
                                           REQERIMIENTO_ACTUALIZACION_SEIA=sum(`REQUERIMIENTO EN EL MARCO SEIA`, na.rm = TRUE),
                                           OTROS=sum(OTROS, na.rm = TRUE)) %>%
  mutate(TOTAL= PREVENTIVA + CARACTER_PARTICULAR + REQERIMIENTO_ACTUALIZACION_SEIA + OTROS) %>%
  adorn_totals("row")
MED_TIPO

#Cantidad de Medidas segun estado de cumplimiento
MED_CUMPL<- REP_CSEP %>%
  mutate(REV= case_when(
    str_detect(ESTADO_MEDIDA,"NO AMERITA") ~ 1,
    str_detect(REVOC_CONF_NUL,"REVOCADA") ~ 1,
    TRUE~ 0 )) %>%
  filter(REV<1) %>%
  group_by(ANO,SECTOR,ESTADO_MEDIDA)%>% summarise(CANTIDAD_MEDIDAS=n()) %>% 
  spread(ESTADO_MEDIDA,CANTIDAD_MEDIDAS) %>%
  group_by(ANO,SECTOR)%>% summarise(CUMPLIDA = sum(CUMPLIDA, na.rm = TRUE),
                                    INCUMPLIDA=sum(INCUMPLIDA, na.rm = TRUE),
                                    EN_EJECUCION = sum(`EN EJECUCION`, na.rm = TRUE),
                                    PDTE_INFORME=sum(`PENDIENTE DE INFORME`, na.rm = TRUE),
                                    PDTE_VERIFICACION=sum(`PENDIENTE DE VERIFICACION`, na.rm = TRUE)) %>%
  mutate(TOTAL= CUMPLIDA + INCUMPLIDA + EN_EJECUCION + PDTE_INFORME + PDTE_VERIFICACION) %>% adorn_totals("row")
MED_CUMPL

#TIEMPO HASTA EL DICTADO DE MEDIDA
MED_DIAS<- REP_CSEP %>%
  mutate(REV= case_when(
    str_detect(ESTADO_MEDIDA,"NO AMERITA") ~ 1,
    str_detect(REVOC_CONF_NUL,"REVOCADA") ~ 1,
    TRUE~ 0 )) %>%
  filter(REV<1) %>%
  group_by(ANO,SECTOR)%>% 
  summarise(DIAS_MIN= round(min(DIAS_INI_HECHOS_HASTA_RESOLUCION,na.rm = TRUE)+1, digits = 0),
            DIAS_MEAN=round(mean(DIAS_INI_HECHOS_HASTA_RESOLUCION,na.rm = TRUE)+1, digits = 0),
            DIAS_MAX= round(max(DIAS_INI_HECHOS_HASTA_RESOLUCION,na.rm = TRUE)+1, digits = 0))
MED_DIAS

#MEDIDAS POR DEPARTAMENTO
MED_DEPTO<- REP_CSEP %>%
  group_by(ANO,SECTOR,REGION,RESOL)%>% summarise(CANTIDAD_MEDIDAS=n()) %>%
  group_by(ANO,SECTOR,REGION) %>% summarise(CANTIDAD_RESOL=n(), 
                                             CANTIDAD_MEDIDAS = sum(CANTIDAD_MEDIDAS, na.rm = TRUE))
MED_DEPTO

}

#Estrategias del mes

#Medidas Admin

# Criterio Medidas.
# Medidas Dictadas con los criterios de metas planefa (Incluye solo medidas dictadas  inicialmente (principales) 
# independinetemente si estas posteriormente han sido variadas o reconsideradas o revocadas o declaradas nulas.
MA<- REP_CSEP %>% 
  mutate_if(is.character,function(x){toupper(x)}) %>%
  mutate(TIPO_DICTADO = case_when(
    str_detect(RESOL,"ACTA") &  str_detect(MED_DIC_EN,"CAMPO") ~ "ACTA",
    TRUE ~ "RESOL" )) %>%
  filter(FECHA_RESOL>="2022-01-01",
         FECHA_RESOL<= EVALUACION,
         str_detect(ORIGEN_RESOL,"PRINCIPAL")) %>%
  group_by(SECTOR,TIPO_DICTADO) %>% 
  summarise(MA=n_distinct(RESOL)) %>%
  spread(TIPO_DICTADO,MA) %>% replace(is.na(.),0)  %>% adorn_totals("col",name = "RD/ACTA")

MA %>% adorn_totals() %>% formattable()
#Multas coercitivas
MC<- MEDIDAS_ADM2 %>% filter(FECHA_RD_MULTA_COERCITIVA >="2022-01-01",
                             FECHA_RD_MULTA_COERCITIVA <= EVALUACION) %>%
  mutate(ANO=year(FECHA_RD_MULTA_COERCITIVA)) %>%
  group_by(#ANO,
           SECTOR) %>% summarise(MC=n_distinct(ULTIMA_RD_MULTA_COERCITIVA))

MC %>% adorn_totals() %>% formattable()
#Acta de Ejecución Forzosa
EF<- MEDIDAS_ADM2 %>% filter(FECHA_ACTA_EJECUCION_FORZOSA >="2022-01-01",
                             FECHA_ACTA_EJECUCION_FORZOSA <= EVALUACION) %>%
  mutate(ANO=year(FECHA_ACTA_EJECUCION_FORZOSA)) %>%
  group_by(#ANO,
           SECTOR) %>% summarise(EF=n_distinct(ACTA_EJECUCION_FORZOSA))

EF %>% adorn_totals() %>% formattable()
  
AC<- ACUERDOS %>% filter(FEC_REUNION>="2022-01-01",
                         FEC_REUNION<= EVALUACION) %>%
  group_by(#ANO, 
           SECTOR) %>% summarise(AC=n_distinct(NRO_ACUERDO))


AC %>% adorn_totals() %>% formattable()
#Unimos las estrategias

ESTRATEGIAS<- MA %>% full_join(MC) %>% full_join(EF) %>% full_join(AC) %>% replace(is.na(.),0) %>% adorn_totals()
ESTRATEGIAS



#---REPORTE EXCEL-------
if(1==1){
#Creamos Libro
#-------------------------------
wb<- createWorkbook()
#Creamos EStilos y modificamos la fuente
#----------------------------------------
modifyBaseFont(wb,fontSize = 10,fontName = "calibri", fontColour = "blue") # Solamente Aplica para writeData
s1 <- createStyle(fontSize=10,fontName = "arial", fontColour = "blue", textDecoration = c("bold","underline"))
s2 <- createStyle(fontSize=10,fontName = "arial", fontColour = "black", textDecoration = c("bold","underline"))
#Crea Hojas
#----------------------
addWorksheet(wb,"BD", gridLines = FALSE)
addWorksheet(wb,"REPORTE", gridLines = FALSE)
#Tabla 1
writeData(wb,"BD",as.data.frame(paste(c("1. REPORTE CONSOLIDADO DE MEDIDAS ADMINISTRATIVAS - (SET 2016-2021) / "),c(Sys.time()), sep = "-"))
          , startRow = 1 ,startCol = "A", colNames = FALSE)
writeDataTable(wb,"BD", REP_CSEP,startRow = 2 ,startCol = "A")
#Tabla 2.1
writeData(wb,"REPORTE",as.data.frame(c("2.1 CANTIDAD DE RESOLUCIONES Y MEDIDAS DICTADAS POR SECTOR - (SET 2016-2021)"))
          , startRow = 2 ,startCol = "A", colNames = FALSE)
writeDataTable(wb,"REPORTE", RESOL_MED,startRow = 3 ,startCol = "A")
#Tabla 2.2
writeData(wb,"REPORTE",as.data.frame(c("2.2 CANTIDAD DE MEDIDAS FIRMES POR SECTOR - (SET 2016-2021)")),
          startRow = 3+dim(RESOL_MED)[1]+3 ,
          startCol = "A", colNames = FALSE)
writeDataTable(wb,"REPORTE", MED_FIRM,
               startRow = 3+dim(RESOL_MED)[1]+4 ,
               startCol = "A")
#Tabla 2.3
writeData(wb,"REPORTE",as.data.frame(c("2.3 CANTIDAD DE MEDIDAS SEGUN TIPO DE MEDIDA POR SECTOR - (SET 2016-2021)")),
          startRow = 3+dim(RESOL_MED)[1]+4+dim(MED_FIRM)[1]+3 ,
          startCol = "A", colNames = FALSE)
writeDataTable(wb,"REPORTE", MED_TIPO,
               startRow = 3+dim(RESOL_MED)[1]+4+dim(MED_FIRM)[1]+4 ,
               startCol = "A")
#Tabla 2.4
writeData(wb,"REPORTE",as.data.frame(c("2.4 REPORTE CONSOLIDADO DEL ESTADO DE LAS MEDIDAS DICTADAS - (SET 2016-2021)")),
          startRow = 3+dim(RESOL_MED)[1]+4+dim(MED_FIRM)[1]+4 + dim(MED_TIPO)[1]+3 ,
          startCol = "A", colNames = FALSE)
writeDataTable(wb,"REPORTE", MED_CUMPL,
               startRow = 3+dim(RESOL_MED)[1]+4+dim(MED_FIRM)[1]+4 + dim(MED_TIPO)[1]+4 ,
               startCol = "A")
#Tabla 2.7
writeData(wb,"REPORTE",as.data.frame(c("2.7 DIAS CALENDARIOS TRANSCURRIDOS ENTRE FIN DE SUPERVISION Y DICTADO DE LA RESOLUCION - (SET 2016-2021)")),
          startRow = 3+dim(RESOL_MED)[1]+4+dim(MED_FIRM)[1]+4 + dim(MED_TIPO)[1]+4 + dim(MED_CUMPL)[1]+3 ,
          startCol = "A", colNames = FALSE)
writeDataTable(wb,"REPORTE", MED_DIAS,
               startRow = 3+dim(RESOL_MED)[1]+4+dim(MED_FIRM)[1]+4 + dim(MED_TIPO)[1]+4 + dim(MED_CUMPL)[1]+4 ,
               startCol = "A")
#Tabla 2.8
writeData(wb,"REPORTE",as.data.frame(c("2.8 CANTIDAD DE RESOLUCIONES Y MEDIDAS DICTADAS POR REGION - (SET 2016-2021)")),
          startRow = 3+dim(RESOL_MED)[1]+4+dim(MED_FIRM)[1]+4 + dim(MED_TIPO)[1]+4 + dim(MED_CUMPL)[1]+4 + dim(MED_DIAS)[1]+3 ,
          startCol = "A", colNames = FALSE)
writeDataTable(wb,"REPORTE", MED_DEPTO,
               startRow = 3+dim(RESOL_MED)[1]+4+dim(MED_FIRM)[1]+4 + dim(MED_TIPO)[1]+4 + dim(MED_CUMPL)[1]+4 + dim(MED_DIAS)[1]+4 ,
               startCol = "A")
#Tabla 2.8
writeData(wb,"REPORTE",as.data.frame(c("2.9 ESTRATEGIAS DE PROMOCIÓN DE CUMPLIMIENTO - (ENE 2020 - 2021)")),
          startRow = 3+dim(RESOL_MED)[1]+4+dim(MED_FIRM)[1]+4 + dim(MED_TIPO)[1]+4 + dim(MED_CUMPL)[1]+4 + dim(MED_DIAS)[1]+4 + dim(MED_DEPTO)[1]+3 ,
          startCol = "A", colNames = FALSE)
writeDataTable(wb,"REPORTE", ESTRATEGIAS,
               startRow = 3+dim(RESOL_MED)[1]+4+dim(MED_FIRM)[1]+4 + dim(MED_TIPO)[1]+4 + dim(MED_CUMPL)[1]+4 + dim(MED_DIAS)[1]+4 + dim(MED_DEPTO)[1]+4 ,
               startCol = "A")

#Estilo Tabla 1
addStyle(wb,sheet = "BD",style = s1 ,rows = c(1),cols = c(1:10))
#Guarda
#---------
saveWorkbook(wb, paste0(ruta_BolOutput,"Medidas_Enero_2022.xlsx"), overwrite = TRUE)

}


