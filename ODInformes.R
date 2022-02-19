#Instalar paquete (Select : NONE)
# devtools::install_github("r-lib/gargle")
# library(gargle)
#install.packages("googlesheets4")
#abrimos las librerias
library(easypackages)
pqt<- c("googlesheets4","gargle", "tidyverse","lubridate", "janitor","formattable", "openxlsx","readxl", "stringr")
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

#BUscar archivos con informacion 
archivos<- gs4_find("POI 2022", n_max=500) %>% filter(!str_detect(name,"@|Copia|MODELO")) %>% arrange(name)

#Confirmas permisos
1

#Creamos una LIsta
file.list1<- list("")
file.list2<- list("")


#Importar datos
for(i in 1:14){
  file.list1[[i]]<- as.data.frame(read_sheet(ss=archivos$id[i],
                                             sheet = "INAF_1",
                                             range = "A2:X",
                                             #skip = 2,
                                             col_types = "icccccccccccccccicccDDcc"))
                                             #col_types = "icccccccDcDcccccccciccccDDcDccc")) para 2021
} #Hasta LORETO

for(i in 15:30){
  file.list2[[i]]<- as.data.frame(read_sheet(ss=archivos$id[i],
                                             sheet = "INAF_1",
                                             range = "A2:X",
                                             #skip = 3,
                                             col_types = "icccccccccccccccicccDDcc"))
                                             #col_types = "icccccccDcDcccccccciccccDDcDccc")) para 2021
} #Hasta LA CONVENCIO

OD1<- bind_rows(file.list1[1:14]) %>% mutate_if(is.character, function(x){str_to_upper(x)})
OD2<- bind_rows(file.list2[15:29]) %>% mutate_if(is.character, function(x){str_to_upper(x)})
#consolidamos la data
od_inf<- bind_rows(OD1,OD2)
#Finalizar conexion
gs4_deauth()
#OD para Boletin


#Guardamos una copia en las descargas

save(od_inf, file = paste0(ruta_descargas,"od_inf_enero.Rdata"))


INF_OD_BOLETIN <- od_inf %>% mutate(COMPET= if_else(str_detect(COMPETENCIA,"RESIDUOS|INTEGRALES"), "OD_RES","OD_HID")) %>%
  filter(!is.na(INFORME)) %>% mutate(INFORME = trimws(INFORME,which = c("both"),whitespace = "[ \t\r\n]"))

#Correcciones
INF_OD_BOLETIN<- INF_OD_BOLETIN %>%
  mutate(INFORME = if_else(str_detect(INFORME,"0006-2022-OEFA/ODES-APU"),"00006-2022-OEFA/ODES-APU",INFORME))


#Meses Evaluados
#EVALUACION<- "ENERO|FEBRERO"#|MARZO|ABRIL|MAYO|JUNIO|JULIO|AGOSTO|SETIEMBRE|SEPTIEMBRE|OCTUBRE|NOVIEMBRE|DICIEMBRE"
EVALUACION<- as.Date("2022-01-31")



#-------ESTADO INFORMES INAPS--------
if(2==2){
  RECOM_INAPS<- Informes_R  %>%
    select(TXCOORDINACION,TXCUC,EXPEDIENTE=TXNUMEXP, FEINICIO, FEFIN,
           IDUF_SIG, 
           TIPOACCION_INAPS=TXACCION,
           TIPOSUP_INAPS= TXTIPSUP,
           INFORME=TXINFORME,
           RECOM_INAPS=TXRECOMENDACION,
           FECHA_REAL= FEINFORME, 
           IDADMIN= IDADMINISTRADO,
           ADMINISTRADO_INAPS=TXADMINISTRADO, 
           UF_INAPS=TXSUBUNIDAD,
           FEINI_ELAB_INF, FE_DERIV_DOC_DERIVACION,
           TX_DOCUMENTO_PREVIO, TX_NUMERO_DOCUMENTO_PREVIO,
           FE_DOCUMENTO_PREVIO, FE_REGISTRO_DOCUMENTO_PREVIO,
           TX_OTRO_DOCUMENTO_PREVIO)%>%
    mutate(YEAR_APROB= year(FECHA_REAL),
           INFORME= str_replace(INFORME,"INF Nº|INFORME N°"," "),
           INFORME = trimws(INFORME,which = c("both"),whitespace = "[ \t\r\n]")) %>%
    mutate(MES_META_DERIV = toupper(as.character(month(as.Date(FE_DERIV_DOC_DERIVACION),label = TRUE, abbr = FALSE )))) %>% #creamos el mes meta con la derivacion
    mutate(MES_META_DERIV = if_else(str_detect(MES_META_DERIV,"SEPTIEMBRE"),"SETIEMBRE",MES_META_DERIV)) %>% #cambiamos el mes
    filter(!is.na(RECOM_INAPS)) %>%
    distinct(TXCUC,IDADMIN,IDUF_SIG,INFORME,EXPEDIENTE,.keep_all = TRUE) %>%
    select(-RECOM_INAPS)%>%
    mutate_at(vars(starts_with("FE")), function(x){as.Date(x)}) %>%
    arrange(FEINICIO) %>%
    rename(CUC_INAPS = TXCUC)
  
  #-------CONSOLIDADO BOLETIN-------------
  INF_OD_CONSOL_BOLETIN<- INF_OD_BOLETIN %>%
    mutate(INFORME= trimws(INFORME, which = c("both")))%>%
    left_join(RECOM_INAPS, by=c("INFORME")) %>%
    mutate(MES_META = MES_META_DERIV) %>% #Desde julio se usa la fecha de derivacion como mes meta
    filter(FE_DERIV_DOC_DERIVACION <= EVALUACION) %>%
    mutate(MES_META= factor(MES_META,levels = c("ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO","JULIO","AGOSTO","SETIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE")))%>%
    mutate(CATEG= if_else(str_detect(COMPETENCIA,"RESID|INTEGRALES"),"OD_RES","OD_HID"))%>%
    mutate(OBLIG_CUMPL = if_else(is.na(OBLIG_CUMPL),as.integer(0),as.integer(OBLIG_CUMPL) )) %>%
    #CREAMOS CUMPLIMIENTOS
    mutate(
      CUMPLIM_ADD = case_when(
        str_detect(RESULT_HECHO,"ARCHIVO") & str_detect(MOTIVO_ARCHIV_CONOC,"LEVE") ~ 1,  #se quitó "CUMPLIMETO$" desde la meta de junio 2021
        str_detect(RESULT_HECHO,"ARCHIVO") & str_detect(MOTIVO_ARCHIV_CONOC,"SUPERVISION_ORIENTATIVA") & str_detect(RECOM_INFORME,"^ARCHIVO$") ~ 1,
        str_detect(RESULT_HECHO,"CONOCIMIENTO_OTROS") & str_detect(MOTIVO_ARCHIV_CONOC,"CUMPLIMIENTO$|LEVE") ~ 1,
        str_detect(RESULT_HECHO,"INFORMAR_DFAI") & str_detect(MOTIVO_ARCHIV_CONOC,"CUMPLIMIENTO_MEDIDA_DFAI") ~ 1,
        TRUE ~ 0
      ),
      #CREAMOS INCUMPLIMIENTOS
      OBL_INCUMPLIDA = case_when(
        str_detect(RESULT_HECHO,"INICIO_PAS") ~ 1,
        str_detect(RESULT_HECHO,"ALERTA|RECOMENDAR_MEDIDA|ARCHIVO ATIPICO") & str_detect(MOTIVO_ARCHIV_CONOC,"SUPERVISION_ORIENTATIVA") ~ 1,
        str_detect(RESULT_HECHO,"RECOMENDAR") & str_detect(MOTIVO_ARCHIV_CONOC,"SUPERVISION_ORIENTATIVA") & str_detect(RECOM_INFORME,"ARCHIVO_ORIENTATIVO") ~ 1,
        str_detect(RESULT_HECHO,"ARCHIVO") & str_detect(MOTIVO_ARCHIV_CONOC,"SUPERVISION_ORIENTATIVA") & str_detect(RECOM_INFORME,"ARCHIVO_ORIENTATIVO") ~ 1,
        TRUE ~ 0
      ),
      #NO CONSIDERAR SI NO TIENE UN RESULTADO CONCLUYENTE
      NO_CONSIDERAR = case_when(
        CUMPLIM_ADD==0 || OBL_INCUMPLIDA==0 ~ 1,
        TRUE ~ 0
      ),
      #SUBSANADOS
      SUBSANADO= case_when(
        str_detect(SUBSANA,"^SUBSANO") & NO_CONSIDERAR ==0 ~ 1,
        TRUE ~ 0
      ),
      SUBS_LEVE = case_when(
        SUBSANADO==1 & str_detect(CLASIF_HECHO,"LEVE") ~ 1,
        TRUE ~ 0
      ),
      SUBS_TRAS = case_when(
        SUBSANADO==1 & str_detect(CLASIF_HECHO,"TRASCENDENTE") ~ 1,
        TRUE ~ 0
      )) %>%
    #FILTRO UNICOS
    group_by(CUC_INAPS,IDADMIN,IDUF_SIG,INFORME) %>%
    mutate(SEQX=as.character(seq(paste(CUC_INAPS,IDADMIN,IDUF_SIG, sep = "")))) %>% #creamos correlativo de cantidad de registros por cuc
    as.data.frame() %>%
    mutate(FILTRO_INF_UNIQ= if_else( duplicated(paste(INFORME,SEQX, sep = ""))==FALSE,1,0)) %>%
    select(-c(SEQX,NO_CONSIDERAR,MES_META_DERIV))
  
  table(is.na(INF_OD_CONSOL_BOLETIN$IDADMIN), exclude = TRUE) #Conocer si hay vacios en los merge
}

#Agregamos el tipo de informe con la ifnormacion del expediente
# INF_OD_CONSOL_BOLETIN <- INF_OD_CONSOL_BOLETIN %>% 
#   left_join(
#     INF_OD_CONSOL_BOLETIN %>% distinct(EXPEDIENTE, CUC_INAPS, FEFIN) %>% arrange(EXPEDIENTE, FEFIN) %>% filter(!duplicated(EXPEDIENTE)) %>%
#       mutate(TIPO_INF = paste0("ACCIONES ",year(FEFIN))) %>% select(EXPEDIENTE, TIPO_INF)
#   )
  
#Agregamos el tipo de tinforme con la ifnormacion del expediente
INF_OD_CONSOL_BOLETIN <- INF_OD_CONSOL_BOLETIN %>% 
  left_join(
    INF_OD_CONSOL_BOLETIN %>% distinct(EXPEDIENTE, CUC_INAPS, FEFIN) %>% arrange(EXPEDIENTE, FEFIN) %>% filter(!duplicated(EXPEDIENTE)) %>%
      mutate(TIPO_INF = paste0("ACCIONES ",year(as.Date(FEFIN)))) %>% select(EXPEDIENTE, TIPO_INF)
  ) %>% relocate(TIPO_INF , .after = MES_META) %>%
  mutate(FE_DOCUMENTO_PREVIO = case_when(
    is.na(FE_DOCUMENTO_PREVIO) ~ FEFIN,
    TRUE ~ FE_DOCUMENTO_PREVIO))

  

#-------------------Guardamos la data--------------
save(INF_OD_CONSOL_BOLETIN, file = paste0(ruta_DatOutput,"ODInf_Febrero_2022.RData"))
load(paste0(ruta_DatOutput,"ODInf_Febrero_2022.RData"))

if(3==3){
  #-------CANTIDAD DE INFORMES
  INF_OD_TOTAL<- INF_OD_CONSOL_BOLETIN %>%
    #filter(str_detect(CUENTA,"META"))%>%
    filter(FE_DERIV_DOC_DERIVACION <= EVALUACION) %>%
    distinct(INFORME, .keep_all = TRUE) %>%
    group_by(MES_META,OD,
             TIPO_INF,CATEG) %>%
    summarise(TOTAL= n_distinct(INFORME) ) %>%
    spread(CATEG, TOTAL) %>%
    replace(.,is.na(.),0) %>%
    arrange(MES_META,
            desc(TIPO_INF)) %>%
    adorn_totals("row") %>% adorn_totals("col") %>% formattable()
  
  INF_OD_TOTAL
  
  #-------RECOMENDACION DE LOS INFORMES
  INF_OD_RECOM<-  INF_OD_CONSOL_BOLETIN %>%
    #filter(str_detect(CUENTA,"META"))%>%
    filter(FE_DERIV_DOC_DERIVACION <= EVALUACION) %>%
    mutate(RECOM_INFORME= if_else(str_detect(RECOM_INFORME,"ARCHIVO"), "ARCHIVO", RECOM_INFORME)) %>%
    distinct(TIPO_INF,MES_META,OD,CATEG,INFORME, RECOM_INFORME) %>%
    group_by(MES_META,OD,CATEG,RECOM_INFORME) %>%
    summarise(TOTAL = n_distinct(INFORME)) %>%
    spread(CATEG, TOTAL)%>% replace(.,is.na(.),0) %>%
    adorn_totals("row") %>% adorn_totals("col") %>%
    formattable()
  
  formattable(INF_OD_RECOM) 
  
  #----------------
  INF_OD_TIPOSUP<- INF_OD_CONSOL_BOLETIN %>%
    distinct(INFORME, CUC_INAPS , .keep_all = TRUE) %>%
    #mutate(ANO_EXP = parse_number(str_extract(EXPEDIENTE ,"(?<=-)[0-9]+"))) %>% #precedido de - 
    group_by(INFORME) %>%
    filter(FEFIN== min(as.Date(FEFIN))) %>%
    distinct(INFORME, MES_META,FEFIN, .keep_all = TRUE) %>%
    group_by(MES_META,CATEG,OD ,
             TIPOSUP_INAPS) %>%
    summarise(INFORMES = n_distinct(INFORME)) %>%
    spread(CATEG, INFORMES) %>%
    replace(is.na(.),0) %>%
    adorn_totals(c("row","col"))
  
  formattable(INF_OD_TIPOSUP)

  #----------------
  INF_OD_ADMIN<-  INF_OD_CONSOL_BOLETIN %>%
    mutate(IDADMIN = if_else(str_detect(ADMINISTRADO_INAPS,"NO DETERMINADO"),NA_character_,IDADMIN)) %>%
    filter(FE_DERIV_DOC_DERIVACION <= EVALUACION) %>%
    distinct(CATEG,OD ,IDADMIN,IDUF_SIG, .keep_all = TRUE) %>%
    group_by(#MES_META,
      CATEG) %>% 
    summarise(ADMINISTRADO = n_distinct(IDADMIN, na.rm = TRUE),
              UF=n_distinct(IDUF_SIG, na.rm = TRUE)) %>%
    gather(key = "ITEM", value = "VALORES", ADMINISTRADO:UF) %>%
    spread(CATEG, VALORES)%>%
    adorn_totals("col")
  
  formattable(INF_OD_ADMIN) 
  
  #----------------
  INF_OD_INCUMPL<-  INF_OD_CONSOL_BOLETIN %>%
    #filter(str_detect(CUENTA,"META"))%>%
    filter(FILTRO_INF_UNIQ==1) %>%
    group_by(CATEG, INFORME, MES_META) %>% 
    summarise(
      OBL_INCUMPLIDA= sum(OBL_INCUMPLIDA, na.rm = TRUE),
      CUMPLIM_ADD = sum(CUMPLIM_ADD, na.rm = TRUE),
      OBLIG_CUMPL = max(OBLIG_CUMPL, na.rm = TRUE),
      SUBS = sum(SUBSANADO, na.rm = TRUE),
      SUBS_LEVE= sum(SUBS_LEVE, na.rm = TRUE),
      SUBS_TRAS= sum(SUBS_TRAS, na.rm = TRUE)) %>%
    mutate(OBL_CUMPLIDAS= CUMPLIM_ADD + OBLIG_CUMPL,
           OBL_TOTAL= OBL_INCUMPLIDA + OBL_CUMPLIDAS) %>%
    group_by(MES_META, #TIPO_INF,
             CATEG) %>% 
    summarise(
      INFORMES = n(),
      OBL_INC= sum(OBL_INCUMPLIDA, na.rm = TRUE),
      OBL_CUM= sum(OBL_CUMPLIDAS, na.rm = TRUE),
      OBL_TOTAL = sum(OBL_TOTAL, na.rm = TRUE),
      SUBSANADO= sum(SUBS, na.rm = TRUE),
      SUBS_LEVE= sum(SUBS_LEVE, na.rm = TRUE),
      SUBS_TRAS= sum(SUBS_TRAS, na.rm = TRUE)) %>%
    arrange(MES_META, CATEG) %>%
    adorn_totals("row") %>%
    formattable()
  
  INF_OD_INCUMPL
}

#Reporte Excel
#library(xlsx)

if(4==4){
  
  #Creamos Libro
  #-------------------------------
  wb<- createWorkbook()
  #Creamos EStilos y modificamos la fuente
  #----------------------------------------
  modifyBaseFont(wb,fontSize = 10,fontName = "calibri", fontColour = "black") # Solamente Aplica para writeData
  s1 <- createStyle(fontSize=10,fontName = "arial", fontColour = "black", textDecoration = c("bold","underline"))
  #s2 <- createStyle(fontSize=10,fontName = "arial", fontColour = "black", textDecoration = c("bold","underline"))
  #Crea Hojas
  #----------------------
  addWorksheet(wb,"OD_INAF_1", gridLines = FALSE)
  addWorksheet(wb,"REPORTE", gridLines = FALSE)
  #Tabla 1
  writeData(wb,"OD_INAF_1",as.data.frame(c("1. CONSOLIDADO DE INFORMES OFICINAS DESCONCENTRADAS"))
            , startRow = 1 ,startCol = "A", colNames = FALSE)
  writeDataTable(wb,"OD_INAF_1", INF_OD_CONSOL_BOLETIN,startRow = 3 ,startCol = "A")
  #Tabla 1.1
  writeData(wb,"REPORTE",as.data.frame(c("1.1 CANTIDAD DE INFORMES CONCLUIDOS POR OD"))
            , startRow = 3 ,startCol = "A", colNames = FALSE)
  writeDataTable(wb,"REPORTE", INF_OD_TOTAL,startRow = 4 ,startCol = "A")
  # #Tabla 1.2
  writeData(wb,"REPORTE",as.data.frame(c("1.2 CANTIDAD DE INFORMES SEGUN RECOMENDACIÓN")),
            startRow = 4+dim(INF_OD_TOTAL)[1]+3 ,
            startCol = "A", colNames = FALSE)
  writeDataTable(wb,"REPORTE", INF_OD_RECOM,
                 startRow = 4+dim(INF_OD_TOTAL)[1]+4 ,
                 startCol = "A") #6 filas
  # #Tabla 1.3
  writeData(wb,"REPORTE",as.data.frame(c("1.3 CANTIDAD INFORMES POR TIPO DE SUPERVISIÓN")),
            startRow = 4+dim(INF_OD_TOTAL)[1]+4 + dim(INF_OD_RECOM)[1]+3 ,
            startCol = "A", colNames = FALSE)
  writeDataTable(wb,"REPORTE",INF_OD_TIPOSUP,
                 startRow = 4+dim(INF_OD_TOTAL)[1]+4 + dim(INF_OD_RECOM)[1]+4 ,
                 startCol = "A") #6 filas
  # #Tabla 1.4
  writeData(wb,"REPORTE",as.data.frame(c("1.4 CANTIDAD DE ADMINISTRADOS Y UNIDADES FISCALIZABLES (ACUMULADO) ")),
            startRow = 4+dim(INF_OD_TOTAL)[1]+4 + dim(INF_OD_RECOM)[1]+4 + dim(INF_OD_TIPOSUP)[1]+3 ,
            startCol = "A", colNames = FALSE)
  writeDataTable(wb,"REPORTE", INF_OD_ADMIN,
                 startRow = 4+dim(INF_OD_TOTAL)[1]+4 + dim(INF_OD_RECOM)[1]+4 + dim(INF_OD_TIPOSUP)[1]+3 ,
                 startCol = "A") #27 filas
  # #Tabla 1.5
  writeData(wb,"REPORTE",as.data.frame(c("1.5 CUMPLIMIENTOS E INCUMPLIMIENTOS ACUMULADOS POR SECTOR ")),
            startRow = 4+dim(INF_OD_TOTAL)[1]+4 + dim(INF_OD_RECOM)[1]+4 + dim(INF_OD_TIPOSUP)[1]+3 + dim(INF_OD_ADMIN)[1]+3 ,
            startCol = "A", colNames = FALSE)
  writeDataTable(wb,"REPORTE", INF_OD_INCUMPL,
                 startRow = 4+dim(INF_OD_TOTAL)[1]+4 + dim(INF_OD_RECOM)[1]+4 + dim(INF_OD_TIPOSUP)[1]+3 + dim(INF_OD_ADMIN)[1]+4 ,
                 startCol = "A") #21 filas
  
  
  #Guarda
  #---------
  saveWorkbook(wb, paste0(ruta_BolOutput,"ODInformes_Enero_2022.xlsx"), overwrite = TRUE)
}




