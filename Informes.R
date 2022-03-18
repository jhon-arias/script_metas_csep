#Este Script se trabajó para el trabajo remoto
#Pegar en esta parte el pedazo del script del equipo (OEFA)
options(java.parameters = "-Xmx1024m") # Ampliar espacio de procesamiento
#------------------------------------------------------------------------------------------------
#Paquetes===============================
library(easypackages)
pqt<- c("tidyverse","readxl","stringr","openxlsx","lubridate","bizdays" ,"janitor", "knitr","formattable")
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
#Carga BD Excel=============================
if(2==2){
  #---CARGAMOS RESULTADOS DE SUPERVISION------
  #DF_RES<- read_excel("Insumos/INFORMES/Informes_Diciembre.xlsx", skip = 3) #Oficina
  DF_RES<- read_excel(paste0(ruta_DataInput,"Informes/2022/Informes_Febrero_2022.xlsx"), skip = 3,
                      col_types = c( "numeric","text","date","text","date","text","text","date","text","date",
                                     "text","text","text","text","text","text","text","text","text","text",
                                     "numeric","numeric","numeric","numeric","numeric","numeric","text","text","text","text",
                                     "text","numeric","numeric","numeric","text","text","text","text","date","text",
                                     "text","date","date","text","date","text","date","text","text","text",
                                     "text","date","text","text","text","text")) 
  
  #Limpieza
  DF_RES<- DF_RES %>%
    filter(!is.na(INFORME)) %>%
    mutate_if(is.character,toupper) %>%
    mutate(INFORME = trimws(INFORME, which = c("both"), whitespace = "[/t/r/n]")) %>%
    mutate_at(vars(starts_with("FECHA")),as.Date) %>%
    mutate(SECTOR= case_when(
      str_detect(INFORME,"AGR") ~ "AGR",
      str_detect(INFORME,"IND") ~ "IND",
      str_detect(INFORME,"PES") ~ "PES",
      str_detect(INFORME,"MIN") ~ "MIN",
      str_detect(INFORME,"HID") ~ "HID",
      str_detect(INFORME,"ELE") ~ "ELE",
      str_detect(INFORME,"CAM") ~ "CAM",
      str_detect(INFORME,"RES") ~ "RES"
    )) %>%
    mutate(SECTOR = factor(SECTOR, levels=c("MIN","HID","ELE","IND","PES","AGR","CAM","RES"))) %>%
    mutate(DIREC= case_when(
      str_detect(SECTOR,"AGR|IND|PES") ~ "DSAP",
      str_detect(SECTOR,"MIN|HID|ELE") ~ "DSEM",
      TRUE ~ "DSIS"
    ))
    
}

# Correcciones


#Meses Evaluados ========
#EVALUACION<- "ENERO"#|FEBRERO|MARZO|ABRIL|MAYO|JUNIO|JULIO|AGOSTO|SETIEMBRE|SEPTIEMBRE|OCTUBRE|NOVIEMBRE|DICIEMBRE"
EVALUACION<- as.Date("2022-02-28")
#Estado de Informes INAPS==========================
if(3==3){ 
  RECOM_INAPS<- Informes_R%>%
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
    filter(!is.na(RECOM_INAPS)) %>%
    arrange(desc(FE_DERIV_DOC_DERIVACION)) %>%
    distinct(TXCUC,IDADMIN,IDUF_SIG,INFORME,EXPEDIENTE,.keep_all = TRUE) %>%
    select(-RECOM_INAPS)%>%
    mutate(MES_META_DERIV = toupper(as.character(month(as.Date(FE_DERIV_DOC_DERIVACION),label = TRUE, abbr = FALSE )))) %>% #creamos el mes meta con la derivacion
    mutate(MES_META_DERIV = if_else(str_detect(MES_META_DERIV,"SEPTIEMBRE"),"SETIEMBRE",MES_META_DERIV)) %>% #cambiamos el mes
    mutate(YEAR_APROB= year(FECHA_REAL)) %>%
    mutate_at(vars(starts_with("FE")), function(x){as.Date(x)}) %>%
    mutate(INFORME = trimws(INFORME,which = c("both"), whitespace = "[/t/r/n]")) %>%
    arrange(FEINICIO)%>%
    rename(CUC_INAPS = TXCUC)
  
  
  #-------CONSOLIDADO BOLETIN-------------
  INF_CONSOL_BOLETIN<- DF_RES %>%
    left_join(RECOM_INAPS) %>%
    filter(!is.na(RECOM_INFORME)) %>%
    mutate(MES_META = MES_META_DERIV) %>% #Desde julio se usa la fecha de derivacion como mes meta
    filter(FE_DERIV_DOC_DERIVACION<= EVALUACION) %>%  #=============CAMBIAR SIEMRPRE
    mutate(MES_META= factor(MES_META,levels = c("ENERO","FEBRERO","MARZO","ABRIL","MAYO","JUNIO","JULIO","AGOSTO","SETIEMBRE","OCTUBRE","NOVIEMBRE","DICIEMBRE")))%>%
    mutate_at(vars(starts_with("FE")), function(x){as.Date(x)}) %>%
    #REEMPLAZAR LOS NA DE LOS CUMPLIMIENTOS
    mutate(OBLIG_CUMPL = if_else(is.na(OBLIG_CUMPL),0,OBLIG_CUMPL)) %>%
    #CREAMOS CUMPLIMIENTOS
    mutate(
      CUMPLIM_ADD = case_when(
        str_detect(RESULT_HECHO,"ARCHIVO") & str_detect(MOTIVO_ARCHIV_CONOC,"LEVE SUBS") ~ 1, #se qutó "CUMPLIMETO$" desde la meta de junio 2021
        str_detect(RESULT_HECHO,"ARCHIVO") & str_detect(MOTIVO_ARCHIV_CONOC,"SUPERVISIÓN_ORIENTATIVA") & str_detect(RECOM_INFORME,"^ARCHIVO$") ~ 1,
        str_detect(RESULT_HECHO,"CONOCIMIENTO_OTROS") & str_detect(MOTIVO_ARCHIV_CONOC,"CUMPLIMIENTO$|LEVE SUBS") ~ 1,
        str_detect(RESULT_HECHO,"INFORMAR_DFAI") & str_detect(MOTIVO_ARCHIV_CONOC,"^CUMPLIMIENTO_DE_MEDIDA_DFAI") ~ 1,
        TRUE ~ 0
      ),
      #CREAMOS INCUMPLIMIENTOS
      OBL_INCUMPLIDA = case_when(
        str_detect(RESULT_HECHO,"INICIO_PAS") ~ 1,
        str_detect(RESULT_HECHO,"ALERTA|RECOMENDAR_MEDIDA|ARCHIVO_ATÍPICO") & str_detect(MOTIVO_ARCHIV_CONOC,"SUPERVISIÓN_ORIENTATIVA") ~ 1,
        str_detect(RESULT_HECHO,"RECOMENDAR") & str_detect(MOTIVO_ARCHIV_CONOC,"SUPERVISIÓN_ORIENTATIVA") & str_detect(RECOM_INFORME,"ARCHIVO_ORIENTATIVO") ~ 1,
        str_detect(RESULT_HECHO,"ARCHIVO") & str_detect(MOTIVO_ARCHIV_CONOC,"SUPERVISIÓN_ORIENTATIVA") & str_detect(RECOM_INFORME,"ARCHIVO_ORIENTATIVO") ~ 1,
        str_detect(RESULT_HECHO,"INFORMAR_DFAI") & str_detect(MOTIVO_ARCHIV_CONOC,"INCUMPLIMIENTO_DE_MEDIDA_DFAI") ~ 1,
        TRUE ~ 0
      ),
      #SUBSANADOS
      SUBS_LEVE = case_when(
        str_detect(SUBSANA,"^SUBSANÓ") & str_detect(CLASIF_HECHO,"LEVE") & CUMPLIM_ADD==1 ~ 1,
        TRUE ~ 0
      ),
      SUBS_TRAS = case_when(
        str_detect(SUBSANA,"^SUBSANÓ") & str_detect(CLASIF_HECHO,"TRASCENDENTE") & OBL_INCUMPLIDA==1  ~ 1,
        TRUE ~ 0
      ),
      SUBSANADO = SUBS_LEVE + SUBS_TRAS) %>%
    #FILTRO UNICOS
    group_by(CUC_INAPS,IDADMIN,IDUF_SIG,INFORME) %>%
    mutate(SEQX=as.character(seq(paste(CUC_INAPS,IDADMIN,IDUF_SIG, sep = "")))) %>% #creamos correlativo de cantidad de registros por cuc/admin/uf
    as.data.frame() %>%
    mutate(FILTRO_INF_UNIQ= if_else(duplicated(paste(INFORME,SEQX, sep = ""))==FALSE,1,0)) %>%
    select(-c(SEQX, MES_META_DERIV))
    
  table(is.na(INF_CONSOL_BOLETIN$IDADMIN), exclude = TRUE) #Que no haya TRUE
  
}        

#Agregamos el tipo de tinforme con la ifnormacion del expediente
  INF_CONSOL_BOLETIN <- INF_CONSOL_BOLETIN %>% select(-TIPO_INF) %>%
  left_join(
    INF_CONSOL_BOLETIN %>% distinct(EXPEDIENTE, CUC_INAPS, FEFIN) %>% arrange(EXPEDIENTE, FEFIN) %>% filter(!duplicated(EXPEDIENTE)) %>%
      mutate(TIPO_INF = paste0("ACCIONES ",year(as.Date(FEFIN)))) %>% select(EXPEDIENTE, TIPO_INF)
  ) %>% relocate(TIPO_INF , .after = MES_META) %>%
  mutate(FE_DOCUMENTO_PREVIO = case_when(
    is.na(FE_DOCUMENTO_PREVIO) ~ FEFIN,
    TRUE ~ FE_DOCUMENTO_PREVIO))


#Guardamos Excel para Boletin==========================
save(INF_CONSOL_BOLETIN, file = paste0(ruta_DatOutput,"BDInf_Febrero_2022.RData"))
load(paste0(ruta_DatOutput,"BDInf_Febrero_2022.RData"))


#Informes por mes====================================
if(4==4){
  INF_TOTAL<- INF_CONSOL_BOLETIN %>%
    mutate(SECTOR = factor(SECTOR,levels = c("MIN","HID","ELE","IND","PES","AGR","CAM","RES")))%>%
    #filter(str_detect(CUENTA,"CONCLUIDO"))%>%
    #mutate(TIPO_INF = if_else(str_detect(TIPO_INF,"PAS|TR|2019|2018"),"ACCIONES 2020",TIPO_INF)) %>%
    filter(FE_DERIV_DOC_DERIVACION<= EVALUACION) %>%
    distinct(INFORME,SECTOR, .keep_all = TRUE) %>%
    group_by(MES_META,
      TIPO_INF,
      SECTOR) %>%
    summarise(TOTAL= n_distinct(INFORME)) %>%
    spread(SECTOR, TOTAL) %>%
    replace(is.na(.),0) %>%
    arrange(MES_META,TIPO_INF
    ) %>%
    adorn_totals(c("row","col")) 
  
  INF_TOTAL %>% formattable()
  
  #Informes por recomendacion============================
  INF_RECOM<-  INF_CONSOL_BOLETIN %>%
    mutate(SECTOR = factor(SECTOR,levels = c("MIN","HID","ELE","IND","PES","AGR","CAM","RES")))%>%
    #filter(str_detect(CUENTA,"CONCLUIDO"))%>%
    filter(FE_DERIV_DOC_DERIVACION<= EVALUACION)%>%
    mutate(RECOM_INFORME= if_else(str_detect(RECOM_INFORME,"ARCHIVO"), "ARCHIVO", RECOM_INFORME)) %>%
    distinct(#TIPO_INF,
      MES_META,SECTOR,INFORME,RECOM_INFORME) %>%
    group_by(#TIPO_INF,
      MES_META, SECTOR,RECOM_INFORME) %>% 
    summarise(TOTAL = n()) %>%
    spread(SECTOR, TOTAL) %>% replace(.,is.na(.),0) %>%
    adorn_totals("row") %>% adorn_totals("col") 
  
  INF_RECOM %>% formattable()
  
  #Informes por tipo de supervisión=======================
  
  INF_TIPO<- INF_CONSOL_BOLETIN %>%
    mutate(SECTOR = factor(SECTOR,levels = c("MIN","HID","ELE","IND","PES","AGR","CAM","RES")))%>%
    filter(FE_DERIV_DOC_DERIVACION<= EVALUACION)%>%
    #mutate(ANO_EXP = parse_number(str_extract(EXPEDIENTE ,"(?<=-)[0-9]+"))) %>% #precedido de - 
    group_by(MES_META, INFORME) %>%
    filter(FEFIN== min(as.Date(FEFIN))) %>%
    distinct(INFORME, MES_META,FEFIN, .keep_all = TRUE) %>% 
    group_by(SECTOR,MES_META ,
             TIPOSUP_INAPS) %>%
    summarise(INFORMES = n_distinct(INFORME)) %>%
    spread(SECTOR, INFORMES) %>%
    replace(is.na(.),0) %>%
    arrange(MES_META, desc(TIPOSUP_INAPS)) %>%
    adorn_totals(c("row","col")) 
  
  INF_TIPO %>% formattable()
  
  
  #Administrados y Unidades Fizcalizables========================
  INF_ADMIN<-  INF_CONSOL_BOLETIN %>%
    mutate(SECTOR = factor(SECTOR,levels = c("MIN","HID","ELE","IND","PES","AGR","CAM","RES")))%>%
    filter(FE_DERIV_DOC_DERIVACION<= EVALUACION)%>%
    mutate(IDADMIN = if_else(str_detect(ADMINISTRADO_INAPS,"NO DETERMINADO"),NA_character_,IDADMIN)) %>%
    distinct(SECTOR, IDADMIN,IDUF_SIG, .keep_all = TRUE
             )%>%
    group_by(SECTOR) %>% 
    summarise(ADMINISTRADO= n_distinct(IDADMIN, na.rm = TRUE),
              UF= n_distinct(IDUF_SIG, na.rm = TRUE)) %>%
    gather(key = "ITEM", value = "VALORES", ADMINISTRADO:UF) %>%
    spread(SECTOR, VALORES) %>%
    adorn_totals("col") 
  
  INF_ADMIN %>% formattable()
  
  #Incumpleimientos por informe===================
  INF_INCUMPL<-  INF_CONSOL_BOLETIN %>%
    mutate(SECTOR = factor(SECTOR,levels = c("MIN","HID","ELE","IND","PES","AGR","CAM","RES")))%>%
    #filter(str_detect(CUENTA,"CONCLUIDO"))%>%
    filter(FE_DERIV_DOC_DERIVACION<= EVALUACION) %>%
    filter(FILTRO_INF_UNIQ==1) %>%
    group_by(SECTOR, INFORME, MES_META) %>% 
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
             SECTOR) %>% 
    summarise(
      INFORMES = n_distinct(INFORME),
      OBL_INC= sum(OBL_INCUMPLIDA, na.rm = TRUE),
      OBL_CUM= sum(OBL_CUMPLIDAS, na.rm = TRUE),
      OBL_TOTAL = sum(OBL_TOTAL, na.rm = TRUE),
      SUBSANADO= sum(SUBS, na.rm = TRUE),
      SUBS_LEVE= sum(SUBS_LEVE, na.rm = TRUE),
      SUBS_TRAS= sum(SUBS_TRAS, na.rm = TRUE)) %>%
    adorn_totals("row") 
  
  INF_INCUMPL
  
  #Dias calendarios elaboracion informe=======================
  source("C:/Users/jach_/OneDrive/Documentos/ShinyApp/Planefa2021/Feriados.R")  #CARGAMOS LOS FERIADOS 2021
  
  create.calendar(name="Peru2021", weekdays=c("sunday", "saturday"),
                  adjust.from=adjust.next, adjust.to=adjust.previous, financial=F,
                  holidays = feriados,
                  start.date = as.Date("1970-01-01"), end.date = as.Date("2071-01-01"))
  bizdays.options$set(default.calendar="Peru2021")

  
  INF_DIASHABIL<- INF_CONSOL_BOLETIN %>%
    filter(FE_DERIV_DOC_DERIVACION<= EVALUACION)%>%
    filter(FILTRO_INF_UNIQ==1) %>%
    group_by(INFORME) %>%
    mutate(FECHA_MAX_ACC = max(as.Date(FEFIN)),
           FECHA_INI= if_else(!is.na(FE_DOCUMENTO_PREVIO),FE_DOCUMENTO_PREVIO,FECHA_MAX_ACC),
           DIAS = bizdays(FECHA_INI,FECHA_INFORME, cal = "Peru2021")+1) %>%
    group_by(DIREC, SECTOR, MES_META) %>%
    summarise(MINIMO=min(DIAS, na.rm = TRUE),
              PROMEDIO=round(mean(DIAS,na.rm = TRUE),0),
              MAXIMO=max(DIAS, na.rm = TRUE)) %>%
    mutate(SECTOR= factor(as.factor(SECTOR),levels=c("MIN","HID","ELE","IND","PES","AGR","CAM","RES"))) %>%
    arrange(SECTOR)
  
  INF_DIASHABIL 
}


#Exportamos la informacion a un excel
if(5==5){

#lista de BD  
df.list <- list(
  c("1.1 CANTIDAD DE INFORMES POR SECTOR"), as.data.frame(INF_TOTAL),
  c("1.2 CANTIDAD DE INFORMES SEGUN RECOMENDACIÓN"), as.data.frame(INF_RECOM),
  c("1.3 CANTIDAD INFORMES POR TIPO DE SUPERVISIÓN"), as.data.frame(INF_TIPO),
  c("1.4 CANTIDAD DE ADMINISTRADOS Y UNIDADES FISCALIZABLES (ACUMULADO)"), as.data.frame(INF_ADMIN),
  c("1.5 CUMPLIMIENTOS E INCUMPLIMIENTOS ACUMULADOS POR SECTOR "),as.data.frame(INF_INCUMPL),
  c("1.6 DIAS HÁBILES PROMEDIO PARA LA ELABORACIÓN DE UN INFORME "),as.data.frame(INF_DIASHABIL)
)

df.row<- list() #Creamos una lista para los indices
for(i in 1:length(df.list)){
   df.row[[i]]<- nrow(as.data.frame(df.list[[i]]))
}
#Convertimos la lista en un dataframe
df.row<- as.data.frame(matrix(unlist(df.row)))
#Con lalista de indices creamos una columna de posicion donde iniciar
df.row<- df.row %>% rename("Filas"=V1)
for(i in 1:nrow(df.row)){
  df.row$Posi[[i]] <- ifelse(i==1,df.row$Filas[[1]],
                             (df.row$Filas[[i-1]]+ df.row$Posi[[i-1]]+1))
}


library(openxlsx)
#-- Creamos libro --
wb1<- createWorkbook()
#-- Creamor hojas --
addWorksheet(wb1,"BD")
addWorksheet(wb1,"REPORTE")
#-- Agrgamos general --
writeData(wb1,"BD","1. CONSOLIDADO DE INFORMES",startRow = 1 ,startCol = "A")
#writeData(wb1,"BD",INF_CONSOL_BOLETIN, startRow = 3 ,startCol = "A", na.string = "") #se debe revisar
#-- Agregamos tablas
#Tabla 01
writeData(wb1,"REPORTE",df.list[[1]],startRow = df.row$Posi[[1]] ,startCol = "A")
writeDataTable(wb1,"REPORTE",df.list[[2]], startRow = df.row$Posi[[2]] ,startCol = "A")
#Tabla 02
writeData(wb1,"REPORTE",df.list[[3]],startRow = df.row$Posi[[3]] ,startCol = "A")
writeDataTable(wb1,"REPORTE",df.list[[4]], startRow = df.row$Posi[[4]] ,startCol = "A")
#Tabla 03
writeData(wb1,"REPORTE",df.list[[5]],startRow = df.row$Posi[[5]] ,startCol = "A")
writeDataTable(wb1,"REPORTE",df.list[[6]], startRow = df.row$Posi[[6]] ,startCol = "A")
#Tabla 04
writeData(wb1,"REPORTE",df.list[[7]],startRow = df.row$Posi[[7]] ,startCol = "A")
writeDataTable(wb1,"REPORTE",df.list[[8]], startRow = df.row$Posi[[8]] ,startCol = "A")
#Tabla 05
writeData(wb1,"REPORTE",df.list[[9]],startRow = df.row$Posi[[9]] ,startCol = "A")
writeDataTable(wb1,"REPORTE",df.list[[10]], startRow = df.row$Posi[[10]] ,startCol = "A")
#Tabla 06
#writeData(wb1,"REPORTE",df.list[[11]],startRow = df.row$Posi[[11]] ,startCol = "A")
#writeDataTable(wb1,"REPORTE",df.list[[12]], startRow = df.row$Posi[[12]] ,startCol = "A")

#-- Guardamos --
saveWorkbook(wb1,paste0(ruta_BolOutput,"Informes_Febrero_2022.xlsx"), overwrite = TRUE)

}

write.csv(INF_CONSOL_BOLETIN,file = paste0(ruta_BolOutput,"Informes_Febrero_2022.csv"), row.names = TRUE, na="")

