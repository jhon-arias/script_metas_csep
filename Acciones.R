#Este Script se trabajó para el trabajo remoto
#Pegar en esta parte el pedazo del script del equipo (OEFA)
#options(java.parameters = "-Xmx1024m") # Ampliar espacio de procesamiento
#------------------------------------------------------------------------------------------------
library(easypackages)
pqt<- c("tidyverse","readxl", "openxlsx","lubridate","bizdays" ,"janitor", "knitr","formattable")
libraries(pqt)
#Guardamos la data===========
setwd("C:/Users/jach_/OneDrive/Documentos/ShinyApp/Planefa2021/Diciembre2021/") 
#save(Informes_R,file = "Inf_Enero_2020.RData")
#Cargamos la data de acciones y docuemntos
load("Doc_Diciembre_2021.RData")
load("Acc_Diciembre_2021.RData")
load("Inf_Diciembre_2021.RData")
# Revisión de carga de documentos.
CargaDocs_R2 <- CargaDocs_R %>%
  #filter(FEFIN>=as.Date("2018-01-01"))%>%
  distinct(TXCUC,TXTIPO_DOC) %>%  # Quitamos Duplicadoa
  group_by(TXCUC,TXTIPO_DOC) %>%  summarise(DOCS=n())%>%  spread(TXTIPO_DOC,DOCS) %>%
  replace(is.na(.),0) %>%
  mutate(ACC_DOC= `Acta de Supervisión`+ `Documento de Registro de Informacion`+ Acta) %>%
  mutate(ACC_DOC= if_else(ACC_DOC>=1,1,0)) %>%
  select(TXCUC,ACC_DOC)

#ACCIONES E NFORMES DESDE 2018 ========================
Acciones_R2 <- Acciones_R %>%
  filter(FEFIN >= as.Date("2021-01-01")  , 
         FEFIN <=  as.Date("2021-12-31")) %>%
  left_join(CargaDocs_R2) %>%
  #left_join(Informes_R) %>%
  mutate(
    ACC_EJEC= case_when(
      TXESTADO=="ANÁLISIS DE RESULTADOS" ~ 1,
      TXESTADO=="CON RESULTADOS" ~ 1,
      TXESTADO=="CONCLUIDO" ~ 1,
      TXESTADO=="EJECUTADA" ~ 1,
      TXESTADO=="EN CUSTODIA" ~ 1,
      TRUE ~ 0),
    ACC_CLEAN= case_when(
      ACC_EJEC==1 & FGSUPUPD_2DOTIEMPO=="1" & TXACCION=="IN SITU" & ACC_DOC==1       ~ 1,
      ACC_EJEC==1 & FGSUPUPD_2DOTIEMPO=="1" & TXACCION=="PRESENCIAL" & ACC_DOC==1    ~ 1,
      ACC_EJEC==1 & FGSUPUPD_2DOTIEMPO=="1" & TXACCION=="EN GABINETE" & ACC_DOC==1   ~ 1,
      ACC_EJEC==1 & FGSUPUPD_2DOTIEMPO=="1" & TXACCION=="NO PRESENCIAL" & ACC_DOC==1 ~ 1,
      ACC_EJEC==1 & FGSUPUPD_2DOTIEMPO=="1" & TXACCION=="EN GABINETE" & ACC_DOC==0   ~ 1,
      ACC_EJEC==1 & FGSUPUPD_2DOTIEMPO=="1" & TXACCION=="NO PRESENCIAL" & ACC_DOC==0 ~ 1)) %>% 
  distinct(TXCUC, .keep_all = TRUE) %>%
  mutate( 
    TRIM= case_when(
      str_detect(TXMES,"ENE|FEB|MAR") ~ "I",
      str_detect(TXMES,"ABR|MAY|JUN") ~ "II",
      str_detect(TXMES,"JUL|AGO|SEP") ~ "III",
      str_detect(TXMES,"OCT|NOV|DIC") ~ "IV",
      TRUE ~ "OTRO"),
    COORD= case_when(
      str_detect(TXCOORDINACION,"MINE") ~ "MIN",
      str_detect(TXCOORDINACION,"HIDR") ~ "HID",
      str_detect(TXCOORDINACION,"ELEC") ~ "ELE",
      str_detect(TXCOORDINACION,"INDU") ~ "IND",
      str_detect(TXCOORDINACION,"PESC") ~ "PES",
      str_detect(TXCOORDINACION,"AGRI") ~ "AGR",
      str_detect(TXCOORDINACION,"CONS") ~ "CAM",
      str_detect(TXCOORDINACION,"RESI") ~ "RES",
      str_detect(TXCOORDINACION,"OFICI") & str_detect(TXSUBSECTOR_UND,"HIDRO")  ~ "OD_HID",
      str_detect(TXCOORDINACION,"OFICI") & str_detect(TXSUBSECTOR_UND,"RESID")  ~ "OD_RES"),
    DIREC= case_when(
      str_detect(COORD,"OD") ~ "CODE",
      str_detect(COORD,"MIN|HID|ELE") ~ "DSEM",
      str_detect(COORD,"IND|PES|AGR") ~ "DSAP",
      str_detect(COORD,"CAM|RES") ~ "DSIS")) %>%
  #Factorizamos los niveles de MES
  mutate(TXMES= factor(TXMES, levels = c("ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
                                         "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"))) %>%
  #Factorizamos los niveles de coordinación
  mutate(COORD= factor(COORD, 
                       levels=c("MIN","HID","ELE","IND","PES","AGR","CAM","RES","OD_HID","OD_RES"))) %>%
  mutate(TXFUENTE= case_when(
    str_detect(TXFUENTE,"OTRAS CIRCUNSTANCIAS") ~ "OTRAS CIRCUNSTANCIAS",
    str_detect(TXFUENTE,"ACCIDENTE O EMERGENCIA DE CARÁCTER AMBIENTAL") ~ "EMERGENCIA AMBIENTAL",
    str_detect(TXFUENTE,"DENUNCIAS") ~ "DENUNCIA AMBIENTAL",
    str_detect(TXFUENTE,"SOLICITUDES DE INTERVENCIÓN FORMULADAS POR ORGANISMOS PÚBLICOS") ~ "SOLICITUD DE INTERVENCIÓN FORMULADA POR UN ORGANISMO PÚBLICO",
    str_detect(TXFUENTE,"VERIFICACIÓN DE MEDIDAS ADMINISTRATIVAS") ~ "VERIFICACIÓN DEL CUMPLIMIENTO DE LAS MEDIDAS ADMINISTRATIVAS ORDENADAS POR EL OEFA",
    str_detect(TXFUENTE,"VERIFICACIÓN DE ACUERDO") ~ "VERIFICACIÓN DE ACUERDO DE CUMPLIMIENTO",
    str_detect(TXFUENTE,"TERMINACIÓN DE ACTIVIDADES") ~ "TERMINACIÓN DE ACTIVIDADES TOTAL O PARCIAL",
    str_detect(TXFUENTE, "PLANEFA") ~ "PLANEFA",
    TRUE ~ TXFUENTE))%>%
  mutate(TXFUENTE= factor(TXFUENTE, levels = c("PLANEFA",
                                               "EMERGENCIA AMBIENTAL",
                                               "DENUNCIA AMBIENTAL",
                                               "SOLICITUD DE INTERVENCIÓN FORMULADA POR UN ORGANISMO PÚBLICO",
                                               "TERMINACIÓN DE ACTIVIDADES TOTAL O PARCIAL",
                                               "VERIFICACIÓN DEL CUMPLIMIENTO DE LAS MEDIDAS ADMINISTRATIVAS ORDENADAS POR EL OEFA",
                                               "VERIFICACIÓN DE ACUERDO DE CUMPLIMIENTO",
                                               "OTRAS CIRCUNSTANCIAS")))

#REGIONES DESAGREGADO===========================
if("dpto"=="dpto"){
  Acc_tem<- Acciones_R2 %>%
    filter(ACC_EJEC==1) %>%
    distinct(DIREC,TXMES,COORD,TXCUC,TXUBIGEO_UND)
  temp<- as.data.frame(str_split(Acc_tem$TXUBIGEO_UND," / ",simplify = TRUE))
  Acc_tem1 <- cbind(Acc_tem,temp)
  for(i in 5:length(colnames(Acc_tem1))){
    Acc_tem1[[i]]<- gsub("\\-.*","",Acc_tem1[[i]])
  }
  Acciones_Dpto<- Acc_tem1 %>%
    gather("VAR","REG_DESAGR", starts_with("V")) %>% 
    mutate(REG_DESAGR = trimws(REG_DESAGR, which = c("both"))) %>%
    select(-VAR, - TXUBIGEO_UND) %>% filter(REG_DESAGR!="") %>%
    distinct(TXMES,TXCUC,COORD,DIREC,REG_DESAGR)
  rm(Acc_tem,temp,Acc_tem1)
}

#ACTIVIDADES DESAGREGADO===========================
if("Activ"=="Activ"){
  Acc_tem<- Acciones_R2 %>%
    filter(ACC_EJEC==1) %>%
    distinct(DIREC,TXMES,COORD,TXCUC,TXACTIVIDAD_UND)
  temp<- as.data.frame(str_split(Acc_tem$TXACTIVIDAD_UND," / ",simplify = TRUE))
  Acc_tem1 <- cbind(Acc_tem,temp)
  # for(i in 3:length(colnames(Acc_tem1))){
  #   Acc_tem1[[i]]<- gsub("\\-.*","",Acc_tem1[[i]])
  # }
  Acciones_Activ<- Acc_tem1 %>%
    gather("VAR","ACTIV_DESAGR", starts_with("V"))%>% 
    mutate(ACTIV_DESAGR = trimws(ACTIV_DESAGR, which = c("both"))) %>%
    select(-VAR) %>% filter(ACTIV_DESAGR!="") %>%
    distinct(TXMES, TXCUC, TXACTIVIDAD_UND, COORD, DIREC, ACTIV_DESAGR)
  rm(Acc_tem,temp,Acc_tem1)
}



#Reportes
if(3==3) {
  #Acciones ejecutadas =========================
  ACC_EJECUTADAS<- Acciones_R2 %>%
    distinct(TXCUC, .keep_all = TRUE) %>%
    filter(ACC_EJEC==1) %>%
    mutate(COORD = factor(COORD,levels = c("MIN","HID","ELE","IND","PES","AGR","CAM","RES","OD_HID","OD_RES")))%>%
    group_by(#TRIM,TXMES,
      TXTIPSUP, TXACCION, COORD) %>% 
    summarise(TOTAL= n_distinct(TXCUC)) %>%
    spread(COORD, TOTAL) %>%
    replace(is.na(.),0) %>%
    arrange(desc(TXTIPSUP), desc(TXACCION)) %>%
    adorn_totals("row") %>%
    adorn_totals("col")
  
  formattable(ACC_EJECUTADAS) 
  
  #Acciones por tipo de supervisión y tipo de acción=====================================
  ACC_EJEC_TIPO<-Acciones_R2 %>%
    distinct(TXCUC, .keep_all = TRUE) %>%
    filter(ACC_EJEC==1)%>%
    mutate(COORD = factor(COORD,levels = c("MIN","HID","ELE","IND","PES","AGR","CAM","RES","OD_HID","OD_RES")))%>%
    mutate(TXACCION= case_when(
      TXACCION=="PRESENCIAL" ~ "IN SITU",
      TXACCION=="NO PRESENCIAL" ~ "EN GABINETE",
      TRUE ~ TXACCION))%>%
    group_by(#DIREC,TRIM,TXMES ,
             TXTIPSUP,TXACCION,COORD)%>%
    summarise(TOTAL= n_distinct(TXCUC)) %>%
    arrange(desc(COORD)) %>%
    spread(COORD,TOTAL) %>%
    replace(is.na(.),0) %>%
    arrange(#MES,
      desc(TXTIPSUP),desc(TXACCION)) %>%
    adorn_totals("row") %>% adorn_totals("col")
  formattable(ACC_EJEC_TIPO) 
  
  #Acciones orientativas===========================
  ACC_EJEC_ORIENTA<- Acciones_R2 %>%
    distinct(TXCUC, .keep_all = TRUE) %>%
    filter(ACC_EJEC==1) %>%
    mutate(COORD = factor(COORD,levels = c("MIN","HID","ELE","IND","PES","AGR","CAM","RES","OD_HID","OD_RES")))%>%
    mutate(
      FGSUP_ORIENTATIVA=case_when(
        is.na(FGSUP_ORIENTATIVA) ~ "NO ORIENTATIVA",
        str_detect(FGSUP_ORIENTATIVA,"1") ~ "ORIENTATIVA",
        str_detect(FGSUP_ORIENTATIVA,"0|9") ~ "NO ORIENTATIVA"
      )) %>%
    group_by(#TRIM,TXMES, 
             FGSUP_ORIENTATIVA ,TXTIPSUP ,COORD) %>% #Trimestral
    summarise(TOTAL= n_distinct(TXCUC)) %>%
    spread(COORD, TOTAL) %>%
    replace(is.na(.),0) %>%
    arrange(desc(FGSUP_ORIENTATIVA), desc(TXTIPSUP)) %>%
    adorn_totals("row") %>% adorn_totals("col")
  
  formattable(ACC_EJEC_ORIENTA)
  
  #Acciones por region============================================= 
  ACC_EJEC_REGION<- Acciones_Dpto %>%
    mutate(REG_DESAGR = str_replace(REG_DESAGR,"ANCASH ..........","ANCASH"))%>%
    mutate(COORD = factor(COORD,levels = c("MIN","HID","ELE","IND","PES","AGR","CAM","RES","OD_HID","OD_RES")))%>%
    distinct(TXCUC, REG_DESAGR , .keep_all = TRUE)%>%
    group_by(REG_DESAGR, COORD) %>%
    summarise(ACCIONES_CONTABILIZADAS= n_distinct(TXCUC)) %>%
    spread(COORD,ACCIONES_CONTABILIZADAS) %>%
    replace(is.na(.),0) %>%
    adorn_totals("row") %>%
    adorn_totals("col")
  
  formattable(ACC_EJEC_REGION) 
  
  #Acciones por admnistrado y unidad fiscalizable ============================
  ACC_EJEC_ADMIN<-Acciones_R2 %>%
    distinct(TXCUC, .keep_all = TRUE) %>%
    filter(ACC_EJEC==1) %>%
    mutate(IDADMINISTRADO = if_else(str_detect(TXADMINISTRADO_ADM,"NO DETERMINADO"),"",IDADMINISTRADO)) %>%
    distinct(COORD, IDADMINISTRADO, IDUF_SIG, .keep_all = TRUE) %>% #Trimestre
    group_by(COORD)%>% #Trimestral
    summarise(ADMINISTRADO= n_distinct(IDADMINISTRADO, na.rm = TRUE),
              UF= n_distinct(IDUF_SIG, na.rm = TRUE)) %>%
    gather(key = "ITEM", value = "VALORES", ADMINISTRADO:UF) %>%
    arrange(COORD) %>%
    spread(COORD, VALORES) %>%
    adorn_totals("col")
  
  formattable(ACC_EJEC_ADMIN)
  
  
  #Acciones por fuente==============================================
  ACC_EJEC_FUENTE<-  Acciones_R2 %>%
    distinct(TXCUC, .keep_all = TRUE) %>%
    filter(ACC_EJEC==1)%>%
    mutate(COORD = factor(COORD,levels = c("MIN","HID","ELE","IND","PES","AGR","CAM","RES","OD_HID","OD_RES")))%>%
    group_by(#TRIM, DIREC,
             TXFUENTE,COORD)%>%
    summarise(TOTAL = n_distinct(TXCUC))%>%
    arrange(desc(COORD)) %>%
    spread(COORD, TOTAL) %>%
    replace(is.na(.),0) %>%
    adorn_totals("row") %>%
    adorn_totals("col")
  
  formattable(ACC_EJEC_FUENTE) 
  
  #Acciones por actividad=========================
  ACC_EJEC_ACTIVIDAD<- Acciones_Activ %>%
    mutate(COORD = factor(COORD,levels = c("MIN","HID","ELE","IND","PES","AGR","CAM","RES","OD_HID","OD_RES")))%>%
    distinct(TXCUC, ACTIV_DESAGR, .keep_all = TRUE)%>%
    mutate(ACTIV_DESAGR= str_replace(ACTIV_DESAGR,"([0-9]+ - )","")) %>%
    group_by(COORD,ACTIV_DESAGR) %>% #Trimestral
    summarise(ACCIONES_CONTABILIZADAS= n_distinct(TXCUC)) %>%
    spread(COORD,ACCIONES_CONTABILIZADAS)%>%
    replace(is.na(.),0) %>%
    adorn_totals("row") %>%
    adorn_totals("col")
  
  formattable(ACC_EJEC_ACTIVIDAD) 
  
  
  #Acciones por estado no consideradas=====================================
  ACC_TOTAL_ESTADO<- Acciones_R2 %>%
    mutate(COORD = factor(COORD,levels = c("MIN","HID","ELE","IND","PES","AGR","CAM","RES","OD_HID","OD_RES")))%>%
    distinct(TXCUC, .keep_all = TRUE) %>%
    group_by(TXESTADO,COORD) %>%
    summarise(NRO_ACCIONES = n_distinct (TXCUC)) %>%
    spread(COORD,NRO_ACCIONES) %>%
    filter(str_detect(TXESTADO,"ANULADA|NO EJECUTADA|EN REVISION" )) %>%
    replace(is.na(.),0) %>%
    adorn_totals("row") %>%
    adorn_totals("col")
  
  formattable(ACC_TOTAL_ESTADO) 
  
  #Supervisiones (Expedientes) Activas=====================================
  TOTAL_SUP<- Acciones_R2 %>%
    mutate(COORD = factor(COORD,levels = c("MIN","HID","ELE","IND","PES","AGR","CAM","RES","OD_HID","OD_RES")))%>%
    distinct(TXCUC, .keep_all = TRUE) %>%
    filter(ACC_EJEC==1) %>%
    mutate(ANO_EXP= parse_number(str_extract(TXNUMEXP,"(?<=-)[0-9]+"))) %>% #precedido de - 
    group_by(TXNUMEXP) %>%
    filter(FEINICIO== min(as.Date(FEINICIO))) %>%
    mutate(
      FGSUP_ORIENTATIVA=case_when(
        is.na(FGSUP_ORIENTATIVA) ~ "NO ORIENTATIVA",
        str_detect(FGSUP_ORIENTATIVA,"1") ~ "ORIENTATIVA",
        str_detect(FGSUP_ORIENTATIVA,"0|9") ~ "NO ORIENTATIVA"
      )) %>%
    group_by(COORD,TXTIPSUP
             ,FGSUP_ORIENTATIVA) %>%
    summarise(SUPERV = n_distinct(TXNUMEXP)) %>%
    spread(COORD, SUPERV) %>%
    arrange(desc(TXTIPSUP)
            ,desc(FGSUP_ORIENTATIVA)) %>%
    replace(is.na(.),0) %>%
    adorn_totals("row") %>%
    adorn_totals("col")
  
  formattable(TOTAL_SUP)
  
}


#Exportamos reporte Boletin=============
if(5==5){
  #setwd("Y:/Sistematización y Gestión de Procesos/33. Informacion Consolidada CSEP/REPORTE/BOLETIN/2020/")
  #setwd("C:/Users/JHON/Documents/OEFA/Outputs/Boletin/")
  setwd("C:/Users/jach_/OneDrive/Documentos/OEFA/Outputs/Boletin/")
  #Creamos Libro
  wb<- createWorkbook()
  modifyBaseFont(wb,fontSize = 10,fontName = "calibri", fontColour = "blue") # Solamente Aplica para writeData
  #Crea Hojas
  #----------------------
  addWorksheet(wb,"Acciones", gridLines = FALSE)
  addWorksheet(wb,"Acciones_Regiones", gridLines = FALSE)
  addWorksheet(wb,"Acciones_Activid", gridLines = FALSE)
  addWorksheet(wb, "Reporte_Meta", gridLines = FALSE)
  #Tabla 1.0
  #writeDataTable(wb,"Acciones", Acciones_R2,startRow = 1 ,startCol = "A")
  #Tabla 1.1
  writeDataTable(wb,"Acciones_Regiones", Acciones_Dpto,startRow = 1 ,startCol = "A")
  #Tabla 1.1
  writeDataTable(wb,"Acciones_Activid", Acciones_Activ,startRow = 1 ,startCol = "A")
  #Tabla 1.1
  writeDataTable(wb,"Reporte_Meta", ACC_EJEC_TIPO,startRow = 1 ,startCol = "A")
  writeDataTable(wb,"Reporte_Meta", ACC_EJEC_ORIENTA,startRow = 9 ,startCol = "A")
  writeDataTable(wb,"Reporte_Meta", ACC_EJEC_REGION,startRow = 17 ,startCol = "A")
  writeDataTable(wb,"Reporte_Meta", ACC_EJEC_ADMIN,startRow = 46 ,startCol = "A")
  writeDataTable(wb,"Reporte_Meta", ACC_EJEC_FUENTE,startRow = 51 ,startCol = "A")
  writeDataTable(wb,"Reporte_Meta", ACC_TOTAL_ESTADO,startRow = 62 ,startCol = "A")
  writeDataTable(wb,"Reporte_Meta", TOTAL_SUP,startRow = 68 ,startCol = "A")
  
  #Guardamos data
  saveWorkbook(wb, "Acciones_Diciembre2021.xlsx", overwrite = TRUE)
}

write.csv(Acciones_R2,file = "Acciones_Diciembre2021a.csv", row.names = TRUE, na="")
