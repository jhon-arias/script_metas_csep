#-------------------- CARGAR PAQUETES
library(easypackages)
pqt<- c("tidyverse","readxl","stringi","openxlsx","lubridate","bizdays" ,"janitor", "knitr","formattable","shadowtext","DT")
libraries(pqt)

#-------------------- PALETA DE COLORES Y TEMA

#Colores institucionales
OEFA.AZUL1<-c("#144AA7")
OEFA.AZUL2<-c("#1d85bf")
OEFA.TURQUEZA<-c("#0BC7E0")
OEFA.JADE<-c("#44bfb5")
OEFA.VERDE<-c("#8CCD3A")
OEFA.GRIS<-c("#696A6A")
OEFA.MOSTAZA<-c("#FFB500")
OEFA.HUMO<-c("#BCBCBC")
OEFA.LIMON<-c("#D5CB00")
OEFA.MORADO<-c("#4F4898")
OEFA.FUCSIA<-c("#E83670")
OEFA.NARANJA<-c("#EF7911")
OEFA.ROJO<-c("#EF0C33")

PALETA.PRINCIPAL<-c(OEFA.AZUL1,OEFA.AZUL2,OEFA.TURQUEZA,OEFA.JADE,OEFA.VERDE,OEFA.GRIS,OEFA.MOSTAZA)
PALETA.SECUNDARIA<-c(OEFA.HUMO,OEFA.LIMON,OEFA.MORADO,OEFA.FUCSIA,OEFA.NARANJA,OEFA.ROJO)

#Tema para DFAI
TEMA=    theme(panel.background = element_blank(),
               plot.title = element_text(hjust = 0.5),
               axis.line = element_line(colour = OEFA.GRIS),
               text=element_text(size=12),
               plot.caption = element_text(hjust = 0),
               axis.title.y.right = element_text( angle = 90)
)



#-------------------- IMPORTAR ARCHIVOS
DFAI_EXP<- read_excel("C:/Users/jach_/Downloads/DFAI OCTUBRE 2021.xlsx", sheet = "RSDRD", skip = 0) %>% filter(!is.na(EXPEDIENTE)) %>%
  rename("COORDINACION" = `COORDINACIÓN`, "RAZON" = ADMINISTRADO, "UNIDAD" = `UNIDAD FISCALIZABLE`, "TIPO" = TIPO_RES, "ANO_RPTE" = `AÑO_RPTE`) %>%
  mutate(FECHA_RESOL =  as.Date(FECHA_RESOL),
         FE_INI_SUP = as.Date(FE_INI_SUP),
         ANO_SUP =  year(FE_INI_SUP))
save(DFAI_EXP,file = "C:/Users/jach_/OneDrive/Documentos/OEFA/Inputs/Data2021/Agosto(gestion)/DFAI_Agosto_2021.RData")

DFAI_CORR<- read_excel("C:/Users/jach_/Downloads/DFAI OCTUBRE 2021.xlsx", sheet = "MEDIDAS CORRECTIVAS", skip = 0) %>% filter(!is.na(EXPEDIENTE)) %>%
  rename("ADMIN" = ADMINISTRADO, "ANO_RPTE"= `AÑO_RPTE`) %>%
  mutate(FECHA_INI_SUP = as.Date(FECHA_INI_SUP),
         FECHA_INF = as.Date(FECHA_INF),
         FECHA_RESOL = as.Date(FECHA_RESOL),
         ANO_RPTE = year(FECHA_RESOL),
         INF_SFAP_SFEM = if_else(is.na(INF_SFAP_SFEM), RESOL_DIRECT,INF_SFAP_SFEM),
         FECHA_INF = if_else(is.na(FECHA_INF),FECHA_RESOL, FECHA_INF))

save(DFAI_CORR,file = "C:/Users/jach_/OneDrive/Documentos/OEFA/Inputs/Data2021/Agosto(gestion)/DFAImc_Agosto_2021.RData")

DFAI_STOCK<- read_excel("C:/Users/jach_/Downloads/DFAI OCTUBRE 2021.xlsx", sheet = "STOCK", skip = 0)%>% filter(!is.na(DOCUMENTO)) %>%
  rename("ESTADO" = `ESTADO...4`, "RESOL" = `RESOL_CONCLUCIÓN` , "ESTADO_2" = `ESTADO...13`)

save(DFAI_STOCK,file = "C:/Users/jach_/OneDrive/Documentos/OEFA/Inputs/Data2021/Agosto(gestion)/DFAIstk_Agosto_2021.RData")

#-------------------- CONSULTAMOS ARCHIVOS
load("C:/Users/jach_/OneDrive/Documentos/ShinyApp/Planefa2021/Progr2021.RData")
#-----
load("C:/Users/jach_/OneDrive/Documentos/OEFA/Inputs/Data2021/Diciembre/DFAI_Diciembre_2021.RData")
load("C:/Users/jach_/OneDrive/Documentos/OEFA/Inputs/Data2021/Diciembre/DFAImc_Diciembre_2021.RData")
load("C:/Users/jach_/OneDrive/Documentos/OEFA/Inputs/Data2021/Diciembre/DFAIstk_Diciembre_2021.RData")

#-------------------- LIMPIESA DE DATOS
FECHA_EVALUADA<- c("2021-12-31")


#5---
DFAI_EXP<- DFAI_EXP %>%
  filter(str_detect(ESTADO,"CONCLUIDO"))%>%
  mutate(ANO_SUP =  year(FE_INI_SUP)) %>%
  mutate_if(is.character, function(X){toupper(X)})%>%
  mutate(SECTOR1= case_when(
    str_detect(SECTOR, "MIN") ~ "MIN",
    str_detect(SECTOR, "HID") ~ "HID",
    str_detect(SECTOR, "ELE") ~ "ELE",
    str_detect(SECTOR, "IND") ~ "IND",
    str_detect(SECTOR, "PES") ~ "PES",
    str_detect(SECTOR, "AGR") ~ "AGR",
    str_detect(SECTOR, "INFR") ~ "INF",
    str_detect(SECTOR, "AMBI") ~ "INF",
    str_detect(SECTOR, "RESI") ~ "INF"),
    SECTOR2= case_when(
      str_detect(SECTOR,"MENOR|COMER") ~ "HID_MEN",
      str_detect(SECTOR,"MAYOR") ~ "HID_MAY",
      TRUE ~  SECTOR1),
    MES_META = substr(MES_META,1,3),
    MES_META = if_else(str_detect(MES_META,"SEP"),"SET",MES_META))
#6---
DFAI_CORR <- DFAI_CORR %>%
  filter(!is.na(INF_SFAP_SFEM)) %>%
  #rename("MES_RPT" = MES) %>%
  mutate_if(is.character, function(X){toupper(X)})%>%
  mutate(RESULTADO = stri_trans_general(RESULTADO,"Latin-ASCII")) %>%
  mutate(SECTOR1= case_when(
    str_detect(SECTOR, "MIN") ~ "MIN",
    str_detect(SECTOR, "HID") ~ "HID",
    str_detect(SECTOR, "ELE") ~ "ELE",
    str_detect(SECTOR, "IND") ~ "IND",
    str_detect(SECTOR, "PES") ~ "PES",
    str_detect(SECTOR, "AGR") ~ "AGR",
    str_detect(SECTOR, "INFR") ~ "INF",
    str_detect(SECTOR, "AMBI") ~ "INF",
    str_detect(SECTOR, "RESI") ~ "INF"),
    SECTOR2= case_when(
      str_detect(SECTOR,"MENOR|COMER") ~ "HID_MEN",
      str_detect(SECTOR,"MAYOR") ~ "HID_MAY",
      TRUE ~  SECTOR1),
    MES_META = substr(MES_META,1,3),
    MES_META = if_else(str_detect(MES_META,"SEP"),"SET",MES_META)) 

#6.1 STock DFAI
DFAI_STOCK<- DFAI_STOCK %>%
  mutate(ID = row_number()) %>% #crea el ID si no viene
  mutate_if(is.character, function(X){toupper(X)})%>%
  filter(!str_detect(PNDNT_TRAMITADO,"CONCLUIDO"))%>%
  mutate(SECTOR= case_when(
    str_detect(SECTOR, "MIN") ~ "MIN",
    str_detect(SECTOR, "ELE") ~ "ELE",
    str_detect(SECTOR, "IND") ~ "IND",
    str_detect(SECTOR, "PES") ~ "PES",
    str_detect(SECTOR, "AGR") ~ "AGR",
    str_detect(SECTOR, "AMBI") ~ "INF",
    str_detect(SECTOR, "RESI") ~ "INF",
    str_detect(SECTOR, "INFRA") ~ "INF",
    str_detect(SECTOR,"MENOR|COMER") ~ "HID_MEN",
    str_detect(SECTOR,"MAYOR") ~ "HID_MAY"))

#-------------------- GENERACION DE TABLAS
EXP.DFAI<- DFAI_EXP %>%
  filter(FECHA_RESOL>= as.Date("2021-01-01"),
         FECHA_RESOL<= FECHA_EVALUADA)%>%
  rename(MES=MES_META) %>%
  group_by(SECTOR1, MES) %>% 
  summarise(NRO = n_distinct(RESOL),.groups = 'drop')%>%
  rename(SECTOR=SECTOR1)

CORR.DFAI<- DFAI_CORR %>%
  filter(FECHA_RESOL>= as.Date("2021-01-01"),
         FECHA_RESOL<= FECHA_EVALUADA)%>%
  rename(MES=MES_META) %>%
  mutate(SECTOR= "MC") %>%
  group_by(SECTOR, MES) %>%
  summarise(NRO= n_distinct(INF_SFAP_SFEM),.groups = 'drop')

EXP.CULM<- bind_rows(EXP.DFAI,CORR.DFAI) %>%
  rename(EJECUTADO = NRO)

#Seguimiento Metas DFAI
META.DFAI<- Progr2021 %>%
  filter(str_detect(FUNCION,"FISCALIZA")) %>%
  left_join(EXP.CULM) %>%
  replace(is.na(.),0)

#Detalle de Metas DFAI
EXP.DETALL<- DFAI_EXP %>%
  filter(FECHA_RESOL>= as.Date("2021-01-01"),
         FECHA_RESOL<= FECHA_EVALUADA)%>%
  rename(MES=MES_META) %>%
  group_by(SECTOR2, MES) %>% 
  summarise(EXP = n_distinct(RESOL),.groups = 'drop')%>%
  rename(SECTOR=SECTOR2)

COR.DETALL<- DFAI_CORR %>%
  filter(FECHA_RESOL>= as.Date("2021-01-01"),
         FECHA_RESOL<= FECHA_EVALUADA) %>%
  rename(MES=MES_META) %>%
  mutate(SECTOR="MC")%>%
  group_by(SECTOR, MES) %>%
  summarise(EXP= n_distinct(INF_SFAP_SFEM),.groups="drop")

META.DETALL<- EXP.DETALL %>% 
  full_join(COR.DETALL) %>%
  replace(is.na(.),0)

#Resultado de los expedientes DFAI
EXP.RESUL<- DFAI_EXP %>%
  filter(FECHA_RESOL>= as.Date("2021-01-01"),
         FECHA_RESOL<= FECHA_EVALUADA)%>%
  rename(MES=MES_META) %>%
  group_by(SECTOR2,TIPO, MES) %>% 
  summarise(EXP = n_distinct(RESOL),.groups = 'drop')%>%
  rename(SECTOR=SECTOR2,
         RESULTADO=TIPO)

COR.RESUL<- DFAI_CORR %>%
  filter(FECHA_RESOL>= as.Date("2021-01-01"),
         FECHA_RESOL<= FECHA_EVALUADA)%>%
  rename(MES=MES_META) %>%
  mutate(SECTOR = "MC") %>%
  group_by(SECTOR,RESULTADO,MES) %>%
  summarise(EXP= n_distinct(INF_SFAP_SFEM),.groups = 'drop') #%>%
# rename(RESULTADO=ESTADO)

META.RESUL<- EXP.RESUL %>% 
  full_join(COR.RESUL) %>%
  replace(is.na(.),0)

#Año de supervisión de expedientes DFAI
EXP.YEAR<- DFAI_EXP %>%
  filter(FECHA_RESOL>= as.Date("2021-01-01"),
         FECHA_RESOL<= FECHA_EVALUADA)%>%
  rename(MES=MES_META) %>%
  group_by(SECTOR2,ANO_SUP, MES) %>% 
  summarise(EXP = n_distinct(RESOL),.groups = 'drop')%>%
  rename(SECTOR=SECTOR2)

COR.YEAR<- DFAI_CORR %>%
  filter(FECHA_RESOL>= as.Date("2021-01-01"),
         FECHA_RESOL<= FECHA_EVALUADA)%>%
  rename(MES=MES_META) %>%
  mutate(SECTOR = "MC") %>%
  group_by(SECTOR,ANO_SUP ,MES) %>%
  summarise(EXP= n_distinct(INF_SFAP_SFEM),.groups = 'drop')

META.YEAR<- EXP.YEAR %>% 
  full_join(COR.YEAR) %>%
  replace(is.na(.),0)

#Motivos de no inicio de supervisión de expedientes DFAI
META.MOTIVO<- DFAI_EXP %>%
  filter(FECHA_RESOL>= as.Date("2021-01-01"),
         FECHA_RESOL<= FECHA_EVALUADA)%>%
  rename(MES=MES_META) %>%
  filter(str_detect(TIPO,"ARCHIVO|NO INICIO")) %>%
  group_by(MOTIVO, MES) %>% 
  summarise(EXP = n_distinct(RESOL),.groups = 'drop')

#Stock DFAI
STOCK.DFAI<- DFAI_STOCK %>%
  mutate(SECTOR = factor(SECTOR,levels = c("MIN","HID_MAY","HID_MEN","ELE","IND","PES","AGR","INF"))) %>%
  group_by(ANO_SUP,SECTOR)%>%
  summarise(EXP=n_distinct(ID),.groups = 'drop') %>%
  spread(SECTOR, EXP) %>%
  replace(is.na(.),0)%>%
  adorn_totals(name = "TOT.EXP", c("row","col"))


#-------------------- GENERACION DE GRAFICOS Y TABLAS

#11.1 Expedientes concluidos DFAI-----------------------------------------------
Graph_1<-  META.DFAI %>%
    group_by(SECTOR) %>%
    summarise(META = sum(META,na.rm = TRUE),
              EJECUTADO=sum(EJECUTADO, na.rm = TRUE),.groups = 'drop') %>%
    mutate(PORCENTAJE =round((EJECUTADO/META)*100,digits = 1)) %>%
    filter(!is.infinite(PORCENTAJE),!is.na(PORCENTAJE)) %>%
    mutate(SECTOR = factor(SECTOR,levels = c("MIN","HID","ELE","IND","PES","AGR","INF","MC"))) %>%
    arrange(desc(PORCENTAJE)) %>%
    ggplot(mapping = aes(SECTOR,PORCENTAJE)) +
    geom_hline(yintercept=100, linetype="dashed", 
               color = OEFA.ROJO, size=1) +
    geom_bar(stat = "identity",fill=OEFA.AZUL1, alpha=1, width=0.65)+
    geom_shadowtext(aes(label=paste(PORCENTAJE,"%",sep = "")),
                    #family = "Times New Roman",
                    vjust=-0.5,
                    hjust=0.5, #????
                    nudge_y = 1.5, #DISTANCIA DEL NUMERO A LA GRAFICA
                    size=4,
                    fontface=2)+
    labs(title = "Porcentaje de Avance de los Expedientes Concluidos, por subsector",
         # subtitle="",
         x="Subsector",
         y="Porcentaje de cumplimiento",
         caption = "Fuente: Reporte - DFAI", 
         size=4)+
    
    scale_y_continuous(limits = c(0, 300,0), breaks = seq(0,100,10), expand = c(0, 0))+
    TEMA

#11.2 Tabla de Expedientes concluidos DFAI-----------------------------------------------


Tabla_1<- META.DFAI %>% 
  mutate(SECTOR = factor(SECTOR,levels = c("MIN","HID","ELE","IND","PES","AGR","INF","MC"))) %>%
  group_by(SECTOR) %>%
  summarise(META = sum(META,na.rm = TRUE), EJECUTADO=sum(EJECUTADO, na.rm = TRUE),.groups = 'drop') %>%
  mutate(PORCENTAJE =paste(round((EJECUTADO/META)*100,digits = 1),"%",sep = " ")) %>%
  rename("Subsector" = SECTOR,"Meta" = META,"Ejecutado" = EJECUTADO,"Porcentaje" = PORCENTAJE)

#11.3 Tabla de porcentaje de avance de Expedientes concluidos DFAI----------------

Tabla_2<-  META.DETALL %>% 
  group_by(SECTOR) %>% summarise(EXP = sum(EXP,na.rm = TRUE),.groups = 'drop') %>%
  mutate(PCTJE = round(prop.table(EXP)*100, digits = 1)) %>%
  arrange(desc(PCTJE)) %>%
  mutate(Acumulado= cumsum(PCTJE)) %>%
  arrange(Acumulado) %>%
  rename("Subsector" =  SECTOR, "Expediente" = EXP, "Porcentaje"=PCTJE)

#11.4 Avance de Expedientes concluidos DFAI por subsector paretto-----------------------------------------------

Graph_2<- META.DETALL %>% 
    group_by(SECTOR) %>% summarise(EXP = sum(EXP,na.rm = TRUE),.groups = 'drop') %>%
    mutate(PCTJE = round(prop.table(EXP)*100, digits = 1)) %>%
    arrange(reorder(SECTOR, desc(EXP))) %>%
    mutate(ACUMUL= cumsum(PCTJE)) %>%
    ggplot(mapping = aes(x= reorder(SECTOR, desc(EXP)))) + 
    geom_bar(aes(y= EXP), stat = "identity", alpha=1, fill=OEFA.TURQUEZA, width = 0.75)+
    
    # geom_line(aes(y=2*ACUMUL, group=1),color="red", size=1,linetype = "solid")+
    
    scale_y_continuous(name = "Expedientes concluidos", sec.axis = sec_axis(~./3, name = "Porcentaje acumulado"), breaks = seq(0,300,50), expand = expansion(add = c(0, 10)))+
    labs(title = "Expedientes concluidos, por subsector",
         caption = "Fuente: Reporte - DFAI", 
         x="Subsector", size=4)+
    
    TEMA+
    geom_shadowtext(aes(y=EXP,label=paste(EXP,sep = ""), ),
                    hjust=0.5, #????
                    vjust=0, #????
                    nudge_y = 4, #DISTANCIA DEL NUMERO A LA GRAFICA
                    size=4,
                    fontface=2)+
    geom_point(aes(y=3*ACUMUL), shape=21,colour = OEFA.MOSTAZA, fill = "white", size = 10, stroke = 1.5)+
    
    geom_shadowtext(aes(y=3*ACUMUL,label=paste(ACUMUL,sep = ""), ),
                    hjust=0.5, #????
                    nudge_y = 0, #DISTANCIA DEL NUMERO A LA GRAFICA
                    size=4,
                    fontface=2)

#11.5 Tabla de resultado de Expedientes concluidos DFAI-----------------------------------------------

Tabla_3<- META.RESUL %>% 
  group_by(SECTOR,RESULTADO) %>% 
  mutate(RESULTADO = str_to_sentence(RESULTADO)) %>%
  summarise(EXP = sum(EXP,na.rm = TRUE),.groups = 'drop') %>%
  spread(SECTOR, EXP) %>% replace(is.na(.),0) %>% arrange(RESULTADO) %>%
  select(RESULTADO,MIN,HID_MAY,HID_MEN,ELE,IND,PES,AGR,INF,MC) %>%
  adorn_totals(c("row","col"))

#11.6 Tabla de motivos de Archivo/No inicio Expedientes concluidos DFAI---------


Tabla_4<- META.MOTIVO %>% 
  group_by(MOTIVO) %>% summarise(EXP = sum(EXP,na.rm = TRUE),.groups = 'drop') %>%
  mutate(MOTIVO = str_to_sentence(MOTIVO)) %>%
  mutate(PCTJE = round(prop.table(EXP)*100, digits = 1)) %>%
  arrange(PCTJE) %>% mutate(ACUMUL= cumsum(PCTJE)) %>%
  rename("Motivo" = MOTIVO,"Expediente" = EXP,"Porcentaje"=PCTJE,"Acumulado"=ACUMUL)

#11.6 Porcentaje de motivos de Archivo/No inicio Expedientes concluidos DFAI---------

Graph_3<-  META.MOTIVO %>% 
    group_by(MOTIVO) %>% summarise(EXP = sum(EXP,na.rm = TRUE),.groups = 'drop') %>%
    mutate(MOTIVO = str_to_sentence(MOTIVO)) %>%
    mutate(PCTJE = round(prop.table(EXP)*100, digits = 1)) %>%
    arrange(PCTJE) %>%
    mutate(ACUMUL= cumsum(PCTJE))%>% #Se creo para darle las listas
    ggplot(mapping = aes(x=2, y=-PCTJE, fill = reorder(str_wrap(MOTIVO,30), -PCTJE)))+
    scale_fill_manual(values = c(OEFA.AZUL1,OEFA.TURQUEZA,OEFA.JADE,OEFA.VERDE,OEFA.MOSTAZA,OEFA.GRIS,OEFA.MORADO))+ #labels=c("1","2","","","","","")
    geom_bar(stat = "identity", color="white",size=1, width=1.2)+
    
    geom_shadowtext(aes(label = paste(round(PCTJE), "%", sep = ""),x = 3),
                    color="white", fontface="bold", size=4,
                    position = position_stack(vjust = 0.5), check_overlap = T) +
    
    coord_polar(theta = "y")+
    labs(title = "Porcentaje de expedientes, por Motivos de Archivo/No Inicio", size=4)+
    theme(panel.background = element_blank(),
          plot.title = element_text(hjust = 0.5),
          text=element_text(size=12),
          legend.key.height=unit(1,"cm"),
          legend.title = element_blank(),
          axis.text = element_blank(),
          axis.ticks = element_blank(),
          axis.title.x = element_blank(),
          axis.title.y = element_blank())+
    xlim(c(1.5-7/3,3))

#11.7 Tabla de año de supervision de los Expedientes concluidos DFAI---------


Tabla_5<- META.YEAR %>%
  group_by(SECTOR,ANO_SUP) %>% summarise(EXP = sum(EXP,na.rm = TRUE),.groups = 'drop') %>%
  spread(SECTOR, EXP) %>% replace(is.na(.),0) %>% arrange(ANO_SUP) %>%
  select(`AÑO SUPERVISION`= ANO_SUP,MIN,HID_MAY,HID_MEN,ELE,IND,PES,AGR,INF,MC) %>%
  rename(`Año supervisión`=`AÑO SUPERVISION`) %>%
  adorn_totals(c("row","col"))%>% 
  filter(Total>0) 


#11.7 Tabla de Stock de expedientes por concluir de la DFAI---------  

Tabla_6<- STOCK.DFAI %>%
  mutate(ANO_SUP = if_else(str_detect(ANO_SUP,"TOT"),"Total",ANO_SUP))%>%
  arrange(ANO_SUP)%>%
  rename(`Año de supervisión` = ANO_SUP) %>%
  rename(`Total` = TOT.EXP)


if(5==5){
setwd("C:/Users/jach_/OneDrive/Documentos/OEFA/Outputs/Boletin/")

#-- Generamos lo lugares para exportar
  #lista de BD  
  df.list <- list(
    c("1.1 CANTIDAD DE EXPEDIENTES CONCLUIDOS RESPECTO A LA META"), as.data.frame(Tabla_1),
    c("1.2 CANTIDAD DE EXPEDIENTES CONCLUIDOS Y PORCENTAJE"), as.data.frame(Tabla_2),
    c("1.3 RESULTADOS DE LOS EXPEDIENTES CONCLUIDOS"), as.data.frame(Tabla_3),
    c("1.4 MOTIVOS DE ARCHIVO / NO INICIO"), as.data.frame(Tabla_4),
    c("1.5 EXPEDIENTES CONCLUIDOS SEGUN AÑO DE SUPERVISION"),as.data.frame(Tabla_5),
    c("1.6 STOCK DE EXPEDINETES "),as.data.frame(Tabla_6)
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
  
#-- Creamos libro --
wb1<- createWorkbook()
#-- Creamor hojas --
addWorksheet(wb1,"RSDRD")
addWorksheet(wb1,"MEDIDAS CORRECTIVAS")
addWorksheet(wb1,"STOCK")
addWorksheet(wb1,"REPORTE")
#-- Agregamos general --
writeData(wb1,"RSDRD","1. CONSOLIDADO RD Y RSD ",startRow = 1 ,startCol = "A")
writeDataTable(wb1,"RSDRD",DFAI_EXP, startRow = 2 ,startCol = "A")

writeData(wb1,"MEDIDAS CORRECTIVAS","1. CONSOLIDADO DE MEDIDAS CORRECTIVAS",startRow = 1 ,startCol = "A")
writeDataTable(wb1,"MEDIDAS CORRECTIVAS",DFAI_CORR, startRow = 2 ,startCol = "A")

writeData(wb1,"STOCK","1. CONSOLIDADO DEL STOCK",startRow = 1 ,startCol = "A")
writeDataTable(wb1,"STOCK",DFAI_STOCK, startRow = 2 ,startCol = "A")

#writeDataTable(wb1,"BD",INF_CONSOL_BOLETIN, startRow = 3 ,startCol = "A", na.string = "") #se debe revisar
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
writeData(wb1,"REPORTE",df.list[[11]],startRow = df.row$Posi[[11]] ,startCol = "A")
writeDataTable(wb1,"REPORTE",df.list[[12]], startRow = df.row$Posi[[12]] ,startCol = "A")

#Graph 01
print(Graph_1)
insertPlot(wb1, "REPORTE", xy = c("N", 2), width = 20, height = 10, fileType = "png", units = "cm" )
#Graph 02
print(Graph_2)
insertPlot(wb1, "REPORTE", xy = c("N", 25), width = 20, height = 10, fileType = "png", units = "cm" )
#Graph 03
print(Graph_3)
insertPlot(wb1, "REPORTE", xy = c("N", 50), width = 20, height = 10, fileType = "png", units = "cm" )

#-- Guardamos --
saveWorkbook(wb1,"DFAIExpedientes_Diciembre2021.xlsx", overwrite = TRUE)

}
