library(tidyverse)
library(dplyr)
library(readxl)
library(lubridate)
library(stringr)
library(openxlsx)



historico<-readxl::read_xlsx("G:/SASData/Marketing/fuentes_compartidas/prendas/historico.xlsx",
                      col_names = T,sheet = "Prenda MI", .name_repair = "unique")

## LA INFORMACI?N DE MATR?CULAS DIARIAS ##
l<-list.files("G:/SASData/Marketing/Code/8. Prendas/matriculas_otros_vehiduclos")
# Sys.setlocale("LC_TIME", "Spanish")

#Cargamos la base de matriculas dirias que nso muestran como va la venta de veh?culos en colombia 2022
for (i in 1:length(l)) {
  y<-readxl::read_xlsx(paste0("G:/SASData/Marketing/Code/8. Prendas/matriculas_otros_vehiduclos/",l[i]),  sheet = "Industria",
                       col_names = T, .name_repair = "unique")
  
  
  FC<-lubridate::ymd(paste("2022",substr(l[i],4,6),substr(l[i],1,2), sep = "-"))
  y$FC<-as.Date(FC)
  file_n = substr(l[i],4,6)
  assign(paste0("matr_dia",file_n), y); rm(y)
  }

#Matr?culas diaria del a?o 2021
suppressMessages(yy<-readxl::read_xlsx("G:/SASData/Marketing/fuentes_compartidas/prendas/matriculas 2021.xlsx",
                            col_names = T, .name_repair = "unique") %>%
                   mutate(
                     FC = lubridate::ymd(paste(AÑO_MATRICULA, MES_MATRICULA, 01, sep = "-"))
                     ))

#Consolido el a?o 2021 y 2022
data<- rbind(matr_diaJan ,matr_diaFeb, yy) %>% 
  mutate(MARCA = case_when(
    MARCA %in% c("CAN-AM", "CAN AM") ~ "fizz",
    MARCA %in% c("MERCEDES-BENZ", "MERCEDES BENZ") ~ "MERCEDES-BENZ",
    MARCA %in% c("MITSUBISHI", "MITSUBISHI FUSO") ~ "MITSUBISHI",
    MARCA %in% c("SINOTRUK", "SINOTRUK SITRAK") ~ "SINOTRUK",
    TRUE ~ MARCA),
    SERVICIO = toupper(SERVICIO),
    SEGMENTO = toupper(SEGMENTO),
    DEPARTAMENTO = toupper(DEPARTAMENTO),
    ZONA = toupper(ZONA),
    NIVEL_EMISIONES = toupper(NIVEL_EMISIONES),
    fila = 1)

colnames(data)[24] <- "PRENDA1"
colnames(data)[46] <- "PRENDA2"

#cargo base de homologaci?n para hacer el cruce con acreedor
hom_entity<-readxl::read_xlsx("G:/SASData/Marketing/fuentes_compartidas/prendas/homologar_entity.xlsx",
                              sheet = "Homologacion Prenda", 
                              col_types = c("numeric", "numeric", "text", "text", "text", "text")) %>% 
  distinct(NIT, .keep_all = T) %>% 
  select(NIT, `ENTIDAD FINAL_VF`)

hom_entity_leasing<-readxl::read_xlsx("G:/SASData/Marketing/fuentes_compartidas/prendas/homologar_entity.xlsx",
                              sheet = "Homologacion Leasing Renting", 
                              col_types = c("numeric", "numeric", "text", "text", "text", "text")) %>% 
  distinct(NIT, .keep_all = T) %>% 
  select(NIT, ENTID_LEAS = `ENTIDAD FINAL_VF`)

#Validar que entidad no esta homologada
data<- data %>% 
  left_join(hom_entity_leasing, by = c("ACREEDOR" = "NIT")) %>% 
  left_join(hom_entity, by = c("ACREEDOR" = "NIT"))

data<- data %>% 
  mutate(NOM_ENTITY = case_when(
    (LEASING == "SI") & (is.na(ENTID_LEAS)) ~ paste0("LEASING ", `ENTIDAD FINAL_VF`),
    (LEASING == "SI") & (is.na(ENTID_LEAS)) ~ ENTID_LEAS,
    PRENDA1 == "SI" ~ `ENTIDAD FINAL_VF`,
    TRUE ~ "NO_INFO")
  )

val_acreedor<- data %>% 
  filter(is.na(`ENTIDAD FINAL_VF`)) %>% 
  group_by(ACREEDOR, `ENTIDAD FINAL_VF`,FC) %>% 
  count() %>% 
  view()

xlsxFileName <- paste("G:/SASData/Marketing/fuentes_compartidas/prendas/val_acreedor.xlsx")
writexl::write_xlsx(list(val_acreedor = val_acreedor), path = xlsxFileName, col_names = TRUE) 



#Ordenamos la informaci?n como la queremos ver para subirla a tableau
MARCA_VEH<- data %>%
  select(FC, MARCA, ZONA, FAMILIA, NOM_ENTITY, fila, PRENDA1, MTM) %>%
  group_by(FC, MARCA, ZONA, FAMILIA, NOM_ENTITY, PRENDA1, MTM) %>%
  count(name = "CANTIDAD")

FILE_PRENDAS<- data %>%
  select(NIT = ACREEDOR, ID_VEHICULO, NOM_ENTITY, AÑO_MATRICULA, MES_MATRICULA, MARCA, FC, MTM)



# MARCA_VEH %>% 
#   filter(MARCA == "RENAULT") %>% 
#   group_by(MARCA, MTM, FC) %>% 
#   count() %>% 
#   view()
# 


xlsxFileName <- paste("G:/SASData/Marketing/Code/8. Prendas/Output/cant_marca.xlsx")
writexl::write_xlsx(list(mysheet = MARCA_VEH), path = xlsxFileName, col_names = TRUE) 

xlsxFileName <- paste("G:/SASData/Marketing/Code/8. Prendas/Output/PRENDA.xlsx")
writexl::write_xlsx(list(mysheet = FILE_PRENDAS), path = xlsxFileName, col_names = TRUE) 


#Ingreso la base de runt del mes corriente y le agrego la base de prendas para
#mirar la entidad que financia y clasificar las ventas financiadas.
a<- (as.character(year(today()-30)))
b<- (str_pad(month(today()-30), 2, side = "left", pad = "0"))

data1 <- data %>% 
  distinct(ID_VEHICULO, .keep_all = TRUE) %>% 
  select( "ID_VEHICULO", "ENTIDAD FINAL_VF", "PRENDA1", "FAMILIA", "FC", "NOM_ENTITY")
  
runt_renault<- readxl::read_excel("G:/SASData/Marketing/fuentes_compartidas/runt_renault/02.2022 RUNT D.H 20 (Febrero) CIERRE.xlsx", 
                                  sheet = "MAT. RUNT", .name_repair = "unique") %>% 
  as_tibble() %>% 
  mutate(VIN = toupper(VIN),
         SALA.INFORME.DIARIO = toupper(`SALA INFORME DIARIO`),
         CONCESIONARIO.INFORME.DIARIO = toupper(`CONCESIONARIO INFORME DIARIO`),
         ANO_IMMAT = `AÑO MI`,
         MES_MI = str_pad(`MES MI`, 2, side = "left", pad = "0")) %>%
  filter(ANO_IMMAT == "2022" & MES_MI >= b)

runt_renault1 <- runt_renault %>% 
  left_join(data1, by= c("CONSECUTIVO" = "ID_VEHICULO"),
            na_matches = c("na", "never")) %>%
  mutate(VEH_CON_SIN_PRENDA = case_when(
    is.na(`ENTIDAD FINAL_VF`) ~ "SIN PRENDA",
    TRUE ~ "CON PRENDA"),
    VIN = VIN,
    `ENTIDAD FINAL` = case_when(
      is.na(`ENTIDAD FINAL_VF`) ~ "SIN PRENDA",
      TRUE ~ `ENTIDAD FINAL_VF`)
  )

runt_renault2<- runt_renault1 %>%
  mutate(`ENTIDAD FINAL` = toupper(`ENTIDAD FINAL`),
         cosecha = paste(ANO_IMMAT, MES_MI),
         ENTIDAD.QUE.FINANCIA.VAL.PRENDA = case_when(
           VEH_CON_SIN_PRENDA == "CON PRENDA" ~ `ENTIDAD FINAL`,
           VEH_CON_SIN_PRENDA == "SIN PRENDA" ~ "SIN PRENDA"
         ),
         TIPO.DE.FINANCIACION.VAL.PRENDA = case_when(
           VEH_CON_SIN_PRENDA == "CON PRENDA" ~ "CREDITO",
           str_detect(`ENTIDAD FINAL`, pattern = "(LEASING)") ~ "LEASING",
           VEH_CON_SIN_PRENDA == "SIN PRENDA" ~ "SIN PRENDA"))

#Ordenamos la informaci?n como la queremos ver para subirla a tableau
MARCA_VEH_runt1<- runt_renault2 %>%
  mutate(fila = 1,
         FAMILIA = toupper(FAMILIA.x)) %>% 
  select(FC, MARCA, ZONA, FAMILIA, ENTIDAD.QUE.FINANCIA.VAL.PRENDA, fila, PRENDA1, TIPO.DE.FINANCIACION.VAL.PRENDA) %>%
  group_by(FC, MARCA, ZONA, FAMILIA, ENTIDAD.QUE.FINANCIA.VAL.PRENDA, PRENDA1, TIPO.DE.FINANCIACION.VAL.PRENDA) %>%
  count(name = "CANTIDAD")

MARCA_VEH_runt2<- runt_renault2 %>%
  mutate(fila = 1,
         FAMILIA = toupper(FAMILIA.x),
         CONCESIONARIO = toupper(`CONCESIONARIO INFORME DIARIO`)) %>% 
  select(FC, MARCA, ZONA, CONCESIONARIO, ENTIDAD.QUE.FINANCIA.VAL.PRENDA, fila, PRENDA1, TIPO.DE.FINANCIACION.VAL.PRENDA) %>%
  group_by(FC, MARCA, ZONA, CONCESIONARIO, ENTIDAD.QUE.FINANCIA.VAL.PRENDA, PRENDA1, TIPO.DE.FINANCIACION.VAL.PRENDA) %>%
  count(name = "CANTIDAD")

MARCA_VEH_runt2<- runt_renault2 %>%
  mutate(fila = 1,
         FAMILIA = toupper(FAMILIA.x),
         
         CONCESIONARIO = toupper(`CONCESIONARIO INFORME DIARIO`)) %>% 
  select(FC, MARCA, ZONA, CONCESIONARIO, ENTIDAD.QUE.FINANCIA.VAL.PRENDA, fila, PRENDA1, TIPO.DE.FINANCIACION.VAL.PRENDA) %>%
  group_by(FC, MARCA, ZONA, CONCESIONARIO, ENTIDAD.QUE.FINANCIA.VAL.PRENDA, PRENDA1, TIPO.DE.FINANCIACION.VAL.PRENDA) %>%
  count(name = "CANTIDAD")

xlsxFileName <- paste("G:/SASData/Marketing/Code/8. Prendas/Output/runt_prenda.xlsx")
writexl::write_xlsx(list(mysheet = runt_renault,
                         conces = data1),
                    path = xlsxFileName, col_names = TRUE) 




            