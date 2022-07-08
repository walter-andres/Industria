#------CONSOLIDAR RUNT
library(tidyverse)
library(dplyr)
library(readxl)
library(lubridate)
# library(stringr)
# library(openxlsx)

## LA INFORMACIï¿½N DE MATRï¿½CULAS DIARIAS ##--------
l<-list.files("G:/SASData/Marketing/fuentes_compartidas/prendas/matriculas_otros_vehiduclos")
Sys.setlocale("LC_TIME", "Spanish")

#Cargamos la base de matriculas dirias que nso muestran como va la venta de vehï¿½culos en colombia 2022
suppressMessages(for (i in 1:length(l)) {
  y<-readxl::read_xlsx(paste0("G:/SASData/Marketing/fuentes_compartidas/prendas/matriculas_otros_vehiduclos/",l[i]),  sheet = "Industria",
                       col_names = T, .name_repair = "unique")
  
  FC<-lubridate::ymd(as.Date(paste("2022",substr(l[i],4,5),substr(l[i],1,2), sep = "-"), format = "%Y-%m-%d"))
  y$FC<-FC
  assign(paste0("matr_dia",FC), y); rm(y)
})

colnames(`matr_dia2022-01-01`)[2] <- "ANO_MATRICULA"
colnames(`matr_dia2022-02-01`)[2] <- "ANO_MATRICULA"
colnames(`matr_dia2022-03-01`)[2] <- "ANO_MATRICULA"
colnames(`matr_dia2022-04-01`)[2] <- "ANO_MATRICULA"
colnames(`matr_dia2022-05-01`)[2] <- "ANO_MATRICULA"
colnames(`matr_dia2022-06-01`)[2] <- "ANO_MATRICULA"

#Matrï¿½culas diaria del ano 2021-----------
yy<-readxl::read_xlsx("G:/SASData/Marketing/fuentes_compartidas/prendas/matriculas 2021.xlsx", 
                      .name_repair = "unique")

colnames(yy)[2] <- "ANO_MATRICULA"

yy <- yy %>% 
  mutate(FC = lubridate::ymd(paste(ANO_MATRICULA, MES_MATRICULA, `DIA MI`, sep = "-")))

#Consolido el ano 2021 y 2022
data<- rbind(`matr_dia2022-01-01` , `matr_dia2022-02-01`, `matr_dia2022-03-01`, `matr_dia2022-04-01`, 
             `matr_dia2022-05-01`, `matr_dia2022-06-01`, yy) %>% 
  filter(MTM %in% c("VP", "VU")) %>%
  mutate(MARCA = case_when(
    MARCA %in% c("CAN-AM", "CAN AM") ~ "fizz",
    MARCA %in% c("MERCEDES-BENZ", "MERCEDES BENZ") ~ "MERCEDES-BENZ",
    MARCA %in% c("MITSUBISHI", "MITSUBISHI FUSO") ~ "MITSUBISHI",
    MARCA %in% c("SINOTRUK", "SINOTRUK SITRAK") ~ "SINOTRUK",
    TRUE ~ MARCA),
    FC = lubridate::ymd(paste(ANO_MATRICULA, MES_MATRICULA, `DIA MI`, sep = "-")),
    SERVICIO = toupper(SERVICIO),
    SEGMENTO = toupper(SEGMENTO),
    DEPARTAMENTO = toupper(DEPARTAMENTO),
    ZONA = toupper(ZONA),
    NIVEL_EMISIONES = toupper(NIVEL_EMISIONES),
    fila = 1)

colnames(data)[24] <- "PRENDA1"
colnames(data)[46] <- "PRENDA2"

#cargo base de homologacion para hacer el cruce con acreedor
hom_entity<-readxl::read_xlsx("G:/SASData/Marketing/fuentes_compartidas/homologacion/homologar_entity.xlsx",
                              sheet = "Homologacion Prenda", 
                              col_types = c("numeric", "numeric", "text", "text", "text", "text")) %>% 
  select(NIT, `ENTIDAD FINAL_VF`) %>% 
  distinct(NIT, .keep_all = T)

colnames(hom_entity)[2] <- "ENTIDAD FINAL"

data<- data %>% 
  left_join(hom_entity, by = c("ACREEDOR" = "NIT"))

data <- data %>%
  mutate(NOM_ENTITY = case_when(
    LEASING == "SI" ~ paste0("LEASING ", data$`ENTIDAD FINAL`),
    PRENDA1 == "SI" ~ (data$`ENTIDAD FINAL`),
    TRUE ~ "NO_INFO"),
    TIPO_FINC = case_when(
      LEASING == "SI" ~ "LEASING ",
      PRENDA1 == "SI" ~ "NO LEASING ",
      TRUE ~ "NO PRENDA"),
    PRENDA1 = case_when(
      LEASING == "SI" | PRENDA1 == "SI" ~ "SI",
      TRUE ~ "NO")
    )

rm(`matr_dia2022-01-01` , `matr_dia2022-02-01`, `matr_dia2022-03-01`, `matr_dia2022-04-01`,
   `matr_dia2022-05-01`, `matr_dia2022-06-01`, yy, hom_entity)
gc()

#Import Runt Info to Bring the Sala-informe-diario
runt2021 <- readxl::read_excel("G:/SASData/Marketing/fuentes_compartidas/runt.BD.R/12.2021 RUNT D.H 22 (Enero-Diciembre) CIERRE.xlsx",
                               sheet = "MAT. RUNT") %>%
  as_tibble()

runt2022 <- readxl::read_excel("G:/SASData/Marketing/fuentes_compartidas/runt.BD.R/05.2022 RUNT D.H.xlsx",
                               sheet = "MAT. RUNT") %>%
  as_tibble()

#paste the df and filter the solumn that we requiered
runt <- rbind(runt2021, runt2022) %>% 
  select(CONSECUTIVO, `SALA INFORME DIARIO`, COD_BIR) %>% 
  mutate(`SALA INFORME DIARIO` = toupper(`SALA INFORME DIARIO`),
         `SALA INFORME DIARIO` = case_when(
           `SALA INFORME DIARIO` %in% c("TAYRONA VALLEDUPAR", "TAYRONA ( VALLEDUPAR )") ~ "TAYRONA VALLEDUPAR",
           `SALA INFORME DIARIO` %in% c("SANAUTOS CÚCUTA", "SANAUTOS CUCUTA") ~ "SANAUTOS CUCUTA",
           `SALA INFORME DIARIO` %in% c("SINCROMOTORS S.A. (CHIA)", "SINCROMOTORS CHIA") ~ "SINCROMOTORS CHIA",
           `SALA INFORME DIARIO` %in% c("SANAUTOS S.A  BARRANCABER", "SANAUTOS BARRANCABERMEJA") ~ "SANAUTOS BARRANCABERMEJA",
           `SALA INFORME DIARIO` %in% c("SANAUTOS CÚCUTA", "SANAUTOS CUCUTA") ~ "SANAUTOS CUCUTA",
           `SALA INFORME DIARIO` %in% c("PASTO MOTORS S.A.S.", "PASTO MOTORS PRINCIPAL") ~ "PASTO MOTORS",
           `SALA INFORME DIARIO` %in% c("MOTOCOSTA S.A.S.", "MOTOCOSTA", "MOTOCOSTA NORTE") ~ "MOTOCOSTA",
           `SALA INFORME DIARIO` %in% c("JUANAUTOS EL CERRO S.A.S", "JUANAUTOS EL CERRO") ~ "JUANAUTOS EL CERRO",
           `SALA INFORME DIARIO` %in% c("CENTRO AUTOMOTOR LTDA", "CENTRO AUTOMOTOR") ~ "CENTRO AUTOMOTOR",
           `SALA INFORME DIARIO` %in% c("CASATORO S.A. VILLAVICENC", "CASA TORO VILLAVICENCIO") ~ "CASA TORO VILLAVICENCIO",
           `SALA INFORME DIARIO` %in% c("CASATORO S.A. IBAGUE", "CASA TORO IBAGUE") ~ "CASA TORO IBAGUE",
           `SALA INFORME DIARIO` %in% c("CASA BRITANICA S.A (ENVIG", "CASA BRITANICA ENVIGADO") ~ "CASA BRITANICA ENVIGADO",
           `SALA INFORME DIARIO` %in% c("CASA BRITANICA MONTERIA", "CASA BRITANICA  S.A (MONT") ~ "CASA BRITANICA MONTERIA",
           `SALA INFORME DIARIO` %in% c("AUTOMOTRIZ CALDAS MOTOR S", "AUTOMOTRIZ CALDAS MOTOR PRINCIPAL") ~ "AUTOMOTRIZ CALDAS MOTOR",
           `SALA INFORME DIARIO` %in% c("ALIANZA MOTOR BOYACA 166", "ALIANZA M. BOYACA 166") ~ "ALIANZA MOTOR BOYACA 166",
           `SALA INFORME DIARIO` %in% c("ALBORAUTOS S.AS. POPAYAN", "ALBORAUTOS POPAYÁN", "ALBORAUTOS POPAYAN") ~ "ALBORAUTOS POPAYAN",
           `SALA INFORME DIARIO` %in% c("AGENCIAUTO PALACÉ", "AGENCIAUTO PALACE") ~ "AGENCIAUTO PALACE",
           TRUE ~ `SALA INFORME DIARIO`))


runt <- runt %>% distinct(CONSECUTIVO, .keep_all = TRUE)

rm(runt2021, runt2022)
gc()

#Import Cod Renault  - 
id.city <- readxl::read_excel("G:/SASData/Marketing/fuentes_compartidas/homologacion/cod.renault.xlsx",
                               sheet = "ID.SITIO") %>%
  as_tibble() %>%
  mutate(Sala = toupper(`Nombre Sitio`))

id.city <- id.city %>% 
  unique()

#We must bring thew zone RCI & Renault from "homo sala comercial"##
#Debemso tomar b- que son las de Sala Renault ycambiar la info de Comercial !a!.
z.rci<- readxl::read_excel("G:/SASData/Marketing/fuentes_compartidas/homologacion/Homologación salas_comercial.xlsx") %>%
  as_tibble() %>%
  mutate(Sala = toupper(`SALA REPORTES`),
         TIPO = toupper(`TIPO DE MERCADO`),
         zona_rci = toupper(`ZONA RCI`),
         Sala = case_when(
           Sala %in% c("AGENCIAUTO APARTADO", "AGENCIAUTO APARTADÓ PA") ~ "AGENCIAUTO APARTADO",
           Sala %in% c("AGENCIAUTO PRO+ INDUSTRIALES") ~ "AGENCIAUTO APARTADO TALLER",
           Sala %in% c("AGENCIAUTO VIVA ENVIGADO") ~ "AGENCIAUTO AYURA",
           Sala %in% c("AGENCIAUTO ITAGÜI") ~ "AGENCIAUTO ITAGUI",
           Sala %in% c("AGENCIAUTO PALACÉ") ~ "AGENCIAUTO PALACE",
           Sala %in% c("AGENCIAUTO LA 77") ~ "AGENCIAUTO PLAZA 77",
           Sala %in% c("ALBORAUTOS BUENAVENTURA") ~ "ALBORAUTOS FLORENCIA",
           Sala %in% c("AGENCIAUTO ITAGÜI") ~ "AGENCIAUTO ITAGUI",
           Sala %in% c("AGENCIAUTO ITAGÜI") ~ "AGENCIAUTO ITAGUI") %>% 
  unique()

z.rci <- z.rci %>% 
  select(Sala, zona_rci, TIPO) %>%
  group_by(Sala, zona_rci, TIPO) %>% 
  count()

z.rci <- subset (z.rci, select = -n)

#Join and select the columns that are requiered
id.city1 <- id.city %>% 
  left_join(z.rci, by = c("Sala" = "Sala"))



#CARGAMOS PARA MIRAR SALA
a<- (as.character(year(today()-30)))
b<- (str_pad(month(today()-30), 2, side = "left", pad = "0"))

runt_renault<- readxl::read_excel("G:/SASData/Marketing/fuentes_compartidas/runt.BD.R/05.2022 RUNT D.H.xlsx", 
                                  sheet = "MAT. RUNT", .name_repair = "unique") %>% 
  as_tibble()

colnames(runt_renault)[3] <- "ANO_MI"

runt_renault <- runt_renault %>%
  mutate(VIN = toupper(VIN),
         SALA.INFORME.DIARIO = toupper(`SALA INFORME DIARIO`),
         CONCESIONARIO.INFORME.DIARIO = toupper(`CONCESIONARIO INFORME DIARIO`),
         ANO_IMMAT = ANO_MI,
         MES_MI = str_pad(`MES MI`, 2, side = "left", pad = "0")) %>%
  filter(ANO_IMMAT == "2022" & MES_MI >= b) %>% 
  select(CONSECUTIVO, SALA.INFORME.DIARIO, CONCESIONARIO.INFORME.DIARIO)

MARCA_VEH2<- data %>%
  left_join(runt_renault, by = c("ID_VEHICULO" = "CONSECUTIVO")) %>% 
  select(FC, MARCA, NOM_ENTITY, fila, PRENDA1, SALA.INFORME.DIARIO, CONCESIONARIO.INFORME.DIARIO, COMBUSTIBLE, TIPO_FINC) %>%
  group_by(FC, MARCA, NOM_ENTITY, PRENDA1, SALA.INFORME.DIARIO, CONCESIONARIO.INFORME.DIARIO, COMBUSTIBLE, TIPO_FINC) %>%
  count(name = "CANTIDAD")

#CARGAMOS PARA MIRAR POR MARCA Y LINEA DE VEH
LINEA_VEH<- data %>% 
  filter(MARCA == "KIA" |
           MARCA == "NISSAN" |
           MARCA == "RENAULT") %>%
  select(FC, MARCA, LINEA, NOM_ENTITY, fila, PRENDA1, COMBUSTIBLE, TIPO_FINC) %>%
  group_by(FC, MARCA, LINEA, NOM_ENTITY, PRENDA1, COMBUSTIBLE, TIPO_FINC) %>%
  count(name = "CANTIDAD")

#Cargamos según cod de renault -ID SITIO
data$ID_VEHICULO <- gsub('\\s+', '', data$ID_VEHICULO)
runt$CONSECUTIVO <- gsub('\\s+', '', runt$CONSECUTIVO)

Matr.city<- data %>%
  filter(MARCA == "RENAULT") %>% 
  left_join(runt, by = c("ID_VEHICULO" = "CONSECUTIVO")) %>% 
  mutate(`SALA INFORME DIARIO` = case_when(
    is.na(`SALA INFORME DIARIO`) ~ "OTHER",
    TRUE ~ `SALA INFORME DIARIO`)
    )

Matr.city <- Matr.city %>%
  left_join(id.city, by = c("COD_BIR" = "ID Sitio")) %>% 
  mutate(Pais = case_when(
    is.na(Pais) ~ "COLOMBIA",
    TRUE ~ Pais)
  )

Matr.city <- Matr.city %>% 
  select(FC, LINEA, NOM_ENTITY, fila, PRENDA1, DEPARTAMENTO, MUNICIPIO, `SALA INFORME DIARIO`, Pais, Ciudad, 
         LONGITUDE, LATITUDE, FAMILIA, COMBUSTIBLE, TIPO_FINC) %>%
  group_by(FC, LINEA, NOM_ENTITY, fila, PRENDA1, DEPARTAMENTO, MUNICIPIO, `SALA INFORME DIARIO`, Pais, Ciudad, 
           LONGITUDE, LATITUDE, FAMILIA, COMBUSTIBLE, TIPO_FINC) %>%
  count(name = "Q")


##DATA FRAME GENERAL
#We must included the Dealer and department
data <- data %>% 
  left_join(runt, by = c("ID_VEHICULO" = "CONSECUTIVO")) %>%
  left_join(id.city, by = c("COD_BIR" = "ID Sitio"))

#Ordenamos la informaciï¿½n como la queremos ver para subirla a tableau
MARCA_VEH<- data %>%
  select(FC, MARCA, ZONA, FAMILIA, NOM_ENTITY, fila, PRENDA1, LINEA, `SALA INFORME DIARIO`, COMBUSTIBLE, TIPO_FINC) %>%
  group_by(FC, MARCA, ZONA, FAMILIA, NOM_ENTITY, PRENDA1, LINEA, `SALA INFORME DIARIO`, COMBUSTIBLE, TIPO_FINC) %>%
  count(name = "CANTIDAD")

#Add working days
dias_habiles<- data.table::fread("G:/SASData/Marketing/fuentes_compartidas/homologacion/dias_habiles.csv") %>% 
  as_tibble() %>% 
  mutate(FECHA = lubridate::dmy(paste0(DIA, sep = "-", MES, sep = "-", ANO)))

dias_habiles <- subset (dias_habiles, select = c(-DIA, -MES, -ANO))
dias_habiles$FECHA <- lubridate::ymd(dias_habiles$FECHA)

#First df
MARCA_VEH2<- MARCA_VEH2 %>%
  left_join(dias_habiles, by = c("FC" = "FECHA"))

#Second df
MARCA_VEH<- MARCA_VEH %>%
  left_join(dias_habiles, by = c("FC" = "FECHA"))

#Third df
LINEA_VEH <- LINEA_VEH %>% 
  left_join(dias_habiles, by = c("FC" = "FECHA"))

#Four df
Conces.veh <- Matr.city %>% 
  left_join(dias_habiles, by = c("FC" = "FECHA"))



#Esta es la que necesito
xlsxFileName <- paste("G:/SASData/Marketing/Code/8. Prendas/Output/cant_marca.xlsx")
writexl::write_xlsx(list(mysheet = MARCA_VEH, 
                         conces = MARCA_VEH2,
                         linea = LINEA_VEH,
                         dealer = Conces.veh), path = xlsxFileName, col_names = TRUE) 

# data2022 <- data %>%
#   filter(year(FC) == "2022")

xlsxFileName <- paste("G:/SASData/Marketing/fuentes_compartidas/prendas/industria.acum.xlsx")
writexl::write_xlsx(list(data = data), path = xlsxFileName, col_names = TRUE)


#-------------INFO WITH CHANNEL AND WITH OUT FLEETS---------------------------
# df.renault <- MARCA_VEH2 %>% 
#   filter(MARCA == "RENAULT" & `ENTIDAD FINAL` %in% c("FINANDINA", "BANCOLOMBIA", "BANCO DE OCCIDENTE",
#                                                      "BANCO DAVIVIENDA SA", "BANCO DE BOGOTA"))


