library(tidyverse)
library(dplyr)
library(readxl)
library(lubridate)
library(stringr)
# library(openxlsx)

#Armar la infromación de prendas
historico<-readxl::read_xlsx("G:/SASData/Marketing/fuentes_compartidas/prendas/historico.xlsx",
                             col_names = T,sheet = "Prenda MI", .name_repair = "unique") %>%
  select(ANO_MATRICULA = `AÑO INSCRIPCION PRENDA`, MES_MATRICULA = `MES INSCRIPCION PRENDA`, MARCA, NIT, ID_VEHICULO = CONSECUTIVO, ENTIDAD.FINAL.VF = `ENTIDAD FINAL`) %>% 
  mutate(ENTIDAD.FINAL.VF = toupper(ENTIDAD.FINAL.VF))

hom.entity.acreedor<-readxl::read_xlsx("G:/SASData/Marketing/fuentes_compartidas/prendas/homologar_entity.xlsx",
                              sheet = "Homologacion Prenda", 
                              col_types = c("numeric", "numeric", "text", "text", "text", "text")) %>% 
  distinct(NIT, .keep_all = T) %>% 
  select(NIT, `ENTIDAD FINAL_VF`) %>% 
  mutate(`ENTIDAD FINAL_VF` = toupper(`ENTIDAD FINAL_VF`))

hom.entity.leasing<-readxl::read_xlsx("G:/SASData/Marketing/fuentes_compartidas/prendas/homologar_entity.xlsx",
                                       sheet = "Homologacion Leasing Renting", 
                                       col_types = c("numeric", "numeric", "text", "text", "text", "text")) %>% 
  distinct(NIT, .keep_all = T) %>% 
  select(NIT, `ENTIDAD FINAL_VF`) %>% 
  mutate(`ENTIDAD FINAL_VF` = toupper(`ENTIDAD FINAL_VF`))

## -----------------LA INFORMACI?N DE MATR?CULAS DIARIAS ##------------------
l<-list.files("G:/SASData/Marketing/fuentes_compartidas/prendas/matriculas_otros_vehiduclos")
# Sys.setlocale("LC_TIME", "Spanish")

#Cargamos la base de matriculas dirias que nso muestran como va la venta de veh?culos en colombia 2022
for (i in 1:length(l)) {
  y<-readxl::read_xlsx(paste0("G:/SASData/Marketing/fuentes_compartidas/prendas/matriculas_otros_vehiduclos/",l[i]),  sheet = "Industria",
                       col_names = T, .name_repair = "unique")
  
  colnames(y)[2] <- "ANO_MATRICULA"
  FC <- as.Date(paste0( substr(l[i],6,9), sep = "-", substr(l[i],4,5), sep = "-","01"))
  y$FC<- FC
  file_n = substr(l[i],4,9)
  assign(paste0("matr_dia",file_n), y); rm(y)
}

prendas.arch.separados<- rbind(matr_dia012022, matr_dia022022, matr_dia032022, matr_dia042022, matr_dia052022,
                               matr_dia062022) %>% 
  mutate(ACREEDOR = as.numeric(ACREEDOR),
         NIT_LEASING = as.numeric(NIT_LEASING))

prendas.arch.separados.acree<- prendas.arch.separados %>% 
  filter(PRENDA...46 == "SI") %>% 
  left_join(hom.entity.acreedor, by = c("ACREEDOR" = "NIT"))

prendas.arch.separados.leas<- prendas.arch.separados %>% 
  filter(LEASING == "SI") %>% 
  left_join(hom.entity.leasing, by = c("NIT_LEASING" = "NIT"))

prendas.arch.separados.otros<- prendas.arch.separados %>% 
  filter(LEASING != "SI" & PRENDA...46 != "SI") %>% 
  left_join(hom.entity.leasing, by = c("NIT_LEASING" = "NIT"))

prendas.acum<- rbind(prendas.arch.separados.acree, prendas.arch.separados.leas, prendas.arch.separados.otros) %>%
  mutate(`ENTIDAD FINAL_VF` = toupper(`ENTIDAD FINAL_VF`))
  

rm(prendas.arch.separados.acree, prendas.arch.separados.leas, prendas.arch.separados.otros)
gc()

#Organizo el archivo solo con informaci´n que necesito para cruzar con matriculas
prendas.acum.mat<- prendas.acum %>%
  mutate(ACREEDOR = as.character(ACREEDOR),
         NIT_LEASING = as.character(NIT_LEASING),
         NIT = case_when(
           PRENDA...46 == "SI" ~ ACREEDOR,
           LEASING == "SI" ~ NIT_LEASING,
           (LEASING != "SI" & PRENDA...46 != "SI") ~ ""),
         ACREEDOR = as.numeric(ACREEDOR),
         NIT_LEASING = as.numeric(NIT_LEASING),
         ENTIDAD.FINAL.VF = toupper(`ENTIDAD FINAL_VF`))


#Historico segun fuentes del area
prendas.acum.mat1<- prendas.acum.mat %>% 
  select(ANO_MATRICULA, MES_MATRICULA, MARCA,NIT, ID_VEHICULO, ENTIDAD.FINAL.VF) 

prendas.acum.mat2<- rbind(historico, prendas.acum.mat1) %>% 
  mutate(ENTIDAD.FINAL.VF = toupper(ENTIDAD.FINAL.VF))

#Escribo el acumulado según funtes del area
xlsxFilePrend <- paste("G:/SASData/Marketing/fuentes_compartidas/prendas/Prendas Acumuladas.xlsx")
writexl::write_xlsx(list(Prendas = prendas.acum.mat2),
                    path = xlsxFilePrend, col_names = TRUE)


#---------------Segú información de INDUSTRIA de RENAULT-------------------------------------------------
#Partimos por traer historico de industria
prenda.ano2021<-readxl::read_xlsx("G:/SASData/Marketing/fuentes_compartidas/prendas/matriculas 2021.xlsx",
                      col_names = T, .name_repair = "unique") %>%
  mutate(
    FC = lubridate::ymd(paste(AÑO_MATRICULA, MES_MATRICULA, 01, sep = "-"))
    )

prendas.separados.acree<- prenda.ano2021 %>% 
  filter(PRENDA...46 == "SI") %>% 
  left_join(hom.entity.acreedor, by = c("ACREEDOR" = "NIT"))

prendas.separados.leas<- prenda.ano2021 %>% 
  filter(LEASING == "SI") %>% 
  left_join(hom.entity.leasing, by = c("NIT_LEASING" = "NIT"))

prendas.separados.otros<- prenda.ano2021 %>% 
  filter(LEASING != "SI" & PRENDA...46 != "SI") %>% 
  left_join(hom.entity.acreedor, by = c("ACREEDOR" = "NIT"))

prendas.separados.na<- prenda.ano2021 %>% 
  filter(is.na(LEASING) | is.na(PRENDA...46)) %>% 
  left_join(hom.entity.acreedor, by = c("ACREEDOR" = "NIT"))


prendas.origen<- rbind(prendas.separados.acree, prendas.separados.leas, 
                       prendas.separados.otros, prendas.separados.na) 

colnames(prendas.origen)[2] <- "ANO_MATRICULA"

prendas.origen1 <- prendas.origen %>% 
  mutate(ACREEDOR = as.character(ACREEDOR),
         NIT_LEASING = as.character(NIT_LEASING),
         NIT = case_when(
           PRENDA...46 == "SI" ~ ACREEDOR,
           LEASING == "SI" ~ NIT_LEASING,
           (LEASING != "SI" & PRENDA...46 != "SI") ~ ""),
         ACREEDOR = as.numeric(ACREEDOR),
         NIT_LEASING = as.numeric(NIT_LEASING),
         ENTIDAD.FINAL.VF = toupper(`ENTIDAD FINAL_VF`))

prendas.origen2<- rbind(prendas.acum.mat, prendas.origen1) %>% 
  mutate(ENTIDAD.FINAL.VF = toupper(`ENTIDAD FINAL_VF`))

xlsxFiletotal <- paste("G:/SASData/Marketing/Code/8. Prendas/Output/prendas.consolidado.xlsx")
writexl::write_xlsx(list(Prendas = prendas.origen2),
                    path = xlsxFiletotal, col_names = TRUE)

