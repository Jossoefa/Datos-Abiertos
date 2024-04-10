library(easypackages)
pqt<- c("googlesheets4","gargle","googledrive","tidyverse" ,"readxl", "openxlsx","lubridate","bizdays" ,"janitor", "formattable")
libraries(pqt)

setwd("S:/Sistematización y Gestión de Procesos/33. Informacion Consolidada CSEP/REPORTE/BOLETIN/2023/BoletinOutput")

Informes_setiembre_2023 <- read.xlsx("Informes_setiembre_2023.xlsx", sheet = "BD")

names(Informes_setiembre_2023) <- Informes_setiembre_2023[1, ]
Informes_setiembre_2023 <- Informes_setiembre_2023[-1, ]

  Informes_setiembre_2023$FECHA_INFORME <- as.numeric(Informes_setiembre_2023$FECHA_INFORME)
  Informes_setiembre_2023$FECHA_INFORME <- as.Date(Informes_setiembre_2023$FECHA_INFORME, origin = "1899-12-30")
  #Informes_setiembre_2023$FECHA_INFORME <- format(Informes_setiembre_2023$FECHA_INFORME, format = "%d/%m/%Y")
  
  Informes_setiembre_2023$FECHA_DOC_CIERRE <- as.numeric(Informes_setiembre_2023$FECHA_DOC_CIERRE)
  Informes_setiembre_2023$FECHA_DOC_CIERRE <- as.Date(Informes_setiembre_2023$FECHA_DOC_CIERRE, origin = "1899-12-30")
  #Informes_setiembre_2023$FECHA_DOC_CIERRE <- format(Informes_setiembre_2023$FECHA_DOC_CIERRE, format = "%d/%m/%Y")
  
  Informes_setiembre_2023$FEFIN <- as.numeric(Informes_setiembre_2023$FEFIN)
  Informes_setiembre_2023$FEFIN <- as.Date(Informes_setiembre_2023$FEFIN, origin = "1899-12-30")
  #Informes_setiembre_2023$FEFIN <- format(Informes_setiembre_2023$FEFIN, format = "%d/%m/%Y")

Informes_setiembre_2023_1 <- Informes_setiembre_2023 %>%
  filter(FECHA_INFORME > as.Date("2023-06-30") & FECHA_INFORME < as.Date("2023-10-01")) %>%
  select(c(EXPEDIENTE,INFORME,FECHA_INFORME,ADMIN,UF,COORDINACION,TIPOSUP_INAPS,OBLIG_CUMPL,
           DOC_CIERRE,NRO_DOC_CIERRE,FECHA_DOC_CIERRE,FEFIN)) %>%
  mutate(MES_DE_APROBACION_DEL_INFORME = month(FECHA_INFORME,
                                               label = TRUE, abbr = FALSE)) %>%
  mutate(MES_DE_SUPERVISION = month(FEFIN,
                                    label = TRUE, abbr = FALSE)) %>%
  mutate(NRO_CARTA = ifelse(str_detect(DOC_CIERRE, "CARTA"),NRO_DOC_CIERRE, "")) %>%
  mutate(NRO_OFICIO = ifelse(str_detect(DOC_CIERRE, "OFICIO"), NRO_DOC_CIERRE, "")) %>%
  mutate(NRO_MEMORANDO = ifelse(str_detect(DOC_CIERRE, "MEMORANDO"), NRO_DOC_CIERRE, "")) %>%
  select(-c(NRO_DOC_CIERRE, FEFIN)) %>%
  relocate(MES_DE_APROBACION_DEL_INFORME, .after = FECHA_INFORME) %>%
  relocate(NRO_MEMORANDO, .after = DOC_CIERRE) %>%
  relocate(NRO_CARTA, .after = NRO_MEMORANDO) %>%
  relocate(NRO_OFICIO, .after = NRO_CARTA) %>%
  relocate(MES_DE_SUPERVISION, .after = COORDINACION) %>%
  rename(NRO_DE_EXPEDIENTE = EXPEDIENTE, NRO_INFORME = INFORME, ADMINISTRADO = ADMIN, UNIDAD_FISCALIZABLE = UF,
         TIPO_DE_SUPERVISION = TIPOSUP_INAPS, CANTIDAD_OBLIGACIONES = OBLIG_CUMPL, DOCUMENTO_DE_CIERRE = DOC_CIERRE,
        FECHA_DE_EMISION_DE_CARTA = FECHA_DOC_CIERRE)

  Informes_setiembre_2023_1$CANTIDAD_OBLIGACIONES <- as.numeric(Informes_setiembre_2023_1$CANTIDAD_OBLIGACIONES)
  

-----------------------------------------------------------------------------------------------------------------------------------------------
setwd("C:/Users/jquichiz/Desktop/OEFA/2023/Datos abiertos/Actualizaciones/Actualizacion II trimestre 2023/Cap II Supervision Ambiental/Informes de supervision")

Informes_II_semestre_2023 <- read.xlsx("Informes de supervisión 2019 - 2023.xlsx", sheet = "INFORMES")

 Informes_II_semestre_2023$FECHA_INFORME <- as.numeric(Informes_II_semestre_2023$FECHA_INFORME)
 Informes_II_semestre_2023$FECHA_INFORME <- as.Date(Informes_II_semestre_2023$FECHA_INFORME, origin = "1899-12-30")
 #Informes_II_semestre_2023$FECHA_INFORME <- format(Informes_II_semestre_2023$FECHA_INFORME, format = "%d/%m/%Y")
 
 Informes_II_semestre_2023$FECHA.DE.EMISION.DE.CARTA <- as.numeric(Informes_II_semestre_2023$FECHA.DE.EMISION.DE.CARTA)
 Informes_II_semestre_2023$FECHA.DE.EMISION.DE.CARTA <- as.Date(Informes_II_semestre_2023$FECHA.DE.EMISION.DE.CARTA, origin = "1899-12-30")
 #Informes_II_semestre_2023$FECHA.DE.EMISION.DE.CARTA <- format(Informes_II_semestre_2023$FECHA.DE.EMISION.DE.CARTA, format = "%d/%m/%Y")
 
Informes_II_semestre_2023_1 <- Informes_II_semestre_2023 %>%
rename(NRO_DE_EXPEDIENTE = NRO.DE.EXPEDIENTE, NRO_INFORME = NRO.INFORME, MES_DE_APROBACION_DEL_INFORME = MES.DE.APROBACIÓN.DEL.INFORME, 
       UNIDAD_FISCALIZABLE = UNIDAD.FISCALIZABLE, MES_DE_SUPERVISION = MES.DE.SUPERVISIÓN, TIPO_DE_SUPERVISION = TIPO.DE.SUPERVISIÓN, 
       CANTIDAD_OBLIGACIONES = CANTIDAD.OBLIGACIONES, DOCUMENTO_DE_CIERRE = DOCUMENTO.DE.CIERRE, 
       NRO_MEMORANDO = NRO.MEMORANDO, NRO_CARTA = NRO.CARTA, NRO_OFICIO = NRO.OFICIO, 
       FECHA_DE_EMISION_DE_CARTA = FECHA.DE.EMISION.DE.CARTA, COORDINACION = COORDINACIÓN)

------------------------------------------------------------------------------------------------------------------------------------------------
  
#Agregar la informacion adicional al II semestre

Informes_de_supervision_2019_2023 <- full_join(Informes_II_semestre_2023_1, Informes_setiembre_2023_1) %>%
  arrange(desc(FECHA_INFORME))

setwd("C:/Users/jquichiz/Desktop/OEFA/2023/Datos abiertos/Actualizaciones/Scripts")
write.xlsx(Informes_de_supervision_2019_2023, file = "Informes_de_supervisión_2019_2023.xlsx", asTable = TRUE)







  
  
  

  

  
  
  
  