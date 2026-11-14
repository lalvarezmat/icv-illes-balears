# PIB_CCAA_VAB_Ramas.R
#
# Peso de cada rama de actividad en el valor añadido bruto (VAB) de cada
# comunidad autonoma: % del VAB total a precios corrientes, 2000-ultimo año.
# Para el estudio de los archipielagos
# (INE_ICV_CCAA/INE_ICV_CCAA_Insulares_Consolidar.R).
#
# Fuente: el mismo .px de Contabilidad Regional de España (INE) que usa
# PIB_CCAA_DataExtraction.R, PIB_precios_corrientes_73423.px (matriz 01001,
# "PIB a precios de mercado y VAB a precios basicos por ramas de actividad.
# Precios corrientes"). Si no existe, ejecutar antes PIB_CCAA_DataExtraction.R,
# que lo descarga. No se vuelve a descargar aqui para que el VAB y el PIB per
# capita del informe salgan de la misma edicion de la Contabilidad Regional.
#
# Ramas: la agregacion A10 de la CNAE-2009 que publica el INE por CCAA (10
# ramas que suman el VAB total). La fila "C. - De las cuales: Industria
# manufacturera" es una PARTE de B_E y no se usa (contaria dos veces).
# Comprobado: las 10 ramas suman el VAB total en todas las CCAA y años.
#
# TRAMPA (la misma que en PIB_CCAA_DataExtraction.R): la fila del PIB sale
# duplicada literalmente en la tabla; aqui no se usa, pero se aplica
# distinct() por si alguna rama tambien lo estuviera.
#
# Genera en esta carpeta:
#   PIB_CCAA_VAB_Ramas.xlsx   hoja "datos": codigo_ccaa, ccaa, anio,
#                             estado_dato, rama_codigo, rama, vab_miles_eur,
#                             pct_vab

library(pxR)
library(dplyr)
library(openxlsx)

CARPETA <- "."  # ruta relativa: ejecutar desde la carpeta de este script
FN_PX <- file.path(CARPETA, "PIB_precios_corrientes_73423.px")
if (!file.exists(FN_PX)) stop("Falta ", FN_PX, ": ejecutar antes PIB_CCAA_DataExtraction.R")

RAMAS_A10 <- c("A", "B_E", "F", "G_I", "J", "K", "L", "M_N", "O_Q", "R_U")
VAB_TOTAL <- "Valor añadido bruto total"

df <- suppressWarnings(as.data.frame(read.px(FN_PX)))
names(df)[grepl("^Comunidades", names(df))] <- "ccaa_px"
names(df)[grepl("^Ramas", names(df))] <- "rama"

df <- df %>%
  mutate(rama = as.character(rama), ccaa_px = as.character(ccaa_px),
         rama_codigo = ifelse(rama == VAB_TOTAL, "TOTAL", sub("\\. .*$", "", rama))) %>%
  filter(rama_codigo %in% c(RAMAS_A10, "TOTAL")) %>%
  distinct(ccaa_px, Periodo, rama_codigo, rama, value) %>%
  mutate(
    codigo_ccaa = ifelse(ccaa_px == "Total Nacional", "00", substr(ccaa_px, 1, 2)),
    ccaa = ifelse(ccaa_px == "Total Nacional", "Total Nacional", trimws(substr(ccaa_px, 4, nchar(ccaa_px)))),
    anio = as.integer(substr(as.character(Periodo), 1, 4)),
    estado_dato = case_when(
      grepl("\\(A\\)", Periodo) ~ "Avance",
      grepl("\\(P\\)", Periodo) ~ "Provisional",
      TRUE ~ "Definitivo"
    )
  )

stopifnot(setequal(unique(df$rama_codigo), c(RAMAS_A10, "TOTAL")),
          !any(duplicated(df[, c("codigo_ccaa", "anio", "rama_codigo")])))

total <- df %>% filter(rama_codigo == "TOTAL") %>% select(codigo_ccaa, anio, vab_total = value)

vab <- df %>%
  filter(rama_codigo != "TOTAL") %>%
  left_join(total, by = c("codigo_ccaa", "anio")) %>%
  mutate(pct_vab = 100 * value / vab_total) %>%
  select(codigo_ccaa, ccaa, anio, estado_dato, rama_codigo, rama,
         vab_miles_eur = value, pct_vab) %>%
  arrange(codigo_ccaa, anio, match(rama_codigo, RAMAS_A10))

# Las 10 ramas suman el VAB total (tolerancia: 0,05 puntos por redondeo)
suma <- vab %>% group_by(codigo_ccaa, anio) %>% summarise(s = sum(pct_vab), .groups = "drop")
stopifnot(all(abs(suma$s - 100) < 0.05), n_distinct(vab$codigo_ccaa) == 20)

message("VAB por ramas: ", min(vab$anio), "-", max(vab$anio), ", ",
        n_distinct(vab$rama_codigo), " ramas, ", n_distinct(vab$codigo_ccaa), " geografias")

wb <- createWorkbook()
addWorksheet(wb, "datos")
writeData(wb, "datos", vab)
addWorksheet(wb, "metadatos")
writeData(wb, "metadatos", data.frame(
  campo = c("fuente", "fichero", "pct_vab", "ramas"),
  valor = c("INE, Contabilidad Regional de España: VAB a precios basicos por ramas de actividad, precios corrientes",
            "PIB_precios_corrientes_73423.px (https://www.ine.es/jaxi/files/tpx/es/px/73423.px)",
            "100 * VAB de la rama / VAB total de la comunidad, a precios corrientes",
            "Agregacion A10 de la CNAE-2009 (10 ramas que suman el VAB total)")
))
saveWorkbook(wb, file.path(CARPETA, "PIB_CCAA_VAB_Ramas.xlsx"), overwrite = TRUE)
