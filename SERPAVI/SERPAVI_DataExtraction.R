# Descarga la base de datos completa del Sistema Estatal de Referencia del
# Precio del Alquiler de Vivienda (SERPAVI, Ministerio de Vivienda y Agenda
# Urbana) y pivota su hoja "CCAA" de formato ANCHO (una fila por CCAA, ~280
# columnas = bloque de variables repetido una vez por año, sufijo "_AA" de
# 2 digitos) a formato LARGO (una fila por CCAA x año x tipo de vivienda).
#
# OJO: el visor publico https://serpavi.mivau.gob.es/ NO es esta fuente --
# es una calculadora por vivienda individual (referencia catastral o
# direccion exacta), no una serie estadistica. La serie esta en el Excel
# masivo enlazado desde https://www.mivau.gob.es/vivienda/alquila-bien-es-tu-derecho/serpavi
# ("BD Sistema Estatal Indices de Alquiler de Vivienda"), verificado en vivo
# el 2026-09-22. La URL lleva la fecha de publicacion en el nombre
# (2026-03_09): si cambia y da 404, volver a sacar el link "Documento Xlsx"
# de esa pagina.
#
# Variables por año (confirmadas contra la hoja "Metadatos" del propio
# fichero; el diccionario usa el prefijo "ALQM2mes_LV_..." pero las columnas
# reales de datos usan "ALQM2_LV_..." sin "mes" -- pequeña discrepancia de
# nombre entre diccionario y datos, no afecta al valor):
#   BI_ALVHEPCO_TVC/TVU_AA        nº de contratos de alquiler considerados
#   ALQM2_LV_{M,25,75}_VC/VU_AA   precio alquiler EUROS/m2/MES (mediana, p25, p75)
#   ALQTBID12_{M,25,75}_VC/VU_AA  importe mensual de TODO el inmueble (no por m2)
#   SLVM2_{M,25,75}_VC/VU_AA      superficie media (m2)
# VC = Vivienda Colectiva (piso), VU = Vivienda Unifamiliar o Rural (chalet).
# La propia fuente ya usa mediana/percentiles en vez de media -- coherente
# con el enfoque de estimadores robustos que ya usa este proyecto.
#
# La hoja "CCAA" NO trae fila de total nacional (19 filas, codigos 01-19,
# mismo esquema de codigo que INE_ICV_CCAA_Consolidado.xlsx: 01 Andalucia
# .. 17 La Rioja, 18 Ceuta, 19 Melilla). Cobertura temporal: 2011-2024.
#
# Genera en esta carpeta:
#   - BD_SERPAVI_2011-2024.xlsx  fichero de origen tal cual se descarga
#   - SERPAVI_CCAA.xlsx          hoja "datos" en formato largo + "metadatos"
#
# No es reanudable por tramos (un unico fichero) -- vuelve a descargar y
# sobrescribe cada vez que se lanza.

Sys.setenv(R_LIBCURL_SSL_REVOKE_BEST_EFFORT = "TRUE")

library(httr)
library(openxlsx)
library(dplyr)

httr::set_config(httr::timeout(600))

CARPETA <- "."  # ruta relativa: ejecutar desde la carpeta de este script
UA <- httr::user_agent("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0 Safari/537.36")
URL <- "https://cdn.mivau.gob.es/portal-web-mivau/vivienda/serpavi/2026-03_09_bd_SERPAVI_2011-2024%20-%20DEFINITIVO%20WEB_v2.xlsx"
FICHERO_ORIGEN <- file.path(CARPETA, "BD_SERPAVI_2011-2024.xlsx")

cat("Descargando SERPAVI (71 MB, puede tardar)...\n")
resp <- httr::GET(URL, UA, httr::write_disk(FICHERO_ORIGEN, overwrite = TRUE))
httr::stop_for_status(resp)
cat("Descargado:", file.size(FICHERO_ORIGEN), "bytes\n")

ccaa <- read.xlsx(FICHERO_ORIGEN, sheet = "CCAA")
# La hoja trae una fila 20 fantasma con CCAA/LITCCAA a NA (algun resto de
# formato en una celda lejana de esa fila hace que openxlsx la incluya en el
# rango detectado) -- se descarta, las 19 CCAA reales tienen codigo no-NA.
ccaa <- ccaa[!is.na(ccaa$CCAA), ]
stopifnot("Se esperaban 19 CCAA (codigos 01-19)" = nrow(ccaa) == 19)

anios <- sprintf("%02d", 11:24)

extraer_anio <- function(sufijo) {
  col <- function(nombre) ccaa[[nombre]]
  bind_rows(
    data.frame(
      ccaa_codigo = ccaa$CCAA, ccaa_nombre = ccaa$LITCCAA,
      anio = 2000L + as.integer(sufijo), tipo_vivienda = "VC",
      n_contratos = col(paste0("BI_ALVHEPCO_TVC_", sufijo)),
      alquiler_m2_mediana = col(paste0("ALQM2_LV_M_VC_", sufijo)),
      alquiler_m2_p25 = col(paste0("ALQM2_LV_25_VC_", sufijo)),
      alquiler_m2_p75 = col(paste0("ALQM2_LV_75_VC_", sufijo)),
      importe_mediana = col(paste0("ALQTBID12_M_VC_", sufijo)),
      importe_p25 = col(paste0("ALQTBID12_25_VC_", sufijo)),
      importe_p75 = col(paste0("ALQTBID12_75_VC_", sufijo)),
      superficie_mediana = col(paste0("SLVM2_M_VC_", sufijo)),
      superficie_p25 = col(paste0("SLVM2_25_VC_", sufijo)),
      superficie_p75 = col(paste0("SLVM2_75_VC_", sufijo)),
      stringsAsFactors = FALSE
    ),
    data.frame(
      ccaa_codigo = ccaa$CCAA, ccaa_nombre = ccaa$LITCCAA,
      anio = 2000L + as.integer(sufijo), tipo_vivienda = "VU",
      n_contratos = col(paste0("BI_ALVHEPCO_TVU_", sufijo)),
      alquiler_m2_mediana = col(paste0("ALQM2_LV_M_VU_", sufijo)),
      alquiler_m2_p25 = col(paste0("ALQM2_LV_25_VU_", sufijo)),
      alquiler_m2_p75 = col(paste0("ALQM2_LV_75_VU_", sufijo)),
      importe_mediana = col(paste0("ALQTBID12_M_VU_", sufijo)),
      importe_p25 = col(paste0("ALQTBID12_25_VU_", sufijo)),
      importe_p75 = col(paste0("ALQTBID12_75_VU_", sufijo)),
      superficie_mediana = col(paste0("SLVM2_M_VU_", sufijo)),
      superficie_p25 = col(paste0("SLVM2_25_VU_", sufijo)),
      superficie_p75 = col(paste0("SLVM2_75_VU_", sufijo)),
      stringsAsFactors = FALSE
    )
  )
}

datos <- bind_rows(lapply(anios, extraer_anio)) %>% arrange(anio, ccaa_codigo, tipo_vivienda)

metadatos <- data.frame(
  campo = c("fuente", "url", "descripcion", "cobertura_geografica", "cobertura_temporal",
            "tipo_vivienda_VC", "tipo_vivienda_VU", "fecha_descarga", "n_filas"),
  valor = c("Ministerio de Vivienda y Agenda Urbana - SERPAVI (explotacion de fuentes tributarias)",
            URL, "Precio de alquiler EUROS/m2/mes por CCAA y año, mediana y percentiles 25/75",
            "España (19 CCAA, sin fila de total nacional)", "2011-2024",
            "Vivienda Colectiva (piso)", "Vivienda Unifamiliar o Rural (chalet)",
            format(Sys.time(), "%Y-%m-%d %H:%M:%S"), as.character(nrow(datos))),
  stringsAsFactors = FALSE
)

wb <- createWorkbook()
addWorksheet(wb, "datos"); writeData(wb, "datos", datos)
addWorksheet(wb, "metadatos"); writeData(wb, "metadatos", metadatos)
saveWorkbook(wb, file.path(CARPETA, "SERPAVI_CCAA.xlsx"), overwrite = TRUE)

cat(nrow(datos), "filas ->", file.path(CARPETA, "SERPAVI_CCAA.xlsx"), "\n")
