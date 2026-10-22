# INE_ICV_CCAA_ConstruirCatalogo.R
#
# Consolida en un unico catalogo (INE_ICV_CCAA.catalogo.xlsx) los datos que
# ya estan repartidos en los 42 <slug>_METADATOS.xlsx de este directorio
# (uno por hoja "resumen"/"medida"/"geografia" -- no se vuelve a leer ningun
# .px, es una relectura local barata de lo que generó
# INE_ICV_CCAA_Procesar.R). Una fila por tabla, con:
#
#   fecha_inicio/fecha_fin  AÑO (entero), no Date: las 42 tablas son anuales
#                      (frecuencia == "Y" en las 42, comprobado el
#                      2026-09-21 con table(catalogo$frecuencia) antes de
#                      cambiar esto) y guardar un Date con dia/mes fijos a
#                      01-01 solo añadia ruido.
#
#   n_atributos       nº de valores medidos por region/año, es decir nº de
#                      hojas "datos*" del <slug>.xlsx (1 en la mayoria de
#                      tablas; >1 solo cuando el indicador tiene categoria de
#                      respuesta propia, ver cabecera de
#                      INE_ICV_CCAA_Procesar.R -- nunca una desagregacion
#                      demografica distinta)
#   explicacion       campo "tabla" de la hoja "resumen" (titulo completo de
#                      la tabla en la fuente INE, incluida la categoria de
#                      respuesta cuando la tabla tiene mas de una hoja de
#                      datos) + la unidad entre parentesis. Se usa "tabla" y
#                      no "contenido" porque "contenido" viene truncado en 2
#                      de los 42 ficheros ("Índice de masa cor", "Asalariados
#                      tem" -- truncamiento ya presente en el _METADATOS.xlsx
#                      de origen, no introducido aqui) mientras que "tabla"
#                      esta completo en los 42.
#
# Reanudable/regenerable en cualquier momento: no llama a ninguna red, solo
# relee los METADATOS ya presentes en el directorio.

library(openxlsx)
library(dplyr)

DIR <- "."
metadatos <- list.files(DIR, pattern = "_METADATOS\\.xlsx$", full.names = FALSE)

fila_catalogo <- function(mf) {
  slug <- sub("_METADATOS\\.xlsx$", "", mf)
  datos_xlsx <- paste0(slug, ".xlsx")

  resumen <- read.xlsx(mf, sheet = "resumen")
  medida <- read.xlsx(mf, sheet = "medida")
  geografia <- read.xlsx(mf, sheet = "geografia")

  val <- function(tabla, campo) {
    v <- tabla$valor[tabla$campo == campo]
    if (length(v) == 0 || is.na(v)) "" else v
  }

  hojas_datos <- grep("^datos", getSheetNames(datos_xlsx), value = TRUE)
  n_regiones <- nrow(geografia)
  n_atributos <- length(hojas_datos)

  unidad <- val(medida, "unidad")
  explicacion <- paste0(val(resumen, "tabla"), " (unidad: ", unidad, ")")

  data.frame(
    fichero = slug,
    tema = val(resumen, "tema"),
    dimension = val(resumen, "dimension"),
    subdimension = val(resumen, "subdimension"),
    explicacion = explicacion,
    unidad = unidad,
    n_atributos = n_atributos,
    n_regiones = n_regiones,
    frecuencia = val(resumen, "frecuencia"),
    fecha_inicio = as.integer(format(as.Date(val(resumen, "fecha_inicio")), "%Y")),
    fecha_fin = as.integer(format(as.Date(val(resumen, "fecha_fin")), "%Y")),
    n_periodos = as.integer(val(resumen, "n_periodos")),
    n_obs = as.integer(val(resumen, "n_obs")),
    tabla_t = val(resumen, "tabla_t"),
    url_tabla = val(resumen, "url_tabla"),
    stringsAsFactors = FALSE
  )
}

catalogo <- bind_rows(lapply(metadatos, fila_catalogo)) %>%
  arrange(desc(fecha_fin), desc(n_periodos))

stopifnot(all(catalogo$frecuencia == "Y")) # si deja de ser anual, fecha_inicio/fecha_fin ya no pueden ser solo el año

wb <- createWorkbook()
addWorksheet(wb, "catalogo")
writeData(wb, "catalogo", catalogo)
setColWidths(wb, "catalogo", cols = 1:ncol(catalogo), widths = "auto")
freezePane(wb, "catalogo", firstRow = TRUE)
saveWorkbook(wb, "INE_ICV_CCAA.catalogo.xlsx", overwrite = TRUE)

cat(nrow(catalogo), "tablas escritas en INE_ICV_CCAA.catalogo.xlsx\n")
