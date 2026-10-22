# INE_ICV_AnadirMetadatos.R
#
# Añade a INE_ICV.catalogo.xlsx columnas derivadas del CONTENIDO de cada
# tabla ya descargada en DATA/<t>.px (lectura en local, sin red -- no es la
# descarga masiva/pivote que se dejó pendiente, solo metadatos):
#
#   ambito_geografico  nombre de la dimension STUB que representa geografia
#                      ("Comunidad autónoma", "Ámbito geográfico"...) o
#                      "Nacional" si la tabla no tiene desagregacion
#                      geografica (viene desagregada por sexo/edad/etc.)
#   n_regiones         nº de valores distintos de esa dimension geografica
#                      (1 si es "Nacional")
#   frecuencia         detectada a partir del FORMATO de los valores de la
#                      dimension HEADING ("Periodo"): Y anual, Q trimestral,
#                      M mensual, ? si no se reconoce (mismo criterio que
#                      INE_TableDatasetGeneration.Rmd de AEDV)
#   n_periodos         nº de fechas distintas de "Periodo"
#   fecha_inicio       fecha (Date) del periodo mas antiguo de "Periodo"
#   fecha_fin          fecha (Date) del periodo mas reciente de "Periodo"
#                      -- mismo patron que fecha_inicio/fecha_fin en
#                      IBESTAT.indicators.xlsx (ver CLAUDE.md)
#   n_obs              nº de filas de la tabla ya en formato largo/tidy
#                      (as.data.frame(px) tras quitar NA) -- equivalente al
#                      N.obs de AEDV
#
# Reanudable: se puede relanzar en cualquier momento, siempre recalcula todas
# las columnas para todas las filas del catalogo a partir de los .px locales
# (rapido, sin red).

library(pxR)
library(dplyr)
library(stringr)
library(openxlsx)

CATALOGO_XLSX <- "INE_ICV.catalogo.xlsx"
DATA_DIR <- "DATA"

catalogo <- openxlsx::read.xlsx(CATALOGO_XLSX, sheet = "catalogo")

GEO_DIMS <- c("Comunidad autónoma", "Ámbito geográfico")

detecta_frecuencia <- function(periodos) {
  p <- unique(as.character(periodos))
  p <- p[!is.na(p)]
  if (length(p) == 0) return(NA_character_)
  if (all(str_detect(p, "^[0-9]{4}$"))) return("Y")
  if (all(str_detect(p, "^[0-9]{4}[QT][0-9]$"))) return("Q")
  if (all(str_detect(p, "^[0-9]{4}M[0-9]{2}$"))) return("M")
  "?"
}

# Convierte un valor de "Periodo" a Date (inicio del periodo). Soporta los
# mismos formatos que detecta_frecuencia(); si el formato no se reconoce
# devuelve NA (no se adivina).
parsear_fecha <- function(periodo) {
  p <- as.character(periodo)
  out <- rep(as.Date(NA), length(p))
  es_y <- str_detect(p, "^[0-9]{4}$")
  out[es_y] <- as.Date(paste0(p[es_y], "-01-01"))
  es_q <- str_detect(p, "^[0-9]{4}[QT][0-9]$")
  if (any(es_q)) {
    anyo <- str_sub(p[es_q], 1, 4)
    trim <- as.integer(str_sub(p[es_q], 6, 6))
    out[es_q] <- as.Date(sprintf("%s-%02d-01", anyo, (trim - 1) * 3 + 1))
  }
  es_m <- str_detect(p, "^[0-9]{4}M[0-9]{2}$")
  if (any(es_m)) {
    anyo <- str_sub(p[es_m], 1, 4)
    mes <- str_sub(p[es_m], 6, 7)
    out[es_m] <- as.Date(sprintf("%s-%s-01", anyo, mes))
  }
  out
}

n_tablas <- n_distinct(catalogo$tabla_t)
tablas_unicas <- catalogo %>% distinct(tabla_t)

resultado <- vector("list", nrow(tablas_unicas))
for (i in seq_len(nrow(tablas_unicas))) {
  t_id <- tablas_unicas$tabla_t[i]
  fn <- file.path(DATA_DIR, paste0(t_id, ".px"))

  if (!file.exists(fn)) {
    resultado[[i]] <- tibble(
      tabla_t = t_id, ambito_geografico = NA_character_, n_regiones = NA_integer_,
      frecuencia = NA_character_, n_periodos = NA_integer_,
      fecha_inicio = as.Date(NA), fecha_fin = as.Date(NA), n_obs = NA_integer_
    )
    next
  }

  px <- tryCatch(read.px(fn), error = function(e) NULL)
  if (is.null(px)) {
    resultado[[i]] <- tibble(
      tabla_t = t_id, ambito_geografico = "ERROR", n_regiones = NA_integer_,
      frecuencia = NA_character_, n_periodos = NA_integer_,
      fecha_inicio = as.Date(NA), fecha_fin = as.Date(NA), n_obs = NA_integer_
    )
    next
  }

  datos <- as.data.frame(px) %>% na.omit()

  cols <- setdiff(names(datos), c("Periodo", "value"))
  # nombres de columna en as.data.frame(px) sustituyen espacios por "."
  geo_col <- cols[str_replace_all(cols, "\\.", " ") %in% GEO_DIMS]

  if (length(geo_col) >= 1) {
    ambito <- str_replace_all(geo_col[1], "\\.", " ")
    n_reg <- n_distinct(datos[[geo_col[1]]])
  } else {
    ambito <- "Nacional"
    n_reg <- 1L
  }

  periodos <- if ("Periodo" %in% names(datos)) datos[["Periodo"]] else NA
  freq <- detecta_frecuencia(periodos)
  n_per <- n_distinct(periodos[!is.na(periodos)])
  fechas <- parsear_fecha(periodos)
  fechas <- fechas[!is.na(fechas)]

  resultado[[i]] <- tibble(
    tabla_t = t_id,
    ambito_geografico = ambito,
    n_regiones = as.integer(n_reg),
    frecuencia = freq,
    n_periodos = as.integer(n_per),
    fecha_inicio = if (length(fechas)) min(fechas) else as.Date(NA),
    fecha_fin = if (length(fechas)) max(fechas) else as.Date(NA),
    n_obs = as.integer(nrow(datos))
  )
}

metadatos <- bind_rows(resultado)

catalogo_ampliado <- catalogo %>%
  select(-any_of(c("ambito_geografico", "n_regiones", "frecuencia", "n_periodos", "fecha_inicio", "fecha_fin", "n_obs"))) %>%
  left_join(metadatos, by = "tabla_t")

write.xlsx(catalogo_ampliado, CATALOGO_XLSX, sheetName = "catalogo", overwrite = TRUE)

cat("Catalogo actualizado:", nrow(catalogo_ampliado), "filas,", n_tablas, "tablas distintas\n")
cat("Resumen frecuencia:\n")
print(table(metadatos$frecuencia, useNA = "always"))
cat("Resumen ambito_geografico:\n")
print(table(metadatos$ambito_geografico, useNA = "always"))
cat("Rango fecha_inicio:", format(min(metadatos$fecha_inicio, na.rm = TRUE)), "-", format(max(metadatos$fecha_inicio, na.rm = TRUE)), "\n")
cat("Rango fecha_fin:", format(min(metadatos$fecha_fin, na.rm = TRUE)), "-", format(max(metadatos$fecha_fin, na.rm = TRUE)), "\n")
cat("Filas sin fecha (formato Periodo no reconocido):", sum(is.na(metadatos$fecha_inicio)), "\n")
