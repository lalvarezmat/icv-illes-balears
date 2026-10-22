# IPC_CCAA_DataExtraction.R
#
# Indice de Precios de Consumo (IPC) general, MEDIA ANUAL, por comunidad
# autonoma, para deflactar en INE_ICV_CCAA_Consolidar.R los indicadores en
# euros corrientes (renta media y mediana, precio de compraventa y de alquiler)
# y pasarlos a euros constantes de 2020, la misma base que el PIB per capita
# real. Peticion explicita del usuario 2026-09-25: deflactar con el IPC de CADA
# comunidad, no solo con el nacional.
#
# Fuente: INE, tabla tempus3 t=24078, "Indice general por comunidades
# autonomas. Series desde enero de 1978" (mensual). Se pide por la API JSON
# del INE (servicios.ine.es/wstempus), sin clave. La tabla trae 20 series:
# "Nacional" + las 17 CCAA + Ceuta y Melilla.
#
# Media anual = media simple de los 12 meses del año. Solo se guardan años
# completos (12 meses); el año en curso, incompleto, se descarta.
#
# Comprobado el 2026-09-25: la inflacion nacional 2007-2024 calculada asi
# (+38,93 %) coincide con la de las medias anuales que publica el INE ya hechas
# en la tabla t=50934 ("Indices nacionales. Medias anuales", base 2021: 2007 =
# 83,015 y 2024 = 115,333). Los NIVELES no coinciden porque la t=24078 viene en
# otra base (2021 vale 84,4), pero para deflactar solo importan los cocientes
# entre años. La tabla 50934 no tiene desglose por CCAA, por eso se usa la
# mensual.
#
# OJO: el IPC por CCAA mide cuanto suben los precios DENTRO de cada comunidad
# a lo largo del tiempo, no cuanto mas caro es vivir en una que en otra (no es
# una paridad de poder adquisitivo regional). Deflactar con el corrige la
# inflacion de cada comunidad, pero no las diferencias de nivel de precios
# entre comunidades.
#
# Genera en esta carpeta:
#   IPC_CCAA_Anual.xlsx   hoja "datos" (una fila por geografia x año:
#                         codigo_ccaa "00".."19", ccaa, anio, ipc, n_meses) +
#                         hoja "metadatos"

Sys.setenv(R_LIBCURL_SSL_REVOKE_BEST_EFFORT = "TRUE")
options(timeout = 120)

library(jsonlite)
library(dplyr)
library(openxlsx)

CARPETA <- "."  # ruta relativa: ejecutar desde la carpeta de este script
TABLA <- 24078
# Pese al titulo de la tabla ("desde enero de 1978"), la API solo devuelve esta
# serie desde enero de 2002 (comprobado 2026-09-25). Basta: lo mas antiguo que
# se deflacta es el precio de compraventa de 2007 (y la renta de la encuesta
# de 2008, que son ingresos de 2007).
ANIO_INICIO <- 2002L
URL <- sprintf("https://servicios.ine.es/wstempus/js/ES/DATOS_TABLA/%d?date=%d0101:%s",
               TABLA, ANIO_INICIO, format(Sys.Date(), "%Y%m%d"))

# Codificacion oficial de CCAA, la misma que usan el consolidado, REGISTRADORES,
# SERPAVI y PIB_CCAA ("00" = total nacional). Los nombres son los que devuelve
# la API en el campo "Nombre" de cada serie.
CODIGOS <- c(
  "Nacional" = "00", "Andalucía" = "01", "Aragón" = "02", "Asturias, Principado de" = "03",
  "Balears, Illes" = "04", "Canarias" = "05", "Cantabria" = "06", "Castilla y León" = "07",
  "Castilla - La Mancha" = "08", "Cataluña" = "09", "Comunitat Valenciana" = "10",
  "Extremadura" = "11", "Galicia" = "12", "Madrid, Comunidad de" = "13",
  "Murcia, Región de" = "14", "Navarra, Comunidad Foral de" = "15", "País Vasco" = "16",
  "Rioja, La" = "17", "Ceuta" = "18", "Melilla" = "19"
)

# Los fallos de red son transitorios: se reintenta y, si no hay manera, se
# para con error en vez de seguir con datos a medias.
leer_api <- function(url, n = 5, sleep = 3) {
  for (i in seq_len(n)) {
    res <- tryCatch(fromJSON(url), error = function(e) NULL)
    if (!is.null(res) && NROW(res) > 0) return(res)
    Sys.sleep(sleep * i)
  }
  stop("No se pudo leer ", url, " tras ", n, " intentos")
}

series <- leer_api(URL)

mensual <- bind_rows(lapply(seq_len(nrow(series)), function(i) {
  dd <- series$Data[[i]]
  data.frame(
    ccaa = sub(". Índice general. Índice. ", "", series$Nombre[i], fixed = TRUE),
    anio = as.integer(dd$Anyo),
    mes = as.integer(dd$FK_Periodo),
    ipc = as.numeric(dd$Valor),
    stringsAsFactors = FALSE
  )
}))

stopifnot("La API no devolvio las 20 geografias esperadas (o cambio el formato del nombre)" =
  setequal(unique(mensual$ccaa), names(CODIGOS)))
stopifnot(!any(duplicated(mensual[, c("ccaa", "anio", "mes")])), !anyNA(mensual$ipc))

datos <- mensual %>%
  group_by(ccaa, anio) %>%
  summarise(ipc = mean(ipc), n_meses = n(), .groups = "drop") %>%
  filter(n_meses == 12) %>%
  mutate(codigo_ccaa = unname(CODIGOS[ccaa])) %>%
  select(codigo_ccaa, ccaa, anio, ipc, n_meses) %>%
  arrange(codigo_ccaa, anio)

# Cordura: serie completa y sin huecos para las 20 geografias, y la inflacion
# nacional 2007-2024 igual a la de las medias anuales de la tabla t=50934
# (115,333 / 83,015, ver cabecera). Se compara el cociente, no el nivel, para
# que no falle si el INE cambia la base del indice.
anios <- range(datos$anio)
stopifnot(anios[1] == ANIO_INICIO,
          all(count(datos, codigo_ccaa)$n == anios[2] - anios[1] + 1))
nac <- datos %>% filter(codigo_ccaa == "00")
stopifnot("La inflacion nacional 2007-2024 no coincide con la tabla t=50934 del INE" =
  abs(nac$ipc[nac$anio == 2024] / nac$ipc[nac$anio == 2007] - 115.333 / 83.015) < 0.001)

metadatos <- data.frame(
  campo = c("fuente", "url", "metodo", "cobertura_anios", "fecha_extraccion", "n_filas", "advertencia"),
  valor = c(
    "INE - IPC, tabla tempus3 t=24078: Índice general por comunidades autónomas. Series desde enero de 1978 (mensual)",
    URL,
    "Media simple de los 12 índices mensuales de cada año; solo años completos",
    paste(anios, collapse = "-"),
    format(Sys.time(), "%Y-%m-%d %H:%M:%S"),
    as.character(nrow(datos)),
    "El IPC por CCAA mide la evolución de los precios dentro de cada comunidad, no las diferencias de nivel de precios entre comunidades"
  ),
  stringsAsFactors = FALSE
)

wb <- createWorkbook()
addWorksheet(wb, "datos")
writeData(wb, "datos", datos)
addWorksheet(wb, "metadatos")
writeData(wb, "metadatos", metadatos)
setColWidths(wb, "metadatos", cols = 1:2, widths = c(20, 120))
saveWorkbook(wb, file.path(CARPETA, "IPC_CCAA_Anual.xlsx"), overwrite = TRUE)

cat(nrow(datos), "filas (", length(unique(datos$codigo_ccaa)), "geografias,",
    paste(anios, collapse = "-"), ") en IPC_CCAA_Anual.xlsx\n")
