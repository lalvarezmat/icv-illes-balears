# POBLACION_CCAA_DataExtraction.R
#
# Poblacion residente a 1 de enero por comunidad autonoma, 1971-ultimo año,
# de la Estadistica Continua de Poblacion (ECP) del INE. La usa el estudio de
# los archipielagos (INE_ICV_CCAA/INE_ICV_CCAA_Insulares_Consolidar.R) para:
#   - el indicador "crecimiento de la poblacion (% anual)";
#   - pasar a "por habitante" los indicadores que llegan en numeros absolutos
#     (viviendas iniciadas, compraventas por no residentes, viviendas
#     turisticas, pernoctaciones en plataformas).
#
# Fuente: tabla tempus3 t=56940 "Poblacion residente por fecha, sexo y edad
# (desde 1971)", via ineapir. Es la serie RETROPOLADA por el INE con la
# metodologia de la ECP: homogenea en todo el periodo, sin el salto de 2021
# que obliga a empalmar Padron (t=2853) y ECP (t=67988) en
# ../PIB_CCAA/PIB_CCAA_DataExtraction.R. Por eso aqui no se empalma nada.
#
# La tabla trae fechas semestrales (1 de enero y 1 de julio) hasta 2021 y
# trimestrales despues; solo se guarda el 1 de enero de cada año.
#
# TRAMPA: sin filtro, la tabla es enorme (100 edades x 3 sexos x 20
# geografias x 117 fechas) y el servidor del INE corta la conexion
# ("schannel: server closed abruptly"). Se filtra en la propia peticion por
# Sexo = Total y Edad = Todas las edades, con los Id de valor que da
# get_metadata_table_varval() (el filtro por texto no funciona).
#
# Genera en esta carpeta:
#   POBLACION_CCAA.xlsx   hoja "datos": codigo_ccaa, ccaa, anio, poblacion
#                         (1 de enero), crecimiento_pct (de 1 de enero del
#                         año al 1 de enero del siguiente; NA en el ultimo)

Sys.setenv(R_LIBCURL_SSL_REVOKE_BEST_EFFORT = "TRUE")

library(ineapir)
library(dplyr)
library(openxlsx)

CARPETA <- "."  # ruta relativa: ejecutar desde la carpeta de este script
ID_TABLA <- 56940

get_ine <- function(f, ..., n = 5, sleep = 3) {
  for (i in seq_len(n)) {
    r <- tryCatch(f(...), error = function(e) {
      message("  fallo intento ", i, ": ", conditionMessage(e))
      NULL
    })
    if (!is.null(r)) return(r)
    Sys.sleep(sleep * i)
  }
  stop("La API del INE no respondio tras ", n, " intentos")
}

# Id de los valores "Total" (sexo) y "Todas las edades" para filtrar en la peticion
varval <- get_ine(get_metadata_table_varval, idTable = ID_TABLA)
filtro_ids <- varval %>% filter(Nombre %in% c("Total", "Todas las edades"))
stopifnot(nrow(filtro_ids) == 2)
filtro <- setNames(as.list(as.character(filtro_ids$Id)), as.character(filtro_ids$Fk_Variable))

pob <- get_ine(get_data_table, idTable = ID_TABLA, filter = filtro, nlast = 500,
               unnest = TRUE, metanames = TRUE, metacodes = TRUE, tip = "AM")

pob <- pob %>%
  filter(substr(Fecha, 6, 10) == "01-01") %>%
  transmute(
    codigo_ccaa = `Comunidades.y.ciudades.autónomas.Codigo`,
    ccaa = `Comunidades.y.ciudades.autónomas`,
    anio = as.integer(substr(Fecha, 1, 4)),
    poblacion = Valor
  ) %>%
  mutate(codigo_ccaa = ifelse(ccaa == "Total Nacional", "00", codigo_ccaa)) %>%
  arrange(codigo_ccaa, anio)

# 20 geografias (Total Nacional + 17 CCAA + Ceuta + Melilla), sin huecos
stopifnot(n_distinct(pob$codigo_ccaa) == 20,
          !any(duplicated(pob[, c("codigo_ccaa", "anio")])),
          all(table(pob$codigo_ccaa) == n_distinct(pob$anio)))
# La suma de las 19 CCAA y ciudades debe dar el Total Nacional
suma <- pob %>% filter(codigo_ccaa != "00") %>% group_by(anio) %>% summarise(s = sum(poblacion))
tot  <- pob %>% filter(codigo_ccaa == "00") %>% select(anio, poblacion)
stopifnot(all(abs(suma$s - tot$poblacion[match(suma$anio, tot$anio)]) <= 5))

pob <- pob %>%
  group_by(codigo_ccaa) %>%
  mutate(crecimiento_pct = 100 * (lead(poblacion) / poblacion - 1)) %>%
  ungroup()

message("Poblacion a 1 de enero, ", min(pob$anio), "-", max(pob$anio), ": ",
        nrow(pob), " filas")

wb <- createWorkbook()
addWorksheet(wb, "datos")
writeData(wb, "datos", pob)
addWorksheet(wb, "metadatos")
writeData(wb, "metadatos", data.frame(
  campo = c("fuente", "tabla", "poblacion", "crecimiento_pct", "descarga"),
  valor = c("INE, Estadistica Continua de Poblacion (serie retropolada desde 1971)",
            paste0("https://www.ine.es/jaxiT3/Tabla.htm?t=", ID_TABLA),
            "Poblacion residente a 1 de enero del año",
            "100 * (poblacion 1 de enero del año siguiente / poblacion 1 de enero del año - 1): crecimiento durante el año",
            format(Sys.Date()))
))
saveWorkbook(wb, file.path(CARPETA, "POBLACION_CCAA.xlsx"), overwrite = TRUE)
