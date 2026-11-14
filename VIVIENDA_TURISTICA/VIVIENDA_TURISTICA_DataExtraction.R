# VIVIENDA_TURISTICA_DataExtraction.R
#
# Vivienda turistica (alquiler vacacional) por comunidad autonoma, de dos
# fuentes que miden cosas distintas, para el estudio de los archipielagos
# (INE_ICV_CCAA/INE_ICV_CCAA_Insulares_Consolidar.R):
#
#   1. OFERTA: viviendas turisticas y sus plazas (INE, estadistica
#      experimental "Medicion del numero de viviendas turisticas en España y
#      su capacidad", tabla tempus3 t=39364). El INE cuenta los anuncios de
#      las grandes plataformas (Airbnb, Booking, Vrbo), quitando duplicados.
#      Es una FOTO en una fecha: agosto de 2020 y despues dos al año (febrero
#      y agosto hasta 2024; mayo y noviembre desde noviembre de 2024). Se
#      guardan todas las fechas; el consolidado hace la media de cada año.
#      La variable geografica mezcla CCAA y provincias, y las uniprovinciales
#      salen dos veces con el mismo nombre: se filtra por los Id de nivel
#      CCAA (Fk_Variable 70) y Total Nacional (349) de los metadatos.
#
#   2. ACTIVIDAD: noches pasadas en alojamientos de corta estancia
#      reservados a traves de plataformas en linea (Airbnb, Booking, Expedia
#      y Tripadvisor), Eurostat, dataset tour_ce_omn12, anual 2018-ultimo
#      año, NUTS 2 (= CCAA en España). Indicador NGT_SP, residencia del
#      huesped TOTAL, mes TOTAL. Es la actividad realmente vendida, no la
#      oferta anunciada.
#
# Genera en esta carpeta:
#   VIVIENDA_TURISTICA_CCAA.xlsx
#     hoja "ine":      codigo_ccaa, ccaa, fecha (AAAA-MM-DD), anio, viviendas, plazas
#     hoja "eurostat": codigo_ccaa, ccaa, anio, pernoctaciones
#     hoja "metadatos"

Sys.setenv(R_LIBCURL_SSL_REVOKE_BEST_EFFORT = "TRUE")

library(ineapir)
library(httr)
library(jsonlite)
library(dplyr)
library(tidyr)
library(openxlsx)

CARPETA <- "."  # ruta relativa: ejecutar desde la carpeta de este script
ID_TABLA_INE <- 39364
URL_EUROSTAT <- paste0("https://ec.europa.eu/eurostat/api/dissemination/statistics/1.0/data/",
                       "tour_ce_omn12?format=JSON&lang=EN&indic_to=NGT_SP&c_resid=TOTAL&month=TOTAL")

NUTS2_INE <- c(ES = "00", ES61 = "01", ES24 = "02", ES12 = "03", ES53 = "04", ES70 = "05",
               ES13 = "06", ES41 = "07", ES42 = "08", ES51 = "09", ES52 = "10", ES43 = "11",
               ES11 = "12", ES30 = "13", ES62 = "14", ES22 = "15", ES21 = "16", ES23 = "17",
               ES63 = "18", ES64 = "19")
NOMBRE_INE <- c("00" = "Total Nacional", "01" = "Andalucía", "02" = "Aragón",
                "03" = "Asturias, Principado de", "04" = "Balears, Illes", "05" = "Canarias",
                "06" = "Cantabria", "07" = "Castilla y León", "08" = "Castilla - La Mancha",
                "09" = "Cataluña", "10" = "Comunitat Valenciana", "11" = "Extremadura",
                "12" = "Galicia", "13" = "Madrid, Comunidad de", "14" = "Murcia, Región de",
                "15" = "Navarra, Comunidad Foral de", "16" = "País Vasco", "17" = "Rioja, La",
                "18" = "Ceuta", "19" = "Melilla")

reintentar <- function(f, ..., n = 5, sleep = 3) {
  for (i in seq_len(n)) {
    r <- tryCatch(f(...), error = function(e) {
      message("  fallo intento ", i, ": ", conditionMessage(e))
      NULL
    })
    if (!is.null(r)) return(r)
    Sys.sleep(sleep * i)
  }
  stop("La fuente no respondio tras ", n, " intentos")
}

# --- 1. INE: viviendas turisticas y plazas ------------------------------------

varval <- reintentar(get_metadata_table_varval, idTable = ID_TABLA_INE)
ids_ccaa <- varval %>% filter(Fk_Variable %in% c(70, 349)) %>% pull(Id)
stopifnot(length(ids_ccaa) == 20)

ine <- reintentar(get_data_table, idTable = ID_TABLA_INE, nlast = 99, unnest = TRUE,
                  metanames = TRUE, metacodes = TRUE, tip = "AM")
ine <- ine %>%
  filter(`Comunidades.Autónomas.y.Provincias.Id` %in% ids_ccaa,
         Viviendas.y.plazas %in% c("Viviendas turísticas", "Plazas")) %>%
  transmute(codigo_ccaa = ifelse(`Comunidades.Autónomas.y.Provincias` == "Total Nacional", "00",
                                 `Comunidades.Autónomas.y.Provincias.Codigo`),
            fecha = substr(Fecha, 1, 10),
            anio = as.integer(substr(Fecha, 1, 4)),
            medida = ifelse(Viviendas.y.plazas == "Plazas", "plazas", "viviendas"),
            valor = Valor) %>%
  pivot_wider(names_from = medida, values_from = valor) %>%
  mutate(ccaa = unname(NOMBRE_INE[codigo_ccaa])) %>%
  select(codigo_ccaa, ccaa, fecha, anio, viviendas, plazas) %>%
  arrange(codigo_ccaa, fecha)

stopifnot(n_distinct(ine$codigo_ccaa) == 20, !anyNA(ine$viviendas),
          all(table(ine$codigo_ccaa) == n_distinct(ine$fecha)))
cuadre <- ine %>% group_by(fecha) %>%
  summarise(esp = viviendas[codigo_ccaa == "00"], suma = sum(viviendas[codigo_ccaa != "00"]))
stopifnot(all(abs(cuadre$suma - cuadre$esp) <= 20))

# --- 2. Eurostat: pernoctaciones en plataformas -------------------------------

resp <- reintentar(function() {
  r <- httr::GET(URL_EUROSTAT)
  stopifnot(httr::status_code(r) == 200)
  r
})
js <- fromJSON(httr::content(resp, as = "text", encoding = "UTF-8"), simplifyVector = FALSE)
# JSON-stat: indice plano = sum(indice_dim * paso); las dimensiones que no son
# geo ni time tienen un solo valor (se filtran en la URL)
dims <- unlist(js$id)
tam <- unlist(js$size)
stopifnot(all(tam[!dims %in% c("geo", "time")] == 1))
geo_idx  <- unlist(js$dimension$geo$category$index)
time_idx <- unlist(js$dimension$time$category$index)
paso <- rev(cumprod(rev(c(tam[-1], 1))))
names(paso) <- dims
valores <- js$value

euro <- expand.grid(geo = names(geo_idx), anio = names(time_idx), stringsAsFactors = FALSE) %>%
  filter(geo %in% names(NUTS2_INE)) %>%
  mutate(pos = geo_idx[geo] * paso[["geo"]] + time_idx[anio] * paso[["time"]],
         pernoctaciones = vapply(as.character(pos), function(p) {
           v <- valores[[p]]
           if (is.null(v)) NA_real_ else as.numeric(v)
         }, numeric(1)),
         codigo_ccaa = unname(NUTS2_INE[geo]),
         ccaa = unname(NOMBRE_INE[codigo_ccaa]),
         anio = as.integer(anio)) %>%
  filter(!is.na(pernoctaciones)) %>%
  select(codigo_ccaa, ccaa, anio, pernoctaciones) %>%
  arrange(codigo_ccaa, anio)

# Solo años con dato para las 17 CCAA
anios_ok <- euro %>% filter(codigo_ccaa %in% sprintf("%02d", 1:17)) %>%
  count(anio) %>% filter(n == 17) %>% pull(anio)
euro <- euro %>% filter(anio %in% anios_ok)
stopifnot(length(anios_ok) >= 5)

message("INE viviendas turisticas: ", paste(unique(ine$fecha), collapse = ", "),
        "\nEurostat pernoctaciones: ", min(euro$anio), "-", max(euro$anio))

wb <- createWorkbook()
addWorksheet(wb, "ine")
writeData(wb, "ine", ine)
addWorksheet(wb, "eurostat")
writeData(wb, "eurostat", euro)
addWorksheet(wb, "metadatos")
writeData(wb, "metadatos", data.frame(
  campo = c("ine", "ine_url", "eurostat", "eurostat_url", "eurostat_actualizado", "descarga"),
  valor = c("INE, estadistica experimental de viviendas turisticas: viviendas y plazas anunciadas en plataformas (foto en cada fecha)",
            paste0("https://www.ine.es/jaxiT3/Tabla.htm?t=", ID_TABLA_INE),
            "Eurostat tour_ce_omn12: noches pasadas en alojamientos de corta estancia reservados en plataformas en linea (Airbnb, Booking, Expedia, Tripadvisor); total de huespedes y meses",
            URL_EUROSTAT,
            if (is.null(js$updated)) "" else js$updated,
            format(Sys.Date()))
))
saveWorkbook(wb, file.path(CARPETA, "VIVIENDA_TURISTICA_CCAA.xlsx"), overwrite = TRUE)
