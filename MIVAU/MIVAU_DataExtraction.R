# MIVAU_DataExtraction.R
#
# Dos estadisticas del Ministerio de Vivienda y Agenda Urbana (MIVAU) por
# comunidad autonoma, para el estudio de los archipielagos
# (INE_ICV_CCAA/INE_ICV_CCAA_Insulares_Consolidar.R):
#
#   1. VIVIENDAS INICIADAS Y TERMINADAS (Estadistica de Vivienda Libre +
#      vivienda protegida), anual, 2008-ultimo año. El MIVAU ya no sirve los
#      .XLS de su antiguo boletin (las URL de transportes.gob.es/sedal/322...
#      dan 404, comprobado 2026-11-14), pero el ISTAC republica la serie
#      completa de TODAS las CCAA y provincias en su API de datos abiertos
#      (la misma API eDatos que usa IBESTAT, sin clave):
#        dataset ISTAC E25004A_000001 "Viviendas libres y protegidas,
#        iniciadas y terminadas. España, comunidades autonomas y provincias
#        por periodos". "~latest" = ultima version publicada.
#      Las CCAA vienen con codigo NUTS 2 (ES53 = Illes Balears, ES70 =
#      Canarias...). Antes de 2008 solo hay dato de Canarias: el resto de
#      CCAA empieza en 2008. El ultimo año viene marcado "P" (provisional).
#      Iniciadas = visados de direccion de obra nueva (libre) + calificaciones
#      provisionales (protegida); terminadas = certificados de fin de obra +
#      calificaciones definitivas.
#
#   2. COMPRAVENTAS DE VIVIENDA SEGUN RESIDENCIA DEL COMPRADOR (Estadistica
#      de Transacciones Inmobiliarias, datos de los notarios), tabla 1.6 del
#      boletin, trimestral desde 2007, un .XLS con una hoja por trimestre
#      ("1t 2007" ... ). Columnas (iguales en las 78 hojas, comprobado
#      2026-11-14): nombre | vacia | TOTAL | Residentes en España (Total,
#      Españoles, Extranjeros, No consta) | No residentes en España (Total,
#      Españoles, Extranjeros, No consta) | No consta. Se suman los cuatro
#      trimestres de cada año y solo se guardan los años completos.
#      "No residentes" se refiere a la RESIDENCIA del comprador, no a su
#      nacionalidad (a diferencia de la ERI de los Registradores, que mide
#      nacionalidad: ver ../REGISTRADORES/REGISTRADORES_Extranjeros_DataExtraction.R).
#
# TRAMPA: en las hojas, las CCAA uniprovinciales (Illes Balears, Madrid...)
# pueden aparecer dos veces, como comunidad y como provincia, con el mismo
# nombre y el mismo valor. Se toma la primera aparicion.
# TRAMPA: el portal exige User-Agent de navegador (como REGISTRADORES y jaxi).
#
# Genera en esta carpeta:
#   MIVAU_Viviendas.xlsx        hoja "datos": codigo_ccaa, ccaa, anio,
#                               iniciadas, terminadas, estado (P = provisional)
#   MIVAU_Transacciones.xlsx    hoja "datos": codigo_ccaa, ccaa, anio, total,
#                               residentes, no_residentes, no_residentes_extranjeros,
#                               n_trimestres
#   ISTAC_E25004A_000001.xlsx / MIVAU_340101d0.XLS   ficheros originales
#   (si existen no se vuelven a descargar: borrarlos para actualizar)

Sys.setenv(R_LIBCURL_SSL_REVOKE_BEST_EFFORT = "TRUE")
options(timeout = 300)

library(httr)
library(readxl)
library(dplyr)
library(tidyr)
library(openxlsx)

CARPETA <- "."  # ruta relativa: ejecutar desde la carpeta de este script
UA <- httr::user_agent("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0 Safari/537.36")

URL_VIVIENDAS <- "https://datos.canarias.es/api/estadisticas/statistical-resources/v1.0/datasets/ISTAC/E25004A_000001/~latest.xlsx"
URL_TRANSACCIONES <- "https://apps.fomento.gob.es/BoletinOnline2/sedal/340101d0.XLS"
FN_VIVIENDAS <- file.path(CARPETA, "ISTAC_E25004A_000001.xlsx")
FN_TRANSACCIONES <- file.path(CARPETA, "MIVAU_340101d0.XLS")

# NUTS 2 -> codigo INE de comunidad (el que usa el resto del proyecto)
NUTS2_INE <- c(ES = "00", ES61 = "01", ES24 = "02", ES12 = "03", ES53 = "04", ES70 = "05",
               ES13 = "06", ES41 = "07", ES42 = "08", ES51 = "09", ES52 = "10", ES43 = "11",
               ES11 = "12", ES30 = "13", ES62 = "14", ES22 = "15", ES21 = "16", ES23 = "17",
               ES63 = "18", ES64 = "19")

# Nombres del boletin del MIVAU -> codigo INE
MIVAU_INE <- c("TOTAL NACIONAL" = "00", "Andalucía" = "01", "Aragón" = "02",
               "Asturias (Principado de)" = "03", "Balears (Illes)" = "04", "Canarias" = "05",
               "Cantabria" = "06", "Castilla y León" = "07", "Castilla-La Mancha" = "08",
               "Cataluña" = "09", "Comunitat Valenciana" = "10", "Extremadura" = "11",
               "Galicia" = "12", "Madrid (Comunidad de)" = "13", "Murcia (Región de)" = "14",
               "Navarra (Comunidad Foral de)" = "15", "País Vasco" = "16", "Rioja (La)" = "17",
               "Ceuta" = "18", "Melilla" = "19")

NOMBRE_INE <- c("00" = "Total Nacional", "01" = "Andalucía", "02" = "Aragón",
                "03" = "Asturias, Principado de", "04" = "Balears, Illes", "05" = "Canarias",
                "06" = "Cantabria", "07" = "Castilla y León", "08" = "Castilla - La Mancha",
                "09" = "Cataluña", "10" = "Comunitat Valenciana", "11" = "Extremadura",
                "12" = "Galicia", "13" = "Madrid, Comunidad de", "14" = "Murcia, Región de",
                "15" = "Navarra, Comunidad Foral de", "16" = "País Vasco", "17" = "Rioja, La",
                "18" = "Ceuta", "19" = "Melilla")

descargar <- function(url, destino, n = 5, sleep = 3) {
  if (file.exists(destino)) return(invisible(TRUE))  # reanudable
  for (i in seq_len(n)) {
    ok <- tryCatch({
      resp <- httr::GET(url, UA, httr::write_disk(destino, overwrite = TRUE))
      httr::status_code(resp) == 200 && file.info(destino)$size > 0
    }, error = function(e) FALSE)
    if (ok) return(invisible(TRUE))
    Sys.sleep(sleep * i)
  }
  stop("No se pudo descargar ", url, " tras ", n, " intentos")
}

# --- 1. Viviendas iniciadas y terminadas (ISTAC) ------------------------------

descargar(URL_VIVIENDAS, FN_VIVIENDAS)
viv <- read.xlsx(FN_VIVIENDAS, sheet = 1)

viv <- viv %>%
  filter(grepl("^[0-9]{4}$", TIME_PERIOD_CODE),               # serie anual
         TERRITORIO_CODE %in% names(NUTS2_INE),               # España y CCAA (no provincias)
         MEDIDAS_CODE %in% c("VIVIENDAS_INICIADAS", "VIVIENDAS_TERMINADAS")) %>%
  transmute(codigo_ccaa = unname(NUTS2_INE[TERRITORIO_CODE]),
            anio = as.integer(TIME_PERIOD_CODE),
            medida = ifelse(MEDIDAS_CODE == "VIVIENDAS_INICIADAS", "iniciadas", "terminadas"),
            valor = suppressWarnings(as.numeric(OBS_VALUE)),
            estado = ESTADO_OBSERVACION_CODE) %>%
  filter(!is.na(valor))

estado_anio <- viv %>% group_by(codigo_ccaa, anio) %>%
  summarise(estado = paste(unique(na.omit(estado[estado != ""])), collapse = ""), .groups = "drop")

viv <- viv %>%
  select(-estado) %>%
  pivot_wider(names_from = medida, values_from = valor) %>%
  left_join(estado_anio, by = c("codigo_ccaa", "anio")) %>%
  mutate(ccaa = unname(NOMBRE_INE[codigo_ccaa])) %>%
  select(codigo_ccaa, ccaa, anio, iniciadas, terminadas, estado) %>%
  arrange(codigo_ccaa, anio)

# Desde 2008, las 20 geografias todos los años, sin huecos
viv08 <- viv %>% filter(anio >= 2008)
stopifnot(n_distinct(viv08$codigo_ccaa) == 20,
          all(table(viv08$codigo_ccaa) == n_distinct(viv08$anio)),
          !anyNA(viv08$iniciadas), !any(duplicated(viv[, c("codigo_ccaa", "anio")])))
# La suma de las CCAA cuadra con España (tolerancia por redondeo de estimaciones)
cuadre <- viv08 %>% group_by(anio) %>%
  summarise(esp = iniciadas[codigo_ccaa == "00"], suma = sum(iniciadas[codigo_ccaa != "00"]))
stopifnot(all(abs(cuadre$suma / cuadre$esp - 1) < 0.01))

# --- 2. Compraventas segun residencia del comprador (MIVAU) -------------------

descargar(URL_TRANSACCIONES, FN_TRANSACCIONES)
hojas <- excel_sheets(FN_TRANSACCIONES)
stopifnot(all(grepl("^[1-4]t [0-9]{4}$", hojas)))

leer_trimestre <- function(hoja) {
  x <- suppressMessages(read_excel(FN_TRANSACCIONES, sheet = hoja, col_names = FALSE,
                                   .name_repair = "minimal"))
  stopifnot(ncol(x) == 12)
  # Comprobacion de la cabecera: la columna 8 es el total de no residentes
  fila_cab <- which(apply(x, 1, function(f) any(f %in% "No residentes en España")))
  stopifnot(length(fila_cab) == 1, identical(x[[8]][fila_cab], "No residentes en España"),
            identical(x[[4]][fila_cab], "Residentes en España"),
            identical(x[[8]][fila_cab + 1], "Total"), identical(x[[10]][fila_cab + 1], "Extranjeros"))
  nombres <- x[[1]]
  filas <- vapply(names(MIVAU_INE), function(n) match(n, nombres), integer(1))  # 1a aparicion
  stopifnot(!anyNA(filas))
  num <- function(col) as.numeric(unlist(x[filas, col]))
  data.frame(codigo_ccaa = unname(MIVAU_INE), hoja = hoja,
             anio = as.integer(substr(hoja, 4, 7)),
             total = num(3), residentes = num(4), no_residentes = num(8),
             no_residentes_extranjeros = num(10))
}

trans <- bind_rows(lapply(hojas, leer_trimestre))
stopifnot(!anyNA(trans[, c("total", "residentes", "no_residentes")]))
# Total >= residentes + no residentes (la diferencia es "no consta")
stopifnot(all(trans$total >= trans$residentes + trans$no_residentes))

trans <- trans %>%
  group_by(codigo_ccaa, anio) %>%
  summarise(across(c(total, residentes, no_residentes, no_residentes_extranjeros), sum),
            n_trimestres = n(), .groups = "drop") %>%
  filter(n_trimestres == 4) %>%                                 # solo años completos
  mutate(ccaa = unname(NOMBRE_INE[codigo_ccaa])) %>%
  select(codigo_ccaa, ccaa, anio, everything()) %>%
  arrange(codigo_ccaa, anio)

cuadre_t <- trans %>% group_by(anio) %>%
  summarise(esp = no_residentes[codigo_ccaa == "00"], suma = sum(no_residentes[codigo_ccaa != "00"]))
stopifnot(all(cuadre_t$esp == cuadre_t$suma))

message("Viviendas: ", min(viv08$anio), "-", max(viv08$anio),
        " | Compraventas por residencia: ", min(trans$anio), "-", max(trans$anio))

guardar <- function(df, fichero, meta) {
  wb <- createWorkbook()
  addWorksheet(wb, "datos")
  writeData(wb, "datos", df)
  addWorksheet(wb, "metadatos")
  writeData(wb, "metadatos", meta)
  saveWorkbook(wb, file.path(CARPETA, fichero), overwrite = TRUE)
}

guardar(viv, "MIVAU_Viviendas.xlsx", data.frame(
  campo = c("fuente", "url", "iniciadas", "terminadas", "estado", "descarga"),
  valor = c("MIVAU, Estadistica de Vivienda Libre y de Vivienda Protegida, republicada por el ISTAC (dataset E25004A_000001)",
            URL_VIVIENDAS,
            "Viviendas iniciadas en el año (libres: visados de direccion de obra; protegidas: calificaciones provisionales)",
            "Viviendas terminadas en el año (libres: certificados de fin de obra; protegidas: calificaciones definitivas)",
            "P = provisional",
            format(Sys.Date()))))

guardar(trans, "MIVAU_Transacciones.xlsx", data.frame(
  campo = c("fuente", "url", "total", "residentes", "no_residentes", "no_residentes_extranjeros", "anio", "descarga"),
  valor = c("MIVAU, Estadistica de Transacciones Inmobiliarias (datos notariales), tabla 1.6: transacciones de vivienda segun residencia del comprador",
            URL_TRANSACCIONES,
            "Compraventas de vivienda (todas: libre y protegida, nueva y usada) en la comunidad donde esta la vivienda",
            "Compradas por residentes en España (españoles y extranjeros)",
            "Compradas por no residentes en España (españoles y extranjeros)",
            "Compradas por extranjeros no residentes",
            "Suma de los 4 trimestres; solo años completos",
            format(Sys.Date()))))
