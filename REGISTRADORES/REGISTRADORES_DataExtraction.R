# Descarga y limpia la Estadistica Registral Inmobiliaria del Colegio de
# Registradores de la Propiedad (opendata.registradores.org), compraventas
# trimestrales de inmuebles por uso residencial y uso comercial.
#
# A diferencia de IBESTAT/INE_ICV, aqui NO hay que recorrer un catalogo de
# indicadores: cada uso (residencial/comercial) es UN UNICO fichero ya
# consolidado con TODOS los trimestres (2007-T1 en adelante) y TODOS los
# niveles geograficos (Nacional/Comunidad/Provincia) juntos en las mismas
# filas -- la columna "geo" distingue el nivel de cada fila. No hace falta
# agregar provincias para obtener el nivel CCAA, ya viene en el propio
# fichero (filas con geo == "Comunidad").
#
# Genera en esta carpeta:
#   - REGISTRADORES_Residencial.xlsx  hoja "datos" (viviendas/garajes/
#     trasteros) + hoja "metadatos"
#   - REGISTRADORES_Comercial.xlsx    hoja "datos" (locales/naves
#     industriales) + hoja "metadatos"
#
# TRAMPA (verificada 2026-09-22): el sitio bloquea peticiones sin cabecera
# User-Agent de navegador ("The requested URL was rejected") -- WebFetch y
# curl por defecto caen en ese bloqueo. httr::user_agent(...) lo evita.
#
# No es reanudable por tramos (son solo 2 ficheros, no un catalogo de
# cientos) -- simplemente vuelve a descargar y sobrescribe cada vez que se
# lanza. Fuente: https://opendata.registradores.org/dataset/dataset

Sys.setenv(R_LIBCURL_SSL_REVOKE_BEST_EFFORT = "TRUE")

library(httr)
library(openxlsx)

httr::set_config(httr::timeout(60))

CARPETA <- "."  # ruta relativa: ejecutar desde la carpeta de este script
UA <- httr::user_agent("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0 Safari/537.36")

FUENTES <- list(
  residencial = list(
    url = "https://opendata.registradores.org/data-integration/compraventas-residencial-trimestres-provincias-es/RP_ComprvResid_2007-2t2026.csv",
    salida = "REGISTRADORES_Residencial.xlsx",
    descripcion = "Compraventas de inmuebles de uso residencial (viviendas, garajes, trasteros)",
    cols_num = c("viv-num", "viv-imp", "viv-pm2", "gar-num", "gar-imp", "gar-pm2", "tras-num", "tras-imp", "tras-pm2")
  ),
  comercial = list(
    url = "https://opendata.registradores.org/data-integration/compraventas-comercial-trimestres-provincias-es/RP_ComprvCom_2007-2t2026.csv",
    salida = "REGISTRADORES_Comercial.xlsx",
    descripcion = "Compraventas de inmuebles de uso comercial (locales, naves industriales)",
    cols_num = c("loc-num", "loc-imp", "loc-pm2", "nav-num", "nav-imp", "nav-pm2")
  )
)

# --- Descarga un CSV con separador ";" y decimal "," (formato tipico
#     espanol), quita la columna vacia que deja el ";" final de cada linea
#     de origen, y convierte a numerico las columnas de importes/precios ---
descargar_y_limpiar <- function(fuente) {
  resp <- httr::GET(fuente$url, UA)
  httr::stop_for_status(resp)
  texto <- httr::content(resp, "text", encoding = "UTF-8")

  datos <- read.csv2(text = texto, encoding = "UTF-8", stringsAsFactors = FALSE)
  datos <- datos[, !grepl("^X$", names(datos)), drop = FALSE]  # columna fantasma del ";" final

  names(datos) <- gsub("\\.", "-", names(datos))
  names(datos)[names(datos) == "cod-ca"] <- "cod_ca"
  names(datos)[names(datos) == "cod-prv"] <- "cod_prv"

  for (col in fuente$cols_num) {
    col_r <- gsub("-", "-", col)  # nombres ya usan "-", coherente tras el gsub anterior
    if (col_r %in% names(datos)) datos[[col_r]] <- suppressWarnings(as.numeric(datos[[col_r]]))
  }

  datos$fecha <- paste0(datos$ano, "-Q", datos$trim)
  datos <- datos[, c("fecha", "ano", "trim", "geo", "cod_ca", "ca", "cod_prv", "prv",
                      setdiff(names(datos), c("fecha", "ano", "trim", "geo", "cod_ca", "ca", "cod_prv", "prv"))), drop = FALSE]
  datos
}

for (nombre in names(FUENTES)) {
  fuente <- FUENTES[[nombre]]
  cat("Descargando", nombre, "...\n")
  datos <- descargar_y_limpiar(fuente)

  metadatos <- data.frame(
    campo = c("fuente", "url", "descripcion", "frecuencia", "cobertura_geografica",
              "niveles_geo", "fecha_descarga", "n_filas"),
    valor = c("Colegio de Registradores de la Propiedad - Estadistica Registral Inmobiliaria",
              fuente$url, fuente$descripcion, "Trimestral", "Espana (Nacional/CCAA/Provincia)",
              "geo == 'Nacional' | 'Comunidad' | 'Provincia' (filtrar segun el nivel deseado)",
              format(Sys.time(), "%Y-%m-%d %H:%M:%S"), as.character(nrow(datos))),
    stringsAsFactors = FALSE
  )

  wb <- createWorkbook()
  addWorksheet(wb, "datos"); writeData(wb, "datos", datos)
  addWorksheet(wb, "metadatos"); writeData(wb, "metadatos", metadatos)
  saveWorkbook(wb, file.path(CARPETA, fuente$salida), overwrite = TRUE)

  cat("  ->", nrow(datos), "filas,", fuente$salida, "\n")
}

cat("Hecho.\n")
