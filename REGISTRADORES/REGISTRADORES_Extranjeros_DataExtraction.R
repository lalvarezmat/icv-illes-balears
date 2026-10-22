# Porcentaje de compras de vivienda realizadas por extranjeros, por comunidad
# autonoma y año, segun la Estadistica Registral Inmobiliaria (ERI) del
# Colegio de Registradores (2026-10-10).
#
# DE DONDE SALE. El portal de datos abiertos (opendata.registradores.org), que
# usa REGISTRADORES_DataExtraction.R, NO tiene este dato: sus 2 ficheros de
# compraventas solo traen numero, importe y precio por m2. La nacionalidad del
# comprador se publica unicamente en los ANUARIOS de la ERI, en PDF (no hay
# Excel ni CSV), en la pagina
#   https://www.registradores.org/actualidad/portal-estadistico-registral/estadisticas-de-propiedad
# Hay anuarios de 2004 a 2025; el dato nacional de extranjeros aparece en el
# de 2006, pero la tabla POR COMUNIDAD AUTONOMA solo desde el anuario 2012.
# Por eso la serie por CCAA es 2012-2025 (14 años).
#
# QUE TABLAS SE LEEN (se localizan por su contenido, no por numero de pagina):
#   1. Tabla anual "Nacionalidad en las compras de vivienda" de cada anuario:
#      % nacionales y % extranjeros de las 17 CCAA en el año del anuario, mas
#      la fila NACIONAL/España (desde 2013). Es la PRIMERA PUBLICACION de cada
#      año. España 2012 sale del texto del anuario 2012 ("El 8,12% de las
#      compras de vivienda realizadas en España durante el año 2012 ..."),
#      porque esa tabla no trae fila nacional.
#   2. Tabla de 8 años "Compras de vivienda por extranjeros en Comunidades
#      Autonomas (%)" de los anuarios 2021 en adelante (el de 2021 trae
#      2014-2021; el de 2025, 2018-2025). Son cifras REVISADAS.
#   3. Grafico "Compras de vivienda por extranjeros (%). Anual" de los mismos
#      anuarios: España 2009-<año del anuario>, revisada. Cada etiqueta de valor
#      se asigna al año cuyo rotulo del eje tiene mas cerca en horizontal.
#
# QUE VALOR SE USA. El anuario 2021 reviso UNA vez 2014-2019 (sobre todo
# 2016-2019: hasta -0,76 pp en una CCAA, un 25 % relativo en las de menos
# compras de extranjeros) y desde entonces los anuarios repiten las mismas
# cifras. El total nacional casi no cambio (2019: 12,61 -> 12,45; 2020: 11,34
# -> 11,32): fue un reparto entre CCAA. Por eso el valor vigente
# (pct_extranjeros) es la ULTIMA cifra publicada de cada año (tablas 2 y 3 del
# anuario mas reciente que lo trae; 2012-2013, que ninguna tabla de 8 años
# cubre, de su propio anuario). La primera publicacion se conserva en
# pct_primera_publicacion, y todas las versiones en la hoja "revisiones".
#
# QUE MIDE. Porcentaje de compraventas de vivienda inscritas en los Registros
# de la Propiedad cuyo comprador es de nacionalidad extranjera. Es la
# NACIONALIDAD, no la residencia: un extranjero residente en España cuenta como
# extranjero y un español residente fuera, como nacional. La ERI dice trabajar
# con practicamente todas las compraventas inscritas ("tamaño censal"), no con
# una muestra (los anuarios de 2006-2009 hablaban todavia de "muestreo").
#
# TRAMPAS:
#   - Los nombres de las CCAA cambian de un anuario a otro ("BALEARES",
#     "BALEARS, ILLES", "Balears, Illes"; "COM. MADRID", "MADRID, COMUNIDAD
#     DE"...). Se normalizan a ASCII en mayusculas sin espacios y se emparejan
#     con el codigo INE 01-17 (el mismo de INE_ICV_CCAA_Consolidado.xlsx).
#   - Hay otras tablas por CCAA con dos porcentajes que suman 100 (personas
#     fisicas/juridicas, tipo de interes fijo/variable...). La tabla anual se
#     reconoce por estar en una pagina con un encabezado "nacionales ...
#     extranjeros" y por tener las 17 CCAA (eso descarta tambien las tablas de
#     provincias, que tienen filas "BALEARES", "MURCIA"... pero no las 17).
#   - Algunos anuarios repiten la tabla en el anexo: se comprueba que las
#     copias coinciden en las 17 CCAA. El dato nacional se toma de la copia del
#     cuerpo (la primera). En el anuario 2012 la copia del cuerpo no tiene fila
#     nacional y la del anexo dice 8,32 %, que contradice el 8,12 % del texto
#     del mismo anuario (y el que repiten para 2012 los anuarios posteriores):
#     se usa el del texto y se avisa de la discrepancia.
#   - El servidor exige User-Agent de navegador, igual que el portal de datos
#     abiertos.
#
# Los PDF se guardan en ERI_Anuarios/ (unos 100 MB; en .gitignore) y no se
# vuelven a descargar si ya estan. Genera REGISTRADORES_Extranjeros.xlsx con
# las hojas "datos" (larga), "serie_ancha" (año x geografia), "revisiones" y
# "metadatos".
#
# NO esta incorporado a INE_ICV_CCAA_Consolidado.xlsx (pendiente, a decision
# del usuario). Lo usa solo REGISTRADORES_Extranjeros.Rmd.

Sys.setenv(R_LIBCURL_SSL_REVOKE_BEST_EFFORT = "TRUE")

library(httr)
library(pdftools)
library(stringi)
library(openxlsx)

httr::set_config(httr::timeout(120))

CARPETA <- "."  # ruta relativa: ejecutar desde la carpeta de este script
CARPETA_PDF <- file.path(CARPETA, "ERI_Anuarios")
PORTAL <- "https://www.registradores.org/actualidad/portal-estadistico-registral/estadisticas-de-propiedad"
UA <- httr::user_agent("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0 Safari/537.36")
PRIMER_ANUARIO_CCAA <- 2012

# Codigo INE y nombre oficial de las 17 CCAA (+ "00" para España), con la
# expresion que reconoce cada una sobre el nombre normalizado.
CCAA <- data.frame(
  cod_ca = c("00", sprintf("%02d", 1:17)),
  ccaa = c("Total Nacional", "Andalucía", "Aragón", "Asturias, Principado de",
           "Balears, Illes", "Canarias", "Cantabria", "Castilla y León",
           "Castilla - La Mancha", "Cataluña", "Comunitat Valenciana",
           "Extremadura", "Galicia", "Madrid, Comunidad de", "Murcia, Región de",
           "Navarra, Comunidad Foral de", "País Vasco", "Rioja, La"),
  patron = c("^(NACIONAL|ESPANA)$", "^ANDALUCIA$", "^ARAGON$", "^ASTURIAS",
             "^BALEAR", "^CANARIAS$", "^CANTABRIA$", "^CASTILLAYLEON$",
             "^CASTILLALAMANCHA$", "^CATALUNA$",
             "^(COMVALENCIANA|COMUNITATVALENCIANA|COMUNIDADVALENCIANA)$",
             "^EXTREMADURA$", "^GALICIA$", "^(COMMADRID|MADRIDCOMUNIDADDE|COMUNIDADDEMADRID)$",
             "^MURCIA", "^NAVARRA", "^PAISVASCO$", "^(LARIOJA|RIOJALA)$"),
  stringsAsFactors = FALSE
)

codigo_ccaa <- function(nombre) {
  clave <- gsub("[^A-Z]", "", toupper(stri_trans_general(nombre, "Latin-ASCII")))
  cod <- vapply(clave, function(k) {
    hit <- CCAA$cod_ca[vapply(CCAA$patron, grepl, logical(1), x = k)]
    if (length(hit) == 1) hit else NA_character_
  }, character(1))
  unname(cod)
}

num_es <- function(x) as.numeric(gsub(",", ".", gsub("\\.", "", x)))

# --- Descarga de los anuarios (solo los que falten) ---
descargar_anuarios <- function() {
  dir.create(CARPETA_PDF, showWarnings = FALSE)
  resp <- httr::GET(PORTAL, UA)
  httr::stop_for_status(resp)
  html <- httr::content(resp, "text", encoding = "UTF-8")
  hrefs <- unique(regmatches(html, gregexpr('/documents/[^"]*Anuario[^"]*\\.pdf[^"]*', html))[[1]])
  anios <- as.integer(sub(".*Anuario[_+](20[0-9]{2}).*", "\\1", hrefs))
  stopifnot(!anyNA(anios), !anyDuplicated(anios))
  enlaces <- data.frame(anio = anios, url = paste0("https://www.registradores.org", hrefs))
  enlaces <- enlaces[enlaces$anio >= PRIMER_ANUARIO_CCAA, ]
  enlaces <- enlaces[order(enlaces$anio), ]
  # Sin huecos: un anuario que faltara dejaria un año sin dato en silencio
  stopifnot(all(diff(enlaces$anio) == 1))

  for (i in seq_len(nrow(enlaces))) {
    destino <- file.path(CARPETA_PDF, sprintf("ERI_Anuario_%d.pdf", enlaces$anio[i]))
    if (file.exists(destino)) next
    cat("Descargando anuario", enlaces$anio[i], "...\n")
    r <- httr::GET(enlaces$url[i], UA, httr::write_disk(destino, overwrite = TRUE))
    httr::stop_for_status(r)
  }
  enlaces$fichero <- file.path(CARPETA_PDF, sprintf("ERI_Anuario_%d.pdf", enlaces$anio))
  enlaces
}

# --- Reconstruye las lineas de una pagina a partir de la posicion de cada
#     palabra (pdf_data), agrupando las que comparten altura. pdf_text() no
#     sirve en los anuarios recientes: alli la columna de nombres y las de
#     valores salen en bloques separados. ---
lineas_pagina <- function(palabras, tol = 3) {
  if (nrow(palabras) == 0) return(character(0))
  palabras$yc <- palabras$y + palabras$height / 2
  palabras <- palabras[order(palabras$yc, palabras$x), ]
  grupo <- cumsum(c(TRUE, diff(palabras$yc) > tol))
  unname(vapply(split(palabras, grupo),
                function(z) paste(z$text[order(z$x)], collapse = " "), character(1)))
}

# Separa una linea en nombre (texto hasta el primer numero) y numeros, anotando
# si cada numero lleva "%" detras.
partir_fila <- function(linea) {
  m <- regexpr("\\s-?[0-9][0-9.]*,[0-9]+", linea)
  if (m < 0) return(NULL)
  nombre <- trimws(substr(linea, 1, m - 1))
  resto <- substr(linea, m, nchar(linea))
  nums <- regmatches(resto, gregexpr("-?[0-9][0-9.]*,[0-9]+\\s*%?", resto))[[1]]
  list(nombre = nombre, valor = num_es(sub("\\s*%$", "", nums)), pct = grepl("%$", nums))
}

# Tabla anual: filas "CCAA  % nacionales  % extranjeros [var. pp]", en una
# pagina con el encabezado "% nacionales % extranjeros" (2012-2020) o
# "Nacionales Extranjeros" (2021 en adelante).
tabla_anual <- function(lineas) {
  if (!any(grepl("nacionales.*extranjeros", lineas, ignore.case = TRUE))) return(NULL)
  filas <- lapply(lineas, partir_fila)
  filas <- Filter(function(f) !is.null(f) && length(f$valor) >= 2 && all(f$pct[1:2]) &&
                    abs(sum(f$valor[1:2]) - 100) < 0.05, filas)
  if (length(filas) == 0) return(NULL)
  df <- data.frame(nombre = vapply(filas, `[[`, "", "nombre"),
                   pct_nacionales = vapply(filas, function(f) f$valor[1], 0),
                   pct_extranjeros = vapply(filas, function(f) f$valor[2], 0))
  df$cod_ca <- codigo_ccaa(df$nombre)
  df <- df[!is.na(df$cod_ca), ]
  if (!all(sprintf("%02d", 1:17) %in% df$cod_ca)) return(NULL)  # p. ej. tablas de provincias
  stopifnot(!anyDuplicated(df$cod_ca))
  df
}

# Tabla de 8 años: cabecera "18 T4 19 T4 ..." y filas "CCAA  v1 % ... v8 %".
tabla_varios_anios <- function(lineas) {
  if (!any(grepl("Compras de vivienda por extranjeros en Comunidades", lineas))) return(NULL)
  cab <- grep("([0-9]{2} T4.*){4,}", lineas, value = TRUE)
  if (length(cab) != 1) return(NULL)
  anios <- 2000 + as.integer(regmatches(cab, gregexpr("[0-9]{2}(?= T4)", cab, perl = TRUE))[[1]])
  filas <- lapply(lineas, partir_fila)
  filas <- Filter(function(f) !is.null(f) && length(f$valor) == length(anios) && all(f$pct), filas)
  df <- do.call(rbind, lapply(filas, function(f)
    data.frame(nombre = f$nombre, anio = anios, pct_extranjeros = f$valor)))
  df$cod_ca <- codigo_ccaa(df$nombre)
  df <- df[!is.na(df$cod_ca), ]
  stopifnot(all(sprintf("%02d", 1:17) %in% df$cod_ca), !anyDuplicated(df[, c("cod_ca", "anio")]))
  df
}

# Grafico de barras "Compras de vivienda por extranjeros (%). Anual" (España).
# Trabaja con las palabras sueltas: cada etiqueta de valor esta encima de su
# barra, asi que se asigna al rotulo de año del eje mas cercano en horizontal.
# Solo cuentan las etiquetas entre el titulo del grafico y el eje (en la misma
# pagina hay una tabla por CCAA mas arriba).
grafico_nacional <- function(palabras) {
  lineas <- lineas_pagina(palabras)
  if (!any(grepl("Compras de vivienda por extranjeros \\(%\\)\\. Anual", lineas))) return(NULL)
  rot <- palabras[grepl("^20[0-9]{2}$", palabras$text), ]
  if (nrow(rot) < 10) return(NULL)
  eje <- rot[abs(rot$y - max(rot$y)) < 3, ]
  eje$xc <- eje$x + eje$width / 2
  titulo_y <- max(palabras$y[palabras$text == "Anual" & palabras$y < min(eje$y)])
  et <- palabras[grepl("^[0-9]+,[0-9]+$", palabras$text) &
                   palabras$y > titulo_y & palabras$y < min(eje$y), ]
  # el centro del bloque "14,98 %" queda un poco a la derecha del numero
  xc <- et$x + et$width / 2 + 4
  cerca <- vapply(xc, function(v) which.min(abs(eje$xc - v)), integer(1))
  dist <- abs(eje$xc[cerca] - xc)
  stopifnot(all(dist < median(diff(sort(eje$xc))) / 4))  # cada etiqueta, claramente sobre una barra
  df <- data.frame(anio = as.integer(eje$text[cerca]), pct_extranjeros = num_es(et$text))
  stopifnot(!anyDuplicated(df$anio), setequal(df$anio, as.integer(eje$text)))
  df[order(df$anio), ]
}

enlaces <- descargar_anuarios()

anual <- list()
varios <- list()
for (i in seq_len(nrow(enlaces))) {
  anio <- enlaces$anio[i]
  cat("Leyendo anuario", anio, "...\n")
  paginas <- pdf_data(enlaces$fichero[i])
  lineas <- lapply(paginas, lineas_pagina)

  # 1. Tabla anual (puede estar repetida en el anexo: las copias deben coincidir)
  copias <- Filter(Negate(is.null), lapply(seq_along(lineas), function(p) {
    t <- tabla_anual(lineas[[p]]); if (!is.null(t)) t$pagina <- p; t
  }))
  if (length(copias) == 0) stop("No se encuentra la tabla anual por CCAA en el anuario ", anio)
  ccaa17 <- function(x) { x <- x[x$cod_ca != "00", ]; x[order(x$cod_ca), "pct_extranjeros"] }
  for (k in seq_along(copias)[-1]) {
    if (any(ccaa17(copias[[1]]) != ccaa17(copias[[k]])))
      stop("Las copias de la tabla anual del anuario ", anio, " no coinciden en las CCAA")
  }
  t <- copias[[1]]
  nac_copias <- unlist(lapply(copias, function(x) x$pct_extranjeros[x$cod_ca == "00"]))
  t$anio <- anio
  t$anuario <- anio

  # 3. España 2012: del texto, porque la tabla del cuerpo no trae fila nacional
  if (!"00" %in% t$cod_ca) {
    texto <- paste(vapply(lineas, paste, "", collapse = " "), collapse = " ")
    patron <- paste0("El\\s+([0-9]+,[0-9]+)%\\s+de\\s+las\\s+compras\\s+de\\s+vivienda\\s+",
                     "realizadas\\s+en\\s+Espa\u00f1a\\s+durante\\s+el\\s+a\u00f1o\\s+", anio)
    m <- regmatches(texto, regexec(patron, texto))[[1]]
    if (length(m) != 2) stop("No se encuentra el porcentaje nacional en el texto del anuario ", anio)
    ext <- num_es(m[2])
    t <- rbind(t, data.frame(nombre = "(texto del anuario)", pct_nacionales = 100 - ext,
                             pct_extranjeros = ext, cod_ca = "00", pagina = NA,
                             anio = anio, anuario = anio))
  }
  nac <- t$pct_extranjeros[t$cod_ca == "00"]
  if (any(nac_copias != nac))
    message("  Aviso: otra copia de la tabla del anuario ", anio, " da para España ",
            paste(format(nac_copias[nac_copias != nac], decimal.mark = ","), collapse = ", "),
            " % (se usa ", format(nac, decimal.mark = ","), " %)")
  stopifnot(nrow(t) == 18)
  anual[[as.character(anio)]] <- t

  # 2. Tabla de 8 años y 3. grafico nacional (anuarios 2021 en adelante)
  v <- Filter(Negate(is.null), lapply(lineas, tabla_varios_anios))
  g <- Filter(Negate(is.null), lapply(paginas, grafico_nacional))
  if (anio >= 2021 && (length(v) == 0 || length(g) == 0))
    stop("Falta la tabla de 8 años o el grafico nacional en el anuario ", anio)
  if (length(v) > 0) {
    v <- v[[1]]
    v$anuario <- anio
    g <- g[[1]]
    g <- g[g$anio >= PRIMER_ANUARIO_CCAA, ]
    g$cod_ca <- "00"
    g$anuario <- anio
    varios[[as.character(anio)]] <- rbind(v[, names(g)], g)
  }
}

anual <- do.call(rbind, anual)
varios <- do.call(rbind, varios)
stopifnot(max(varios$anuario) == max(anual$anio))  # el ultimo anuario trae su tabla de 8 años

# --- Todas las versiones publicadas de cada dato: la primera (tabla anual) y
#     las revisadas (tablas de 8 años y grafico nacional) ---
versiones <- rbind(
  data.frame(anio = anual$anio, cod_ca = anual$cod_ca, anuario = anual$anuario,
             pct_extranjeros = anual$pct_extranjeros,
             tabla = ifelse(is.na(anual$pagina), "texto", paste0("tabla anual, p. ", anual$pagina))),
  data.frame(anio = varios$anio, cod_ca = varios$cod_ca, anuario = varios$anuario,
             pct_extranjeros = varios$pct_extranjeros,
             tabla = ifelse(varios$cod_ca == "00", "grafico anual", "tabla de 8 años"))
)
# Un mismo anuario no puede dar dos cifras distintas para el mismo dato
stopifnot(!anyDuplicated(unique(versiones[, c("anio", "cod_ca", "anuario", "pct_extranjeros")])[
  , c("anio", "cod_ca", "anuario")]))

# --- Hoja "datos": serie larga. Vigente = ultima publicacion ---
# (los anuarios 2021+ dan el dato de su propio año dos veces, en la tabla anual
#  y en la de 8 años, con la misma cifra: como fuente se cita la tabla anual)
versiones$anual <- versiones$tabla %in% "texto" | startsWith(versiones$tabla, "tabla anual")
versiones <- versiones[order(versiones$cod_ca, versiones$anio, versiones$anuario, versiones$anual), ]
ultima <- versiones[!duplicated(versiones[, c("cod_ca", "anio")], fromLast = TRUE), ]
primera <- versiones[!duplicated(versiones[, c("cod_ca", "anio")]), ]
datos <- merge(ultima, primera[, c("cod_ca", "anio", "pct_extranjeros")],
               by = c("cod_ca", "anio"), suffixes = c("", "_primera"))
names(datos)[names(datos) == "pct_extranjeros_primera"] <- "pct_primera_publicacion"
datos <- merge(datos, CCAA[, c("cod_ca", "ccaa")], by = "cod_ca")
datos$fuente <- paste0("ERI Anuario ", datos$anuario, ", ", datos$tabla)
datos <- datos[order(datos$cod_ca, datos$anio),
               c("anio", "cod_ca", "ccaa", "pct_extranjeros", "pct_primera_publicacion", "fuente")]
rownames(datos) <- NULL

# Serie completa en las 18 geografias y sin huecos
stopifnot(nrow(datos) == 18 * length(unique(datos$anio)), !anyNA(datos$pct_extranjeros),
          setequal(datos$anio, PRIMER_ANUARIO_CCAA:max(anual$anio)))

# --- Hoja "serie_ancha": un año por fila, una geografia por columna ---
ancha <- reshape(datos[, c("anio", "ccaa", "pct_extranjeros")], idvar = "anio",
                 timevar = "ccaa", direction = "wide")
names(ancha) <- sub("^pct_extranjeros\\.", "", names(ancha))
ancha <- ancha[, c("anio", CCAA$ccaa)]

# --- Hoja "revisiones": lo que dicen los anuarios posteriores de cada dato,
#     frente a su primera publicacion ---
rev <- merge(versiones[versiones$anuario > versiones$anio, ],
             primera[, c("cod_ca", "anio", "pct_extranjeros")],
             by = c("cod_ca", "anio"), suffixes = c("", "_primera"))
rev <- merge(rev, CCAA[, c("cod_ca", "ccaa")], by = "cod_ca")
rev$diferencia_pp <- round(rev$pct_extranjeros - rev$pct_extranjeros_primera, 2)
rev <- rev[order(rev$cod_ca, rev$anio, rev$anuario),
           c("anio", "cod_ca", "ccaa", "pct_extranjeros_primera", "anuario", "tabla",
             "pct_extranjeros", "diferencia_pp")]
names(rev)[5:7] <- c("anuario_revision", "tabla_revision", "pct_extranjeros_revisado")
rownames(rev) <- NULL
cat("Revisiones: ", nrow(rev), " valores comparados con su primera publicacion; ",
    "diferencia maxima ", max(abs(rev$diferencia_pp)), " pp; ",
    sum(datos$pct_extranjeros != datos$pct_primera_publicacion),
    " datos vigentes distintos de su primera publicacion\n", sep = "")

metadatos <- data.frame(
  campo = c("fuente", "url", "indicador", "definicion", "frecuencia", "cobertura",
            "valor_vigente", "revisiones", "fecha_descarga"),
  valor = c("Colegio de Registradores - Estadistica Registral Inmobiliaria (ERI), anuarios",
            PORTAL,
            "Porcentaje de compraventas de vivienda realizadas por extranjeros",
            "Nacionalidad del comprador (no residencia); compraventas de vivienda inscritas en los Registros de la Propiedad",
            "Anual",
            paste0("17 CCAA + España, ", min(datos$anio), "-", max(datos$anio),
                   " (la tabla por CCAA empieza en el anuario 2012)"),
            "pct_extranjeros = ultima cifra publicada de cada año; pct_primera_publicacion = la del anuario de ese año",
            "Hoja 'revisiones': cada version publicada frente a la primera (el anuario 2021 reviso 2014-2019)",
            format(Sys.time(), "%Y-%m-%d %H:%M:%S")),
  stringsAsFactors = FALSE
)

wb <- createWorkbook()
addWorksheet(wb, "datos"); writeData(wb, "datos", datos)
addWorksheet(wb, "serie_ancha"); writeData(wb, "serie_ancha", ancha)
addWorksheet(wb, "revisiones"); writeData(wb, "revisiones", rev)
addWorksheet(wb, "metadatos"); writeData(wb, "metadatos", metadatos)
saveWorkbook(wb, file.path(CARPETA, "REGISTRADORES_Extranjeros.xlsx"), overwrite = TRUE)

cat("Hecho:", nrow(datos), "filas,", min(datos$anio), "-", max(datos$anio),
    "-> REGISTRADORES_Extranjeros.xlsx\n")
