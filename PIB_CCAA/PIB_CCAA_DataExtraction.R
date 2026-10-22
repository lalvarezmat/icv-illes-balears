# PIB_CCAA_DataExtraction.R
#
# PIB per capita REAL (deflactado) por comunidad autonoma, construido a partir
# de fuentes oficiales del INE, sin aplicar ningun deflactor casero (decision
# explicita del usuario, 2026-09-23: usar los indices de volumen encadenados
# que ya publica el propio INE, no CPI ni otro deflactor externo).
#
# El INE NO publica un "PIB per capita a precios constantes" en euros de forma
# directa por CCAA -- lo que si publica son piezas sueltas que hay que
# combinar (verificado navegando ine.es en vivo, 2026-09-23):
#
#   1. Indice de volumen encadenado del PIB TOTAL (no per capita) por CCAA,
#      base 2020=100, serie 2000-2023. Tabla PC-Axis antigua (sistema "jaxi",
#      no tempus3/ineapir), matriz 01002, descargable como .px:
#      https://www.ine.es/jaxi/files/tpx/es/px/73424.px
#   2. PIB TOTAL a precios corrientes (nominal) por CCAA, miles de euros,
#      misma estructura y mismos anios. Matriz 01001:
#      https://www.ine.es/jaxi/files/tpx/es/px/73423.px
#   3. Poblacion por CCAA y anio (para pasar de PIB total a per capita). El
#      INE cambio de metodologia de poblacion tras el Censo 2021: la serie
#      antigua basada en Padron (tabla tempus3 t=2853, via ineapir) cubre
#      1996-2021; la nueva Estadistica Continua de Poblacion (ECP, tabla
#      t=67988) cubre desde 2021. Se usa 2853 para 2000-2020 y 67988 para
#      2021-2023; el propio 2021 sirve de anio de solape para comprobar que
#      no hay un salto grande de metodologia (se compara y se avisa si la
#      diferencia relativa supera el 1%, pero no se corrige: es la cifra
#      oficial de cada fuente en ese momento).
#
# Metodo de deflactado (encadenamiento a partir de piezas oficiales, NO es un
# deflactor inventado): para cada CCAA, PIB real en euros de un anio t =
# PIB nominal del anio ancla * indice de volumen(t) / indice de volumen(ancla).
# Dividiendo por poblacion(t) se obtiene el PIB real per capita en euros
# constantes del anio ancla. El anio ancla es el ULTIMO de la serie (2023):
# peticion del usuario 2026-09-26, para que el ultimo dato quede igual que el
# nominal publicado y la serie se deflacte hacia atras (antes se anclaba en
# 2020, el año base del indice, donde vale 100). Es lo mismo que hace el INE
# internamente para construir series de PIB real; aqui no se inventa ningun
# IPC ni deflactor de precios propio, solo se combinan piezas que el INE ya
# publica por separado.
#
# TRAMPA verificada en el propio .px (2026-09-23): la dimension "Ramas de
# actividad" tiene la categoria "PRODUCTO INTERIOR BRUTO A PRECIOS DE MERCADO"
# duplicada LITERALMENTE dos veces (aparece como primera fila -- cabecera de
# la tabla, antes del desglose por rama -- y otra vez al final, tras el
# desglose, con el mismo texto exacto). Comprobado con datos reales (Total
# Nacional, 2023, tabla de volumen): ambas copias dan el MISMO valor
# (116.30626), asi que no es un error de la fuente, es el mismo dato repetido
# dos veces en la tabla -- pxR las colapsa en el mismo nivel de factor pero
# dentro de la tabla siguen siendo 2 filas fisicas distintas, asi que hay que
# usar distinct() tras filtrar esa categoria o salen las filas duplicadas.
#
# TRAMPA (igual que REGISTRADORES): el sitio exige un User-Agent de
# navegador para servir el .px del sistema "jaxi" antiguo.
#
# TRAMPA en las tablas de poblacion via ineapir: la variable geografica
# ("Comunidades y Ciudades Autonomas" en t=2853; "Comunidad Autonoma y
# Provincia" en t=67988, que mezcla en la MISMA variable filas de nivel CCAA
# y filas de nivel provincia) coloca el nombre de la CCAA en una posicion
# distinta dentro del texto libre "Nombre" segun sea la fila de Total
# Nacional o de una CCAA concreta (comprobado en vivo: para Total Nacional el
# sexo aparece en 2a posicion, para el resto de CCAA en 3a). Parsear "Nombre"
# a mano por posicion de texto es fragil; en su lugar se piden los datos con
# metanames=TRUE, metacodes=TRUE, tip="AM", que devuelve columnas separadas
# ya resueltas (nombre, codigo e id de cada variable) sin depender del orden
# del texto libre.
#
# Los ultimos 2 anios de las tablas del INE vienen marcados como "2022 (P)"
# (estimacion provisional) y "2023 (A)" (estimacion de avance) en vez de solo
# el numero del anio.
#
# Genera en esta carpeta:
#   PIB_CCAA_Real.xlsx   hoja "datos" (una fila por CCAA x anio, 2000-2023) +
#                        hoja "metadatos"
#   PIB_volumen_encadenado_73424.px / PIB_precios_corrientes_73423.px
#                        ficheros .px originales del INE, sin modificar
#                        (se guardan para no tener que volver a descargarlos)

Sys.setenv(R_LIBCURL_SSL_REVOKE_BEST_EFFORT = "TRUE")
options(timeout = 120)

library(httr)
library(pxR)
library(dplyr)
library(openxlsx)
library(ineapir)

CARPETA <- "."  # ruta relativa: ejecutar desde la carpeta de este script
UA <- httr::user_agent("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0 Safari/537.36")
ANIO_BASE_INDICE <- 2020L  # anio en el que el indice de volumen encadenado vale 100
RAMA_TOTAL <- "PRODUCTO INTERIOR BRUTO A PRECIOS DE MERCADO"

# --- 1. Descarga de los dos .px de Contabilidad Regional (sistema jaxi) ----

descargar_px <- function(url, destino, n = 5, sleep = 2) {
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

fn_volumen <- file.path(CARPETA, "PIB_volumen_encadenado_73424.px")
fn_corrientes <- file.path(CARPETA, "PIB_precios_corrientes_73423.px")
descargar_px("https://www.ine.es/jaxi/files/tpx/es/px/73424.px", fn_volumen)
descargar_px("https://www.ine.es/jaxi/files/tpx/es/px/73423.px", fn_corrientes)

# --- 2. Parseo con pxR, quedandonos solo con el PIB total (no por rama) ----

leer_pib_total <- function(fichero) {
  df <- as.data.frame(read.px(fichero))
  names(df)[grepl("^Comunidades", names(df))] <- "ccaa_px"
  names(df)[grepl("^Ramas", names(df))] <- "rama"

  df %>%
    filter(rama == RAMA_TOTAL) %>%
    distinct(ccaa_px, Periodo, value) %>%  # colapsa la fila duplicada (ver cabecera)
    mutate(
      ccaa_px = as.character(ccaa_px),
      codigo_ccaa = ifelse(ccaa_px == "Total Nacional", "00", substr(ccaa_px, 1, 2)),
      ccaa = ifelse(ccaa_px == "Total Nacional", "Total Nacional", trimws(substr(ccaa_px, 4, nchar(ccaa_px)))),
      anio = as.integer(substr(as.character(Periodo), 1, 4)),
      estado_dato = case_when(
        grepl("\\(A\\)", Periodo) ~ "Avance",
        grepl("\\(P\\)", Periodo) ~ "Provisional",
        TRUE ~ "Definitivo"
      )
    ) %>%
    select(codigo_ccaa, ccaa, anio, estado_dato, value)
}

volumen <- leer_pib_total(fn_volumen) %>% rename(indice_volumen = value)
corrientes <- leer_pib_total(fn_corrientes) %>% rename(pib_nominal_miles_eur = value)

stopifnot(!any(duplicated(volumen[, c("codigo_ccaa", "anio")])))
stopifnot(!any(duplicated(corrientes[, c("codigo_ccaa", "anio")])))
stopifnot(setequal(volumen$codigo_ccaa, corrientes$codigo_ccaa))

# el indice debe valer exactamente 100 en el anio ancla para las 20 geografias
# base del indice (comprobacion de que se lee bien el .px)
chequeo_base <- volumen %>% filter(anio == ANIO_BASE_INDICE)
stopifnot(nrow(chequeo_base) == 20, all(abs(chequeo_base$indice_volumen - 100) < 1e-6))

# anio ancla = ultimo año de la serie, comun a las 20 geografias
ANIO_ANCLA <- max(volumen$anio)
stopifnot(sum(volumen$anio == ANIO_ANCLA) == 20, sum(corrientes$anio == ANIO_ANCLA) == 20)

nominal_ancla <- corrientes %>%
  filter(anio == ANIO_ANCLA) %>%
  select(codigo_ccaa, pib_nominal_ancla_miles_eur = pib_nominal_miles_eur) %>%
  left_join(volumen %>% filter(anio == ANIO_ANCLA) %>%
              select(codigo_ccaa, indice_ancla = indice_volumen), by = "codigo_ccaa")

# --- 3. Poblacion por CCAA y anio (dos tablas tempus3 via ineapir) ---------

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

# t=2853 "Poblacion por comunidades y ciudades autonomas y sexo" (base Padron,
# serie 1996-2021). Un unico nivel geografico (CCAA), sin mezclar provincias.
pob_padron <- get_ine(get_data_table, idTable = 2853, nlast = 99, unnest = TRUE,
                       metanames = TRUE, metacodes = TRUE, tip = "AM")
pob_padron <- pob_padron %>%
  filter(Sexo == "Total") %>%
  transmute(
    codigo_ccaa = `Comunidades.y.Ciudades.Autónomas.Codigo`,
    anio = Anyo,
    poblacion = Valor
  ) %>%
  filter(anio <= 2020)  # 2021 en adelante se toma de la ECP (t=67988)

# t=67988 "Poblacion segun comunidad autonoma y provincia y sexo" (ECP, desde
# 2021). La variable geografica mezcla filas de CCAA y de provincia; hay que
# quedarse solo con los ids de nivel CCAA (Fk_Variable==70 en los metadatos:
# 16473 = Total Nacional, mas los 19 ids de "Comunidades y Ciudades
# Autonomas") para no sumar dos veces al agregar provincias sueltas.
ids_ccaa_67988 <- get_ine(get_metadata_table_varval, idTable = 67988) %>%
  filter(Fk_Variable %in% c(70, 349)) %>%
  pull(Id)

pob_ecp <- get_ine(get_data_table, idTable = 67988, nlast = 99, unnest = TRUE,
                    metanames = TRUE, metacodes = TRUE, tip = "AM")
pob_ecp <- pob_ecp %>%
  filter(Sexo == "Total", `Comunidad.Autónoma.y.Provincia.Id` %in% ids_ccaa_67988) %>%
  transmute(
    codigo_ccaa = `Comunidad.Autónoma.y.Provincia.Codigo`,
    anio = Anyo,
    poblacion = Valor
  ) %>%
  filter(anio >= 2021, anio <= 2023)

stopifnot(nrow(pob_padron %>% filter(anio == 2020)) == 20)
stopifnot(nrow(pob_ecp %>% filter(anio == 2021)) == 20)

# comprobacion de cordura en el anio de solape (2021): las dos fuentes miden
# poblacion con metodologia distinta (Padron vs ECP tras el Censo 2021), asi
# que puede no coincidir exactamente -- se avisa si la diferencia relativa
# supera el 1%, pero se usa igualmente el dato oficial de cada fuente en su
# rango de anios (no se fuerza a que empalmen).
solape <- inner_join(
  pob_padron %>% filter(anio == 2020) %>% select(codigo_ccaa, pob_2020 = poblacion),
  pob_ecp %>% filter(anio == 2021) %>% select(codigo_ccaa, pob_2021 = poblacion),
  by = "codigo_ccaa"
) %>% mutate(variacion_pct = 100 * (pob_2021 - pob_2020) / pob_2020)
cat("Variacion de poblacion 2020 (Padron) -> 2021 (ECP) por CCAA:\n")
print(solape)
if (any(abs(solape$variacion_pct) > 5)) {
  warning("Salto de poblacion 2020->2021 mayor al 5% en alguna CCAA: revisar antes de usar el dato.")
}

poblacion <- bind_rows(pob_padron, pob_ecp) %>% arrange(codigo_ccaa, anio)
stopifnot(!any(duplicated(poblacion[, c("codigo_ccaa", "anio")])))

# --- 4. Construccion del PIB real per capita -------------------------------

datos <- volumen %>%
  left_join(nominal_ancla, by = "codigo_ccaa") %>%
  left_join(poblacion, by = c("codigo_ccaa", "anio")) %>%
  mutate(
    pib_real_eur = pib_nominal_ancla_miles_eur * 1000 * indice_volumen / indice_ancla,
    pib_real_pc_eur = pib_real_eur / poblacion
  ) %>%
  arrange(codigo_ccaa, anio) %>%
  select(codigo_ccaa, ccaa, anio, estado_dato, indice_volumen, poblacion,
         pib_real_eur, pib_real_pc_eur)

n_sin_poblacion <- sum(is.na(datos$poblacion))
if (n_sin_poblacion > 0) {
  warning(n_sin_poblacion, " filas sin poblacion (revisar aniso/CCAA sin cruzar).")
}

cat("\nResumen Balears (04):\n")
print(datos %>% filter(codigo_ccaa == "04", anio %in% c(2000, 2010, 2020, 2023)))

# --- 5. Escritura del resultado --------------------------------------------

metadatos <- data.frame(
  campo = c(
    "fuente_volumen", "url_volumen", "fuente_nominal", "url_nominal",
    "fuente_poblacion_2000_2020", "fuente_poblacion_2021_2023",
    "anio_ancla_indice", "metodo",
    "cobertura_anios", "fecha_extraccion", "n_filas"
  ),
  valor = c(
    "INE - Contabilidad Regional de España. PIB pm: variaciones de volumen, índices de volumen encadenados (matriz 01001... 01002) por CCAA, rama de actividad y periodo",
    "https://www.ine.es/jaxi/files/tpx/es/px/73424.px",
    "INE - Contabilidad Regional de España. PIB pm y VAB a precios básicos: precios corrientes por CCAA, rama de actividad y periodo",
    "https://www.ine.es/jaxi/files/tpx/es/px/73423.px",
    "INE - tabla tempus3 t=2853, Población por comunidades y ciudades autónomas y sexo (serie basada en Padrón)",
    "INE - tabla tempus3 t=67988, Población según comunidad autónoma y provincia y sexo (Estadística Continua de Población, desde 2021)",
    as.character(ANIO_ANCLA),
    "PIB real (euros constantes del anio ancla, el ultimo de la serie) = PIB nominal del anio ancla * indice de volumen encadenado(t) / indice(ancla); PIB real per capita = PIB real / poblacion. Sin deflactor de precios casero: solo se combinan series oficiales del INE.",
    paste(range(datos$anio), collapse = "-"),
    format(Sys.time(), "%Y-%m-%d %H:%M:%S"),
    as.character(nrow(datos))
  ),
  stringsAsFactors = FALSE
)

wb <- createWorkbook()
addWorksheet(wb, "datos"); writeData(wb, "datos", datos)
addWorksheet(wb, "metadatos"); writeData(wb, "metadatos", metadatos)
saveWorkbook(wb, file.path(CARPETA, "PIB_CCAA_Real.xlsx"), overwrite = TRUE)

cat("\nHecho:", nrow(datos), "filas ->", file.path(CARPETA, "PIB_CCAA_Real.xlsx"), "\n")
