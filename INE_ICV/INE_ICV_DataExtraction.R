# INE_ICV_DataExtraction.R
#
# Descarga completa de los Indicadores de Calidad de Vida (ICV) del INE
# (https://ine.es/dynt3/ICV/), lo más detallado posible.
#
# El ICV NO está en el catálogo tempus3/wstempus que usa `ineapir` (verificado:
# ninguna de las 5507 tablas de INE.datasets.raw.xlsx en AEDV pertenece a una
# operación "Calidad de Vida"/"Bienestar"). Es un sistema de navegación propio
# ("dynt3"), en árbol de 4 niveles:
#
#   Dimensión (ind=11250..11258, 9 en total)
#     -> Subdimensión (index.htm?padre=NNNNN)
#          -> Tema (index.htm?padre=NNNNN&capsel=MMMMM)
#               -> Tabla jaxiT3 (Tabla.htm?t=PPPPP)
#
# Cada tabla hoja SÍ es una tabla jaxiT3 normal, descargable como .px con la
# MISMA URL que usa `ineapir`/tempus3 para sus tablas:
#   https://www.ine.es/jaxiT3/files/t/px/{t}.px
# (verificado en vivo: HTTP 200, contenido PC-Axis válido).
#
# Las páginas HTML del árbol declaran <meta charset="ISO-8859-15"> pero el
# Content-Type real del servidor es UTF-8 (verificado con curl -I): hay que
# forzar encoding="UTF-8" en rvest::read_html, si no los acentos se corrompen.
#
# NOTA (petición explícita del usuario, 2026-09-20): esta pasada SOLO descarga
# en bruto -- catálogo del árbol + ficheros .px originales tal cual los sirve
# el INE, sin parsear su contenido con pxR ni convertir a xlsx/tidy. El
# procesado (parseo de los .px, pivote, limpieza) se hará en un script aparte
# más adelante, una vez decidido el indicador final del trabajo.
#
# Salida:
#   INE_ICV.catalogo.xlsx   catálogo completo (dimension/subdimension/tema/tabla)
#   DATA/<t>.px             fichero .px original de cada tabla, sin modificar.
#
# Reanudable: si ya existe INE_ICV.catalogo.xlsx se reutiliza (no se vuelve a
# recorrer el árbol); si ya existe DATA/<t>.px para una tabla, se omite su
# descarga. Para forzar el reprocesado de una tabla o del árbol completo, basta
# con borrar el fichero correspondiente.

Sys.setenv(R_LIBCURL_SSL_REVOKE_BEST_EFFORT = "TRUE")
options(timeout = 60)

library(rvest)
library(dplyr)
library(stringr)
library(openxlsx)
library(tibble)

select <- dplyr::select
mutate <- dplyr::mutate
filter <- dplyr::filter

LOG <- "INE_ICV_DataExtraction.log"
log_msg <- function(...) {
  cat(format(Sys.time()), "-", sprintf(...), "\n", file = LOG, append = TRUE)
}
cat(format(Sys.time()), "- inicio\n", file = LOG)

CATALOGO_XLSX <- "INE_ICV.catalogo.xlsx"
DATA_DIR <- "DATA"
if (!dir.exists(DATA_DIR)) dir.create(DATA_DIR)

# --- Descarga robusta de una página HTML, con reintentos ------------------
read_html_retry <- function(url, n = 5, sleep = 2) {
  for (i in seq_len(n)) {
    pg <- tryCatch(read_html(url, encoding = "UTF-8"), error = function(e) NULL)
    if (!is.null(pg)) return(pg)
    Sys.sleep(sleep * i)
  }
  stop("No se pudo leer la página tras ", n, " intentos: ", url)
}

# --- 1. CRAWL DEL ÁRBOL DE 4 NIVELES ---------------------------------------

crawl_catalogo <- function() {
  IND_IDS <- 11250:11258 # 9 dimensiones del ICV (verificado en vivo, 2026-09-20)

  filas <- list()
  for (ind in IND_IDS) {
    url_dim <- paste0("https://ine.es/dynt3/ICV/indicador.html?ind=", ind)
    pg_dim <- read_html_retry(url_dim)
    dim_name <- pg_dim %>% html_element("h1") %>% html_text2()
    log_msg("Dimension ind=%s: %s", ind, dim_name)

    arts <- pg_dim %>% html_elements("article")
    for (a in arts) {
      subdim_name <- a %>% html_element("header") %>% html_text2() %>% str_squish()
      link <- a %>% html_element("td[data-cell=Resultados] a")
      href <- link %>% html_attr("href")
      if (is.na(href)) next
      padre <- str_match(href, "padre=(\\d+)")[, 2]
      if (is.na(padre)) next

      url_sub <- paste0("https://ine.es", href)
      pg_sub <- read_html_retry(url_sub)
      temas <- pg_sub %>% html_elements("a[id^=c_]")
      log_msg("  Subdimension padre=%s (%s): %d temas", padre, subdim_name, length(temas))

      for (tm in temas) {
        capsel <- str_remove(tm %>% html_attr("id"), "^c_")
        tema_name <- tm %>% html_text2() %>% str_squish()

        url_tema <- paste0("https://ine.es/dynt3/inebase/es/index.htm?padre=", padre, "&capsel=", capsel)
        pg_tema <- read_html_retry(url_tema)
        tablas <- pg_tema %>% html_elements("a[id^=t_]")

        for (tb in tablas) {
          t_id <- str_remove(tb %>% html_attr("id"), "^t_")
          tabla_name <- tb %>% html_text2() %>% str_squish()
          filas[[length(filas) + 1]] <- tibble(
            dimension_ind = ind,
            dimension = dim_name,
            subdimension_padre = padre,
            subdimension = subdim_name,
            tema_capsel = capsel,
            tema = tema_name,
            tabla_t = t_id,
            tabla = tabla_name
          )
        }
        log_msg("    Tema capsel=%s (%s): %d tablas", capsel, tema_name, length(tablas))
      }
    }
  }
  bind_rows(filas)
}

if (file.exists(CATALOGO_XLSX)) {
  log_msg("Catalogo ya existe, se reutiliza: %s", CATALOGO_XLSX)
  catalogo <- openxlsx::read.xlsx(CATALOGO_XLSX, sheet = "catalogo")
} else {
  log_msg("Recorriendo el arbol de 4 niveles del ICV...")
  catalogo <- crawl_catalogo()
  write.xlsx(catalogo, CATALOGO_XLSX, sheetName = "catalogo")
  log_msg("Catalogo escrito: %d tablas distintas de %d dimensiones", nrow(catalogo), length(unique(catalogo$dimension_ind)))
}

log_msg("Total tablas jaxiT3 en el catalogo: %d (unicas: %d)", nrow(catalogo), n_distinct(catalogo$tabla_t))

# --- 2. DESCARGA EN BRUTO DEL .px DE CADA TABLA (sin parsear) -------------

download_px_retry <- function(url, destfile, n = 5, sleep = 2) {
  for (i in seq_len(n)) {
    ok <- tryCatch({ download.file(url, destfile, quiet = TRUE, mode = "wb"); TRUE }, error = function(e) FALSE)
    if (ok && file.exists(destfile) && file.info(destfile)$size > 0) return(TRUE)
    Sys.sleep(sleep * i)
  }
  FALSE
}

tablas_unicas <- catalogo %>% distinct(tabla_t, tabla, dimension, subdimension, tema)

n_ok <- 0L
n_fail <- 0L
for (i in seq_len(nrow(tablas_unicas))) {
  t_id <- tablas_unicas$tabla_t[i]
  titulo <- tablas_unicas$tabla[i]
  fn <- file.path(DATA_DIR, paste0(t_id, ".px"))

  if (file.exists(fn)) next # reanudacion: tabla ya descargada

  url_px <- paste0("https://www.ine.es/jaxiT3/files/t/px/", t_id, ".px")
  t0 <- Sys.time()
  ok <- download_px_retry(url_px, fn)
  seg <- as.numeric(difftime(Sys.time(), t0, units = "secs"))

  if (!ok) {
    n_fail <- n_fail + 1L
    log_msg("FALLO tabla t=%s (%s) tras reintentos (%.1fs)", t_id, titulo, seg)
    next
  }

  n_ok <- n_ok + 1L
  log_msg("OK tabla t=%s (%s): %.1fs", t_id, titulo, seg)
}

log_msg("Fin. Tablas descargadas en esta ejecucion: %d, fallos: %d", n_ok, n_fail)
cat(format(Sys.time()), "- fin\n", file = LOG, append = TRUE)
