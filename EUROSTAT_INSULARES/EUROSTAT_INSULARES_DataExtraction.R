# EUROSTAT_INSULARES_DataExtraction.R
#
# Series de Eurostat para el estudio de las regiones insulares europeas
# (EUROSTAT_INSULARES.Rmd): que distingue a los archipielagos de la parte
# continental de su pais.
#
# Cada serie se descarga DOS veces con la API JSON de Eurostat
# (https://ec.europa.eu/eurostat/api/dissemination/statistics/1.0/data/):
#   - geoLevel = nuts2   : todas las regiones NUTS 2 (las insulares y las
#                          continentales, estas para dar contexto europeo);
#   - geoLevel = country : el total de cada pais. La referencia de cada isla
#                          es "su pais menos sus regiones insulares", que se
#                          calcula restando valores ABSOLUTOS (numeradores y
#                          denominadores por separado) en
#                          EUROSTAT_INSULARES_Consolidar.R.
# Por eso se descargan sobre todo series absolutas (personas, euros, noches,
# viviendas) y no los cocientes que publica Eurostat: un cociente no se puede
# "restar". Las que solo existen como tasa (pobreza, esperanza de vida...)
# se descargan como tasa y se restan ponderando por poblacion.
#
# Filtros: cada serie fija una sola categoria en todas las dimensiones salvo
# geo y time (y el mes en la estacionalidad). Los codigos se comprobaron con
# el "content constraint" de la API (2026-11-14).
#
# TRAMPA: las tablas no estan todas en la misma version NUTS. Las de cuentas
# regionales (nama_10r_*) usan NUTS 2024; otras siguen en NUTS 2021 (Portugal
# continental PT16-PT18, Paises Bajos NL31/NL33, Croacia HR04) y las de
# turismo arrastran codigos de NUTS 2013 en años viejos. A las regiones
# insulares no les afecta (sus codigos no han cambiado) y la referencia
# continental se calcula con el total del pais, que tampoco cambia: por eso
# se descarga el nivel pais. Solo el contexto europeo (percentiles entre
# regiones continentales) pierde las pocas regiones con codigo distinto.
#
# Genera en esta carpeta:
#   EUROSTAT_INSULARES_Series.xlsx   hoja "series": serie, nivel, geo, anio,
#                                    mes, valor (formato largo)
#                                    hoja "catalogo": serie, dataset, filtros, url

Sys.setenv(R_LIBCURL_SSL_REVOKE_BEST_EFFORT = "TRUE")

library(httr)
library(jsonlite)
library(dplyr)
library(tibble)
library(openxlsx)

CARPETA <- "."  # ruta relativa: ejecutar desde la carpeta de este script
API <- "https://ec.europa.eu/eurostat/api/dissemination/statistics/1.0/data/"
DESDE <- 2012

# serie | dataset | filtros
SERIES <- tribble(
  ~serie,              ~dataset,           ~filtros,
  # --- Poblacion -------------------------------------------------------------
  "pob_media",         "nama_10r_3popgdp", list(unit = "THS"),
  "pob_total",         "demo_r_pjanaggr3", list(unit = "NR", sex = "T", age = "TOTAL"),
  "pob_lt15",          "demo_r_pjanaggr3", list(unit = "NR", sex = "T", age = "Y_LT15"),
  "pob_ge65",          "demo_r_pjanaggr3", list(unit = "NR", sex = "T", age = "Y_GE65"),
  "pob_1enero",        "demo_r_gind3",     list(indic_de = "JAN"),
  "pob_crecimiento",   "demo_r_gind3",     list(indic_de = "GROW"),
  "pob_saldo_migr",    "demo_r_gind3",     list(indic_de = "CNMIGRAT"),
  "pob_crec_natural",  "demo_r_gind3",     list(indic_de = "NATGROW"),
  "superficie_tierra", "reg_area3",        list(landuse = "L0008", unit = "KM2"),
  "lfs_pob_15_74",     "lfst_r_lfsd2pwc",  list(unit = "THS_PER", wstatus = "POP", c_birth = "TOTAL", sex = "T", age = "Y15-74"),
  "lfs_nacidos_fuera", "lfst_r_lfsd2pwc",  list(unit = "THS_PER", wstatus = "POP", c_birth = "FOR", sex = "T", age = "Y15-74"),
  # --- Economia ---------------------------------------------------------------
  "vab_TOTAL",         "nama_10r_3gva",    list(unit = "CP_MEUR", nace_r2 = "TOTAL"),
  "vab_A",             "nama_10r_3gva",    list(unit = "CP_MEUR", nace_r2 = "A"),
  "vab_B-E",           "nama_10r_3gva",    list(unit = "CP_MEUR", nace_r2 = "B-E"),
  "vab_F",             "nama_10r_3gva",    list(unit = "CP_MEUR", nace_r2 = "F"),
  "vab_G-I",           "nama_10r_3gva",    list(unit = "CP_MEUR", nace_r2 = "G-I"),
  "vab_J",             "nama_10r_3gva",    list(unit = "CP_MEUR", nace_r2 = "J"),
  "vab_K",             "nama_10r_3gva",    list(unit = "CP_MEUR", nace_r2 = "K"),
  "vab_L",             "nama_10r_3gva",    list(unit = "CP_MEUR", nace_r2 = "L"),
  "vab_M_N",           "nama_10r_3gva",    list(unit = "CP_MEUR", nace_r2 = "M_N"),
  "vab_O-Q",           "nama_10r_3gva",    list(unit = "CP_MEUR", nace_r2 = "O-Q"),
  "vab_R-U",           "nama_10r_3gva",    list(unit = "CP_MEUR", nace_r2 = "R-U"),
  "pib_meur",          "nama_10r_2gdp",    list(unit = "MIO_EUR"),
  "pib_mpps",          "nama_10r_2gdp",    list(unit = "MIO_PPS_EU27_2020"),
  # Indice de volumen del PIB (2020 = 100): para deflactar. El PIB real de cada
  # año = PIB del ultimo año (PPA) * indice(año) / indice(ultimo año)
  "pib_vol_i20",       "nama_10r_2gvagr",  list(unit = "I20", na_item = "B1GQ"),
  "renta_hog_mpps",    "nama_10r_2hhinc",  list(unit = "MIO_PPS_EU27_2020", na_item = "B6N", direct = "BAL"),
  "horas_trabajadas",  "nama_10r_2emhrw",  list(unit = "THS", wstatus = "EMP", nace_r2 = "TOTAL"),
  "fbcf_meur",         "nama_10r_2gfcf",   list(currency = "MIO_EUR", sector = "S1", nace_r2 = "TOTAL"),
  "id_meur",           "rd_e_gerdreg",     list(sectperf = "TOTAL", unit = "MIO_EUR"),
  "empresas_nacidas",  "bd_hgnace_r",      list(indic_sbs = "ENT_BRTH_NR", nace_r2 = "B-S_X_O_S94"),
  "empresas_activas",  "bd_hgnace_r",      list(indic_sbs = "ENT_NR", nace_r2 = "B-S_X_O_S94"),
  # --- Trabajo (Encuesta de Poblacion Activa de la UE) -----------------------
  "lfs_ocup_20_64",    "lfst_r_lfe2emp",   list(unit = "THS_PER", sex = "T", age = "Y20-64"),
  "lfs_pob_20_64",     "lfst_r_lfsd2pop",  list(unit = "THS_PER", sex = "T", age = "Y20-64", isced11 = "TOTAL"),
  "lfs_parados",       "lfst_r_lfu3pers",  list(unit = "THS_PER", sex = "T", age = "Y15-74", isced11 = "TOTAL"),
  "lfs_activos",       "lfst_r_lfp2act",   list(unit = "THS_PER", sex = "T", age = "Y15-74"),
  "lfs_parados_ld",    "lfst_r_lfu2ltu",   list(unit = "THS_PER", sex = "T", age = "Y15-74", isced11 = "TOTAL"),
  "lfs_ocup_total",    "lfst_r_lfe2eftpt", list(unit = "THS_PER", sex = "T", age = "Y15-74", wstatus = "EMP", worktime = "TOTAL"),
  "lfs_ocup_parcial",  "lfst_r_lfe2eftpt", list(unit = "THS_PER", sex = "T", age = "Y15-74", wstatus = "EMP", worktime = "PT"),
  "neet_15_29_pc",     "edat_lfse_22",     list(unit = "PC", sex = "T", age = "Y15-29", training = "NO_FE_NO_NFE", wstatus = "NEMP"),
  # --- Turismo ----------------------------------------------------------------
  "noches_total",      "tour_occ_nin2",    list(unit = "NR", nace_r2 = "I551-I553", c_resid = "TOTAL"),
  "noches_no_resid",   "tour_occ_nin2",    list(unit = "NR", nace_r2 = "I551-I553", c_resid = "FOR"),
  "noches_mensuales",  "tour_occ_nin2m",   list(unit = "NR", nace_r2 = "I551-I553", c_resid = "TOTAL"),
  "noches_plataformas","tour_ce_omn12",    list(indic_to = "NGT_SP", c_resid = "TOTAL", month = "TOTAL", unit = "NR"),
  "plazas_turisticas", "tour_cap_nuts2",   list(unit = "NR", nace_r2 = "I551-I553", accomunit = "BEDPL"),
  # --- Vivienda (Censo 2021) ---------------------------------------------------
  "viv_total",         "cens_21dwob_r3",   list(unit = "NR", housing = "DW", building = "TOTAL"),
  "viv_no_ocupadas",   "cens_21dwob_r3",   list(unit = "NR", housing = "DW_NOC", building = "TOTAL"),
  "viv_total_periodo", "cens_21dwop_r3",   list(unit = "NR", housing = "DW", y_const = "TOTAL"),
  "viv_2011_2015",     "cens_21dwop_r3",   list(unit = "NR", housing = "DW", y_const = "Y2011-2015"),
  "viv_desde_2016",    "cens_21dwop_r3",   list(unit = "NR", housing = "DW", y_const = "Y_GE2016"),
  "viv_propiedad",     "cens_21dwbo_r2",   list(unit = "NR", housing = "DW_OWN", building = "TOTAL"),
  "viv_alquiler",      "cens_21dwbo_r2",   list(unit = "NR", housing = "DW_RENT", building = "TOTAL"),
  "viv_otra_tenencia", "cens_21dwbo_r2",   list(unit = "NR", housing = "DW_OTH", building = "TOTAL"),
  # --- Educacion ---------------------------------------------------------------
  "lfs_pob_25_64",     "lfst_r_lfsd2pop",  list(unit = "THS_PER", sex = "T", age = "Y25-64", isced11 = "TOTAL"),
  "lfs_superiores",    "lfst_r_lfsd2pop",  list(unit = "THS_PER", sex = "T", age = "Y25-64", isced11 = "ED5-8"),
  "formacion_cont_pc", "trng_lfse_04",     list(unit = "PC", sex = "T", age = "Y25-64"),
  # --- Salud -------------------------------------------------------------------
  "esperanza_vida",    "demo_r_mlifexp",   list(unit = "YR", sex = "T", age = "Y_LT1"),
  "muertes_infantiles","demo_r_minf",      list(unit = "NR"),
  "nacimientos",       "demo_r_births",    list(unit = "NR"),
  "medicos",           "hlth_rs_physreg",  list(unit = "NR"),
  "camas_hosp",        "hlth_rs_bdsrg2",   list(unit = "NR"),
  # --- Pobreza (EU-SILC: solo como tasa) ---------------------------------------
  "arope_pc",          "ilc_peps11n",      list(unit = "PC"),
  "privacion_sev_pc",  "ilc_mdsd18",       list(unit = "PC"),
  "sobrecarga_viv_pc", "ilc_lvho07_r",     list(unit = "PC"),
  "retrasos_pagos_pc", "ilc_mdes05_r",     list(unit = "PC"),
  # --- Transporte y seguridad --------------------------------------------------
  "pasajeros_aereos",  "tran_r_avpa_nm",   list(tra_meas = "PAS_CRD", unit = "THS_PAS"),
  "turismos",          "tran_r_vehst",     list(vehicle = "CAR", unit = "NR"),
  "muertos_carretera", "tran_r_acci",      list(victim = "KIL", unit = "NR"),
  "homicidios",        "crim_gen_reg",     list(iccs = "ICCS0101", unit = "NR"),
  "robos_vivienda",    "crim_gen_reg",     list(iccs = "ICCS0501", unit = "NR")
)
stopifnot(!any(duplicated(SERIES$serie)))

reintentar <- function(f, ..., n = 5, sleep = 3) {
  for (i in seq_len(n)) {
    r <- tryCatch(f(...), error = function(e) {
      message("  fallo intento ", i, ": ", conditionMessage(e))
      NULL
    })
    if (!is.null(r)) return(r)
    Sys.sleep(sleep * i)
  }
  stop("Eurostat no respondio tras ", n, " intentos")
}

# JSON-stat de Eurostat -> data.frame (geo, anio, [mes], valor). El indice
# plano de cada valor se descompone en el indice de cada dimension con los
# "pasos" de la matriz (la ultima dimension varia mas deprisa).
leer_json_stat <- function(js) {
  ids <- unlist(js$id)
  tam <- as.integer(unlist(js$size))
  if (length(js$value) == 0) return(NULL)
  codigos <- lapply(ids, function(d) {
    ix <- unlist(js$dimension[[d]]$category$index)
    names(ix)[order(ix)]
  })
  names(codigos) <- ids
  k <- as.numeric(names(js$value))
  v <- vapply(js$value, function(x) if (is.null(x)) NA_real_ else as.numeric(x), numeric(1))
  paso <- rev(cumprod(rev(c(tam[-1], 1))))
  dims <- lapply(seq_along(ids), function(j) codigos[[j]][(k %/% paso[j]) %% tam[j] + 1])
  names(dims) <- ids
  otras <- setdiff(ids[tam > 1], c("geo", "time", "month"))
  if (length(otras) > 0) stop("Dimensiones sin filtrar: ", paste(otras, collapse = ", "))
  out <- data.frame(geo = dims$geo, anio = as.integer(substr(dims$time, 1, 4)),
                    mes = if ("month" %in% ids) dims$month else NA_character_,
                    valor = unname(v), stringsAsFactors = FALSE)
  out[!is.na(out$valor), ]
}

descargar <- function(dataset, filtros, nivel) {
  q <- c(list(format = "JSON", lang = "EN", geoLevel = nivel, sinceTimePeriod = DESDE), filtros)
  r <- reintentar(function() {
    resp <- httr::GET(paste0(API, dataset), query = q, httr::timeout(300))
    if (httr::status_code(resp) != 200) stop("HTTP ", httr::status_code(resp))
    resp
  })
  leer_json_stat(fromJSON(httr::content(r, as = "text", encoding = "UTF-8"), simplifyVector = FALSE))
}

# Cache por serie y nivel (reanudable): borrar la carpeta cache/ para volver a
# descargar todo
dir.create(file.path(CARPETA, "cache"), showWarnings = FALSE)
todo <- list()
for (i in seq_len(nrow(SERIES))) {
  s <- SERIES[i, ]
  for (nivel in c("nuts2", "country")) {
    fc <- file.path(CARPETA, "cache", paste0(s$serie, "_", nivel, ".rds"))
    if (file.exists(fc)) {
      x <- readRDS(fc)
    } else {
      message(s$serie, " (", s$dataset, ", ", nivel, ")")
      x <- descargar(s$dataset, s$filtros[[1]], nivel)
      if (is.null(x) || nrow(x) == 0) stop("Sin datos: ", s$serie, " ", nivel)
      saveRDS(x, fc)
    }
    todo[[length(todo) + 1]] <- x %>% mutate(serie = s$serie, nivel = nivel, .before = 1)
  }
}
series <- bind_rows(todo) %>% arrange(serie, nivel, geo, anio, mes)

# Las regiones insulares y sus paises deben estar en cada serie (salvo las
# ausencias ya conocidas, que el consolidado trata como dato que falta)
ISLAS <- c("ES53", "ES70", "PT20", "PT30", "EL41", "EL42", "EL62", "FI20", "MT00", "FRY1",
           "ITG1", "ITG2", "FRM0", "EL43", "CY00", "FRY2", "FRY4", "FRY5")
cobertura <- series %>% filter(geo %in% ISLAS) %>% distinct(serie, geo) %>% count(serie, name = "n_islas")
print(as.data.frame(cobertura %>% filter(n_islas < length(ISLAS))))
# Total nacional de los 6 paises con islas: sin el, no hay referencia
# continental para esa serie y ese pais (se avisa; el consolidado lo trata
# como dato que falta)
faltan_pais <- expand.grid(serie = SERIES$serie, geo = c("ES", "PT", "EL", "IT", "FR", "FI"),
                           stringsAsFactors = FALSE) %>%
  anti_join(series %>% filter(nivel == "country") %>% distinct(serie, geo), by = c("serie", "geo"))
if (nrow(faltan_pais) > 0) {
  message("Series sin total nacional para algun pais con islas:")
  print(faltan_pais)
}

catalogo <- SERIES %>%
  mutate(filtros = vapply(filtros, function(f) paste(names(f), unlist(f), sep = "=", collapse = "&"), character(1)),
         url = paste0(API, dataset, "?", filtros))

message(nrow(series), " valores de ", nrow(SERIES), " series")
wb <- createWorkbook()
addWorksheet(wb, "series")
writeData(wb, "series", series)
addWorksheet(wb, "catalogo")
writeData(wb, "catalogo", catalogo)
addWorksheet(wb, "metadatos")
writeData(wb, "metadatos", data.frame(
  campo = c("fuente", "api", "desde", "descarga"),
  valor = c("Eurostat, base de datos regional (NUTS 2) y nacional",
            API, DESDE, format(Sys.time(), "%Y-%m-%d %H:%M"))))
saveWorkbook(wb, file.path(CARPETA, "EUROSTAT_INSULARES_Series.xlsx"), overwrite = TRUE)
