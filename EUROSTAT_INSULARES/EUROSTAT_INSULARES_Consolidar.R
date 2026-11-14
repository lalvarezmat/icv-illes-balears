# EUROSTAT_INSULARES_Consolidar.R
#
# Convierte las series de EUROSTAT_INSULARES_Series.xlsx en indicadores
# RELATIVOS comparables entre regiones y calcula, para cada region insular, su
# REFERENCIA: la parte continental de su pais (pais menos sus regiones
# insulares). Esquema acordado con el usuario el 2026-11-14:
#
# 1. De absoluto a relativo (catalogo INDICADORES, abajo):
#    - "cociente": numerador / denominador * escala (por habitante, % de un
#      total...). Siempre que se puede, numerador y denominador salen de la
#      misma tabla de Eurostat.
#    - "tasa": indicadores que Eurostat solo publica como tasa (pobreza,
#      esperanza de vida...). Se combinan ponderando por poblacion (u otro
#      peso).
#    - Casos especiales: densidad (superficie del ultimo año), estacionalidad
#      (% de las noches en los 3 meses de mas actividad) y el PIB REAL
#      (peticion del usuario, 2026-11-14: "PIB per capita, si es posible
#      deflactado"). El PIB real de cada año se obtiene encadenando el indice
#      de volumen regional de Eurostat (nama_10r_2gvagr, 2020 = 100) sobre el
#      PIB en PPA del ultimo año (ANIO_ANCLA): PIB real(t) = PIB PPA(ancla) *
#      indice(t) / indice(ancla), en PPA constantes del año ancla. Asi el
#      nivel sigue siendo comparable entre paises (PPA) y la serie esta
#      deflactada en el tiempo (el mismo metodo que PIB_CCAA_DataExtraction.R
#      con datos del INE). De ahi salen tres indicadores: el PIB real por
#      habitante y el crecimiento medio anual del PIB real, total y por
#      habitante, en los ultimos VENTANA_CREC años (el total menos el por
#      habitante es, aproximadamente, el crecimiento de la poblacion). La
#      referencia continental resta los PIB reales de las islas del del pais:
#      con volumenes encadenados la suma no es exacta, pero en el año ancla si
#      lo es y la desviacion en los demas es pequeña.
# 2. Referencia continental de cada pais = total del pais menos sus regiones
#    insulares y menos Ceuta, Melilla y Guyane (que no son islas ni
#    continente europeo). Se restan los ABSOLUTOS (numerador y denominador por
#    separado) y luego se divide; en las tasas se resta ponderando. Una region
#    sin dato ese año no se puede restar: queda dentro de la referencia y se
#    anota en "no_restadas".
# 3. Poblacion presente: para los indicadores de servicios o sucesos que
#    afectan tambien a los turistas (medicos, camas, homicidios, robos,
#    muertos en carretera) se calcula ademas una version por poblacion
#    presente = residentes + (noches en alojamientos + noches en
#    plataformas) / 365. Es un minimo: no cuenta segundas residencias ni
#    alojamientos no registrados.
# 4. Valor tipico: mediana de los ultimos 5 años en que hay dato a la vez para
#    la isla y su referencia, sin 2020 ni 2021 (pandemia; ver ANIOS_PANDEMIA)
#    (o suma de los 5 años, "agrupado", en los
#    indicadores de sucesos poco numerosos: homicidios, robos, muertos en
#    carretera y mortalidad infantil).
# 5. Comparacion: "indice" (isla / referencia * 100) para las intensidades
#    (por habitante, por superficie) o "diferencia" (isla - referencia) para
#    porcentajes, tasas y años. Ademas, una medida comun a todos para poder
#    compararlos entre si: z = (f(isla) - f(referencia)) / desviacion tipica
#    de f entre las regiones continentales de la UE, con f = logaritmo en los
#    indices y f = identidad en las diferencias.
# 6. Malta y Chipre no tienen parte continental: su referencia es la mediana
#    de las regiones continentales de la UE.
#
# Regiones continentales de la UE (contexto y desviacion tipica): todas las
# NUTS 2 de la UE salvo las 18 insulares, Ceuta, Melilla, Guyane, las de
# Irlanda (isla) y las dos danesas de la isla de Selandia (DK01, DK02).
#
# Genera: EUROSTAT_INSULARES_Indicadores.xlsx, hojas "catalogo", "valores"
# (indicador x geografia x año, con las referencias continentales XX_CONT),
# "comparacion" (isla x indicador) y "contexto_ue" (indicador).

library(openxlsx)
library(dplyr)
library(tidyr)
library(tibble)

ARCHIPIELAGOS <- c("ES53", "ES70", "PT20", "PT30", "EL41", "EL42", "EL62", "FI20", "MT00", "FRY1")
ISLAS_GRANDES <- c("ITG1", "ITG2", "FRM0", "EL43", "CY00", "FRY2", "FRY4", "FRY5")
INSULARES <- c(ARCHIPIELAGOS, ISLAS_GRANDES)
ISLAS_PAIS <- list(ES = c("ES53", "ES70"), PT = c("PT20", "PT30"),
                   EL = c("EL41", "EL42", "EL62", "EL43"), IT = c("ITG1", "ITG2"),
                   FR = c("FRM0", "FRY1", "FRY2", "FRY4", "FRY5"), FI = "FI20")
FUERA_PAIS <- list(ES = c("ES63", "ES64"), FR = "FRY3")
NO_CONTINENTAL_UE <- c("IE04", "IE05", "IE06", "DK01", "DK02", "ES63", "ES64", "FRY3")
UE <- c("AT", "BE", "BG", "CY", "CZ", "DE", "DK", "EE", "EL", "ES", "FI", "FR", "HR", "HU", "IE",
        "IT", "LT", "LU", "LV", "MT", "NL", "PL", "PT", "RO", "SE", "SI", "SK")
N_ANIOS <- 5
VENTANA_CREC <- 10   # años del crecimiento medio del PIB real
# Los años de la pandemia no entran en el valor tipico (2026-11-14): en 2020 y
# 2021 el turismo se hundio (en Illes Balears los turistas equivalian al 2 % de
# la poblacion en 2020, frente al 18 % de un año normal) y con ellos los
# indicadores turisticos, la poblacion presente y el PIB. Excepcion: si un
# indicador solo tiene datos de esos años (el Censo de 2021), se usan.
ANIOS_PANDEMIA <- c(2020, 2021)
elegir_anios <- function(anios) {
  a <- sort(unique(anios), decreasing = TRUE)
  sin <- setdiff(a, ANIOS_PANDEMIA)
  if (length(sin) > 0) a <- sin
  a[seq_len(min(N_ANIOS, length(a)))]
}

# --- Catalogo de indicadores --------------------------------------------------
# num/den: series de la extraccion; "a+b" = suma. pob_presente es virtual.
INDICADORES <- tribble(
  ~id, ~nombre, ~tema, ~unidad, ~tipo, ~num, ~den, ~escala, ~comparacion, ~agrupado,
  "crec_pob",      "Crecimiento de la población",          "Población", "‰ anual",           "cociente", "pob_crecimiento",  "pob_1enero",   1000, "diferencia", FALSE,
  "saldo_migr",    "Saldo migratorio",                      "Población", "‰ anual",           "cociente", "pob_saldo_migr",   "pob_1enero",   1000, "diferencia", FALSE,
  "crec_nat",      "Crecimiento natural",                   "Población", "‰ anual",           "cociente", "pob_crec_natural", "pob_1enero",   1000, "diferencia", FALSE,
  "pc_lt15",       "Menores de 15 años",                    "Población", "% de la población", "cociente", "pob_lt15",         "pob_total",     100, "diferencia", FALSE,
  "pc_ge65",       "65 años y más",                         "Población", "% de la población", "cociente", "pob_ge65",         "pob_total",     100, "diferencia", FALSE,
  "densidad",      "Densidad de población",                 "Población", "hab./km²",          "densidad", "pob_total",        "superficie_tierra", 1, "indice", FALSE,
  "nacidos_fuera", "Nacidos en el extranjero",              "Población", "% (15-74 años)",    "cociente", "lfs_nacidos_fuera","lfs_pob_15_74", 100, "diferencia", FALSE,
  "vab_A",         "VAB: agricultura y pesca",              "Estructura productiva", "% del VAB", "cociente", "vab_A",   "vab_TOTAL", 100, "diferencia", FALSE,
  "vab_B-E",       "VAB: industria y energía",              "Estructura productiva", "% del VAB", "cociente", "vab_B-E", "vab_TOTAL", 100, "diferencia", FALSE,
  "vab_F",         "VAB: construcción",                     "Estructura productiva", "% del VAB", "cociente", "vab_F",   "vab_TOTAL", 100, "diferencia", FALSE,
  "vab_G-I",       "VAB: comercio, transporte y hostelería","Estructura productiva", "% del VAB", "cociente", "vab_G-I", "vab_TOTAL", 100, "diferencia", FALSE,
  "vab_J",         "VAB: información y comunicaciones",     "Estructura productiva", "% del VAB", "cociente", "vab_J",   "vab_TOTAL", 100, "diferencia", FALSE,
  "vab_K",         "VAB: finanzas y seguros",               "Estructura productiva", "% del VAB", "cociente", "vab_K",   "vab_TOTAL", 100, "diferencia", FALSE,
  "vab_L",         "VAB: actividades inmobiliarias",        "Estructura productiva", "% del VAB", "cociente", "vab_L",   "vab_TOTAL", 100, "diferencia", FALSE,
  "vab_M_N",       "VAB: act. profesionales y administrativas", "Estructura productiva", "% del VAB", "cociente", "vab_M_N", "vab_TOTAL", 100, "diferencia", FALSE,
  "vab_O-Q",       "VAB: AAPP, educación y sanidad",        "Estructura productiva", "% del VAB", "cociente", "vab_O-Q", "vab_TOTAL", 100, "diferencia", FALSE,
  "vab_R-U",       "VAB: ocio y otros servicios",           "Estructura productiva", "% del VAB", "cociente", "vab_R-U", "vab_TOTAL", 100, "diferencia", FALSE,
  "pib_real_hab",  "PIB real por habitante",                "Economía", "PPA constantes de {A} por hab.", "pib_real", "pib_mpps", "pob_media", 1000, "indice", FALSE,
  "crec_pib_hab",  "Crecimiento del PIB real por habitante","Economía", "% medio anual, {A0}-{A}", "crec_real_hab", "pib_mpps", "pob_media", 100, "diferencia", FALSE,
  "crec_pib_total","Crecimiento del PIB real (total)",      "Economía", "% medio anual, {A0}-{A}", "crec_real_total", "pib_mpps", "pob_media", 100, "diferencia", FALSE,
  "renta_hab",     "Renta disponible de los hogares por hab.", "Economía", "PPA por hab.",    "cociente", "renta_hog_mpps",   "pob_media",    1000, "indice", FALSE,
  "productividad", "PIB por hora trabajada",                "Economía", "PPA por hora",       "cociente", "pib_mpps",         "horas_trabajadas", 1000, "indice", FALSE,
  "inversion",     "Inversión (formación bruta de capital)","Economía", "% del PIB",          "cociente", "fbcf_meur",        "pib_meur",      100, "diferencia", FALSE,
  "id_pib",        "Gasto en I+D",                          "Economía", "% del PIB",          "cociente", "id_meur",          "pib_meur",      100, "diferencia", FALSE,
  "natalidad_emp", "Natalidad empresarial",                 "Economía", "% de las empresas activas", "cociente", "empresas_nacidas", "empresas_activas", 100, "diferencia", FALSE,
  "empleo",        "Tasa de empleo (20-64 años)",           "Trabajo", "%",                   "cociente", "lfs_ocup_20_64",   "lfs_pob_20_64", 100, "diferencia", FALSE,
  "paro",          "Tasa de paro",                          "Trabajo", "% de los activos",    "cociente", "lfs_parados",      "lfs_activos",   100, "diferencia", FALSE,
  "paro_ld",       "Paro de larga duración",                "Trabajo", "% de los activos",    "cociente", "lfs_parados_ld",   "lfs_activos",   100, "diferencia", FALSE,
  "parcial",       "Empleo a tiempo parcial",               "Trabajo", "% del empleo",        "cociente", "lfs_ocup_parcial", "lfs_ocup_total",100, "diferencia", FALSE,
  "neet",          "Jóvenes que ni estudian ni trabajan (15-29)", "Trabajo", "%",             "tasa",     "neet_15_29_pc",    "pob_media",       1, "diferencia", FALSE,
  "noches_hab",    "Noches en alojamientos turísticos por hab.", "Turismo", "noches por hab.", "cociente", "noches_total",   "pob_media",   0.001, "indice", FALSE,
  "plataformas_hab","Noches en plataformas por hab.",       "Turismo", "noches por hab.",     "cociente", "noches_plataformas", "pob_media", 0.001, "indice", FALSE,
  "plazas_hab",    "Plazas turísticas por 1.000 hab.",      "Turismo", "por 1.000 hab.",      "cociente", "plazas_turisticas","pob_total",    1000, "indice", FALSE,
  "no_resid",      "Noches de no residentes",               "Turismo", "% de las noches",     "cociente", "noches_no_resid",  "noches_total",  100, "diferencia", FALSE,
  "estacionalidad","Noches en los 3 meses de más actividad","Turismo", "% de las noches del año", "estacionalidad", "noches_mensuales", NA, 100, "diferencia", FALSE,
  "viv_no_ocup",   "Viviendas no ocupadas (2021)",          "Vivienda", "% del parque",       "cociente", "viv_no_ocupadas",  "viv_total",     100, "diferencia", FALSE,
  "viv_recientes", "Viviendas construidas en 2011-2021",    "Vivienda", "% del parque",       "cociente", "viv_2011_2015+viv_desde_2016", "viv_total_periodo", 100, "diferencia", FALSE,
  "viv_propiedad", "Viviendas ocupadas en propiedad (2021)","Vivienda", "% de las ocupadas",  "cociente", "viv_propiedad",    "viv_propiedad+viv_alquiler+viv_otra_tenencia", 100, "diferencia", FALSE,
  "superiores",    "Estudios superiores (25-64 años)",      "Educación", "%",                 "cociente", "lfs_superiores",   "lfs_pob_25_64", 100, "diferencia", FALSE,
  "formacion",     "Formación continua (25-64 años)",       "Educación", "%",                 "tasa",     "formacion_cont_pc","lfs_pob_25_64",   1, "diferencia", FALSE,
  "esperanza",     "Esperanza de vida al nacer",            "Salud", "años",                  "tasa",     "esperanza_vida",   "pob_media",       1, "diferencia", FALSE,
  "mort_infantil", "Mortalidad infantil",                   "Salud", "‰ nacidos vivos",       "cociente", "muertes_infantiles","nacimientos",  1000, "diferencia", TRUE,
  "medicos",       "Médicos por 100.000 hab.",              "Salud", "por 100.000 hab.",      "cociente", "medicos",          "pob_media",     100, "indice", FALSE,
  "camas",         "Camas hospitalarias por 100.000 hab.",  "Salud", "por 100.000 hab.",      "cociente", "camas_hosp",       "pob_media",     100, "indice", FALSE,
  "arope",         "Riesgo de pobreza o exclusión social",  "Pobreza", "%",                   "tasa",     "arope_pc",         "pob_media",       1, "diferencia", FALSE,
  "privacion",     "Privación material y social severa",    "Pobreza", "%",                   "tasa",     "privacion_sev_pc", "pob_media",       1, "diferencia", FALSE,
  "sobrecarga",    "Sobrecarga del gasto en vivienda",      "Pobreza", "%",                   "tasa",     "sobrecarga_viv_pc","pob_media",       1, "diferencia", FALSE,
  "retrasos",      "Retrasos en pagos (hipoteca, alquiler, recibos)", "Pobreza", "%",         "tasa",     "retrasos_pagos_pc","pob_media",       1, "diferencia", FALSE,
  "pasajeros_hab", "Pasajeros aéreos por habitante",        "Transporte y seguridad", "pasajeros por hab.", "cociente", "pasajeros_aereos", "pob_media", 1, "indice", FALSE,
  "turismos_hab",  "Turismos por 1.000 hab.",               "Transporte y seguridad", "por 1.000 hab.", "cociente", "turismos", "pob_total",   1000, "indice", FALSE,
  "muertos_carr",  "Muertos en carretera por millón de hab.", "Transporte y seguridad", "por millón de hab.", "cociente", "muertos_carretera", "pob_media", 1000, "indice", TRUE,
  "homicidios",    "Homicidios por 100.000 hab.",           "Transporte y seguridad", "por 100.000 hab.", "cociente", "homicidios", "pob_media",   100, "indice", TRUE,
  "robos",         "Robos en vivienda por 100.000 hab.",    "Transporte y seguridad", "por 100.000 hab.", "cociente", "robos_vivienda", "pob_media", 100, "indice", TRUE
)
# Version por poblacion presente de los indicadores de servicios y sucesos
PRESENTE <- c("medicos", "camas", "muertos_carr", "homicidios", "robos")
INDICADORES <- bind_rows(
  INDICADORES,
  INDICADORES %>% filter(id %in% PRESENTE) %>%
    mutate(id = paste0(id, "_presente"), den = "pob_presente",
           nombre = paste(nombre, "(población presente)"))
) %>% mutate(orden = row_number())
stopifnot(!any(duplicated(INDICADORES$id)))

# --- Series -------------------------------------------------------------------
S <- read.xlsx("EUROSTAT_INSULARES_Series.xlsx", sheet = "series")
S <- S %>% filter(!grepl("ZZ$", geo))   # "extra-regio" (no es territorio)
universo <- S %>% filter(serie == "pob_media", nivel == "nuts2", nchar(geo) == 4,
                         substr(geo, 1, 2) %in% UE) %>% distinct(geo) %>% pull(geo)
CONTINENTAL_UE <- setdiff(universo, c(INSULARES, NO_CONTINENTAL_UE))
PAISES <- names(ISLAS_PAIS)
GEOS <- c(universo, PAISES)

# Año ancla del PIB real: el ultimo con PIB en PPA e indice de volumen para los
# 6 paises con islas y todas sus regiones insulares
ancla_de <- function(g) {
  a <- S %>% filter(serie == "pib_mpps", geo == g) %>% pull(anio)
  b <- S %>% filter(serie == "pib_vol_i20", geo == g) %>% pull(anio)
  max(intersect(a, b))
}
ANIO_ANCLA <- min(vapply(c(PAISES, unlist(ISLAS_PAIS)), ancla_de, numeric(1)))
ANIO_CREC0 <- ANIO_ANCLA - VENTANA_CREC
INDICADORES <- INDICADORES %>%
  mutate(unidad = gsub("{A0}", ANIO_CREC0, gsub("{A}", ANIO_ANCLA, unidad, fixed = TRUE), fixed = TRUE))

serie <- function(id) {
  if (id == "pob_presente") {
    p <- serie("pob_media")
    n <- serie("noches_total") %>% rename(n1 = valor) %>%
      full_join(serie("noches_plataformas") %>% rename(n2 = valor), by = c("geo", "anio"))
    return(p %>% inner_join(n, by = c("geo", "anio")) %>%
             transmute(geo, anio, valor = valor + (n1 + coalesce(n2, 0)) / 365 / 1000))
  }
  partes <- strsplit(id, "+", fixed = TRUE)[[1]]
  if (length(partes) > 1) {
    l <- lapply(partes, serie)
    out <- Reduce(function(a, b) inner_join(a, b, by = c("geo", "anio")), l)
    return(data.frame(geo = out$geo, anio = out$anio, valor = rowSums(out[, -(1:2), drop = FALSE])))
  }
  # Series anuales: sin mes, o con mes "TOTAL" (tour_ce_omn12 trae la dimension mes)
  x <- S %>% filter(serie == id, geo %in% GEOS, is.na(mes) | mes == "TOTAL") %>% select(geo, anio, valor)
  stopifnot(nrow(x) > 0, !any(duplicated(x[, c("geo", "anio")])))
  x
}

# Referencia continental: pais menos sus islas (y fuera), restando absolutos
restar_islas <- function(base, cols) {
  bind_rows(lapply(PAISES, function(p) {
    comps <- c(ISLAS_PAIS[[p]], FUERA_PAIS[[p]])
    nac <- base %>% filter(geo == p)
    sub <- base %>% filter(geo %in% comps) %>% group_by(anio) %>%
      summarise(across(all_of(cols), sum), restadas = paste(sort(geo), collapse = ","), .groups = "drop")
    nac %>% left_join(sub, by = "anio", suffix = c("", "_islas")) %>%
      mutate(across(all_of(cols), ~ .x - coalesce(get(paste0(cur_column(), "_islas")), 0)),
             no_restadas = vapply(restadas, function(r) paste(setdiff(comps, strsplit(coalesce(r, ""), ",")[[1]]), collapse = ","), character(1)),
             geo = paste0(p, "_CONT")) %>%
      select(geo, anio, all_of(cols), no_restadas)
  }))
}

calcular <- function(ind) {
  if (ind$tipo %in% c("cociente", "densidad")) {
    N <- serie(ind$num) %>% rename(num = valor)
    D <- serie(ind$den) %>% rename(den = valor)
    if (ind$tipo == "densidad") {   # superficie: la del ultimo año de cada geografia, para todos los años
      D <- D %>% group_by(geo) %>% filter(anio == max(anio)) %>% ungroup() %>% select(geo, den)
      base <- N %>% inner_join(D, by = "geo")
    } else {
      base <- N %>% inner_join(D, by = c("geo", "anio"))
    }
    out <- bind_rows(base %>% mutate(no_restadas = NA_character_), restar_islas(base, c("num", "den")))
    out %>% mutate(valor = ind$escala * num / den)
  } else if (ind$tipo %in% c("pib_real", "crec_real_hab", "crec_real_total")) {
    # PIB real en PPA constantes del año ancla (millones) y poblacion media (miles)
    ancla <- serie("pib_mpps") %>% filter(anio == ANIO_ANCLA) %>% select(geo, pps_ancla = valor)
    vol <- serie("pib_vol_i20")
    vol_ancla <- vol %>% filter(anio == ANIO_ANCLA) %>% select(geo, vol_ancla = valor)
    base <- vol %>% inner_join(ancla, by = "geo") %>% inner_join(vol_ancla, by = "geo") %>%
      transmute(geo, anio, num = pps_ancla * valor / vol_ancla) %>%
      inner_join(serie("pob_media") %>% rename(den = valor), by = c("geo", "anio"))
    real <- bind_rows(base %>% mutate(no_restadas = NA_character_), restar_islas(base, c("num", "den"))) %>%
      mutate(valor = 1000 * num / den)
    if (ind$tipo == "pib_real") return(real)
    # Crecimiento medio anual entre ANIO_CREC0 y ANIO_ANCLA (total o por habitante)
    real %>% filter(anio %in% c(ANIO_CREC0, ANIO_ANCLA)) %>%
      group_by(geo) %>% filter(n() == 2) %>% arrange(anio, .by_group = TRUE) %>%
      summarise(valor = 100 * ((if (ind$tipo == "crec_real_total") num[2] / num[1] else valor[2] / valor[1])^(1 / VENTANA_CREC) - 1),
                no_restadas = last(no_restadas), .groups = "drop") %>%
      mutate(anio = ANIO_ANCLA, num = NA_real_, den = NA_real_, periodo = paste0(ANIO_CREC0, "–", ANIO_ANCLA))
  } else if (ind$tipo == "tasa") {
    R <- serie(ind$num) %>% rename(tasa = valor)
    W <- serie(ind$den) %>% rename(peso = valor)
    base <- R %>% inner_join(W, by = c("geo", "anio")) %>% mutate(num = tasa * peso, den = peso)
    out <- bind_rows(base %>% mutate(no_restadas = NA_character_), restar_islas(base, c("num", "den")))
    out %>% mutate(valor = ind$escala * num / den) %>% select(-any_of(c("tasa", "peso")))
  } else if (ind$tipo == "estacionalidad") {
    M <- S %>% filter(serie == ind$num, geo %in% GEOS, !is.na(mes), mes != "TOTAL", anio >= 2022) %>%
      select(geo, anio, mes, valor)
    # Pais menos islas, mes a mes (absolutos), y despues el % de los 3 meses punta
    cont <- bind_rows(lapply(PAISES, function(p) {
      comps <- c(ISLAS_PAIS[[p]], FUERA_PAIS[[p]])
      M %>% filter(geo == p) %>%
        left_join(M %>% filter(geo %in% comps) %>% group_by(anio, mes) %>%
                    summarise(isl = sum(valor), restadas = paste(sort(geo), collapse = ","), .groups = "drop"),
                  by = c("anio", "mes")) %>%
        mutate(valor = valor - coalesce(isl, 0), geo = paste0(p, "_CONT"),
               no_restadas = vapply(restadas, function(r) paste(setdiff(comps, strsplit(coalesce(r, ""), ",")[[1]]), collapse = ","), character(1))) %>%
        select(geo, anio, mes, valor, no_restadas)
    }))
    bind_rows(M %>% mutate(no_restadas = NA_character_), cont) %>%
      group_by(geo, anio) %>% filter(n() == 12) %>%
      summarise(num = sum(sort(valor, decreasing = TRUE)[1:3]), den = sum(valor),
                no_restadas = first(no_restadas), .groups = "drop") %>%
      mutate(valor = ind$escala * num / den)
  }
}

valores <- bind_rows(lapply(seq_len(nrow(INDICADORES)), function(i) {
  ind <- INDICADORES[i, ]
  calcular(ind) %>% filter(is.finite(valor)) %>% mutate(indicador = ind$id, .before = 1)
}))
if (!"periodo" %in% names(valores)) valores$periodo <- NA_character_

# --- Valor tipico, contexto europeo y comparacion ------------------------------
f_trans <- function(v, comparacion, suelo) if (comparacion == "indice") log(pmax(v, suelo)) else v

tipico <- function(x, agrupado, escala) {   # x: filas de un indicador y una geografia, años ya elegidos
  if (agrupado) escala * sum(x$num) / sum(x$den) else median(x$valor)
}

contexto <- bind_rows(lapply(seq_len(nrow(INDICADORES)), function(i) {
  ind <- INDICADORES[i, ]
  v <- valores %>% filter(indicador == ind$id, geo %in% CONTINENTAL_UE) %>%
    group_by(geo) %>% filter(anio %in% elegir_anios(anio)) %>%
    summarise(t = tipico(cur_data(), ind$agrupado, ind$escala), .groups = "drop")
  suelo <- if (any(v$t > 0)) min(v$t[v$t > 0]) / 2 else NA
  ft <- f_trans(v$t, ind$comparacion, suelo)
  tibble(indicador = ind$id, n_regiones_ue = nrow(v), mediana_ue = median(v$t),
         sd_f_ue = sd(ft), mediana_f_ue = median(ft), suelo = suelo, valores_ue = list(v$t))
}))

pais_isla <- function(g) {
  p <- names(ISLAS_PAIS)[vapply(ISLAS_PAIS, function(v) g %in% v, logical(1))]
  if (length(p) == 0) NA_character_ else p
}

comparacion <- bind_rows(lapply(seq_len(nrow(INDICADORES)), function(i) {
  ind <- INDICADORES[i, ]
  ctx <- contexto %>% filter(indicador == ind$id)
  vi <- valores %>% filter(indicador == ind$id)
  bind_rows(lapply(INSULARES, function(g) {
    p <- pais_isla(g)
    xi <- vi %>% filter(geo == g)
    if (nrow(xi) == 0) return(NULL)
    if (is.na(p)) {             # Malta y Chipre: frente a la UE continental
      anios <- elegir_anios(xi$anio)
      t_i <- tipico(xi %>% filter(anio %in% anios), ind$agrupado, ind$escala)
      t_r <- ctx$mediana_ue; ref <- "UE_CONT"; no_rest <- NA_character_
    } else {
      ref <- paste0(p, "_CONT")
      xr <- vi %>% filter(geo == ref)
      comunes <- intersect(xi$anio, xr$anio)
      if (length(comunes) == 0) return(NULL)
      anios <- elegir_anios(comunes)
      t_i <- tipico(xi %>% filter(anio %in% anios), ind$agrupado, ind$escala)
      t_r <- tipico(xr %>% filter(anio %in% anios), ind$agrupado, ind$escala)
      no_rest <- paste(unique(na.omit(xr$no_restadas[xr$anio %in% anios & xr$no_restadas != ""])), collapse = ";")
    }
    comp <- if (ind$comparacion == "indice") 100 * t_i / t_r else t_i - t_r
    z <- (f_trans(t_i, ind$comparacion, ctx$suelo) - f_trans(t_r, ind$comparacion, ctx$suelo)) / ctx$sd_f_ue
    periodo <- unique(na.omit(xi$periodo[xi$anio %in% anios]))
    tibble(indicador = ind$id, geo = g, pais = coalesce(p, substr(g, 1, 2)), referencia = ref,
           anios = if (length(periodo) == 1) periodo else paste0(min(anios), "–", max(anios)),
           n_anios = length(anios),
           valor_isla = t_i, valor_ref = t_r, comparacion = comp, z = z,
           percentil_ue = 100 * mean(ctx$valores_ue[[1]] < t_i),
           no_restadas = no_rest)
  }))
})) %>%
  mutate(grupo = ifelse(geo %in% ARCHIPIELAGOS, "Archipiélago", "Isla grande"))

stopifnot(!any(duplicated(comparacion[, c("indicador", "geo")])))
message(nrow(INDICADORES), " indicadores; ", nrow(valores), " valores; ", nrow(comparacion), " comparaciones")
print(comparacion %>% count(indicador, name = "n_islas") %>% filter(n_islas < length(INSULARES)) %>% as.data.frame())

wb <- createWorkbook()
addWorksheet(wb, "catalogo"); writeData(wb, "catalogo", INDICADORES)
addWorksheet(wb, "valores"); writeData(wb, "valores", valores)
addWorksheet(wb, "comparacion"); writeData(wb, "comparacion", comparacion)
addWorksheet(wb, "contexto_ue"); writeData(wb, "contexto_ue", contexto %>% select(-valores_ue))
saveWorkbook(wb, "EUROSTAT_INSULARES_Indicadores.xlsx", overwrite = TRUE)
