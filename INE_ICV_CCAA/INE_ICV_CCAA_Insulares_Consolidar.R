# INE_ICV_CCAA_Insulares_Consolidar.R
#
# Indicadores NUEVOS del estudio de los archipielagos
# (INE_ICV_CCAA_Insulares.Rmd), con el mismo formato que
# INE_ICV_CCAA_Consolidado.xlsx (una fila por indicador y año, una columna
# por geografia) para que el Rmd pueda apilarlos con los 82 del informe
# principal. NO toca INE_ICV_CCAA_Consolidado.xlsx: el informe publicado
# (INE_ICV_CCAA_all.Rmd) sigue con sus 82 indicadores.
#
# Indicadores (15):
#   - Poblacion y vivienda:
#       crecimiento de la poblacion (% anual)            ../POBLACION_CCAA
#       viviendas iniciadas por 1.000 habitantes         ../MIVAU
#       compraventas de vivienda por no residentes
#         por 1.000 habitantes                           ../MIVAU
#   - Turismo:
#       viviendas turisticas por 1.000 habitantes        ../VIVIENDA_TURISTICA (INE)
#       pernoctaciones en alojamientos reservados en
#         plataformas por habitante                      ../VIVIENDA_TURISTICA (Eurostat)
#   - Estructura productiva:
#       peso de cada una de las 10 ramas A10 en el VAB   ../PIB_CCAA/PIB_CCAA_VAB_Ramas.xlsx
#
# "Por habitante" = entre la poblacion residente a 1 de enero del mismo año
# (ECP del INE, ../POBLACION_CCAA). Es poblacion RESIDENTE: en las
# comunidades turisticas la poblacion presente a lo largo del año es mayor,
# y eso es justo lo que estos cocientes quieren reflejar (presion por
# residente), no un error.
#
# Viviendas turisticas (INE): la estadistica es una foto en 2-3 fechas al
# año (ver ../VIVIENDA_TURISTICA); el valor anual es la media de las fotos
# del año, y solo se guardan los años con al menos 2 fotos (2021-2025: 2020
# solo tiene agosto y 2026, mayo).
#
# Se empieza en ANIO_INI = 2004, como el informe principal; las series que
# empiezan despues (viviendas iniciadas 2008, compraventas 2007, plataformas
# 2018, viviendas turisticas 2021) empiezan en su primer año.
#
# Requisitos (ejecutar antes, desde su carpeta):
#   ../POBLACION_CCAA/POBLACION_CCAA_DataExtraction.R
#   ../MIVAU/MIVAU_DataExtraction.R
#   ../VIVIENDA_TURISTICA/VIVIENDA_TURISTICA_DataExtraction.R
#   ../PIB_CCAA/PIB_CCAA_VAB_Ramas.R
#
# Genera: INE_ICV_CCAA_Insulares_Indicadores.xlsx (hoja "datos")

library(openxlsx)
library(dplyr)
library(tidyr)

ANIO_INI <- 2004

# Columnas de geografia exactamente como las tiene el consolidado principal
# (read.xlsx las devuelve con puntos: "04.Balears,.Illes")
geo_cols <- setdiff(names(read.xlsx("INE_ICV_CCAA_Consolidado.xlsx", sheet = "datos", rows = 1:2)),
                    c("atributo", "nombre_corto", "dimension", "fuente", "unidad", "fecha"))
stopifnot(length(geo_cols) == 20, geo_cols[1] == "Total.Nacional")
codigo_a_columna <- setNames(geo_cols, c("00", substr(geo_cols[-1], 1, 2)))

# largo (codigo_ccaa, anio, valor) -> ancho con las columnas del consolidado
a_ancho <- function(df, atributo, nombre_corto, dimension, fuente, unidad) {
  df %>%
    filter(anio >= ANIO_INI, codigo_ccaa %in% names(codigo_a_columna)) %>%
    mutate(columna = unname(codigo_a_columna[codigo_ccaa])) %>%
    select(fecha = anio, columna, valor) %>%
    pivot_wider(names_from = columna, values_from = valor) %>%
    arrange(fecha) %>%
    mutate(atributo = atributo, nombre_corto = nombre_corto, dimension = dimension,
           fuente = fuente, unidad = unidad) %>%
    select(atributo, nombre_corto, dimension, fuente, unidad, fecha, any_of(geo_cols))
}

pob <- read.xlsx("../POBLACION_CCAA/POBLACION_CCAA.xlsx", sheet = "datos")
por_hab <- function(df, col, escala) {
  df %>%
    inner_join(pob %>% select(codigo_ccaa, anio, poblacion), by = c("codigo_ccaa", "anio")) %>%
    mutate(valor = escala * .data[[col]] / poblacion)
}

# --- Poblacion y vivienda -----------------------------------------------------
crec <- pob %>% filter(!is.na(crecimiento_pct)) %>% transmute(codigo_ccaa, anio, valor = crecimiento_pct)

viv <- read.xlsx("../MIVAU/MIVAU_Viviendas.xlsx", sheet = "datos") %>%
  filter(anio >= 2008) %>%                       # antes de 2008 solo hay Canarias
  por_hab("iniciadas", 1000)

trans <- read.xlsx("../MIVAU/MIVAU_Transacciones.xlsx", sheet = "datos") %>%
  por_hab("no_residentes", 1000)

# --- Turismo ------------------------------------------------------------------
vut <- read.xlsx("../VIVIENDA_TURISTICA/VIVIENDA_TURISTICA_CCAA.xlsx", sheet = "ine") %>%
  group_by(codigo_ccaa, anio) %>%
  summarise(viviendas = mean(viviendas), n_fotos = n(), .groups = "drop") %>%
  filter(n_fotos >= 2) %>%
  por_hab("viviendas", 1000)

plat <- read.xlsx("../VIVIENDA_TURISTICA/VIVIENDA_TURISTICA_CCAA.xlsx", sheet = "eurostat") %>%
  por_hab("pernoctaciones", 1)

# --- Estructura productiva ----------------------------------------------------
RAMAS <- c(
  A   = "VAB: agricultura y pesca",
  B_E = "VAB: industria y energía",
  F   = "VAB: construcción",
  G_I = "VAB: comercio, transporte y hostelería",
  J   = "VAB: información y comunicaciones",
  K   = "VAB: finanzas y seguros",
  L   = "VAB: actividades inmobiliarias",
  M_N = "VAB: act. profesionales y administrativas",
  O_Q = "VAB: AAPP, educación y sanidad",
  R_U = "VAB: ocio y otros servicios"
)
vab <- read.xlsx("../PIB_CCAA/PIB_CCAA_VAB_Ramas.xlsx", sheet = "datos")
stopifnot(setequal(unique(vab$rama_codigo), names(RAMAS)))
filas_vab <- bind_rows(lapply(names(RAMAS), function(r) {
  x <- vab %>% filter(rama_codigo == r)
  a_ancho(x %>% transmute(codigo_ccaa, anio, valor = pct_vab),
          atributo = paste0("Peso en el VAB de la rama ", unique(x$rama)),
          nombre_corto = RAMAS[[r]], dimension = "Estructura productiva",
          fuente = "INE", unidad = "% del VAB")
}))

nuevos <- bind_rows(
  a_ancho(crec, "Crecimiento de la población residente durante el año (1 de enero a 1 de enero)",
          "Crecimiento de la población", "Población y vivienda", "INE", "% anual"),
  a_ancho(viv, "Viviendas iniciadas (libres y protegidas) por 1.000 habitantes",
          "Viviendas iniciadas", "Población y vivienda", "MIVAU", "Por 1.000 hab."),
  a_ancho(trans, "Compraventas de vivienda por compradores no residentes en España por 1.000 habitantes",
          "Compras no residentes", "Población y vivienda", "MIVAU", "Por 1.000 hab."),
  a_ancho(vut, "Viviendas turísticas anunciadas en plataformas por 1.000 habitantes (media de las fotos del año)",
          "Viviendas turísticas", "Turismo", "INE", "Por 1.000 hab."),
  a_ancho(plat, "Pernoctaciones en alojamientos de corta estancia reservados en plataformas en línea por habitante",
          "Noches en plataformas", "Turismo", "Eurostat", "Noches por habitante"),
  filas_vab
)

# Guardas: un valor por (indicador, año); las 17 CCAA completas en cada año;
# nombres cortos que no chocan con los 82 del consolidado principal
stopifnot(nrow(count(nuevos, nombre_corto, fecha) %>% filter(n > 1)) == 0)
ccaa17 <- geo_cols[2:18]
stopifnot(!anyNA(nuevos[, ccaa17]))
principales <- unique(read.xlsx("INE_ICV_CCAA_Consolidado.xlsx", sheet = "datos")$nombre_corto)
stopifnot(!any(unique(nuevos$nombre_corto) %in% principales))

wb <- createWorkbook()
addWorksheet(wb, "datos")
writeData(wb, "datos", nuevos)
setColWidths(wb, "datos", cols = 1:ncol(nuevos), widths = "auto")
freezePane(wb, "datos", firstRow = TRUE, firstCol = TRUE)
saveWorkbook(wb, "INE_ICV_CCAA_Insulares_Indicadores.xlsx", overwrite = TRUE)

cat(nrow(nuevos), "filas,", n_distinct(nuevos$nombre_corto),
    "indicadores en INE_ICV_CCAA_Insulares_Indicadores.xlsx\n")
print(nuevos %>% group_by(nombre_corto) %>% summarise(ini = min(fecha), fin = max(fecha), .groups = "drop"),
      n = 20)
