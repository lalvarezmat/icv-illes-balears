# INE_ICV_CCAA_Procesar.R
#
# Procesa (parsea/pivota) las tablas del ICV que tienen desagregacion por
# Comunidad Autonoma (ambito_geografico == "Comunidad autónoma" en
# INE_ICV.catalogo.xlsx, 58 tablas) Y CUYO fecha_fin >= 2023-01-01 (42 de
# esas 58 -- petición explícita del usuario, quedarse solo con lo
# actualizado despues de 2023 incluido) y genera, EN OTRA CARPETA
# (../INE_ICV_CCAA), el mismo estilo de fichero que se uso en IBESTAT:
#
#   <slug>.xlsx            una hoja "datos" (o "datos_<categoria>" si la
#                          tabla tiene, ademas de CCAA, otra dimension propia
#                          del indicador -- ocurre en 32 de las 58, y esa
#                          dimension SIEMPRE es la propia categoria de
#                          respuesta del indicador, nunca una variable
#                          demografica; verificado en vivo, ver conversacion)
#                          con fecha + una columna por Comunidad Autonoma.
#   <slug>_METADATOS.xlsx   hojas "resumen" (fechas, frecuencia, tema,
#                          dimension/subdimension, nota, url), "geografia"
#                          (codigo INE + nombre de cada CCAA), "medida"
#                          (unidad, decimales, contenido) y "categorias"
#                          (solo si hay dimension extra: hoja -> nombre de
#                          categoria COMPLETO, sin truncar -- el nombre de
#                          hoja del <slug>.xlsx si esta truncado a 31
#                          caracteres por el limite de Excel, pero el valor
#                          de la categoria en si (antes de convertirlo en
#                          nombre de hoja) no lo esta, se guarda tal cual
#                          sale del .px. Lo usa INE_ICV_CCAA_Consolidar.R
#                          para poner el nombre completo en la columna
#                          "atributo" del consolidado en vez del recortado).
#
# Nombre de fichero (<slug>): version corta de la columna "tema" del
# catalogo (pedido explicito del usuario). Las 58 tablas CCAA tienen 56
# temas distintos -- el unico caso con 3 tablas por el mismo tema
# (contaminacion por PM10/PM2,5) se desambigua anadiendo un sufijo tomado
# del titulo de tabla.
#
# Fuente: los .px YA DESCARGADOS en DATA/ (lectura en local, sin red).
# Reanudable: si ya existen <slug>.xlsx y <slug>_METADATOS.xlsx se omite esa
# tabla.

library(pxR)
library(dplyr)
library(stringr)
library(stringi)
library(openxlsx)
library(tidyr)

CATALOGO <- "INE_ICV.catalogo.xlsx"
DATA_DIR <- "DATA"
OUT_DIR <- file.path("..", "INE_ICV_CCAA")
if (!dir.exists(OUT_DIR)) dir.create(OUT_DIR)

GEO_DIMS <- c("Comunidad autónoma", "Ámbito geográfico")

catalogo <- openxlsx::read.xlsx(CATALOGO, sheet = "catalogo", detectDates = TRUE)
ccaa <- catalogo %>%
  filter(ambito_geografico == "Comunidad autónoma", fecha_fin >= as.Date("2023-01-01"))

slugify <- function(x, maxlen = 45) {
  s <- stri_trans_general(x, "Latin-ASCII")
  s <- str_to_lower(s)
  s <- str_replace_all(s, "[^a-z0-9]+", "_")
  s <- str_replace_all(s, "^_+|_+$", "")
  str_sub(s, 1, maxlen)
}

# Desambiguacion en dos pasadas: 1) si el tema se repite, se anade un trozo
# del titulo de tabla; 2) si AUN asi coincide (ocurre cuando varias tablas
# del mismo tema tienen ademas titulos con el mismo prefijo largo -- p.ej.
# las 3 de contaminacion PM10/PM2,5, o las de "Nivel de formación..."), se
# garantiza unicidad anadiendo el id numerico de tabla (tabla_t).
ccaa <- ccaa %>%
  mutate(slug = slugify(tema)) %>%
  group_by(slug) %>%
  mutate(slug = if (n() > 1) paste0(slug, "_", slugify(tabla, 20)) else slug) %>%
  ungroup() %>%
  group_by(slug) %>%
  mutate(slug = if (n() > 1) paste0(slug, "_", tabla_t) else slug) %>%
  ungroup()

stopifnot(!any(duplicated(ccaa$slug)))

# categorias: vector con un valor por cada dimension extra que tenga la
# tabla (0, 1 o mas -- ver bug corregido en la cabecera del bucle principal).
# Se reparte el presupuesto de 31 caracteres (limite de nombre de hoja de
# Excel) a partes iguales entre las categorias para que quepan todas.
sheet_name <- function(prefijo, categorias) {
  if (length(categorias) == 0) return(prefijo)
  budget <- max(4, floor((31 - nchar(prefijo) - 1 - (length(categorias) - 1)) / length(categorias)))
  partes <- vapply(categorias, function(c) slugify(c, budget), character(1))
  s <- paste0(prefijo, "_", paste(partes, collapse = "_"))
  str_sub(s, 1, 31)
}

n_ok <- 0L
n_ya <- 0L
n_fail <- 0L

for (i in seq_len(nrow(ccaa))) {
  slug <- ccaa$slug[i]
  t_id <- ccaa$tabla_t[i]
  fn_datos <- file.path(OUT_DIR, paste0(slug, ".xlsx"))
  fn_meta <- file.path(OUT_DIR, paste0(slug, "_METADATOS.xlsx"))

  if (file.exists(fn_datos) && file.exists(fn_meta)) { n_ya <- n_ya + 1L; next }

  px <- tryCatch(read.px(file.path(DATA_DIR, paste0(t_id, ".px"))), error = function(e) NULL)
  if (is.null(px)) { n_fail <- n_fail + 1L; next }

  datos <- as.data.frame(px) %>% na.omit()
  cols <- setdiff(names(datos), c("Periodo", "value"))
  geo_col <- cols[str_replace_all(cols, "\\.", " ") %in% GEO_DIMS][1]
  # character(0) si no hay dimension extra; puede haber MAS de una (bug
  # corregido el 2026-09-20: antes se cogia solo la primera con [1] y las
  # demas quedaban sin pivotar, mezcladas como texto separado por comas en
  # cada celda -- ver nota en CLAUDE.md sobre la tabla t=69774).
  extra_cols <- setdiff(cols, geo_col)

  datos <- datos %>% mutate(fecha = as.Date(paste0(Periodo, "-01-01"))) %>% select(-Periodo)

  # una fila de combinaciones por cada cruce de categorias de las dimensiones
  # extra (producto cartesiano); una sola fila vacia si no hay ninguna.
  combinaciones <- if (length(extra_cols) == 0) {
    tibble(.dummy = 1)
  } else {
    # as.character() es imprescindible: las columnas de dimension que trae
    # pxR son factor, y c() sobre un factor devuelve el codigo entero, no la
    # etiqueta (bug encontrado el 2026-09-20: los nombres de hoja salian
    # "datos_1_1", "datos_2_1"... en vez del nombre de cada categoria).
    niveles <- lapply(extra_cols, function(ec) sort(unique(as.character(datos[[ec]]))))
    names(niveles) <- extra_cols
    expand.grid(niveles, stringsAsFactors = FALSE, KEEP.OUT.ATTRS = FALSE)
  }

  wb <- createWorkbook()
  hojas_usadas <- character(0)
  categorias_completas <- tibble(hoja = character(0), categoria_completa = character(0))
  for (j in seq_len(nrow(combinaciones))) {
    d <- datos
    categ_partes <- character(0)
    for (ec in extra_cols) {
      val <- combinaciones[[ec]][j]
      d <- d %>% filter(.data[[ec]] == val)
      categ_partes <- c(categ_partes, val)
    }
    ancho <- d %>%
      select(fecha, all_of(geo_col), value) %>%
      pivot_wider(id_cols = fecha, names_from = all_of(geo_col), values_from = value) %>%
      arrange(fecha)
    # quita columnas geograficas 100% vacias (mismo criterio que IBESTAT)
    vacias <- names(ancho)[-1][sapply(ancho[-1], function(c) all(is.na(c)))]
    if (length(vacias) > 0) ancho <- ancho %>% select(-all_of(vacias))

    hoja <- sheet_name("datos", categ_partes)
    stopifnot(!hoja %in% hojas_usadas) # nombres de hoja distintos tras truncar a 31 caracteres
    hojas_usadas <- c(hojas_usadas, hoja)
    addWorksheet(wb, hoja)
    writeData(wb, hoja, ancho)
    if (length(categ_partes) > 0) {
      categorias_completas <- bind_rows(categorias_completas,
        tibble(hoja = hoja, categoria_completa = paste(categ_partes, collapse = " - ")))
    }
  }
  saveWorkbook(wb, fn_datos, overwrite = TRUE)

  # --- metadatos ---
  geo_codes <- px$CODES[[geo_col]]
  geo_names <- px$VALUES[[geo_col]]
  geografia <- tibble(codigo = geo_codes, nombre = geo_names)

  resumen <- tibble(
    campo = c("tabla_t", "tabla", "tema", "dimension", "subdimension", "fecha_inicio",
              "fecha_fin", "frecuencia", "n_periodos", "n_obs", "dimension_extra", "nota", "url_tabla"),
    valor = c(
      as.character(t_id), ccaa$tabla[i], ccaa$tema[i], ccaa$dimension[i], ccaa$subdimension[i],
      format(ccaa$fecha_inicio[i]), format(ccaa$fecha_fin[i]), ccaa$frecuencia[i],
      as.character(ccaa$n_periodos[i]), as.character(ccaa$n_obs[i]),
      if (length(extra_cols) == 0) "" else paste(str_replace_all(extra_cols, "\\.", " "), collapse = " x "),
      ifelse(is.null(px$NOTE$value), "", px$NOTE$value),
      paste0("https://www.ine.es/jaxiT3/Tabla.htm?t=", t_id)
    )
  )

  medida <- tibble(
    campo = c("unidad", "decimales", "contenido"),
    valor = c(
      ifelse(is.null(px$UNITS$value), "", px$UNITS$value),
      ifelse(is.null(px$DECIMALS$value), "", px$DECIMALS$value),
      ifelse(is.null(px$CONTENTS$value), "", px$CONTENTS$value)
    )
  )

  wb2 <- createWorkbook()
  addWorksheet(wb2, "resumen"); writeData(wb2, "resumen", resumen)
  addWorksheet(wb2, "geografia"); writeData(wb2, "geografia", geografia)
  addWorksheet(wb2, "medida"); writeData(wb2, "medida", medida)
  if (nrow(categorias_completas) > 0) {
    addWorksheet(wb2, "categorias"); writeData(wb2, "categorias", categorias_completas)
  }
  saveWorkbook(wb2, fn_meta, overwrite = TRUE)

  n_ok <- n_ok + 1L
}

cat("Procesadas:", n_ok, "| ya existian:", n_ya, "| fallos:", n_fail, "| total CCAA:", nrow(ccaa), "\n")
