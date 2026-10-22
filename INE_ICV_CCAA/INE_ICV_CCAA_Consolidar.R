# INE_ICV_CCAA_Consolidar.R
#
# Junta en un unico fichero (INE_ICV_CCAA_Consolidado.xlsx, una sola hoja
# "datos") el contenido de las 42 tablas de este directorio, apiladas por
# filas en el mismo orden que INE_ICV_CCAA.catalogo.xlsx (que ya viene
# ordenado desc(fecha_fin), desc(n_periodos) -- ejecutar
# INE_ICV_CCAA_ConstruirCatalogo.R antes si el catalogo no existe o esta
# desactualizado).
#
# Columnas del resultado: atributo | dimension | fuente | unidad | fecha |
# <una por comunidad autonoma, mismo formato ancho que las tablas de origen>.
#
#   dimension   la DIMENSION del marco ICV del INE a la que pertenece el
#               atributo (columna "dimension" de INE_ICV_CCAA.catalogo.xlsx,
#               que a su vez viene del .px de origen): "Condiciones
#               materiales de vida", "Trabajo", "Salud", "Educación",
#               "Ocio y relaciones sociales", "Seguridad física y personal",
#               "Entorno y medioambiente", "Experiencia general de la vida".
#               Añadida el 2026-09-21 para poder agrupar los atributos por
#               tema en INE_ICV_CCAA_Analisis.Rmd (79 indicadores sueltos son
#               ilegibles; por dimension son 8 bloques con sentido).
#   fuente      organismo/origen del atributo: "INE" (las 42 tablas del ICV,
#               vía INE_ICV_CCAA_Procesar.R), "REGISTRADORES" (precio de
#               compraventa, Colegio de Registradores), "SERPAVI" (precio de
#               alquiler, Ministerio de Vivienda) o "PIB_CCAA" (PIB per
#               cápita real, construido a partir de 3 tablas del INE -- ver
#               bloque "PIB_CCAA" más abajo). Añadida el 2026-09-22 al
#               incorporar precio de vivienda -- ver también el bloque
#               "REGISTRADORES + SERPAVI" para el detalle de esas 2 fuentes.
#               Desde el 2026-09-25 la renta media y mediana y los dos precios
#               de vivienda van en euros CONSTANTES (deflactados con el IPC de
#               cada CCAA a partir del primer año de cada serie), y su atributo
#               lo dice entre paréntesis -- ver el bloque "IPC_CCAA".
#   atributo    nombre del atributo medido. Si la tabla tiene una sola hoja
#               de datos, es el "tema" del catalogo. Si tiene varias (ver
#               columna n_atributos del catalogo), se añade cada hoja como
#               fila independiente, con "tema" + " - " + el nombre COMPLETO
#               de la categoria (columna "categoria_completa" de la hoja
#               "categorias" de cada <slug>_METADATOS.xlsx -- añadida el
#               2026-09-21 en INE_ICV_CCAA_Procesar.R precisamente para esto,
#               nunca el nombre de hoja truncado a 31 caracteres que se usaba
#               antes). De paso este nombre completo ya distingue por si solo
#               los casos que antes colisionaban (ver bug de PM10/PM2,5 mas
#               abajo), asi que no hace falta ninguna coletilla añadida a
#               mano.
#   fecha       AÑO (entero), no Date: las 42 tablas son anuales (ver
#               comprobacion y motivo en INE_ICV_CCAA_ConstruirCatalogo.R).
#   unidad      la de la hoja "medida" de cada tabla. Se aplica la MISMA
#               unidad a todos los atributos de una tabla que tenga varias
#               hojas, aunque alguna de sus hojas mida en realidad otra cosa
#               -- ocurre en 4 de las 42 tablas: satisfaccion_con_el_trabajo
#               (4 hojas en %, 1 en "Número medio") y las 3 tablas de
#               contaminacion urbana poblacion_urbana_expuesta_... (µg/m3,
#               miles de personas y nº de estaciones mezclados). La hoja
#               "medida" de origen no guarda la unidad por hoja, solo una
#               por tabla, así que no se puede separar sin volver a mirar la
#               tabla INE original.
#
# Las 42 tablas no tienen todas las mismas columnas geograficas:
# salarios_bajos.xlsx no trae Ceuta ni Melilla (18 comunidades en vez de
# 20) -- quedan como NA en esas dos columnas para esa tabla.
#
# Bug ya corregido (detectado el 2026-09-21 al preparar nombres cortos para
# los atributos, antes de que existiera la hoja "categorias"): las 3 tablas
# de contaminacion urbana PM10/PM2,5 (fichero *_74699, *_74702, *_74705)
# comparten el mismo "tema" Y ademas 74699/74702 usaban los MISMOS sufijos de
# hoja truncados ("datos_media_ponderada_con_la_po" etc.), asi que el
# "atributo" resultante era identico para la media anual de PM10 (74699) y
# de PM2,5 (74702) -- se apilaban 42 filas (2 por año) bajo el mismo rotulo,
# sin forma de saber cual era cual. Al pasar a usar categoria_completa (que
# ya dice "PM10"/"PM2.5" explicitamente en su texto) el problema desaparece
# solo; se deja la guarda de cordura de mas abajo por si algun otro caso
# futuro vuelve a colisionar.
#
# Reanudable/regenerable: no llama a ninguna red, solo relee lo que ya esta
# en el directorio.

library(openxlsx)
library(dplyr)
library(tidyr)

CATALOGO <- "INE_ICV_CCAA.catalogo.xlsx"
EXCLUIDOS <- character(0) # nivel_de_formacion..._69774 ya corregido el 2026-09-20 (bug de pivote en INE_ICV_CCAA_Procesar.R), vuelve a incluirse
catalogo <- read.xlsx(CATALOGO, sheet = "catalogo") %>%
  filter(!fichero %in% EXCLUIDOS)

# Atributos individuales excluidos del consolidado (a diferencia de EXCLUIDOS
# arriba, que quita una tabla entera): peticion explicita del usuario el
# 2026-09-21 de quitar, de las 3 tablas de contaminacion urbana PM10/PM2,5,
# las hojas de "poblacion considerada"/"estaciones consideradas" -- son el
# TAMAÑO DE MUESTRA usado para calcular la media ponderada (nº de estaciones,
# nº de personas cubiertas), no una medida de contaminacion en si, asi que no
# tiene sentido compararlas entre Balears y España como si fueran un
# indicador mas. Se queda solo la hoja "media ponderada" de cada una de las 3
# tablas (PM10 anual, PM2,5 anual, PM10 dias de superacion).
ATRIBUTOS_EXCLUIDOS <- c(
  "Población urbana expuesta a contaminación del aire (micropartículas PM10, PM2,5) - Estaciones consideradas para media ponderada de concentración media anual PM10",
  "Población urbana expuesta a contaminación del aire (micropartículas PM10, PM2,5) - Población considerada para la media ponderada de concentración media anual PM10",
  "Población urbana expuesta a contaminación del aire (micropartículas PM10, PM2,5) - Estaciones consideradas para media ponderada de concentración media anual PM2.5",
  "Población urbana expuesta a contaminación del aire (micropartículas PM10, PM2,5) - Población considerada para la media ponderada de concentración media anual PM2.5",
  "Población urbana expuesta a contaminación del aire (micropartículas PM10, PM2,5) - Estaciones en el nº días que superan concentración límite diaria PM10",
  "Población urbana expuesta a contaminación del aire (micropartículas PM10, PM2,5) - Población considerada en el nº días/año que superan concentración límite diaria de PM10"
)

fila_atributo <- function(fichero, tema, dimension, unidad, hoja, multiple, categorias) {
  datos <- read.xlsx(paste0(fichero, ".xlsx"), sheet = hoja, detectDates = TRUE)
  atributo <- if (multiple) {
    categoria_completa <- categorias$categoria_completa[categorias$hoja == hoja]
    stopifnot(length(categoria_completa) == 1) # la hoja "categorias" debe tener exactamente una fila por hoja de datos
    paste0(tema, " - ", categoria_completa)
  } else {
    tema
  }
  bind_cols(
    data.frame(atributo = atributo, dimension = dimension, unidad = unidad, stringsAsFactors = FALSE),
    datos
  )
}

bloques <- lapply(seq_len(nrow(catalogo)), function(i) {
  fichero <- catalogo$fichero[i]
  tema <- catalogo$tema[i]
  dimension <- catalogo$dimension[i]

  mf <- paste0(fichero, "_METADATOS.xlsx")
  medida <- read.xlsx(mf, sheet = "medida")
  unidad <- medida$valor[medida$campo == "unidad"]

  hojas <- grep("^datos", getSheetNames(paste0(fichero, ".xlsx")), value = TRUE)
  multiple <- length(hojas) > 1
  categorias <- if (multiple) read.xlsx(mf, sheet = "categorias") else NULL

  bind_rows(lapply(hojas, function(h) fila_atributo(fichero, tema, dimension, unidad, h, multiple, categorias)))
})

consolidado <- bind_rows(bloques)
consolidado$fecha <- as.integer(format(consolidado$fecha, "%Y"))
consolidado$fuente <- "INE"

# --- REGISTRADORES (compraventa) + SERPAVI (alquiler): precio de vivienda ---
#
# Ninguna de las 42 tablas del INE mide precio de vivienda en euros/m2 (solo
# indicadores indirectos: % gasto en vivienda, hacinamiento...), asi que se
# incorporan como dos atributos nuevos, dimension "Condiciones materiales de
# vida", con las mismas columnas geograficas (CCAA en ancho) que el resto.
# Peticion explicita del usuario 2026-09-22. Fuentes generadas por
# REGISTRADORES/REGISTRADORES_DataExtraction.R y
# SERPAVI/SERPAVI_DataExtraction.R -- ejecutarlos antes si no existen o estan
# desactualizados.
#
# El codigo de CCAA (1-19) de ambas fuentes coincide exactamente con el
# usado en las columnas geograficas del consolidado (01 Andalucia .. 17 La
# Rioja, 18 Ceuta, 19 Melilla) -- verificado en vivo en ambas fuentes, no
# solo supuesto. Por eso la columna destino de cada codigo se deriva del
# propio nombre de columna con una expresion regular (prefijo numerico),
# en vez de hardcodear los 19 nombres a mano (serian faciles de desincronizar
# si cambia la grafia de alguna CCAA).
geo_cols <- setdiff(names(consolidado), c("atributo", "dimension", "unidad", "fecha", "fuente"))
NACIONAL_COL <- "Total.Nacional"
ccaa_cols <- setdiff(geo_cols, NACIONAL_COL)
stopifnot("Se esperaban 19 columnas de CCAA (sin Total Nacional)" = length(ccaa_cols) == 19)
codigo_a_columna <- setNames(ccaa_cols, as.integer(sub("^([0-9]{2})\\..*", "\\1", ccaa_cols)))

completar_columnas_geo <- function(df) {
  faltan <- setdiff(geo_cols, names(df))
  for (col in faltan) df[[col]] <- NA_real_
  df
}

# --- IPC_CCAA: renta y precios de vivienda en euros constantes -------------
#
# Peticion explicita del usuario 2026-09-25: los 4 atributos en euros
# corrientes (renta media y mediana por unidad de consumo, precio de
# compraventa y de alquiler) se deflactan con el IPC de CADA comunidad (el
# nacional para Total.Nacional), no con uno solo. El PIB per capita no se toca:
# ya viene en euros constantes desde PIB_CCAA (indices de volumen, anclados
# tambien en el ultimo año de su serie).
#
# Año base: el IPC vale 100 en el ULTIMO año de cada serie, para que el ultimo
# dato (el mas reciente, el que se compara con las cifras publicadas hoy) no
# cambie y los anteriores se deflacten hacia atras a euros de ese año.
# Peticion del usuario 2026-09-26; hasta entonces la base era el PRIMER año
# (euros de 2007 en renta y compraventa, de 2011 en alquiler). La base es
# comun a todas las CCAA de un mismo atributo; si no, cada columna estaria en
# euros de un año distinto y no se podrian comparar. En los 4 atributos las 17
# CCAA terminan el mismo año (se comprueba abajo), asi que el ultimo dato de
# cada una coincide con el nominal.
#
# Desfase: la renta de la Encuesta de Condiciones de Vida del año t son los
# ingresos del año t-1, asi que se deflacta con el IPC de t-1 (desfase = 1) y
# su año base es el de esos ingresos (encuesta 2024 -> euros de 2023). Los
# precios de vivienda son del propio año (desfase = 0, euros de 2024).
#
# OJO: el IPC por CCAA corrige la inflacion de cada comunidad a lo largo del
# tiempo, no las diferencias de nivel de precios entre comunidades (no es una
# paridad de poder adquisitivo regional).
#
# Fuente: IPC_CCAA/IPC_CCAA_DataExtraction.R (media anual del IPC general por
# CCAA, tabla t=24078 del INE) -- ejecutarlo antes si no existe.
ipc <- read.xlsx("../IPC_CCAA/IPC_CCAA_Anual.xlsx", sheet = "datos") %>%
  mutate(columna = if_else(codigo_ccaa == "00", NACIONAL_COL, codigo_a_columna[as.character(as.integer(codigo_ccaa))])) %>%
  select(anio, columna, ipc) %>%
  pivot_wider(names_from = columna, values_from = ipc)
stopifnot("Faltan geografias en IPC_CCAA_Anual.xlsx" = all(geo_cols %in% names(ipc)))

anio_base <- function(df, desfase) max(df$fecha) - desfase

# df en formato ancho (fecha + columnas geograficas), en euros corrientes.
deflactar <- function(df, desfase) {
  i_t <- match(df$fecha - desfase, ipc$anio)
  i_base <- match(anio_base(df, desfase), ipc$anio)
  stopifnot("Falta el IPC de algun año necesario para deflactar" = !anyNA(c(i_t, i_base)))
  for (col in geo_cols) df[[col]] <- df[[col]] * ipc[[col]][i_base] / ipc[[col]][i_t]
  df
}

ccaa17_cols <- codigo_a_columna[as.character(1:17)]
RENTA <- c("Renta media y mediana - Renta media por unidad de consumo",
           "Renta media y mediana - Renta mediana por unidad de consumo")
stopifnot(all(RENTA %in% consolidado$atributo))
for (a in RENTA) {
  filas <- consolidado$atributo == a
  bloque <- consolidado[filas, ]
  base <- anio_base(bloque, desfase = 1)
  stopifnot("La ultima renta no es del mismo año en las 17 CCAA" =
    !anyNA(unlist(bloque[bloque$fecha == max(bloque$fecha), ccaa17_cols])))
  consolidado[filas, geo_cols] <- deflactar(bloque, desfase = 1)[, geo_cols]
  consolidado$unidad[filas] <- paste("Euros constantes de", base)
  consolidado$atributo[filas] <- paste0(a, " (euros constantes de ", base, ")")
}

# Compraventa: media anual (2004-2024, aunque la fuente solo empieza en
# 2007-T1) de las 4 medias trimestrales de precio €/m2 de vivienda
# (columna "viv-pm2", solo vivienda -- se descartan garajes/trasteros del
# mismo fichero). geo=="Nacional"/"Comunidad" ya vienen desagregados en el
# propio CSV de origen, no hace falta sumar provincias.
reg <- read.xlsx("../REGISTRADORES/REGISTRADORES_Residencial.xlsx", sheet = "datos")
names(reg) <- gsub("-", "_", names(reg))
reg_anual <- reg %>%
  filter(geo %in% c("Nacional", "Comunidad"), ano >= 2004, ano <= 2024) %>%
  group_by(ano, geo, cod_ca) %>%
  summarise(precio_m2 = mean(viv_pm2, na.rm = TRUE), .groups = "drop") %>%
  mutate(columna = if_else(geo == "Nacional", NACIONAL_COL, codigo_a_columna[as.character(cod_ca)])) %>%
  select(fecha = ano, columna, precio_m2) %>%
  pivot_wider(names_from = columna, values_from = precio_m2) %>%
  arrange(fecha) %>%
  completar_columnas_geo()
base_compra <- anio_base(reg_anual, desfase = 0)
stopifnot("El ultimo precio de compraventa no es del mismo año en las 17 CCAA" =
  !anyNA(unlist(reg_anual[reg_anual$fecha == base_compra, ccaa17_cols])))
reg_anual <- deflactar(reg_anual, desfase = 0)

fila_compra <- bind_cols(
  data.frame(atributo = paste0("Precio de compraventa de vivienda (euros constantes de ", base_compra, ")"),
             dimension = "Condiciones materiales de vida",
             unidad = paste("Euros constantes de", base_compra, "por metro cuadrado"),
             fuente = "REGISTRADORES", stringsAsFactors = FALSE)[rep(1, nrow(reg_anual)), ],
  reg_anual
)

# Alquiler: mediana anual de precio €/m2/mes, solo Vivienda Colectiva (VC,
# piso -- se descarta VU, vivienda unifamiliar/chalet, para tener un unico
# numero por CCAA/año comparable con el de compraventa, que tampoco separa
# tipologias).
#
# SERPAVI no publica fila de total nacional en la hoja CCAA (a diferencia de
# REGISTRADORES, que si la trae). Peticion explicita del usuario 2026-09-22:
# en vez de dejar Total.Nacional a NA, se rellena con la MEDIA SIMPLE de las
# CCAA, para que el indicador tenga una linea de referencia estatal con la que
# compararse igual que el resto.
#
# OJO al interpretarla: es la media de las 19 comunidades/ciudades autonomas
# dando a cada una el MISMO peso, no una media ponderada por poblacion ni un
# agregado nacional real -- La Rioja pesa aqui lo mismo que Madrid, asi que
# queda por debajo de lo que daria un agregado ponderado (las CCAA mas caras
# son tambien de las mas pobladas). Sirve como referencia de "comunidad
# tipica", no como "precio medio en España". El informe lo dice explicitamente
# donde usa esta serie.
serp <- read.xlsx("../SERPAVI/SERPAVI_CCAA.xlsx", sheet = "datos")
serp_anual <- serp %>%
  filter(tipo_vivienda == "VC", anio >= 2004, anio <= 2024) %>%
  mutate(columna = codigo_a_columna[as.character(as.integer(ccaa_codigo))]) %>%
  select(fecha = anio, columna, alquiler_m2_mediana) %>%
  pivot_wider(names_from = columna, values_from = alquiler_m2_mediana) %>%
  arrange(fecha) %>%
  completar_columnas_geo()
# Se deflacta ANTES de la media simple, para que Total.Nacional sea la media
# de las CCAA ya deflactadas, cada una con su propio IPC.
base_alquiler <- anio_base(serp_anual, desfase = 0)
stopifnot("El ultimo precio de alquiler no es del mismo año en las 17 CCAA" =
  !anyNA(unlist(serp_anual[serp_anual$fecha == base_alquiler, ccaa17_cols])))
serp_anual <- deflactar(serp_anual, desfase = 0)
serp_anual[[NACIONAL_COL]] <- rowMeans(serp_anual[, ccaa_cols], na.rm = TRUE)

fila_alquiler <- bind_cols(
  data.frame(atributo = paste0("Precio de alquiler de vivienda (euros constantes de ", base_alquiler, ")"),
             dimension = "Condiciones materiales de vida",
             unidad = paste("Euros constantes de", base_alquiler, "por metro cuadrado al mes"),
             fuente = "SERPAVI", stringsAsFactors = FALSE)[rep(1, nrow(serp_anual)), ],
  serp_anual
)

# --- PIB_CCAA: PIB per capita REAL (deflactado) por CCAA -------------------
#
# Peticion explicita del usuario 2026-09-23: incorporar el PIB per capita
# deflactado como atributo nuevo, dimension "Condiciones materiales de vida"
# (mismo bloque tematico que renta y vivienda -- es el nivel de riqueza
# economica de la comunidad, no un indicador de encuesta del ICV).
#
# Fuente: PIB_CCAA/PIB_CCAA_DataExtraction.R (ver cabecera de ese script para
# el detalle completo del metodo). Resumen: el INE no publica un "PIB per
# capita a precios constantes" directo por CCAA, asi que se construye
# encadenando 3 piezas oficiales del INE -- indice de volumen encadenado del
# PIB total (base 2020=100), PIB nominal del ultimo año, 2023 (para anclar
# el indice en euros) y poblacion por CCAA (Padron 2000-2020 + Estadistica Continua de
# Poblacion 2021-2023) -- sin aplicar ningun deflactor de precios casero
# (decision explicita del usuario, 2026-09-23: usar solo series oficiales del
# INE, no CPI). Se usa fuente = "PIB_CCAA" (no "INE") precisamente para dejar
# claro que es un indicador DERIVADO/construido a partir de varias tablas del
# INE, no un numero que el INE publique ya calculado tal cual.
#
# El codigo de CCAA de PIB_CCAA_Real.xlsx (columna codigo_ccaa, "00".."19")
# usa la misma codificacion oficial que REGISTRADORES/SERPAVI, con "00" para
# Total Nacional -- verificado en vivo al construir la fuente.
pib <- read.xlsx("../PIB_CCAA/PIB_CCAA_Real.xlsx", sheet = "datos")
anio_ancla_pib <- max(pib$anio)
pib_anual <- pib %>%
  mutate(columna = if_else(codigo_ccaa == "00", NACIONAL_COL, codigo_a_columna[as.character(as.integer(codigo_ccaa))])) %>%
  select(fecha = anio, columna, pib_real_pc_eur) %>%
  pivot_wider(names_from = columna, values_from = pib_real_pc_eur) %>%
  arrange(fecha) %>%
  completar_columnas_geo()

fila_pib <- bind_cols(
  data.frame(atributo = paste0("PIB per cápita real (precios constantes de ", anio_ancla_pib, ")"),
             dimension = "Condiciones materiales de vida",
             unidad = paste("Euros constantes de", anio_ancla_pib, "por habitante"), fuente = "PIB_CCAA", stringsAsFactors = FALSE)[rep(1, nrow(pib_anual)), ],
  pib_anual
)

consolidado <- bind_rows(consolidado, fila_compra, fila_alquiler, fila_pib) %>%
  select(atributo, dimension, fuente, unidad, fecha, all_of(geo_cols))

# Guarda de cordura: cada (atributo, fecha) debe salir de UNA sola tabla de
# origen (una CCAA por columna). Si esto falla es que dos ficheros distintos
# se estan apilando bajo el mismo atributo sin desambiguar, como paso con
# PM10/PM2,5 antes del arreglo de arriba.
stopifnot(nrow(count(consolidado, atributo, fecha) %>% filter(n > 1)) == 0)

stopifnot("ATRIBUTOS_EXCLUIDOS ya no coincide con ningun atributo del consolidado (revisar el texto exacto)" =
  all(ATRIBUTOS_EXCLUIDOS %in% consolidado$atributo))
consolidado <- consolidado %>% filter(!atributo %in% ATRIBUTOS_EXCLUIDOS)

wb <- createWorkbook()
addWorksheet(wb, "datos")
writeData(wb, "datos", consolidado)
setColWidths(wb, "datos", cols = 1:ncol(consolidado), widths = "auto")
freezePane(wb, "datos", firstRow = TRUE, firstCol = TRUE)
saveWorkbook(wb, "INE_ICV_CCAA_Consolidado.xlsx", overwrite = TRUE)

cat(nrow(consolidado), "filas,", length(unique(consolidado$atributo)), "atributos distintos, en INE_ICV_CCAA_Consolidado.xlsx\n")
cat("Excluidas:", paste(EXCLUIDOS, collapse = ", "), "\n")
