# INE_ICV_CCAA_NombresCortos.R
#
# Añade a la hoja "datos" de INE_ICV_CCAA_Consolidado.xlsx una columna
# "nombre_corto" (justo despues de "atributo") con un nombre corto
# identificativo pensado para usarse como etiqueta en graficos (leyendas,
# ejes) donde el nombre largo completo de "atributo" no cabe.
#
# El nombre corto es una tabla de traduccion fija (NOMBRES_CORTOS mas abajo),
# no una regla automatica de abreviado: es una decision editorial (que
# palabras quitar, que quedarse) que no se puede derivar mecanicamente del
# nombre largo sin perder criterio.
#
# El orden de NOMBRES_CORTOS es el orden de aparicion de unique(atributo) en
# la hoja "datos" en el momento de escribir este script (2026-09-21, 79
# atributos -- 85 menos 6 hojas de "poblacion/estaciones consideradas" de las
# tablas de contaminacion PM10/PM2,5, excluidas en INE_ICV_CCAA_Consolidar.R
# por ser tamaño de muestra, no una medida en si; ver ATRIBUTOS_EXCLUIDOS en
# ese script -- "atributo" ya con el nombre COMPLETO desde que
# INE_ICV_CCAA_Consolidar.R usa la hoja "categorias" de cada _METADATOS.xlsx
# en vez del nombre de hoja truncado a 31 caracteres). Si se añaden/quitan
# tablas de origen y cambia el conjunto de atributos, este script para con
# un error explicito en vez de dejar nombres cortos huerfanos o atributos
# sin nombre corto -- hay que actualizar NOMBRES_CORTOS a mano cuando eso
# pase, no se puede automatizar sin perder el criterio editorial de las
# abreviaturas.
#
# Reanudable/regenerable: no llama a ninguna red, solo relee y reescribe
# INE_ICV_CCAA_Consolidado.xlsx (debe existir ya, generado por
# INE_ICV_CCAA_Consolidar.R -- ejecutar este script SIEMPRE DESPUES de
# aquel, porque Consolidar.R reescribe el fichero entero desde cero y se
# llevaria por delante la columna "nombre_corto" si se ejecuta despues).

library(openxlsx)
library(dplyr)

FICHERO <- "INE_ICV_CCAA_Consolidado.xlsx"

NOMBRES_CORTOS <- c(
  "Abandono temprano de la educación-formación en la población de 18 a 24 años" = "Abandono educativo 18-24",
  "Carencia material - Carencia material (carencia en al menos 3 conceptos)" = "Carencia material",
  "Carencia material - Carencia material severa (carencia en al menos 4 conceptos)" = "Carencia material severa",
  "Dificultades para llegar a fin de mes - Con cierta dificultad" = "Fin mes: cierta dificultad",
  "Dificultades para llegar a fin de mes - Con dificultad" = "Fin mes: dificultad",
  "Dificultades para llegar a fin de mes - Con mucha dificultad" = "Fin mes: mucha dificultad",
  "Incapacidad de hacer frente a gastos económicos imprevistos" = "Gastos imprevistos",
  "Morbilidad crónica. Personas con enfermedades o problemas de salud de larga duración" = "Morbilidad crónica",
  "Necesidades no satisfechas de cuidados médicos. - Otras razones" = "Nec. médicas: otras razones",
  "Necesidades no satisfechas de cuidados médicos. - Razones" = "Nec. médicas: razones",
  "Nivel de formación alcanzado por la población total - Nivel (0-2): preescolar, primaria y secundaria de 1ª etapa" = "Form. total: nivel 0-2",
  "Nivel de formación alcanzado por la población total - Nivel (3-4): secundaria de 2ª etapa y postsecundaria no superior" = "Form. total: nivel 3-4",
  "Nivel de formación alcanzado por la población total - Nivel (3-8): Educación secudaria o superior" = "Form. total: nivel 3-8",
  "Nivel de formación alcanzado por la población total - Nivel (5-8): primer y segundo ciclo de educación superior y doctorado" = "Form. total: nivel 5-8",
  "Nivel de formación alcanzado por la población adulta - De 25 a 64 años - Nivel (0-2): preescolar, primaria y secundaria de 1ª etapa" = "Form. 25-64: nivel 0-2",
  "Nivel de formación alcanzado por la población adulta - De 65 y más años - Nivel (0-2): preescolar, primaria y secundaria de 1ª etapa" = "Form. 65+: nivel 0-2",
  "Nivel de formación alcanzado por la población adulta - De 25 a 64 años - Nivel (3-4): secundaria de 2ª etapa y postsecundaria no superior" = "Form. 25-64: nivel 3-4",
  "Nivel de formación alcanzado por la población adulta - De 65 y más años - Nivel (3-4): secundaria de 2ª etapa y postsecundaria no superior" = "Form. 65+: nivel 3-4",
  "Nivel de formación alcanzado por la población adulta - De 25 a 64 años - Nivel (3-8): Educación secudaria o superior" = "Form. 25-64: nivel 3-8",
  "Nivel de formación alcanzado por la población adulta - De 65 y más años - Nivel (3-8): Educación secudaria o superior" = "Form. 65+: nivel 3-8",
  "Nivel de formación alcanzado por la población adulta - De 25 a 64 años - Nivel (5-8): primer y segundo ciclo de educación superior y doctorado" = "Form. 25-64: nivel 5-8",
  "Nivel de formación alcanzado por la población adulta - De 65 y más años - Nivel (5-8): primer y segundo ciclo de educación superior y doctorado" = "Form. 65+: nivel 5-8",
  "Nivel de formación alcanzado por la población joven (de 18 a 24 años) - Nivel (0-2): preescolar, primaria y secundaria de 1ª etapa" = "Form. 18-24: nivel 0-2",
  "Nivel de formación alcanzado por la población joven (de 18 a 24 años) - Nivel (3-4): secundaria de 2ª etapa y postsecundaria no superior" = "Form. 18-24: nivel 3-4",
  "Nivel de formación alcanzado por la población joven (de 18 a 24 años) - Nivel (3-8): Educación secudaria o superior" = "Form. 18-24: nivel 3-8",
  "Nivel de formación alcanzado por la población joven (de 18 a 24 años) - Nivel (5-8): primer y segundo ciclo de educación superior y doctorado" = "Form. 18-24: nivel 5-8",
  "Personas con limitaciones en la actividad diaria en los últimos 6 meses - Gravemente limitado" = "Limit. actividad: grave",
  "Personas con limitaciones en la actividad diaria en los últimos 6 meses - Limitado pero no gravemente" = "Limit. actividad: leve",
  "Personas con limitaciones en la actividad diaria en los últimos 6 meses - Ninguna limitación" = "Limit. actividad: ninguna",
  "Personas de 25 a 64 años que han recibido formación durante las últimas 4 semanas" = "Formación continua 25-64",
  "Población con falta de espacio en la vivienda" = "Falta espacio vivienda",
  "Población urbana expuesta a contaminación del aire (micropartículas PM10, PM2,5) - Media ponderada con la población de la concentración media anual de PM10 (µg/m3)" = "PM10 anual",
  "Población urbana expuesta a contaminación del aire (micropartículas PM10, PM2,5) - Media ponderada con la población de la concentración media anual de PM2.5 (µg/m3)" = "PM2,5 anual",
  "Población urbana expuesta a contaminación del aire (micropartículas PM10, PM2,5) - Media ponderada con población del nº días/ año que se supera concentracion límite diaria  de PM10" = "PM10 días superación",
  "Retrasos en los pagos" = "Retrasos en pagos",
  "Salud autopercibida - Bueno" = "Salud: buena",
  "Salud autopercibida - Malo" = "Salud: mala",
  "Salud autopercibida - Muy bueno" = "Salud: muy buena",
  "Salud autopercibida - Muy malo" = "Salud: muy mala",
  "Salud autopercibida - Regular" = "Salud: regular",
  "Empleo involuntario a tiempo parcial (porcentaje sobre el empleo total a tiempo parcial)" = "Empleo parcial involuntario",
  "Jornadas largas - De 40 a 48 horas" = "Jornada larga 40-48h",
  "Jornadas largas - De 49 y más horas" = "Jornada larga 49h+",
  "Tasa de empleo" = "Tasa de empleo",
  "Tasa de paro de larga duración (porcentaje sobre la población activa)" = "Paro de larga duración",
  "Tasa de paro" = "Tasa de paro",
  "Trabajo temporal" = "Trabajo temporal",
  "Desigualdad (S80/S20)" = "Desigualdad S80/S20",
  "Población con gasto elevado en vivienda" = "Gasto elevado vivienda",
  "Población en riesgo de pobreza con umbral fijo en el año 2008" = "Pobreza umbral fijo 2008",
  "Población en riesgo de pobreza relativa según distintos umbrales" = "Pobreza relativa",
  "Renta media y mediana - Renta media por unidad de consumo (euros constantes de 2023)" = "Renta media real",
  "Renta media y mediana - Renta mediana por unidad de consumo (euros constantes de 2023)" = "Renta mediana real",
  "Homicidios y Criminalidad - Tasa de criminalidad" = "Tasa de criminalidad",
  "Homicidios y Criminalidad - Tasa de homicidios" = "Tasa de homicidios",
  "Confianza en los demás" = "Confianza en los demás",
  "Satisfacción global con la vida" = "Satisfacción vital",
  "Esperanza de vida a diferentes edades" = "Esperanza de vida",
  "Delincuencia o vandalismo en la zona" = "Delincuencia/vandalismo",
  "Población que sufre problemas de contaminación y otros problemas ambientales" = "Contaminación ambiental",
  "Población que sufre problemas de ruidos producidos por vecinos o del exterior" = "Ruidos vecinos/exterior",
  "Población que vive en hogares con determinadas deficiencias en la vivienda" = "Deficiencias en vivienda",
  "Salarios bajos" = "Salarios bajos",
  "Fumadores diarios" = "Fumadores diarios",
  "Índice de masa corporal - Normopeso (18,5 kg/m2 <= IMC < 25 kg/m2)" = "IMC: normopeso",
  "Índice de masa corporal - Obesidad (IMC >= 30 kg/m2)" = "IMC: obesidad",
  "Índice de masa corporal - Peso insuficiente (IMC < 18,5 kg/m2)" = "IMC: bajo peso",
  "Índice de masa corporal - Sobrepeso (25 kg/m2 <= IMC < 30 kg/m2)" = "IMC: sobrepeso",
  "Ejercicio físico regular y sedentarismo en el tiempo libre - Ejercicio fisico regular" = "Ejercicio físico regular",
  "Ejercicio físico regular y sedentarismo en el tiempo libre - Sedentarismo" = "Sedentarismo",
  "Satisfacción con el trabajo - De 0 a 4" = "Satisf. trabajo 0-4",
  "Satisfacción con el trabajo - De 5 a 6" = "Satisf. trabajo 5-6",
  "Satisfacción con el trabajo - De 7 a 8" = "Satisf. trabajo 7-8",
  "Satisfacción con el trabajo - De 9 a 10" = "Satisf. trabajo 9-10",
  "Satisfacción con el trabajo - Satisfacción media" = "Satisf. trabajo (media)",
  "Satisfacción con la vivienda - Algo insatisfecho" = "Satisf. vivienda: algo insat.",
  "Satisfacción con la vivienda - Muy insatisfecho" = "Satisf. vivienda: muy insat.",
  "Satisfacción con la vivienda - Muy satisfecho" = "Satisf. vivienda: muy satisf.",
  "Satisfacción con la vivienda - Satisfecho" = "Satisf. vivienda: satisfecho",
  "Precio de compraventa de vivienda (euros constantes de 2024)" = "Precio compraventa (€/m²)",
  "Precio de alquiler de vivienda (euros constantes de 2024)" = "Precio alquiler (€/m²/mes)",
  "PIB per cápita real (precios constantes de 2023)" = "PIB per cápita real"
)

datos <- read.xlsx(FICHERO, sheet = "datos", detectDates = TRUE)
atributos_actuales <- unique(datos$atributo)

faltan <- setdiff(atributos_actuales, names(NOMBRES_CORTOS))
sobran <- setdiff(names(NOMBRES_CORTOS), atributos_actuales)
if (length(faltan) > 0) {
  stop("Atributos en '", FICHERO, "' sin nombre corto asignado (actualizar NOMBRES_CORTOS): ",
       paste(faltan, collapse = " | "))
}
if (length(sobran) > 0) {
  stop("NOMBRES_CORTOS tiene entradas que ya no existen en '", FICHERO, "' (quitar del vector): ",
       paste(sobran, collapse = " | "))
}
stopifnot("los nombres cortos deben ser unicos" = !any(duplicated(NOMBRES_CORTOS)))

datos <- datos %>%
  mutate(nombre_corto = NOMBRES_CORTOS[atributo], .after = atributo)

wb <- createWorkbook()
addWorksheet(wb, "datos")
writeData(wb, "datos", datos)
setColWidths(wb, "datos", cols = 1:ncol(datos), widths = "auto")
freezePane(wb, "datos", firstRow = TRUE, firstCol = TRUE)
saveWorkbook(wb, FICHERO, overwrite = TRUE)

cat("Columna 'nombre_corto' añadida para", length(atributos_actuales), "atributos en", FICHERO, "\n")
