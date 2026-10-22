# INE_ICV_CCAA_Baleares_Singularidades.R
#
# Compara, para cada uno de los 79 atributos de INE_ICV_CCAA_Consolidado.xlsx,
# Balears, Illes contra las otras 16 comunidades autonomas (se excluyen
# Ceuta/Melilla por ser ciudades autonomas, no comunidades; Total Nacional
# queda solo como referencia informativa, no entra en el ranking ni en la
# media). Calcula el puesto de Balears (1=mas alto, 17=mas bajo) y su
# z-score frente a la media/sd de las otras 16, para detectar en que
# atributos Balears es un valor atipico respecto al resto de España EN NIVEL
# (ver INE_ICV_CCAA_Baleares_Tendencias.R para la misma pregunta pero sobre
# la VELOCIDAD de cambio en el tiempo).
#
# MEDIANA de los ultimos 5 años disponibles, no el ultimo año a solas
# (cambiado el 2026-09-21, peticion explicita del usuario): un solo año
# puede tener una fluctuacion interanual fuerte (una encuesta con muestra
# pequeña, un año atipico...) que desvirtue el ranking/z-score si se usa tal
# cual. La MEDIANA (no la media) de los ultimos 5 años de cada comunidad es
# mas robusta que un año suelto y menos sensible a un año atipico dentro de
# la ventana que la media aritmetica. Los 5 años son los ultimos CON DATO
# para ese atributo en concreto (no un rango de calendario fijo), igual
# criterio que ya usaba INE_ICV_CCAA_Consolidar.R para "los ultimos años
# disponibles". Si un atributo tiene menos de 5 años de historia se usan los
# que haya (la funcion `mediana()` no exige un minimo de años, solo que cada
# comunidad tenga al menos un dato no-NA en la ventana).
#
# Salida: INE_ICV_CCAA_Baleares_Singularidades.xlsx, ordenado por cercania al
# extremo del ranking (puesto 1/2 o 16/17 primero) y dentro de eso por
# |z-score| descendente.
#
# Reanudable/regenerable: no llama a ninguna red, solo relee
# INE_ICV_CCAA_Consolidado.xlsx (regenerarlo antes con
# INE_ICV_CCAA_Consolidar.R si el catalogo cambio).

suppressMessages({
  library(openxlsx)
  library(dplyr)
})

VENTANA_ANIOS <- 5

CATALOGO <- "INE_ICV_CCAA_Consolidado.xlsx"
d <- read.xlsx(CATALOGO, sheet = "datos", detectDates = TRUE)

ccaa17 <- c("01.Andalucía","02.Aragón","03.Asturias,.Principado.de","04.Balears,.Illes",
            "05.Canarias","06.Cantabria","07.Castilla.y.León","08.Castilla.-.La.Mancha",
            "09.Cataluña","10.Comunitat.Valenciana","11.Extremadura","12.Galicia",
            "13.Madrid,.Comunidad.de","14.Murcia,.Región.de","15.Navarra,.Comunidad.Foral.de",
            "16.País.Vasco","17.Rioja,.La")
bal <- "04.Balears,.Illes"

# los nombres de columna reales llevan "." donde el original tenia espacio
# (efecto de leer con check.names=TRUE por defecto al construir el
# consolidado) -- se comprueba que existen antes de seguir
faltan <- setdiff(c(bal, ccaa17, "Total.Nacional"), names(d))
if (length(faltan) > 0) stop("Columnas geograficas no encontradas: ", paste(faltan, collapse = " | "))

atributos <- unique(d$atributo)

fila <- function(atrib) {
  sub <- d %>% filter(atributo == atrib)
  sub_val <- sub %>% filter(!is.na(.data[[bal]]))
  if (nrow(sub_val) == 0) return(NULL)
  ventana <- sub_val %>% filter(fecha > max(fecha) - VENTANA_ANIOS)

  medianas <- sapply(intersect(ccaa17, names(ventana)), function(reg) median(ventana[[reg]], na.rm = TRUE))
  medianas <- medianas[!is.na(medianas) & !is.nan(medianas)]

  if (length(medianas) < 5 || !(bal %in% names(medianas))) return(NULL)

  val_bal <- medianas[[bal]]
  otras <- medianas[setdiff(names(medianas), bal)]
  media_otras <- mean(otras)
  sd_otras <- sd(otras)
  z <- if (sd_otras > 0) (val_bal - media_otras) / sd_otras else NA
  rank_desc <- rank(-medianas, ties.method = "min")[[bal]] # 1 = valor mas alto
  n <- length(medianas)

  data.frame(
    atributo = atrib,
    unidad = sub$unidad[1],
    periodo = paste0(min(ventana$fecha), "-", max(ventana$fecha)),
    n_anios = n_distinct(ventana$fecha),
    valor_baleares = round(val_bal, 2),
    media_resto_ccaa = round(media_otras, 2),
    min_resto = round(min(otras), 2),
    max_resto = round(max(otras), 2),
    rank_baleares = rank_desc,
    n_ccaa = n,
    z_score = round(z, 2),
    total_nacional = round(median(ventana[["Total.Nacional"]], na.rm = TRUE), 2),
    stringsAsFactors = FALSE
  )
}

resultado <- bind_rows(lapply(atributos, fila)) %>%
  mutate(extremo = pmin(rank_baleares, n_ccaa - rank_baleares + 1)) %>%
  arrange(extremo, desc(abs(z_score)))

wb <- createWorkbook()
addWorksheet(wb, "baleares")
writeData(wb, "baleares", resultado)
setColWidths(wb, "baleares", cols = 1:ncol(resultado), widths = "auto")
freezePane(wb, "baleares", firstRow = TRUE)
saveWorkbook(wb, "INE_ICV_CCAA_Baleares_Singularidades.xlsx", overwrite = TRUE)

cat(nrow(resultado), "atributos analizados (mediana de los ultimos", VENTANA_ANIOS, "años disponibles)\n\n")
cat("=== Baleares en 1er/2o puesto (mas alto) o penultimo/ultimo (mas bajo) entre 17 CCAA ===\n")
extremos <- resultado %>% filter(extremo <= 2)
print(extremos[, c("atributo","unidad","periodo","valor_baleares","rank_baleares","n_ccaa","media_resto_ccaa","z_score")], row.names = FALSE)
