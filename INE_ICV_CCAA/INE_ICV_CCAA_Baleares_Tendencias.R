# INE_ICV_CCAA_Baleares_Tendencias.R
#
# Mismo objetivo que INE_ICV_CCAA_Baleares_Singularidades.R (detectar en que
# atributos Balears, Illes es un valor atipico frente a las otras 16
# comunidades autonomas) pero mirando la VELOCIDAD DE CAMBIO en el tiempo en
# vez del nivel actual: para cada atributo con al menos 3 años de dato,
# calcula la pendiente anual por comunidad y compara la de Balears contra la
# media/sd de las otras 16 (rank + z-score, mismo patron que el script de
# niveles). Se guarda tambien el cambio total (ultimo valor menos primero)
# como metrica complementaria mas facil de leer que la pendiente en si.
#
# Estimador de pendiente ROBUSTO (Theil-Sen), no lm() (cambiado el
# 2026-09-21, peticion explicita del usuario): lm(valor ~ año) es una media
# de minimos cuadrados, y un solo año con una fluctuacion interanual fuerte
# (outlier) puede desviar bastante la pendiente estimada -- justo el mismo
# problema que motivo usar mediana en vez de un año suelto en
# Singularidades.R. El estimador de Theil-Sen (Sen, 1968) es la MEDIANA de
# las pendientes entre TODOS los pares de puntos (y_j-y_i)/(x_j-x_i), i<j:
# un unico año atipico solo contamina las parejas que lo incluyen, no las
# demas, así que la mediana final apenas se mueve (lm en cambio lo nota en
# la pendiente entera, un unico punto de mucho leverage puede dominarla).
# Implementado a mano con combn() en vez de un paquete (`trend::sens.slope`
# exige una serie regularmente espaciada; aqui hace falta la version
# general de Sen, que admite huecos en los años, dividiendo por la
# diferencia real de año entre cada pareja).
#
# Por que hace falta ademas del analisis de nivel: un atributo puede tener un
# nivel actual nada llamativo pero estar mejorando o empeorando mucho mas
# rapido que el resto de España (o al reves, un buen nivel actual que se esta
# estancando mientras el resto mejora) -- eso solo se ve mirando la
# trayectoria completa, no la ultima foto.
#
# Salida: INE_ICV_CCAA_Baleares_Tendencias.xlsx, ordenado igual que el de
# niveles (cercania al extremo del ranking, luego |z-score|).
#
# Reanudable/regenerable: no llama a ninguna red, solo relee
# INE_ICV_CCAA_Consolidado.xlsx.

suppressMessages({
  library(openxlsx)
  library(dplyr)
})

CATALOGO <- "INE_ICV_CCAA_Consolidado.xlsx"
d <- read.xlsx(CATALOGO, sheet = "datos", detectDates = TRUE)

ccaa17 <- c("01.Andalucía","02.Aragón","03.Asturias,.Principado.de","04.Balears,.Illes",
            "05.Canarias","06.Cantabria","07.Castilla.y.León","08.Castilla.-.La.Mancha",
            "09.Cataluña","10.Comunitat.Valenciana","11.Extremadura","12.Galicia",
            "13.Madrid,.Comunidad.de","14.Murcia,.Región.de","15.Navarra,.Comunidad.Foral.de",
            "16.País.Vasco","17.Rioja,.La")
bal <- "04.Balears,.Illes"

d$anio <- d$fecha # "fecha" ya es el año como entero desde el 2026-09-21 (INE_ICV_CCAA_Consolidar.R), no un Date

atributos <- unique(d$atributo)

pendiente <- function(anio, valor) {
  ok <- !is.na(valor)
  anio <- anio[ok]; valor <- valor[ok]
  n <- sum(ok)
  if (n < 3) return(c(slope = NA, n = n))
  pares <- combn(n, 2)
  dx <- anio[pares[2, ]] - anio[pares[1, ]]
  dy <- valor[pares[2, ]] - valor[pares[1, ]]
  validos <- dx != 0
  c(slope = median(dy[validos] / dx[validos]), n = n)
}

fila <- function(atrib) {
  sub <- d %>% filter(atributo == atrib)
  unidad <- sub$unidad[1]

  slopes <- sapply(ccaa17, function(reg) pendiente(sub$anio, sub[[reg]])["slope"])
  ns <- sapply(ccaa17, function(reg) pendiente(sub$anio, sub[[reg]])["n"])
  names(slopes) <- ccaa17
  names(ns) <- ccaa17 # imprescindible: sin esto ns queda nombrado "<reg>.n" y no cruza con slopes (bug ya corregido)

  # cambio total (ultimo valor disponible - primero disponible) como metrica
  # complementaria, mas facil de interpretar que la pendiente
  cambio_total <- sapply(ccaa17, function(reg) {
    v <- sub[[reg]]
    f <- sub$anio
    ok <- !is.na(v)
    if (sum(ok) < 2) return(NA)
    fo <- f[ok]; vo <- v[ok]
    ord <- order(fo)
    vo[ord][length(vo)] - vo[ord][1]
  })
  names(cambio_total) <- ccaa17

  slopes_ok <- slopes[!is.na(slopes) & ns[names(slopes)] >= 3]
  if (length(slopes_ok) < 6 || !(bal %in% names(slopes_ok))) return(NULL)

  slope_bal <- slopes_ok[[bal]]
  otras <- slopes_ok[setdiff(names(slopes_ok), bal)]
  media_otras <- mean(otras)
  sd_otras <- sd(otras)
  z <- if (sd_otras > 0) (slope_bal - media_otras) / sd_otras else NA
  rank_desc <- rank(-slopes_ok, ties.method = "min")[[bal]]
  n_ccaa <- length(slopes_ok)

  data.frame(
    atributo = atrib,
    unidad = unidad,
    n_anios_baleares = ns[[bal]],
    pendiente_baleares_anual = round(slope_bal, 3),
    media_pendiente_resto = round(media_otras, 3),
    cambio_total_baleares = round(cambio_total[[bal]], 2),
    media_cambio_total_resto = round(mean(cambio_total[setdiff(names(cambio_total), bal)], na.rm = TRUE), 2),
    rank_baleares = rank_desc,
    n_ccaa = n_ccaa,
    z_score = round(z, 2),
    stringsAsFactors = FALSE
  )
}

resultado <- bind_rows(lapply(atributos, fila)) %>%
  mutate(extremo = pmin(rank_baleares, n_ccaa - rank_baleares + 1)) %>%
  arrange(extremo, desc(abs(z_score)))

wb <- createWorkbook()
addWorksheet(wb, "tendencia_baleares")
writeData(wb, "tendencia_baleares", resultado)
setColWidths(wb, "tendencia_baleares", cols = 1:ncol(resultado), widths = "auto")
freezePane(wb, "tendencia_baleares", firstRow = TRUE)
saveWorkbook(wb, "INE_ICV_CCAA_Baleares_Tendencias.xlsx", overwrite = TRUE)

cat(nrow(resultado), "atributos con serie temporal suficiente (>=3 años, >=6 CCAA con dato)\n\n")
cat("=== Baleares en 1o/2o o penultimo/ultimo puesto por VELOCIDAD DE CAMBIO entre 17 CCAA ===\n")
extremos <- resultado %>% filter(extremo <= 2)
print(extremos[, c("atributo","unidad","pendiente_baleares_anual","media_pendiente_resto","rank_baleares","n_ccaa","z_score","cambio_total_baleares","media_cambio_total_resto")], row.names = FALSE)
