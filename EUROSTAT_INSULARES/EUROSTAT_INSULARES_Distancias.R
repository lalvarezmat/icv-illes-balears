# EUROSTAT_INSULARES_Distancias.R
#
# Distancia de cada region insular al CONTINENTE de su pais (peticion del
# usuario, 2026-11-14: "la distancia fisica en km del archipielago al pais de
# referencia, en el continente").
#
# Definicion:
#   - Territorio continental de un pais = la masa de tierra principal de cada
#     una de sus regiones NUTS 2 no insulares (de cada region se toma su
#     poligono mas grande). Asi quedan fuera las islas que pertenecen a
#     regiones continentales (Eubea, Elba, Salamina...): "en el continente".
#     España sin Ceuta y Melilla (Africa); Francia sin Corse ni ultramar.
#   - Cada region insular se representa por sus islas o grupos de islas
#     (unidades NUTS 3), cada una por su punto de etiqueta oficial de GISCO.
#   - Distancia de una unidad NUTS 3 = distancia en linea recta (gran
#     circulo) de su punto al vertice mas cercano del territorio continental
#     de su pais. Con la escala 1:3 millones los vertices de la costa estan a
#     1-3 km unos de otros: ese es el orden del error.
#   - Distancia de la region = media de las de sus unidades NUTS 3, ponderada
#     por su poblacion (a 1 de enero del ultimo año con dato).
#     TRAMPA (comprobada): en la geometria 1:3M, Eubea forma un solo poligono
#     con el continente (el estrecho de Calcis mide unos 40 m y tiene dos
#     puentes), asi que cuenta como continente. Afecta a las islas del Egeo,
#     cuyo punto mas cercano es a menudo la costa de Eubea: unas decenas de km.
#   - Malta y Chipre no tienen parte continental: se mide la distancia al
#     territorio continental de la Union Europea mas cercano.
#   - Ademas, la distancia a la capital del pais (mismo promedio ponderado),
#     como dato complementario.
#
# Fuentes: GISCO (Eurostat), NUTS 2021, escala 1:3 millones (regiones NUTS 2)
# y puntos de etiqueta NUTS 3:
#   https://gisco-services.ec.europa.eu/distribution/v2/nuts/geojson/
# Poblacion NUTS 3: Eurostat demo_r_pjanaggr3.
#
# Genera: EUROSTAT_INSULARES_Distancias.xlsx (hojas "regiones" y "nuts3")
# y guarda los dos GeoJSON originales en esta carpeta (reanudable).

Sys.setenv(R_LIBCURL_SSL_REVOKE_BEST_EFFORT = "TRUE")

library(httr)
library(jsonlite)
library(dplyr)
library(openxlsx)

CARPETA <- "."
GISCO <- "https://gisco-services.ec.europa.eu/distribution/v2/nuts/geojson/"
FN_REG <- file.path(CARPETA, "NUTS_RG_03M_2021_4326_LEVL_2.geojson")
FN_LB  <- file.path(CARPETA, "NUTS_LB_2021_4326_LEVL_3.geojson")

ISLAS_PAIS <- list(ES = c("ES53", "ES70"), PT = c("PT20", "PT30"),
                   EL = c("EL41", "EL42", "EL62", "EL43"), IT = c("ITG1", "ITG2"),
                   FR = c("FRM0", "FRY1", "FRY2", "FRY4", "FRY5"), FI = "FI20",
                   MT = "MT00", CY = "CY00")
FUERA <- c("ES63", "ES64", "FRY3")             # Ceuta, Melilla, Guyane: ni islas ni continente europeo
NO_CONTINENTAL_UE <- c("IE04", "IE05", "IE06", "DK01", "DK02")  # isla de Irlanda y de Selandia
UE <- c("AT", "BE", "BG", "CY", "CZ", "DE", "DK", "EE", "EL", "ES", "FI", "FR", "HR", "HU", "IE",
        "IT", "LT", "LU", "LV", "MT", "NL", "PL", "PT", "RO", "SE", "SI", "SK")
CAPITAL <- list(ES = c(40.4168, -3.7038), PT = c(38.7223, -9.1393), EL = c(37.9838, 23.7275),
                IT = c(41.9028, 12.4964), FR = c(48.8566, 2.3522), FI = c(60.1699, 24.9384))

descargar <- function(url, destino) {
  if (file.exists(destino)) return(invisible(TRUE))
  r <- httr::GET(url, httr::write_disk(destino, overwrite = TRUE), httr::timeout(300))
  stopifnot(httr::status_code(r) == 200)
}
descargar(paste0(GISCO, basename(FN_REG)), FN_REG)
descargar(paste0(GISCO, basename(FN_LB)), FN_LB)

# Distancia de gran circulo (km) de un punto a muchos
haversine <- function(lat1, lon1, lat2, lon2) {
  rad <- pi / 180
  a <- sin((lat2 - lat1) * rad / 2)^2 + cos(lat1 * rad) * cos(lat2 * rad) * sin((lon2 - lon1) * rad / 2)^2
  2 * 6371.0088 * asin(pmin(1, sqrt(a)))
}

# --- Territorio continental: poligono principal de cada region NUTS 2 ---------
reg <- fromJSON(FN_REG, simplifyVector = FALSE)$features
area_anillo <- function(m) {  # area aproximada (grados^2 corregidos por latitud), solo para comparar
  x <- m[, 1] * cos(mean(m[, 2]) * pi / 180); y <- m[, 2]
  abs(sum(x * c(y[-1], y[1]) - c(x[-1], x[1]) * y)) / 2
}
anillo <- function(r) do.call(rbind, lapply(r, function(p) c(p[[1]], p[[2]])))
principal <- lapply(reg, function(f) {
  g <- f$geometry
  polis <- if (g$type == "MultiPolygon") g$coordinates else list(g$coordinates)
  exteriores <- lapply(polis, function(p) anillo(p[[1]]))
  m <- exteriores[[which.max(vapply(exteriores, area_anillo, numeric(1)))]]
  data.frame(nuts2 = f$properties$NUTS_ID, pais = f$properties$CNTR_CODE, lon = m[, 1], lat = m[, 2])
})
principal <- bind_rows(principal)
islas_todas <- unlist(ISLAS_PAIS)
continental <- principal %>% filter(!nuts2 %in% c(islas_todas, FUERA))

# --- Islas: unidades NUTS 3 con su punto de etiqueta y su poblacion -------------
lb <- fromJSON(FN_LB, simplifyVector = FALSE)$features
lb <- bind_rows(lapply(lb, function(f) data.frame(
  nuts3 = f$properties$NUTS_ID, nombre = f$properties$NAME_LATN,
  lon = f$geometry$coordinates[[1]], lat = f$geometry$coordinates[[2]])))
nuts3 <- lb %>% filter(substr(nuts3, 1, 4) %in% islas_todas) %>%
  mutate(nuts2 = substr(nuts3, 1, 4))

q <- list(format = "JSON", lang = "EN", unit = "NR", sex = "T", age = "TOTAL", sinceTimePeriod = 2020)
url <- paste0("https://ec.europa.eu/eurostat/api/dissemination/statistics/1.0/data/demo_r_pjanaggr3?",
              paste0(names(q), "=", unlist(q), collapse = "&"),
              paste0("&geo=", nuts3$nuts3, collapse = ""))
js <- fromJSON(content(GET(url, timeout(300)), as = "text", encoding = "UTF-8"), simplifyVector = FALSE)
gi <- unlist(js$dimension$geo$category$index); ti <- unlist(js$dimension$time$category$index)
nt <- length(ti); anio_pob <- names(ti)[which.max(ti)]
pob <- vapply(names(gi), function(g) {
  # ultimo año con dato de esa unidad
  for (t in rev(names(sort(ti)))) {
    v <- js$value[[as.character(gi[[g]] * nt + ti[[t]])]]
    if (!is.null(v)) return(as.numeric(v))
  }
  NA_real_
}, numeric(1))
nuts3$poblacion <- unname(pob[nuts3$nuts3])
stopifnot(!anyNA(nuts3$poblacion))

pais_de <- function(n2) names(ISLAS_PAIS)[vapply(ISLAS_PAIS, function(v) n2 %in% v, logical(1))]
nuts3 <- nuts3 %>% rowwise() %>% mutate(pais = pais_de(nuts2)) %>% ungroup()

# Distancia de cada unidad NUTS 3 al continente de su pais (o de la UE)
dist <- lapply(seq_len(nrow(nuts3)), function(i) {
  u <- nuts3[i, ]
  objetivo <- if (u$pais %in% c("MT", "CY")) {
    continental %>% filter(pais %in% UE, !nuts2 %in% NO_CONTINENTAL_UE)
  } else {
    continental %>% filter(pais == u$pais)
  }
  d <- haversine(u$lat, u$lon, objetivo$lat, objetivo$lon)
  k <- which.min(d)
  cap <- CAPITAL[[u$pais]]
  data.frame(dist_continente_km = d[k], region_continental_mas_cercana = objetivo$nuts2[k],
             dist_capital_km = if (is.null(cap)) NA_real_ else haversine(u$lat, u$lon, cap[1], cap[2]))
})
nuts3 <- bind_cols(nuts3, bind_rows(dist))

regiones <- nuts3 %>%
  group_by(nuts2, pais) %>%
  # summarise() evalua en orden: las columnas que se sobrescriben van al final
  summarise(n_unidades_nuts3 = n(),
            dist_continente_min_km = min(dist_continente_km),
            dist_capital_km = if (all(is.na(dist_capital_km))) NA_real_ else weighted.mean(dist_capital_km, poblacion),
            dist_continente_km = weighted.mean(dist_continente_km, poblacion),
            poblacion = sum(poblacion),
            .groups = "drop") %>%
  mutate(referencia = ifelse(pais %in% c("MT", "CY"), "Continente de la UE", "Continente de su país")) %>%
  arrange(dist_continente_km)

print(as.data.frame(regiones %>% mutate(across(where(is.numeric), ~ round(.x)))))

wb <- createWorkbook()
addWorksheet(wb, "regiones"); writeData(wb, "regiones", regiones)
addWorksheet(wb, "nuts3"); writeData(wb, "nuts3", nuts3)
addWorksheet(wb, "metadatos")
writeData(wb, "metadatos", data.frame(
  campo = c("geometrias", "puntos", "poblacion", "anio_poblacion", "definicion"),
  valor = c(basename(FN_REG), basename(FN_LB), "Eurostat demo_r_pjanaggr3 (1 de enero)", anio_pob,
            "Media ponderada por poblacion de la distancia de cada unidad NUTS 3 al vertice mas cercano del poligono principal de las regiones NUTS 2 continentales de su pais")))
saveWorkbook(wb, file.path(CARPETA, "EUROSTAT_INSULARES_Distancias.xlsx"), overwrite = TRUE)
