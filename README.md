# La calidad de vida en Illes Balears: una visión multidimensional

¿En qué se diferencia Illes Balears del resto de España en calidad de vida, y
hacia dónde se mueve? Este repositorio contiene el estudio completo, con su
código y sus datos, para que cualquiera pueda reproducirlo.

**Leer el informe:** <https://lalvarezmat.github.io/icv-illes-balears/>
(versión compilada de `INE_ICV_CCAA/INE_ICV_CCAA_all.Rmd`, en `docs/index.html`).

El estudio analiza **82 indicadores de calidad de vida** de las 17 comunidades
autónomas entre **2004 y 2024**: 79 de los Indicadores de Calidad de Vida (ICV)
del INE, el precio de compraventa y de alquiler de la vivienda y el PIB per
cápita real. Los compara con varias técnicas (nivel con medianas de cinco años,
tendencias de Theil-Sen, componentes principales, k-means y análisis de rutas),
explicadas sin fórmulas para un lector no especialista.

**Autores**

- **Dr Luis Álvarez León**. Doctor en Matemática Aplicada (UCM). Catedrático ULPGC.
- **Dr Bartomeu Coll Vicens**. Doctor en Ciencias Matemáticas. Catedrático UIB. Área Ciencias Matemáticas e Informática. Matemática Aplicada.
- **Jose Vicens Colom**. Economista. Máster en Dirección Financiera. Director del Cercle d’Economia de Mallorca.
- **Dr. José Mª Vicens Gómez**. Doctor en Medicina. Especialista Medicina Interna. Experto en Estadística de la Salud y Administración Sanitaria. Ex-presidente del Cercle d’Economia de Mallorca (2018-2026). Académico RAMIB.

## Estudio complementario: ¿qué distingue a los archipiélagos?

**Leerlo:** <https://lalvarezmat.github.io/icv-illes-balears/insulares.html>
(versión compilada de `INE_ICV_CCAA/INE_ICV_CCAA_Insulares.Rmd`, en `docs/insulares.html`).

Compara **Illes Balears y Canarias con las 15 comunidades peninsulares** para
ver qué rasgos comparten los dos archipiélagos y cuáles no. Usa los 82
indicadores del informe y **15 más**: el crecimiento de la población, las
viviendas iniciadas, las compras de vivienda por no residentes, las viviendas
turísticas y las noches vendidas en plataformas (por habitante), y el peso de
cada una de las 10 ramas de actividad en el valor añadido bruto. Cada
indicador se clasifica según la distancia de cada isla a la media peninsular,
y el resultado se compara con el de las 136 parejas posibles de comunidades.

## Estudio europeo: ¿qué distingue a los archipiélagos de Europa?

**Leerlo:** <https://lalvarezmat.github.io/icv-illes-balears/europa.html>
(versión compilada de `EUROSTAT_INSULARES/EUROSTAT_INSULARES.Rmd`, en `docs/europa.html`).

Compara **10 archipiélagos y 8 islas grandes de la Unión Europea** (regiones
NUTS 2) **con la parte continental de su propio país**, con 53 indicadores de
Eurostat convertidos en relativos (por habitante, en porcentaje o sobre el
PIB) y la **distancia de cada archipiélago al continente**. La referencia de
cada isla es su país sin sus islas, calculada restando valores absolutos, para
que los indicadores sean comparables entre países. Incluye el PIB real por
habitante, versiones por población presente (residentes más turistas) de los
servicios y sucesos, y una prueba de qué rasgos se repiten en todos los
archipiélagos.

## Contenido

| Carpeta | Qué hay |
|---|---|
| `INE_ICV_CCAA/` | **El informe** (`INE_ICV_CCAA_all.Rmd`), los scripts que construyen el conjunto de datos del estudio y sus resultados: `INE_ICV_CCAA_Consolidado.xlsx` (los 82 indicadores por comunidad y año), el nivel y la tendencia de Illes Balears frente a las demás comunidades, y las 42 tablas del INE ya procesadas. También **el estudio de los archipiélagos** (`INE_ICV_CCAA_Insulares.Rmd`) y sus 15 indicadores nuevos (`INE_ICV_CCAA_Insulares_Consolidar.R` → `INE_ICV_CCAA_Insulares_Indicadores.xlsx`). |
| `INE_ICV/` | Los ficheros `.px` originales de los Indicadores de Calidad de Vida del INE (las 480 tablas del catálogo, sin modificar), su catálogo y los scripts que los descargan y procesan. |
| `REGISTRADORES/` | Precio de compraventa de vivienda (Colegio de Registradores, datos abiertos) y porcentaje de compras por extranjeros (anuarios de la Estadística Registral Inmobiliaria). |
| `SERPAVI/` | Precio de alquiler de vivienda (Ministerio de Vivienda y Agenda Urbana), ya pivotado por comunidad y año. |
| `PIB_CCAA/` | PIB per cápita real por comunidad: las dos tablas `.px` originales del INE y el resultado. Además, el peso de cada rama de actividad en el VAB (`PIB_CCAA_VAB_Ramas.R`), para el estudio de los archipiélagos. |
| `IPC_CCAA/` | IPC por comunidad (INE), usado para descontar la inflación de la renta y de los precios. |
| `POBLACION_CCAA/` | Población residente a 1 de enero por comunidad, 1971-2025, y su crecimiento (INE, Estadística Continua de Población). Estudio de los archipiélagos. |
| `MIVAU/` | Viviendas iniciadas (estadística del Ministerio de Vivienda y Agenda Urbana, en la serie que republica el ISTAC) y compraventas de vivienda según la residencia del comprador (Ministerio de Vivienda y Agenda Urbana), con los ficheros originales. Estudio de los archipiélagos. |
| `VIVIENDA_TURISTICA/` | Viviendas turísticas (INE, estadística experimental) y noches en alojamientos reservados por plataformas en línea (Eurostat). Estudio de los archipiélagos. |
| `EUROSTAT_INSULARES/` | **El estudio europeo** (`EUROSTAT_INSULARES.Rmd`) y sus scripts: la descarga de 69 series de Eurostat por región y por país (`EUROSTAT_INSULARES_Series.xlsx`), las distancias al continente (`EUROSTAT_INSULARES_Distancias.xlsx`) y los indicadores relativos con sus referencias continentales y comparaciones (`EUROSTAT_INSULARES_Indicadores.xlsx`). |

## Cómo reproducirlo

Todo está escrito en **R** (probado con R 4.4.2). Cada script se ejecuta
**desde su propia carpeta** (con `setwd()` o abriéndolo en RStudio como
proyecto de esa carpeta), porque usa rutas relativas.

Paquetes de CRAN:

```r
install.packages(c("openxlsx", "dplyr", "tidyr", "tibble", "stringr", "stringi",
                   "ggplot2", "ggrepel", "plotly", "gt", "scales", "cluster",
                   "pxR", "httr", "jsonlite", "rvest", "trend", "pdftools",
                   "readxl", "knitr", "rmarkdown"))
```

Solo para volver a descargar el PIB, la población o las viviendas turísticas
hace falta además `ineapir`, del INE:
`remotes::install_github("es-ine/ineapir")`.

Para generar el HTML del informe hace falta pandoc, que viene incluido con
RStudio y con Quarto.

### Opción 1: solo el informe

Todos los datos que usa el informe ya están en el repositorio. Basta con
abrir `INE_ICV_CCAA/INE_ICV_CCAA_all.Rmd` en RStudio y pulsar *Knit*. Lo mismo
para el estudio de los archipiélagos, con `INE_ICV_CCAA/INE_ICV_CCAA_Insulares.Rmd`,
y para el europeo, con `EUROSTAT_INSULARES/EUROSTAT_INSULARES.Rmd`.

### Opción 2: rehacer el conjunto de datos sin conexión

A partir de los ficheros originales incluidos (los `.px` del INE y los ficheros
de Registradores, SERPAVI, PIB e IPC), en este orden:

```
INE_ICV/INE_ICV_CCAA_Procesar.R                 # .px  ->  las 42 tablas por comunidad
INE_ICV_CCAA/INE_ICV_CCAA_ConstruirCatalogo.R
INE_ICV_CCAA/INE_ICV_CCAA_Consolidar.R          # junta las fuentes y deflacta con el IPC
INE_ICV_CCAA/INE_ICV_CCAA_NombresCortos.R       # siempre después de Consolidar
INE_ICV_CCAA/INE_ICV_CCAA_Baleares_Singularidades.R
INE_ICV_CCAA/INE_ICV_CCAA_Baleares_Tendencias.R
INE_ICV_CCAA/INE_ICV_CCAA_all.Rmd               # el informe
```

Se ha comprobado que, partiendo de una copia limpia, esta cadena regenera
exactamente los mismos ficheros que hay en el repositorio.

Para el estudio de los archipiélagos, después de la cadena anterior:

```
PIB_CCAA/PIB_CCAA_VAB_Ramas.R                    # peso de cada rama en el VAB (del .px del PIB)
MIVAU/MIVAU_DataExtraction.R                     # usa los ficheros originales incluidos
INE_ICV_CCAA/INE_ICV_CCAA_Insulares_Consolidar.R # los 15 indicadores nuevos, por habitante
INE_ICV_CCAA/INE_ICV_CCAA_Insulares.Rmd          # el estudio
```

La población y las viviendas turísticas se obtienen de las API del INE y de
Eurostat, sin fichero original que guardar: sus resultados
(`POBLACION_CCAA/POBLACION_CCAA.xlsx` y
`VIVIENDA_TURISTICA/VIVIENDA_TURISTICA_CCAA.xlsx`) están incluidos.

Para el estudio europeo, a partir de las series de Eurostat y las distancias
incluidas:

```
EUROSTAT_INSULARES/EUROSTAT_INSULARES_Consolidar.R   # indicadores relativos y referencias continentales
EUROSTAT_INSULARES/EUROSTAT_INSULARES.Rmd            # el estudio
```

También se ha comprobado que el estudio de los archipiélagos y el europeo
regeneran exactamente los mismos datos a partir de lo incluido.

### Opción 3: volver a descargar todo de las fuentes

Los scripts de descarga son independientes entre sí y se ejecutan antes que
los de la opción 2:

```
INE_ICV/INE_ICV_DataExtraction.R          # recorre el catálogo del ICV y descarga los .px
INE_ICV/INE_ICV_AnadirMetadatos.R
REGISTRADORES/REGISTRADORES_DataExtraction.R
REGISTRADORES/REGISTRADORES_Extranjeros_DataExtraction.R   # descarga los anuarios en PDF (~100 MB)
SERPAVI/SERPAVI_DataExtraction.R          # descarga la base de datos completa (71 MB)
PIB_CCAA/PIB_CCAA_DataExtraction.R
IPC_CCAA/IPC_CCAA_DataExtraction.R
POBLACION_CCAA/POBLACION_CCAA_DataExtraction.R            # estudio de los archipiélagos
MIVAU/MIVAU_DataExtraction.R              # borrar antes sus dos ficheros originales para descargarlos de nuevo
VIVIENDA_TURISTICA/VIVIENDA_TURISTICA_DataExtraction.R
EUROSTAT_INSULARES/EUROSTAT_INSULARES_DataExtraction.R    # estudio europeo: API de Eurostat (caché en cache/)
EUROSTAT_INSULARES/EUROSTAT_INSULARES_Distancias.R        # descarga las geometrías de GISCO y calcula las distancias
```

Los anuarios de la Estadística Registral Inmobiliaria y la base de datos
completa de SERPAVI no se incluyen en el repositorio por su tamaño: los
descargan estos scripts. Si las fuentes han publicado datos nuevos o han
cambiado la dirección de algún fichero, los resultados pueden variar.

## Fuentes y licencias de los datos

| Fuente | Datos | Condiciones de reutilización |
|---|---|---|
| [INE](https://www.ine.es), Indicadores de Calidad de Vida, Contabilidad Regional de España, Cifras de Población, Estadística Continua de Población, IPC y estadística experimental de viviendas turísticas | 79 indicadores, PIB y VAB por ramas, población, IPC, viviendas turísticas | [CC BY 4.0](https://www.ine.es/dyngs/AYU/index.htm?cid=125), citando: «Elaboración propia con datos extraídos del sitio web del INE: www.ine.es» |
| [Colegio de Registradores](https://opendata.registradores.org), portal de datos abiertos | Precio de compraventa de vivienda | [ODbL 1.0](https://opendatacommons.org/licenses/odbl/1-0/) |
| Colegio de Registradores, anuarios de la Estadística Registral Inmobiliaria | % de compras de vivienda por extranjeros | Datos publicados por el Colegio; se cita la fuente |
| [Ministerio de Vivienda y Agenda Urbana](https://www.mivau.gob.es/vivienda/alquila-bien-es-tu-derecho/serpavi), SERPAVI | Precio de alquiler de vivienda | [Aviso legal del Ministerio](https://www.mivau.gob.es/ministerio/aviso-legal): reutilización citando la fuente y la fecha de actualización (base de datos 2011-2024, publicada en 2026) |
| [Ministerio de Vivienda y Agenda Urbana](https://apps.fomento.gob.es/BoletinOnline2/?nivel=2&orden=34000000), Estadística de Transacciones Inmobiliarias (tabla 1.6) | Compraventas de vivienda según la residencia del comprador | [Aviso legal del Ministerio](https://www.mivau.gob.es/ministerio/aviso-legal): reutilización citando la fuente y la fecha de actualización (descarga de noviembre de 2026, datos hasta el 2.º trimestre de 2026) |
| [Instituto Canario de Estadística (ISTAC)](https://datos.canarias.es/catalogos/estadisticas/dataset/urn-siemac-org-siemac-metamac-infomodel-statisticalresources-dataset-istac-e25004a_000001), datos abiertos: serie de viviendas libres y protegidas iniciadas y terminadas del Ministerio de Vivienda y Agenda Urbana | Viviendas iniciadas | [Aviso legal del ISTAC](https://www.gobiernodecanarias.org/istac/aviso_legal.html): reutilización libre citando «Fuente: Instituto Canario de Estadística (ISTAC)» |
| [Eurostat](https://ec.europa.eu/eurostat/databrowser/product/page/tour_ce_omn12), `tour_ce_omn12` | Noches en alojamientos de corta estancia reservados por plataformas en línea | [Aviso de copyright de Eurostat](https://ec.europa.eu/eurostat/help/copyright-notice): reutilización libre citando la fuente e indicando las modificaciones (aquí, el cociente por habitante) |
| [Eurostat](https://ec.europa.eu/eurostat/web/regions/database), base de datos regional (NUTS 2) y nacional | Las 69 series del estudio europeo (catálogo en `EUROSTAT_INSULARES_Series.xlsx`) | [Aviso de copyright de Eurostat](https://ec.europa.eu/eurostat/help/copyright-notice): reutilización libre citando la fuente e indicando las modificaciones (aquí, los cocientes, las referencias continentales y las comparaciones) |
| [GISCO](https://ec.europa.eu/eurostat/web/gisco/geodata/statistical-units) (Eurostat), regiones NUTS 2021 | Geometrías para calcular las distancias al continente | Uso **no comercial**, citando «© EuroGeographics for the administrative boundaries». No se incluyen en el repositorio: las descarga `EUROSTAT_INSULARES_Distancias.R` |

El precio de compraventa procede de una base de datos con licencia ODbL, que
exige que las bases de datos derivadas se compartan con la misma licencia: los
ficheros de `REGISTRADORES/` y la parte del consolidado que procede de ellos se
distribuyen bajo **ODbL 1.0**.

## Licencia del estudio

- **Código** (`.R`, `.Rmd`): [MIT](LICENSE).
- **Texto del informe, figuras y datos elaborados**: [CC BY 4.0](LICENSE-CC-BY-4.0.txt),
  salvo los datos derivados del Colegio de Registradores, que mantienen la ODbL 1.0,
  y las distancias del estudio europeo (`EUROSTAT_INSULARES_Distancias.xlsx`), que
  salen de las geometrías de GISCO y mantienen su condición de uso no comercial.

Los datos de origen mantienen las condiciones de cada fuente (tabla de arriba).

## Cómo citarlo

> Álvarez León, L., Coll Vicens, B., Vicens Colom, J. y Vicens Gómez, J. M.
> (2026). *La calidad de vida en Illes Balears: una visión multidimensional*.
> https://lalvarezmat.github.io/icv-illes-balears/
