# La calidad de vida en Illes Balears: una visión multidimensional

¿En qué se diferencia Illes Balears del resto de España en calidad de vida, y
hacia dónde se mueve? Este repositorio contiene el estudio completo, con su
código y sus datos, para que cualquiera pueda reproducirlo.

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

## Contenido

| Carpeta | Qué hay |
|---|---|
| `INE_ICV_CCAA/` | **El informe** (`INE_ICV_CCAA_all.Rmd`), los scripts que construyen el conjunto de datos del estudio y sus resultados: `INE_ICV_CCAA_Consolidado.xlsx` (los 82 indicadores por comunidad y año), el nivel y la tendencia de Illes Balears frente a las demás comunidades, y las 42 tablas del INE ya procesadas. |
| `INE_ICV/` | Los ficheros `.px` originales de los Indicadores de Calidad de Vida del INE (las 480 tablas del catálogo, sin modificar), su catálogo y los scripts que los descargan y procesan. |
| `REGISTRADORES/` | Precio de compraventa de vivienda (Colegio de Registradores, datos abiertos) y porcentaje de compras por extranjeros (anuarios de la Estadística Registral Inmobiliaria). |
| `SERPAVI/` | Precio de alquiler de vivienda (Ministerio de Vivienda y Agenda Urbana), ya pivotado por comunidad y año. |
| `PIB_CCAA/` | PIB per cápita real por comunidad: las dos tablas `.px` originales del INE y el resultado. |
| `IPC_CCAA/` | IPC por comunidad (INE), usado para descontar la inflación de la renta y de los precios. |

## Cómo reproducirlo

Todo está escrito en **R** (probado con R 4.4.2). Cada script se ejecuta
**desde su propia carpeta** (con `setwd()` o abriéndolo en RStudio como
proyecto de esa carpeta), porque usa rutas relativas.

Paquetes de CRAN:

```r
install.packages(c("openxlsx", "dplyr", "tidyr", "tibble", "stringr", "stringi",
                   "ggplot2", "ggrepel", "plotly", "gt", "scales", "cluster",
                   "pxR", "httr", "jsonlite", "rvest", "trend", "pdftools",
                   "knitr", "rmarkdown"))
```

Solo para volver a descargar el PIB hace falta además `ineapir`, del INE:
`remotes::install_github("es-ine/ineapir")`.

Para generar el HTML del informe hace falta pandoc, que viene incluido con
RStudio y con Quarto.

### Opción 1: solo el informe

Todos los datos que usa el informe ya están en el repositorio. Basta con
abrir `INE_ICV_CCAA/INE_ICV_CCAA_all.Rmd` en RStudio y pulsar *Knit*.

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
```

Los anuarios de la Estadística Registral Inmobiliaria y la base de datos
completa de SERPAVI no se incluyen en el repositorio por su tamaño: los
descargan estos scripts. Si las fuentes han publicado datos nuevos o han
cambiado la dirección de algún fichero, los resultados pueden variar.

## Fuentes y licencias de los datos

| Fuente | Datos | Condiciones de reutilización |
|---|---|---|
| [INE](https://www.ine.es), Indicadores de Calidad de Vida, Contabilidad Regional de España, Cifras de Población e IPC | 79 indicadores, PIB, población, IPC | [CC BY 4.0](https://www.ine.es/dyngs/AYU/index.htm?cid=125), citando: «Elaboración propia con datos extraídos del sitio web del INE: www.ine.es» |
| [Colegio de Registradores](https://opendata.registradores.org), portal de datos abiertos | Precio de compraventa de vivienda | [ODbL 1.0](https://opendatacommons.org/licenses/odbl/1-0/) |
| Colegio de Registradores, anuarios de la Estadística Registral Inmobiliaria | % de compras de vivienda por extranjeros | Datos publicados por el Colegio; se cita la fuente |
| [Ministerio de Vivienda y Agenda Urbana](https://www.mivau.gob.es/vivienda/alquila-bien-es-tu-derecho/serpavi), SERPAVI | Precio de alquiler de vivienda | [Aviso legal del Ministerio](https://www.mivau.gob.es/ministerio/aviso-legal): reutilización citando la fuente y la fecha de actualización (base de datos 2011-2024, publicada en 2026) |

El precio de compraventa procede de una base de datos con licencia ODbL, que
exige que las bases de datos derivadas se compartan con la misma licencia: los
ficheros de `REGISTRADORES/` y la parte del consolidado que procede de ellos se
distribuyen bajo **ODbL 1.0**.

## Licencia del estudio

- **Código** (`.R`, `.Rmd`): [MIT](LICENSE).
- **Texto del informe, figuras y datos elaborados**: [CC BY 4.0](LICENSE-CC-BY-4.0.txt),
  salvo los datos derivados del Colegio de Registradores, que mantienen la ODbL 1.0.

Los datos de origen mantienen las condiciones de cada fuente (tabla de arriba).

## Cómo citarlo

> Álvarez León, L., Coll Vicens, B., Vicens Colom, J. y Vicens Gómez, J. M.
> (2026). *La calidad de vida en Illes Balears: una visión multidimensional*.
> https://github.com/lalvarezmat/icv-illes-balears
