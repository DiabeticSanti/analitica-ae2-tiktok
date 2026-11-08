# AE3 – Transformación y enriquecimiento de datos

(Ver también el `README.md` de la raíz del repositorio.)


Continúa con el mismo dataset y la misma necesidad de AE2 (escenario hipotético: el equipo de contenidos de una agencia que gestiona cuentas de marcas en TikTok quiere saber cómo se comportan el alcance y la interacción según el formato y el carácter comercial de las publicaciones).

### Preguntas de negocio y tablas que las responden
| N.º | Pregunta | Tabla (`AE3/datos/tablas/`) |
|---|---|---|
| P1 | ¿Cómo se distribuyen las vistas y la interacción en la muestra? | `p1a_resumen_general.csv`, `p1b_rango_vistas.csv` |
| P2 | ¿Cómo varían según el tipo de publicación (video o foto)? | `p2_tipo_publicacion.csv` |
| P3 | ¿Cómo varían según el rango de duración? | `p3_rango_duracion.csv` |
| P4 | ¿Cómo varían según el país? | `p4_pais.csv`, `p4b_pais_completo.csv` |
| P5 | ¿Cómo varían según el carácter comercial (publicidad, TikTok Shop, contenido de marca)? | `p5a_categoria_comercial.csv`, `p5b_indicadores_comerciales.csv` |
| P6 | ¿Cómo cambian según el año de publicación? (pregunta de contexto añadida) | `p6_anio_publicacion.csv` |

Los ajustes respecto a AE2 (desglose en P1–P6, pregunta P6, país agrupado, duración solo en videos, tasa agregada y selección de variables) están justificados en el notebook (sección 1) y en el documento (sección 8.1).

### Transformaciones (se registran en `datos/calidad/bitacora_transformaciones.csv`)
1. **T1 – Tipos:** conteos `uint64` → `Int64`; fechas a UTC; indicadores 0/1 → booleanos.
2. **T2 – Duplicados** por `video_id` (se conserva el registro con métricas más recientes; se informa aunque no haya).
3. **T3 – Normalización:** `type` en minúsculas; `country` en mayúsculas y vacío → `SIN DATO`; `branded_content` vacío → `ninguno`.
4. **T4 – Controles de calidad** (`datos/calidad/controles_calidad.csv`): fechas incoherentes, vistas nulas, publicaciones sin vistas, me gusta > vistas, duración inconsistente, tipo no reconocido, indicadores sin dato. Los registros se **marcan** (`f_*`), no se eliminan; solo se excluyen de los cálculos que dependen del campo dudoso.
5. **T5 – Selección de columnas** y redondeo de tasas (6 decimales).

### Enriquecimiento (`datos/calidad/variables_derivadas.csv`)
`anio_publicacion`, `edad_dias`, `duracion_s`, `rango_duracion`, `rango_vistas`, `interacciones_totales`, `tasa_interaccion` (con `tasa_imputada`), `categoria_comercial`, `n_senales_comerciales`, `n_hashtags`, `vistas_por_dia`, `pais_agrupado` y las marcas `f_*`, `analisis_vistas`, `analisis_tasa`. No se usa ninguna fuente externa.

### Cómo ejecutar
```bash
pip install -r requirements.txt
cd AE3
jupyter notebook AE3_Transformacion_Enriquecimiento.ipynb   # Kernel > Restart & Run All
```
Requiere conexión a internet (lee la muestra desde Hugging Face). El notebook exporta los datos, las tablas y el resumen `.txt`, verifica que lo exportado coincida con lo calculado y, en su última celda, actualiza el Word (`python generar_documento.py` hace lo mismo por separado). Guarda el notebook **con las salidas visibles** antes de hacer commit.

### Rutas de salida (relativas a `AE3/`)
| Archivo | Contenido |
|---|---|
| `datos/preparado/tiktok_preparado.csv` | Dataset preparado (transformado y enriquecido) |
| `datos/preparado/diccionario_actualizado.csv` | Diccionario de datos actualizado |
| `datos/tablas/*.csv` | Una tabla por pregunta (listas para graficar en AE4) |
| `datos/calidad/*.csv` | Bitácora, controles de calidad, validaciones, comparación antes/después, variables derivadas, verificaciones |
| `resultados/resumen_dataframe.txt` | Resumen descriptivo del DataFrame **final** (generado con `DatasetSummarizer`, se regenera en cada ejecución) |
| `resultados/metadatos_ejecucion.json`, `interpretaciones.json` | Parámetros de la ejecución y textos interpretativos |
| `documento/Documento_AE3_Avance.docx` | Word de AE2 actualizado con metodología, diccionario, resultados y conclusiones |

### Código
- `hf_dataset_loader.py`: clase `HFDatasetLoader` (carga una muestra reproducible desde Hugging Face).
- `dataset_summarizer.py`: clase `DatasetSummarizer` (resumen `.txt`).
- `generar_documento.py`: actualiza el Word de AE2 a partir de los archivos exportados.

### Datos y licencia
El dataset (`datasocial/tiktok-5.6B-videos`) se publica bajo **CC BY-NC 4.0**: uso no comercial con atribución. El dataset preparado es una muestra derivada y se comparte solo con fines académicos. No se incluyen textos libres (`caption`, `mentions`, `on_screen_text`) ni identificadores de creadores. El CSV preparado pesa unos 45 MB con 200.000 filas; si GitHub lo rechaza o advierte (límite de aviso: 50 MB), reduzca `N_FILAS` en el notebook.
