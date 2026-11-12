# AGENTS.md — Asistente de Cierre de Aulas COIN (Biblioteca)

Herramienta **local y offline** para extraer, cruzar, validar y preparar la
información del cierre de aulas de los cursos COIN de la biblioteca. Procesa 4
archivos Excel oficiales en memoria y genera un archivo auxiliar de cierre (o
aplica el cierre directo sobre los oficiales).

## 🔒 Invariantes de seguridad y privacidad (NO romper)

Estas reglas son la razón de ser del proyecto. Cualquier cambio debe respetarlas
y los tests las protegen:

1. **100% offline-first**: cero llamadas de red, cero APIs externas, cero
   telemetría. `security/offline_check.py` escanea que ningún `.py` importe
   módulos de red (`tests/unit/test_offline.py`).
2. **Fuentes inmutables (solo lectura)**: los 4 archivos oficiales nunca se
   modifican ni sobrescriben sin respaldo explícito. El Modo 5 es la única vía
   de escritura y exige confirmación `"SI"` + backup automático. Se mantiene
   **un único respaldo por archivo** (`_backups_cierre/<nombre>_backup.xlsx`):
   en cada cierre se actualiza con el estado previo a la automatización y se
   consolidan (eliminan) las copias antiguas con marca de tiempo del mismo
   archivo.
3. **Auditoría SHA-256**: hash antes y después de cada operación
   (`security/integrity.py`, `tests/unit/test_immutability.py`).
4. **Escritura segura y quirúrgica**: solo se escriben celdas estrictamente
   vacías (y los nombres a MAYÚSCULAS); ninguna celda preexistente con datos se
   altera y **ninguna otra parte del libro se reserializa**: el resto del paquete
   `.xlsx` se conserva byte a byte (`output/xlsx_patcher.py`). Notas es siempre
   solo lectura. **Excepciones autorizadas**: (a) si el usuario lo pide en el
   Modo 5 (`actualizar_nota_final_existente`), se corrige la columna de nota
   final de Sistematización cuando ya tiene un valor distinto al de Notas;
   (b) la columna *Fecha de envío* de Sistematización se corrige a `No aplica`
   cuando el estudiante no aprobó o abandonó y la celda traía otra cosa.
5. **Matching ultra-conservador**: prohibido asignar notas por coincidencia
   difusa de nombres. Se exige doble factor de identidad
   (Documento + Nombre/Correo) — `matching/engine.py`.
6. **Aislamiento de errores por registro**: un estudiante inconsistente no
   bloquea al resto.
7. **Enmascaramiento PII**: documentos, correos y nombres se enmascaran en
   logs/reportes de consola y de descubrimiento (`security/privacy.py`).

## 🚀 Comandos

```bash
# Instalar (única dependencia: openpyxl)
pip install -r requirements.txt

# Ejecutar el menú interactivo (Python 3.10+)
python main.py

# Suite de pruebas completa (recomendado: incluye también las clases estilo pytest)
python -m pytest tests -q

# Alternativa con unittest (solo ejecuta las clases unittest.TestCase)
python -m unittest discover -s tests -p "test_*.py" -v

# Inspectores de descubrimiento directo (Fase 0) sobre archivos reales
python tests/discovery/inspect_real_files.py --control "..." --sistematizacion "..." --notas "..." --certificados "..." --aula "ABBUEI154"

# Diagnóstico de cruce de identidades (sin datos personales; útil si hay 0 matches)
# También disponible en el menú de main.py como opción [7].
python tests/discovery/diagnose_matching.py --control "..." --sistematizacion "..." --notas "..." --certificados "..." --aula "CMI120"

# Inspector del archivo de Certificados: ubica un curso, sus filas y si están ocultas
# o dentro de celdas combinadas (sin datos personales).
python tests/discovery/inspect_certificados.py --certificados "Control_codigos_certificados_ABBUEI.xlsx" --curso ABBUEI227
```

El menú de `main.py` (fuente de verdad; el README puede estar desactualizado):

| Opción | Modo |
|---|---|
| 1 | Descubrimiento Técnico Local (Fase 0) |
| 2 | Simulación Previa (validación en memoria, no escribe) |
| 3 | Generar Archivo Auxiliar de Cierre (5 hojas) |
| 4 | Verificación de Arquitectura Offline |
| 5 | Aplicar Cierre Directo en oficiales (con backup) |
| 6 | Restaurar Último Backup |
| 7 | Diagnóstico de Cruce de Identidades (sin datos personales) |
| 8 | Cierre por Lotes (varias aulas; Sistematización y Control compartidos) |
| 9 | Diagnosticar / Reparar un `.xlsx` dañado (crea `<nombre>_reparado.xlsx`) |

## 🧩 Arquitectura (módulos)

- `main.py` — CLI / menú interactivo (modos 0–8, incluido el cierre por lotes).
- `core/` — `enums.py` (MatchKey, StudentStatus, AulaReadiness…),
  `models.py` (AulaMetadata, AulaCierreResult, ClaseGroup, Estudiante*),
  `traceability.py` (`TracedValue` y `ValidationLog`: valor + procedencia
  celda a celda / auditoría).
- `config/` — `aliases.py`, `colors.py` (detección del naranja), `settings.py`.
- `input/` — `file_selector.py` (diálogos nativos de archivos).
- `parsing/` — parsers de los 4 archivos: `control_parser.py`,
  `sistematizacion_parser.py`, `notas_parser.py`, `certificados_parser.py`.
- `normalization/` — `dates.py` (`format_ciclo_date`, `format_date_certificados`),
  `headers.py` (`analyze_headers`, `find_module_columns`),
  `text.py` (`normalize_document`, `normalize_email`, `normalize_name_component`).
- `matching/` — `engine.py` (cruce conservador por clase,
  `verify_cross_identity`, `match_students_in_class`).
- `processing/` — `cierre_orchestrator.py` (`run_cierre_pipeline`),
  `rules_engine.py` (`apply_rules_to_match`, `evaluate_certification_status`).
- `output/` — `direct_updater.py` (`apply_direct_cierre`, `restore_last_backup`,
  `create_timestamped_backup`), `xlsx_patcher.py` (`XlsxPatcher`: edición
  quirúrgica de celdas + `verify_preservation`: verificación de pérdida cero),
  `excel_generator.py` (`generate_cierre_excel`,
  5 hojas: RESUMEN, SISTEMATIZACION, CERTIFICADOS, ACTUALIZACION_CONTROL,
  VALIDACIONES_Y_TRAZABILIDAD).
- `security/` — `integrity.py`, `offline_check.py`, `privacy.py`,
  `validation_files.py`.
- `validation/` — validación cruzada y chequeos de consistencia (paquete
  actualmente con solo `__init__.py`).
- `tests/` — `unit/`, `integration/` (14 escenarios end-to-end),
  `discovery/inspect_real_files.py`, `fixtures/generate_fixtures.py`
  (datos 100% sintéticos, sin datos reales).

## 📝 Convenciones

- Todo en español (código, comentarios, mensajes, informes).
- Estados oficiales de certificación: `Aprobó`, `No aprobó`, `Abandonó`,
  `No aplica`.
- Los fixtures de pruebas son **sintéticos**; jamás committear archivos
  `.xlsx` reales (están en `.gitignore`, salvo `tests/fixtures/*.xlsx`). Se
  regeneran con fecha fija (`_FIXTURE_TIMESTAMP`) para que el contenido sea
  idéntico en cada ejecución y no ensucie el repositorio.
- Reportes generados en `reports/` y `discovery_*.txt` son locales y están
  gitignored.
- Al modificar código, regenerar el grafo de conocimiento: `graphify update .`
  (el grafo vive en `graphify-out/`).

## Estado actual

Fase 0 (descubrimiento) confirmada; modos 2/3/5/6 habilitados en `main.py`.
El proyecto está en fase de ajustes finales de compatibilidad con los archivos
oficiales de la biblioteca. Ajustes recientes (formato oficial):
- Tablas de Sistematización y Certificados: **todos los bordes en todas las
  celdas escritas**; la celda de fin de clase conserva sus tres lados finos y
  añade el borde inferior `medium` (grueso). Bordes en negro (`000000`).
- Columna **Semestre** de Certificados en formato oficial `AAAA-S` (ej. `2026-1`),
  calculada en `normalization/dates.py:format_semester_label`.
- Regla de semestre: S1 = 1 nov → 30 abr; S2 = 1 may → 31 oct
  (`extract_semester`).
- Fuente de la tabla de Certificados: **Zurich Cn BT, tamaño 11**.
- Formato de la columna **Ciclo**: `CLASE {clase} - {catálogo} - {programa} - {MES DÍA} A {MES DÍA}`
  (guion entre clase y catálogo), p. ej. `CLASE 2216 - PD1002 - ...`.
- Formato de la columna **Docente COIN**: `({CÓDIGO-CON-GUIÓN}) {NOMBRE}`,
  p. ej. `(CMI-120) NOMBRE` (`normalization/text.py:format_docente_coin`,
  `format_ciclo`).
- Columna **Fecha de envío** de Certificados en minúsculas (mes abreviado),
  p. ej. `28-sep-2026` (no se aplica `.upper()`).
- Certificados: **toda la información centrada** y la columna **N° (col 1) es la
  única sin borde inferior grueso** en la fila de fin de clase.
- Saneo al guardar oficiales: **ya no aplica**. El cierre no reserializa el libro
  (ver "Preservación TOTAL del libro original"), así que no se eliminan tablas
  dinámicas, referencias externas ni se sustituyen fórmulas. Por eso tampoco
  aparece el aviso de "reparación" de Excel ("Registros quitados: Fórmula de
  sheetN.xml"). `_resolve_external_link_formulas`, `_strip_external_links` y
  `_strip_pivot_tables` quedan como utilidades de legado.
- Fix de claves del mapa de columnas en `apply_direct_cierre`: el parser usa
  `estado`, `elaboro`, `codigo_cert`, `envio` (antes se buscaban
  `estado_certificacion`, etc.), por lo que las columnas *Estado para
  certificación*, *Se elaboró*, *Código del certificado* y *Se envió* quedaban
  vacías.
- **Notas sin columna 'Sección' (p. ej. cursos CMI)**: la sección es opcional
  (`col_seccion=None`); la asignación clase↔notas se hace por identidad
  (documento/correo/nombre) y cada estudiante se asigna a UNA sola clase (sin
  duplicar excluidos ni no-match) — `processing/cierre_orchestrator.py`.
- **Detección de color naranja restringida** a las columnas del estudiante
  (identidad, nombre y calificaciones). Así columnas de estado/semáforo con
  rellenos ámbar/dorado (`FFC000`, etc.) NO marcan a todo el curso como excluido.
- **Nota final**: se reconoce también `Promedio curso completo`
  (`promedio_curso`) además de `Calculated Final`/`Nota Final`
  (`normalization/headers.py`).

- **Preservación TOTAL del libro original (arquitectura definitiva)**: el cierre
  directo **ya no guarda el libro con openpyxl**. `openpyxl` no es un editor
  *round-trip* fiel y al guardar descarta lo que no sabe representar: listas de
  validación modernas (`x14:dataValidations`), `extLst` de hoja y de libro
  (formato condicional moderno, minigráficos, segmentaciones, líneas de tiempo),
  controles de formulario (`vmlDrawing`), tablas dinámicas, parte de los estilos
  `dxf`, gráficos con extensiones, vínculos externos y `definedNames` externos.

  Ahora `output/xlsx_patcher.py` (`XlsxPatcher`) edita el `.xlsx` como lo que es,
  un ZIP de partes XML: localiza la hoja y **reemplaza únicamente las celdas
  afectadas** (`<c>` dentro de `<row>` dentro de `<sheetData>`); todo el resto del
  paquete se copia **byte a byte**. Solo se tocan, y solo si hace falta:
  `xl/worksheets/sheetN.xml` (celdas), `xl/styles.xml` (se **añaden** fuentes,
  bordes y `xf`; nada se reescribe) y `xl/workbook.xml` (se marca
  `fullCalcOnLoad="1"` para que Excel recalcule las fórmulas dependientes).

  Ya **no** se eliminan tablas dinámicas, referencias externas ni se sustituyen
  fórmulas: todo eso ahora sobrevive intacto, y por tanto tampoco aparece el aviso
  de "reparación" de Excel. Las funciones `_strip_pivot_tables`,
  `_strip_external_links`, `_resolve_external_link_formulas` y
  `preserve_sheet_data_blocks` siguen disponibles como utilidades de **legado**
  (y siguen cubiertas por tests), pero el Modo 5 no las usa.

- **Verificación de pérdida cero**: tras cada archivo, `verify_preservation`
  (`output/xlsx_patcher.py`) compara el respaldo y el resultado y falla si (a)
  desaparece cualquier parte del paquete, (b) cambia una parte no prevista, (c)
  falta alguna lista de validación (huella por huella, estándar y `x14`), regla de
  formato condicional o nombre definido, o (d) se pierden tablas dinámicas,
  gráficos, imágenes o controles de formulario. El resultado se imprime en el
  Modo 5 (`report.validaciones_preservadas`) y cualquier problema se añade a
  `report.warnings`.

  > **Fix crítico (reporte real CMI121, 2026-10-23):** el detector de prefijos
  > (`_prefixes_used`) confundía rangos de celdas de un `sqref` multirrango
  > (`"B2:B9 AD2:AF2"`) con prefijos de namespace, descartaba el `extLst` completo
  > y se perdían las 8 listas modernas (x14). Ahora un atributo con prefijo debe
  > llevar `=` tras el nombre local, y los prefijos implícitos (`xml`) se ignoran.

- **Edición quirúrgica de celdas (`XlsxPatcher`)**: escribe cadenas como
  `t="inlineStr"` para **no tocar `xl/sharedStrings.xml`**; clona el `xf` de la
  celda para conservar fuente, relleno, alineación y formato numérico, y solo
  cambia el borde; hereda el estilo de la columna (`<col style>`) cuando la celda
  no existía; amplía `<dimension>` si añade filas; y no reescribe el archivo si no
  hubo cambios. Escapa `&`, `<`, `>` y usa `xml:space="preserve"`.

- **Fecha de envío en Sistematización**: la columna *Fecha de envío* se escribe
  como fecha oficial `DD-mes-AAAA` con mes abreviado en minúsculas (ej.
  `24-mar-2026`), nunca como `str(datetime)` (`2026-03-24 00:00:00`, que además
  arrastraba la hora en ceros) — `processing/rules_engine.py:_format_fecha_envio_sistematizacion`.
  En el archivo auxiliar (Modo 3) esa columna se excluye del `upper()` aplicado
  al resto de la fila — `output/excel_generator.py`.
- **La fecha de envío depende del estado**: la columna *Fecha de envío* de
  Sistematización lleva la fecha de finalización de la clase **solo si el
  estudiante aprobó**; si `No aprobó` o `Abandonó`, dice `No aplica`
  (`processing/rules_engine.py:apply_rules_to_match`). Si una celda de un no
  aprobado ya trae una fecha escrita (p. ej. de un cierre anterior), el cierre
  directo la corrige a `No aplica`.
- **Aviso de módulos faltantes**: Sistematización tiene 5 columnas de módulos,
  pero muchos cursos tienen menos. Las columnas sobrantes ya no quedan vacías:
  llevan `El curso solo tiene X módulos` (singular `1 módulo`), donde X es el
  número de módulos del curso (columnas de módulo detectadas en Notas, sin
  contar la nota final / promedio del curso completo). Se calcula en
  `processing/rules_engine.py:_texto_modulos_faltantes` y solo rellena celdas
  vacías (nunca pisa una nota existente). En el archivo auxiliar el aviso
  conserva su redacción (no se pasa a MAYÚSCULAS).
- El modo **7 (Diagnóstico)** incluye una **simulación completa del Modo 5**
  sobre copias (parsers, matcheo, edición quirúrgica y escritura) que valida el
  XML de cada hoja resultante con expat, confirma que las listas de validación
  sobreviven y verifica parte por parte (`verify_preservation`) que no se pierda
  nada del libro original; reporta el error exacto si lo hubiera.

- **Selección automática de hoja en Notas**: si el archivo tiene varias hojas,
  se elige la que tenga columnas de módulos/nota final + identidad (p. ej.
  'CURSO REALIZADO'), en lugar de asumir 'Calificaciones' por nombre. Caso real
  CMI: 'Calificaciones' traía solo identidad y sección de otro curso, y las notas
  reales estaban en 'CURSO REALIZADO' — `parsing/notas_parser.py:_select_notas_sheet`.

- **Detección de módulos flexible**: además de "módulo N" se reconoce
  "Cuestionario N", "Buzón N", "Foro N", "Taller N", etc.
  (`normalization/headers.py:_MODULE_REGEX`). Caso real CMI121: la columna
  "Cuestionario 1: ..." no se detectaba y el módulo 1 se ignoraba.

- **Respaldo POSICIONAL de Notas (encabezados atípicos)**: si los encabezados no
  permiten reconocer los módulos o la nota final, `parsing/notas_parser.py:
  _infer_grade_columns` infiere las columnas por posición: de izquierda a
  derecha, módulo 1…N y la **última columna numérica** es la nota final. Solo se
  activa cuando falta el mapeo por encabezado (nunca lo reemplaza). Cada columna
  se valida con sus **datos**: mayoría de valores numéricos, sin booleanos, sin
  fechas, sin formato de porcentaje, sin identidad y con valores en rango
  plausible (máx. 1000, mediana ≤ 100). Se toma el grupo de columnas numéricas
  más a la derecha (tolera un hueco) y se limita a **5 módulos** (regla oficial).
  Si solo faltan algunos módulos y los detectados son un prefijo 1..k, se
  completan por posición (`metodo_modulos` = `"mixto"`).
  Las columnas simbólicas (`Aprobado`, `No aplica`, etc.) **no** se usan como
  nota final en este respaldo: podrían ser un estado; se avisa. El diagnóstico
  (hoja, método, columnas de módulos y final, avisos) se muestra en pantalla y
  viaja en `AulaCierreResult.notas_deteccion`; también se usa para elegir la
  hoja correcta cuando los títulos son atípicos (`_count_numeric_grade_columns`).
- **Identidad en Notas**: se reconoce la columna "Documento" como documento del
  estudiante (además de "OrgDefinedId") — `parsing/notas_parser.py`.

> Nota: el color de borde de las tablas escritas se fijó en negro para que sea
> visible y consistente. Si el documento oficial usa otro color, se cambia en la
> constante `_BORDER_COLOR` (ARGB, `FF000000`) de `output/direct_updater.py` y en
> `output/excel_generator.py`.

- **Anexado de Certificados a prueba de fallos (reporte real ABBUEI227, 2026-11-06)**:
  el punto de anexado ya **no** se calculaba con "última fila con algo en A o B".
  Una nota, total o leyenda suelta debajo de la tabla desplazaba los certificados
  cientos/miles de filas y el usuario concluía que "no se guardaron". Ahora
  `_inspect_certificate_table` reconoce los **registros reales** (código del
  certificado en H o nombres+apellidos en D/E) e ignora lo ajeno a la tabla;
  `_fila_certificados_vacia` impide sobrescribir contenido; `show_row` de
  `XlsxPatcher` desoculta filas de plantilla; se **omiten duplicados** (mismo
  documento y curso) y, tras guardar, `_verify_certificate_rows` **relee el archivo**
  y confirma fila por fila lo anexado (N°, documento y código). El Modo 5 informa
  la hoja, el rango exacto de filas y el resultado de la verificación, además de
  imprimir las rutas completas de los 4 archivos seleccionados. Nuevo inspector
  local `tests/discovery/inspect_certificados.py` (sin PII) para ubicar un curso
  en el archivo, ver si sus filas están ocultas, filtradas o combinadas.

- **La nota final se copia, NUNCA se promedia (reporte real ABBUEI227, 2026-11-06)**:
  `evaluate_certification_status` ya no usa `sum(notas)/len(notas)` como valor de
  la nota final. La columna *Promedio* de Sistematización recibe **exactamente** el
  valor de la nota final del archivo de Notas (numérico o simbólico); si Notas no
  la trae, la celda queda vacía y el Modo 5 lo advierte
  (`sistematizacion_promedios_no_copiados`). El promedio de módulos se conserva
  solo para la regla oficial de abandono (`AVERAGE(AS:AV)=0`). Además:
  - `parsing/notas_parser.py` lee una segunda vez con `data_only=True`, de modo
    que una nota final escrita como **fórmula** usa su valor calculado en caché
    (con `data_only=False` se leía el texto `=AVERAGE(...)` y se perdía el valor).
  - `normalization/headers.py` reconoce como columna de nota final, sin
    distinciones ni descartes por el nombre, `Calculated Final Grade`,
    `Calculated Final Grade Scheme Symbol`, `Calculated Final`, `Final`,
    `Final Grade`, `Nota final`, `Nota definitiva`, `Calificación final`,
    `Definitiva`, `Final Score`, `Puntaje final`, `Promedio final` y
    `Promedio curso completo`. `_detectar_nota_final` puntúa los encabezados y
    descarta falsos positivos de fecha/hora/identidad/actividades (`Fecha
    Final`, `Hora final`, `Examen final`, `Trabajo final`, etc.).
  - **La elección entre varias columnas candidatas es por DATOS, no por el
    nombre**: `parsing/notas_parser.py:_elegir_columna_nota_final` prefiere la
    columna con más valores numéricos, luego la de más valores no vacíos y, al
    final, el mayor puntaje del encabezado. Así, si en un archivo las notas
    reales están en `Calculated Final Grade Scheme Symbol` (aunque el nombre
    sugiera símbolos), se usa esa columna.
  - En el archivo auxiliar (Modo 3) la columna de promedio tampoco se pasa a
    mayúsculas, para conservar el valor literal de Notas.
- **Un solo respaldo por archivo**: `create_timestamped_backup` ahora crea/actualiza
  `_backups_cierre/<nombre>_backup.<ext>` (sin marca de tiempo) con el estado
  previo a cada cierre, de forma atómica (`.tmp` + `os.replace`), y elimina las
  copias antiguas `<nombre>_backup_<timestamp>.<ext>` del mismo archivo.
  `find_latest_backup` devuelve el respaldo único (o el más reciente del formato
  anterior si aún no migró) y el Modo 6 restaura ese estado previo al último
  cierre.

- **Cierre a prueba de bloqueos y re-ejecuciones (reporte real, 2026-11-06)**:
  - **Preflight**: antes de crear respaldos, `files_not_writable` abre cada
    oficial en `r+b`; si está abierto en Excel o es de solo lectura, el Modo 5
    aborta con un mensaje accionable y no modifica nada (antes el error aparecía
    a mitad del cierre como `[WinError 5] Acceso denegado`).
  - **Rollback**: si un paso falla de forma crítica (`fallo_critico`: escritura,
    bloqueo o verificación de certificados), `_restore_after_error` devuelve los
    tres oficiales a su estado previo usando los respaldos del inicio. Un error
    de un registro (p. ej. discrepancia de identidad) NO revierte todo:
    se mantiene el aislamiento de errores por registro.
  - `XlsxPatcher._write` traduce el `PermissionError` de `os.replace` a un
    mensaje claro (archivo abierto en Excel / solo lectura / OneDrive).
  - **Corrección opcional de la nota final ya escrita**: el Modo 5 pregunta
    *"¿Corregir notas finales existentes que no coincidan con Notas? (SI/NO)"*.
    Con `actualizar_nota_final_existente=True` se sobrescribe **solo** la columna
    de promedio cuando difiere del valor de Notas (contador
    `sistematizacion_promedios_actualizados`); por defecto se respeta la regla de
    no sobrescribir.

- **Cierre por lotes (Modo 8)**: `main.py:handle_batch_cierre_mode` permite cerrar
  varias aulas en una sola pasada. La **Sistematización** y el **Control de Aulas**
  se piden una sola vez (compartidos); las **Notas** se eligen en un diálogo de
  selección múltiple y el **archivo de Certificados** puede ser compartido o
  distinto por aula (con opción de reutilizar el anterior con ENTER). El código de
  aula se deduce del nombre de las Notas (`_infer_aula_code`, ej.
  `ABBUEI227_...xlsx → ABBUEI227`). Primero simula todas las aulas y muestra un
  resumen consolidado; con una sola confirmación `SI` aplica cada aula como un
  cierre independiente (respaldo único actualizado por aula, deduplicación y
  numeración real de certificados al escribir). Si un aula falla de forma crítica,
  se detiene el lote y se informa cuáles quedaron cerradas.

- **Blindaje contra libros dañados y recuperación (reporte real «CARGA BLOQUEADA»)**:
  - **Fórmulas compartidas**: Excel guarda las fórmulas arrastradas como un grupo
    (`<f t="shared" ref="L2:L3" si="0">…</f>` en la maestra y
    `<f t="shared" si="0"/>` en las esclavas). `XlsxPatcher.set_cell` **nunca**
    sobrescribe la celda maestra ni una fórmula matricial: lanza
    `CeldaProtegidaError` y `apply_direct_cierre` omite esa celda y lo informa
    (`sistematizacion_celdas_protegidas`). Antes, la corrección de la nota final
    reemplazaba la maestra, las esclavas quedaban huérfanas y Excel "reparaba" el
    libro (perdiendo fórmulas) y mostraba «CARGA BLOQUEADA — Se ha producido un
    problema con este archivo…».
  - **Validación antes de escribir**: `XlsxPatcher.save` valida con
    `validate_sheet_xml` (XML bien formado, sin caracteres de control, sin
    esclavas huérfanas, sin celdas duplicadas) y `styles.xml`/`workbook.xml`
    (un solo `<calcPr>`); si algo falla, **no escribe nada** y `apply_direct_cierre`
    revierte. `_apply_recalc` ya no duplica `<calcPr>` cuando no es autocerrado.
  - **Preflight del cierre**: `_diagnosticar_libros_oficiales` analiza los tres
    oficiales antes de crear respaldos; si uno ya está dañado, aborta para no
    empeorarlo ni sobrescribir un respaldo sano.
  - **Recuperación**: `output/xlsx_recovery.py` (`diagnose_xlsx`, `repair_xlsx`;
    opción [9]) diagnostica ZIP/XML, fórmulas compartidas huérfanas, celdas
    duplicadas, estilos e índices de `sharedStrings` fuera de rango, `<calcPr>`
    duplicado y ampersands sin escapar. Repara en una **copia**
    (`<nombre>_reparado.xlsx`), nunca sobre el original, y deja un
    `<nombre>_diagnostico.txt`. El Modo 6 valida el respaldo único con
    `diagnose_xlsx` antes de restaurar y advierte si también está dañado (en ese
    caso restaurar no arregla el libro).
  - **Robustez adicional**: se sanean caracteres de control XML al serializar
    valores; `nan`/`inf` no se escriben; el borde de Sistematización se acota al
    ancho real de la tabla (`_ancho_tabla_sistematizacion`, máx. 55 columnas) para
    no crear miles de celdas si `max_column` viene inflado; las filas/celdas sin
    atributo `r` (opcional en OOXML) se conservan asignando su posición implícita
    en lugar de descartarse.
  - **Conflictos de combinación (reporte real AR10)**: el cierre se detiene si el
    libro está abierto (archivo de bloqueo `~$<nombre>.xlsx`), si es un **libro
    compartido** (`<fileSharing>` o `xl/revisions/`) o si cambió mientras se
    procesaba (`XlsxPatcher` compara mtime/tamaño justo antes de `os.replace`).
    El aviso «Conflicto de combinación / Cambios sin combinar» aparece cuando
    Excel mantiene el libro cargado o el libro usa control de cambios compartido:
    la automatización no debe escribir sobre esa sesión. En carpetas de
    OneDrive/SharePoint se pide confirmación explícita de que el archivo está
    cerrado (escritorio y web) y sincronizado.
  - Se informa el **mapa de columnas** usado en Sistematización (letra, índice y
    encabezado) para detectar de inmediato una detección equivocada (p. ej. un
    `No aplica` escrito en AR en lugar de la columna real). El cierre conserva
    tablas dinámicas, listas de validación, gráficos e imágenes: solo edita las
    celdas y añade estilos; nunca elimina partes del paquete.

Última verificación: **183 tests OK** (`python -m pytest tests -q`). El grafo de
conocimiento se regenera con `graphify update .` tras cada cambio de código.
