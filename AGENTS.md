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
   de escritura y exige confirmación `"SI"` + backup automático.
3. **Auditoría SHA-256**: hash antes y después de cada operación
   (`security/integrity.py`, `tests/unit/test_immutability.py`).
4. **Escritura segura**: solo se escriben celdas estrictamente vacías; nombres
   a MAYÚSCULAS; ninguna celda preexistente con datos se altera. Notas es
   siempre solo lectura.
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

# Suite de pruebas completa
python -m unittest discover -s tests -p "test_*.py" -v

# Inspectores de descubrimiento directo (Fase 0) sobre archivos reales
python tests/discovery/inspect_real_files.py --control "..." --sistematizacion "..." --notas "..." --certificados "..." --aula "ABBUEI154"
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

## 🧩 Arquitectura (módulos)

- `main.py` — CLI / menú interactivo (modos 0–6).
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
  `create_timestamped_backup`), `excel_generator.py` (`generate_cierre_excel`,
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
  `.xlsx` reales (están en `.gitignore`, salvo `tests/fixtures/*.xlsx`).
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

> Nota: el color de borde de las tablas escritas se fijó en negro para que sea
> visible y consistente. Si el documento oficial usa otro color, se cambia en la
> constante `_BORDER_COLOR` de `output/direct_updater.py` y
> `output/excel_generator.py`.

Última verificación: 34 tests OK. El grafo de conocimiento se regenera con
`graphify update .` tras cada cambio de código.
