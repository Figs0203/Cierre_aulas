# Asistente de Cierre de Aulas COIN (Biblioteca)
**Herramienta Local de Extracción, Cruce, Validación y Preparación de Información**

---

## 🔒 Garantía Fundamental de Privacidad y Seguridad (Offline-First)

1. **100% Local y Desconectado:** Los archivos Excel oficiales contienen información confidencial y son procesados exclusivamente en la memoria de su computador. El programa no realiza peticiones de red, no utiliza APIs externas, no tiene telemetría ni comparte datos con ninguna IA o servicio en la nube.
2. **Fuentes Inmutables (Solo Lectura):** Los archivos oficiales nunca se modifican ni se sobrescriben.
3. **Auditoría Criptográfica SHA-256:** Se calcula el hash SHA-256 de cada archivo antes de abrirlo y se verifica nuevamente al terminar para certificar matemáticamente que permaneció intacto.
4. **Protección de Datos Personales (PII):** Los reportes de consola y de descubrimiento enmascaran automáticamente documentos (`****4050`), correos (`juan.p***@eafit.edu.co`) y nombres para resguardar la privacidad.

---

## 📋 Requisitos e Instalación

### Requisitos del Sistema
- **Python 3.10 o superior** instalado.
- Sistema Operativo: Windows, macOS o Linux.

### Instalación de Dependencias
La única dependencia externa requerida es `openpyxl` para la manipulación segura de hojas de cálculo `.xlsx`:

```bash
cd d:\Monitoria_COIN\Cierre_de_aulas
pip install -r requirements.txt
```

---

## 🚀 Cómo Ejecutar la Herramienta

### Opción 1: Menú Interactivo Principal (Recomendado)

Ejecute en su terminal:

```bash
python main.py
```

El asistente desplegará un menú interactivo con las siguientes opciones:

```text
MENU PRINCIPAL:
  [1] Modo 1 — Descubrimiento Técnico Local (Fase 0) [HABILITADO]
  [2] Modo 2 — Simulación Previa (Validación en Memoria) [HABILITADO]
  [3] Modo 3 — Generar Archivo Auxiliar de Cierre [HABILITADO]
  [4] Verificación de Arquitectura Offline
  [5] Modo 5 — Aplicar Cierre Directo en Archivos Oficiales (Con Backup y Protección)
  [6] Modo 6 — Restaurar Último Backup en Archivos Oficiales
  [7] Diagnóstico de Cruce de Identidades (Notas ↔ Sistematización, sin datos personales)
  [8] Modo 8 — Cierre por Lotes (varias aulas; Sistematización y Control compartidos)
  [0] Salir
```

1. Seleccione la opción `1`.
2. Se abrirán ventanas nativas de selección de archivos para elegir los 4 archivos Excel oficiales desde cualquier carpeta de su equipo:
   - **Archivo 1:** `Control_Aulas_Formadores_2026.xlsx`
   - **Archivo 2:** `Sistematizacion_Cursos_COIN_2026.xlsx`
   - **Archivo 3:** Notas del aula a cerrar (ej. `Notas_ABBUEI154.xlsx`)
   - **Archivo 4:** Control de certificados correspondiente (ej. `Control_codigos_certificados_ABBUEI.xlsx`)
3. Opcionalmente ingrese el código del aula a buscar (ej. `ABBUEI154`).
4. El sistema analizará los 4 archivos en memoria y generará un informe técnico local en `reports/discovery_YYYYMMDD_HHMMSS.txt`.
5. Al finalizar, el programa certificará la inmutabilidad de los 4 archivos comparando sus hashes SHA-256.

### Cierre por lotes (Modo 8)

Para cerrar varias aulas sin repetir la selección de los archivos compartidos:

1. Seleccione la opción `8`.
2. Elija **una vez** el `Control_Aulas_Formadores_2026.xlsx` y la
   `Sistematizacion_Cursos_COIN_2026.xlsx` (se reutilizan para todas las aulas).
3. Seleccione **varias Notas a la vez** en la ventana de archivos.
4. Indique si todas las aulas usan el mismo archivo de certificados (si no,
   puede reutilizar el anterior con ENTER para cada aula).
5. Confirme el código de aula de cada Notas (se deduce del nombre del archivo) y
   el código de curso para certificados (por defecto, el mismo código de aula).
6. Revise el resumen consolidado y confirme `SI` una sola vez.

Cada aula se aplica como un cierre independiente: respaldo único actualizado,
certificados verificados y numeración calculada sobre el archivo real. Si un aula
falla de forma crítica, el lote se detiene e informa cuáles quedaron cerradas.

---

### Opción 2: Ejecución Directa por Línea de Comandos

Si prefiere ejecutar directamente el inspector indicando las rutas:

```bash
python tests/discovery/inspect_real_files.py ^
  --control "C:\Ruta\Control_Aulas_Formadores_2026.xlsx" ^
  --sistematizacion "C:\Ruta\Sistematizacion_Cursos_COIN_2026.xlsx" ^
  --notas "C:\Ruta\Notas_ABBUEI154.xlsx" ^
  --certificados "C:\Ruta\Control_codigos_certificados_ABBUEI.xlsx" ^
  --aula "ABBUEI154"
```

---

## 🔍 Qué Hace la Fase 0 (Descubrimiento Local) y qué Informe Produce

La Fase 0 inspecciona profundamente la estructura de los 4 archivos y genera un archivo de texto en `reports/discovery_YYYYMMDD_HHMMSS.txt` que documenta:

1. **Hashes SHA-256 y Tamaños:** Para certificar la integridad de cada fuente.
2. **Hojas y Dimensiones:** Número de filas, columnas y rangos.
3. **Celdas Combinadas y Filas/Columnas Ocultas:** Para prevenir errores de lectura.
4. **Encabezados Exactos:** Posición de cada columna y su correspondencia preliminar.
5. **Detección de Módulos:** Módulos identificados mediante análisis semántico.
6. **Fórmulas Detectadas:** Especialmente la fórmula exacta utilizada en la columna `Estado para certificación`.
7. **Colores de Relleno (Fills):** Información técnica detallada (RGB, temas, tints, indexados) para identificar con exactitud el color naranja de estudiantes excluidos.
8. **Estadísticas Estructurales (Non-PII):** Conteo de registros, unicidad de documentos y correos con datos enmascarados.
9. **Búsqueda del Aula:** Localización de las celdas donde aparece el código del aula seleccionada.

---

## 📊 Matriz de Reglas: Estado Actual

| Concepto / Regla | Estado | Descripción |
|---|---|---|
| Inmutabilidad (Solo lectura + SHA-256) | `[CONFIRMADO]` | Implementado y verificado por tests unitarios. |
| Procesamiento 100% Local y Offline | `[CONFIRMADO]` | Sin conexiones de red, telemetría ni llamadas externas. |
| Matching Ultra-Conservador por Doc / Correo | `[CONFIRMADO]` | Prohibido asignar notas por coincidencia difusa de nombres. |
| Aislamiento de Errores por Registro | `[CONFIRMADO]` | Un estudiante con inconsistencia no bloquea a los demás. |
| Fórmula de Estado para Certificación | `[PENDIENTE DE ARCHIVOS REALES]` | Se inspeccionará en el informe de Fase 0. |
| Representación técnica del color naranja | `[PENDIENTE DE ARCHIVOS REALES]` | Se identificará el RGB/Theme exacto en Fase 0. |
| Relación `Sección` (Notas) $\leftrightarrow$ `Clase` (Sist) | `[PENDIENTE DE ARCHIVOS REALES]` | Se verificará la concordancia en Fase 0. |
| Columnas y formato real de Certificados | `[PENDIENTE DE ARCHIVOS REALES]` | Se auditarán los encabezados en Fase 0. |
| Fecha de envío = Fecha Final | `[REGLA OPERATIVA — A VALIDAR]` | Se comprobará contra registros históricos. |
| Total certificados = 1 | `[REGLA OPERATIVA — A VALIDAR]` | Se comprobará contra registros históricos. |
| Confirmación de Planner / Teams | `[PENDIENTE DE CONFIRMACIÓN]` | Validación manual por parte del usuario. |

---

## 🧪 Pruebas Automatizadas

El proyecto incluye fixtures sintéticos (100% ficticios, sin datos reales) y una suite de pruebas unitarias e integración:

```bash
python -m unittest discover -s tests -p "test_*.py" -v
```

---

## 🛣️ Próximos Pasos (Transición de Fase 0 a Fase 2)

1. **Paso 1 (Usuario):** Ejecutar `python main.py` (Opción 1) seleccionando sus 4 archivos reales en su computador.
2. **Paso 2 (Usuario):** Abrir el informe generado en `reports/discovery_...txt` y observar los hallazgos (fórmula exacta, RGB de color naranja, nombres de columnas).
3. **Paso 3:** Con las reglas confirmadas a partir de la evidencia del informe, se desbloquearán y construirán los parsers definitivos (Fase 2), validadores (Fase 3) y el generador del Excel auxiliar de 5 hojas (Fase 4).
