"""
Configuración general del Asistente de Cierre de Aulas COIN.

Contiene constantes, formatos y parámetros configurables del sistema.
Los valores definidos aquí son los predeterminados; algunos se pueden
ajustar tras la Fase 0 según los hallazgos de los archivos reales.
"""

from __future__ import annotations

from pathlib import Path


# ============================================================
# FORMATO DEL ARCHIVO DE SALIDA
# ============================================================

# Prefijo del nombre del archivo auxiliar generado
OUTPUT_FILE_PREFIX = "Cierre_Aula"

# Extensión del archivo de salida
OUTPUT_FILE_EXTENSION = ".xlsx"

# Nombres de las hojas del archivo de salida
SHEET_NAMES = {
    "resumen": "1. RESUMEN",
    "sistematizacion": "2. SISTEMATIZACION",
    "certificados": "3. CERTIFICADOS",
    "control_update": "4. ACTUALIZACION_CONTROL",
    "validaciones": "5. VALIDACIONES_Y_TRAZABILIDAD",
}


# ============================================================
# FORMATOS DE FECHA
# ============================================================

# Formato de fecha para la columna de Sistematización
DATE_FORMAT_SISTEMATIZACION = "%d/%m/%Y"  # DD/MM/AAAA

# Formato de fecha para certificados se maneja en normalization/dates.py
# porque usa nombres de meses en español (16-jun-2025)


# ============================================================
# VALORES OPERATIVOS
# ============================================================
# NOTA: Los siguientes valores tienen estado
# [REGLA OPERATIVA REPORTADA — PENDIENTE DE VALIDACIÓN]
# y se confirmarán o ajustarán en la Fase 0.

# Valor para "Se elaboró certificado" cuando el estudiante NO aprobó
CERTIFICADO_NO = "NO"

# Valor para "Se elaboró certificado" cuando el estudiante SÍ aprobó
CERTIFICADO_SI = "SI"

# Valor para campos no aplicables en estudiantes no aprobados
# PENDIENTE DE VALIDACIÓN: confirmar si "No aplica" es el texto oficial
VALOR_NO_APLICA = "No aplica"

# Total de certificados por estudiante aprobado
# PENDIENTE DE VALIDACIÓN: confirmar que siempre es 1
TOTAL_CERTIFICADOS_DEFAULT = 1


# ============================================================
# DETECCIÓN DE ARCHIVOS
# ============================================================

# Extensiones válidas para archivos de entrada
VALID_EXTENSIONS = {".xlsx"}

# Prefijo de archivos auxiliares generados (para evitar seleccionarlos como fuente)
AUXILIARY_FILE_PREFIX = "Cierre_Aula_"


# ============================================================
# LOGGING Y PRIVACIDAD
# ============================================================

# Mostrar datos enmascarados en consola
MASK_PII_IN_CONSOLE = True

# Número de caracteres visibles en documentos enmascarados
DOCUMENT_VISIBLE_CHARS = 4

# Separador visual entre clases/grupos en el reporte
CLASS_SEPARATOR = "=" * 60
