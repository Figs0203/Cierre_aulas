"""
Configuración general del Asistente de Cierre de Aulas COIN.

Contiene constantes, formatos y parámetros configurables del sistema.
Los valores definidos aquí son los predeterminados; algunos se pueden
ajustar tras la Fase 0 según los hallazgos de los archivos reales.
"""

from __future__ import annotations

from pathlib import Path


# ============================================================
# HOJAS DEL ARCHIVO CONTROL_AULAS_FORMADORES (REGLA 1 - CONFIRMADO)
# ============================================================

# Nombres exactos de las hojas del archivo Control_Aulas_Formadores_2026.
# La hoja 'Creación aulas GDA' debe ignorarse SIEMPRE.
CONTROL_AULAS_ALL_SHEETS = [
    "Tabla_CMI",
    "Tabla_ABBUEI",
    "Tabla_GICE",
    "Tabla_EPA",
    "Tabla_LATEX",
    "Tabla_IPI",
    "Tabla_UEI",
    "Tabla_CAI",
    "CONTROL",
]

# Hoja que debe ignorarse siempre, sin excepción.
CONTROL_AULAS_IGNORED_SHEET = "Creación aulas GDA"


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
# NOTA DE ESTADO POR CAMPO:
#   [CONFIRMADO]   = Regla validada por el usuario.
#   [PENDIENTE]    = Requiere evidencia de archivos reales.

# [CONFIRMADO] Valor para "Se elaboró certificado" cuando el estudiante NO aprobó
CERTIFICADO_NO = "NO"

# [CONFIRMADO] Valor para "Se elaboró certificado" cuando el estudiante SÍ aprobó
CERTIFICADO_SI = "SI"

# [CONFIRMADO] Valor en "Código del certificado" para estudiantes no aprobados.
# Se aplica en SISTEMATIZACIÓN y en CERTIFICADOS.
VALOR_NO_APLICA = "No aplica"

# [PENDIENTE] Total de certificados por estudiante aprobado — confirmar en Fase 0.
TOTAL_CERTIFICADOS_DEFAULT = 1


# ============================================================
# FORMATO DE COLUMNAS DE CERTIFICADOS (REGLAS 5 Y 6 - CONFIRMADO)
# ============================================================

# [CONFIRMADO] Formato de la columna Ciclo en Certificados:
# "CLASE {clase} {catalogo} - {programa} - {mes_inicio} {dia_inicio} A {mes_fin} {dia_fin}"
# Ejemplo: "CLASE 5535 MF7001 - M. ESTUDIOS JURÍDICOS - MARZO 16 A ABRIL 27"
CICLO_FORMAT = "CLASE {clase} {catalogo} - {programa} - {mes_inicio} {dia_inicio} A {mes_fin} {dia_fin}"

# [CONFIRMADO] Formato de la columna Docente COIN en Certificados:
# "{codigo_guion} {nombre_formador}"
# Ejemplo: "ABBUEI-205 Manuela Restrepo"
DOCENTE_COIN_FORMAT = "{codigo_guion} {nombre_formador}"

# Meses en español en mayúsculas para el formato de Ciclo
MESES_ES_MAYUSCULAS = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
    5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
    9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE",
}


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
