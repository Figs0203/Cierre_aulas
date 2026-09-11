"""
Normalización y formateo de fechas para el Asistente de Cierre de Aulas COIN.

Maneja la conversión entre los diferentes formatos de fecha encontrados
en los archivos Excel:
- DD/MM/AAAA (Sistematización)
- DD-mes-AAAA como "16-jun-2025" (Certificados)
- Fechas nativas de Excel (datetime)
- Fechas almacenadas como texto en diversos formatos.

Principio de diseño:
    Las fechas deben provenir de fuentes verificables.
    La herramienta normaliza FORMATO, nunca CONTENIDO.
    No se crean fechas nuevas salvo regla explícita confirmada.
"""

from __future__ import annotations

from datetime import date, datetime
from typing import Any


# Meses en español para formato de certificados (16-jun-2025)
_MESES_ES = {
    1: "ene", 2: "feb", 3: "mar", 4: "abr",
    5: "may", 6: "jun", 7: "jul", 8: "ago",
    9: "sep", 10: "oct", 11: "nov", 12: "dic",
}

# Meses en español inverso para parseo
_MESES_ES_INV = {v: k for k, v in _MESES_ES.items()}
# Agregar variantes con tilde y sin tilde
_MESES_ES_INV.update({
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4,
    "mayo": 5, "junio": 6, "julio": 7, "agosto": 8,
    "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12,
})


def parse_date(value: Any) -> date | None:
    """Intenta parsear un valor como fecha.

    Maneja:
    - datetime y date nativos.
    - Strings en formato DD/MM/AAAA.
    - Strings en formato DD-MM-AAAA.
    - Strings en formato DD-mes-AAAA (ej. "16-jun-2025").
    - Strings en formato AAAA-MM-DD (ISO).
    - Valores numéricos de Excel (serial date).

    Args:
        value: Valor a interpretar como fecha.

    Returns:
        date si el parseo fue exitoso, None si no se pudo interpretar.
        Nunca lanza excepción; devuelve None ante datos no interpretables.
    """
    if value is None:
        return None

    # Ya es una fecha nativa
    if isinstance(value, datetime):
        return value.date()

    if isinstance(value, date):
        return value

    # Intentar como string
    if isinstance(value, str):
        text = value.strip()
        if not text:
            return None
        return _parse_date_string(text)

    # Intentar como número (serial date de Excel)
    if isinstance(value, (int, float)):
        return _parse_excel_serial(value)

    return None


def _parse_date_string(text: str) -> date | None:
    """Parsea una cadena de texto como fecha."""

    # Formato DD/MM/AAAA
    parts = text.split("/")
    if len(parts) == 3:
        return _try_dmy(parts[0], parts[1], parts[2])

    # Formato ISO AAAA-MM-DD (ej. "2026-06-16"). Debe evaluarse antes del
    # formato DD-MM-AAAA para no confundir año y día.
    if len(text) == 10 and text[4] == "-" and text[7] == "-":
        try:
            return date.fromisoformat(text)
        except ValueError:
            pass

    # Formato DD-MM-AAAA o DD-mes-AAAA
    parts = text.split("-")
    if len(parts) == 3:
        # Intentar primero DD-mes-AAAA (ej. "16-jun-2025")
        month_num = _MESES_ES_INV.get(parts[1].lower().strip())
        if month_num is not None:
            return _try_dmy(parts[0], str(month_num), parts[2])
        # Intentar DD-MM-AAAA
        return _try_dmy(parts[0], parts[1], parts[2])

    return None


def _try_dmy(day_str: str, month_str: str, year_str: str) -> date | None:
    """Intenta construir una fecha a partir de día, mes y año como strings."""
    try:
        day = int(day_str.strip())
        month = int(month_str.strip())
        year = int(year_str.strip())

        # Ajustar años de 2 dígitos
        if year < 100:
            year += 2000

        return date(year, month, day)
    except (ValueError, OverflowError):
        return None


def _parse_excel_serial(serial: int | float) -> date | None:
    """Convierte un número serial de Excel a fecha.

    Excel usa el 1 de enero de 1900 como base (serial = 1).
    Nota: Excel tiene un bug histórico que considera 1900 como bisiesto.
    """
    try:
        serial_int = int(serial)
        if serial_int < 1 or serial_int > 2958465:  # Rango válido de Excel
            return None

        # Ajustar el bug de Excel (29-feb-1900 no existe)
        if serial_int > 59:
            serial_int -= 1

        # Base: 31-dic-1899 (serial 0 en el sistema corregido)
        from datetime import timedelta
        base = date(1899, 12, 31)
        return base + timedelta(days=serial_int)
    except (ValueError, OverflowError):
        return None


def format_date_sistematizacion(d: date | None) -> str:
    """Formatea una fecha para la columna de Sistematización: DD/MM/AAAA.

    Args:
        d: Fecha a formatear.

    Returns:
        String en formato "DD/MM/AAAA" o "" si la fecha es None.
    """
    if d is None:
        return ""
    return d.strftime("%d/%m/%Y")


def format_date_certificados(d: Any) -> str:
    """Formatea una fecha para el archivo de Certificados: DD-mes-AAAA.

    Ejemplo: date(2025, 6, 16) → "16-jun-2025"
    """
    if d is None:
        return ""
    if isinstance(d, str):
        d = parse_date(d)
        if d is None:
            return ""
    if isinstance(d, datetime):
        d = d.date()
    if not isinstance(d, date):
        return ""

    month_name = _MESES_ES.get(d.month, "???")
    return f"{d.day}-{month_name}-{d.year}"


def extract_year(d: Any) -> int | None:
    """Extrae el año de una fecha o cadena de fecha."""
    if d is None:
        return None
    if isinstance(d, str):
        d = parse_date(d)
    if isinstance(d, (date, datetime)):
        return d.year
    return None


# Meses de inicio de cada semestre (regla operativa confirmada por el usuario):
#   Semestre 1: 1 de noviembre  →  último día de abril
#   Semestre 2: 1 de mayo       →  31 de octubre
_MESES_SEMESTRE_1 = frozenset({11, 12, 1, 2, 3, 4})
_MESES_SEMESTRE_2 = frozenset({5, 6, 7, 8, 9, 10})


def extract_semester(d: Any) -> int | None:
    """Determina el semestre (1 o 2) a partir de una fecha o cadena de fecha.

    Regla operativa oficial (confirmada):
        - Semestre 1: del 1 de noviembre al último día de abril.
        - Semestre 2: del 1 de mayo al 31 de octubre.

    Args:
        d: Fecha (date/datetime), cadena de fecha o serial de Excel.

    Returns:
        1, 2, o None si la fecha no se puede interpretar.
    """
    if d is None:
        return None
    if isinstance(d, str):
        d = parse_date(d)
    if isinstance(d, datetime):
        d = d.date()
    if isinstance(d, date):
        if d.month in _MESES_SEMESTRE_1:
            return 1
        if d.month in _MESES_SEMESTRE_2:
            return 2
    return None


def format_semester_label(d: Any) -> str:
    """Formatea el periodo académico para la columna 'Semestre' de Certificados.

    Formato oficial confirmado con datos reales del archivo de Certificados
    (ej. fecha 2022-02-07 → '2022-1'): la etiqueta es 'AAAA-S', donde AAAA es el
    año de la fecha y S es el semestre según la regla operativa (Nov–Abr = 1,
    May–Oct = 2).

    Ejemplos:
        2022-02-07 → '2022-1'
        2022-03-28 → '2022-1'
        2025-10-01 → '2025-2'
        2026-12-15 → '2026-1'

    Args:
        d: Fecha (date/datetime), cadena de fecha o serial de Excel.

    Returns:
        String 'AAAA-S', o '' si la fecha no se puede interpretar.
    """
    if d is None:
        return ""
    if isinstance(d, str):
        d = parse_date(d)
        if d is None:
            return ""
    if isinstance(d, datetime):
        d = d.date()
    if not isinstance(d, date):
        return ""
    semestre = extract_semester(d)
    if semestre is None:
        return ""
    return f"{d.year}-{semestre}"


# ============================================================
# FORMATO PARA COLUMNA "CICLO" EN CERTIFICADOS (REGLA 6 - CONFIRMADO)
# ============================================================

# Meses en español en MAYÚSCULAS para la columna Ciclo
_MESES_ES_MAYUSCULAS = {
    1: "ENERO", 2: "FEBRERO", 3: "MARZO", 4: "ABRIL",
    5: "MAYO", 6: "JUNIO", 7: "JULIO", 8: "AGOSTO",
    9: "SEPTIEMBRE", 10: "OCTUBRE", 11: "NOVIEMBRE", 12: "DICIEMBRE",
}


def format_ciclo_date(d: Any) -> str:
    """Formatea una fecha al estilo de la columna Ciclo: 'MES DÍA'.

    Formato confirmado por el monitor (2026-09-04):
        2026-03-16 → 'MARZO 16'
        2026-04-27 → 'ABRIL 27'
    """
    if d is None:
        return ""
    if isinstance(d, str):
        d = parse_date(d)
        if d is None:
            return ""
    if isinstance(d, datetime):
        d = d.date()
    if not isinstance(d, date):
        return ""
    mes = _MESES_ES_MAYUSCULAS.get(d.month, "")
    if not mes:
        return ""
    return f"{mes} {d.day}"


# Alias de compatibilidad
format_certificate_date = format_date_certificados
get_semester = extract_semester
get_semester_label = format_semester_label
