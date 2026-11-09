"""
Parser especializado para el archivo de Notas descargado de Interactiva.

Extrae estudiantes, notas por módulo, nota final, identificadores
(OrgDefinedId, Username, Nombres) y la información de la columna Sección,
además de detectar colores de relleno (estudiantes excluidos en naranja).

Encabezados confirmados en Fase 0:
  Col  1 (A): OrgDefinedId
  Col  2 (B): Username
  Col  3 (C): Last Name
  Col  4 (D): First Name
  Col  5 (E): Sección (ej. 'EG7003-6872 COMB FI0')
  Col  6 (F): Cuestionario módulo 1
  Col  7 (G): Buzón módulo 2
  Col  8 (H): Cuestionario módulo 3
  Col  9 (I): Cuestionario módulo 4
  Col 10 (J): Calculated Final

Si los encabezados no permiten reconocer los módulos o la nota final, se usa el
**respaldo posicional** `_infer_grade_columns`: de izquierda a derecha, módulo
1..N y la última columna numérica es la nota final. Solo se activa cuando falta
el mapeo por encabezado y valida cada columna con sus datos (sin identidad,
fechas, porcentajes ni estados).
"""

from __future__ import annotations

import math
import re
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import openpyxl
from openpyxl.utils import get_column_letter

from config.colors import (
    FILL_EXCLUDED_ORANGE,
    is_orange_fill_exact,
)
from core.models import EstudianteNotas
from core.traceability import TracedValue
from normalization.headers import (
    final_grade_candidates,
    find_final_grade_column,
    find_module_columns,
)
from normalization.text import (
    normalize_document,
    normalize_for_comparison,
    normalize_whitespace,
)


def _parse_nota(raw: Any) -> Optional[float]:
    """Convierte un valor crudo de celda de notas a float.

    El archivo oficial de Interactiva puede usar tanto punto como coma como
    separador decimal (ej. '3,5' o '3.5' o 3.5). Esta función normaliza
    ambas representaciones para el procesamiento interno. Los archivos de
    salida (auxiliar, sistematización) usarán coma como separador decimal
    si los valores se formatean como texto, o el tipo float nativo de Excel
    si se escriben como número.

    Args:
        raw: Valor de celda (str, int, float o None).

    Returns:
        float si se pudo convertir, None si el valor es vacío o inválido.
    """
    if raw is None:
        return None
    if isinstance(raw, (int, float)):
        return float(raw)
    s = str(raw).strip()
    if not s:
        return None
    # Soportar ambos separadores decimales: punto y coma
    # Primero reemplazar coma decimal por punto (para convertir a float)
    # Caso: '3,5' -> '3.5'; '3.5' -> '3.5'
    # Cuidado: si hay separador de miles con punto (ej. '1.234,56') se normaliza correctamente
    if ',' in s and '.' in s:
        # Formato con miles: '1.234,56' → quitar puntos de miles, sustituir coma decimal
        s = s.replace('.', '').replace(',', '.')
    elif ',' in s:
        # Solo coma: '3,5' → '3.5'
        s = s.replace(',', '.')
    # Si solo tiene punto ya está en formato correcto
    try:
        return float(s)
    except ValueError:
        return None


def parse_seccion_info(seccion_str: str | None) -> Tuple[str | None, str | None]:
    """Extrae (Catálogo, Clase) a partir del texto de la columna Sección.

    Ejemplos:
        'EG7003-6872 COMB FI0' → ('EG7003', '6872')
        'ABBUEI154-001'        → ('ABBUEI154', '001')
        'CAT101-002'           → ('CAT101', '002')
        '001'                  → (None, '001')

    Args:
        seccion_str: Texto en la columna Sección.

    Returns:
        Tupla (catalogo, clase). Los elementos pueden ser None si no se encuentran.
    """
    if not seccion_str:
        return None, None

    clean = normalize_whitespace(str(seccion_str)).strip()

    # Patrón 1: CATALOGO-CLASE (ej. EG7003-6872 COMB FI0 o ABBUEI154-001)
    m = re.search(r"([A-Za-z0-9]+)[-_ ]+(\d+)", clean)
    if m:
        return m.group(1).upper(), m.group(2)

    # Patrón 2: Solo dígitos de clase (ej. "001" o "6872")
    m_num = re.search(r"(\d+)", clean)
    if m_num:
        return None, m_num.group(1)

    return None, clean


def _is_document_header(h_lower: str) -> bool:
    """Indica si un encabezado corresponde a la columna del número de documento.

    Algunos cursos exportan la identidad como 'Documento' en lugar de
    'OrgDefinedId'. Se excluye 'tipo de documento'.
    """
    hl = str(h_lower).strip()
    if "tipo de documento" in hl:
        return False
    return (
        "orgdefinedid" in hl
        or "org defined id" in hl
        or hl in (
            "documento", "documento de identidad", "identificación", "identificacion",
            "número de documento", "numero de documento", "nro documento", "n. documento",
        )
    )


def _count_identity_headers(headers: List[Optional[str]]) -> int:
    """Cuenta cuántos encabezados corresponden a columnas de identidad del estudiante."""
    count = 0
    for h in headers:
        if not h:
            continue
        hl = str(h).lower()
        if (
            _is_document_header(hl) or "username" in hl
            or "last name" in hl or "first name" in hl
            or "apellidos" in hl or "nombres" in hl
        ):
            count += 1
    return count


# ---------------------------------------------------------------------------
# Respaldo POSICIONAL cuando los encabezados de Notas no se reconocen
# ---------------------------------------------------------------------------
# Regla indicada por la biblioteca: en los archivos de Interactiva las columnas
# con notas van, de izquierda a derecha, del módulo 1 al módulo N y la última es
# la nota final. Este respaldo SOLO se activa cuando la detección por encabezados
# no encuentra módulos o nota final, y valida cada columna con sus DATOS para no
# confundir identidad, fechas, porcentajes, estados ni códigos con notas.

_GRADE_SYMBOLS = frozenset({
    "aprobado", "aprobada", "aprobo", "aprobó",
    "no aprobado", "no aprobada", "no aprobo", "no aprobó",
    "reprobado", "reprobada", "reprobo", "reprobó",
    "no aplica", "n/a", "na", "pendiente", "cancelado", "cancelada",
    "exento", "exenta", "ganado", "ganada", "perdido", "perdida",
    "satisfactorio", "no satisfactorio",
})

# Encabezados que NUNCA son una columna de notas en este respaldo.
_FALLBACK_NON_GRADE_RE = re.compile(
    r"\bfecha\b|\bhora\b|nombre|apellido|correo|email|documento|identificaci|"
    r"orgdefinedid|org defined id|usuario|username|seccion|programa|catalogo|"
    r"clase|aula|semestre|periodo|\bano\b|\bmes\b|porcentaje|peso|intento|"
    r"estado|estatus|observaci|comentario|comentarios|codigo|cedula|telefono|"
    r"celular|sede|jornada|grupo|credito|matricul|identificador|"
    r"\baprob|\breprob|\bganad|\bperdid|\bcertific",
    re.IGNORECASE,
)

_MAX_MODULOS_OFICIALES = 5


def _header_is_non_grade(header: Optional[str]) -> bool:
    """True si el encabezado describe una columna que no contiene notas."""
    if not header or not str(header).strip():
        return False
    return bool(_FALLBACK_NON_GRADE_RE.search(normalize_for_comparison(str(header))))


def _rightmost_cluster(columns: List[int], max_gap: int = 1) -> List[int]:
    """Devuelve el grupo de columnas más a la derecha, tolerando huecos pequeños.

    Ej.: [6, 7, 8, 12, 14] con ``max_gap=1`` empieza en 14, tolera el hueco de
    la 13 y añade 12; el hueco hasta 8 es mayor, así que se detiene → [12, 14].
    """
    if not columns:
        return []
    grupo = [columns[-1]]
    for c in reversed(columns[:-1]):
        if grupo[0] - c - 1 <= max_gap:
            grupo.insert(0, c)
        else:
            break
    return grupo


def _datos_columna_notas(
    ws,
    col: int,
    filas: List[int],
    valor_efectivo,
) -> Dict[str, Any]:
    """Resume si una columna parece contener notas, usando solo sus datos."""
    no_vacias = numericos = fechas = booleanos = porcentajes = 0
    valores_num: List[float] = []
    valores_texto: List[str] = []
    for r in filas:
        celda = ws.cell(r, col)
        v = valor_efectivo(celda, r, col)
        if v is None or str(v).strip() == "":
            continue
        no_vacias += 1
        if isinstance(v, bool):
            booleanos += 1
            continue
        if isinstance(v, (date, datetime)):
            fechas += 1
            continue
        if "%" in str(getattr(celda, "number_format", "") or ""):
            porcentajes += 1
        n = _parse_nota(v)
        if n is not None:
            numericos += 1
            valores_num.append(n)
        else:
            valores_texto.append(str(v).strip())
    return {
        "no_vacias": no_vacias,
        "numericos": numericos,
        "fechas": fechas,
        "booleanos": booleanos,
        "porcentajes": porcentajes,
        "valores_num": valores_num,
        "valores_texto": valores_texto,
    }


def _cap_modulos(modulos: List[int], avisos: List[str]) -> List[int]:
    """Limita los módulos a los 5 oficiales, informando de lo descartado."""
    if len(modulos) <= _MAX_MODULOS_OFICIALES:
        return modulos
    avisos.append(
        "Se encontraron más de 5 columnas con notas antes de la nota final; se "
        "usarán las 5 primeras como módulos 1–5 y se ignorarán "
        + ", ".join(get_column_letter(c) for c in modulos[_MAX_MODULOS_OFICIALES:])
        + ". Verifique los encabezados del archivo de Notas."
    )
    return modulos[:_MAX_MODULOS_OFICIALES]


def _infer_grade_columns(
    ws,
    headers: List[Optional[str]],
    valor_efectivo,
    max_col: int,
    max_row: int,
    identity_cols: set[int],
    known_mod_cols: Dict[int, int],
    known_final_col: Optional[int],
) -> Tuple[Dict[int, int], Optional[int], List[str], Dict[str, Any]]:
    """Infiere por posición las columnas de módulos y nota final.

    Se usa **solo** cuando la detección por encabezados no encontró módulos o
    nota final. Supone que, de izquierda a derecha, las columnas con notas son
    módulo 1, módulo 2, …, módulo N y, la última, la nota final.

    Para no confundir identidad, fechas, porcentajes, estados o códigos con
    notas, cada columna candidata se valida con los **datos**: mayoría de valores
    numéricos, sin booleanos, sin fechas, sin formato de porcentaje y con valores
    en un rango plausible de nota. Las columnas de identidad y las ya detectadas
    por encabezado quedan excluidas.

    Returns:
        ``(mod_map, final_col, avisos, info)``. ``mod_map`` es
        ``{num_modulo: columna}``. Si había módulos detectados por encabezado,
        devuelve el mapa completo (los reconocidos + los completados por
        posición) para no reemplazarlos; el llamador decide si hubo cambios.
    """
    avisos: List[str] = []
    info: Dict[str, Any] = {"filas_analizadas": 0}

    # 1. Filas de estudiantes: al menos un dato de identidad no vacío.
    filas: List[int] = []
    for r in range(2, min(max_row, 1000) + 1):
        if any(
            ws.cell(r, c).value not in (None, "")
            for c in identity_cols
            if c and c <= max_col
        ):
            filas.append(r)
    if not filas:
        # Respaldo: filas con algo en las primeras 5 columnas.
        for r in range(2, min(max_row, 1000) + 1):
            if any(
                ws.cell(r, c).value not in (None, "")
                for c in range(1, min(max_col, 5) + 1)
            ):
                filas.append(r)
    info["filas_analizadas"] = len(filas)
    if not filas:
        avisos.append(
            "No se pudieron identificar filas de estudiantes para inferir las "
            "columnas de notas."
        )
        return {}, None, avisos, info

    total = len(filas)
    minimo = max(1, math.ceil(total * 0.1))
    excluidas = set(known_mod_cols.values()) if known_mod_cols else set()
    if known_final_col:
        excluidas.add(known_final_col)

    numericas: List[int] = []
    simbolicas: List[int] = []

    for c in range(1, max_col + 1):
        if c in identity_cols or c in excluidas:
            continue
        if _header_is_non_grade(headers[c - 1] if c - 1 < len(headers) else None):
            continue

        datos = _datos_columna_notas(ws, c, filas, valor_efectivo)
        if datos["no_vacias"] < minimo:
            continue

        ratio_num = datos["numericos"] / max(datos["no_vacias"], 1)
        rango_ok = False
        if datos["valores_num"]:
            ordenados = sorted(datos["valores_num"])
            mediana = ordenados[len(ordenados) // 2]
            # Descarta documentos/códigos: las notas no suelen pasar de 1000
            # ni tener una mediana de cientos/miles.
            rango_ok = max(datos["valores_num"]) <= 1000 and mediana <= 100

        es_numerica = (
            datos["booleanos"] == 0
            and datos["fechas"] == 0
            and datos["porcentajes"] == 0
            and datos["numericos"] >= minimo
            and ratio_num >= 0.6
            and rango_ok
        )
        textos_ok = bool(datos["valores_texto"]) and all(
            t.lower() in _GRADE_SYMBOLS for t in datos["valores_texto"]
        )
        es_simbolica = (
            datos["numericos"] == 0
            and textos_ok
            and datos["no_vacias"] >= math.ceil(total * 0.5)
        )
        if es_numerica:
            numericas.append(c)
        elif es_simbolica:
            simbolicas.append(c)

    # 2. Si la nota final ya se detectó por encabezado, solo faltan los módulos.
    if known_final_col and not known_mod_cols:
        modulos = _rightmost_cluster(
            [c for c in numericas if c < known_final_col], max_gap=1
        )
        modulos = _cap_modulos(modulos, avisos)
        if modulos:
            avisos.append(
                "No se reconocieron los encabezados de los módulos; se infirieron "
                "por posición antes de la nota final detectada: "
                + ", ".join(
                    f"{i}={get_column_letter(c)}" for i, c in enumerate(modulos, 1)
                )
                + "."
            )
        return {i: c for i, c in enumerate(modulos, 1)}, known_final_col, avisos, info

    # 3. Si los módulos ya se detectaron por encabezado, se infiere la nota final
    #    y, si la numeración es un prefijo 1..k, se completan por posición los
    #    módulos que falten (sin reemplazar los ya reconocidos).
    if known_mod_cols:
        modulos = dict(known_mod_cols)
        numeros = sorted(modulos)
        prefijo = numeros == list(range(1, len(numeros) + 1))
        ultimo_mod = max(modulos.values())
        candidatas = [c for c in numericas if c > ultimo_mod]

        final = known_final_col
        if final:
            candidatas = [c for c in candidatas if c < final]
            grupo = _rightmost_cluster(candidatas, max_gap=1)
        else:
            grupo = _rightmost_cluster(candidatas, max_gap=1)
            if grupo:
                final = grupo[-1]
                grupo = grupo[:-1]
            else:
                final = None

        if prefijo and grupo:
            completos = [modulos[i] for i in numeros] + grupo
            completos = _cap_modulos(completos, avisos)
            modulos = {i: c for i, c in enumerate(completos, 1)}
            avisos.append(
                "Se completaron módulos por posición después del último encabezado "
                "reconocido: "
                + ", ".join(f"{i}={get_column_letter(c)}" for i, c in modulos.items())
                + "."
            )
        elif grupo:
            avisos.append(
                "Hay columnas numéricas no asignadas entre los módulos y la nota "
                "final, pero la numeración de módulos detectada no es consecutiva "
                "desde 1; no se infirieron. Revise los encabezados."
            )

        if not known_final_col and final:
            avisos.append(
                "No se reconoció el encabezado de la nota final; se tomó la última "
                f"columna numérica después de los módulos: {get_column_letter(final)}."
            )
        elif not final:
            avisos.append(
                "Se reconocieron los módulos por encabezado, pero no se pudo "
                "determinar la columna de la nota final."
            )
        return modulos, final, avisos, info

    # 4. Sin encabezados reconocidos: módulos + nota final por posición.
    if not numericas:
        if simbolicas:
            avisos.append(
                "No se encontraron notas numéricas; no se asignó nota final para "
                "no confundir una columna simbólica (p. ej. 'Aprobado'/'No aplica') "
                "con un estado. Ajuste el encabezado de la nota final para que la "
                "aplicación lo reconozca."
            )
        else:
            avisos.append(
                "No se encontraron columnas con notas numéricas para inferir los "
                "módulos por posición."
            )
        return {}, None, avisos, info

    grupo = _rightmost_cluster(numericas, max_gap=1)
    final_col = grupo[-1]
    modulos = list(grupo)
    if final_col in modulos:
        modulos = [c for c in modulos if c != final_col]

    # Una columna simbólica después de las notas puede ser un estado o una nota
    # final simbólica. Por seguridad NO se usa como nota final: se avisa para
    # que el usuario revise el encabezado.
    simbolicas_vecinas = [c for c in simbolicas if final_col < c <= final_col + 3]
    if simbolicas_vecinas:
        avisos.append(
            "Se detectó una columna simbólica después de las notas ("
            + ", ".join(get_column_letter(c) for c in simbolicas_vecinas)
            + "); no se usó como nota final para no confundirla con un estado. "
            "Si es la nota final, ajuste su encabezado para que la aplicación lo "
            "reconozca."
        )

    modulos = _cap_modulos(modulos, avisos)
    if not modulos:
        avisos.append(
            "Solo se encontró una columna con notas; se tomó como nota final y no "
            "se asignaron módulos."
        )
    else:
        avisos.append(
            "No se reconocieron los encabezados de módulos/nota final; se "
            "infirieron por posición. Verifique el mapeo mostrado: módulos "
            + ", ".join(f"{i}={get_column_letter(c)}" for i, c in enumerate(modulos, 1))
            + f"; nota final={get_column_letter(final_col)}."
        )
        avisos.append(
            "La inferencia por posición supone que todos los módulos tienen notas. "
            "Si algún módulo está completamente vacío, la numeración puede "
            "desplazarse."
        )

    return {i: c for i, c in enumerate(modulos, start=1)}, final_col, avisos, info


def _count_numeric_grade_columns(ws, limite: int = 60) -> int:
    """Cuenta columnas con notas numéricas (para elegir la hoja correcta).

    Se usa solo cuando la detección por encabezados no encontró módulos ni nota
    final: permite reconocer una hoja con notas aunque sus títulos sean atípicos.
    """
    max_col = ws.max_column or 1
    max_row = min(ws.max_row or 1, limite)
    filas = list(range(2, max_row + 1))
    if not filas:
        return 0
    total = len(filas)
    minimo = max(1, min(3, total))
    encontradas = 0
    for c in range(1, max_col + 1):
        if _header_is_non_grade(ws.cell(1, c).value):
            continue
        datos = _datos_columna_notas(
            ws, c, filas, lambda celda, _fila, _col: celda.value
        )
        if (
            datos["booleanos"] == 0
            and datos["fechas"] == 0
            and datos["porcentajes"] == 0
            and datos["numericos"] >= minimo
            and datos["numericos"] / max(datos["no_vacias"], 1) >= 0.6
        ):
            valores = datos["valores_num"]
            if valores and max(valores) <= 1000 and sorted(valores)[len(valores) // 2] <= 100:
                encontradas += 1
    return encontradas


def _select_notas_sheet(wb) -> str:
    """Elige la hoja de Notas más adecuada del libro.

    Algunos archivos traen varias hojas: una con el listado general de la
    plataforma (solo identidad, sin calificaciones) y otra con las
    calificaciones reales del aula (p. ej. 'CURSO REALIZADO'). Se elige la hoja
    que tenga columnas de módulos/nota final Y columnas de identidad, en lugar
    de asumir por el nombre.

    Criterio (en orden):
        1. Más columnas de módulos/nota final detectadas.
        2. Más columnas de identidad.
        3. Preferencia por los nombres habituales ('Calificaciones'/'Grades').
        4. Más filas.

    Args:
        wb: Workbook de Notas cargado.

    Returns:
        Nombre de la hoja seleccionada.
    """
    candidates: List[Tuple[int, int, int, int, str]] = []
    for name in wb.sheetnames:
        ws = wb[name]
        if (ws.max_row or 1) < 2:
            continue
        max_col = ws.max_column or 1
        headers = [ws.cell(1, c).value for c in range(1, max_col + 1)]
        grade_cols = len(find_module_columns(headers)) + (1 if find_final_grade_column(headers) else 0)
        if grade_cols == 0:
            # Encabezados atípicos: la hoja aún puede tener notas numéricas.
            grade_cols = _count_numeric_grade_columns(ws)
        ident_cols = _count_identity_headers(headers)
        if grade_cols >= 1 and ident_cols >= 1:
            prefer = 1 if name in ("Calificaciones", "Grades") else 0
            candidates.append((grade_cols, ident_cols, prefer, ws.max_row or 0, name))

    if candidates:
        candidates.sort(reverse=True)
        return candidates[0][4]

    # Respaldo: nombres habituales o la primera hoja
    if "Calificaciones" in wb.sheetnames:
        return "Calificaciones"
    if "Grades" in wb.sheetnames:
        return "Grades"
    return wb.sheetnames[0]


def _elegir_columna_nota_final(
    ws,
    headers: List[Optional[str]],
    valor_efectivo=None,
    limite: int = 300,
) -> Optional[int]:
    """Elige la columna que realmente contiene la nota final.

    Puede haber varias columnas candidatas por su nombre ('Calculated Final
    Grade', 'Calculated Final Grade Scheme Symbol', 'Final', etc.). En lugar de
    decidir solo por el encabezado, se inspeccionan los **datos**: gana la
    columna con más valores numéricos, luego la de más valores no vacíos y, en
    último término, la de mayor puntaje de coincidencia del encabezado.

    Así, si en un archivo la columna real de notas se llama 'Calculated Final
    Grade Scheme Symbol' (aunque el nombre sugiera símbolos), se usa esa y no se
    descarta.

    Args:
        ws: Hoja de openpyxl con las notas.
        headers: Encabezados de la hoja.
        valor_efectivo: Función opcional (celda, fila, col) → valor, para leer el
            resultado de fórmulas. Si es None se usa ``celda.value``.
        limite: Filas a inspeccionar para decidir (las suficientes para detectar
            la columna con datos).

    Returns:
        Índice 1-based de la columna elegida, o None si no hay candidatas.
    """
    candidatos = final_grade_candidates(headers)
    if not candidatos:
        return None
    if len(candidatos) == 1:
        return candidatos[0][0]

    leer = valor_efectivo or (lambda celda, fila, col: celda.value)
    max_fila = min(ws.max_row or 1, limite)
    mejor_clave = None
    mejor_col = candidatos[0][0]

    for col, puntaje, _canonical in candidatos:
        numericos = 0
        no_vacios = 0
        for fila in range(2, max_fila + 1):
            valor = leer(ws.cell(fila, col), fila, col)
            if valor is None or str(valor).strip() == "":
                continue
            no_vacios += 1
            if _parse_nota(valor) is not None:
                numericos += 1
        clave = (numericos, no_vacios, puntaje, -col)
        if mejor_clave is None or clave > mejor_clave:
            mejor_clave = clave
            mejor_col = col

    return mejor_col


def parse_notas_archivo(
    file_path: Path | str,
    sheet_name: Optional[str] = None,
    diagnostics: Optional[Dict[str, Any]] = None,
) -> Tuple[List[EstudianteNotas], Dict[str, Dict[str, str]]]:
    """Extrae las notas de estudiantes desde el archivo de Interactiva.

    Args:
        file_path: Ruta al archivo Excel de Notas.
        sheet_name: Nombre de la hoja (si es None, usa 'Calificaciones' o la activa).
        diagnostics: Diccionario opcional que se rellena con la detección de
            columnas (hoja, método, módulos, nota final y avisos). Sirve para que
            la aplicación muestre qué columnas se usaron y advierta cuando el
            mapeo se infirió por posición.

    Returns:
        Tupla de:
          - Lista de EstudianteNotas
          - Diccionario de secciones encontradas: {seccion_raw: {"catalogo": ..., "clase": ...}}
    """
    path = Path(file_path)
    file_name = path.name

    wb = openpyxl.load_workbook(path, data_only=False, read_only=False)

    if sheet_name and sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
    else:
        # Selección automática de la hoja con calificaciones reales
        sheet_name = _select_notas_sheet(wb)
        ws = wb[sheet_name]

    # Segunda lectura con VALORES CALCULADOS. Los archivos de Interactiva pueden
    # traer la nota final (o los módulos) como fórmula; con `data_only=False`
    # openpyxl devuelve el texto de la fórmula y se perdería el valor real.
    wb_valores = None
    ws_valores = None
    try:
        wb_valores = openpyxl.load_workbook(path, data_only=True, read_only=False)
        if sheet_name in wb_valores.sheetnames:
            ws_valores = wb_valores[sheet_name]
    except Exception:  # noqa: BLE001 - si falla, se usa la lectura normal
        wb_valores = None
        ws_valores = None

    def valor_efectivo(celda, fila: int, col: int):
        """Valor real de una celda: si es fórmula, su resultado en caché si existe."""
        valor = celda.value
        if isinstance(valor, str) and valor.startswith("="):
            if ws_valores is not None:
                cacheado = ws_valores.cell(fila, col).value
                if cacheado is not None:
                    return cacheado
        return valor

    # 1. Leer encabezados
    max_col = ws.max_column or 1
    raw_headers: List[Optional[str]] = []
    for c in range(1, max_col + 1):
        v = ws.cell(1, c).value
        raw_headers.append(str(v).strip() if v is not None else None)

    max_row = ws.max_row or 1

    # 2. Columnas de identificación básica (se detectan ANTES del respaldo
    #    posicional para poder excluirlas con certeza).
    col_org_id = 1
    col_username = 2
    col_last_name = 3
    col_first_name = 4
    # La columna 'Sección' es OPCIONAL: algunos cursos (p. ej. CMI) exportan las
    # notas sin ella. Si no se detecta por encabezado, se deja en None para no
    # confundir una columna de calificaciones con la sección.
    col_seccion: Optional[int] = None

    for c_idx, h in enumerate(raw_headers, start=1):
        if not h:
            continue
        h_lower = h.lower()
        if _is_document_header(h_lower):
            col_org_id = c_idx
        elif "username" in h_lower:
            col_username = c_idx
        elif "last name" in h_lower or "apellidos" in h_lower:
            col_last_name = c_idx
        elif "first name" in h_lower or "nombres" in h_lower:
            col_first_name = c_idx
        elif "sección" in h_lower or "seccion" in h_lower:
            col_seccion = c_idx

    # 3. Mapeo semántico de columnas de módulos y final (por encabezados).
    mod_cols = find_module_columns(raw_headers)
    final_col = _elegir_columna_nota_final(
        ws, raw_headers, valor_efectivo=valor_efectivo
    )
    mod_por_encabezado = bool(mod_cols)
    final_por_encabezado = bool(final_col)
    metodo_modulos = "encabezados" if mod_por_encabezado else "no_detectado"
    metodo_final = "encabezados" if final_por_encabezado else "no_detectada"
    avisos_notas: List[str] = []
    info_inferencia: Dict[str, Any] = {}

    # 4. Respaldo POSICIONAL: solo si faltan módulos o nota final. Así los
    #    archivos con títulos atípicos (p. ej. 'Puntaje A', 'Nota 1', columnas
    #    sin título) siguen funcionando sin inventar columnas cuando los
    #    encabezados sí se reconocen.
    if not mod_cols or not final_col:
        identity_cols = {
            int(c)
            for c in (col_org_id, col_username, col_last_name, col_first_name, col_seccion)
            if c
        }
        inf_mods, inf_final, avisos, info_inferencia = _infer_grade_columns(
            ws=ws,
            headers=raw_headers,
            valor_efectivo=valor_efectivo,
            max_col=max_col,
            max_row=max_row,
            identity_cols=identity_cols,
            known_mod_cols=mod_cols,
            known_final_col=final_col,
        )
        if inf_mods:
            if not mod_cols:
                mod_cols = inf_mods
                metodo_modulos = "posicional"
            elif inf_mods != mod_cols:
                mod_cols = inf_mods
                metodo_modulos = "mixto"
        if not final_col and inf_final:
            final_col = inf_final
            metodo_final = "posicional"
        avisos_notas.extend(avisos)

    if not final_col:
        avisos_notas.append(
            "No se pudo determinar la columna de la nota final en el archivo de Notas."
        )
    if not mod_cols:
        avisos_notas.append(
            "No se pudieron determinar las columnas de los módulos en el archivo de Notas."
        )

    # 5. Diagnóstico para la aplicación (qué columnas se usaron y por qué).
    if diagnostics is not None:
        def _col_info(col: Optional[int]) -> Optional[Dict[str, Any]]:
            if not col:
                return None
            idx = int(col)
            encabezado = raw_headers[idx - 1] if 0 < idx <= len(raw_headers) else None
            return {
                "columna": idx,
                "letra": get_column_letter(idx),
                "encabezado": encabezado or "",
            }

        diagnostics.clear()
        diagnostics.update({
            "hoja": sheet_name,
            "metodo_modulos": metodo_modulos,
            "metodo_nota_final": metodo_final,
            "modulos": {
                str(num): _col_info(col) for num, col in sorted(mod_cols.items())
            },
            "nota_final": _col_info(final_col),
            "avisos": list(dict.fromkeys(avisos_notas)),
            "filas_analizadas": info_inferencia.get("filas_analizadas"),
        })

    # Columnas donde tiene sentido buscar el relleno naranja de "estudiante excluido".
    # El color de exclusión marca al ESTUDIANTE (sus columnas de identidad/nombre) y,
    # en algunos formatos, también las de calificaciones. Se EXCLUYEN columnas extra
    # como la de semáforo o columnas en blanco, que suelen tener rellenos de color
    # (ámbar/dorado) y provocarían falsos positivos de exclusión.
    student_cols: set[int] = {col_org_id, col_username, col_last_name, col_first_name}
    if col_seccion:
        student_cols.add(col_seccion)
    student_cols.update(int(c) for c in mod_cols.values() if c)
    if final_col:
        student_cols.add(int(final_col))
    student_cols = {c for c in student_cols if c and c >= 1}

    estudiantes: List[EstudianteNotas] = []
    secciones_map: Dict[str, Dict[str, str]] = {}

    for r in range(2, max_row + 1):
        # Si la fila está completamente vacía, saltarla
        vals = [ws.cell(r, c).value for c in range(1, min(max_col + 1, 15))]
        if all(v is None for v in vals):
            continue

        # Detección de fila oculta/filtrada (se procesa igualmente)
        is_hidden = bool(ws.row_dimensions[r].hidden)

        def make_traced(val: Any, col_name: str, c_idx: int) -> TracedValue:
            return TracedValue(
                value=val,
                source_file=file_name,
                source_sheet=sheet_name,
                source_column=col_name,
                source_col_idx=c_idx,
                source_row=r,
            )

        org_id_v = ws.cell(r, col_org_id).value
        username_v = ws.cell(r, col_username).value
        last_name_v = ws.cell(r, col_last_name).value
        first_name_v = ws.cell(r, col_first_name).value
        seccion_v = ws.cell(r, col_seccion).value if col_seccion else None

        # Registrar metadatos de sección
        if seccion_v:
            sec_str = str(seccion_v).strip()
            if sec_str not in secciones_map:
                cat, cla = parse_seccion_info(sec_str)
                secciones_map[sec_str] = {
                    "catalogo": cat or "",
                    "clase": cla or "",
                }

        # Extraer calificaciones por módulo
        califs: Dict[str, TracedValue] = {}
        for m_num, c_idx in mod_cols.items():
            cell_val = valor_efectivo(ws.cell(r, c_idx), r, c_idx)
            nota_num = _parse_nota(cell_val)
            if nota_num is None and cell_val is not None:
                nota_num = cell_val  # Preservar valor original si no es parseable como número
            califs[f"modulo_{m_num}"] = make_traced(
                nota_num, raw_headers[c_idx - 1] or f"Módulo {m_num}", c_idx
            )

        # Nota final
        nota_final_traced = None
        if final_col:
            final_v = valor_efectivo(ws.cell(r, final_col), r, final_col)
            fn_num = _parse_nota(final_v)
            if fn_num is not None:
                nota_final_traced = make_traced(
                    fn_num, raw_headers[final_col - 1] or "Calculated Final", final_col
                )
            elif final_v is not None:
                # Valor no parseable como número: preservar tal cual con trazabilidad
                nota_final_traced = make_traced(
                    final_v, raw_headers[final_col - 1] or "Calculated Final", final_col
                )

        # Detección de color de relleno (naranja / excluido)
        # Se inspeccionan únicamente las columnas del estudiante (identidad, nombre y
        # calificaciones). Así se evita que columnas de estado/semáforo con colores
        # ámbar/dorado marquen por error a todo el curso como excluido.
        is_orange = False
        fill_details: Dict[str, Any] = {}
        for c in sorted(student_cols):
            c_obj = ws.cell(r, c)
            f = c_obj.fill
            if f and f.fill_type and f.fill_type != "none":
                fg = getattr(f, "fgColor", None)
                if not fg:
                    continue

                raw_rgb = getattr(fg, "rgb", None)
                rgb_str = None
                if raw_rgb is not None:
                    if isinstance(raw_rgb, str):
                        clean_raw = raw_rgb.strip().lstrip("#")
                        rgb_str = clean_raw[2:] if len(clean_raw) == 8 else clean_raw
                    else:
                        # Si es un descriptor o instancia RGB de openpyxl
                        try:
                            s = str(raw_rgb).strip().lstrip("#")
                            if len(s) in (6, 8):
                                rgb_str = s[2:] if len(s) == 8 else s
                        except Exception:
                            rgb_str = None

                theme_val = getattr(fg, "theme", None)
                if not isinstance(theme_val, int):
                    theme_val = None

                tint_val = getattr(fg, "tint", None)
                if not isinstance(tint_val, (int, float)):
                    tint_val = None

                if is_orange_fill_exact(rgb_str, theme_val, tint_val):
                    is_orange = True
                    fill_details = {
                        "hex": rgb_str,
                        "theme": theme_val,
                        "tint": tint_val,
                        "fill_type": f.fill_type,
                        "col": c,
                    }
                    break

        est = EstudianteNotas(
            row_number=r,
            is_hidden_row=is_hidden,
            org_defined_id=make_traced(org_id_v, "OrgDefinedId", col_org_id) if org_id_v is not None else None,
            username=make_traced(username_v, "Username", col_username) if username_v is not None else None,
            first_name=make_traced(first_name_v, "First Name", col_first_name) if first_name_v is not None else None,
            last_name=make_traced(last_name_v, "Last Name", col_last_name) if last_name_v is not None else None,
            seccion=make_traced(seccion_v, "Sección", col_seccion) if seccion_v is not None else None,
            calificaciones_modulos=califs,
            nota_final=nota_final_traced,
            is_orange_fill=is_orange,
            fill_details=fill_details,
        )
        estudiantes.append(est)

    wb.close()
    if wb_valores is not None:
        wb_valores.close()
    return estudiantes, secciones_map
