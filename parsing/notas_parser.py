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
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import openpyxl

from config.colors import (
    FILL_EXCLUDED_ORANGE,
    is_orange_fill_exact,
)
from core.models import EstudianteNotas
from core.traceability import TracedValue
from normalization.headers import find_final_grade_column, find_module_columns
from normalization.text import normalize_document, normalize_whitespace


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


def parse_notas_archivo(
    file_path: Path | str,
    sheet_name: Optional[str] = None,
) -> Tuple[List[EstudianteNotas], Dict[str, Dict[str, str]]]:
    """Extrae las notas de estudiantes desde el archivo de Interactiva.

    Args:
        file_path: Ruta al archivo Excel de Notas.
        sheet_name: Nombre de la hoja (si es None, usa 'Calificaciones' o la activa).

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
    elif "Calificaciones" in wb.sheetnames:
        sheet_name = "Calificaciones"
        ws = wb["Calificaciones"]
    elif "Grades" in wb.sheetnames:
        sheet_name = "Grades"
        ws = wb["Grades"]
    else:
        sheet_name = wb.sheetnames[0]
        ws = wb[sheet_name]

    # 1. Leer encabezados
    max_col = ws.max_column or 1
    raw_headers: List[Optional[str]] = []
    for c in range(1, max_col + 1):
        v = ws.cell(1, c).value
        raw_headers.append(str(v).strip() if v is not None else None)

    # Mapeo semántico de columnas de módulos y final
    mod_cols = find_module_columns(raw_headers)
    final_col = find_final_grade_column(raw_headers)

    # Columnas de identificación básica
    col_org_id = 1
    col_username = 2
    col_last_name = 3
    col_first_name = 4
    col_seccion = 5

    for c_idx, h in enumerate(raw_headers, start=1):
        if not h:
            continue
        h_lower = h.lower()
        if "orgdefinedid" in h_lower:
            col_org_id = c_idx
        elif "username" in h_lower:
            col_username = c_idx
        elif "last name" in h_lower or "apellidos" in h_lower:
            col_last_name = c_idx
        elif "first name" in h_lower or "nombres" in h_lower:
            col_first_name = c_idx
        elif "sección" in h_lower or "seccion" in h_lower:
            col_seccion = c_idx

    estudiantes: List[EstudianteNotas] = []
    secciones_map: Dict[str, Dict[str, str]] = {}
    max_row = ws.max_row or 1

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
        seccion_v = ws.cell(r, col_seccion).value

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
            cell_val = ws.cell(r, c_idx).value
            # Normalizar float/int
            nota_num = None
            if cell_val is not None:
                try:
                    nota_num = float(str(cell_val).replace(",", "."))
                except ValueError:
                    nota_num = cell_val
            califs[f"modulo_{m_num}"] = make_traced(
                nota_num, raw_headers[c_idx - 1] or f"Módulo {m_num}", c_idx
            )

        # Nota final
        nota_final_traced = None
        if final_col:
            final_v = ws.cell(r, final_col).value
            if final_v is not None:
                try:
                    fn_num = float(str(final_v).replace(",", "."))
                except ValueError:
                    fn_num = final_v
                nota_final_traced = make_traced(
                    fn_num, raw_headers[final_col - 1] or "Calculated Final", final_col
                )

        # Detección de color de relleno (naranja / excluido)
        # Inspecciona exhaustivamente todas las celdas de la fila
        is_orange = False
        fill_details: Dict[str, Any] = {}
        for c in range(1, max_col + 1):
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
    return estudiantes, secciones_map
