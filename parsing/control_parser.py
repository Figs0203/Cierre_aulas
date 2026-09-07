"""
Parser especializado para el archivo Control_Aulas_Formadores_2026.xlsx.

Itera sobre las hojas activas (Tabla_ABBUEI, Tabla_CMI, etc., ignorando
siempre 'Creación aulas GDA') y extrae los metadatos de las clases asociadas:
- Programa Académico
- Catálogo y Clase
- Nombre de la Sección
- Docente / Formador Principal
- Fechas de Inicio, Finalización y Fin 2
- Estado de la columna CERTIFICADOS (para advertencia de aula ya cerrada)
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List, Optional

import openpyxl

from config.settings import CONTROL_AULAS_ALL_SHEETS, CONTROL_AULAS_IGNORED_SHEET
from core.models import AulaMetadata
from core.traceability import TracedValue


def parse_control_aulas(
    file_path: Path | str,
    target_aula: Optional[str] = None,
    clases_objetivo: Optional[List[str]] = None,
) -> Dict[str, AulaMetadata]:
    """Extrae metadatos de clases desde Control_Aulas_Formadores.

    Busca en todas las hojas activas por coincidencia de aula o número de clase.

    Args:
        file_path: Ruta al archivo Control_Aulas_Formadores.
        target_aula: Código del aula (ej. 'ABBUEI205').
        clases_objetivo: Lista opcional de clases a buscar (ej. ['6872', '001']).

    Returns:
        Diccionario {clase_id: AulaMetadata}
    """
    path = Path(file_path)
    file_name = path.name
    clean_target = target_aula.strip().upper() if target_aula else None
    clean_clases = {str(c).strip() for c in clases_objetivo} if clases_objetivo else set()

    wb = openpyxl.load_workbook(path, data_only=True, read_only=False)

    metadatos_por_clase: Dict[str, AulaMetadata] = {}

    # Filtrar hojas a examinar
    sheets_to_check = [
        s for s in wb.sheetnames
        if s != CONTROL_AULAS_IGNORED_SHEET and (s in CONTROL_AULAS_ALL_SHEETS or s.startswith("Tabla_"))
    ]

    for sheet_name in sheets_to_check:
        ws = wb[sheet_name]
        max_col = ws.max_column or 1
        max_row = ws.max_row or 1

        # Mapear encabezados de la fila 1
        col_map: Dict[str, int] = {}
        for c in range(1, max_col + 1):
            val = ws.cell(1, c).value
            if val:
                norm_h = str(val).strip().upper()
                col_map[norm_h] = c

        col_programa = col_map.get("PROGRAMA ACADÉMICO") or col_map.get("PROGRAMA ACADEMICO")
        col_catalogo = col_map.get("CATÁLOGO") or col_map.get("CATALOGO")
        col_clase = col_map.get("CLASE")
        col_seccion = col_map.get("NOMBRE DE LA SECCIÓN") or col_map.get("NOMBRE DE LA SECCION")
        col_docente = (
            col_map.get("DOCENTE PRINCIPAL")
            or col_map.get("FORMADOR PRINCIPAL")
            or col_map.get("PROFESOR(A) PRINCIPAL")
        )
        col_f_inicio = col_map.get("FECHA DE INICIO DEL CURSO")
        col_f_fin = col_map.get("FECHA DE FINALIZACIÓN DEL CURSO") or col_map.get("FECHA DE FINALIZACION DEL CURSO")
        col_f_fin_2 = col_map.get("FECHA 2 DE FINALIZACIÓN DEL CURSO") or col_map.get("FECHA 2 DE FINALIZACION DEL CURSO")
        col_certificados = col_map.get("CERTIFICADOS")

        for r in range(2, max_row + 1):
            clase_val = ws.cell(r, col_clase).value if col_clase else None
            clase_str = str(clase_val).strip() if clase_val is not None else ""
            if isinstance(clase_val, float) and clase_val.is_integer():
                clase_str = str(int(clase_val))

            # Verificar coincidencia por clase o por contenido de fila
            match_row = False
            if clean_clases and clase_str in clean_clases:
                match_row = True
            elif clean_target:
                # Revisar si target_aula aparece en alguna celda de esta fila
                row_vals = [str(ws.cell(r, c).value or "") for c in range(1, max_col + 1)]
                if any(clean_target in v.upper() for v in row_vals):
                    match_row = True

            if not match_row:
                continue

            def make_traced(val: Any, col_name: str, c_idx: Optional[int]) -> Optional[TracedValue]:
                if val is None or c_idx is None:
                    return None
                return TracedValue(
                    value=val,
                    source_file=file_name,
                    source_sheet=sheet_name,
                    source_column=col_name,
                    source_col_idx=c_idx,
                    source_row=r,
                )

            prog_v = ws.cell(r, col_programa).value if col_programa else None
            cat_v = ws.cell(r, col_catalogo).value if col_catalogo else None
            doc_v = ws.cell(r, col_docente).value if col_docente else None
            f_ini_v = ws.cell(r, col_f_inicio).value if col_f_inicio else None
            f_fin_v = ws.cell(r, col_f_fin).value if col_f_fin else None
            f_fin_2_v = ws.cell(r, col_f_fin_2).value if col_f_fin_2 else None
            cert_v = ws.cell(r, col_certificados).value if col_certificados else None

            meta = AulaMetadata(
                codigo_aula=TracedValue(clean_target or "", file_name, sheet_name, "Aula", 0, r),
                programa_academico=make_traced(prog_v, "PROGRAMA ACADÉMICO", col_programa),
                catalogo=make_traced(cat_v, "CATÁLOGO", col_catalogo),
                docente_principal=make_traced(doc_v, "DOCENTE PRINCIPAL", col_docente),
                fecha_inicio=make_traced(f_ini_v, "FECHA DE INICIO DEL CURSO", col_f_inicio),
                fecha_fin=make_traced(f_fin_v, "FECHA DE FINALIZACIÓN DEL CURSO", col_f_fin),
                fecha_fin_2=make_traced(f_fin_2_v, "FECHA 2 DE FINALIZACIÓN DEL CURSO", col_f_fin_2),
                estado_certificados=make_traced(cert_v, "CERTIFICADOS", col_certificados),
            )

            key = clase_str if clase_str else f"row_{r}"
            metadatos_por_clase[key] = meta

    wb.close()
    return metadatos_por_clase
