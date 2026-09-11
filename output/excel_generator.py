"""
Generador del Archivo Excel Auxiliar de Cierre.

Produce el libro 'Cierre_Aula_[CODIGO]_[YYYYMMDD_HHMMSS].xlsx' con 5 hojas
estructuradas exactamente en el orden y formato requeridos para que el usuario
pueda copiar y pegar manualmente hacia los archivos oficiales:

1. RESUMEN: Ficha técnica, métricas por clase, aprobados, excluidos y alertas.
2. SISTEMATIZACION: Tabla con las 55 columnas oficiales lista para copiar/pegar.
3. CERTIFICADOS: Tabla con las 13 columnas oficiales (solo aprobados).
4. ACTUALIZACION_CONTROL: Propuesta de actualización manual en Control de Aulas.
5. VALIDACIONES_Y_TRAZABILIDAD: Auditoría celda a celda (TracedValue) y discrepancias.
"""

from __future__ import annotations

from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from core.models import AulaCierreResult, AulaMetadata, ClaseGroup
from normalization.dates import (
    format_certificate_date,
    format_ciclo_date,
    format_semester_label,
    parse_date,
)
from normalization.text import format_ciclo, format_docente_coin


# Paleta de colores profesionales para la salida
_FONT_HEADER = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
_FONT_TITLE = Font(name="Calibri", size=14, bold=True, color="1F4E79")
_FONT_SUBTITLE = Font(name="Calibri", size=11, bold=True, color="333333")
_FONT_BODY = Font(name="Calibri", size=10)
_FONT_MUTED = Font(name="Calibri", size=9, italic=True, color="666666")

# Fuente oficial de la tabla de Certificados (confirmado: Zurich Cn BT 11).
_FONT_CERTIFICADO = Font(name="Zurich Cn BT", size=11)

_FILL_PRIMARY = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
_FILL_HEADER = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
_FILL_SUCCESS = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
_FILL_WARNING = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
_FILL_ALERT = PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid")

_BORDER_COLOR = "000000"

_BORDER_THIN = Border(
    left=Side(style="thin", color=_BORDER_COLOR),
    right=Side(style="thin", color=_BORDER_COLOR),
    top=Side(style="thin", color=_BORDER_COLOR),
    bottom=Side(style="thin", color=_BORDER_COLOR),
)

# La celda de fin de clase conserva los cuatro bordes: tres lados finos + inferior grueso.
_BORDER_CLASS_SEPARATOR = Border(
    left=Side(style="thin", color=_BORDER_COLOR),
    right=Side(style="thin", color=_BORDER_COLOR),
    top=Side(style="thin", color=_BORDER_COLOR),
    bottom=Side(style="medium", color=_BORDER_COLOR),
)


def generate_cierre_excel(
    result: AulaCierreResult,
    codigo_curso: str,
    sist_headers: List[str],
    start_consecutivo: int = 1,
    output_dir: Optional[Path | str] = None,
    metadatos_control: Optional[Dict[str, AulaMetadata]] = None,
) -> Path:
    """Construye y guarda el archivo Excel auxiliar de cierre con las 5 hojas.

    Args:
        result: AulaCierreResult con las clases y matches procesados.
        codigo_curso: Código del curso ingresado por el monitor (ej. 'ABBUEI205').
        sist_headers: Lista de las 55 columnas oficiales de Sistematización.
        start_consecutivo: Número correlativo inicial para la hoja Certificados.
        output_dir: Directorio de salida (por defecto 'output/').
        metadatos_control: Metadatos extraídos de Control de Aulas por clase.

    Returns:
        Path del archivo Excel generado.
    """
    if output_dir is None:
        out_path_dir = Path("output")
    else:
        out_path_dir = Path(output_dir)

    out_path_dir.mkdir(parents=True, exist_ok=True)

    target_aula_code = result.aula.codigo_aula.value if result.aula and result.aula.codigo_aula else codigo_curso
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"Cierre_Aula_{target_aula_code}_{timestamp}.xlsx"
    out_file = out_path_dir / filename

    wb = openpyxl.Workbook()

    # ============================================================
    # HOJA 1: RESUMEN
    # ============================================================
    ws_resumen = wb.active
    ws_resumen.title = "RESUMEN"
    ws_resumen.views.sheetView[0].showGridLines = True

    ws_resumen["A1"] = f"INFORME DE CIERRE DE AULA — {target_aula_code}"
    ws_resumen["A1"].font = _FONT_TITLE
    ws_resumen["A2"] = f"Generado el: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | Modo: Asistente Local COIN"
    ws_resumen["A2"].font = _FONT_MUTED

    # Ficha Técnica
    ws_resumen["A4"] = "FICHA TÉCNICA DEL AULA"
    ws_resumen["A4"].font = _FONT_SUBTITLE

    ficha_rows = [
        ("Código de Aula / Solicitud:", str(target_aula_code)),
        ("Código del Curso (Certificados):", str(codigo_curso)),
        ("Total Clases Detectadas:", str(len(result.clases))),
        ("Total Estudiantes en Sistematización:", str(result.total_estudiantes)),
        ("Total Estudiantes Matched con Notas:", str(result.total_matched)),
        ("Total Estudiantes Excluidos (Relleno Naranja):", str(result.total_excluded)),
        ("Total Estudiantes sin Notas en Sistematización:", str(result.total_unmatched_sist)),
        ("Total Registros en Notas sin Match:", str(result.total_unmatched_notas)),
    ]

    for idx, (label, val) in enumerate(ficha_rows, start=5):
        ws_resumen.cell(idx, 1, label).font = Font(name="Calibri", size=10, bold=True)
        ws_resumen.cell(idx, 2, val).font = _FONT_BODY

    # Métricas por Clase
    start_r = len(ficha_rows) + 6
    ws_resumen.cell(start_r, 1, "MÉTRICAS POR CLASE").font = _FONT_SUBTITLE

    class_headers = ["Clase", "Total Sistematización", "Con Notas", "Aprobados", "No Aprobados", "Abandonaron", "Excluidos (Naranja)"]
    for c_idx, h in enumerate(class_headers, start=1):
        cell = ws_resumen.cell(start_r + 1, c_idx, h)
        cell.font = _FONT_HEADER
        cell.fill = _FILL_HEADER
        cell.alignment = Alignment(horizontal="center")

    curr_r = start_r + 2
    for cg in result.clases:
        aprobados = sum(1 for m in cg.matches if m.estudiante_sist.estado_calculado and "APROB" in m.estudiante_sist.estado_calculado.upper() and "NO" not in m.estudiante_sist.estado_calculado.upper())
        no_aprobados = sum(1 for m in cg.matches if m.estudiante_sist.estado_calculado and "NO APROB" in m.estudiante_sist.estado_calculado.upper())
        abandonaron = sum(1 for m in cg.matches if m.estudiante_sist.estado_calculado and "ABANDON" in m.estudiante_sist.estado_calculado.upper())

        row_vals = [
            str(cg.clase_id),
            len(cg.estudiantes_sist),
            len(cg.matches),
            aprobados,
            no_aprobados,
            abandonaron,
            len(cg.excluded_orange),
        ]
        for c_idx, val in enumerate(row_vals, start=1):
            cell = ws_resumen.cell(curr_r, c_idx, val)
            cell.font = _FONT_BODY
            cell.alignment = Alignment(horizontal="center")
            cell.border = _BORDER_THIN
        curr_r += 1

    # ============================================================
    # HOJA 2: SISTEMATIZACION (55 Columnas Oficiales)
    # ============================================================
    ws_sist = wb.create_sheet(title="SISTEMATIZACION")
    ws_sist.views.sheetView[0].showGridLines = True

    # Encabezados en fila 1 (en mayúsculas oficiales)
    for c_idx, h_text in enumerate(sist_headers, start=1):
        cell = ws_sist.cell(1, c_idx, str(h_text).upper())
        cell.font = _FONT_HEADER
        cell.fill = _FILL_PRIMARY
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Recopilar todos los estudiantes del aula y ordenarlos por su fila original en 'Formato'
    all_sist_students: List[EstudianteSistematizacion] = []
    for cg in result.clases:
        for m in cg.matches:
            all_sist_students.append(m.estudiante_sist)
        for es in cg.unmatched_sist:
            all_sist_students.append(es)

    # ORDEN EXACTO: el mismo orden de filas de Sistematizacion_Cursos_COIN_2026.xlsx
    all_sist_students.sort(key=lambda x: x.row_number)

    col_correo_idx = None
    if all_sist_students and all_sist_students[0].column_indices:
        col_correo_idx = all_sist_students[0].column_indices.get("correo")

    SPECIAL_CASING_VALUES = {
        "aprobó": "Aprobó",
        "aprobo": "Aprobó",
        "no aprobó": "No aprobó",
        "no aprobo": "No aprobó",
        "abandonó": "Abandonó",
        "abandono": "Abandonó",
        "no aplica": "No aplica",
    }

    sist_r = 2
    for s_idx, es in enumerate(all_sist_students):
        # Determinar si esta fila cierra una clase (comparando con el siguiente estudiante)
        current_clase = str(es.clase.value if es.clase and es.clase.value else "").strip()
        is_last_of_class = (s_idx == len(all_sist_students) - 1)
        if not is_last_of_class:
            next_es = all_sist_students[s_idx + 1]
            next_clase = str(next_es.clase.value if next_es.clase and next_es.clase.value else "").strip()
            if current_clase != next_clase:
                is_last_of_class = True
        row_border = _BORDER_CLASS_SEPARATOR if is_last_of_class else _BORDER_THIN

        for col_idx in range(1, len(sist_headers) + 1):
            val = es.raw_row_data.get(col_idx, None)
            # Todo en mayúsculas, excepto correos electrónicos y estados con formato oficial exacto
            if isinstance(val, str) and not val.startswith("="):
                is_email = (col_idx == col_correo_idx) or ("@" in val and "." in val)
                if is_email:
                    pass  # Mantener correos tal cual
                else:
                    val_clean = val.strip()
                    val_lower = val_clean.lower()
                    if val_lower in SPECIAL_CASING_VALUES:
                        val = SPECIAL_CASING_VALUES[val_lower]
                    else:
                        val = val.upper()
            cell = ws_sist.cell(sist_r, col_idx, val)
            cell.font = _FONT_BODY
            cell.border = row_border
        sist_r += 1

    # ============================================================
    # HOJA 3: CERTIFICADOS (13 Columnas Oficiales)
    # ============================================================
    ws_cert = wb.create_sheet(title="CERTIFICADOS")
    ws_cert.views.sheetView[0].showGridLines = True

    cert_headers = [
        "N°", "Ciclo", "Docente COIN", "Nombres", "Apellidos",
        "Documento de identidad", "Correo Electrónico", "Código del certificado",
        "Fecha de envío", "Año de certificación", "Semestre", "Total certificados",
        "Fecha de totalización",
    ]

    for c_idx, h_text in enumerate(cert_headers, start=1):
        cell = ws_cert.cell(1, c_idx, h_text.upper())
        cell.font = _FONT_HEADER
        cell.fill = _FILL_HEADER
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # Recopilar todos los estudiantes aprobados de todas las clases
    all_approved: List[Tuple[StudentMatch, ClaseGroup]] = []
    for cg in result.clases:
        for m in cg.matches:
            est_status = str(m.estudiante_sist.estado_calculado or "").strip().lower()
            if "aprob" in est_status and "no" not in est_status:
                all_approved.append((m, cg))

    # ORDEN EXACTO: el mismo orden de aparición en la hoja 'Formato' de Sistematización
    all_approved.sort(key=lambda item: item[0].estudiante_sist.row_number)

    cert_r = 2
    current_n = start_consecutivo

    for idx, (m, cg) in enumerate(all_approved):
        es = m.estudiante_sist
        meta = (metadatos_control or {}).get(cg.clase_id) or AulaMetadata()
        programa_str = meta.programa_academico.value if meta.programa_academico else ""
        catalogo_str = meta.catalogo.value if meta.catalogo else ""
        f_ini_dt = meta.fecha_inicio.value if meta.fecha_inicio else None
        f_fin_dt = parse_date(meta.fecha_fin.value) if meta.fecha_fin else None

        f_ini_str = format_ciclo_date(f_ini_dt) if f_ini_dt else ""
        f_fin_str = format_ciclo_date(f_fin_dt) if f_fin_dt else ""
        f_envio_cert_str = format_certificate_date(f_fin_dt) if f_fin_dt else ""

        ciclo_val = format_ciclo(cg.clase_id, catalogo_str, programa_str, f_ini_str, f_fin_str).upper()

        formador_lider = ""
        if es.formador_lider and es.formador_lider.value:
            formador_lider = str(es.formador_lider.value).strip()
        elif meta.docente_principal and meta.docente_principal.value:
            formador_lider = str(meta.docente_principal.value).strip()

        # Docente COIN completamente en MAYÚSCULAS
        docente_coin_val = format_docente_coin(codigo_curso, formador_lider).strip().upper()
        # Columna 'Semestre': periodo académico con formato oficial 'AAAA-S' (ej. '2026-1').
        semestre_val = format_semester_label(f_fin_dt) if f_fin_dt else ""
        ano_val = f_fin_dt.year if isinstance(f_fin_dt, (date, datetime)) else 2026

        nom = str(es.nombres.value if es.nombres else "").strip().upper()
        ape = str(es.apellidos.value if es.apellidos else "").strip().upper()
        doc = es.documento.value if es.documento else ""
        cor = str(es.correo.value if es.correo else "").strip()  # CORREOS SIN MAYÚSCULAS

        row_cert = [
            current_n,
            ciclo_val,
            docente_coin_val,
            nom,
            ape,
            doc,
            cor,
            str(codigo_curso).strip().upper(),
            f_envio_cert_str.upper(),
            ano_val,
            semestre_val,
            1,  # Siempre 1 por cada estudiante aprobado
            None,
        ]

        # Separación con borde inferior grueso al terminar cada clase distinta
        current_clase = str(es.clase.value if es.clase and es.clase.value else cg.clase_id).strip()
        is_last_of_class = False
        if idx == len(all_approved) - 1:
            is_last_of_class = True
        else:
            next_m, next_cg = all_approved[idx + 1]
            next_es = next_m.estudiante_sist
            next_clase = str(next_es.clase.value if next_es.clase and next_es.clase.value else next_cg.clase_id).strip()
            if current_clase != next_clase:
                is_last_of_class = True

        border_cell = _BORDER_CLASS_SEPARATOR if is_last_of_class else _BORDER_THIN

        for c_idx, val in enumerate(row_cert, start=1):
            cell = ws_cert.cell(cert_r, c_idx, val)
            cell.font = _FONT_CERTIFICADO
            cell.border = border_cell
            if c_idx in (1, 10, 11, 12):
                cell.alignment = Alignment(horizontal="center")

        current_n += 1
        cert_r += 1

    # ============================================================
    # HOJA 4: ACTUALIZACION_CONTROL
    # ============================================================
    ws_ctrl = wb.create_sheet(title="ACTUALIZACION_CONTROL")
    ws_ctrl.views.sheetView[0].showGridLines = True

    ctrl_headers = ["Hoja en Control_Aulas", "Clase", "Columna Objetivo", "Valor a Colocar", "Acción Manual a Realizar"]
    for c_idx, h in enumerate(ctrl_headers, start=1):
        cell = ws_ctrl.cell(1, c_idx, h)
        cell.font = _FONT_HEADER
        cell.fill = _FILL_HEADER
        cell.alignment = Alignment(horizontal="center")

    ctrl_r = 2
    fecha_cierre_hoy = datetime.now().strftime("%d/%m/%Y")
    for cg in result.clases:
        row_ctrl = [
            f"Tabla correspondiente al curso",
            str(cg.clase_id),
            "CERTIFICADOS",
            fecha_cierre_hoy,
            f"Colocar la fecha {fecha_cierre_hoy} en la columna CERTIFICADOS de la fila de la clase {cg.clase_id} para marcarla como cerrada.",
        ]
        for c_idx, val in enumerate(row_ctrl, start=1):
            cell = ws_ctrl.cell(ctrl_r, c_idx, val)
            cell.font = _FONT_BODY
            cell.border = _BORDER_THIN
        ctrl_r += 1

    # ============================================================
    # HOJA 5: VALIDACIONES_Y_TRAZABILIDAD
    # ============================================================
    ws_traz = wb.create_sheet(title="VALIDACIONES_Y_TRAZABILIDAD")
    ws_traz.views.sheetView[0].showGridLines = True

    traz_headers = [
        "Clase", "Estudiante", "Documento", "Correo", "Método de Match",
        "Fila en Sistematización", "Fila en Notas", "Estado Calculado", "Advertencias / Notas"
    ]
    for c_idx, h in enumerate(traz_headers, start=1):
        cell = ws_traz.cell(1, c_idx, h.upper())
        cell.font = _FONT_HEADER
        cell.fill = _FILL_PRIMARY
        cell.alignment = Alignment(horizontal="center")

    traz_r = 2
    for cg in result.clases:
        for m in cg.matches:
            es = m.estudiante_sist
            en = m.estudiante_nota
            nom = f"{es.nombres.value if es.nombres else ''} {es.apellidos.value if es.apellidos else ''}".strip().upper()
            doc = str(es.documento.value) if es.documento else ""
            cor = str(es.correo.value) if es.correo else ""
            adv_parts = []
            if m.cross_validation_label:
                adv_parts.append(m.cross_validation_label)
            if es.is_hidden_row:
                adv_parts.append("Fila oculta en Sistematización (procesada)")
            if en.is_hidden_row:
                adv_parts.append("Fila oculta en Notas (procesada)")
            if m.warnings:
                adv_parts.extend(m.warnings)

            adv = " | ".join(adv_parts) if adv_parts else "OK"

            row_traz = [
                str(cg.clase_id).upper(),
                nom,
                str(doc).strip().upper(),
                cor,
                str(m.match_key).upper(),
                es.row_number,
                en.row_number,
                str(es.estado_calculado or ""),
                adv.upper(),
            ]
            for c_idx, val in enumerate(row_traz, start=1):
                cell = ws_traz.cell(traz_r, c_idx, val)
                cell.font = _FONT_BODY
                cell.border = _BORDER_THIN
            traz_r += 1

        # Agregar advertencias para los excluidos (naranja)
        for en in cg.excluded_orange:
            nom = f"{en.first_name.value if en.first_name else ''} {en.last_name.value if en.last_name else ''}".strip().upper()
            adv_orange = "Estudiante con relleno naranja en Notas (retirado/excluido del proceso)."
            if en.is_hidden_row:
                adv_orange += " [Fila oculta en Notas]"
            row_traz = [
                str(cg.clase_id).upper(),
                nom,
                str(en.org_defined_id.value) if en.org_defined_id else "",
                str(en.username.value) if en.username else "",
                "NO ASOCIADO (EXCLUIDO)",
                "N/A",
                en.row_number,
                "EXCLUIDO",
                adv_orange.upper(),
            ]
            for c_idx, val in enumerate(row_traz, start=1):
                cell = ws_traz.cell(traz_r, c_idx, val)
                cell.font = _FONT_BODY
                cell.fill = _FILL_ALERT
                cell.border = _BORDER_THIN
            traz_r += 1

        # Agregar advertencias para los que quedaron sin notas
        for es in cg.unmatched_sist:
            nom = f"{es.nombres.value if es.nombres else ''} {es.apellidos.value if es.apellidos else ''}".strip().upper()
            adv_sist = "Estudiante registrado en Sistematización pero sin registro en el archivo de Notas."
            if es.is_hidden_row:
                adv_sist += " [Fila oculta en Sistematización]"
            row_traz = [
                str(cg.clase_id).upper(),
                nom,
                str(es.documento.value if es.documento else "").upper(),
                str(es.correo.value if es.correo else ""),
                "SIN NOTAS",
                es.row_number,
                "N/A",
                "SIN NOTAS",
                adv_sist.upper(),
            ]
            for c_idx, val in enumerate(row_traz, start=1):
                cell = ws_traz.cell(traz_r, c_idx, val)
                cell.font = _FONT_BODY
                cell.fill = _FILL_WARNING
                cell.border = _BORDER_THIN
            traz_r += 1

    # Configuración de visibilidad, filtros y auto-ajuste de ancho de columnas
    for ws_curr in wb.worksheets:
        max_r = ws_curr.max_row or 1
        max_c = ws_curr.max_column or 1

        # Asegurar que ninguna fila ni columna quede oculta en la salida auxiliar
        for r_idx in range(1, max_r + 1):
            ws_curr.row_dimensions[r_idx].hidden = False
        for c_idx in range(1, max_c + 1):
            col_letter = get_column_letter(c_idx)
            ws_curr.column_dimensions[col_letter].hidden = False

        # Activar autofiltro en las tablas de datos para facilitar revisión y copiado
        if ws_curr.title in ("SISTEMATIZACION", "CERTIFICADOS", "ACTUALIZACION_CONTROL", "VALIDACIONES_Y_TRAZABILIDAD") and max_r > 1:
            ws_curr.auto_filter.ref = f"A1:{get_column_letter(max_c)}{max_r}"

        # Ancho de columnas adaptativo
        for col in ws_curr.columns:
            max_len = max(len(str(cell.value or "")) for cell in col[:100])
            col_letter = get_column_letter(col[0].column)
            ws_curr.column_dimensions[col_letter].width = max(max_len + 3, 11)

    wb.save(out_file)
    return out_file
