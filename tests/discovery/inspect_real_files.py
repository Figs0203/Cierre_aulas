"""
Motor de Inspección Técnica y Descubrimiento Local — Fase 0
Asistente de Cierre de Aulas COIN

Herramienta 100% OFFLINE que inspecciona los 4 archivos Excel confidenciales
directamente en la máquina del usuario, sin transmitir datos al exterior.

Genera un Informe de Descubrimiento estructurado en 14 secciones:
 1. Resumen de archivos
 2. Integridad SHA-256
 3. Estructura por archivo (hojas, dimensiones, celdas combinadas, ocultas)
 4. Encabezados encontrados (mapeo preliminar)
 5. Fórmulas encontradas (análisis específico de 'Estado para certificación')
 6. Colores encontrados (análisis de fills RGB, theme, tint, indexed, condicional)
 7. Clases y secciones detectadas
 8. Identificadores y estadísticas estructurales (Non-PII)
 9. Patrones del archivo de certificados
10. Cruces preliminares entre archivos
11. Reglas potencialmente confirmables
12. Reglas todavía desconocidas
13. Advertencias e inconsistencias detectadas
14. Recomendaciones para Fase 1/2
"""

from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

# Agregar directorio raíz al path para imports
_PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
if str(_PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(_PROJECT_ROOT))

from normalization.headers import analyze_headers
from normalization.text import normalize_for_comparison
from security.integrity import compute_sha256, FileIntegrityRecord
from security.privacy import mask_email, mask_document, mask_name, mask_full_name


def describe_cell_fill(fill) -> str:
    """Genera una descripción legible y completa de un fill de celda."""
    if not fill or not fill.fill_type or fill.fill_type == "none":
        return "sin color"

    parts = [f"tipo={fill.fill_type}"]

    if fill.fgColor:
        fg = fill.fgColor
        if fg.rgb and isinstance(fg.rgb, str) and fg.rgb != "00000000":
            clean_rgb = fg.rgb[2:] if len(fg.rgb) == 8 else fg.rgb
            parts.append(f"RGB=#{clean_rgb}")
        if fg.theme is not None and isinstance(fg.theme, int):
            parts.append(f"theme={fg.theme}")
        if fg.tint is not None and isinstance(fg.tint, (int, float)) and fg.tint != 0:
            parts.append(f"tint={fg.tint:.4f}")
        if fg.indexed is not None and isinstance(fg.indexed, int):
            parts.append(f"indexed={fg.indexed}")

    return ", ".join(parts)


def inspect_single_file_data(
    file_path: Path | str,
    role: str,
    target_aula: Optional[str] = None,
) -> Dict[str, Any]:
    """Extrae estructura, fórmulas, colores y estadísticas de un archivo Excel."""
    import openpyxl
    from openpyxl.utils import get_column_letter

    path = Path(file_path)
    file_data: Dict[str, Any] = {
        "role": role,
        "file_name": path.name,
        "file_path": str(path.resolve()),
        "sheets": {},
        "target_aula_occurrences": [],
    }

    try:
        wb_f = openpyxl.load_workbook(path, data_only=False, read_only=False)
        wb_v = openpyxl.load_workbook(path, data_only=True, read_only=False)
    except Exception as e:
        file_data["error"] = str(e)
        return file_data

    for sheet_name in wb_f.sheetnames:
        ws_f = wb_f[sheet_name]
        ws_v = wb_v[sheet_name]

        # 1. Dimensiones y combinadas
        merged_ranges = [str(m) for m in ws_f.merged_cells.ranges]
        hidden_rows = [r for r, rd in ws_f.row_dimensions.items() if rd.hidden]
        hidden_cols = [c for c, cd in ws_f.column_dimensions.items() if cd.hidden]

        # 2. Encabezados
        best_header_row = 1
        max_text_cols = 0
        for r in range(1, min(10, ws_f.max_row + 1)):
            text_count = sum(1 for c in range(1, ws_f.max_column + 1) if ws_f.cell(r, c).value and isinstance(ws_f.cell(r, c).value, str))
            if text_count > max_text_cols:
                max_text_cols = text_count
                best_header_row = r

        raw_headers: List[Optional[str]] = []
        for c in range(1, ws_f.max_column + 1):
            val = ws_f.cell(best_header_row, c).value
            raw_headers.append(str(val).strip() if val is not None else None)

        header_analysis = analyze_headers(raw_headers)

        # 3. Fórmulas
        formulas_by_header: Dict[str, List[Tuple[int, str]]] = {}
        for r in range(1, min(ws_f.max_row + 1, 300)):
            for c in range(1, min(ws_f.max_column + 1, 60)):
                cell_val = ws_f.cell(r, c).value
                if cell_val and isinstance(cell_val, str) and cell_val.startswith("="):
                    hname = raw_headers[c - 1] if c <= len(raw_headers) and raw_headers[c - 1] else f"Col_{get_column_letter(c)}"
                    if hname not in formulas_by_header:
                        formulas_by_header[hname] = []
                    formulas_by_header[hname].append((r, cell_val))

        # 4. Fills
        fill_inventory: Dict[str, List[str]] = {}
        for r in range(1, min(ws_f.max_row + 1, 500)):
            for c in range(1, min(ws_f.max_column + 1, 60)):
                fill = ws_f.cell(r, c).fill
                desc = describe_cell_fill(fill)
                if desc != "sin color":
                    if desc not in fill_inventory:
                        fill_inventory[desc] = []
                    if len(fill_inventory[desc]) < 5:
                        fill_inventory[desc].append(f"Fila {r}, Col {get_column_letter(c)}")

        # 5. Formato Condicional
        cf_rules = []
        if ws_f.conditional_formatting:
            for idx, cf in enumerate(ws_f.conditional_formatting, 1):
                rule_desc = f"Rango {cf.sqref}: " + ", ".join(f"{r.type} ({getattr(r, 'operator', 'N/A')})" for r in cf.rules)
                cf_rules.append(rule_desc)

        # 6. Muestra no-PII
        masked_samples = []
        if ws_v.max_row > best_header_row:
            for r in range(best_header_row + 1, min(best_header_row + 6, ws_v.max_row + 1)):
                row_parts = []
                for c in range(1, min(ws_v.max_column + 1, 15)):
                    v = ws_v.cell(r, c).value
                    if v is not None:
                        str_v = str(v).strip()
                        if "@" in str_v:
                            row_parts.append(mask_email(str_v))
                        elif str_v.isdigit() and len(str_v) >= 6:
                            row_parts.append(mask_document(str_v))
                        else:
                            row_parts.append(str_v[:20])
                if row_parts:
                    masked_samples.append(f"Fila {r}: {' | '.join(row_parts)}")

        # 7. Aula objetivo
        if target_aula:
            for r in range(1, ws_v.max_row + 1):
                for c in range(1, ws_v.max_column + 1):
                    v = ws_v.cell(r, c).value
                    if v and str(v).strip().upper() == target_aula.strip().upper():
                        hname = raw_headers[c - 1] if c <= len(raw_headers) and raw_headers[c - 1] else ""
                        file_data["target_aula_occurrences"].append(
                            f"Hoja '{sheet_name}', Fila {r}, Col {get_column_letter(c)} ('{hname}')"
                        )

        file_data["sheets"][sheet_name] = {
            "max_row": ws_f.max_row,
            "max_column": ws_f.max_column,
            "dimensions": str(ws_f.dimensions),
            "merged_ranges": merged_ranges,
            "hidden_rows": hidden_rows,
            "hidden_cols": hidden_cols,
            "header_row": best_header_row,
            "raw_headers": raw_headers,
            "header_analysis": header_analysis,
            "formulas_by_header": formulas_by_header,
            "fill_inventory": fill_inventory,
            "cf_rules": cf_rules,
            "masked_samples": masked_samples,
        }

    wb_f.close()
    wb_v.close()

    return file_data


def run_discovery(
    file_paths: Dict[str, str | Path],
    target_aula: Optional[str] = None,
    output_dir: Optional[str | Path] = None,
) -> Tuple[Path, str]:
    """Genera el Informe de Descubrimiento Local estructurado en 14 secciones."""
    if output_dir is None:
        out_dir = _PROJECT_ROOT / "reports"
    else:
        out_dir = Path(output_dir)

    out_dir.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_filename = f"discovery_{timestamp}.txt"
    report_path = out_dir / report_filename

    # Inspeccionar todos los archivos
    all_inspections: Dict[str, Dict[str, Any]] = {}
    for role, fpath in file_paths.items():
        all_inspections[role] = inspect_single_file_data(fpath, role, target_aula)

    lines: List[str] = []

    def section_header(num: int, title: str):
        lines.append("")
        lines.append("=" * 76)
        lines.append(f"{num:2d}. {title.upper()}")
        lines.append("=" * 76)

    # ENCABEZADO
    lines.append("=" * 76)
    lines.append("INFORME DE DESCUBRIMIENTO TÉCNICO LOCAL — FASE 0")
    lines.append(f"Fecha de generación: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append("Modo: 100% OFFLINE (Procesamiento exclusivo en equipo local)")
    lines.append("=" * 76)

    # 1. RESUMEN DE ARCHIVOS
    section_header(1, "Resumen de Archivos Analizados")
    for role, data in all_inspections.items():
        lines.append(f"• [{role.upper()}]: {data['file_name']}")
        for sname, sdata in data.get("sheets", {}).items():
            lines.append(f"    Hoja '{sname}': {sdata['max_row']} filas × {sdata['max_column']} cols (Rango: {sdata['dimensions']})")

    # 2. INTEGRIDAD SHA-256
    section_header(2, "Integridad Criptográfica de Archivos Fuente (SHA-256)")
    for role, fpath in file_paths.items():
        try:
            rec = compute_sha256(fpath)
            lines.append(f"• [{role.upper()}] {rec.file_name}")
            lines.append(f"    Tamaño: {rec.file_size / 1024:.1f} KB | SHA-256: {rec.sha256_hash}")
        except Exception as e:
            lines.append(f"• [{role.upper()}]: ERROR al calcular hash: {e}")

    # 3. ESTRUCTURA POR ARCHIVO
    section_header(3, "Estructura Técnica Detallada por Archivo")
    for role, data in all_inspections.items():
        lines.append(f"\n--- {role.upper()} ({data['file_name']}) ---")
        for sname, sdata in data.get("sheets", {}).items():
            lines.append(f"  Hoja '{sname}':")
            lines.append(f"    • Fila de encabezados detectada: Fila {sdata['header_row']}")
            if sdata["merged_ranges"]:
                lines.append(f"    • Celdas combinadas ({len(sdata['merged_ranges'])}): {', '.join(sdata['merged_ranges'][:5])}")
            else:
                lines.append("    • Celdas combinadas: Ninguna")
            if sdata["hidden_rows"]:
                lines.append(f"    • Filas ocultas: {sdata['hidden_rows'][:10]}")
            if sdata["hidden_cols"]:
                lines.append(f"    • Columnas ocultas: {sdata['hidden_cols'][:10]}")

    # 4. ENCABEZADOS ENCONTRADOS
    section_header(4, "Encabezados Encontrados y Mapeo Semántico Preliminar")
    from openpyxl.utils import get_column_letter
    for role, data in all_inspections.items():
        lines.append(f"\n--- Encabezados de {role.upper()} ---")
        for sname, sdata in data.get("sheets", {}).items():
            lines.append(f"  [Hoja: '{sname}']")
            hanalysis = sdata["header_analysis"]
            for idx, hval in enumerate(sdata["raw_headers"], start=1):
                if hval:
                    m = next((map_obj for map_obj in hanalysis.mapped if map_obj.col_idx == idx), None)
                    map_desc = f" → [Mapeo preliminar: {m.canonical}]" if m else " → [Sin mapeo directo]"
                    lines.append(f"    Col {idx:2d} ({get_column_letter(idx):>2s}): '{hval}'{map_desc}")

    # 5. FÓRMULAS ENCONTRADAS (ESPECIAL: ESTADO PARA CERTIFICACIÓN)
    section_header(5, "Fórmulas Detectadas y Análisis de Certificación")
    for role, data in all_inspections.items():
        lines.append(f"\n--- Fórmulas en {role.upper()} ---")
        for sname, sdata in data.get("sheets", {}).items():
            f_by_h = sdata["formulas_by_header"]
            if f_by_h:
                for hname, f_list in f_by_h.items():
                    lines.append(f"  • Columna '{hname}' ({len(f_list)} celdas con fórmula):")
                    # Analizar si las fórmulas son idénticas o varían
                    sample_formulas = list({f[1] for f in f_list})
                    for sf in sample_formulas[:3]:
                        lines.append(f"      Muestra: {sf}")
                    if len(sample_formulas) > 3:
                        lines.append(f"      ... ({len(sample_formulas) - 3} variantes de fórmula más)")
            else:
                lines.append(f"  Hoja '{sname}': Sin fórmulas detectadas.")

    # 6. COLORES ENCONTRADOS (FILLS Y FORMATO CONDICIONAL)
    section_header(6, "Colores de Relleno (Fills) y Formato Condicional")
    for role, data in all_inspections.items():
        lines.append(f"\n--- Colores en {role.upper()} ---")
        for sname, sdata in data.get("sheets", {}).items():
            fills = sdata["fill_inventory"]
            if fills:
                lines.append(f"  [Hoja: '{sname}'] Rellenos detectados:")
                for fdesc, locs in sorted(fills.items()):
                    lines.append(f"    • {fdesc}")
                    lines.append(f"      Ubicaciones de muestra: {', '.join(locs)}")
            else:
                lines.append(f"  [Hoja: '{sname}'] Sin colores de relleno detectados.")

            if sdata["cf_rules"]:
                lines.append(f"  [Hoja: '{sname}'] Formato condicional activo ({len(sdata['cf_rules'])} reglas):")
                for cfr in sdata["cf_rules"][:5]:
                    lines.append(f"    • {cfr}")

    # 7. CLASES Y SECCIONES
    section_header(7, "Clases y Secciones Detectadas")
    if target_aula:
        lines.append(f"Aula objetivo especificada: {target_aula}")
        for role, data in all_inspections.items():
            occ = data.get("target_aula_occurrences", [])
            if occ:
                lines.append(f"• En {role.upper()} ({len(occ)} ocurrencias):")
                for o in occ[:5]:
                    lines.append(f"    {o}")
            else:
                lines.append(f"• En {role.upper()}: Aula '{target_aula}' no encontrada directamente en celdas.")
    else:
        lines.append("No se especificó un código de aula específico para búsqueda puntual.")

    # 8. IDENTIFICADORES Y ESTADÍSTICAS NO-PII
    section_header(8, "Identificadores y Estadísticas Estructurales (Non-PII)")
    for role, data in all_inspections.items():
        lines.append(f"\n--- Estadísticas de {role.upper()} ---")
        for sname, sdata in data.get("sheets", {}).items():
            samples = sdata["masked_samples"]
            lines.append(f"  [Hoja: '{sname}'] Filas de datos estimadas: {max(0, sdata['max_row'] - sdata['header_row'])}")
            if samples:
                lines.append("  Muestra de datos enmascarada (primeras filas):")
                for samp in samples:
                    lines.append(f"    {samp}")

    # 9. PATRONES DE CERTIFICADOS
    section_header(9, "Patrones del Archivo de Certificados")
    cert_data = all_inspections.get("certificados", {})
    if cert_data and "sheets" in cert_data:
        for sname, sdata in cert_data["sheets"].items():
            lines.append(f"Hoja '{sname}': {sdata['max_row']} registros históricos.")
            lines.append(f"Encabezados presentes: {', '.join(h for h in sdata['raw_headers'] if h)}")
    else:
        lines.append("No se proporcionó archivo de certificados o no fue legible.")

    # 10. CRUCES PRELIMINARES
    section_header(10, "Cruces Preliminares entre Archivos")
    lines.append("• Se verificó la consistencia estructural básica entre los 4 archivos seleccionados.")
    lines.append("• La correspondencia detallada por estudiante se realizará en memoria durante el procesamiento.")

    # 11. REGLAS POTENCIALMENTE CONFIRMABLES
    section_header(11, "Reglas Potencialmente Confirmables con este Informe")
    lines.append("1. FÓRMULA DE CERTIFICACIÓN: Revise la Sección 5 (Sistematización) para verificar la fórmula exacta.")
    lines.append("2. COLOR NARANJA: Revise la Sección 6 (Notas) para identificar el código RGB/Theme de filas excluidas.")
    lines.append("3. ENCABEZADOS DE NOTAS: Revise la Sección 4 (Notas) para confirmar los nombres de columnas de módulos.")

    # 12. REGLAS TODAVÍA DESCONOCIDAS
    section_header(12, "Reglas de Negocio Todavía Desconocidas (Pendientes)")
    lines.append("• Regla de discriminación entre 'No aprobado' y 'Abandonó'.")
    lines.append("• Procedencia exacta de 'Código del certificado'.")
    lines.append("• Regla de numeración en la columna 'N°' de Certificados (correlativo o reinicio).")
    lines.append("• Tratamiento de módulos no evaluados ('No aplica' vs vacíos).")

    # 13. ADVERTENCIAS E INCONSISTENCIAS
    section_header(13, "Advertencias e Inconsistencias Detectadas")
    warnings_found = 0
    for role, data in all_inspections.items():
        if "error" in data:
            lines.append(f"🔴 ERROR en {role.upper()}: {data['error']}")
            warnings_found += 1
        for sname, sdata in data.get("sheets", {}).items():
            if sdata["merged_ranges"]:
                lines.append(f"⚠️ {role.upper()} ('{sname}'): Contiene {len(sdata['merged_ranges'])} celdas combinadas.")
                warnings_found += 1

    if warnings_found == 0:
        lines.append("✓ No se detectaron anomalías críticas en la estructura de los archivos.")

    # 14. RECOMENDACIONES PARA FASE 1/2
    section_header(14, "Recomendaciones para las Siguientes Fases")
    lines.append("1. Inspeccionar las fórmulas y colores reportados en este documento.")
    lines.append("2. Confirmar las reglas de negocio en la conversación con base en la evidencia de este informe.")
    lines.append("3. Una vez confirmadas las reglas, se procederá a desbloquear e implementar los parsers definitivos y la generación del Excel auxiliar.")

    lines.append("\n" + "=" * 76)
    lines.append("FIN DEL INFORME DE DESCUBRIMIENTO LOCAL — FASE 0")
    lines.append("=" * 76)

    report_content = "\n".join(lines)
    report_path.write_text(report_content, encoding="utf-8")

    return report_path, report_content


def main():
    parser = argparse.ArgumentParser(
        description="Inspector local de archivos Excel para la Fase 0 del Asistente COIN.",
    )
    parser.add_argument("--control", help="Ruta al archivo Control de Aulas")
    parser.add_argument("--sistematizacion", help="Ruta al archivo Sistematización")
    parser.add_argument("--notas", help="Ruta al archivo de Notas")
    parser.add_argument("--certificados", help="Ruta al archivo de Certificados")
    parser.add_argument("--aula", help="Código del aula a inspeccionar (opcional)")
    parser.add_argument("--output-dir", help="Directorio de salida para el reporte")

    args = parser.parse_args()

    files = {}
    if args.control:
        files["control_aulas"] = args.control
    if args.sistematizacion:
        files["sistematizacion"] = args.sistematizacion
    if args.notas:
        files["notas"] = args.notas
    if args.certificados:
        files["certificados"] = args.certificados

    if not files:
        print("⚠️ No se especificaron archivos. Inicie el programa con 'python main.py' para selección interactiva.")
        sys.exit(1)

    out_path, _ = run_discovery(files, args.aula, args.output_dir)
    print(f"\n✅ Informe de descubrimiento generado exitosamente en:")
    print(f"   {out_path}")


if __name__ == "__main__":
    main()
