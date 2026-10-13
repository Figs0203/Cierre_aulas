"""
Diagnóstico de Cruce de Identidades (SIN exponer datos personales)
Asistente de Cierre de Aulas COIN — 100% LOCAL y OFFLINE.

Uso (en el equipo con los archivos reales):

    python tests/discovery/diagnose_matching.py ^
        --control "Control_Aulas_Formadores_2026_Programacion.xlsx" ^
        --sistematizacion "Sistematizacion_Cursos_COIN_2026.xlsx" ^
        --notas "CMI120_ECONOMIA_2.xlsx" ^
        --certificados "Control_codigos_certificados_CMI.xlsx" ^
        --aula "CMI120"

Qué hace:
  1. Muestra la ESTRUCTURA de cada archivo (hojas, encabezados, dimensiones).
  2. Muestra qué columnas detecta el programa (identidad, módulos, nota final).
  3. Muestra los rellenos de color por columna (para detectar semáforos).
  4. Hace un SONDEO de cruce Notas <-> Sistematización comparando documento,
     correo y nombre normalizados, y reporta cuántos coinciden.

Privacidad:
  - NUNCA imprime documentos, correos ni nombres completos.
  - Todo dato personal se enmascara con security/privacy.py.
  - No realiza ninguna conexión de red.
"""

from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path

# Agregar el directorio raíz al path para imports
_PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
if str(_PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(_PROJECT_ROOT))

import openpyxl

from config.colors import is_orange_fill_exact
from normalization.headers import (
    analyze_headers,
    find_final_grade_column,
    find_module_columns,
)
from normalization.text import (
    normalize_document,
    normalize_email,
    normalize_name_component,
)
from parsing.control_parser import parse_control_aulas
from parsing.notas_parser import parse_notas_archivo
from parsing.sistematizacion_parser import parse_sistematizacion_aula
from security.privacy import mask_document, mask_email, mask_full_name

_SAMPLE_N = 6


def _section(title: str, lines: list[str]) -> None:
    print("\n" + "=" * 76)
    print(title)
    print("=" * 76)
    for ln in lines:
        print(ln)


def _open(path: Path):
    return openpyxl.load_workbook(path, data_only=False, read_only=False)


def _headers_of(ws, max_col: int | None = None) -> list[str]:
    n = max_col or (ws.max_column or 1)
    out = []
    for c in range(1, n + 1):
        v = ws.cell(1, c).value
        out.append(str(v).strip() if v is not None else f"<vacío:{c}>")
    return out


def _col_letter(idx: int) -> str:
    from openpyxl.utils import get_column_letter

    return get_column_letter(idx)


def inspect_structure(path: Path, label: str) -> None:
    wb = _open(path)
    lines = [f"Archivo: {path.name}", f"Hojas ({len(wb.sheetnames)}): {wb.sheetnames}"]
    for name in wb.sheetnames:
        ws = wb[name]
        merged = len(ws.merged_cells.ranges)
        hidden_rows = [r for r, d in ws.row_dimensions.items() if d.hidden]
        hidden_cols = [c for c, d in ws.column_dimensions.items() if d.hidden]
        lines.append(
            f"  • Hoja '{name}': {ws.max_row} filas × {ws.max_column} cols | "
            f"combinadas={merged} | filas_ocultas={len(hidden_rows)} | cols_ocultas={hidden_cols}"
        )
        lines.append(f"      Encabezados (fila 1): {_headers_of(ws)}")
    wb.close()
    _section(f"ESTRUCTURA — {label}", lines)


def inspect_notas(path: Path) -> None:
    wb = _open(path)
    # Hoja que elegiría el parser
    if "Calificaciones" in wb.sheetnames:
        chosen = "Calificaciones"
    elif "Grades" in wb.sheetnames:
        chosen = "Grades"
    else:
        chosen = wb.sheetnames[0]
    ws = wb[chosen]

    headers = _headers_of(ws)
    analysis = analyze_headers(headers)
    mod_cols = find_module_columns(headers)
    final_col = find_final_grade_column(headers)

    lines = [f"Hoja usada por el parser: '{chosen}' ({ws.max_row} filas × {ws.max_column} cols)", ""]
    lines.append("Mapeo de encabezados detectado:")
    for m in analysis.mapped:
        lines.append(f"  • '{m.original}' (col {m.col_idx} = {_col_letter(m.col_idx)}) → {m.canonical} [{m.matched_via}]")
    if analysis.unmapped:
        for name, idx in analysis.unmapped:
            lines.append(f"  • '{name}' (col {idx} = {_col_letter(idx)}) → SIN MAPEAR ⚠️")
    lines.append(f"Módulos detectados: {mod_cols}")
    lines.append(f"Nota final detectada en col: {final_col}")

    # Identidad: conteos y ejemplos enmascarados
    id_cols = {
        "OrgDefinedId": next((m.col_idx for m in analysis.mapped if m.canonical == "org_defined_id"), 1),
        "Username": next((m.col_idx for m in analysis.mapped if m.canonical == "username"), 2),
        "Last Name": next((m.col_idx for m in analysis.mapped if m.canonical == "apellidos"), 3),
        "First Name": next((m.col_idx for m in analysis.mapped if m.canonical == "nombres"), 4),
    }
    lines.append("")
    lines.append("Columnas de identidad:")
    for label, c in id_cols.items():
        vals = []
        for r in range(2, (ws.max_row or 1) + 1):
            v = ws.cell(r, c).value
            if v is not None and str(v).strip() != "":
                vals.append(str(v))
        formulas = sum(1 for v in vals if v.startswith("="))
        lines.append(f"  • {label} (col {c} = {_col_letter(c)}): no vacíos={len(vals)} | fórmulas={formulas}")
        lines.append(f"      ejemplos: {[mask_document(v) for v in vals[:_SAMPLE_N]]}")

    # Rellenos por columna (para detectar semáforos / exclusiones)
    fills_by_col: dict[int, dict[str, int]] = {}
    excluded_rows = 0
    for r in range(2, (ws.max_row or 1) + 1):
        row_orange = False
        for c in range(1, (ws.max_column or 1) + 1):
            cell = ws.cell(r, c)
            f = cell.fill
            if not f or not f.fill_type or f.fill_type == "none":
                continue
            fg = getattr(f, "fgColor", None)
            if not fg:
                continue
            rgb = getattr(fg, "rgb", None)
            rgb_s = None
            if isinstance(rgb, str):
                s = rgb.strip().lstrip("#")
                rgb_s = s[2:] if len(s) == 8 else s
            theme = getattr(fg, "theme", None)
            theme = theme if isinstance(theme, int) else None
            tint = getattr(fg, "tint", None)
            tint = tint if isinstance(tint, (int, float)) else None
            key = f"rgb={rgb_s} theme={theme} tint={tint}"
            fills_by_col.setdefault(c, {})
            fills_by_col[c][key] = fills_by_col[c].get(key, 0) + 1
            if is_orange_fill_exact(rgb_s, theme, tint):
                row_orange = True
        if row_orange:
            excluded_rows += 1

    lines.append("")
    lines.append("Rellenos de color por columna (para detectar semáforos/exclusiones):")
    for c in sorted(fills_by_col):
        header = headers[c - 1] if c - 1 < len(headers) else "?"
        lines.append(f"  • col {c} = {_col_letter(c)} ('{header}'):")
        for key, cnt in fills_by_col[c].items():
            lines.append(f"      {key}  × {cnt}")
    lines.append(f"Filas que el detector marcaría como EXCLUIDAS (naranja): {excluded_rows}")
    wb.close()
    _section("NOTAS — COLUMNAS Y COLORES DETECTADOS", lines)

    # Lo que realmente extrae el parser
    try:
        estudiantes, secciones = parse_notas_archivo(path)
    except Exception as e:  # noqa: BLE001
        _section("NOTAS — PARSER", [f"ERROR al parsear: {e}"])
        return
    lines = [
        f"Estudiantes extraídos por el parser: {len(estudiantes)}",
        f"Secciones detectadas: {len(secciones)}",
    ]
    for es in estudiantes[:_SAMPLE_N]:
        lines.append(
            "  " + " | ".join([
                f"doc={mask_document(str(es.org_defined_id.value) if es.org_defined_id else '')}",
                f"user={mask_email(str(es.username.value) if es.username else '')}",
                f"nombre={mask_full_name(es.first_name.value if es.first_name else '', es.last_name.value if es.last_name else '')}",
                f"seccion={es.seccion.value if es.seccion else None}",
                f"excluido={es.is_orange_fill}",
            ])
        )
    _section("NOTAS — LO QUE VE EL PARSER (enmascarado)", lines)


def inspect_sistematizacion(path: Path, aula: str) -> object:
    wb = _open(path)
    sheet = "Formato" if "Formato" in wb.sheetnames else wb.sheetnames[0]
    ws = wb[sheet]
    headers = _headers_of(ws)
    lines = [
        f"Hoja: '{sheet}' ({ws.max_row} filas × {ws.max_column} cols)",
        f"Hoja 'Formato' presente: {'Formato' in wb.sheetnames}",
    ]
    # localizar columnas por encabezado (mismo criterio que el parser)
    def find(pred) -> int | None:
        for c, h in enumerate(headers, start=1):
            hl = h.lower()
            if pred(hl):
                return c
        return None

    cols = {
        "Aula Virtual / Solicitud": find(lambda h: "aula virtual" in h or "solicitud" in h),
        "Nombres": find(lambda h: h == "nombres"),
        "Apellidos": find(lambda h: h == "apellidos"),
        "Documento": find(lambda h: "número de documento" in h or "numero de documento" in h),
        "Correo": find(lambda h: "correo" in h),
        "Clase": find(lambda h: h == "clase"),
        "Módulo 1": find(lambda h: "módulo 1" in h or "modulo 1" in h),
        "Promedio": find(lambda h: "promedio" in h),
        "Estado": find(lambda h: "estado para certific" in h),
    }
    lines.append("Columnas detectadas por encabezado:")
    for k, v in cols.items():
        lines.append(f"  • {k}: {'col ' + str(v) + ' (' + _col_letter(v) + ')' if v else 'NO ENCONTRADA ⚠️'}")
    wb.close()
    _section("SISTEMATIZACIÓN — ESTRUCTURA", lines)

    try:
        por_clase, headers2 = parse_sistematizacion_aula(path, target_aula=aula)
    except Exception as e:  # noqa: BLE001
        _section("SISTEMATIZACIÓN — PARSER", [f"ERROR al parsear: {e}"])
        return None
    lines = [f"Clases detectadas para el aula '{aula}': {list(por_clase.keys())}"]
    total = sum(len(v) for v in por_clase.values())
    lines.append(f"Total estudiantes en Sistematización: {total}")
    for clase_id, ests in por_clase.items():
        lines.append(f"  [Clase {clase_id}] {len(ests)} estudiantes")
        for es in ests[:_SAMPLE_N]:
            lines.append(
                "    " + " | ".join([
                    f"doc={mask_document(str(es.documento.value) if es.documento else '')}",
                    f"correo={mask_email(str(es.correo.value) if es.correo else '')}",
                    f"nombre={mask_full_name(es.nombres.value if es.nombres else '', es.apellidos.value if es.apellidos else '')}",
                ])
            )
    _section("SISTEMATIZACIÓN — LO QUE VE EL PARSER (enmascarado)", lines)
    return por_clase


def inspect_control(path: Path, aula: str) -> None:
    wb = _open(path)
    lines = []
    tablas = [s for s in wb.sheetnames if s.startswith("Tabla_") or s == "CONTROL"]
    lines.append(f"Hojas Tabla_*/CONTROL: {tablas}")
    found_any = False
    for name in tablas:
        ws = wb[name]
        hits = []
        for r in range(1, (ws.max_row or 1) + 1):
            for c in range(1, (ws.max_column or 1) + 1):
                v = ws.cell(r, c).value
                if v is not None and aula.upper() in str(v).upper():
                    hits.append((r, c))
                    break
            if len(hits) >= 3:
                break
        if hits:
            found_any = True
            lines.append(f"  • '{name}': aula '{aula}' encontrada en filas {[h[0] for h in hits]}")
    if not found_any:
        lines.append(f"  ⚠️ El aula '{aula}' NO se encontró en ninguna hoja Tabla_*.")
    wb.close()

    try:
        metas = parse_control_aulas(path, target_aula=aula)
        lines.append(f"Metadatos de Control extraídos: {list(metas.keys())}")
        for k, m in metas.items():
            prog = m.programa_academico.value if m.programa_academico else None
            cat = m.catalogo.value if m.catalogo else None
            lines.append(f"  [Clase {k}] catálogo={cat} | programa={prog}")
    except Exception as e:  # noqa: BLE001
        lines.append(f"ERROR al parsear Control: {e}")
    _section("CONTROL DE AULAS — UBICACIÓN DEL AULA", lines)


def cross_probe(notas_path: Path, sist_path: Path, aula: str) -> None:
    """Sondeo de cruce replicando la normalización del motor de matching."""
    estudiantes_notas, _ = parse_notas_archivo(notas_path)
    por_clase, _ = parse_sistematizacion_aula(sist_path, target_aula=aula)
    ests_sist = [es for v in por_clase.values() for es in v]

    def norm_doc(v) -> str:
        return normalize_document(str(v)) if v not in (None, "") else ""

    def norm_mail(v) -> str:
        return normalize_email(str(v)) if v not in (None, "") else ""

    def norm_name(first, last) -> str:
        return f"{normalize_name_component(str(first or ''))} {normalize_name_component(str(last or ''))}".strip()

    notas_docs = [norm_doc(en.org_defined_id.value) if en.org_defined_id else "" for en in estudiantes_notas]
    notas_mails = [norm_mail(en.username.value) if en.username else "" for en in estudiantes_notas]
    notas_names = [norm_name(en.first_name.value if en.first_name else "", en.last_name.value if en.last_name else "") for en in estudiantes_notas]

    sist_docs = [norm_doc(es.documento.value) if es.documento else "" for es in ests_sist]
    sist_mails = [norm_mail(es.correo.value) if es.correo else "" for es in ests_sist]
    sist_names = [norm_name(es.nombres.value if es.nombres else "", es.apellidos.value if es.apellidos else "") for es in ests_sist]

    def non_empty(lst):
        return [x for x in lst if x]

    def inter(a, b):
        return len(set(non_empty(a)) & set(non_empty(b)))

    lines = [
        f"Notas: {len(estudiantes_notas)} estudiantes | Sistematización: {len(ests_sist)} estudiantes",
        "",
        "Conteo de identificadores NO vacíos:",
        f"  • Documento  → Notas: {len(non_empty(notas_docs))} | Sist: {len(non_empty(sist_docs))}",
        f"  • Correo     → Notas: {len(non_empty(notas_mails))} | Sist: {len(non_empty(sist_mails))}",
        f"  • Nombre     → Notas: {len(non_empty(notas_names))} | Sist: {len(non_empty(sist_names))}",
        "",
        "Coincidencias (normalizadas):",
        f"  • Documento: {inter(notas_docs, sist_docs)}",
        f"  • Correo:    {inter(notas_mails, sist_mails)}",
        f"  • Nombre:    {inter(notas_names, sist_names)}",
        "",
        f"Muestras Notas (enmascaradas, primeras {_SAMPLE_N}):",
    ]
    for i in range(min(_SAMPLE_N, len(estudiantes_notas))):
        en = estudiantes_notas[i]
        lines.append(
            "  " + " | ".join([
                f"doc={mask_document(str(en.org_defined_id.value) if en.org_defined_id else '')}→{notas_docs[i][:3]}***",
                f"correo={mask_email(str(en.username.value) if en.username else '')}",
                f"nombre={mask_full_name(en.first_name.value if en.first_name else '', en.last_name.value if en.last_name else '')}",
            ])
        )
    lines.append("")
    lines.append(f"Muestras Sistematización (enmascaradas, primeras {_SAMPLE_N}):")
    for i in range(min(_SAMPLE_N, len(ests_sist))):
        es = ests_sist[i]
        lines.append(
            "  " + " | ".join([
                f"doc={mask_document(str(es.documento.value) if es.documento else '')}→{sist_docs[i][:3]}***",
                f"correo={mask_email(str(es.correo.value) if es.correo else '')}",
                f"nombre={mask_full_name(es.nombres.value if es.nombres else '', es.apellidos.value if es.apellidos else '')}",
            ])
        )
    _section("SONDEO DE CRUCE NOTAS ↔ SISTEMATIZACIÓN", lines)


def main() -> None:
    parser = argparse.ArgumentParser(description="Diagnóstico de cruce (sin datos personales)")
    parser.add_argument("--control", required=True)
    parser.add_argument("--sistematizacion", required=True)
    parser.add_argument("--notas", required=True)
    parser.add_argument("--certificados", required=False)
    parser.add_argument("--aula", required=True)
    args = parser.parse_args()

    control = Path(args.control)
    sist = Path(args.sistematizacion)
    notas = Path(args.notas)
    cert = Path(args.certificados) if args.certificados else None

    print("=" * 76)
    print("DIAGNÓSTICO DE CRUCE — SIN DATOS PERSONALES (100% LOCAL)")
    print(f"Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"Aula objetivo: {args.aula}")
    print("=" * 76)

    inspect_structure(notas, "NOTAS")
    inspect_notas(notas)
    inspect_sistematizacion(sist, args.aula)
    inspect_control(control, args.aula)
    if cert and cert.exists():
        inspect_structure(cert, "CERTIFICADOS")
    cross_probe(notas, sist, args.aula)

    print("\n" + "=" * 76)
    print("FIN DEL DIAGNÓSTICO. Puede compartir esta salida: NO contiene datos personales.")
    print("=" * 76)


if __name__ == "__main__":
    main()
