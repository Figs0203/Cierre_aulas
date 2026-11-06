"""
Inspector local del archivo de Certificados (SIN datos personales).

Responde, sobre el archivo real, preguntas como:

  • ¿en qué hoja está la tabla de certificados?
  • ¿cuál es la última fila real de la tabla?
  • ¿existen filas del curso ABBUEI227 y en qué filas?
  • ¿esas filas están ocultas o dentro de celdas combinadas?
  • ¿hay filtros activos?

Uso (equipo con los archivos reales):

    python tests/discovery/inspect_certificados.py ^
        --certificados "Control_codigos_certificados_ABBUEI.xlsx" ^
        --curso ABBUEI227

Privacidad:
  - NUNCA imprime documentos, correos ni nombres completos (se enmascaran).
  - No abre conexiones de red.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

_PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent
if str(_PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(_PROJECT_ROOT))

import openpyxl

from output.direct_updater import _inspect_certificate_table
from security.privacy import mask_document, mask_email, mask_full_name


def _seccion(titulo: str, lineas: list[str]) -> None:
    print("\n" + "=" * 76)
    print(titulo)
    print("=" * 76)
    for linea in lineas:
        print(linea)


def _rangos_combinados(ws) -> list:
    try:
        return list(ws.merged_cells.ranges)
    except Exception:  # noqa: BLE001
        return []


def _rango_combinado_de(ws, fila: int, col: int) -> str | None:
    for rango in _rangos_combinados(ws):
        if rango.min_row <= fila <= rango.max_row and rango.min_col <= col <= rango.max_col:
            return str(rango)
    return None


def inspeccionar(path: Path, curso: str, hoja: str | None = None) -> None:
    wb = openpyxl.load_workbook(path, data_only=True)
    try:
        lineas = [f"Archivo: {path.name}", f"Hojas: {wb.sheetnames}"]
        for nombre in wb.sheetnames:
            ws = wb[nombre]
            ocultas = sum(1 for d in ws.row_dimensions.values() if d.hidden)
            lineas.append(
                f"  • '{nombre}': {ws.max_row} filas × {ws.max_column} columnas "
                f"| filas ocultas: {ocultas}"
            )
        _seccion("ESTRUCTURA DEL ARCHIVO DE CERTIFICADOS", lineas)

        nombre_hoja = hoja
        if not nombre_hoja:
            for candidata in ("Códigos", "Certificados", "CERTIFICADOS"):
                if candidata in wb.sheetnames:
                    nombre_hoja = candidata
                    break
        if not nombre_hoja:
            nombre_hoja = wb.sheetnames[0]

        ws = wb[nombre_hoja]
        ultima, max_n, uso_respaldo = _inspect_certificate_table(ws)

        filtros = []
        try:
            if ws.auto_filter and ws.auto_filter.ref:
                filtros.append(f"rango {ws.auto_filter.ref}")
            if getattr(ws.auto_filter, "filterColumn", None):
                filtros.append(f"{len(ws.auto_filter.filterColumn)} filtro(s) activos")
        except Exception:  # noqa: BLE001
            pass

        combinados_cerca = [
            str(r) for r in _rangos_combinados(ws)
            if r.min_row <= ultima <= r.max_row or r.min_row >= max(1, ultima - 5)
        ]

        _seccion(
            f"TABLA DE CERTIFICADOS — hoja '{nombre_hoja}'",
            [
                f"Última fila de la tabla (registro real): {ultima or '(sin registros)'}",
                f"Último N° usado: {max_n}",
                f"Detección con respaldo (formato no estándar): {'SÍ' if uso_respaldo else 'no'}",
                f"Filtros: {', '.join(filtros) if filtros else 'ninguno'}",
                f"Combinados junto al final de la tabla: "
                f"{', '.join(combinados_cerca) if combinados_cerca else 'ninguno'}",
            ],
        )

        objetivo = str(curso or "").strip().upper()
        coincidencias = []
        for r in range(2, min(ws.max_row or 1, 200_000) + 1):
            codigo = str(ws.cell(r, 8).value or "").strip().upper()
            if codigo != objetivo:
                continue
            coincidencias.append(
                "  • fila {fila}: N°={n} | documento={doc} | nombre={nombre} | "
                "oculta={oculta} | combinado(D/E)={comb} | combinado(1..13)={comb_full}".format(
                    fila=r,
                    n=ws.cell(r, 1).value,
                    doc=mask_document(ws.cell(r, 6).value),
                    nombre=mask_full_name(ws.cell(r, 4).value, ws.cell(r, 5).value),
                    oculta=bool(ws.row_dimensions[r].hidden),
                    comb=_rango_combinado_de(ws, r, 4) or _rango_combinado_de(ws, r, 5) or "-",
                    comb_full=next(
                        (
                            str(rango) for rango in _rangos_combinados(ws)
                            if rango.min_row <= r <= rango.max_row
                            and rango.min_col <= 1 and rango.max_col >= 13
                        ),
                        "-",
                    ),
                )
            )

        resumen = [
            f"Curso buscado: {objetivo}",
            f"Coincidencias en la columna H ('Código del certificado'): {len(coincidencias)}",
        ]
        if coincidencias:
            resumen.append("")
            resumen.extend(coincidencias)
            resumen.append("")
            resumen.append(
                "Si las filas existen pero no se ven en Excel, revise filtros activos, "
                "filas ocultas/agrupadas o que esté abriendo otra copia del archivo."
            )
        else:
            resumen.append("")
            resumen.append(
                "No hay certificados de ese curso en esta hoja. Compruebe que el archivo "
                "y la hoja sean los correctos (y que la copia no sea un respaldo o una "
                "versión anterior)."
            )
        _seccion("CERTIFICADOS DEL CURSO (datos enmascarados)", resumen)
    finally:
        wb.close()


def main() -> None:
    parser = argparse.ArgumentParser(description="Inspector local de Certificados (sin PII)")
    parser.add_argument("--certificados", required=True)
    parser.add_argument("--curso", required=True, help="Ej. ABBUEI227")
    parser.add_argument("--hoja", required=False, help="Hoja a inspeccionar (por defecto Códigos)")
    args = parser.parse_args()

    path = Path(args.certificados)
    if not path.exists():
        print(f"No existe el archivo: {path}")
        return
    inspeccionar(path, args.curso, args.hoja)
    print("\n" + "=" * 76)
    print("FIN DEL INSPECTOR. No se imprimieron datos personales.")
    print("=" * 76)


if __name__ == "__main__":
    main()
