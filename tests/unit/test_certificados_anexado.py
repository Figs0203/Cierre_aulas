"""
Pruebas de regresión del anexado de Certificados.

Cubren tres fallos detectados en uso real:

1. El punto de anexado se calculaba con la "última fila con algo en A o B".
   Una nota o total suelto debajo de la tabla desplazaba los certificados cientos
   de filas y el usuario no los encontraba (parecía que no se habían escrito).
2. Las filas de plantilla ocultas que ya existían en la hoja recibían el
   certificado pero Excel las seguía mostrando ocultas.
3. Repetir el cierre del mismo curso duplicaba los certificados.
"""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

import openpyxl

from core.models import (
    AulaCierreResult,
    AulaMetadata,
    ClaseGroup,
    EstudianteNotas,
    EstudianteSistematizacion,
    StudentMatch,
)
from core.traceability import TracedValue
from output.direct_updater import (
    _inspect_certificate_table,
    _verify_certificate_rows,
    apply_direct_cierre,
)

_HEADERS_CERT = [
    "N°", "Ciclo", "Docente COIN", "Nombres", "Apellidos",
    "Documento de identidad", "Correo Electrónico", "Código del certificado",
    "Fecha de envío", "Año de certificación", "Semestre", "Total certificados",
    "Fecha de totalización",
]


def _tv(valor):
    return TracedValue(valor, "S", "Formato", 1, 1, 2)


def _crear_archivos(tmpdir: Path, filas_tabla: int = 8) -> dict:
    """Crea los 3 archivos oficiales mínimos (certificados, sistematización y control)."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Códigos"
    ws.append(_HEADERS_CERT)
    for i in range(2, filas_tabla + 2):
        ws.append([
            i - 1, "CLASE X", "(ABBUEI-205) DOC", f"NOM{i}", f"APE{i}",
            f"{100000 + i}", f"x{i}@y.co", "ABBUEI205", "01-ene-2026", 2026,
            "2026-1", 1, None,
        ])
    path_cert = tmpdir / "Control_codigos_certificados_ABBUEI.xlsx"
    wb.save(path_cert)
    wb.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Formato"
    ws.append(["Nombres", "Apellidos", "Clase", "Documento", "Correo"])
    for i in range(2, 8):
        ws.append([f"NOM{i}", f"APE{i}", "6401", f"{200000 + i}", f"y{i}@z.co"])
    path_sist = tmpdir / "Sistematizacion_Cursos_COIN_2026.xlsx"
    wb.save(path_sist)
    wb.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Tabla_ABBUEI"
    ws.append([
        "PROGRAMA ACADÉMICO", "CATÁLOGO", "CLASE", "DOCENTE PRINCIPAL",
        "FECHA DE INICIO DEL CURSO", "FECHA DE FINALIZACIÓN DEL CURSO",
        "CERTIFICADOS Y CIERRE DE CURSO",
    ])
    ws.append(["FINANZAS", "FN01", "6401", "DOCENTE", "16/03/2026", "27/04/2026", None])
    path_ctrl = tmpdir / "Control_Aulas_Formadores_2026.xlsx"
    wb.save(path_ctrl)
    wb.close()

    return {
        "sistematizacion": path_sist,
        "certificados": path_cert,
        "control_aulas": path_ctrl,
        "notas": path_cert,
    }


def _resultado_y_metadatos():
    matches = []
    for i in range(2, 8):
        es = EstudianteSistematizacion(
            row_number=i,
            nombres=_tv(f"NOM{i}"),
            apellidos=_tv(f"APE{i}"),
            clase=_tv("6401"),
            documento=_tv(f"{200000 + i}"),
            correo=_tv(f"y{i}@z.co"),
            estado_calculado="Aprobó",
            column_indices={
                "nombres": 1, "apellidos": 2, "clase": 3, "documento": 4, "correo": 5
            },
            raw_row_data={},
        )
        matches.append(
            StudentMatch(
                estudiante_sist=es,
                estudiante_nota=EstudianteNotas(row_number=i),
                match_key="doc",
                match_value=f"{200000 + i}",
            )
        )

    meta = {
        "6401": AulaMetadata(
            programa_academico=TracedValue("FINANZAS", "C", "Tabla", 1, 1, 2),
            catalogo=TracedValue("FN01", "C", "Tabla", 2, 1, 2),
            docente_principal=TracedValue("DOCENTE", "C", "Tabla", 4, 1, 2),
            fecha_inicio=TracedValue("16/03/2026", "C", "Tabla", 5, 1, 2),
            fecha_fin=TracedValue("27/04/2026", "C", "Tabla", 6, 1, 2),
        )
    }
    result = AulaCierreResult(
        aula=AulaMetadata(),
        clases=[ClaseGroup(clase_id="6401", matches=matches)],
        metadatos_control=meta,
        start_consecutivo=100,
    )
    return result, meta


class TestAnexadoCertificados(unittest.TestCase):
    def test_punto_de_anexado_ignora_notas_debajo_de_la_tabla(self):
        """Una nota suelta lejos de la tabla no debe mover el punto de anexado."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Códigos"
        ws.append(_HEADERS_CERT)
        for i in range(2, 12):
            ws.append([i - 1, "C", "D", f"N{i}", f"A{i}", f"{i}", "x@y.co",
                       "ABBUEI205", "01-ene-2026", 2026, "2026-1", 1, None])
        ws["A90"] = "NOTA: los certificados se relacionan por orden de envío"
        ws["B90"] = "Actualizado por la biblioteca"

        ultima, max_n, uso_respaldo = _inspect_certificate_table(ws)
        wb.close()

        self.assertEqual(ultima, 11)
        self.assertEqual(max_n, 10)
        self.assertFalse(uso_respaldo)

    def test_anexa_al_final_de_la_tabla_y_no_despues_de_la_nota(self):
        with tempfile.TemporaryDirectory() as tmp_str:
            tmpdir = Path(tmp_str)
            files = _crear_archivos(tmpdir)

            # Nota ajena muy por debajo de la tabla
            wb = openpyxl.load_workbook(files["certificados"])
            ws = wb["Códigos"]
            ws["A500"] = "Nota de la biblioteca"
            wb.save(files["certificados"])
            wb.close()

            result, meta = _resultado_y_metadatos()
            report = apply_direct_cierre(
                files, result, "ABBUEI227", metadatos_control=meta,
                start_consecutivo=result.start_consecutivo, backup_dir=tmpdir / "bk",
            )

            self.assertTrue(report.is_successful, report.errors)
            self.assertEqual(report.certificados_rows_appended, 6)
            self.assertEqual(report.certificados_rows_verificados, 6)
            # Tabla de 8 registros (filas 2..9) -> se anexa desde la fila 10
            self.assertEqual(report.certificados_fila_inicial, 10)
            self.assertEqual(report.certificados_fila_final, 15)

            wb = openpyxl.load_workbook(files["certificados"])
            ws = wb["Códigos"]
            filas = [
                r for r in range(2, ws.max_row + 1)
                if str(ws.cell(r, 8).value or "").upper() == "ABBUEI227"
            ]
            wb.close()
            self.assertEqual(filas, list(range(10, 16)))

    def test_fila_de_plantilla_oculta_se_hace_visible(self):
        with tempfile.TemporaryDirectory() as tmp_str:
            tmpdir = Path(tmp_str)
            files = _crear_archivos(tmpdir)

            # Fila 10 oculta (justo donde toca anexar)
            wb = openpyxl.load_workbook(files["certificados"])
            ws = wb["Códigos"]
            ws.row_dimensions[10].hidden = True
            wb.save(files["certificados"])
            wb.close()

            result, meta = _resultado_y_metadatos()
            report = apply_direct_cierre(
                files, result, "ABBUEI227", metadatos_control=meta,
                start_consecutivo=result.start_consecutivo, backup_dir=tmpdir / "bk",
            )

            self.assertTrue(report.is_successful, report.errors)
            self.assertTrue(any("ocultas" in adv for adv in report.warnings))

            wb = openpyxl.load_workbook(files["certificados"])
            ws = wb["Códigos"]
            ocultas = [r for r in range(10, 16) if ws.row_dimensions[r].hidden]
            wb.close()
            self.assertEqual(ocultas, [])

    def test_repetir_el_cierre_no_duplica_certificados(self):
        with tempfile.TemporaryDirectory() as tmp_str:
            tmpdir = Path(tmp_str)
            files = _crear_archivos(tmpdir)
            result, meta = _resultado_y_metadatos()

            primero = apply_direct_cierre(
                files, result, "ABBUEI227", metadatos_control=meta,
                start_consecutivo=result.start_consecutivo, backup_dir=tmpdir / "bk",
            )
            self.assertEqual(primero.certificados_rows_appended, 6)

            segundo = apply_direct_cierre(
                files, result, "ABBUEI227", metadatos_control=meta,
                start_consecutivo=result.start_consecutivo, backup_dir=tmpdir / "bk",
            )
            self.assertTrue(segundo.is_successful, segundo.errors)
            self.assertEqual(segundo.certificados_rows_appended, 0)
            self.assertEqual(segundo.certificados_omitidos_duplicados, 6)

            wb = openpyxl.load_workbook(files["certificados"])
            ws = wb["Códigos"]
            filas = [
                r for r in range(2, ws.max_row + 1)
                if str(ws.cell(r, 8).value or "").upper() == "ABBUEI227"
            ]
            wb.close()
            self.assertEqual(len(filas), 6)

    def test_verificacion_detecta_filas_que_no_quedaron(self):
        with tempfile.TemporaryDirectory() as tmp_str:
            tmpdir = Path(tmp_str)
            files = _crear_archivos(tmpdir)

            confirmadas, fallidas = _verify_certificate_rows(
                files["certificados"], "Códigos",
                [(2, 99, "100002", "ABBUEI227"), (999, 100, "100003", "ABBUEI227")],
            )
            # La fila 2 existe pero con otros datos; la 999 no existe.
            self.assertEqual(confirmadas, 0)
            self.assertEqual(fallidas, [2, 999])


if __name__ == "__main__":
    unittest.main()
