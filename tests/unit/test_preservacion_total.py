"""
Prueba de "pérdida cero" sobre los archivos oficiales.

Regresión reportada: al aplicar el cierre directo (Modo 5) se perdían las listas
de validación de datos y otras funciones de Excel de los archivos originales de
la biblioteca, porque `openpyxl` no es un editor *round-trip* fiel: al guardar
descarta todo lo que no sabe representar (validaciones `x14`, `extLst`, controles
de formulario, tablas dinámicas, parte del formato condicional, etc.).

El cierre ahora NO vuelve a escribir el libro: `output/xlsx_patcher.py` modifica
solo las celdas necesarias y copia el resto del paquete byte a byte. Estas pruebas
lo verifican parte por parte, incluso sobre hojas que SÍ se modifican.
"""

from __future__ import annotations

import base64
import io
import re
import tempfile
import unittest
import zipfile
from pathlib import Path
from xml.parsers import expat

import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.comments import Comment
from openpyxl.drawing.image import Image as XLImage
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Font, PatternFill
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

from core.models import (
    AulaCierreResult,
    AulaMetadata,
    ClaseGroup,
    EstudianteNotas,
    EstudianteSistematizacion,
    StudentMatch,
)
from core.traceability import TracedValue
from output.direct_updater import apply_direct_cierre
from output.xlsx_patcher import XlsxPatcher, verify_preservation

# PNG 1x1 válido (para probar que las imágenes sobreviven al cierre).
_PNG_1PX = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
)

_SHEET1 = "xl/worksheets/sheet1.xml"

_X14_EXT = (
    '<extLst><ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" '
    'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" '
    'xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">'
    '<x14:dataValidations count="1">'
    '<x14:dataValidation type="list" allowBlank="1">'
    '<x14:formula1><xm:f>Formato!$A$1:$A$3</xm:f></x14:formula1>'
    '<xm:sqref>J2:J500</xm:sqref></x14:dataValidation>'
    '</x14:dataValidations></ext></extLst>'
)

# Partes ajenas a openpyxl: nunca se leen, pero tampoco deben desaparecer.
_EXTRA_PARTS = {
    "xl/pivotTables/pivotTable1.xml": b"<pivotTableDefinition/>",
    "xl/pivotCache/pivotCacheDefinition1.xml": b"<pivotCacheDefinition/>",
    "xl/charts/chart9.xml": b"<chartSpace/>",
    "xl/media/logo_biblioteca.png": b"\x89PNG\r\n\x1a\n",
    "xl/worksheets/vmlDrawing1.vml": b"<xml/>",
    "customXml/item1.xml": b"<custom/>",
}


def _rewrite_zip(path: Path, transform) -> None:
    """Aplica una transformación a las partes del .xlsx conservando el resto."""
    with zipfile.ZipFile(path) as zf:
        members = {name: zf.read(name) for name in zf.namelist()}
    transform(members)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in members.items():
            zf.writestr(name, data)


def _build_rich_files(tmpdir: Path) -> dict:
    """Crea archivos oficiales sintéticos con TODAS las funciones frágiles."""
    # ---------------- Sistematización ----------------
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Formato"
    ws.append([
        "ID", "Nombres", "Apellidos", "Clase", "Documento", "Correo",
        "Módulo 1", "Módulo 2", "Módulo 3", "Módulo 4", "Módulo 5",
        "Promedio curso completo", "Estado", "Se elaboró certificado",
        "Código del certificado", "Se envió certificado", "Fecha de envío",
    ])
    ws.append([
        1, "carlos andres", "perez gomez", "5535", "1001", "cperez@eafit.edu.co",
        4.5, None, None, None, None, None, None, None, None, None, None,
    ])
    ws.append([
        2, "Ana Maria", "Lopez Duque", "5535", "1002", "alopez@eafit.edu.co",
        None, None, None, None, None, None, None, None, None, None, None,
    ])
    # Fórmula real en una fila que el cierre SÍ toca (bordes): debe sobrevivir.
    ws["L3"] = "=SUM(G3:K3)"
    # Funciones de Excel que hoy se perderían:
    dv = DataValidation(type="list", formula1='"SI,NO"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("M2:M500")
    ws.conditional_formatting.add(
        "L2:L500",
        ColorScaleRule(start_type="min", start_color="FF0000",
                       end_type="max", end_color="00FF00"),
    )
    ws["A1"].comment = Comment("Encabezado oficial", "Biblioteca")
    ws["A1"].font = Font(name="Zurich Cn BT", size=11, bold=True)
    ws["A1"].fill = PatternFill("solid", fgColor="FFCC00")
    ws.row_dimensions[1].height = 28
    ws.column_dimensions["B"].width = 24
    ws.freeze_panes = "C2"
    image = XLImage(io.BytesIO(_PNG_1PX))
    ws.add_image(image, "S1")
    wb.defined_names.add(
        DefinedName("ListaEstados", attr_text="Formato!$M$1:$M$2")
    )

    # Hoja adicional con un gráfico que apunta a la hoja modificada.
    ws_chart = wb.create_sheet("Resumen")
    for i in range(1, 6):
        ws_chart.cell(i, 1, i)
        ws_chart.cell(i, 2, i * i)
    chart = BarChart()
    chart.add_data(Reference(ws_chart, min_col=2, min_row=1, max_row=5))
    ws_chart.add_chart(chart, "D2")

    path_sist = tmpdir / "Sistematizacion_Cursos_COIN_2026.xlsx"
    wb.save(path_sist)
    wb.close()

    # Inyectar validaciones modernas (x14) y partes que openpyxl desconoce.
    def _transform(members: dict) -> None:
        xml = members[_SHEET1].decode("utf-8")
        members[_SHEET1] = xml.replace(
            "</worksheet>", _X14_EXT + "</worksheet>"
        ).encode("utf-8")
        for name, data in _EXTRA_PARTS.items():
            if name not in members:
                members[name] = data

    _rewrite_zip(path_sist, _transform)

    # ---------------- Certificados ----------------
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Códigos"
    ws.append([
        "N°", "Ciclo", "Docente COIN", "Nombres", "Apellidos",
        "Documento de identidad", "Correo Electrónico", "Código del certificado",
        "Fecha de envío", "Año de certificación", "Semestre", "Total certificados",
        "Fecha de totalización",
    ])
    ws.append([
        10, "CLASE PREVIA", "DOCENTE PREVIO", "PREVIO", "PREVIO",
        "9999", "previo@eafit.edu.co", "ABBUEI205", "15-mar-2026",
        2026, "2026-1", 1, None,
    ])
    dv_cert = DataValidation(type="list", formula1='"SI,NO"', allow_blank=True)
    ws.add_data_validation(dv_cert)
    dv_cert.add("N2:N900")
    path_cert = tmpdir / "Control_codigos_certificados_ABBUEI.xlsx"
    wb.save(path_cert)
    wb.close()

    # ---------------- Control de Aulas ----------------
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Tabla_ABBUEI"
    ws.append([
        "PROGRAMA ACADÉMICO", "CATÁLOGO", "CLASE", "DOCENTE PRINCIPAL",
        "FECHA DE INICIO DEL CURSO", "FECHA DE FINALIZACIÓN DEL CURSO",
        "CERTIFICADOS",
    ])
    ws.append(["ADMINISTRACION", "AD01", "5535", "Manuela Restrepo",
               "16/03/2026", "27/04/2026", None])
    dv_ctrl = DataValidation(type="list", formula1='"SI,NO"', allow_blank=True)
    ws.add_data_validation(dv_ctrl)
    dv_ctrl.add("H2:H50")
    path_ctrl = tmpdir / "Control_Aulas_Formadores_2026.xlsx"
    wb.save(path_ctrl)
    wb.close()

    # ---------------- Notas (solo lectura) ----------------
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["OrgDefinedId", "Username", "Last Name", "First Name", "Calculated Final"])
    ws.append(["1001", "cperez", "perez", "carlos", 4.8])
    ws.append(["1002", "alopez", "lopez", "ana", 4.2])
    path_notas = tmpdir / "Notas_ABBUEI205.xlsx"
    wb.save(path_notas)
    wb.close()

    return {
        "sistematizacion": path_sist,
        "certificados": path_cert,
        "control_aulas": path_ctrl,
        "notas": path_notas,
    }


def _build_result() -> AulaCierreResult:
    col_map = {
        "nombres": 2, "apellidos": 3, "clase": 4, "documento": 5, "correo": 6,
        "modulo_1": 7, "modulo_2": 8, "promedio": 12, "estado_certificacion": 13,
        "elaboro_certificado": 14, "codigo_certificado": 15,
        "envio_certificado": 16, "fecha_envio": 17,
    }

    def tv(val, col, row):
        return TracedValue(val, "S", "Formato", col, 1, row)

    matches = []
    for row, (nom, ape, doc, correo, nota) in enumerate(
        [("carlos andres", "perez gomez", "1001", "cperez@eafit.edu.co", 4.8),
         ("Ana Maria", "Lopez Duque", "1002", "alopez@eafit.edu.co", 4.2)],
        start=2,
    ):
        es = EstudianteSistematizacion(
            row_number=row,
            nombres=tv(nom, "Nombres", row),
            apellidos=tv(ape, "Apellidos", row),
            clase=tv("5535", "Clase", row),
            documento=tv(doc, "Documento", row),
            correo=tv(correo, "Correo", row),
            estado_calculado="Aprobó",
            column_indices=col_map,
            raw_row_data={
                7: nota, 8: nota, 12: nota, 13: "Aprobó", 14: "SI",
                15: "ABBUEI205", 16: "SI", 17: "27-abr-2026",
            },
        )
        matches.append(
            StudentMatch(
                estudiante_sist=es,
                estudiante_nota=EstudianteNotas(row_number=row),
                match_key="doc",
                match_value=doc,
            )
        )
    return AulaCierreResult(
        aula=AulaMetadata(), clases=[ClaseGroup(clase_id="5535", matches=matches)]
    )


def _meta() -> dict:
    def tv(val, col, row):
        return TracedValue(val, "C", "Tabla", col, 1, row)

    return {
        "5535": AulaMetadata(
            programa_academico=tv("ADMINISTRACION", "P", 2),
            catalogo=tv("AD01", "C", 2),
            docente_principal=tv("Manuela Restrepo", "D", 2),
            fecha_inicio=tv("16/03/2026", "FI", 2),
            fecha_fin=tv("27/04/2026", "FF", 2),
        )
    }


def _snapshot(path: Path) -> dict:
    with zipfile.ZipFile(path) as zf:
        return {name: zf.read(name) for name in zf.namelist()}


def _assert_xml_valido(data: bytes, name: str) -> None:
    parser = expat.ParserCreate(namespace_separator="|")
    try:
        parser.Parse(data, True)
    except expat.ExpatError as exc:  # pragma: no cover - mensaje de diagnóstico
        raise AssertionError(f"XML inválido en {name}: {exc}") from exc


class TestPreservacionTotal(unittest.TestCase):
    """El cierre no puede perder ninguna parte de los libros oficiales."""

    def _apply(self, tmpdir: Path):
        files = _build_rich_files(tmpdir)
        before = {role: _snapshot(path) for role, path in files.items()}
        report = apply_direct_cierre(
            file_paths=files,
            result=_build_result(),
            codigo_curso="ABBUEI205",
            metadatos_control=_meta(),
            start_consecutivo=1,
        )
        return files, before, report

    def test_ninguna_parte_del_libro_se_pierde(self):
        with tempfile.TemporaryDirectory() as tmp_str:
            tmpdir = Path(tmp_str)
            files, before, report = self._apply(tmpdir)

            self.assertTrue(report.is_successful, report.errors)
            self.assertEqual(report.warnings, [], report.warnings)

            for role in ("sistematizacion", "certificados", "control_aulas"):
                after = _snapshot(files[role])
                original = before[role]

                # 1. Ni una sola parte del paquete puede desaparecer.
                faltantes = sorted(set(original) - set(after))
                self.assertEqual(
                    faltantes, [], f"{role}: se perdieron partes del libro: {faltantes}"
                )

                # 2. Todo lo que no es hoja/estilos/libro queda byte a byte igual.
                permitidas = {"xl/styles.xml", "xl/workbook.xml"}
                for name, data in original.items():
                    if name in permitidas or name.startswith("xl/worksheets/sheet"):
                        continue
                    self.assertEqual(
                        data, after[name],
                        f"{role}: la parte '{name}' se modificó y no debía",
                    )

                # 3. Todo el XML resultante es válido (Excel no pedirá reparar).
                for name, data in after.items():
                    if name.endswith(".xml") or name.endswith(".rels") or name.endswith(".vml"):
                        _assert_xml_valido(data, f"{role}:{name}")

                with zipfile.ZipFile(files[role]) as zf:
                    self.assertIsNone(zf.testzip())

    def test_se_conservan_las_funciones_de_excel(self):
        with tempfile.TemporaryDirectory() as tmp_str:
            tmpdir = Path(tmp_str)
            files, before, report = self._apply(tmpdir)

            sist = _snapshot(files["sistematizacion"])

            # Listas de validación estándar y modernas (x14) intactas.
            sheet1 = sist[_SHEET1].decode("utf-8", "ignore")
            self.assertIn('<dataValidation ', sheet1)
            self.assertIn("M2:M500", sheet1)
            self.assertIn("x14:dataValidations", sheet1)
            self.assertIn("J2:J500", sheet1)
            self.assertEqual(sheet1.count("<extLst"), 1)

            # Formato condicional, comentario, imagen, gráfico y nombres definidos.
            self.assertIn("<cfRule", sheet1)
            self.assertIn("<legacyDrawing", sheet1)
            self.assertIn("<drawing", sheet1)
            self.assertIn("ListaEstados", sist["xl/workbook.xml"].decode("utf-8", "ignore"))
            self.assertTrue(any(n.startswith("xl/charts/chart") for n in sist))
            self.assertTrue(any(n.startswith("xl/media/") for n in sist))
            self.assertTrue(
                any(n.endswith(".vml") for n in sist), "se perdió el VML (comentarios/controles)"
            )
            # Tablas dinámicas y customXml inyectados.
            self.assertIn("xl/pivotTables/pivotTable1.xml", sist)
            self.assertIn("customXml/item1.xml", sist)
            # Ancho de columna, alto de fila, panel congelado y relleno del encabezado.
            self.assertIn('width="24"', sheet1)
            self.assertIn('ht="28"', sheet1)
            self.assertIn("pane", sheet1)

            # La fórmula de una fila modificada no se sustituye por un valor y el
            # libro queda marcado para recalcular al abrir.
            self.assertIn("<f>SUM(G3:K3)</f>", sheet1)
            self.assertIn(
                'fullCalcOnLoad="1"', sist["xl/workbook.xml"].decode("utf-8", "ignore")
            )
            wb = openpyxl.load_workbook(files["sistematizacion"])
            self.assertEqual(wb["Formato"]["L3"].value, "=SUM(G3:K3)")
            wb.close()

            # En Certificados y Control también sobreviven sus listas.
            cert = _snapshot(files["certificados"])
            self.assertIn(
                "N2:N900",
                cert["xl/worksheets/sheet1.xml"].decode("utf-8", "ignore"),
            )
            ctrl = _snapshot(files["control_aulas"])
            self.assertIn(
                "H2:H50",
                ctrl["xl/worksheets/sheet1.xml"].decode("utf-8", "ignore"),
            )

            # El veredicto del Modo 5 es positivo para los tres archivos.
            self.assertEqual(len(report.validaciones_preservadas), 3)
            for resumen in report.validaciones_preservadas.values():
                self.assertIn("✅", resumen)
                self.assertIn("conservadas", resumen)

    def test_los_cambios_del_cierre_si_se_aplican(self):
        with tempfile.TemporaryDirectory() as tmp_str:
            tmpdir = Path(tmp_str)
            files, _before, report = self._apply(tmpdir)

            self.assertTrue(report.is_successful, report.errors)
            self.assertEqual(report.sistematizacion_names_uppercased, 4)
            self.assertEqual(report.certificados_rows_appended, 2)
            self.assertEqual(report.control_aulas_classes_marked, 1)

            wb = openpyxl.load_workbook(files["sistematizacion"])
            ws = wb["Formato"]
            self.assertEqual(ws["B2"].value, "CARLOS ANDRES")
            self.assertEqual(ws["C3"].value, "LOPEZ DUQUE")
            self.assertEqual(ws["G2"].value, 4.5)   # preexistente intacta
            self.assertEqual(ws["H2"].value, 4.8)   # vacía rellenada
            self.assertEqual(ws["M2"].value, "Aprobó")
            self.assertEqual(ws["Q2"].value, "27-abr-2026")
            self.assertEqual(ws["A2"].border.bottom.style, "thin")
            self.assertEqual(ws["A3"].border.bottom.style, "medium")  # fin de clase
            self.assertEqual(ws["A3"].border.left.style, "thin")
            # El formato del encabezado no se tocó.
            self.assertTrue(ws["A1"].font.b)
            self.assertEqual(ws["A1"].fill.fgColor.rgb, "00FFCC00")
            wb.close()

            wb = openpyxl.load_workbook(files["certificados"])
            ws = wb["Códigos"]
            self.assertEqual(ws["A2"].value, 10)          # fila previa intacta
            self.assertEqual(ws["A3"].value, 11)          # consecutivo
            self.assertEqual(ws["D3"].value, "CARLOS ANDRES")
            self.assertEqual(ws["A3"].font.name, "Zurich Cn BT")
            self.assertEqual(ws["A3"].font.size, 11)
            self.assertEqual(ws["A3"].alignment.horizontal, "center")
            wb.close()

    def test_verificacion_detecta_parte_perdida(self):
        """La verificación debe fallar si algo se pierde (red de seguridad)."""
        with tempfile.TemporaryDirectory() as tmp_str:
            tmpdir = Path(tmp_str)
            files = _build_rich_files(tmpdir)
            original = tmpdir / "original.xlsx"
            copia = files["sistematizacion"]
            import shutil

            shutil.copy2(copia, original)

            # Simular una pérdida real: eliminar el gráfico del resultado.
            def _drop(members: dict) -> None:
                for name in [n for n in members if n.startswith("xl/charts/")]:
                    del members[name]

            _rewrite_zip(copia, _drop)

            info = verify_preservation(original, copia)
            self.assertFalse(info.ok)
            self.assertTrue(any("gráficos" in p for p in info.problemas), info.problemas)


class TestXlsxPatcher(unittest.TestCase):
    """Pruebas unitarias del editor quirúrgico de celdas."""

    def _book(self, tmpdir: Path) -> Path:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Hoja"
        ws.append(["A", "B", "C"])
        ws.append([1, "texto", None])
        ws["B2"].font = Font(name="Zurpi", size=14, bold=True)
        ws["B2"].number_format = "@"
        ws.column_dimensions["C"].width = 12
        path = tmpdir / "libro.xlsx"
        wb.save(path)
        wb.close()
        return path

    def test_escribe_valores_y_conserva_formato_de_la_celda(self):
        with tempfile.TemporaryDirectory() as tmp_str:
            tmpdir = Path(tmp_str)
            path = self._book(tmpdir)
            before = _snapshot(path)

            patcher = XlsxPatcher(path)
            sheet = patcher.sheet("Hoja")
            sheet.set_cell(2, 2, "TEXTO", has_value=True)          # texto existente
            sheet.set_cell(2, 3, 4.75, has_value=True)             # celda vacía
            patcher.save()

            wb = openpyxl.load_workbook(path)
            ws = wb["Hoja"]
            self.assertEqual(ws["B2"].value, "TEXTO")
            self.assertEqual(ws["B2"].font.name, "Zurpi")          # se conserva
            self.assertEqual(ws["B2"].font.sz, 14)
            self.assertTrue(ws["B2"].font.b)
            self.assertEqual(ws["B2"].number_format, "@")
            self.assertEqual(ws["C2"].value, 4.75)
            wb.close()

            after = _snapshot(path)
            # sharedStrings no se toca: las cadenas nuevas van como inlineStr.
            if "xl/sharedStrings.xml" in before:
                self.assertEqual(
                    before["xl/sharedStrings.xml"], after["xl/sharedStrings.xml"]
                )

    def test_crea_filas_nuevas_y_actualiza_dimension(self):
        with tempfile.TemporaryDirectory() as tmp_str:
            tmpdir = Path(tmp_str)
            path = self._book(tmpdir)

            patcher = XlsxPatcher(path)
            sheet = patcher.sheet("Hoja")
            for col, value in enumerate([7, "nuevo", 3.5], start=1):
                sheet.set_cell(10, col, value, has_value=True)
            patcher.save()

            wb = openpyxl.load_workbook(path)
            ws = wb["Hoja"]
            self.assertEqual([ws.cell(10, c).value for c in (1, 2, 3)], [7, "nuevo", 3.5])
            wb.close()

            xml = _snapshot(path)[_SHEET1].decode("utf-8")
            self.assertIn('ref="A1:C10"', xml)

    def test_escapa_caracteres_especiales(self):
        with tempfile.TemporaryDirectory() as tmp_str:
            tmpdir = Path(tmp_str)
            path = self._book(tmpdir)

            patcher = XlsxPatcher(path)
            sheet = patcher.sheet("Hoja")
            texto = 'Pérez & Cía. <2026> "Ñ"'
            sheet.set_cell(2, 3, texto, has_value=True)
            patcher.save()

            wb = openpyxl.load_workbook(path)
            self.assertEqual(wb["Hoja"]["C2"].value, texto)
            wb.close()

            self.assertIn("&amp;", _snapshot(path)[_SHEET1].decode("utf-8"))

    def test_sin_cambios_no_reescribe_el_archivo(self):
        with tempfile.TemporaryDirectory() as tmp_str:
            tmpdir = Path(tmp_str)
            path = self._book(tmpdir)
            before_bytes = path.read_bytes()

            patcher = XlsxPatcher(path)
            patcher.sheet("Hoja")   # se abre pero no se modifica nada
            patcher.save()

            self.assertEqual(path.read_bytes(), before_bytes)


if __name__ == "__main__":
    unittest.main()
