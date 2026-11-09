"""
Pruebas de la protección contra fórmulas compartidas y de la reparación de libros.

Regresión real: al autorizar la corrección de la nota final, el cierre reemplazaba
la celda **maestra** de una fórmula compartida de Excel. Las demás celdas quedaban
con `si="N"` apuntando a un maestro inexistente; Excel "reparaba" el libro y
después mostraba «CARGA BLOQUEADA — Se ha producido un problema con este archivo».
"""

from __future__ import annotations

import re
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock

import openpyxl

from output import xlsx_patcher
from output.xlsx_patcher import CeldaProtegidaError, XlsxPatcher
from output.xlsx_recovery import diagnose_xlsx, repair_xlsx

SHEET = "xl/worksheets/sheet1.xml"


def _parts(hoja: str, calcpr: str | None = None) -> dict[str, bytes]:
    """Paquete .xlsx mínimo y válido con la hoja XML indicada."""
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
        "</Types>"
    )
    rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        "</Relationships>"
    )
    calc = calcpr if calcpr is not None else '<calcPr calcId="191029"/>'
    workbook = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="Formato" sheetId="1" r:id="rId1"/></sheets>'
        f"{calc}</workbook>"
    )
    workbook_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
        "</Relationships>"
    )
    styles = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
        '<fills count="1"><fill><patternFill/></fill></fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        "</styleSheet>"
    )
    shared_strings = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'count="1" uniqueCount="1"><si><t>Encabezado</t></si></sst>'
    )
    return {
        "[Content_Types].xml": content_types.encode("utf-8"),
        "_rels/.rels": rels.encode("utf-8"),
        "xl/workbook.xml": workbook.encode("utf-8"),
        "xl/_rels/workbook.xml.rels": workbook_rels.encode("utf-8"),
        "xl/styles.xml": styles.encode("utf-8"),
        "xl/sharedStrings.xml": shared_strings.encode("utf-8"),
        SHEET: hoja.encode("utf-8"),
    }


def _sheet(
    promedio_celda: str = '<c r="L2" s="1"><f>SUM(G2:K2)</f><v>4</v></c>',
    formula_extra: str = "",
    estilo_fuera_rango: bool = False,
    string_fuera_rango: bool = False,
) -> str:
    l3 = '<c r="L3" s="1"><v>4.5</v></c>'
    texto = '<c r="B2" t="inlineStr"><is><t>Texto</t></is></c>'
    estilo = '<c r="C2" s="99"><v>0</v></c>' if estilo_fuera_rango else ""
    if formula_extra:
        l3 = formula_extra
    if string_fuera_rango:
        texto = '<c r="B2" t="s"><v>5</v></c>'
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<dimension ref="A1:L3"/>'
        "<sheetData>"
        '<row r="1"><c r="A1" t="s"><v>0</v></c></row>'
        f'<row r="2"><c r="A2" s="1"><v>1</v></c>{texto}{estilo}{promedio_celda}</row>'
        f'<row r="3"><c r="A3" s="1"><v>2</v></c>{l3}</row>'
        "</sheetData></worksheet>"
    )


def _build(
    tmpdir: Path,
    hoja: str,
    calcpr: str | None = None,
    nombre: str = "libro.xlsx",
) -> Path:
    path = tmpdir / nombre
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in _parts(hoja, calcpr).items():
            zf.writestr(name, data)
    return path


def _read(path: Path, parte: str = SHEET) -> str:
    with zipfile.ZipFile(path) as zf:
        return zf.read(parte).decode("utf-8")


class TestPatcherFormulasCompartidas(unittest.TestCase):
    """El patcher no puede dejar fórmulas compartidas huérfanas."""

    def test_no_sobrescribe_la_maestra_de_una_formula_compartida(self):
        hoja = _sheet(
            promedio_celda=(
                '<c r="L2" s="1"><f t="shared" ref="L2:L3" si="0">'
                "SUM(G2:K2)</f><v>4</v></c>"
            ),
            formula_extra='<c r="L3" s="1"><f t="shared" si="0"/><v>4</v></c>',
        )
        with tempfile.TemporaryDirectory() as tmp:
            path = _build(Path(tmp), hoja)
            antes = path.read_bytes()

            patcher = XlsxPatcher(path)
            sheet = patcher.sheet("Formato")
            with self.assertRaises(CeldaProtegidaError) as ctx:
                sheet.set_cell(2, 12, 3.5, has_value=True)
            self.assertIn("fórmula compartida", ctx.exception.motivo)
            patcher.save()

            self.assertEqual(path.read_bytes(), antes, "el archivo no debía cambiar")
            xml = _read(path)
            self.assertIn('t="shared" ref="L2:L3"', xml)
            self.assertIn('<f t="shared" si="0"/>', xml)

    def test_si_permite_reemplazar_una_formula_normal(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = _build(Path(tmp), _sheet())
            patcher = XlsxPatcher(path)
            sheet = patcher.sheet("Formato")
            sheet.set_cell(2, 12, 3.5, has_value=True)  # no debe lanzar
            patcher.save()

            xml = _read(path)
            self.assertIn('<c r="L2" s="1"><v>3.5</v></c>', xml)
            self.assertNotIn("SUM(G2:K2)", xml)

    def test_la_formula_matricial_tambien_esta_protegida(self):
        hoja = _sheet(
            promedio_celda=(
                '<c r="L2" s="1"><f t="array" ref="L2:L3">SUM(G2:K2)</f><v>4</v></c>'
            )
        )
        with tempfile.TemporaryDirectory() as tmp:
            path = _build(Path(tmp), hoja)
            patcher = XlsxPatcher(path)
            sheet = patcher.sheet("Formato")
            with self.assertRaises(CeldaProtegidaError) as ctx:
                sheet.set_cell(2, 12, 3.5, has_value=True)
            self.assertIn("matricial", ctx.exception.motivo)

    def test_sanitiza_caracteres_de_control(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = _build(Path(tmp), _sheet())
            patcher = XlsxPatcher(path)
            sheet = patcher.sheet("Formato")
            sheet.set_cell(2, 2, "A\x0bB\x00C", has_value=True)
            patcher.save()

            xml = _read(path)
            self.assertNotIn("\x0b", xml)
            self.assertNotIn("\x00", xml)
            self.assertIn("ABC", xml)

    def test_save_aborta_si_el_xml_resultante_es_invalido(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = _build(Path(tmp), _sheet())
            antes = path.read_bytes()
            patcher = XlsxPatcher(path)
            sheet = patcher.sheet("Formato")
            sheet.set_cell(2, 2, "X", has_value=True)

            with mock.patch.object(
                xlsx_patcher._SheetEditor, "render", return_value="<worksheet>"
            ):
                with self.assertRaises(ValueError) as ctx:
                    patcher.save()
            self.assertIn("quedaría inválida", str(ctx.exception))
            self.assertEqual(path.read_bytes(), antes, "el original no debe tocarse")

    def test_conserva_celdas_y_filas_sin_atributo_r(self):
        """`r` es opcional en OOXML: no se pueden descartar esas celdas."""
        hoja = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            "<sheetData>"
            '<row><c t="inlineStr"><is><t>Uno</t></is></c><c><v>1</v></c></row>'
            '<row><c t="inlineStr"><is><t>Dos</t></is></c><c><v>2</v></c></row>'
            "</sheetData></worksheet>"
        )
        with tempfile.TemporaryDirectory() as tmp:
            path = _build(Path(tmp), hoja)
            patcher = XlsxPatcher(path)
            patcher.sheet("Formato").set_cell(1, 3, "Nuevo", has_value=True)
            patcher.save()

            xml = _read(path)
            for texto in ("Uno", "Dos", "Nuevo"):
                self.assertIn(texto, xml)
            self.assertIn('r="A1"', xml)
            self.assertIn('r="A2"', xml)

    def test_aborta_si_el_archivo_cambio_mientras_se_procesa(self):
        """Si Excel/OneDrive guarda el libro durante el cierre, no se sobrescribe."""
        with tempfile.TemporaryDirectory() as tmp:
            path = _build(Path(tmp), _sheet())
            patcher = XlsxPatcher(path)
            patcher.sheet("Formato").set_cell(2, 2, "X", has_value=True)

            # Simular un guardado externo (cambia el tamaño y la fecha).
            path.write_bytes(path.read_bytes() + b"\0")

            with self.assertRaises(RuntimeError) as ctx:
                patcher.save()
            self.assertIn("cambió mientras", str(ctx.exception))

    def test_no_duplica_calcpr_si_no_es_autocerrado(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = _build(
                Path(tmp), _sheet(), calcpr='<calcPr calcId="1"></calcPr>'
            )
            patcher = XlsxPatcher(path)
            patcher.sheet("Formato").set_cell(2, 2, "X", has_value=True)
            patcher.save()

            wb = _read(path, "xl/workbook.xml")
            self.assertEqual(wb.count("<calcPr"), 1)
            self.assertIn('fullCalcOnLoad="1"', wb)

    def test_save_aborta_si_ya_hay_calcpr_duplicado(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = _build(
                Path(tmp),
                _sheet(),
                calcpr='<calcPr calcId="1"/><calcPr calcId="2"/>',
            )
            antes = path.read_bytes()
            patcher = XlsxPatcher(path)
            patcher.sheet("Formato").set_cell(2, 2, "X", has_value=True)
            with self.assertRaises(ValueError) as ctx:
                patcher.save()
            self.assertIn("calcPr", str(ctx.exception))
            self.assertEqual(path.read_bytes(), antes)

    def test_no_escribe_celdas_en_columnas_lejanas_inexistentes(self):
        """El borde no debe crear miles de celdas si max_column viene inflado."""
        with tempfile.TemporaryDirectory() as tmp:
            hoja = _sheet().replace(
                '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
                '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                '<cols><col min="1" max="16384" width="10"/></cols>',
            )
            path = _build(Path(tmp), hoja)
            patcher = XlsxPatcher(path)
            sheet = patcher.sheet("Formato")
            # Aplicar borde solo hasta la columna 12 (ancho real de la tabla).
            sheet.set_row_border(2, 1, 12, patcher.styles.border_id(
                xlsx_patcher.BorderSpec()
            ), patcher.styles)
            patcher.save()
            xml = _read(path)
            self.assertNotIn('r="XFD2"', xml)


class TestRecuperacionXlsx(unittest.TestCase):
    """El diagnóstico detecta y la reparación corrige los daños conocidos."""

    def _daniado(self, tmpdir: Path) -> Path:
        hoja = _sheet(
            promedio_celda=(
                '<c r="L2" s="1"><f t="shared" ref="L2:L3" si="0">'
                "SUM(G2:K2)</f><v>4</v></c>"
            ),
            formula_extra='<c r="L3" s="1"><f t="shared" si="0"/><v>4</v></c>',
            estilo_fuera_rango=True,
            string_fuera_rango=True,
        )
        # Simular el daño del cierre: la maestra se reemplazó por un valor y la
        # esclava quedó apuntando a un grupo inexistente.
        hoja = hoja.replace(
            '<c r="L2" s="1"><f t="shared" ref="L2:L3" si="0">SUM(G2:K2)</f><v>4</v></c>',
            '<c r="L2" s="1"><v>4.5</v></c>',
        )
        return _build(
            tmpdir,
            hoja,
            calcpr='<calcPr calcId="1"/><calcPr calcId="2"/>',
        )

    def test_diagnostico_detecta_los_danos(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = self._daniado(Path(tmp))
            report = diagnose_xlsx(path)
            self.assertFalse(report.ok)
            self.assertEqual(report.esclavas_huerfanas, 1)
            self.assertEqual(report.calcpr_duplicados, 1)
            self.assertEqual(report.estilos_fuera_de_rango, 1)
            self.assertEqual(report.strings_fuera_de_rango, 1)

    def test_reparacion_deja_un_libro_que_excel_puede_abrir(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = self._daniado(Path(tmp))
            original = path.read_bytes()

            resultado = repair_xlsx(path)
            self.assertTrue(resultado.reparado, resultado.problemas_restantes)
            self.assertIsNotNone(resultado.salida)
            self.assertTrue(resultado.salida.exists())
            self.assertEqual(path.read_bytes(), original, "el original no debe cambiar")

            xml = _read(resultado.salida)
            self.assertNotIn('t="shared"', xml)
            self.assertIn("<v>4</v>", xml)  # valor de la esclava conservado
            self.assertNotIn('s="99"', xml)

            diagnostico = diagnose_xlsx(resultado.salida)
            self.assertTrue(diagnostico.ok, diagnostico.problemas)

            wb = openpyxl.load_workbook(resultado.salida)
            self.assertIn("Formato", wb.sheetnames)
            wb.close()

    def test_conserva_los_grupos_de_formulas_compartidas_sanos(self):
        hoja = _sheet(
            promedio_celda=(
                '<c r="L2" s="1"><f t="shared" ref="L2:L3" si="0">'
                "SUM(G2:K2)</f><v>4</v></c>"
            ),
            formula_extra='<c r="L3" s="1"><f t="shared" si="0"/><v>4</v></c>',
            estilo_fuera_rango=True,
        )
        with tempfile.TemporaryDirectory() as tmp:
            path = _build(
                Path(tmp),
                hoja,
                calcpr='<calcPr calcId="1"/><calcPr calcId="2"/>',
            )
            resultado = repair_xlsx(path)
            self.assertTrue(resultado.reparado, resultado.problemas_restantes)
            xml = _read(resultado.salida)
            self.assertIn('t="shared" ref="L2:L3"', xml)
            self.assertIn('<f t="shared" si="0"/>', xml)

    def test_repara_ampersand_sin_escapar(self):
        with tempfile.TemporaryDirectory() as tmp:
            hoja = _sheet().replace("Texto", "A & B")
            path = _build(Path(tmp), hoja)
            self.assertFalse(diagnose_xlsx(path).ok)

            resultado = repair_xlsx(path)
            self.assertTrue(resultado.reparado, resultado.problemas_restantes)
            self.assertIn("&amp;", _read(resultado.salida))
            self.assertTrue(diagnose_xlsx(resultado.salida).ok)

    def test_nunca_escribe_sobre_el_original(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = self._daniado(Path(tmp))
            with self.assertRaises(ValueError):
                repair_xlsx(path, output_path=path)

    def test_no_pisa_una_copia_reparada_anterior(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = self._daniado(Path(tmp))
            primera = repair_xlsx(path)
            segunda = repair_xlsx(path)
            self.assertNotEqual(primera.salida, segunda.salida)
            self.assertTrue(primera.salida.exists())
            self.assertTrue(segunda.salida.exists())


class TestReparacionPreservaElLibro(unittest.TestCase):
    """Reparar NO debe perder tablas dinámicas, gráficos, imágenes ni fórmulas."""

    def test_conserva_las_partes_fragiles_y_las_formulas(self):
        from tests.unit.test_preservacion_total import _build_rich_files

        with tempfile.TemporaryDirectory() as tmp:
            tmpdir = Path(tmp)
            files = _build_rich_files(tmpdir)
            path = files["sistematizacion"]

            # Simular el daño: esclavo de fórmula compartida sin maestro.
            with zipfile.ZipFile(path) as zf:
                antes = {n: zf.read(n) for n in zf.namelist()}
            xml = antes[SHEET].decode("utf-8")
            xml = re.sub(
                r'(<row r="3"[^>]*>)(.*?)(</row>)',
                r'\1<c r="Z3"><f t="shared" si="9"/></c>\2\3',
                xml,
                count=1,
                flags=re.S,
            )
            antes[SHEET] = xml.encode("utf-8")
            with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
                for name, data in antes.items():
                    zf.writestr(name, data)

            resultado = repair_xlsx(path)
            self.assertTrue(resultado.reparado, resultado.problemas_restantes)

            with zipfile.ZipFile(resultado.salida) as zf:
                despues = {n: zf.read(n) for n in zf.namelist()}

            # 1. No puede desaparecer ninguna parte del paquete.
            self.assertEqual(sorted(set(antes) - set(despues)), [])

            # 2. Todo lo que no sea la hoja reparada queda byte a byte igual.
            for nombre, datos in antes.items():
                if nombre == SHEET:
                    continue
                self.assertEqual(
                    datos, despues.get(nombre),
                    f"la parte '{nombre}' se modificó y no debía",
                )

            # 3. Tablas dinámicas, gráficos, imágenes y VML siguen ahí.
            self.assertIn("xl/pivotTables/pivotTable1.xml", despues)
            self.assertIn("xl/pivotCache/pivotCacheDefinition1.xml", despues)
            self.assertTrue(any(n.startswith("xl/charts/") for n in despues))
            self.assertTrue(any(n.startswith("xl/media/") for n in despues))
            self.assertTrue(any(n.endswith(".vml") for n in despues))

            # 4. Las fórmulas normales (no huérfanas) se conservan.
            hoja = despues[SHEET].decode("utf-8")
            self.assertIn("<f>SUM(G3:K3)</f>", hoja)
            # 5. Las listas de validación modernas (x14) también.
            self.assertIn("x14:dataValidations", hoja)


class TestPreflightCierre(unittest.TestCase):
    """El cierre no debe editar (ni respaldar) un libro que ya está dañado."""

    def test_cierre_detenido_antes_de_crear_respaldos(self):
        from output.direct_updater import apply_direct_cierre
        from tests.unit.test_preservacion_total import (
            _build_result,
            _build_rich_files,
            _meta,
        )

        with tempfile.TemporaryDirectory() as tmp:
            tmpdir = Path(tmp)
            files = _build_rich_files(tmpdir)

            # Simular el daño: un esclavo de fórmula compartida sin maestro.
            with zipfile.ZipFile(files["sistematizacion"]) as zf:
                parts = {n: zf.read(n) for n in zf.namelist()}
            xml = parts[SHEET].decode("utf-8")
            self.assertIn('<row r="3"', xml)
            xml = re.sub(
                r'(<row r="3"[^>]*>)(.*?)(</row>)',
                r'\1<c r="N3"><f t="shared" si="7"/></c>\2\3',
                xml,
                count=1,
                flags=re.S,
            )
            parts[SHEET] = xml.encode("utf-8")
            with zipfile.ZipFile(
                files["sistematizacion"], "w", zipfile.ZIP_DEFLATED
            ) as zf:
                for name, data in parts.items():
                    zf.writestr(name, data)
            danado = files["sistematizacion"].read_bytes()

            report = apply_direct_cierre(
                file_paths=files,
                result=_build_result(),
                codigo_curso="ABBUEI205",
                metadatos_control=_meta(),
                backup_dir=tmpdir / "backups",
            )

            self.assertFalse(report.is_successful)
            self.assertTrue(
                any("daños estructurales" in e for e in report.errors),
                report.errors,
            )
            self.assertEqual(report.backups_created, [])
            self.assertFalse((tmpdir / "backups").exists())
            # El libro dañado no se tocó (sigue exactamente como estaba).
            self.assertEqual(files["sistematizacion"].read_bytes(), danado)


if __name__ == "__main__":
    unittest.main()
