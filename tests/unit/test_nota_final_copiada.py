"""
Reglas de la nota final en Sistematización.

La nota final debe copiarse **tal cual** del archivo de Notas. La herramienta no
debe promediar módulos ni inventar fórmulas: los porcentajes y ponderaciones son
del curso y pueden variar.

Casos cubiertos:
  1. El motor de reglas ya no devuelve un promedio de módulos.
  2. La columna 'Promedio' recibe el valor exacto de Notas (no la media).
  3. Si Notas no trae nota final, la celda queda vacía (no se promedia).
  4. Una nota final simbólica ('A') se copia tal cual.
  5. Se prefiere la columna numérica sobre '... Scheme Symbol'.
  6. Una nota final con fórmula usa el valor calculado en caché.
  7. El cierre directo escribe en el archivo el valor exacto, no la media.
"""

from __future__ import annotations

import tempfile
import unittest
import zipfile
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
from normalization.headers import _detectar_nota_final, find_final_grade_column
from parsing.notas_parser import parse_notas_archivo
from output.direct_updater import apply_direct_cierre
from processing.rules_engine import apply_rules_to_match, evaluate_certification_status


def _tv(valor, col: int = 1, fila: int = 1):
    return TracedValue(valor, "N", "Calificaciones", "col", col, fila)


def _match(nota_final, califs=None, col_promedio: int = 8):
    colifs = califs or {"modulo_1": 5.0, "modulo_2": 4.0}  # media = 4.5
    es = EstudianteSistematizacion(
        row_number=2,
        nombres=_tv("ANA", 1),
        apellidos=_tv("PEREZ", 2),
        clase=_tv("6401", 3),
        documento=_tv("1001", 4),
        correo=_tv("ana@eafit.edu.co", 5),
        column_indices={
            "nombres": 1, "apellidos": 2, "clase": 3, "documento": 4, "correo": 5,
            "modulo_1": 6, "modulo_2": 7, "promedio": col_promedio, "estado": 9,
        },
    )
    en = EstudianteNotas(
        row_number=2,
        calificaciones_modulos={k: _tv(v) for k, v in colifs.items()},
        nota_final=_tv(nota_final) if nota_final is not None else None,
    )
    return StudentMatch(
        estudiante_sist=es, estudiante_nota=en, match_key="doc", match_value="1001"
    )


class TestNotaFinalNoSeCalcula(unittest.TestCase):
    def test_motor_de_reglas_no_devuelve_promedio_de_modulos(self):
        estado, nota = evaluate_certification_status({"modulo_1": 5.0, "modulo_2": 4.0})
        self.assertEqual(estado, "Aprobó")
        # Antes devolvía 4.5 (media de módulos). Ahora no hay nota final → None.
        self.assertIsNone(nota)

    def test_motor_de_reglas_respeta_la_nota_final_de_notas(self):
        estado, nota = evaluate_certification_status(
            {"modulo_1": 5.0, "modulo_2": 4.0}, nota_final_calculated=4.2
        )
        self.assertEqual(estado, "Aprobó")
        self.assertEqual(nota, 4.2)

    def test_apply_rules_copia_la_nota_final_exacta(self):
        match = _match(nota_final=4.2)
        apply_rules_to_match(match, codigo_curso="ABBUEI227", clase_id="6401")
        self.assertEqual(match.estudiante_sist.raw_row_data[8], 4.2)
        self.assertNotEqual(match.estudiante_sist.raw_row_data[8], 4.5)

    def test_sin_nota_final_no_se_escribe_promedio(self):
        match = _match(nota_final=None)
        apply_rules_to_match(match, codigo_curso="ABBUEI227", clase_id="6401")
        self.assertNotIn(8, match.estudiante_sist.raw_row_data)

    def test_nota_final_simbolica_se_copia_tal_cual(self):
        match = _match(nota_final="A")
        apply_rules_to_match(match, codigo_curso="ABBUEI227", clase_id="6401")
        self.assertEqual(match.estudiante_sist.raw_row_data[8], "A")

    def test_formula_sin_valor_no_se_copia_como_formula(self):
        match = _match(nota_final="=AVERAGE(F2:I2)")
        apply_rules_to_match(match, codigo_curso="ABBUEI227", clase_id="6401")
        self.assertNotIn(8, match.estudiante_sist.raw_row_data)

    def test_prefiere_columna_numerica_sobre_scheme_symbol(self):
        headers = [
            "Calculated Final Grade",
            "Calculated Final Grade Scheme Symbol",
        ]
        # Con ambas columnas gana la numérica (columna 1).
        self.assertEqual(find_final_grade_column(headers), 1)
        # Si solo existe el símbolo, se usa como último recurso (se copia tal cual).
        self.assertEqual(find_final_grade_column([headers[1]]), 1)


class TestDeteccionNotaFinalRobusta(unittest.TestCase):
    """La columna de nota final debe reconocerse con los nombres reales."""

    def test_nombres_habituales(self):
        esperados = {
            "Calculated Final Grade": "nota_final",
            "Calculated Final": "nota_final",
            "Final": "nota_final",
            "Final Grade": "nota_final",
            "Nota Final": "nota_final",
            "Nota definitiva": "nota_final",
            "Calificación final": "nota_final",
            "Definitiva": "nota_final",
            "Final Score": "nota_final",
            "Puntaje final": "nota_final",
            "Promedio curso completo": "promedio_curso",
            "Promedio final": "nota_final",
            "Calculated Final Grade Scheme Symbol": "nota_final",
        }
        for header, canonical in esperados.items():
            detectada = _detectar_nota_final(header)
            self.assertIsNotNone(detectada, f"No se detectó: {header}")
            self.assertEqual(detectada[1], canonical, header)

    def test_no_confunde_fechas_ni_actividades(self):
        for header in [
            "Fecha Final (DD/MM/AAAA)",
            "Hora final",
            "Examen final",
            "Trabajo final",
            "Proyecto final",
            "Nombre final",
            "Sección final",
            "Cuestionario módulo 1",
        ]:
            self.assertIsNone(_detectar_nota_final(header), header)

    def test_reconoce_scheme_symbol_como_columna_valida(self):
        headers = [
            "Calculated Final Grade Scheme Symbol",
            "Calculated Final Grade",
        ]
        # Ambas son candidatas válidas; con solo encabezados se toma la primera
        # de mayor puntaje (empate). El parser decide luego con los datos.
        self.assertEqual(find_final_grade_column(headers), 1)

    def test_encabezado_final_en_cualquier_posicion(self):
        self.assertEqual(find_final_grade_column(["Documento", "Nombre", "Final"]), 3)
        self.assertEqual(
            find_final_grade_column(["Documento", "Calculated Final", "Correo"]), 2
        )

    def test_sin_columna_de_nota_final(self):
        self.assertIsNone(
            find_final_grade_column(["Documento", "Nombre", "Fecha Final", "Correo"])
        )


class TestNotaFinalConFormula(unittest.TestCase):
    def _crear_notas(self, path: Path, con_cache: bool) -> None:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Calificaciones"
        ws.append([
            "OrgDefinedId", "Username", "Last Name", "First Name", "Sección",
            "Cuestionario módulo 1", "Buzón módulo 2", "Cuestionario módulo 3",
            "Cuestionario módulo 4", "Calculated Final",
        ])
        ws.append(["1001", "user", "PEREZ", "ANA", "SEC-1", 5, 4, 5, 4, None])
        ws["J2"] = "=AVERAGE(F2:I2)"
        wb.save(path)
        wb.close()

        if not con_cache:
            return

        # openpyxl no guarda el valor calculado de las fórmulas: se inyecta el
        # `<v>` en el XML para simular un archivo que sí pasó por Excel.
        tmp = path.with_name(path.stem + "_tmp.xlsx")
        with zipfile.ZipFile(path) as zin, zipfile.ZipFile(
            tmp, "w", zipfile.ZIP_DEFLATED
        ) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "xl/worksheets/sheet1.xml":
                    xml = data.decode("utf-8")
                    xml = xml.replace(
                        "<f>AVERAGE(F2:I2)</f>", "<f>AVERAGE(F2:I2)</f><v>4.35</v>"
                    )
                    data = xml.encode("utf-8")
                zout.writestr(item, data)
        tmp.replace(path)

    def test_usa_el_valor_calculado_en_cache(self):
        with tempfile.TemporaryDirectory() as tmp_str:
            path = Path(tmp_str) / "Notas_ABBUEI227.xlsx"
            self._crear_notas(path, con_cache=True)
            estudiantes, _ = parse_notas_archivo(path)
            self.assertEqual(len(estudiantes), 1)
            self.assertAlmostEqual(float(estudiantes[0].nota_final.value), 4.35)

    def test_sin_cache_la_nota_final_no_es_numerica(self):
        with tempfile.TemporaryDirectory() as tmp_str:
            path = Path(tmp_str) / "Notas_ABBUEI227.xlsx"
            self._crear_notas(path, con_cache=False)
            estudiantes, _ = parse_notas_archivo(path)
            # Queda el texto de la fórmula; el cierre lo tratará como ausente y
            # NO lo convertirá en promedio ni escribirá la fórmula.
            self.assertEqual(estudiantes[0].nota_final.value, "=AVERAGE(F2:I2)")


class TestSeleccionColumnaPorDatos(unittest.TestCase):
    """Con varias columnas candidatas, decide por los datos reales, no por el nombre."""

    def _crear_notas(self, path: Path, filas: list[list], headers: list[str]) -> None:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Calificaciones"
        ws.append(headers)
        for fila in filas:
            ws.append(fila)
        wb.save(path)
        wb.close()

    def test_usa_scheme_symbol_si_alli_estan_las_notas(self):
        headers = [
            "OrgDefinedId", "Username", "Last Name", "First Name",
            "Calculated Final Grade", "Calculated Final Grade Scheme Symbol",
        ]
        with tempfile.TemporaryDirectory() as tmp_str:
            path = Path(tmp_str) / "Notas.xlsx"
            self._crear_notas(path, [
                ["1001", "u1", "PEREZ", "ANA", None, 4.7],
                ["1002", "u2", "GOMEZ", "LUIS", None, 3.9],
            ], headers)
            estudiantes, _ = parse_notas_archivo(path)
            self.assertEqual(estudiantes[0].nota_final.value, 4.7)
            self.assertEqual(estudiantes[1].nota_final.value, 3.9)

    def test_usa_la_columna_numerica_si_la_otra_trae_letras(self):
        headers = [
            "OrgDefinedId", "Username", "Last Name", "First Name",
            "Calculated Final Grade", "Calculated Final Grade Scheme Symbol",
        ]
        with tempfile.TemporaryDirectory() as tmp_str:
            path = Path(tmp_str) / "Notas.xlsx"
            self._crear_notas(path, [
                ["1001", "u1", "PEREZ", "ANA", 4.8, "A"],
                ["1002", "u2", "GOMEZ", "LUIS", 3.2, "C"],
            ], headers)
            estudiantes, _ = parse_notas_archivo(path)
            self.assertEqual(estudiantes[0].nota_final.value, 4.8)
            self.assertEqual(estudiantes[1].nota_final.value, 3.2)

    def test_encabezado_final_es_suficiente(self):
        headers = [
            "OrgDefinedId", "Username", "Last Name", "First Name", "Final",
        ]
        with tempfile.TemporaryDirectory() as tmp_str:
            path = Path(tmp_str) / "Notas.xlsx"
            self._crear_notas(path, [["1001", "u1", "PEREZ", "ANA", 4.1]], headers)
            estudiantes, _ = parse_notas_archivo(path)
            self.assertEqual(estudiantes[0].nota_final.value, 4.1)


class TestCierreEscribeNotaExacta(unittest.TestCase):
    def _crear_archivos(self, tmpdir: Path) -> dict:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Formato"
        ws.append([
            "Nombres", "Apellidos", "Clase", "Documento", "Correo",
            "Módulo 1", "Módulo 2", "Promedio curso completo", "Estado",
        ])
        ws.append(["ANA", "PEREZ", "6401", "1001", "ana@eafit.edu.co", 5, 4, None, None])
        path_sist = tmpdir / "Sistematizacion_Cursos_COIN_2026.xlsx"
        wb.save(path_sist)
        wb.close()

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Códigos"
        ws.append([
            "N°", "Ciclo", "Docente COIN", "Nombres", "Apellidos",
            "Documento de identidad", "Correo Electrónico", "Código del certificado",
            "Fecha de envío", "Año de certificación", "Semestre",
            "Total certificados", "Fecha de totalización",
        ])
        path_cert = tmpdir / "Control_codigos_certificados_ABBUEI.xlsx"
        wb.save(path_cert)
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

    def test_escribe_4_2_y_no_la_media_4_5(self):
        with tempfile.TemporaryDirectory() as tmp_str:
            tmpdir = Path(tmp_str)
            files = self._crear_archivos(tmpdir)

            es = EstudianteSistematizacion(
                row_number=2,
                nombres=_tv("ANA", 1, 2),
                apellidos=_tv("PEREZ", 2, 2),
                clase=_tv("6401", 3, 2),
                documento=_tv("1001", 4, 2),
                correo=_tv("ana@eafit.edu.co", 5, 2),
                estado_calculado="Aprobó",
                column_indices={
                    "nombres": 1, "apellidos": 2, "clase": 3, "documento": 4,
                    "correo": 5, "modulo_1": 6, "modulo_2": 7, "promedio": 8,
                    "estado": 9,
                },
                # La nota final viene copiada de Notas: 4.2 (no 4.5)
                raw_row_data={6: 5.0, 7: 4.0, 8: 4.2, 9: "Aprobó"},
            )
            match = StudentMatch(
                estudiante_sist=es,
                estudiante_nota=EstudianteNotas(row_number=2),
                match_key="doc",
                match_value="1001",
            )
            meta = {
                "6401": AulaMetadata(
                    programa_academico=_tv("FINANZAS"),
                    catalogo=_tv("FN01"),
                    docente_principal=_tv("DOCENTE"),
                    fecha_inicio=_tv("16/03/2026"),
                    fecha_fin=_tv("27/04/2026"),
                )
            }
            result = AulaCierreResult(
                aula=AulaMetadata(),
                clases=[ClaseGroup(clase_id="6401", matches=[match])],
                metadatos_control=meta,
                start_consecutivo=1,
            )

            report = apply_direct_cierre(
                files, result, "ABBUEI227", metadatos_control=meta,
                start_consecutivo=1, backup_dir=tmpdir / "bk",
            )
            self.assertTrue(report.is_successful, report.errors)
            self.assertEqual(report.sistematizacion_promedios_no_copiados, 0)

            wb = openpyxl.load_workbook(files["sistematizacion"])
            ws = wb["Formato"]
            self.assertEqual(ws.cell(2, 8).value, 4.2)
            self.assertNotEqual(ws.cell(2, 8).value, 4.5)
            wb.close()


class TestPipelineCopiaLaNotaFinal(unittest.TestCase):
    """Verifica de extremo a extremo que la hoja SISTEMATIZACION copie Notas."""

    def test_la_hoja_sistematizacion_trae_la_nota_final_de_notas(self):
        from processing.cierre_orchestrator import run_cierre_pipeline
        from tests.fixtures.generate_fixtures import generate_all_fixtures

        fixtures = generate_all_fixtures()
        with tempfile.TemporaryDirectory() as tmp_str:
            result, out_file, intact, _ = run_cierre_pipeline(
                file_paths=fixtures,
                target_aula="ABBUEI154",
                codigo_curso="ABBUEI154",
                is_simulation=False,
                output_dir=tmp_str,
            )
            self.assertTrue(intact)
            self.assertIsNotNone(out_file)

            wb = openpyxl.load_workbook(out_file, data_only=True)
            ws = wb["SISTEMATIZACION"]
            headers = [str(c.value or "").strip().upper() for c in ws[1]]
            col_prom = headers.index("PROMEDIO CURSO COMPLETO") + 1
            col_nom = headers.index("NOMBRES") + 1

            promedios = {}
            for r in range(2, (ws.max_row or 1) + 1):
                nombre = ws.cell(r, col_nom).value
                if nombre:
                    promedios[str(nombre).strip().upper()] = ws.cell(r, col_prom).value
            wb.close()

            # Notas dice 4.8 para Juan Camilo; el promedio de sus módulos sería
            # 4.825 (4.83). La herramienta debe copiar 4.8, no promediar.
            self.assertIn("JUAN CAMILO", promedios)
            self.assertEqual(promedios["JUAN CAMILO"], 4.8)
            self.assertNotEqual(promedios["JUAN CAMILO"], 4.83)


if __name__ == "__main__":
    unittest.main()
