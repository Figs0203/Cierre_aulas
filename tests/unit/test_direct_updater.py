"""
Pruebas unitarias para el módulo de actualización directa sobre archivos oficiales (direct_updater.py).

Verifica:
1. Creación de copias de seguridad automáticas (backup) antes de modificar.
2. Protección estricta de celdas preexistentes (no sobreescritura).
3. Estandarización de nombres y apellidos a MAYÚSCULAS sostenidas.
4. Inserción de notas y estados oficiales solo en celdas vacías.
5. Aplicación de borde inferior grueso entre clases en Sistematización y Certificados.
6. Anexado correcto en Certificados respetando registros previos.
7. Marcado de fecha de cierre en Control de Aulas.
8. Conservación intacta del archivo de notas (solo lectura).
"""

from datetime import datetime
from pathlib import Path
import tempfile
import unittest
import openpyxl
from openpyxl.styles import Border, Side

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
    apply_direct_cierre,
    create_timestamped_backup,
    _strip_external_links,
)


class TestDirectUpdater(unittest.TestCase):

    def _create_mock_files(self, tmpdir: Path):
        """Genera archivos Excel de prueba simulando la estructura oficial."""
        # 1. Sistematización
        wb_sist = openpyxl.Workbook()
        ws_sist = wb_sist.active
        ws_sist.title = "Formato"
        headers_sist = [
            "ID", "Nombres", "Apellidos", "Clase", "Documento", "Correo",
            "Módulo 1", "Módulo 2", "Módulo 3", "Módulo 4", "Módulo 5",
            "Promedio curso completo", "Estado", "Se elaboró certificado",
            "Código del certificado", "Se envió certificado", "Fecha de envío",
        ]
        ws_sist.append(headers_sist)
        # Fila 2: Estudiante con nombres en minúsculas y celda de Módulo 1 ya diligenciada (preexistente)
        ws_sist.append([
            1, "carlos andres", "perez gomez", "5535", "1001", "cperez@eafit.edu.co",
            4.5, None, None, None, None,  # Módulo 1 = 4.5 preexistente
            None, None, None, None, None, None,
        ])
        # Fila 3: Estudiante con nombres mixtos y celdas de cierre vacías
        ws_sist.append([
            2, "Ana Maria", "Lopez Duque", "5536", "1002", "alopez@eafit.edu.co",
            None, None, None, None, None,
            None, None, None, None, None, None,
        ])
        path_sist = tmpdir / "Sistematizacion_Cursos_COIN_2026.xlsx"
        wb_sist.save(path_sist)

        # 2. Certificados
        wb_cert = openpyxl.Workbook()
        ws_cert = wb_cert.active
        ws_cert.title = "Códigos"
        headers_cert = [
            "N°", "Ciclo", "Docente COIN", "Nombres", "Apellidos",
            "Documento de identidad", "Correo Electrónico", "Código del certificado",
            "Fecha de envío", "Año de certificación", "Semestre", "Total certificados",
            "Fecha de totalización",
        ]
        ws_cert.append(headers_cert)
        # Fila previa existente (debe quedar intacta)
        ws_cert.append([
            10, "CLASE PREVIA", "DOCENTE PREVIO", "PREVIO", "PREVIO",
            "9999", "previo@eafit.edu.co", "ABBUEI205", "15 DE MARZO DE 2026",
            2026, 1, 1, None,
        ])
        path_cert = tmpdir / "Control_codigos_certificados_ABBUEI.xlsx"
        wb_cert.save(path_cert)

        # 3. Control de Aulas
        wb_ctrl = openpyxl.Workbook()
        ws_ctrl = wb_ctrl.active
        ws_ctrl.title = "Tabla_ABBUEI"
        headers_ctrl = [
            "PROGRAMA ACADÉMICO", "CATÁLOGO", "CLASE", "DOCENTE PRINCIPAL",
            "FECHA DE INICIO DEL CURSO", "FECHA DE FINALIZACIÓN DEL CURSO", "CERTIFICADOS",
        ]
        ws_ctrl.append(headers_ctrl)
        ws_ctrl.append(["ADMINISTRACION", "AD01", "5535", "Manuela Restrepo", "16/03/2026", "27/04/2026", None])
        ws_ctrl.append(["FINANZAS", "FN01", "5536", "Manuela Restrepo", "16/03/2026", "27/04/2026", "01/01/2026"]) # Ya tiene fecha previa
        path_ctrl = tmpdir / "Control_Aulas_Formadores_2026.xlsx"
        wb_ctrl.save(path_ctrl)

        # 4. Notas (Solo Lectura)
        wb_notas = openpyxl.Workbook()
        ws_notas = wb_notas.active
        ws_notas.append(["OrgDefinedId", "Username", "Last Name", "First Name", "Calculated Final"])
        ws_notas.append(["1001", "cperez", "perez", "carlos", 4.8])
        ws_notas.append(["1002", "alopez", "lopez", "ana", 4.2])
        path_notas = tmpdir / "Notas_ABBUEI205.xlsx"
        wb_notas.save(path_notas)

        return {
            "sistematizacion": path_sist,
            "certificados": path_cert,
            "control_aulas": path_ctrl,
            "notas": path_notas,
        }

    def test_backup_creation(self):
        with tempfile.TemporaryDirectory() as tmp_str:
            tmpdir = Path(tmp_str)
            files = self._create_mock_files(tmpdir)
            bk = create_timestamped_backup(files["sistematizacion"])
            self.assertTrue(bk.exists())
            self.assertEqual(bk.parent.name, "_backups_cierre")
            self.assertIn("Sistematizacion_Cursos_COIN_2026_backup_", bk.name)
            self.assertEqual(bk.stat().st_size, files["sistematizacion"].stat().st_size)

    def test_apply_direct_cierre_full_flow(self):
        with tempfile.TemporaryDirectory() as tmp_str:
            tmpdir = Path(tmp_str)
            files = self._create_mock_files(tmpdir)

            col_map = {
                "nombres": 2, "apellidos": 3, "clase": 4, "documento": 5, "correo": 6,
                "modulo_1": 7, "modulo_2": 8, "modulo_3": 9, "modulo_4": 10, "modulo_5": 11,
                "promedio": 12, "estado_certificacion": 13, "elaboro_certificado": 14,
                "codigo_certificado": 15, "envio_certificado": 16, "fecha_envio": 17,
            }

            def tv(val, col, row):
                return TracedValue(val, "S", "Formato", col, 1, row)

            # Estudiante 1 (Clase 5535, fila 2)
            es1 = EstudianteSistematizacion(
                row_number=2,
                nombres=tv("carlos andres", "Nombres", 2),
                apellidos=tv("perez gomez", "Apellidos", 2),
                clase=tv("5535", "Clase", 2),
                documento=tv("1001", "Documento", 2),
                correo=tv("cperez@eafit.edu.co", "Correo", 2),
                estado_calculado="Aprobó",
                column_indices=col_map,
                raw_row_data={
                    7: 5.0,  # Nota nueva para Mod 1 (pero ya existía 4.5 en celda original)
                    8: 4.8,  # Mod 2 vacía
                    12: 4.8, # Promedio (Calculated Final)
                    13: "Aprobó",
                    14: "SI",
                    15: "ABBUEI205",
                    16: "SI",
                    17: "27 DE ABRIL DE 2026",
                },
            )
            en1 = EstudianteNotas(row_number=2)
            m1 = StudentMatch(estudiante_sist=es1, estudiante_nota=en1, match_key="doc", match_value="1001")

            # Estudiante 2 (Clase 5536, fila 3)
            es2 = EstudianteSistematizacion(
                row_number=3,
                nombres=tv("Ana Maria", "Nombres", 3),
                apellidos=tv("Lopez Duque", "Apellidos", 3),
                clase=tv("5536", "Clase", 3),
                documento=tv("1002", "Documento", 3),
                correo=tv("alopez@eafit.edu.co", "Correo", 3),
                estado_calculado="Aprobó",
                column_indices=col_map,
                raw_row_data={
                    7: 4.0,
                    8: 4.2,
                    12: 4.2,
                    13: "Aprobó",
                    14: "SI",
                    15: "ABBUEI205",
                    16: "SI",
                    17: "27 DE ABRIL DE 2026",
                },
            )
            en2 = EstudianteNotas(row_number=3)
            m2 = StudentMatch(estudiante_sist=es2, estudiante_nota=en2, match_key="doc", match_value="1002")

            cg1 = ClaseGroup(clase_id="5535", matches=[m1])
            cg2 = ClaseGroup(clase_id="5536", matches=[m2])
            res = AulaCierreResult(aula=AulaMetadata(), clases=[cg1, cg2])

            meta_dict = {
                "5535": AulaMetadata(
                    programa_academico=tv("ADMINISTRACION", "P", 2),
                    catalogo=tv("AD01", "C", 2),
                    docente_principal=tv("Manuela Restrepo", "D", 2),
                    fecha_inicio=tv("16/03/2026", "FI", 2),
                    fecha_fin=tv("27/04/2026", "FF", 2),
                ),
                "5536": AulaMetadata(
                    programa_academico=tv("FINANZAS", "P", 3),
                    catalogo=tv("FN01", "C", 3),
                    docente_principal=tv("Manuela Restrepo", "D", 3),
                    fecha_inicio=tv("16/03/2026", "FI", 3),
                    fecha_fin=tv("27/04/2026", "FF", 3),
                ),
            }

            # Guardar tamaño previo de notas para verificar que permanezca intacto
            notas_stat_before = files["notas"].stat().st_mtime_ns

            # Ejecutar actualización directa
            report = apply_direct_cierre(
                file_paths=files,
                result=res,
                codigo_curso="ABBUEI205",
                metadatos_control=meta_dict,
                start_consecutivo=1,
            )

            # Verificar que el reporte indique éxito y backups creados
            self.assertTrue(report.is_successful)
            self.assertEqual(len(report.backups_created), 3)

            # -------------------------------------------------------------
            # VERIFICACIÓN EN SISTEMATIZACIÓN
            # -------------------------------------------------------------
            wb_sist = openpyxl.load_workbook(files["sistematizacion"], data_only=True)
            ws_sist = wb_sist["Formato"]

            # 1. Nombres y apellidos estandarizados a MAYÚSCULAS
            self.assertEqual(ws_sist.cell(2, 2).value, "CARLOS ANDRES")
            self.assertEqual(ws_sist.cell(2, 3).value, "PEREZ GOMEZ")
            self.assertEqual(ws_sist.cell(3, 2).value, "ANA MARIA")
            self.assertEqual(ws_sist.cell(3, 3).value, "LOPEZ DUQUE")

            # 2. Celda preexistente intacta: Módulo 1 en fila 2 tenía 4.5, NO DEBE SER 5.0
            self.assertEqual(ws_sist.cell(2, 7).value, 4.5)

            # 3. Celda vacía actualizada: Módulo 2 en fila 2 estaba vacía, ahora debe tener 4.8
            self.assertEqual(ws_sist.cell(2, 8).value, 4.8)

            # 4. Promedio y Estado insertados correctamente
            self.assertEqual(ws_sist.cell(2, 12).value, 4.8)
            self.assertEqual(ws_sist.cell(2, 13).value, "Aprobó")
            self.assertEqual(ws_sist.cell(2, 14).value, "SI")
            self.assertEqual(ws_sist.cell(2, 15).value, "ABBUEI205")

            # 5. Bordes entre clases: Fila 2 (fin de 5535) y Fila 3 (fin de 5536) con borde medium
            self.assertEqual(ws_sist.cell(2, 1).border.bottom.style, "medium")
            self.assertEqual(ws_sist.cell(3, 1).border.bottom.style, "medium")

            # 5b. TODAS las celdas de las filas procesadas tienen los cuatro bordes.
            #     La celda de fin de clase conserva los tres lados finos además del inferior grueso.
            for r in (2, 3):
                for c in range(1, 16):
                    b = ws_sist.cell(r, c).border
                    self.assertIsNotNone(b.left.style, f"Falta borde izquierdo en Sist r{r}c{c}")
                    self.assertIsNotNone(b.right.style, f"Falta borde derecho en Sist r{r}c{c}")
                    self.assertIsNotNone(b.top.style, f"Falta borde superior en Sist r{r}c{c}")
                    self.assertIsNotNone(b.bottom.style, f"Falta borde inferior en Sist r{r}c{c}")
            # Los lados de la celda de fin de clase son finos (no vacíos) y solo el inferior es grueso
            self.assertEqual(ws_sist.cell(2, 1).border.left.style, "thin")
            self.assertEqual(ws_sist.cell(2, 1).border.top.style, "thin")
            wb_sist.close()

            # -------------------------------------------------------------
            # VERIFICACIÓN EN CERTIFICADOS
            # -------------------------------------------------------------
            wb_cert = openpyxl.load_workbook(files["certificados"], data_only=True)
            ws_cert = wb_cert["Códigos"]

            # Fila 2 previa intacta (N° = 10)
            self.assertEqual(ws_cert.cell(2, 1).value, 10)
            self.assertEqual(ws_cert.cell(2, 4).value, "PREVIO")

            # Fila 3 anexada para Carlos Andrés (N° = 11 correlativo)
            self.assertEqual(ws_cert.cell(3, 1).value, 11)
            self.assertEqual(ws_cert.cell(3, 3).value, "ABBUEI-205 MANUELA RESTREPO")
            self.assertEqual(ws_cert.cell(3, 4).value, "CARLOS ANDRES")
            self.assertEqual(ws_cert.cell(3, 5).value, "PEREZ GOMEZ")
            self.assertEqual(ws_cert.cell(3, 7).value, "cperez@eafit.edu.co")
            self.assertEqual(ws_cert.cell(3, 8).value, "ABBUEI205")
            self.assertEqual(ws_cert.cell(3, 12).value, 1)
            # La columna N° (1) es la ÚNICA sin borde inferior grueso en fin de clase.
            self.assertEqual(ws_cert.cell(3, 1).border.bottom.style, "thin")
            self.assertEqual(ws_cert.cell(3, 2).border.bottom.style, "medium")  # Borde fin de clase 5535

            # Fila 4 anexada para Ana María (N° = 12 correlativo)
            self.assertEqual(ws_cert.cell(4, 1).value, 12)
            self.assertEqual(ws_cert.cell(4, 4).value, "ANA MARIA")
            self.assertEqual(ws_cert.cell(4, 12).value, 1)
            self.assertEqual(ws_cert.cell(4, 1).border.bottom.style, "thin")
            self.assertEqual(ws_cert.cell(4, 2).border.bottom.style, "medium")  # Borde fin de clase 5536

            # Toda la información de Certificados va centrada.
            self.assertEqual(ws_cert.cell(3, 4).alignment.horizontal, "center")
            self.assertEqual(ws_cert.cell(3, 7).alignment.horizontal, "center")
            self.assertEqual(ws_cert.cell(4, 5).alignment.horizontal, "center")

            # Fuente oficial de la tabla de Certificados: Zurich Cn BT 11
            self.assertEqual(ws_cert.cell(3, 1).font.name, "Zurich Cn BT")
            self.assertEqual(ws_cert.cell(3, 1).font.size, 11)
            self.assertEqual(ws_cert.cell(4, 1).font.name, "Zurich Cn BT")
            self.assertEqual(ws_cert.cell(4, 1).font.size, 11)

            # TODAS las celdas anexadas tienen los cuatro bordes (incluida la de fin de clase)
            for r in (3, 4):
                for c in range(1, 14):
                    b = ws_cert.cell(r, c).border
                    self.assertIsNotNone(b.left.style, f"Falta borde izquierdo en Cert r{r}c{c}")
                    self.assertIsNotNone(b.right.style, f"Falta borde derecho en Cert r{r}c{c}")
                    self.assertIsNotNone(b.top.style, f"Falta borde superior en Cert r{r}c{c}")
                    self.assertIsNotNone(b.bottom.style, f"Falta borde inferior en Cert r{r}c{c}")

            # Columna 'Semestre' (11) en formato oficial 'AAAA-S'
            import re as _re
            for r in (3, 4):
                sem = str(ws_cert.cell(r, 11).value or "")
                self.assertRegex(sem, _re.compile(r"^\d{4}-[12]$"), f"Semestre con formato inesperado: {sem!r}")
            wb_cert.close()

            # -------------------------------------------------------------
            # VERIFICACIÓN EN CONTROL DE AULAS
            # -------------------------------------------------------------
            wb_ctrl = openpyxl.load_workbook(files["control_aulas"], data_only=True)
            ws_ctrl = wb_ctrl["Tabla_ABBUEI"]

            # Fila 2 (5535): estaba vacía, ahora debe tener fecha de hoy DD/MM/YYYY
            fecha_hoy = datetime.now().strftime("%d/%m/%Y")
            self.assertEqual(ws_ctrl.cell(2, 7).value, fecha_hoy)

            # Fila 3 (5536): ya tenía "01/01/2026", NO DEBE SOBREESCRIBIRSE
            self.assertEqual(ws_ctrl.cell(3, 7).value, "01/01/2026")
            wb_ctrl.close()

            # -------------------------------------------------------------
            # VERIFICACIÓN EN NOTAS (SOLO LECTURA)
            # -------------------------------------------------------------
            notas_stat_after = files["notas"].stat().st_mtime_ns
            self.assertEqual(notas_stat_before, notas_stat_after)

    def test_row_identity_discrepancy_protection(self):
        """Verifica que si la fila en Sistematización no corresponde a la persona esperada, se proteja y no se escriba."""
        with tempfile.TemporaryDirectory() as tmp_str:
            tmpdir = Path(tmp_str)
            files = self._create_mock_files(tmpdir)

            col_map = {
                "nombres": 2, "apellidos": 3, "clase": 4, "documento": 5, "correo": 6,
                "modulo_1": 7, "modulo_2": 8, "promedio": 12, "estado_certificacion": 13,
            }

            def tv(val, col, row):
                return TracedValue(val, "S", "Formato", col, 1, row)

            # Estudiante configurado con row_number=2, pero con documento y nombre COMPLETAMENTE DISTINTOS
            # a lo que hay en la fila 2 del archivo ("carlos andres perez gomez", doc "1001")
            es_discrepante = EstudianteSistematizacion(
                row_number=2,
                nombres=tv("ZULMA", "Nombres", 2),
                apellidos=tv("RESTREPO", "Apellidos", 2),
                clase=tv("5535", "Clase", 2),
                documento=tv("99999999", "Documento", 2),
                correo=tv("zrestrepo@eafit.edu.co", "Correo", 2),
                estado_calculado="Aprobó",
                column_indices=col_map,
                raw_row_data={7: 5.0, 8: 5.0, 12: 5.0, 13: "Aprobó"},
            )
            en = EstudianteNotas(row_number=2)
            m = StudentMatch(estudiante_sist=es_discrepante, estudiante_nota=en, match_key="doc", match_value="99999999")
            cg = ClaseGroup(clase_id="5535", matches=[m])
            res = AulaCierreResult(aula=AulaMetadata(), clases=[cg])

            report = apply_direct_cierre(
                file_paths=files,
                result=res,
                codigo_curso="ABBUEI205",
                start_consecutivo=1,
            )

            # Debe haber registrado un error de discrepancia para proteger la fila
            self.assertFalse(report.is_successful)
            self.assertTrue(any("DISCREPANCIA EN FILA 2" in err for err in report.errors))

            # Y la celda original en fila 2 debe permanecer intacta (sin sobreescritura)
            wb_sist = openpyxl.load_workbook(files["sistematizacion"], data_only=True)
            self.assertEqual(wb_sist["Formato"].cell(2, 7).value, 4.5)
            wb_sist.close()

    def test_strip_external_links_removes_external_references(self):
        """Las referencias externas deben eliminarse para que Excel no pida reparar el libro."""
        import zipfile
        from openpyxl.workbook.external_link.external import ExternalLink
        from openpyxl.packaging.relationship import Relationship

        with tempfile.TemporaryDirectory() as tmp_str:
            tmpdir = Path(tmp_str)
            src = tmpdir / "con_links.xlsx"

            wb = openpyxl.Workbook()
            wb.active["A1"] = 1
            link = ExternalLink()
            link.file_link = Relationship(type="externalLink", Target="externalLink1.xml")
            link.file_link.TargetMode = "External"
            wb._external_links.append(link)
            wb.save(src)

            # El archivo original SÍ contiene referencias externas
            with zipfile.ZipFile(src) as z:
                self.assertTrue(any("external" in n.lower() for n in z.namelist()))

            # Tras sanear y guardar, ya no debe quedar ninguna referencia externa
            wb2 = openpyxl.load_workbook(src)
            self.assertEqual(len(wb2._external_links), 1)
            removed = _strip_external_links(wb2)
            self.assertEqual(removed, 1)
            clean = tmpdir / "sin_links.xlsx"
            wb2.save(clean)

            with zipfile.ZipFile(clean) as z:
                self.assertFalse(any("external" in n.lower() for n in z.namelist()))
                self.assertNotIn("externalReference", z.read("xl/workbook.xml").decode("utf-8", "ignore"))


if __name__ == "__main__":
    unittest.main()
