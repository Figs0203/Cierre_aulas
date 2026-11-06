"""
Cierre por lotes: varias aulas seguidas sobre la misma Sistematización y Control.

Verifica que:
  1. El código de aula se pueda inferir del nombre de las Notas.
  2. Aplicar dos aulas en secuencia sobre archivos compartidos funciona.
  3. La numeración de certificados continúa correctamente entre aulas.
  4. El lote se aborta de forma limpia si un archivo compartido está bloqueado.
"""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest import mock

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
from main import _infer_aula_code, handle_batch_cierre_mode
from output.direct_updater import apply_direct_cierre


def _tv(valor, col: int = 1, fila: int = 1):
    return TracedValue(valor, "S", "Formato", "col", col, fila)


def _crear_archivos(tmpdir: Path) -> dict:
    # Sistematización con dos aulas en clases distintas
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Formato"
    ws.append([
        "Nombres", "Apellidos", "Clase", "Documento", "Correo",
        "Módulo 1", "Módulo 2", "Promedio curso completo", "Estado",
    ])
    ws.append(["ANA", "PEREZ", "6401", "1001", "ana@eafit.edu.co", 5, 4, None, None])
    ws.append(["LUIS", "GOMEZ", "6402", "1002", "luis@eafit.edu.co", 4, 4, None, None])
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
    ws.append(["FINANZAS", "FN02", "6402", "DOCENTE", "16/03/2026", "27/04/2026", None])
    path_ctrl = tmpdir / "Control_Aulas_Formadores_2026.xlsx"
    wb.save(path_ctrl)
    wb.close()

    return {
        "sistematizacion": path_sist,
        "certificados": path_cert,
        "control_aulas": path_ctrl,
        "notas": path_cert,
    }


def _resultado(fila: int, clase: str, doc: str, nombre: str, nota: float, curso: str):
    es = EstudianteSistematizacion(
        row_number=fila,
        nombres=_tv(nombre, 1, fila),
        apellidos=_tv("APELLIDO", 2, fila),
        clase=_tv(clase, 3, fila),
        documento=_tv(doc, 4, fila),
        correo=_tv(f"{doc}@eafit.edu.co", 5, fila),
        estado_calculado="Aprobó",
        column_indices={
            "nombres": 1, "apellidos": 2, "clase": 3, "documento": 4, "correo": 5,
            "modulo_1": 6, "modulo_2": 7, "promedio": 8, "estado": 9,
        },
        raw_row_data={6: 5.0, 7: 4.0, 8: nota, 9: "Aprobó"},
    )
    match = StudentMatch(
        estudiante_sist=es,
        estudiante_nota=EstudianteNotas(row_number=fila),
        match_key="doc",
        match_value=doc,
    )
    meta = {
        clase: AulaMetadata(
            programa_academico=_tv("FINANZAS"),
            catalogo=_tv("FN01"),
            docente_principal=_tv("DOCENTE"),
            fecha_inicio=_tv("16/03/2026"),
            fecha_fin=_tv("27/04/2026"),
        )
    }
    result = AulaCierreResult(
        aula=AulaMetadata(),
        clases=[ClaseGroup(clase_id=clase, matches=[match])],
        metadatos_control=meta,
        start_consecutivo=1,  # simula el consecutivo calculado antes de escribir
    )
    return result, meta


class TestLote(unittest.TestCase):
    def test_infiere_codigo_de_aula_desde_el_nombre(self):
        self.assertEqual(
            _infer_aula_code(Path("ABBUEI227_E_FINANZAS_MAF_1.xlsx")), "ABBUEI227"
        )
        self.assertEqual(
            _infer_aula_code(Path("CMI122_PSICOLOGIA_2.xlsx")), "CMI122"
        )
        self.assertEqual(_infer_aula_code(Path("notas_sin_codigo.xlsx")), "")

    def test_dos_aulas_en_secuencia_comparten_sistematizacion_y_control(self):
        with tempfile.TemporaryDirectory() as tmp_str:
            tmpdir = Path(tmp_str)
            files = _crear_archivos(tmpdir)

            result1, meta1 = _resultado(2, "6401", "1001", "ANA", 4.2, "ABBUEI227")
            result2, meta2 = _resultado(3, "6402", "1002", "LUIS", 3.8, "ABBUEI228")

            rep1 = apply_direct_cierre(
                files, result1, "ABBUEI227", metadatos_control=meta1,
                start_consecutivo=1, backup_dir=tmpdir / "bk",
            )
            self.assertTrue(rep1.is_successful, rep1.errors)

            rep2 = apply_direct_cierre(
                files, result2, "ABBUEI228", metadatos_control=meta2,
                start_consecutivo=1, backup_dir=tmpdir / "bk",
            )
            self.assertTrue(rep2.is_successful, rep2.errors)

            # Certificados: dos filas, numeración continuada 1 y 2
            wb = openpyxl.load_workbook(files["certificados"])
            ws = wb["Códigos"]
            filas = [
                (ws.cell(r, 1).value, str(ws.cell(r, 6).value), ws.cell(r, 8).value)
                for r in range(2, ws.max_row + 1)
                if ws.cell(r, 1).value is not None
            ]
            wb.close()
            self.assertEqual(
                filas,
                [(1, "1001", "ABBUEI227"), (2, "1002", "ABBUEI228")],
            )

            # Sistematización: cada aula conserva su nota final exacta
            wb = openpyxl.load_workbook(files["sistematizacion"])
            ws = wb["Formato"]
            self.assertEqual(ws.cell(2, 8).value, 4.2)
            self.assertEqual(ws.cell(3, 8).value, 3.8)
            wb.close()

            # Control: las dos clases quedaron marcadas
            wb = openpyxl.load_workbook(files["control_aulas"])
            ws = wb["Tabla_ABBUEI"]
            self.assertIsNotNone(ws.cell(2, 7).value)
            self.assertIsNotNone(ws.cell(3, 7).value)
            wb.close()

            # Un solo respaldo por archivo (política de respaldo único)
            backups = tmpdir / "bk"
            self.assertTrue(
                (backups / "Sistematizacion_Cursos_COIN_2026_backup.xlsx").exists()
            )
            self.assertTrue(
                (backups / "Control_codigos_certificados_ABBUEI_backup.xlsx").exists()
            )
            self.assertTrue(
                (backups / "Control_Aulas_Formadores_2026_backup.xlsx").exists()
            )

    def test_el_lote_aborta_si_los_compartidos_estan_bloqueados(self):
        with tempfile.TemporaryDirectory() as tmp_str:
            tmpdir = Path(tmp_str)
            files = _crear_archivos(tmpdir)

            # Se simula que el Control compartido está bloqueado (solo lectura)
            import os
            import stat
            os.chmod(files["control_aulas"], stat.S_IREAD)
            try:
                with mock.patch(
                    "main.prompt_for_file",
                    side_effect=[files["control_aulas"], files["sistematizacion"]],
                ), mock.patch("main.prompt_for_multiple_files") as multi:
                    handle_batch_cierre_mode()
                    multi.assert_not_called()
            finally:
                os.chmod(files["control_aulas"], stat.S_IWRITE)


if __name__ == "__main__":
    unittest.main()
