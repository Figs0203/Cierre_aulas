"""
Robustez del cierre directo ante bloqueos y re-ejecuciones.

Cubre:
  1. Detección previa de archivos abiertos en Excel / solo lectura.
  2. Reversión automática si un paso falla a mitad del cierre.
  3. Corrección opcional de una nota final ya escrita (re-ejecución).
"""

from __future__ import annotations

import os
import stat
import tempfile
import unittest
import zipfile
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
from output.direct_updater import apply_direct_cierre, files_not_writable
from output.xlsx_patcher import XlsxPatcher


def _tv(valor, col: int = 1, fila: int = 1):
    return TracedValue(valor, "S", "Formato", "col", col, fila)


def _crear_archivos(tmpdir: Path, promedio_existente=4.5) -> dict:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Formato"
    ws.append([
        "Nombres", "Apellidos", "Clase", "Documento", "Correo",
        "Módulo 1", "Módulo 2", "Promedio curso completo", "Estado",
    ])
    # El promedio ya está escrito (como en un cierre anterior)
    ws.append([
        "ANA", "PEREZ", "6401", "1001", "ana@eafit.edu.co",
        5, 4, promedio_existente, None,
    ])
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


def _resultado(promedio=4.2) -> tuple[AulaCierreResult, dict]:
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
            "correo": 5, "modulo_1": 6, "modulo_2": 7, "promedio": 8, "estado": 9,
        },
        raw_row_data={6: 5.0, 7: 4.0, 8: promedio, 9: "Aprobó"},
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
    return result, meta


class TestCierreRobusto(unittest.TestCase):
    def test_detecta_archivo_de_solo_lectura_antes_de_escribir(self):
        with tempfile.TemporaryDirectory() as tmp_str:
            tmpdir = Path(tmp_str)
            files = _crear_archivos(tmpdir)
            os.chmod(files["sistematizacion"], stat.S_IREAD)
            try:
                mensajes = files_not_writable(files)
                self.assertTrue(mensajes)
                self.assertIn("Sistematizacion_Cursos_COIN_2026.xlsx", mensajes[0])

                result, meta = _resultado()
                report = apply_direct_cierre(
                    files, result, "ABBUEI227", metadatos_control=meta,
                    start_consecutivo=1, backup_dir=tmpdir / "bk",
                )
                self.assertFalse(report.is_successful)
                self.assertTrue(report.errors)
                # No se creó respaldo ni se modificó nada
                self.assertFalse((tmpdir / "bk").exists())
            finally:
                os.chmod(files["sistematizacion"], stat.S_IWRITE)

    def test_detecta_archivo_de_bloqueo_de_excel(self):
        """Excel deja '~$<nombre>.xlsx' mientras el libro está abierto."""
        with tempfile.TemporaryDirectory() as tmp_str:
            tmpdir = Path(tmp_str)
            files = _crear_archivos(tmpdir)
            bloqueo = files["sistematizacion"].with_name(
                "~$" + files["sistematizacion"].name
            )
            bloqueo.write_text("", encoding="utf-8")
            try:
                mensajes = files_not_writable(files)
                self.assertTrue(mensajes)
                self.assertIn("ABIERTO", mensajes[0])
            finally:
                bloqueo.unlink()

    def test_detecta_libro_compartido(self):
        """Los libros compartidos generan 'Conflicto de combinación'."""
        with tempfile.TemporaryDirectory() as tmp_str:
            tmpdir = Path(tmp_str)
            files = _crear_archivos(tmpdir)
            path = files["sistematizacion"]
            with zipfile.ZipFile(path) as zf:
                parts = {n: zf.read(n) for n in zf.namelist()}
            wb = parts["xl/workbook.xml"].decode("utf-8")
            wb = wb.replace(
                "</workbook>",
                '<fileSharing readOnlyRecommended="0"/></workbook>',
            )
            parts["xl/workbook.xml"] = wb.encode("utf-8")
            with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
                for name, data in parts.items():
                    zf.writestr(name, data)

            mensajes = files_not_writable(files)
            self.assertTrue(
                any("LIBRO COMPARTIDO" in m for m in mensajes), mensajes
            )

    def test_revierte_si_un_paso_falla(self):
        with tempfile.TemporaryDirectory() as tmp_str:
            tmpdir = Path(tmp_str)
            files = _crear_archivos(tmpdir)
            antes_sist = files["sistematizacion"].read_bytes()
            antes_cert = files["certificados"].read_bytes()
            antes_ctrl = files["control_aulas"].read_bytes()

            result, meta = _resultado()
            guardar_original = XlsxPatcher.save

            def guardar_que_falla_en_control(self, *args, **kwargs):
                if self.path.name.startswith("Control_Aulas"):
                    raise PermissionError("simulado: archivo abierto en Excel")
                return guardar_original(self, *args, **kwargs)

            with mock.patch.object(XlsxPatcher, "save", guardar_que_falla_en_control):
                report = apply_direct_cierre(
                    files, result, "ABBUEI227", metadatos_control=meta,
                    start_consecutivo=1, backup_dir=tmpdir / "bk",
                )

            self.assertFalse(report.is_successful)
            self.assertTrue(any("revertida" in w.lower() for w in report.warnings))
            # Los tres oficiales vuelven al estado previo al cierre
            self.assertEqual(files["sistematizacion"].read_bytes(), antes_sist)
            self.assertEqual(files["certificados"].read_bytes(), antes_cert)
            self.assertEqual(files["control_aulas"].read_bytes(), antes_ctrl)

    def test_no_corrige_la_nota_final_existente_por_defecto(self):
        with tempfile.TemporaryDirectory() as tmp_str:
            tmpdir = Path(tmp_str)
            files = _crear_archivos(tmpdir, promedio_existente=4.5)
            result, meta = _resultado(promedio=4.2)

            report = apply_direct_cierre(
                files, result, "ABBUEI227", metadatos_control=meta,
                start_consecutivo=1, backup_dir=tmpdir / "bk",
            )
            self.assertTrue(report.is_successful, report.errors)
            self.assertEqual(report.sistematizacion_promedios_actualizados, 0)

            wb = openpyxl.load_workbook(files["sistematizacion"])
            self.assertEqual(wb["Formato"].cell(2, 8).value, 4.5)
            wb.close()

    def test_corrige_la_nota_final_si_se_autoriza(self):
        with tempfile.TemporaryDirectory() as tmp_str:
            tmpdir = Path(tmp_str)
            files = _crear_archivos(tmpdir, promedio_existente=4.5)
            result, meta = _resultado(promedio=4.2)

            report = apply_direct_cierre(
                files, result, "ABBUEI227", metadatos_control=meta,
                start_consecutivo=1, backup_dir=tmpdir / "bk",
                actualizar_nota_final_existente=True,
            )
            self.assertTrue(report.is_successful, report.errors)
            self.assertEqual(report.sistematizacion_promedios_actualizados, 1)

            wb = openpyxl.load_workbook(files["sistematizacion"])
            self.assertEqual(wb["Formato"].cell(2, 8).value, 4.2)
            wb.close()


if __name__ == "__main__":
    unittest.main()
