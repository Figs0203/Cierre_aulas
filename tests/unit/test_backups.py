"""
Respaldo único por archivo oficial.

Cada cierre debe dejar **exactamente un** respaldo por archivo en
``_backups_cierre/``, actualizado con el estado previo a esa ejecución. Los
respaldos antiguos con marca de tiempo del mismo archivo se consolidan.
"""

from __future__ import annotations

import shutil
import tempfile
import unittest
from pathlib import Path

from output.direct_updater import (
    apply_direct_cierre,
    backup_path_for,
    create_timestamped_backup,
    find_latest_backup,
)
from processing.cierre_orchestrator import run_cierre_pipeline
from tests.fixtures.generate_fixtures import generate_all_fixtures

_ROLES = ("sistematizacion", "certificados", "control_aulas")


class TestRespaldoUnico(unittest.TestCase):
    def test_find_latest_backup_prefiere_el_unico(self):
        with tempfile.TemporaryDirectory() as tmp_str:
            tmpdir = Path(tmp_str)
            original = tmpdir / "Archivo.xlsx"
            original.write_bytes(b"original")
            backup_dir = tmpdir / "_backups_cierre"
            backup_dir.mkdir()

            antiguos = [
                backup_dir / "Archivo_backup_20260101_010101.xlsx",
                backup_dir / "Archivo_backup_20260201_010101.xlsx",
            ]
            for p in antiguos:
                p.write_bytes(b"viejo")
            self.assertEqual(
                find_latest_backup(original, backup_dir), antiguos[-1].resolve()
            )

            unico = create_timestamped_backup(original, backup_dir=backup_dir)
            self.assertEqual(find_latest_backup(original, backup_dir), unico)
            self.assertEqual(unico.read_bytes(), b"original")
            self.assertEqual(len(list(backup_dir.iterdir())), 1)

    def test_cada_cierre_deja_un_solo_respaldo_por_archivo(self):
        fixtures = generate_all_fixtures()
        with tempfile.TemporaryDirectory() as tmp_str:
            tmpdir = Path(tmp_str)
            copias = {}
            for role, ruta in fixtures.items():
                destino = tmpdir / Path(ruta).name
                shutil.copy2(ruta, destino)
                copias[role] = destino

            # Estado inicial de los oficiales (Notas no se respalda)
            antes_1 = {r: copias[r].read_bytes() for r in _ROLES}

            result, _, _, _ = run_cierre_pipeline(
                file_paths=copias,
                target_aula="ABBUEI154",
                codigo_curso="ABBUEI154",
                is_simulation=True,
            )
            report_1 = apply_direct_cierre(
                copias, result, "ABBUEI154",
                metadatos_control=result.metadatos_control,
                start_consecutivo=result.start_consecutivo,
            )
            self.assertTrue(report_1.is_successful, report_1.errors)

            for role in _ROLES:
                backup_dir = copias[role].parent / "_backups_cierre"
                unico = backup_path_for(copias[role], backup_dir)
                self.assertTrue(unico.exists(), f"{role}: falta el respaldo único")
                # No deben quedar copias con marca de tiempo de ESTE archivo.
                legados = [
                    p for p in backup_dir.iterdir()
                    if p.name.startswith(copias[role].stem + "_backup_")
                ]
                self.assertEqual(legados, [], f"{role}: quedaron respaldos antiguos")
                self.assertEqual(
                    unico.read_bytes(), antes_1[role],
                    f"{role}: el respaldo debe ser el estado previo al cierre",
                )

            # Segundo cierre: el respaldo se ACTUALIZA con el estado actual
            antes_2 = {r: copias[r].read_bytes() for r in _ROLES}
            result_2, _, _, _ = run_cierre_pipeline(
                file_paths=copias,
                target_aula="ABBUEI154",
                codigo_curso="ABBUEI154",
                is_simulation=True,
            )
            report_2 = apply_direct_cierre(
                copias, result_2, "ABBUEI154",
                metadatos_control=result_2.metadatos_control,
                start_consecutivo=result_2.start_consecutivo,
            )
            self.assertTrue(report_2.is_successful, report_2.errors)

            for role in _ROLES:
                backup_dir = copias[role].parent / "_backups_cierre"
                unico = backup_path_for(copias[role], backup_dir)
                self.assertTrue(unico.exists())
                self.assertEqual(unico.read_bytes(), antes_2[role])
                legados = [
                    p for p in backup_dir.iterdir()
                    if p.name.startswith(copias[role].stem + "_backup_")
                ]
                self.assertEqual(legados, [])


if __name__ == "__main__":
    unittest.main()
