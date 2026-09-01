"""
Pruebas de verificación de Inmutabilidad Absoluta y Ausencia de Escritura sobre Fuentes.

Verifica:
1. Que el inspector de descubrimiento y las funciones de lectura nunca modifiquen los archivos de entrada.
2. Que en todo el código base no exista ninguna llamada a .save() sobre archivos o rutas de entrada.
"""

import ast
import tempfile
import unittest
from pathlib import Path

from security.integrity import compute_sha256, verify_integrity
from tests.discovery.inspect_real_files import run_discovery
from tests.fixtures.generate_fixtures import generate_all_fixtures

_PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent


class TestImmutability(unittest.TestCase):

    def test_no_save_on_input_files_in_codebase(self):
        """Audita que ningún archivo en input/, parsing/, security/, normalization/ llame a .save()."""
        scanned_dirs = ["security", "normalization", "input", "parsing", "core", "config"]
        violations = []

        for sdir in scanned_dirs:
            dir_path = _PROJECT_ROOT / sdir
            if not dir_path.exists():
                continue
            for py_file in dir_path.glob("*.py"):
                tree = ast.parse(py_file.read_text(encoding="utf-8"))
                for node in ast.walk(tree):
                    if isinstance(node, ast.Call):
                        if isinstance(node.func, ast.Attribute) and node.func.attr == "save":
                            violations.append(f"{py_file.name}: llamada a .save() detectada")

        self.assertEqual(
            len(violations), 0,
            f"Se detectaron llamadas a .save() no autorizadas en módulos de solo lectura:\n"
            + "\n".join(violations)
        )

    def test_sha256_strictly_identical_after_discovery(self):
        """Verifica que tras ejecutar el descubrimiento completo los 4 hashes sigan siendo idénticos."""
        with tempfile.TemporaryDirectory() as tmpdir:
            fixtures = generate_all_fixtures(output_dir=Path(tmpdir) / "fixtures")
            hashes_before = {k: compute_sha256(v) for k, v in fixtures.items()}

            # Ejecutar descubrimiento
            report_path, _ = run_discovery(fixtures, target_aula="ABBUEI154", output_dir=Path(tmpdir) / "reports")

            # Verificar inmutabilidad
            intact, msgs = verify_integrity(hashes_before, fixtures)
            self.assertTrue(intact)

            for role, path in fixtures.items():
                hash_after = compute_sha256(path)
                self.assertEqual(
                    hashes_before[role].sha256_hash,
                    hash_after.sha256_hash,
                    f"El hash SHA-256 del archivo {role} cambió después del descubrimiento."
                )


if __name__ == "__main__":
    unittest.main()
