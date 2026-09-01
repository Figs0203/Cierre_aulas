"""
Pruebas de verificación de arquitectura Offline y ausencia de red.

Verifica que ningún archivo del proyecto importe módulos de red ni dependencias externas
no autorizadas, asegurando que la herramienta opera de forma 100% aislada.
"""

import ast
import os
import unittest
from pathlib import Path

_PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent

PROHIBITED_IMPORTS = {
    "urllib", "requests", "http", "socket", "ftplib", "smtplib",
    "xmlrpc", "telnetlib", "ssl", "aiohttp", "httpx", "boto3",
    "google", "azure", "firebase",
}


class TestOfflineArchitecture(unittest.TestCase):

    def test_no_prohibited_imports_in_codebase(self):
        """Escanea todos los archivos .py del proyecto para asegurar que no importen red."""
        py_files = list(_PROJECT_ROOT.glob("**/*.py"))
        violations = []

        for py_file in py_files:
            # Omitir entornos virtuales si existieran
            if "venv" in py_file.parts or ".venv" in py_file.parts:
                continue

            try:
                tree = ast.parse(py_file.read_text(encoding="utf-8"))
            except Exception:
                continue

            for node in ast.walk(tree):
                if isinstance(node, ast.Import):
                    for alias in node.names:
                        mod = alias.name.split(".")[0]
                        if mod in PROHIBITED_IMPORTS:
                            violations.append(f"{py_file.name}: import {alias.name}")
                elif isinstance(node, ast.ImportFrom):
                    if node.module:
                        mod = node.module.split(".")[0]
                        if mod in PROHIBITED_IMPORTS:
                            violations.append(f"{py_file.name}: from {node.module} import ...")

        self.assertEqual(
            len(violations), 0,
            f"Se detectaron importaciones de red prohibidas:\n" + "\n".join(violations)
        )


if __name__ == "__main__":
    unittest.main()
