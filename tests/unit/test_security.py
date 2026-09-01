"""
Pruebas unitarias para los módulos de seguridad e integridad.
"""

import tempfile
import unittest
from pathlib import Path

from security.integrity import compute_sha256, verify_integrity
from security.privacy import mask_email, mask_document, mask_name, mask_full_name, mask_for_log
from security.validation_files import (
    validate_no_duplicate_files,
    validate_not_auxiliary_file,
    validate_file_extension,
    validate_file_exists,
    validate_all_input_files,
)


class TestSecurity(unittest.TestCase):

    def test_sha256_computation_and_verification(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            f1 = Path(tmpdir) / "test1.xlsx"
            f1.write_bytes(b"CONTENIDO DE PRUEBA 12345")

            record = compute_sha256(f1)
            self.assertEqual(record.file_name, "test1.xlsx")
            self.assertEqual(record.file_size, len(b"CONTENIDO DE PRUEBA 12345"))
            self.assertEqual(len(record.sha256_hash), 64)

            # Verificación intacta
            records = {"rol1": record}
            paths = {"rol1": f1}
            intact, msgs = verify_integrity(records, paths)
            self.assertTrue(intact)
            self.assertTrue(any("Integridad verificada" in m for m in msgs))

            # Modificación del archivo
            f1.write_bytes(b"CONTENIDO MODIFICADO")
            intact, msgs = verify_integrity(records, paths)
            self.assertFalse(intact)
            self.assertTrue(any("INTEGRIDAD COMPROMETIDA" in m for m in msgs))

    def test_privacy_masking(self):
        # Emails
        self.assertEqual(mask_email("juan.perez@eafit.edu.co"), "juan.p***@eafit.edu.co")
        self.assertEqual(mask_email("a@b.com"), "a***@b.com")
        self.assertEqual(mask_email(None), "")

        # Documentos
        self.assertEqual(mask_document("1020304050"), "****4050")
        self.assertEqual(mask_document("1.020.304.050"), "****4050")
        self.assertEqual(mask_document("123"), "****123")
        self.assertEqual(mask_document(None), "")

        # Nombres
        self.assertEqual(mask_name("Carlos Alberto Restrepo"), "Carlos R.")
        self.assertEqual(mask_name("Juan"), "Juan")
        self.assertEqual(mask_name(None), "")

        # Nombres combinados
        self.assertEqual(mask_full_name("Juan Carlos", "Perez Gomez"), "Juan P.")

        # Logs
        log_doc = mask_for_log(document="1020304050")
        self.assertIn("****4050", log_doc)

    def test_file_validation(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            f1 = Path(tmpdir) / "Control_2026.xlsx"
            f2 = Path(tmpdir) / "Sistematizacion_2026.xlsx"
            f1.write_bytes(b"Archivo 1")
            f2.write_bytes(b"Archivo 2")

            # Archivo válido
            self.assertEqual(len(validate_file_exists(f1)), 0)
            self.assertEqual(len(validate_file_extension(f1)), 0)
            self.assertEqual(len(validate_not_auxiliary_file(f1)), 0)

            # Archivo auxiliar prohibido
            faux = Path(tmpdir) / "Cierre_Aula_ABBUEI154.xlsx"
            faux.write_bytes(b"Auxiliar")
            self.assertGreater(len(validate_not_auxiliary_file(faux)), 0)

            # Extensión inválida
            fcsv = Path(tmpdir) / "archivo.csv"
            fcsv.write_bytes(b"csv")
            self.assertGreater(len(validate_file_extension(fcsv)), 0)

            # Archivo inexistente
            fnox = Path(tmpdir) / "no_existe.xlsx"
            self.assertGreater(len(validate_file_exists(fnox)), 0)

            # Archivos duplicados (mismo path)
            dups = validate_no_duplicate_files({"rol1": f1, "rol2": f1})
            self.assertGreater(len(dups), 0)


if __name__ == "__main__":
    unittest.main()
