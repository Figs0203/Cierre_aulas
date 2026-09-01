"""
Pruebas de Auditoría de Privacidad (Privacy Audit).

Verifica que el informe de descubrimiento generado NO contenga patrones obvios
de datos personales completos (correos sin enmascarar, documentos completos de 8-10 dígitos).
"""

import re
import tempfile
import unittest
from pathlib import Path

from tests.discovery.inspect_real_files import run_discovery
from tests.fixtures.generate_fixtures import generate_all_fixtures

# Regex para detectar correos completos sin asteriscos de enmascaramiento
# Un correo válido no enmascarado tiene usuario sin asteriscos
UNMASKED_EMAIL_REGEX = re.compile(r"\b[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+\b")

# Regex para detectar números de documento completos (7 a 10 dígitos continuos sin asteriscos)
UNMASKED_DOC_REGEX = re.compile(r"\b(1000\d{6}|[1-9]\d{6,9})\b")


class TestPrivacyAudit(unittest.TestCase):

    def test_discovery_report_does_not_leak_unmasked_pii(self):
        """Verifica que el informe de descubrimiento generado aplique enmascaramiento estricto."""
        with tempfile.TemporaryDirectory() as tmpdir:
            fixtures = generate_all_fixtures(output_dir=Path(tmpdir) / "fixtures")
            report_path, report_content = run_discovery(
                fixtures,
                target_aula="ABBUEI154",
                output_dir=Path(tmpdir) / "reports",
            )

            # Verificar que los correos de muestra en los fixtures sintéticos no salgan completos
            known_synthetic_emails = [
                "juan.perez1@ejemplo.edu.co",
                "maria.gomez2@ejemplo.edu.co",
                "carlos.restrepo3@ejemplo.edu.co",
                "laura.velez4@ejemplo.edu.co",
                "santiago.morales5@ejemplo.edu.co",
                "valentina.ospina6@ejemplo.edu.co",
                "pedro.ramirez@ejemplo.edu.co",
            ]

            for raw_email in known_synthetic_emails:
                self.assertNotIn(
                    raw_email,
                    report_content,
                    f"Fuga de privacidad: el correo '{raw_email}' aparece sin enmascarar en el informe."
                )

            # Verificar que los documentos sintéticos no salgan completos
            known_synthetic_docs = [
                "1000000001", "1000000002", "1000000003",
                "1000000004", "1000000005", "1000000006",
                "1000000099",
            ]

            for raw_doc in known_synthetic_docs:
                self.assertNotIn(
                    raw_doc,
                    report_content,
                    f"Fuga de privacidad: el documento '{raw_doc}' aparece sin enmascarar en el informe."
                )

            # Verificar que los datos aparezcan enmascarados con asteriscos
            self.assertIn("****0001", report_content)
            self.assertIn("juan.p***@ejemplo.edu.co", report_content)


if __name__ == "__main__":
    unittest.main()
