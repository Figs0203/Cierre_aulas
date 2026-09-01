"""
Pruebas de integración para el Motor de Descubrimiento Local (Fase 0).

Ejecuta el descubrimiento completo sobre los fixtures sintéticos y verifica:
1. Inmutabilidad: Los hashes SHA-256 coinciden exactamente antes y después.
2. Detección estructural: Identifica hojas, encabezados y módulos.
3. Detección de fórmulas: Encuentra la fórmula de Estado para certificación.
4. Detección de colores: Identifica las celdas con relleno naranja y verde.
5. Privacidad: Los datos personales quedan enmascarados en el reporte generado.
"""

import tempfile
import unittest
from pathlib import Path

from security.integrity import compute_sha256, verify_integrity
from tests.discovery.inspect_real_files import run_discovery
from tests.fixtures.generate_fixtures import generate_all_fixtures


class TestDiscoveryIntegration(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.fixtures = generate_all_fixtures()

    def test_discovery_end_to_end(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            # 1. Calcular hashes iniciales
            hashes_before = {k: compute_sha256(v) for k, v in self.fixtures.items()}

            # 2. Ejecutar descubrimiento con aula objetivo
            report_path, report_content = run_discovery(
                file_paths=self.fixtures,
                target_aula="ABBUEI154",
                output_dir=tmpdir,
            )

            # Verificar que el reporte fue creado
            self.assertTrue(report_path.exists())
            self.assertGreater(len(report_content), 1000)

            # 3. Verificar inmutabilidad post-ejecución
            intact, msgs = verify_integrity(hashes_before, self.fixtures)
            self.assertTrue(intact, "Los archivos originales fueron modificados durante la lectura")

            # 4. Verificar secciones clave en el reporte
            self.assertIn("INFORME DE DESCUBRIMIENTO TÉCNICO LOCAL", report_content)
            self.assertIn("INTEGRIDAD CRIPTOGRÁFICA DE ARCHIVOS FUENTE", report_content)
            self.assertIn("ESTRUCTURA TÉCNICA DETALLADA POR ARCHIVO", report_content)
            self.assertIn("ENCABEZADOS ENCONTRADOS Y MAPEO SEMÁNTICO PRELIMINAR", report_content)

            # Verificar que detectó módulos
            self.assertIn("modulo_1", report_content)
            self.assertIn("modulo_2", report_content)
            self.assertIn("modulo_3", report_content)
            self.assertIn("modulo_4", report_content)

            # Verificar que detectó fórmulas (especialmente Estado para certificación)
            self.assertIn("FÓRMULAS DETECTADAS Y ANÁLISIS DE CERTIFICACIÓN", report_content)
            self.assertIn("Estado para certificación", report_content)

            # Verificar que detectó colores de relleno (naranja / verde)
            self.assertIn("COLORES DE RELLENO (FILLS) Y FORMATO CONDICIONAL", report_content)
            self.assertIn("RGB=#FFA500", report_content)  # Naranja sintético
            self.assertIn("RGB=#C6EFCE", report_content)  # Verde sintético

            # Verificar búsqueda del aula objetivo
            self.assertIn("CLASES Y SECCIONES DETECTADAS", report_content)
            self.assertIn("ABBUEI154", report_content)

            # Verificar privacidad: los correos deben estar enmascarados
            self.assertNotIn("juan.perez1@ejemplo.edu.co", report_content)
            self.assertIn("juan.p***@ejemplo.edu.co", report_content)


if __name__ == "__main__":
    unittest.main()
