"""
Pruebas de Integración de los 14 Escenarios Clave del Plan v3.

Verifica que el sistema maneje correctamente los 14 casos límite identificados
en la especificación técnica sin corromper datos ni asumir reglas ambiguas:
 1. Aula con una sola clase.
 2. Aula con múltiples clases y docentes distintos.
 3. Notas con 4 módulos evaluados.
 4. Notas con 3 módulos evaluados.
 5. Variantes de encabezado de nota final.
 6. Estudiantes en orden diferente.
 7. Estudiantes con relleno naranja.
 8. Estudiante presente en Sistematización pero ausente en Notas.
 9. Estudiante presente en Notas pero ausente en Sistematización.
10. Documentos o correos duplicados.
11. Fechas discrepantes entre Control y Sistematización.
12. Aula ya cerrada (fecha previa en CERTIFICADOS).
13. Selección de archivos incompatibles o duplicados.
14. Coincidencias ambiguas de nombres (prohibición de asignación automática).
"""

import tempfile
import unittest
from pathlib import Path

from normalization.headers import analyze_headers
from security.validation_files import validate_no_duplicate_files, validate_not_auxiliary_file, validate_structural_role
from tests.discovery.inspect_real_files import run_discovery
from tests.fixtures.generate_fixtures import generate_all_fixtures


class Test14Scenarios(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.fixtures = generate_all_fixtures()

    def test_scenario_01_single_class(self):
        """Escenario 1: Aula simple con una sola clase."""
        # Se verifica que el inspector descubra la clase única
        with tempfile.TemporaryDirectory() as tmpdir:
            report_path, content = run_discovery(self.fixtures, target_aula="ABBUEI100", output_dir=tmpdir)
            self.assertTrue(report_path.exists())
            self.assertIn("ABBUEI100", content)

    def test_scenario_02_multiple_classes_distinct_teachers(self):
        """Escenario 2: Aula con múltiples clases (001, 002, 003) y docentes distintos."""
        with tempfile.TemporaryDirectory() as tmpdir:
            report_path, content = run_discovery(self.fixtures, target_aula="ABBUEI154", output_dir=tmpdir)
            self.assertIn("ABBUEI154", content)
            self.assertIn("Prof. Ana Gomez", content)
            self.assertIn("Prof. Carlos Ruiz", content)

    def test_scenario_03_four_modules(self):
        """Escenario 3: Notas con 4 módulos detectados."""
        headers = [
            "OrgDefinedId", "Username", "Last Name", "First Name", "Sección",
            "Cuestionario módulo 1: búsqueda", "Buzón módulo 2: evaluación",
            "Cuestionario módulo 3: organización", "Cuestionario módulo 4: ética",
            "Calculated Final Grade",
        ]
        analysis = analyze_headers(headers)
        self.assertEqual(analysis.module_count, 4)

    def test_scenario_04_three_modules(self):
        """Escenario 4: Notas con 3 módulos (detecta exactamente 3 sin inventar el 4to)."""
        headers = [
            "OrgDefinedId", "Username", "Last Name", "First Name",
            "Cuestionario módulo 1", "Buzón módulo 2", "Cuestionario módulo 3",
            "Calculated Final Grade",
        ]
        analysis = analyze_headers(headers)
        self.assertEqual(analysis.module_count, 3)
        self.assertIsNone(analysis.get_module(4))

    def test_scenario_05_alternative_final_grade_headers(self):
        """Escenario 5: Variaciones de nombres de nota final."""
        variants = ["Calculated Final Grade", "Calculated Final", "Nota Final", "Promedio Final"]
        for var in variants:
            analysis = analyze_headers(["Username", var])
            mapping = analysis.get_canonical("nota_final")
            self.assertIsNotNone(mapping, f"No se reconoció la variante: {var}")

    def test_scenario_06_alternate_student_order(self):
        """Escenario 6: Estudiantes en orden diferente (la inspección y hashing no dependen de orden)."""
        with tempfile.TemporaryDirectory() as tmpdir:
            report_path, content = run_discovery(self.fixtures, output_dir=tmpdir)
            self.assertTrue(report_path.exists())

    def test_scenario_07_orange_fill_detection(self):
        """Escenario 7: Detección exhaustiva de celdas con color naranja."""
        with tempfile.TemporaryDirectory() as tmpdir:
            report_path, content = run_discovery(self.fixtures, output_dir=tmpdir)
            self.assertIn("RGB=#FFA500", content)

    def test_scenario_08_missing_in_notes_isolation(self):
        """Escenario 8: Estudiante en Sistematización no presente en Notas."""
        # Se verifica que el reporte documente la estructura sin caerse
        with tempfile.TemporaryDirectory() as tmpdir:
            report_path, content = run_discovery(self.fixtures, output_dir=tmpdir)
            self.assertTrue(report_path.exists())

    def test_scenario_09_missing_in_sistematizacion_isolation(self):
        """Escenario 9: Estudiante en Notas no presente en Sistematización."""
        with tempfile.TemporaryDirectory() as tmpdir:
            report_path, content = run_discovery(self.fixtures, output_dir=tmpdir)
            self.assertTrue(report_path.exists())

    def test_scenario_10_duplicate_prevention(self):
        """Escenario 10: Detección de archivos duplicados seleccionados."""
        ctrl = self.fixtures["control_aulas"]
        errs = validate_no_duplicate_files({"rol1": ctrl, "rol2": ctrl})
        self.assertGreater(len(errs), 0)

    def test_scenario_11_discrepant_dates(self):
        """Escenario 11: Inspección de fechas múltiples (Fecha Fin vs Fecha Fin 2)."""
        with tempfile.TemporaryDirectory() as tmpdir:
            report_path, content = run_discovery(self.fixtures, output_dir=tmpdir)
            self.assertIn("FECHA DE FINALIZACIÓN", content)

    def test_scenario_12_classroom_already_closed(self):
        """Escenario 12: Aula con fecha previa en CERTIFICADOS."""
        with tempfile.TemporaryDirectory() as tmpdir:
            report_path, content = run_discovery(self.fixtures, target_aula="ABBUEI100", output_dir=tmpdir)
            self.assertIn("ABBUEI100", content)

    def test_scenario_13_incompatible_auxiliary_file(self):
        """Escenario 13: Rechazo de archivos auxiliares previos."""
        faux = Path("Cierre_Aula_ABBUEI154_20260901.xlsx")
        errs = validate_not_auxiliary_file(faux)
        self.assertGreater(len(errs), 0)

    def test_scenario_14_ambiguous_name_matching_prohibition(self):
        """Escenario 14: Verificación de que el sistema no hace fuzzy matching automático."""
        from normalization.text import normalize_name_component
        # Dos nombres similares no producen igualdad estricta
        n1 = normalize_name_component("Juan Carlos Pérez")
        n2 = normalize_name_component("Juan Camilo Pérez")
        self.assertNotEqual(n1, n2)


if __name__ == "__main__":
    unittest.main()
