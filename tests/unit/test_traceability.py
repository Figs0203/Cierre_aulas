"""
Pruebas unitarias para el sistema de trazabilidad y log de validaciones.
"""

import unittest

from core.traceability import TracedValue, ValidationEntry, ValidationLog
from core.enums import Severity


class TestTraceability(unittest.TestCase):

    def test_traced_value(self):
        tv = TracedValue(
            value="Profesor Ejemplo",
            source_file="Control_Aulas_2026.xlsx",
            source_sheet="Hoja1",
            source_column="DOCENTE PRINCIPAL",
            source_col_idx=5,
            source_row=12,
        )

        self.assertEqual(tv.value, "Profesor Ejemplo")
        self.assertEqual(str(tv), "Profesor Ejemplo")
        self.assertTrue(bool(tv))
        self.assertIn("Control_Aulas_2026.xlsx", tv.origin_summary())
        self.assertIn("DOCENTE PRINCIPAL", tv.origin_summary())
        self.assertIn("fila 12", tv.origin_summary())

    def test_validation_log(self):
        log = ValidationLog()

        log.add_info("structure", "Archivo leído correctamente")
        log.add_warning("dates", "Fecha 2 de finalización difiere de Fecha 1")
        log.add_error("matching", "Estudiante no encontrado en sistematización")

        self.assertEqual(len(log.entries), 3)
        self.assertEqual(len(log.errors), 1)
        self.assertEqual(len(log.warnings), 1)
        self.assertEqual(len(log.infos), 1)
        self.assertTrue(log.has_errors)
        self.assertTrue(log.has_warnings)

        summary = log.summary()
        self.assertIn("1 errores", summary)
        self.assertIn("1 advertencias", summary)


if __name__ == "__main__":
    unittest.main()
