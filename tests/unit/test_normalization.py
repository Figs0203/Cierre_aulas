"""
Pruebas unitarias para los módulos de normalización (texto, fechas, encabezados).
"""

import unittest
from datetime import date

from normalization.text import (
    remove_accents,
    normalize_whitespace,
    normalize_text,
    normalize_for_comparison,
    normalize_document,
    normalize_email,
    normalize_name_component,
)
from normalization.dates import (
    parse_date,
    format_date_sistematizacion,
    format_date_certificados,
    extract_year,
    extract_semester,
)
from normalization.headers import analyze_headers


class TestNormalization(unittest.TestCase):

    def test_text_normalization(self):
        # Tildes
        self.assertEqual(remove_accents("Búsqueda y uso ético"), "Busqueda y uso etico")
        self.assertEqual(remove_accents("Año"), "Año")  # Preserva ñ

        # Espacios
        self.assertEqual(normalize_whitespace("  Texto   con   espacios  "), "Texto con espacios")

        # Normalización completa
        self.assertEqual(normalize_text("  CALIFICACIÓN MÓDULO 1  "), "calificacion modulo 1")

        # Normalización para comparación (sin puntuación)
        self.assertEqual(normalize_for_comparison("Aula Virtual / Solicitud:"), "aula virtual solicitud")

        # Documento y correo
        self.assertEqual(normalize_document("1.020.304.050"), "1020304050")
        self.assertEqual(normalize_document("10-203-040-50"), "1020304050")
        self.assertEqual(normalize_email(" Juan.Perez@EAFIT.EDU.CO "), "juan.perez@eafit.edu.co")

        # Nombre
        self.assertEqual(normalize_name_component("  MARÍA JOSÉ  "), "maria jose")

    def test_date_normalization(self):
        # Parseo de formatos variados
        d1 = parse_date("16/06/2025")
        self.assertEqual(d1, date(2025, 6, 16))

        d2 = parse_date("16-06-2025")
        self.assertEqual(d2, date(2025, 6, 16))

        d3 = parse_date("16-jun-2025")
        self.assertEqual(d3, date(2025, 6, 16))

        d4 = parse_date(date(2025, 6, 16))
        self.assertEqual(d4, date(2025, 6, 16))

        d_invalid = parse_date("fecha-invalida")
        self.assertIsNone(d_invalid)

        # Formatos de salida
        self.assertEqual(format_date_sistematizacion(d1), "16/06/2025")
        self.assertEqual(format_date_certificados(d1), "16-jun-2025")
        self.assertEqual(format_date_sistematizacion(None), "")
        self.assertEqual(format_date_certificados(None), "")

        # Año y semestre
        self.assertEqual(extract_year(d1), 2025)
        self.assertEqual(extract_semester(d1), 1)
        d_sem2 = date(2025, 10, 1)
        self.assertEqual(extract_semester(d_sem2), 2)

    def test_header_analysis(self):
        headers = [
            "OrgDefinedId",
            "Username",
            "Last Name",
            "First Name",
            "Sección",
            "Cuestionario módulo 1: estrategia y herramientas de búsqueda",
            "Buzón módulo 2: evaluación de la información",
            "Cuestionario módulo 3: organización de información",
            "Cuestionario módulo 4: uso ético de la información",
            "Calculated Final Grade",
            "Columna Desconocida Extra",
        ]

        analysis = analyze_headers(headers)

        # Validar mapeos conocidos
        self.assertIsNotNone(analysis.get_canonical("org_defined_id"))
        self.assertIsNotNone(analysis.get_canonical("username"))
        self.assertIsNotNone(analysis.get_canonical("apellidos"))
        self.assertIsNotNone(analysis.get_canonical("nombres"))
        self.assertIsNotNone(analysis.get_canonical("seccion"))
        self.assertIsNotNone(analysis.get_canonical("nota_final"))

        # Validar módulos detectados
        self.assertEqual(analysis.module_count, 4)
        self.assertIsNotNone(analysis.get_module(1))
        self.assertIsNotNone(analysis.get_module(2))
        self.assertIsNotNone(analysis.get_module(3))
        self.assertIsNotNone(analysis.get_module(4))

        # Validar columna no mapeada
        self.assertEqual(len(analysis.unmapped), 1)
        self.assertEqual(analysis.unmapped[0][0], "Columna Desconocida Extra")


if __name__ == "__main__":
    unittest.main()
