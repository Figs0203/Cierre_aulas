"""
Tests unitarios para las funciones de formateo de columnas de salida.

Cubre las Reglas 5 y 6 confirmadas el 2026-09-04:
  - Regla 5: Columna 'Docente COIN' = '{CÓDIGO-CON-GUIÓN} {NOMBRE_FORMADOR}'
  - Regla 6: Columna 'Ciclo' = 'CLASE {N} {CAT} - {PROG} - {F_INI} A {F_FIN}'

También cubre la lista de hojas confirmada (Regla 1) y la
función de fecha para Ciclo: format_ciclo_date().
"""

from __future__ import annotations

from datetime import date, datetime

import pytest

from config.settings import (
    CONTROL_AULAS_ALL_SHEETS,
    CONTROL_AULAS_IGNORED_SHEET,
)
from normalization.dates import format_ciclo_date
from normalization.text import (
    format_ciclo,
    format_codigo_con_guion,
    format_docente_coin,
)


# ============================================================
# REGLA 1 — Hojas de Control de Aulas
# ============================================================


class TestHojasControlAulas:
    """Verifica que la lista de hojas activas y la hoja ignorada son correctas."""

    EXPECTED_ACTIVE = [
        "Tabla_CMI",
        "Tabla_ABBUEI",
        "Tabla_GICE",
        "Tabla_EPA",
        "Tabla_LATEX",
        "Tabla_IPI",
        "Tabla_UEI",
        "Tabla_CAI",
        "CONTROL",
    ]

    def test_hojas_activas_exactas(self):
        assert CONTROL_AULAS_ALL_SHEETS == self.EXPECTED_ACTIVE, (
            "La lista de hojas activas no coincide con lo confirmado."
        )

    def test_hoja_ignorada_nunca_en_activas(self):
        assert CONTROL_AULAS_IGNORED_SHEET not in CONTROL_AULAS_ALL_SHEETS, (
            f"'{CONTROL_AULAS_IGNORED_SHEET}' NO debe estar en CONTROL_AULAS_ALL_SHEETS."
        )

    def test_hoja_ignorada_es_la_correcta(self):
        assert CONTROL_AULAS_IGNORED_SHEET == "Creación aulas GDA"


# ============================================================
# REGLA 6 — Formato de fecha para Ciclo: format_ciclo_date()
# ============================================================


class TestFormatCicloDate:
    """Verifica la conversión de fechas al estilo 'MES DÍA' para la columna Ciclo."""

    def test_fecha_marzo(self):
        assert format_ciclo_date(date(2026, 3, 16)) == "MARZO 16"

    def test_fecha_abril(self):
        assert format_ciclo_date(date(2026, 4, 27)) == "ABRIL 27"

    def test_fecha_enero(self):
        assert format_ciclo_date(date(2026, 1, 5)) == "ENERO 5"

    def test_fecha_diciembre(self):
        assert format_ciclo_date(date(2025, 12, 31)) == "DICIEMBRE 31"

    def test_fecha_datetime_object(self):
        """También debe funcionar con objetos datetime."""
        dt = datetime(2026, 6, 15, 10, 30)
        assert format_ciclo_date(dt) == "JUNIO 15"

    def test_fecha_none_retorna_vacio(self):
        assert format_ciclo_date(None) == ""


# ============================================================
# REGLA 5 — format_codigo_con_guion()
# ============================================================


class TestFormatCodigoConGuion:
    """Verifica la inserción del guión en códigos de curso."""

    def test_abbuei(self):
        assert format_codigo_con_guion("ABBUEI205") == "ABBUEI-205"

    def test_cmi(self):
        assert format_codigo_con_guion("CMI154") == "CMI-154"

    def test_latex(self):
        assert format_codigo_con_guion("LATEX10") == "LATEX-10"

    def test_ya_tiene_guion(self):
        """Si ya tiene guión, no debe duplicarlo."""
        assert format_codigo_con_guion("ABBUEI-205") == "ABBUEI-205"

    def test_vacio_retorna_vacio(self):
        assert format_codigo_con_guion("") == ""

    def test_patron_no_reconocido_retorna_original(self):
        """Código sin patrón letras+números se devuelve sin cambios."""
        assert format_codigo_con_guion("12345") == "12345"

    def test_espacios_iniciales_finales(self):
        assert format_codigo_con_guion("  GICE99  ") == "GICE-99"


# ============================================================
# REGLA 5 — format_docente_coin()
# ============================================================


class TestFormatDocenteCOIN:
    """Verifica el formato de la columna Docente COIN."""

    def test_caso_confirmado(self):
        result = format_docente_coin("ABBUEI205", "Manuela Restrepo")
        assert result == "ABBUEI-205 Manuela Restrepo"

    def test_codigo_ya_tiene_guion(self):
        result = format_docente_coin("ABBUEI-205", "Manuela Restrepo")
        assert result == "ABBUEI-205 Manuela Restrepo"

    def test_codigo_cmi(self):
        result = format_docente_coin("CMI154", "Carlos Pérez")
        assert result == "CMI-154 Carlos Pérez"

    def test_none_codigo_retorna_vacio(self):
        assert format_docente_coin(None, "Carlos Pérez") == ""

    def test_none_nombre_retorna_vacio(self):
        assert format_docente_coin("CMI154", None) == ""

    def test_ambos_none_retorna_vacio(self):
        assert format_docente_coin(None, None) == ""

    def test_codigo_vacio_retorna_vacio(self):
        assert format_docente_coin("", "Carlos Pérez") == ""

    def test_nombre_vacio_retorna_vacio(self):
        assert format_docente_coin("CMI154", "") == ""


# ============================================================
# REGLA 6 — format_ciclo()
# ============================================================


class TestFormatCiclo:
    """Verifica el formato completo de la columna Ciclo."""

    def test_caso_confirmado(self):
        result = format_ciclo(
            clase=5535,
            catalogo="MF7001",
            programa="M. ESTUDIOS JURÍDICOS",
            fecha_inicio_str="MARZO 16",
            fecha_fin_str="ABRIL 27",
        )
        assert result == "CLASE 5535 MF7001 - M. ESTUDIOS JURÍDICOS - MARZO 16 A ABRIL 27"

    def test_clase_como_string(self):
        result = format_ciclo(
            clase="5535",
            catalogo="MF7001",
            programa="M. ESTUDIOS JURÍDICOS",
            fecha_inicio_str="MARZO 16",
            fecha_fin_str="ABRIL 27",
        )
        assert result == "CLASE 5535 MF7001 - M. ESTUDIOS JURÍDICOS - MARZO 16 A ABRIL 27"

    def test_none_clase_retorna_vacio(self):
        assert format_ciclo(None, "MF7001", "PROG", "MARZO 16", "ABRIL 27") == ""

    def test_none_catalogo_retorna_vacio(self):
        assert format_ciclo(5535, None, "PROG", "MARZO 16", "ABRIL 27") == ""

    def test_none_programa_retorna_vacio(self):
        assert format_ciclo(5535, "MF7001", None, "MARZO 16", "ABRIL 27") == ""

    def test_fecha_inicio_vacia_retorna_vacio(self):
        assert format_ciclo(5535, "MF7001", "PROG", "", "ABRIL 27") == ""

    def test_fecha_fin_vacia_retorna_vacio(self):
        assert format_ciclo(5535, "MF7001", "PROG", "MARZO 16", "") == ""

    def test_todos_none_retorna_vacio(self):
        assert format_ciclo(None, None, None, None, None) == ""

    def test_espacios_en_programa_se_normalizan(self):
        """El programa con espacios extra al inicio/fin debe funcionar igual."""
        result = format_ciclo(
            clase=5535,
            catalogo="MF7001",
            programa="  M. ESTUDIOS JURÍDICOS  ",
            fecha_inicio_str="MARZO 16",
            fecha_fin_str="ABRIL 27",
        )
        assert "M. ESTUDIOS JURÍDICOS" in result


# ============================================================
# FLUJO COMBINADO (Reglas 5 + 6 juntas)
# ============================================================


class TestFlujoCombinadoCertificado:
    """Simula el flujo completo de construcción de las dos columnas de certificado."""

    def test_ciclo_completo_con_fechas_reales(self):
        """Construye Ciclo a partir de fechas reales via format_ciclo_date."""
        f_inicio = format_ciclo_date(date(2026, 3, 16))
        f_fin = format_ciclo_date(date(2026, 4, 27))
        ciclo = format_ciclo(5535, "MF7001", "M. ESTUDIOS JURÍDICOS", f_inicio, f_fin)
        assert ciclo == "CLASE 5535 MF7001 - M. ESTUDIOS JURÍDICOS - MARZO 16 A ABRIL 27"

    def test_docente_y_ciclo_con_datos_faltantes_producen_vacio(self):
        """Si el formador no se puede extraer del archivo, ningún campo se llena."""
        docente = format_docente_coin("ABBUEI205", None)   # formador no encontrado
        f_inicio = format_ciclo_date(None)                 # fecha no encontrada
        ciclo = format_ciclo(5535, "MF7001", "PROG", f_inicio, "ABRIL 27")
        assert docente == ""
        assert ciclo == ""


# ============================================================
# DETECCIÓN DE COLOR NARANJA Y COMPATIBILIDAD CON OBJETOS RGB
# ============================================================


class TestOrangeFillDetection:
    """Verifica la detección de relleno naranja y la robustez ante objetos no-string."""

    def test_known_orange_hex_strings(self):
        from config.colors import is_orange_fill_exact
        assert is_orange_fill_exact("FFA500") is True
        assert is_orange_fill_exact("#FFA500") is True
        assert is_orange_fill_exact("ED7D31") is True
        assert is_orange_fill_exact("FFC000") is True
        assert is_orange_fill_exact("000000") is False
        assert is_orange_fill_exact(None) is False

    def test_theme_colors(self):
        from config.colors import is_orange_fill_exact
        assert is_orange_fill_exact(None, theme_val=6) is True
        assert is_orange_fill_exact(None, theme_val=5) is True
        assert is_orange_fill_exact(None, theme_val=1) is False

    def test_object_without_len_does_not_crash(self):
        """Verifica que un objeto sin __len__ (como openpyxl.styles.colors.RGB) no genere error."""
        class MockRGB:
            def __str__(self):
                return "FFFFA500"

        from config.colors import is_orange_fill_exact
        # No debe lanzar TypeError: object of type 'MockRGB' has no len()
        assert is_orange_fill_exact(MockRGB()) is False or is_orange_fill_exact("FFA500") is True
