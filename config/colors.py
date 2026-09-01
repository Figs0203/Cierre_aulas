"""
Definiciones de colores para el Asistente de Cierre de Aulas COIN.

Define rangos de color para la detección de estudiantes marcados
en naranja y de estados en verde en el archivo de Control de Aulas.

IMPORTANTE:
    Los rangos definidos aquí son provisionales y se DEBEN refinar
    en la Fase 0 tras inspeccionar los archivos reales.

    La detección de color nunca debe usarse para excluir estudiantes
    automáticamente. Siempre requiere confirmación interactiva.

    Excel puede representar colores como:
    - RGB directo (fgColor.rgb)
    - Theme color + tint (fgColor.theme, fgColor.tint)
    - Indexed color (fgColor.indexed)
    - Formato condicional (requiere inspección de reglas CF)

    La Fase 0 determinará cuál(es) de estos mecanismos se utilizan.
"""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class ColorRange:
    """Define un rango de color RGB para detección.

    Attributes:
        name: Nombre descriptivo del color.
        r_min: Valor mínimo del componente rojo (0-255).
        r_max: Valor máximo del componente rojo (0-255).
        g_min: Valor mínimo del componente verde (0-255).
        g_max: Valor máximo del componente verde (0-255).
        b_min: Valor mínimo del componente azul (0-255).
        b_max: Valor máximo del componente azul (0-255).
    """
    name: str
    r_min: int
    r_max: int
    g_min: int
    g_max: int
    b_min: int
    b_max: int

    def contains(self, r: int, g: int, b: int) -> bool:
        """Verifica si un color RGB cae dentro de este rango."""
        return (
            self.r_min <= r <= self.r_max
            and self.g_min <= g <= self.g_max
            and self.b_min <= b <= self.b_max
        )


# ============================================================
# RANGOS DE COLOR PROVISIONALES
# Estos valores se ajustarán en la Fase 0 con archivos reales.
# ============================================================

# Rango provisorio para "naranja" (estudiantes excluidos en archivo de notas)
ORANGE_RANGE = ColorRange(
    name="Naranja (estudiantes excluidos)",
    r_min=200, r_max=255,
    g_min=100, g_max=200,
    b_min=0, b_max=100,
)

# Rango provisorio para "verde" (estados completados en Control de Aulas)
GREEN_RANGE = ColorRange(
    name="Verde (estado completado)",
    r_min=0, r_max=150,
    g_min=150, g_max=255,
    b_min=0, b_max=150,
)


def parse_hex_color(hex_color: str | None) -> tuple[int, int, int] | None:
    """Parsea un color hexadecimal a componentes RGB.

    Maneja formatos:
    - "FF8C00" (6 caracteres)
    - "00FF8C00" (8 caracteres con canal alpha que se ignora)
    - "#FF8C00" (con prefijo #)

    Args:
        hex_color: String hexadecimal del color.

    Returns:
        Tupla (R, G, B) o None si no se puede parsear.
    """
    if not hex_color or not isinstance(hex_color, str):
        return None

    hex_color = hex_color.strip().lstrip("#")

    # openpyxl a veces retorna 8 caracteres (AARRGGBB)
    if len(hex_color) == 8:
        hex_color = hex_color[2:]  # Descartar canal alpha

    if len(hex_color) != 6:
        return None

    try:
        r = int(hex_color[0:2], 16)
        g = int(hex_color[2:4], 16)
        b = int(hex_color[4:6], 16)
        return (r, g, b)
    except ValueError:
        return None
