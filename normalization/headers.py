"""
Normalización y detección flexible de encabezados para el Asistente de Cierre de Aulas COIN.

Implementa un sistema de alias canónicos que permite reconocer columnas
independientemente de variaciones en:
- Mayúsculas/minúsculas.
- Tildes y diacríticos.
- Espacios simples, dobles o trailing.
- Signos de puntuación.
- Nombres alternativos del mismo concepto.

Cada columna mapeada se registra en el log de trazabilidad para auditoría:
    "Columna 'Calculated Final Grade' (col H) → interpretada como 'nota_final'"

Principio de diseño:
    Si un encabezado no puede mapearse con certeza → WARNING, no asumir.
    El sistema de alias es extensible y se refinará en la Fase 0 con archivos reales.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from normalization.text import normalize_for_comparison


@dataclass(frozen=True)
class HeaderMapping:
    """Resultado del mapeo de un encabezado original a su nombre canónico.

    Attributes:
        original: Nombre original tal como aparece en el archivo.
        canonical: Nombre canónico interno asignado.
        col_idx: Índice de columna (1-based).
        matched_via: Indica cómo se realizó el match (alias exacto, regex, etc.).
    """
    original: str
    canonical: str
    col_idx: int
    matched_via: str  # "alias", "module_regex", "final_grade_regex", "exact"

    def trace_description(self) -> str:
        """Descripción legible para el reporte de trazabilidad."""
        return (
            f"Columna '{self.original}' (col {self.col_idx}) "
            f"→ interpretada como '{self.canonical}' "
            f"(vía {self.matched_via})"
        )


@dataclass
class HeaderAnalysisResult:
    """Resultado completo del análisis de encabezados de un archivo.

    Attributes:
        mapped: Encabezados mapeados exitosamente.
        unmapped: Encabezados no reconocidos.
        modules_detected: Módulos detectados con su número y nombre original.
    """
    mapped: list[HeaderMapping] = field(default_factory=list)
    unmapped: list[tuple[str, int]] = field(default_factory=list)  # (nombre, col_idx)
    modules_detected: list[HeaderMapping] = field(default_factory=list)

    def get_canonical(self, canonical_name: str) -> HeaderMapping | None:
        """Busca un mapeo por nombre canónico."""
        for m in self.mapped:
            if m.canonical == canonical_name:
                return m
        return None

    def get_module(self, module_number: int) -> HeaderMapping | None:
        """Busca un módulo por su número."""
        target = f"modulo_{module_number}"
        for m in self.modules_detected:
            if m.canonical == target:
                return m
        return None

    @property
    def module_count(self) -> int:
        """Cantidad de módulos detectados."""
        return len(self.modules_detected)

    def summary(self) -> str:
        """Resumen legible del análisis de encabezados."""
        lines = [
            f"Columnas mapeadas: {len(self.mapped)}",
            f"Módulos detectados: {self.module_count}",
            f"Columnas no reconocidas: {len(self.unmapped)}",
        ]
        if self.unmapped:
            unmap_names = [f"'{name}' (col {idx})" for name, idx in self.unmapped]
            lines.append(f"  Sin mapear: {', '.join(unmap_names)}")
        return "\n".join(lines)


# ============================================================
# Regex para detección de módulos
# ============================================================

# Patrón principal: captura "módulo N" independientemente del prefijo
# Ejemplos que matchean:
#   "Cuestionario módulo 1: estrategia y herramientas"
#   "Buzón módulo 2: evaluación de la información"
#   "módulo 3"
#   "Modulo 4"
_MODULE_REGEX = re.compile(
    r"m[oó]dulo\s*(\d+)",
    re.IGNORECASE,
)

# Patrón para nota final
_FINAL_GRADE_PATTERNS = [
    re.compile(r"calculated\s*final\s*(grade)?", re.IGNORECASE),
    re.compile(r"nota\s*final", re.IGNORECASE),
    re.compile(r"promedio\s*final", re.IGNORECASE),
    re.compile(r"final\s*grade", re.IGNORECASE),
]


def analyze_headers(
    headers: list[str | None],
    alias_table: dict[str, str] | None = None,
) -> HeaderAnalysisResult:
    """Analiza y mapea los encabezados de un archivo Excel.

    Proceso:
    1. Para cada encabezado, normalizar el texto.
    2. Buscar en la tabla de alias.
    3. Intentar detectar patrón de módulo.
    4. Intentar detectar patrón de nota final.
    5. Si no hay match → registrar como no mapeado.

    Args:
        headers: Lista de encabezados tal como aparecen en el archivo
                 (puede contener None para celdas vacías).
        alias_table: Diccionario de aliases {texto_normalizado: nombre_canónico}.
                     Si es None, se usa la tabla por defecto de config/aliases.py.

    Returns:
        HeaderAnalysisResult con todos los mapeos y columnas no reconocidas.
    """
    if alias_table is None:
        from config.aliases import HEADER_ALIASES
        alias_table = HEADER_ALIASES

    result = HeaderAnalysisResult()

    for col_idx_0, raw_header in enumerate(headers):
        col_idx = col_idx_0 + 1  # 1-based como Excel

        if raw_header is None or (isinstance(raw_header, str) and not raw_header.strip()):
            continue

        original = str(raw_header).strip()
        normalized = normalize_for_comparison(original)

        # Intento 1: Búsqueda exacta en tabla de alias
        if normalized in alias_table:
            canonical = alias_table[normalized]

            # Verificar si es un módulo
            if canonical.startswith("modulo_"):
                mapping = HeaderMapping(
                    original=original,
                    canonical=canonical,
                    col_idx=col_idx,
                    matched_via="alias",
                )
                result.modules_detected.append(mapping)
                result.mapped.append(mapping)
            else:
                result.mapped.append(HeaderMapping(
                    original=original,
                    canonical=canonical,
                    col_idx=col_idx,
                    matched_via="alias",
                ))
            continue

        # Intento 2: Detección de módulo por regex
        module_match = _MODULE_REGEX.search(normalized)
        if module_match:
            module_num = int(module_match.group(1))
            canonical = f"modulo_{module_num}"
            mapping = HeaderMapping(
                original=original,
                canonical=canonical,
                col_idx=col_idx,
                matched_via="module_regex",
            )
            result.modules_detected.append(mapping)
            result.mapped.append(mapping)
            continue

        # Intento 3: Detección de nota final por regex
        is_final = False
        for pattern in _FINAL_GRADE_PATTERNS:
            if pattern.search(normalized):
                result.mapped.append(HeaderMapping(
                    original=original,
                    canonical="nota_final",
                    col_idx=col_idx,
                    matched_via="final_grade_regex",
                ))
                is_final = True
                break

        if is_final:
            continue

        # Sin match: registrar como no mapeado
        result.unmapped.append((original, col_idx))

    return result


def find_module_columns(headers: list[str | None]) -> dict[int, int]:
    """Retorna un diccionario {num_modulo: col_idx} mapeando módulos a su columna (1-based)."""
    analysis = analyze_headers(headers)
    mod_map = {}
    for m in analysis.modules_detected:
        try:
            num = int(m.canonical.split("_")[1])
            mod_map[num] = m.col_idx
        except (ValueError, IndexError):
            pass
    return mod_map


def find_final_grade_column(headers: list[str | None]) -> int | None:
    """Retorna el índice de columna (1-based) de la nota final o None."""
    analysis = analyze_headers(headers)
    mapping = analysis.get_canonical("nota_final")
    return mapping.col_idx if mapping else None
