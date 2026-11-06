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

# Patrón principal: captura el número de módulo aunque el encabezado NO diga
# literalmente "módulo". En los archivos de Interactiva cada módulo puede venir
# como "Cuestionario 1", "Buzón 2", "Foro 3", etc.
# Ejemplos que matchean:
#   "Cuestionario módulo 1: estrategia y herramientas"  → 1
#   "Buzón módulo 2: evaluación de la información"       → 2
#   "Cuestionario 1: estrategia y herramientas"         → 1  (sin la palabra módulo)
#   "módulo 3" / "Modulo 4" / "N.º 5"                     → 3 / 4 / 5
_MODULE_REGEX = re.compile(
    r"(?:"
    r"m[oó]dulo|cuestionario|buz[oó]n|foro|taller|unidad|actividad|"
    r"evaluaci[oó]n|tarea|quiz|examen|parcial|entrega|lecci[oó]n"
    r")"
    r"\s*(?:n[.\u00ba\u00b0]?\s*)?(\d+)",
    re.IGNORECASE,
)

# Encabezados que NO deben confundirse con la nota final: fechas, horas,
# identidad, datos estructurales y columnas de actividades/módulos.
_FINAL_GRADE_NON_GRADE = re.compile(
    r"\bfecha\b|\bhora\b|nombre|apellido|correo|documento|identificaci|usuario|"
    r"username|seccion|programa|catalogo|clase|aula|semestre|periodo|\bano\b|\bmes\b|"
    r"porcentaje|peso|intento|modulo|cuestionario|foro|buzon|taller|examen|"
    r"actividad|tarea|proyecto|trabajo|unidad|leccion|entrega|semana|"
    r"sede|jornada|grupo|credito",
    re.IGNORECASE,
)

# Detección robusta de la nota final. Se evalúa sobre el encabezado normalizado
# (minúsculas, sin tildes ni signos). El puntaje resuelve empates cuando el
# archivo trae varias columnas parecidas. OJO: 'Calculated Final Grade Scheme
# Symbol' es, en muchos archivos, el nombre real de la columna que contiene la
# nota final; se trata igual que las demás y NUNCA se descarta por su nombre.
_FINAL_GRADE_CANDIDATES: list[tuple[str, int]] = [
    (r"calculated\s*final\s*grade", 100),          # Calculated Final Grade
    (r"calculated\s*final", 96),                   # Calculated Final
    (r"nota\s*(final|definitiva)", 98),            # Nota final / Nota definitiva
    (r"calificacion\s*(final|definitiva)", 98),    # Calificación final
    (r"\bdefinitiva\b", 95),                       # Definitiva
    (r"final\s*grade", 96),                        # Final Grade
    (r"final\s*score", 92),                        # Final Score
    (r"puntaje\s*final", 92),                      # Puntaje final
    (r"promedio\s*(final|curso\s*completo|general)", 90),  # Promedio curso completo
    (r"nota\s*(del\s*)?curso", 84),                # Nota del curso
    (r"calificacion\s*(del\s*)?curso", 84),        # Calificación del curso
    (r"^final$", 70),                              # Encabezado exacto "Final"
]


def _detectar_nota_final(header: str | None) -> tuple[int, str] | None:
    """Devuelve (prioridad, canónico) si el encabezado es una nota final.

    Reconoce variantes como 'Calculated Final Grade', 'Calculated Final',
    'Final', 'Final Grade', 'Nota final', 'Definitiva', 'Promedio curso
    completo' o 'Calculated Final Grade Scheme Symbol'. Esta última se trata
    como una columna de nota final más (en muchos archivos es la columna real de
    notas). Descarta encabezados de fechas, horas, identidad y actividades
    (p. ej. 'Fecha Final', 'Examen final').

    Returns:
        ``(puntaje, canonical)`` donde ``canonical`` es ``nota_final`` o
        ``promedio_curso``; ``None`` si no aplica.
    """
    normalized = normalize_for_comparison(header)
    if not normalized:
        return None
    if _FINAL_GRADE_NON_GRADE.search(normalized):
        return None

    puntaje = None
    for patron, valor in _FINAL_GRADE_CANDIDATES:
        if re.search(patron, normalized):
            puntaje = max(puntaje or 0, valor)

    if puntaje is None:
        return None

    if "promedio curso completo" in normalized or normalized.startswith("promedio general"):
        canonical = "promedio_curso"
    else:
        canonical = "nota_final"
    return puntaje, canonical


def final_grade_candidates(
    headers: list[str | None],
) -> list[tuple[int, int, str]]:
    """Todas las columnas candidatas a nota final: [(col_idx, puntaje, canónico)].

    Se exponen todas para que el parser pueda elegir, además del nombre, según
    los **datos reales** de cada columna (cuál contiene las notas).
    """
    candidatos: list[tuple[int, int, str]] = []
    for idx, header in enumerate(headers, start=1):
        detectada = _detectar_nota_final(header)
        if detectada is not None:
            candidatos.append((idx, detectada[0], detectada[1]))
    return candidatos


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

        # Intento 3: Detección robusta de nota final
        detectada = _detectar_nota_final(original)
        if detectada is not None:
            _puntaje, canonical = detectada
            result.mapped.append(HeaderMapping(
                original=original,
                canonical=canonical,
                col_idx=col_idx,
                matched_via="final_grade_regex",
            ))
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
    """Retorna el índice de columna (1-based) de la nota final o None.

    Reconoce variantes como 'Calculated Final Grade', 'Calculated Final',
    'Final', 'Final Grade', 'Nota final', 'Definitiva', 'Promedio curso completo'
    o 'Calculated Final Grade Scheme Symbol'. Todas se consideran columnas de
    nota final; **no** se descarta la de 'Scheme Symbol' por su nombre.

    Cuando hay varias columnas candidatas, esta función (solo con encabezados)
    toma la de mayor puntaje y, en empate, la primera. El parser de Notas
    refina la elección con los datos reales de cada columna.
    """
    candidatos = final_grade_candidates(headers)
    if candidatos:
        candidatos.sort(key=lambda c: (-c[1], c[0]))
        return candidatos[0][0]

    # Respaldo: alias canónicos ya mapeados (por si el alias no está cubierto
    # por el detector anterior).
    analysis = analyze_headers(headers)
    for canonical in ("nota_final", "promedio_curso"):
        mapping = analysis.get_canonical(canonical)
        if mapping:
            return mapping.col_idx
    return None
