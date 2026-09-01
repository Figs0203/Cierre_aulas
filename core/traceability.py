"""
Sistema de trazabilidad para el Asistente de Cierre de Aulas COIN.

Proporciona la clase TracedValue que encapsula cualquier dato extraído
conservando su procedencia exacta (archivo, hoja, columna, fila).

Esto permite auditar el origen de cada celda generada en el archivo
auxiliar de salida y responder siempre:
    ¿De dónde salió este dato?
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass(frozen=True)
class TracedValue:
    """Encapsula un valor junto con la información completa de su origen.

    Attributes:
        value: El dato extraído (puede ser str, int, float, date, None, etc.).
        source_file: Nombre del archivo de origen (sin ruta completa por privacidad).
        source_sheet: Nombre de la hoja de origen dentro del archivo.
        source_column: Nombre del encabezado original de la columna.
        source_col_idx: Índice de la columna (1-based, como en Excel).
        source_row: Número de fila exacto (1-based, como en Excel).
    """
    value: Any
    source_file: str
    source_sheet: str
    source_column: str
    source_col_idx: int
    source_row: int

    def __repr__(self) -> str:
        return (
            f"TracedValue({self.value!r}, "
            f"from='{self.source_file}' / '{self.source_sheet}' / "
            f"col '{self.source_column}' [idx={self.source_col_idx}] / "
            f"row {self.source_row})"
        )

    def __str__(self) -> str:
        """Devuelve la representación del valor subyacente."""
        return str(self.value) if self.value is not None else ""

    def __bool__(self) -> bool:
        """Permite verificar si el valor subyacente es truthy."""
        return self.value is not None and self.value != ""

    def origin_summary(self) -> str:
        """Resumen legible del origen para la hoja de validaciones.

        Returns:
            String como: "Control_Aulas_2026.xlsx / Hoja1 / DOCENTE PRINCIPAL (col 5) / fila 45"
        """
        return (
            f"{self.source_file} / {self.source_sheet} / "
            f"{self.source_column} (col {self.source_col_idx}) / "
            f"fila {self.source_row}"
        )


@dataclass
class ValidationEntry:
    """Una entrada individual en el registro de validaciones.

    Attributes:
        severity: Nivel de severidad (ERROR, WARNING, INFO).
        category: Categoría funcional (matching, grades, dates, structure, etc.).
        message: Descripción legible del hallazgo o problema.
        field: Campo afectado (opcional).
        value: Valor relevante (opcional).
        source_file: Archivo donde se detectó el problema (opcional).
        source_sheet: Hoja donde se detectó el problema (opcional).
        source_row: Fila donde se detectó el problema (opcional).
        source_column: Columna donde se detectó el problema (opcional).
    """
    severity: Any  # Severity enum, usamos Any para evitar import circular
    category: str
    message: str
    field: str | None = None
    value: Any = None
    source_file: str | None = None
    source_sheet: str | None = None
    source_row: int | None = None
    source_column: str | None = None

    def format_for_display(self) -> str:
        """Formato legible para consola y reportes."""
        icon = {
            "ERROR": "🔴",
            "WARNING": "🟡",
            "INFO": "🟢",
        }.get(self.severity.name if hasattr(self.severity, 'name') else str(self.severity), "⚪")

        parts = [f"{icon} [{self.severity.name if hasattr(self.severity, 'name') else self.severity}]"]
        parts.append(f"[{self.category}]")
        parts.append(self.message)

        if self.source_file:
            location = self.source_file
            if self.source_sheet:
                location += f" / {self.source_sheet}"
            if self.source_row:
                location += f" / fila {self.source_row}"
            if self.source_column:
                location += f" / col '{self.source_column}'"
            parts.append(f"(Fuente: {location})")

        return " ".join(parts)


@dataclass
class ValidationLog:
    """Registro acumulativo de todas las validaciones del proceso.

    Proporciona métodos para agregar entradas y consultar por severidad.
    """
    entries: list[ValidationEntry] = field(default_factory=list)

    def add(self, entry: ValidationEntry) -> None:
        """Agrega una entrada de validación al registro."""
        self.entries.append(entry)

    def add_error(self, category: str, message: str, **kwargs) -> None:
        """Atajo para agregar un ERROR."""
        from core.enums import Severity
        self.entries.append(ValidationEntry(
            severity=Severity.ERROR, category=category, message=message, **kwargs
        ))

    def add_warning(self, category: str, message: str, **kwargs) -> None:
        """Atajo para agregar un WARNING."""
        from core.enums import Severity
        self.entries.append(ValidationEntry(
            severity=Severity.WARNING, category=category, message=message, **kwargs
        ))

    def add_info(self, category: str, message: str, **kwargs) -> None:
        """Atajo para agregar un INFO."""
        from core.enums import Severity
        self.entries.append(ValidationEntry(
            severity=Severity.INFO, category=category, message=message, **kwargs
        ))

    @property
    def errors(self) -> list[ValidationEntry]:
        """Todas las entradas de tipo ERROR."""
        from core.enums import Severity
        return [e for e in self.entries if e.severity == Severity.ERROR]

    @property
    def warnings(self) -> list[ValidationEntry]:
        """Todas las entradas de tipo WARNING."""
        from core.enums import Severity
        return [e for e in self.entries if e.severity == Severity.WARNING]

    @property
    def infos(self) -> list[ValidationEntry]:
        """Todas las entradas de tipo INFO."""
        from core.enums import Severity
        return [e for e in self.entries if e.severity == Severity.INFO]

    @property
    def has_errors(self) -> bool:
        """True si existe al menos un ERROR."""
        return len(self.errors) > 0

    @property
    def has_warnings(self) -> bool:
        """True si existe al menos un WARNING."""
        return len(self.warnings) > 0

    def summary(self) -> str:
        """Resumen numérico de las validaciones."""
        return (
            f"Validaciones: {len(self.errors)} errores, "
            f"{len(self.warnings)} advertencias, "
            f"{len(self.infos)} informativas"
        )
