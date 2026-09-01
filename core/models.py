"""
Modelos de datos del dominio para el Asistente de Cierre de Aulas COIN.

Define las estructuras de datos que representan la información extraída
de los cuatro archivos fuente. Cada campo relevante se envuelve en
TracedValue para mantener trazabilidad completa del origen.

IMPORTANTE:
    Estos modelos representan la ESTRUCTURA del dominio, no las reglas
    de negocio. Las reglas (aprobación, certificación, etc.) se definen
    en los módulos de processing/ y se confirmarán en la Fase 0.

    Los modelos están diseñados para tolerar información faltante:
    la mayoría de los campos son opcionales (None) y se dejan vacíos
    cuando no se pueden determinar con certeza.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from core.traceability import TracedValue


# ============================================================
# METADATOS DEL AULA (desde Control de Aulas)
# ============================================================

@dataclass
class AulaMetadata:
    """Información del aula extraída del archivo Control_Aulas_Formadores.

    Cada campo puede ser None si no se encontró o no se pudo determinar.
    Los campos TracedValue conservan la procedencia exacta.
    """
    codigo_aula: TracedValue | None = None
    curso: TracedValue | None = None
    programa_academico: TracedValue | None = None
    catalogo: TracedValue | None = None
    docente_principal: TracedValue | None = None
    fecha_inicio: TracedValue | None = None
    fecha_fin: TracedValue | None = None
    fecha_fin_2: TracedValue | None = None
    estado_certificados: TracedValue | None = None
    estado_envio_notas: TracedValue | None = None
    estado_cierre_curso: TracedValue | None = None
    color_envio_notas: str | None = None  # Hex del color de la celda
    color_cierre_curso: str | None = None  # Hex del color de la celda

    # Campos adicionales que se extraerán si existen en el archivo
    extra_fields: dict[str, TracedValue] = field(default_factory=dict)


# ============================================================
# INFORMACIÓN POR CLASE (unidad lógica de procesamiento)
# ============================================================

@dataclass
class ClaseGroup:
    """Contenedor de una clase individual dentro de un aula.

    Representa la unidad fundamental de procesamiento:
    Aula + Clase = contenedor estanco.

    Ningún dato debe compartirse entre ClaseGroups sin validación.
    """
    clase_id: str
    seccion_id: str | None = None  # Valor de "Sección" en Notas, si existe

    # Metadatos específicos de esta clase (pueden diferir de otra clase del aula)
    docente: TracedValue | None = None
    ciclo: TracedValue | None = None
    catalogo: TracedValue | None = None
    programa: TracedValue | None = None
    fecha_inicio: TracedValue | None = None
    fecha_fin: TracedValue | None = None

    # Listas de estudiantes por fuente
    estudiantes_sist: list[EstudianteSistematizacion] = field(default_factory=list)
    estudiantes_notas: list[EstudianteNotas] = field(default_factory=list)

    # Resultados del matching
    matches: list[StudentMatch] = field(default_factory=list)
    unmatched_sist: list[EstudianteSistematizacion] = field(default_factory=list)
    unmatched_notas: list[EstudianteNotas] = field(default_factory=list)
    excluded_orange: list[EstudianteNotas] = field(default_factory=list)


# ============================================================
# ESTUDIANTE EN SISTEMATIZACIÓN
# ============================================================

@dataclass
class EstudianteSistematizacion:
    """Registro de un estudiante tal como aparece en el archivo de Sistematización.

    Todos los campos académicos se almacenan como TracedValue para
    poder rastrear su origen exacto.
    """
    row_number: int  # Fila en el archivo original (1-based)

    # Datos personales (se enmascaran en logs, completos en el Excel auxiliar)
    nombres: TracedValue | None = None
    apellidos: TracedValue | None = None
    tipo_documento: TracedValue | None = None
    documento: TracedValue | None = None
    correo: TracedValue | None = None

    # Datos académicos
    clase: TracedValue | None = None
    catalogo: TracedValue | None = None
    programa: TracedValue | None = None
    formador_lider: TracedValue | None = None
    formador_acompanante: TracedValue | None = None

    # Calificaciones ya existentes en la sistematización
    calificaciones_existentes: dict[str, TracedValue] = field(default_factory=dict)
    promedio_existente: TracedValue | None = None

    # Estado de certificación (puede contener fórmula)
    estado_certificacion_raw: TracedValue | None = None
    formula_estado_certificacion: str | None = None  # Fórmula Excel si existe

    # Campos adicionales extraídos
    extra_fields: dict[str, TracedValue] = field(default_factory=dict)


# ============================================================
# ESTUDIANTE EN NOTAS
# ============================================================

@dataclass
class EstudianteNotas:
    """Registro de un estudiante tal como aparece en el archivo de Notas.

    Incluye información sobre el color de relleno para detección de
    estudiantes excluidos (naranja).
    """
    row_number: int  # Fila en el archivo original (1-based)

    # Identificadores
    org_defined_id: TracedValue | None = None
    username: TracedValue | None = None
    first_name: TracedValue | None = None
    last_name: TracedValue | None = None
    seccion: TracedValue | None = None

    # Calificaciones por módulo: {"modulo_1": TracedValue, "modulo_2": TracedValue, ...}
    calificaciones_modulos: dict[str, TracedValue] = field(default_factory=dict)

    # Columnas de notas adicionales no mapeadas a módulos conocidos
    modulos_adicionales: dict[str, TracedValue] = field(default_factory=dict)

    # Nota final
    nota_final: TracedValue | None = None

    # Información de color de relleno para detección de excluidos
    is_orange_fill: bool = False
    fill_details: dict[str, Any] = field(default_factory=dict)
    # fill_details puede contener:
    #   "hex": "FF8C00"
    #   "theme": 6
    #   "tint": 0.4
    #   "indexed": 52
    #   "fill_type": "solid"
    #   "conditional_format": True/False


# ============================================================
# RESULTADO DE MATCHING
# ============================================================

@dataclass
class StudentMatch:
    """Resultado de la correspondencia entre un estudiante de Notas y Sistematización.

    Solo se crea cuando existe una correspondencia determinística unívoca.
    """
    estudiante_sist: EstudianteSistematizacion
    estudiante_nota: EstudianteNotas
    match_key: str  # "documento", "correo", "id_institucional", "nombre_exacto"
    match_value: str  # El valor que produjo la coincidencia
    warnings: list[str] = field(default_factory=list)


# ============================================================
# RESULTADO COMPLETO DEL PROCESAMIENTO DE UN AULA
# ============================================================

@dataclass
class AulaCierreResult:
    """Resultado completo del procesamiento de cierre de un aula.

    Agrupa toda la información necesaria para generar el archivo
    auxiliar de salida.
    """
    # Metadatos del aula
    aula: AulaMetadata

    # Clases procesadas
    clases: list[ClaseGroup] = field(default_factory=list)

    # Hashes de integridad
    hashes_before: dict[str, str] = field(default_factory=dict)
    hashes_after: dict[str, str] = field(default_factory=dict)

    # Confirmación manual del Planner
    planner_confirmed: bool | None = None  # None = no preguntado

    # Fecha de generación del cierre
    fecha_generacion: str = ""

    @property
    def total_estudiantes(self) -> int:
        """Total de estudiantes procesados en todas las clases."""
        return sum(len(c.estudiantes_sist) for c in self.clases)

    @property
    def total_matched(self) -> int:
        """Total de matches exitosos."""
        return sum(len(c.matches) for c in self.clases)

    @property
    def total_excluded(self) -> int:
        """Total de estudiantes excluidos (naranja)."""
        return sum(len(c.excluded_orange) for c in self.clases)

    @property
    def total_unmatched_notas(self) -> int:
        """Total de estudiantes en notas sin correspondencia."""
        return sum(len(c.unmatched_notas) for c in self.clases)

    @property
    def total_unmatched_sist(self) -> int:
        """Total de estudiantes en sistematización sin notas."""
        return sum(len(c.unmatched_sist) for c in self.clases)
