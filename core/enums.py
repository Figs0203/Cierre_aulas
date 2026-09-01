"""
Enumeraciones del dominio para el Asistente de Cierre de Aulas COIN.

Define los niveles de severidad para validaciones y los estados
de procesamiento utilizados en todo el sistema.
"""

from enum import Enum, auto


class Severity(Enum):
    """Nivel de severidad de una entrada de validación.

    - ERROR: Impide la asignación segura de datos para el registro o bloque afectado.
    - WARNING: Requiere revisión humana obligatoria antes de aceptar el resultado.
    - INFO: Registro de auditoría y trazabilidad; no requiere acción.
    """
    ERROR = auto()
    WARNING = auto()
    INFO = auto()


class MatchKey(Enum):
    """Tipo de clave utilizada para establecer correspondencia entre estudiantes.

    Define la jerarquía de matching ultra-conservador:
    1. DOCUMENTO — Coincidencia exacta por número de documento (prioridad máxima).
    2. CORREO — Coincidencia exacta por correo electrónico.
    3. ID_INSTITUCIONAL — Username/OrgDefinedId solo si la Fase 0 confirma
       correspondencia determinística.
    4. NOMBRE_EXACTO — Nombre + apellido exactos y únicos (respaldo seguro).
    5. NO_MATCH — Sin correspondencia confiable.
    """
    DOCUMENTO = "documento"
    CORREO = "correo"
    ID_INSTITUCIONAL = "id_institucional"
    NOMBRE_EXACTO = "nombre_exacto"
    NO_MATCH = "no_match"


class StudentStatus(Enum):
    """Estado de procesamiento de un estudiante individual.

    - MATCHED: Correspondencia unívoca confirmada entre Notas y Sistematización.
    - UNMATCHED_NOTAS: Presente en Notas pero no encontrado en Sistematización.
    - UNMATCHED_SIST: Presente en Sistematización pero no encontrado en Notas.
    - EXCLUDED_ORANGE: Marcado con color naranja y confirmado como excluido.
    - DUPLICATE: Documento o correo duplicado detectado.
    - AMBIGUOUS: Múltiples candidatos posibles; requiere revisión manual.
    - PENDING_REVIEW: Pendiente de revisión humana por cualquier motivo.
    """
    MATCHED = "matched"
    UNMATCHED_NOTAS = "unmatched_notas"
    UNMATCHED_SIST = "unmatched_sist"
    EXCLUDED_ORANGE = "excluded_orange"
    DUPLICATE = "duplicate"
    AMBIGUOUS = "ambiguous"
    PENDING_REVIEW = "pending_review"


class AulaReadiness(Enum):
    """Estado de preparación del aula para el cierre.

    - READY: Cumple todos los requisitos verificables automáticamente.
    - ALREADY_CLOSED: Ya tiene fecha en la columna CERTIFICADOS.
    - MISSING_PREREQUISITES: Faltan fecha o estado en envío de notas o cierre.
    - PLANNER_UNCONFIRMED: Pendiente de confirmación manual de Planner/Teams.
    - NOT_FOUND: El aula no existe en el archivo de Control.
    """
    READY = "ready"
    ALREADY_CLOSED = "already_closed"
    MISSING_PREREQUISITES = "missing_prerequisites"
    PLANNER_UNCONFIRMED = "planner_unconfirmed"
    NOT_FOUND = "not_found"


class FileRole(Enum):
    """Rol asignado a cada archivo de entrada en el sistema.

    Utilizado para validar que un archivo no se seleccione para más de un rol
    y que su estructura sea compatible con el rol esperado.
    """
    CONTROL_AULAS = "control_aulas"
    SISTEMATIZACION = "sistematizacion"
    NOTAS = "notas"
    CERTIFICADOS = "certificados"
