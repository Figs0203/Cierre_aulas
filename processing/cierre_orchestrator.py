"""
Orquestador Integral de Cierre de Aulas COIN.

Coordina todo el flujo operativo de extremo a extremo:
1. Verificación de integridad inicial (SHA-256 de los 4 archivos).
2. Extracción mediante Parsers especializados (Control, Sistematización, Notas, Certificados).
3. Agrupación por clases y correspondencia determinística (Matching Engine).
4. Aplicación de la fórmula oficial de certificación y reglas de negocio.
5. Generación del libro Excel auxiliar de cierre (o reporte de simulación previa).
6. Verificación de inmutabilidad final post-procesamiento.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from core.models import AulaCierreResult, AulaMetadata, ClaseGroup
from matching.engine import match_students_in_class
from normalization.text import (
    normalize_document,
    normalize_email,
    normalize_name_component,
)
from output.excel_generator import generate_cierre_excel
from parsing.certificados_parser import parse_certificados_consecutivo
from parsing.control_parser import parse_control_aulas
from parsing.notas_parser import parse_notas_archivo
from parsing.sistematizacion_parser import parse_sistematizacion_aula
from processing.rules_engine import apply_rules_to_match
from security.integrity import compute_sha256, verify_integrity


def _nota_identity_matches_class(
    en,  # EstudianteNotas
    ests_sist: List,
) -> bool:
    """Determina si una nota de Interactiva pertenece a la clase dada comparando identidad.

    Se usa cuando el archivo de Notas NO trae columna 'Sección' (p. ej. cursos CMI),
    para asignar cada estudiante a la clase correspondiente por documento, correo o
    nombre en lugar de asignar el pool completo a todas las clases (lo que duplicaría
    excluidos y no-match).

    Returns:
        True si la nota corresponde a algún estudiante de Sistematización de la clase.
    """
    en_doc = normalize_document(en.org_defined_id.value if en.org_defined_id else "")
    en_mail = normalize_email(en.username.value if en.username else "")
    en_name = (
        f"{normalize_name_component(en.first_name.value if en.first_name else '')} "
        f"{normalize_name_component(en.last_name.value if en.last_name else '')}"
    ).strip()

    for es in ests_sist:
        es_doc = normalize_document(es.documento.value if es.documento else "")
        es_mail = normalize_email(es.correo.value if es.correo else "")
        es_name = (
            f"{normalize_name_component(es.nombres.value if es.nombres else '')} "
            f"{normalize_name_component(es.apellidos.value if es.apellidos else '')}"
        ).strip()

        if en_doc and es_doc and en_doc == es_doc:
            return True
        if en_mail and es_mail and (en_mail == es_mail or es_mail.startswith(en_mail + "@")):
            return True
        if en_name and es_name and en_name == es_name:
            return True

    return False


def run_cierre_pipeline(
    file_paths: Dict[str, str | Path],
    target_aula: str,
    codigo_curso: Optional[str] = None,
    is_simulation: bool = False,
    output_dir: Optional[str | Path] = None,
) -> Tuple[AulaCierreResult, Optional[Path], bool, List[str]]:
    """Ejecuta el pipeline completo de cierre de un aula virtual.

    Args:
        file_paths: Diccionario con rutas a 'control_aulas', 'sistematizacion', 'notas', 'certificados'.
        target_aula: Código del aula a cerrar (ej. 'ABBUEI205').
        codigo_curso: Código para certificados (si es None, usa target_aula).
        is_simulation: Si True, solo simula y valida sin generar archivo Excel.
        output_dir: Directorio de salida para el Excel auxiliar.

    Returns:
        Tupla (resultado, ruta_excel_generado, integridad_ok, mensajes_integridad)
    """
    clean_aula = target_aula.strip().upper()
    cod_curso = codigo_curso.strip() if codigo_curso else clean_aula

    # 1. Calcular hashes SHA-256 antes de abrir archivos
    hashes_before = {role: compute_sha256(p) for role, p in file_paths.items()}

    # 2. Parsear Sistematización
    estudiantes_sist_por_clase, sist_headers = parse_sistematizacion_aula(
        file_paths["sistematizacion"],
        target_aula=clean_aula,
    )

    if not estudiantes_sist_por_clase:
        raise ValueError(
            f"No se encontraron registros para el aula '{clean_aula}' en la hoja de Sistematización."
        )

    clases_detectadas = list(estudiantes_sist_por_clase.keys())

    # 3. Parsear Notas
    estudiantes_notas, secciones_map = parse_notas_archivo(file_paths["notas"])

    # 4. Parsear Control de Aulas
    metadatos_control = parse_control_aulas(
        file_paths["control_aulas"],
        target_aula=clean_aula,
        clases_objetivo=clases_detectadas,
    )

    # 5. Parsear Certificados (consecutivo)
    max_consecutivo, prev_cert_count = parse_certificados_consecutivo(
        file_paths["certificados"],
        target_aula=clean_aula,
    )

    # 6. Agrupar y procesar por cada Clase detectada
    clases_groups: List[ClaseGroup] = []

    # Determinar si el archivo de notas tiene una sola sección o múltiples

    # Pre-pasada: asignar cada nota a UNA sola clase (sección; si no, por identidad).
    # Esto evita duplicar el mismo estudiante (y su exclusión/no-match) en varias clases
    # cuando el archivo de Notas no trae columna 'Sección'.
    multi_clase = len(estudiantes_sist_por_clase) > 1
    notas_por_clase: Dict[str, list] = {}
    assigned_rows: set[int] = set()

    for clase_id, ests_sist in estudiantes_sist_por_clase.items():
        if not multi_clase:
            pool = list(estudiantes_notas)
        else:
            pool = [
                en for en in estudiantes_notas
                if en.row_number not in assigned_rows
                and en.seccion and clase_id in str(en.seccion.value)
            ]
            if not pool:
                pool = [
                    en for en in estudiantes_notas
                    if en.row_number not in assigned_rows
                    and _nota_identity_matches_class(en, ests_sist)
                ]
        notas_por_clase[clase_id] = pool
        assigned_rows.update(en.row_number for en in pool)

    # Notas que no corresponden a ninguna clase: se reportan una sola vez en la
    # primera clase (queda constancia en la hoja de validaciones).
    if multi_clase:
        leftovers = [en for en in estudiantes_notas if en.row_number not in assigned_rows]
        if leftovers:
            first_clase = next(iter(notas_por_clase))
            notas_por_clase[first_clase].extend(leftovers)

    for clase_id, ests_sist in estudiantes_sist_por_clase.items():
        cg = ClaseGroup(clase_id=clase_id)
        cg.estudiantes_sist = ests_sist

        # Asociar metadatos de Control si existen
        meta_clase = metadatos_control.get(clase_id) or AulaMetadata()
        cg.programa = meta_clase.programa_academico
        cg.catalogo = meta_clase.catalogo
        cg.docente = meta_clase.docente_principal
        cg.fecha_inicio = meta_clase.fecha_inicio
        cg.fecha_fin = meta_clase.fecha_fin

        # Notas ya asignadas a esta clase en la pre-pasada.
        cg.estudiantes_notas = list(notas_por_clase.get(clase_id, []))

        # 7. Ejecutar Matching de Estudiantes en esta clase
        cg = match_students_in_class(cg)

        # 8. Aplicar Reglas Oficiales de Certificación a cada match
        for m in cg.matches:
            apply_rules_to_match(m, codigo_curso=cod_curso, clase_id=clase_id, aula_meta=meta_clase)

        clases_groups.append(cg)

    # Construir objeto resultado
    primary_meta = list(metadatos_control.values())[0] if metadatos_control else AulaMetadata()
    start_consecutivo = max_consecutivo + 1 if max_consecutivo > 0 else 1
    cierre_result = AulaCierreResult(
        aula=primary_meta,
        clases=clases_groups,
        hashes_before={k: v.sha256_hash for k, v in hashes_before.items()},
        metadatos_control=metadatos_control,
        start_consecutivo=start_consecutivo,
    )

    # 9. Generar Excel Auxiliar si no es simulación
    output_file: Optional[Path] = None
    if not is_simulation:
        output_file = generate_cierre_excel(
            result=cierre_result,
            codigo_curso=cod_curso,
            sist_headers=sist_headers,
            start_consecutivo=start_consecutivo,
            output_dir=output_dir,
            metadatos_control=metadatos_control,
        )

    # 10. Verificación final de integridad de los 4 archivos fuente
    intact, integrity_msgs = verify_integrity(hashes_before, file_paths)
    cierre_result.hashes_after = {role: compute_sha256(p).sha256_hash for role, p in file_paths.items()}

    return cierre_result, output_file, intact, integrity_msgs
