"""
Motor de Matching Ultra-Conservador de Estudiantes.

Realiza la correspondencia unívoca entre estudiantes de Notas (Interactiva)
y estudiantes de Sistematización mediante claves determinísticas jerárquicas:
1. Documento de identidad normalizado (solo dígitos).
2. Correo electrónico institucional normalizado.
3. Username / OrgDefinedId.
4. Nombre + Apellido exactos (solo si es estrictamente unívoco).

Principio de seguridad:
  Aislamiento de errores: Un estudiante sin match o duplicado solo afecta a su
  propio registro, nunca detiene el procesamiento de la clase completa.
  Cero fuzzy matching: Prohibidas las coincidencias difusas o aproximadas.
"""

from __future__ import annotations

from typing import Dict, List, Set, Tuple

from core.models import (
    ClaseGroup,
    EstudianteNotas,
    EstudianteSistematizacion,
    StudentMatch,
)
from normalization.text import (
    normalize_document,
    normalize_email,
    normalize_name_component,
)


def match_students_in_class(
    clase_group: ClaseGroup,
) -> ClaseGroup:
    """Realiza el cruce conservador de estudiantes dentro de un ClaseGroup aislado.

    Modifica y retorna el ClaseGroup con sus listas de matches, unmatched_sist,
    unmatched_notas y excluded_orange pobladas.

    Args:
        clase_group: Contenedor ClaseGroup con estudiantes_sist y estudiantes_notas.

    Returns:
        El mismo ClaseGroup con los resultados de matching actualizados.
    """
    sist_pool = list(clase_group.estudiantes_sist)
    notas_pool = list(clase_group.estudiantes_notas)

    matches: List[StudentMatch] = []
    unmatched_sist: List[EstudianteSistematizacion] = []
    unmatched_notas: List[EstudianteNotas] = []
    excluded_orange: List[EstudianteNotas] = []

    # 1. Separar estudiantes de Notas con relleno naranja (excluidos/retirados)
    notas_validas: List[EstudianteNotas] = []
    for en in notas_pool:
        if en.is_orange_fill:
            excluded_orange.append(en)
        else:
            notas_validas.append(en)

    # 2. Construir índices sobre Sistematización para búsqueda determinística rápida
    # Si un documento o correo se repite en Sistematización, se marca como conflictivo
    doc_to_sist: Dict[str, List[EstudianteSistematizacion]] = {}
    email_to_sist: Dict[str, List[EstudianteSistematizacion]] = {}
    name_to_sist: Dict[str, List[EstudianteSistematizacion]] = {}

    for es in sist_pool:
        doc_raw = str(es.documento.value) if es.documento and es.documento.value is not None else ""
        doc_norm = normalize_document(doc_raw)
        if doc_norm:
            doc_to_sist.setdefault(doc_norm, []).append(es)

        mail_raw = str(es.correo.value) if es.correo and es.correo.value is not None else ""
        mail_norm = normalize_email(mail_raw)
        if mail_norm:
            email_to_sist.setdefault(mail_norm, []).append(es)

        nom_raw = str(es.nombres.value) if es.nombres and es.nombres.value is not None else ""
        ape_raw = str(es.apellidos.value) if es.apellidos and es.apellidos.value is not None else ""
        full_name_norm = f"{normalize_name_component(nom_raw)} {normalize_name_component(ape_raw)}".strip()
        if full_name_norm:
            name_to_sist.setdefault(full_name_norm, []).append(es)

    matched_sist_rows: Set[int] = set()
    matched_notas_rows: Set[int] = set()

    # 3. Ronda 1: Coincidencia exacta por Documento
    for en in notas_validas:
        if en.row_number in matched_notas_rows:
            continue

        doc_raw = str(en.org_defined_id.value) if en.org_defined_id and en.org_defined_id.value is not None else ""
        doc_norm = normalize_document(doc_raw)

        if doc_norm and doc_norm in doc_to_sist:
            candidates = doc_to_sist[doc_norm]
            if len(candidates) == 1:
                target_es = candidates[0]
                if target_es.row_number not in matched_sist_rows:
                    matches.append(StudentMatch(
                        estudiante_sist=target_es,
                        estudiante_nota=en,
                        match_key="documento",
                        match_value=doc_norm,
                    ))
                    matched_sist_rows.add(target_es.row_number)
                    matched_notas_rows.add(en.row_number)
            else:
                # Documento duplicado en sistematización: no asociar automáticamente para evitar colisión
                pass

    # 4. Ronda 2: Coincidencia exacta por Correo Institucional
    for en in notas_validas:
        if en.row_number in matched_notas_rows:
            continue

        username_raw = str(en.username.value) if en.username and en.username.value is not None else ""
        user_norm = normalize_email(username_raw)

        # Si el username no tiene @, probar construyendo el correo institucional
        email_candidates = [user_norm]
        if "@" not in user_norm and user_norm:
            email_candidates.append(f"{user_norm}@eafit.edu.co")

        found = False
        for em in email_candidates:
            if em in email_to_sist:
                cands = email_to_sist[em]
                if len(cands) == 1:
                    target_es = cands[0]
                    if target_es.row_number not in matched_sist_rows:
                        matches.append(StudentMatch(
                            estudiante_sist=target_es,
                            estudiante_nota=en,
                            match_key="correo",
                            match_value=em,
                        ))
                        matched_sist_rows.add(target_es.row_number)
                        matched_notas_rows.add(en.row_number)
                        found = True
                        break
        if found:
            continue

    # 5. Ronda 3: Coincidencia unívoca por Nombre + Apellido exactos
    for en in notas_validas:
        if en.row_number in matched_notas_rows:
            continue

        first_raw = str(en.first_name.value) if en.first_name and en.first_name.value is not None else ""
        last_raw = str(en.last_name.value) if en.last_name and en.last_name.value is not None else ""
        full_norm = f"{normalize_name_component(first_raw)} {normalize_name_component(last_raw)}".strip()

        if full_norm and full_norm in name_to_sist:
            cands = name_to_sist[full_norm]
            if len(cands) == 1:
                target_es = cands[0]
                if target_es.row_number not in matched_sist_rows:
                    matches.append(StudentMatch(
                        estudiante_sist=target_es,
                        estudiante_nota=en,
                        match_key="nombre_exacto",
                        match_value=full_norm,
                        warnings=["Coincidencia realizada únicamente por nombre y apellido."],
                    ))
                    matched_sist_rows.add(target_es.row_number)
                    matched_notas_rows.add(en.row_number)

    # 6. Identificar registros no asociados (unmatched)
    for en in notas_validas:
        if en.row_number not in matched_notas_rows:
            unmatched_notas.append(en)

    for es in sist_pool:
        if es.row_number not in matched_sist_rows:
            unmatched_sist.append(es)

    clase_group.matches = matches
    clase_group.unmatched_sist = unmatched_sist
    clase_group.unmatched_notas = unmatched_notas
    clase_group.excluded_orange = excluded_orange

    return clase_group
