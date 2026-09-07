"""
Motor de Reglas de Negocio Oficiales y Cálculo de Certificación.

Implementa la fórmula exacta confirmada en Fase 0 desde Sistematización:
  =IF(OR(AVERAGE(AS:AV)=0, AND(AS=0,AT=0,AU=0,AV=0), AND(AU=0,AV=0)), "Abandonó",
      IF(D="Pregrado",
          IF(OR(AS<3, AT<3, AU<3, AV<3), "No aprobó", "Aprobó"),
          IF(D="Posgrado",
              IF(OR(AS<3.5, AT<3.5, AU<3.5, AV<3.5), "No aprobó", "Aprobó")
          )
      )
  )

Además gestiona los formatos de:
- Se elaboró certificado SI/NO: 'SI' para Aprobó, 'NO' en otro caso.
- Código del certificado: Código del curso para SI, 'No aplica' para NO.
- Docente COIN: '{CÓDIGO-GUIÓN} {FORMADOR_LÍDER}'.
- Ciclo: 'CLASE {clase} {catalogo} - {programa} - {fecha_ini} A {fecha_fin}'.
"""

from __future__ import annotations

from datetime import date, datetime
from typing import Any, Dict, List, Optional, Tuple

from config.settings import (
    CERTIFICADO_NO,
    CERTIFICADO_SI,
    VALOR_NO_APLICA,
)
from core.models import AulaMetadata, StudentMatch
from normalization.dates import (
    format_certificate_date,
    format_ciclo_date,
    get_semester,
)
from normalization.text import (
    format_ciclo,
    format_docente_coin,
)


def evaluate_certification_status(
    calificaciones: Dict[str, float],
    tipo_grupo: Optional[str] = None,
    nota_final_calculated: Optional[float] = None,
) -> Tuple[str, float]:
    """Evalúa el estado de certificación oficial de un estudiante.

    Args:
        calificaciones: Diccionario {'modulo_1': nota, 'modulo_2': nota, ...}
        tipo_grupo: 'Pregrado', 'Posgrado' o None (default Pregrado).
        nota_final_calculated: Valor de 'Calculated Final' extraído directamente de Notas.

    Returns:
        Tupla (estado, promedio):
          - estado: 'APROBÓ', 'NO APROBÓ' o 'ABANDONÓ' (en mayúsculas oficiales).
          - promedio: Nota de 'Calculated Final' (o promedio de módulos como fallback).
    """
    # Extraer valores numéricos de módulos
    mod_keys = sorted([k for k in calificaciones if k.startswith("modulo_")])
    notas_vals: List[float] = []
    for k in mod_keys:
        v = calificaciones[k]
        try:
            notas_vals.append(float(v))
        except (ValueError, TypeError):
            pass

    # Asignar promedio: prioridad a Calculated Final de notas
    if nota_final_calculated is not None:
        promedio = nota_final_calculated
    elif notas_vals:
        promedio = round(sum(notas_vals) / len(notas_vals), 2)
    else:
        promedio = 0.0

    if not notas_vals and nota_final_calculated is None:
        return "Abandonó", 0.0

    # 1. Regla oficial de "Abandonó":
    # Promedio = 0, ó todas las notas en 0, ó los módulos 3 y 4 en 0
    all_zero = all(n == 0.0 for n in notas_vals) if notas_vals else (promedio == 0.0)
    m3_zero = False
    m4_zero = False
    if len(notas_vals) >= 3 and notas_vals[2] == 0.0:
        m3_zero = True
    if len(notas_vals) >= 4 and notas_vals[3] == 0.0:
        m4_zero = True

    if promedio == 0.0 or all_zero or (m3_zero and m4_zero):
        return "Abandonó", promedio

    # 2. Umbral según tipo de grupo (Pregrado >= 3.0, Posgrado >= 3.5)
    tipo_norm = str(tipo_grupo).strip().lower() if tipo_grupo else "pregrado"
    is_posgrado = "posgrado" in tipo_norm

    threshold = 3.5 if is_posgrado else 3.0

    # Todos los módulos deben cumplir el umbral mínimo
    if any(n < threshold for n in notas_vals):
        return "No aprobó", promedio

    return "Aprobó", promedio


def apply_rules_to_match(
    match: StudentMatch,
    codigo_curso: str,
    clase_id: str,
    aula_meta: Optional[AulaMetadata] = None,
) -> None:
    """Aplica las reglas oficiales a un match de estudiante, actualizando sus campos.

    Modifica in-place match.estudiante_sist con los valores calculados.
    """
    es = match.estudiante_sist
    en = match.estudiante_nota

    # Extraer calificaciones numéricas de módulos
    califs_dict: Dict[str, float] = {}
    for m_key, tv in en.calificaciones_modulos.items():
        if tv and tv.value is not None:
            try:
                califs_dict[m_key] = float(str(tv.value).replace(",", "."))
            except ValueError:
                pass

    # Extraer valor exacto de Calculated Final de Notas
    nota_final_val = None
    if en.nota_final and en.nota_final.value is not None:
        try:
            nota_final_val = float(str(en.nota_final.value).replace(",", "."))
        except (ValueError, TypeError):
            nota_final_val = en.nota_final.value

    tipo_grupo_str = str(es.tipo_grupo.value) if es.tipo_grupo and es.tipo_grupo.value else "Pregrado"

    estado, promedio = evaluate_certification_status(
        califs_dict,
        tipo_grupo=tipo_grupo_str,
        nota_final_calculated=nota_final_val if isinstance(nota_final_val, (int, float)) else None,
    )

    es.estado_calculado = estado
    is_aprobado = estado == "Aprobó"

    es.elaboro_certificado_calculado = CERTIFICADO_SI if is_aprobado else CERTIFICADO_NO
    es.codigo_certificado_calculado = str(codigo_curso).strip().upper() if is_aprobado else VALOR_NO_APLICA
    es.envio_certificado_calculado = CERTIFICADO_SI if is_aprobado else CERTIFICADO_NO

    # Fecha de finalización para la fecha de envío
    f_fin = None
    if aula_meta and aula_meta.fecha_fin and aula_meta.fecha_fin.value:
        f_fin = aula_meta.fecha_fin.value
    elif aula_meta and aula_meta.fecha_fin_2 and aula_meta.fecha_fin_2.value:
        f_fin = aula_meta.fecha_fin_2.value

    # Actualizar raw_row_data para reproducción exacta de las columnas de Sistematización
    col_map = es.column_indices if es.column_indices else {}
    c_m1 = col_map.get("modulo_1", 45)
    c_m2 = col_map.get("modulo_2", 46)
    c_m3 = col_map.get("modulo_3", 47)
    c_m4 = col_map.get("modulo_4", 48)
    c_m5 = col_map.get("modulo_5", 49)
    c_prom = col_map.get("promedio", 50)
    c_est = col_map.get("estado", 51)
    c_elab = col_map.get("elaboro", 52)
    c_cod = col_map.get("codigo_cert", 53)
    c_env = col_map.get("envio", 54)
    c_fenv = col_map.get("fecha_envio", 55)

    if "modulo_1" in califs_dict:
        es.raw_row_data[c_m1] = califs_dict["modulo_1"]
    if "modulo_2" in califs_dict:
        es.raw_row_data[c_m2] = califs_dict["modulo_2"]
    if "modulo_3" in califs_dict:
        es.raw_row_data[c_m3] = califs_dict["modulo_3"]
    if "modulo_4" in califs_dict:
        es.raw_row_data[c_m4] = califs_dict["modulo_4"]
    if "modulo_5" in califs_dict:
        es.raw_row_data[c_m5] = califs_dict["modulo_5"]

    # Copiar y pegar directamente Calculated Final en la columna de promedio
    es.raw_row_data[c_prom] = promedio
    es.raw_row_data[c_est] = es.estado_calculado
    es.raw_row_data[c_elab] = es.elaboro_certificado_calculado
    es.raw_row_data[c_cod] = es.codigo_certificado_calculado
    es.raw_row_data[c_env] = es.envio_certificado_calculado
    if f_fin:
        es.raw_row_data[c_fenv] = str(f_fin).upper()
