"""
Motor de Reglas de Negocio Oficiales y Cálculo de Certificación.

Regla oficial de certificación:

* ``Abandonó``: todas las notas de los módulos están en 0 (si el curso no trae
  módulos, cuando no hay nota final o es 0).
* ``Aprobó`` (Pregrado): la nota general (nota final de Notas; si no viene, el
  promedio de los módulos) es 3.0 o más.
* ``Aprobó`` (Posgrado): la nota final es 3.5 o más **y** ninguna nota de módulo
  está por debajo de 3.5.
* ``No aprobó``: cualquier otro caso con al menos una nota distinta de 0.

Además gestiona los formatos de:
- Se elaboró certificado SI/NO: 'SI' para Aprobó, 'NO' en otro caso.
- Código del certificado: Código del curso para SI, 'No aplica' para NO.
- Fecha de envío: fecha de finalización de la clase si Aprobó; 'No aplica' si
  No aprobó o Abandonó.
- Docente COIN: '{CÓDIGO-GUIÓN} {FORMADOR_LÍDER}'.
- Ciclo: 'CLASE {clase} {catalogo} - {programa} - {fecha_ini} A {fecha_fin}'.
"""

from __future__ import annotations

from datetime import date, datetime
from typing import Any, Dict, List, Optional, Tuple

from config.settings import (
    CERTIFICADO_NO,
    CERTIFICADO_SI,
    PREFIJO_TEXTO_MODULOS,
    VALOR_NO_APLICA,
)
from core.models import AulaMetadata, StudentMatch
from normalization.dates import (
    format_certificate_date,
    format_ciclo_date,
    get_semester,
    parse_date,
)
from normalization.text import (
    format_ciclo,
    format_docente_coin,
)


def evaluate_certification_status(
    calificaciones: Dict[str, float],
    tipo_grupo: Optional[str] = None,
    nota_final_calculated: Optional[float] = None,
) -> Tuple[str, Optional[float]]:
    """Evalúa el estado de certificación oficial de un estudiante.

    Regla oficial:

    * ``Abandonó``: todas las notas de los módulos están en 0 (si el curso no
      trae notas de módulos, cuando no hay nota final o esta es 0).
    * ``Aprobó`` (Pregrado): la **nota general** es 3.0 o más.
    * ``Aprobó`` (Posgrado): la **nota final** es 3.5 o más **y** ninguna nota de
      módulo está por debajo de 3.5.
    * ``No aprobó``: cualquier otro caso (tiene al menos una nota distinta de 0).

    La *nota general* es la nota final del archivo de Notas (``Calculated
    Final`` y equivalentes). Si Notas no la trae, se usa el promedio de los
    módulos, solo para clasificar; **nunca** se escribe ese promedio en la
    columna de nota final.

    Args:
        calificaciones: Diccionario {'modulo_1': nota, 'modulo_2': nota, ...}
        tipo_grupo: 'Pregrado', 'Posgrado' o None (default Pregrado).
        nota_final_calculated: Valor de 'Calculated Final' extraído directamente de Notas.

    Returns:
        Tupla (estado, nota_final):
          - estado: 'Aprobó', 'No aprobó' o 'Abandonó'.
          - nota_final: **exactamente** el valor de la nota final de Notas
            (puede ser None si el archivo no la trae). NUNCA se calcula un
            promedio para escribir: los porcentajes y ponderaciones del curso
            son potestad del archivo de Notas (Calculated Final).
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

    # La nota final se COPIA TAL CUAL de Notas; si no existe, queda None.
    promedio = nota_final_calculated

    # 1. "Abandonó": todas las notas de los módulos están en 0. Si el curso no
    #    trae módulos, se considera abandono cuando no hay nota final o es 0.
    #    Ya NO se marca abandonó por tener los módulos 3 y 4 en 0: solo cuando
    #    TODAS las notas están en 0.
    if notas_vals:
        if all(n == 0.0 for n in notas_vals):
            return "Abandonó", promedio
    elif promedio is None or promedio == 0.0:
        return "Abandonó", promedio

    # Nota general para decidir la aprobación: la nota final de Notas; si no
    # viene (o no es numérica), el promedio de los módulos (solo para clasificar).
    promedio_modulos = (
        round(sum(notas_vals) / len(notas_vals), 2) if notas_vals else None
    )
    nota_general = promedio if promedio is not None else promedio_modulos

    # 2. Umbral según tipo de grupo (Pregrado >= 3.0, Posgrado >= 3.5)
    tipo_norm = str(tipo_grupo).strip().lower() if tipo_grupo else "pregrado"
    is_posgrado = "posgrado" in tipo_norm
    umbral = 3.5 if is_posgrado else 3.0

    aprobado = nota_general is not None and nota_general >= umbral

    # Posgrado: además, ninguna nota de módulo puede quedar por debajo de 3.5.
    if is_posgrado:
        aprobado = aprobado and all(n >= umbral for n in notas_vals)

    return ("Aprobó" if aprobado else "No aprobó"), promedio


def _texto_modulos_faltantes(num_modulos: int) -> str:
    """Texto informativo para las columnas de módulos que el curso no tiene.

    Ej.: ``'El curso solo tiene 3 módulos'`` (o ``'... 1 módulo'`` en singular).
    """
    unidad = "módulo" if num_modulos == 1 else "módulos"
    return f"{PREFIJO_TEXTO_MODULOS}{num_modulos} {unidad}"


def _format_fecha_envio_sistematizacion(value: Any) -> str:
    """Formatea la fecha de envío para la columna de Sistematización.

    Formato oficial confirmado: 'DD-mes-AAAA' con mes abreviado en minúsculas
    (ej. '24-mar-2026'), idéntico al usado en la columna 'Fecha de envío' del
    archivo de Certificados.

    Se evita deliberadamente el `str(datetime)` (que devolvía
    '2026-03-24 00:00:00', arrastrando la hora en ceros).

    Args:
        value: Fecha (date/datetime), cadena de fecha o serial de Excel.

    Returns:
        Cadena 'DD-mes-AAAA' o '' si el valor no es una fecha interpretable.
    """
    fecha = parse_date(value)
    if fecha is None:
        return ""
    return format_certificate_date(fecha)


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

    # Extraer la nota final EXACTA de Notas ('Calculated Final' o equivalente).
    # Se copia tal cual: no se recalcula, no se promedia y no se reemplaza por
    # una fórmula. Si Notas no la trae, no se escribe nada en Sistematización.
    nota_final_origen: Any = None
    if en.nota_final and en.nota_final.value is not None:
        valor = en.nota_final.value
        if isinstance(valor, str):
            texto = valor.strip()
            # Una fórmula sin valor calculado no debe copiarse como fórmula a los
            # archivos oficiales: se trata como nota ausente.
            if texto.startswith("=") or texto == "":
                nota_final_origen = None
            else:
                nota_final_origen = valor
        else:
            nota_final_origen = valor

    # Valor numérico solo para comparaciones internas (regla de abandono).
    nota_final_num: Optional[float] = None
    if nota_final_origen is not None:
        try:
            nota_final_num = float(str(nota_final_origen).replace(",", "."))
        except (ValueError, TypeError):
            nota_final_num = None

    tipo_grupo_str = str(es.tipo_grupo.value) if es.tipo_grupo and es.tipo_grupo.value else "Pregrado"

    estado, promedio = evaluate_certification_status(
        califs_dict,
        tipo_grupo=tipo_grupo_str,
        nota_final_calculated=nota_final_num,
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

    # La columna 'Fecha de envío' de Sistematización se escribe como fecha oficial
    # 'DD-mes-AAAA' (ej. '24-mar-2026'), NUNCA como el repr de un datetime
    # (ej. '2026-03-24 00:00:00'), que además arrastraba la hora en ceros.
    # REGLA OFICIAL: solo los aprobados tienen fecha de envío; si no aprobó o
    # abandonó, la columna dice 'No aplica'.
    if is_aprobado:
        fecha_envio_txt = _format_fecha_envio_sistematizacion(f_fin)
    else:
        fecha_envio_txt = VALOR_NO_APLICA

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

    # Relleno informativo de las columnas de módulos que el curso no tiene.
    # X = número de módulos del curso (columnas de módulos detectadas en Notas,
    # sin contar la nota final / promedio del curso completo). Si el curso tiene
    # menos de 5, las columnas restantes no se dejan vacías. Nunca se pisa un
    # valor que ya exista en esa celda de Sistematización.
    modulos_curso = len(en.calificaciones_modulos or {})
    if 0 < modulos_curso < 5:
        texto_modulos = _texto_modulos_faltantes(modulos_curso)
        for n in range(modulos_curso + 1, 6):
            c_mod = col_map.get(f"modulo_{n}", 44 + n)
            valor_actual = es.raw_row_data.get(c_mod)
            if valor_actual is None or str(valor_actual).strip() == "":
                es.raw_row_data[c_mod] = texto_modulos

    # Copiar y pegar la nota final TAL CUAL viene del archivo de Notas.
    # Si Notas no trae nota final, la celda de Sistematización se deja vacía:
    # NUNCA se inventa un promedio de módulos.
    if nota_final_origen is not None:
        es.raw_row_data[c_prom] = nota_final_origen
    es.raw_row_data[c_est] = es.estado_calculado
    es.raw_row_data[c_elab] = es.elaboro_certificado_calculado
    es.raw_row_data[c_cod] = es.codigo_certificado_calculado
    es.raw_row_data[c_env] = es.envio_certificado_calculado
    if fecha_envio_txt:
        es.fecha_envio_calculada = fecha_envio_txt
        es.raw_row_data[c_fenv] = fecha_envio_txt
