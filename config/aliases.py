"""
Tabla de alias semánticos de encabezados para el Asistente de Cierre de Aulas COIN.

Mapea encabezados normalizados (minúsculas, sin tildes, sin puntuación)
a nombres canónicos internos del sistema.

IMPORTANTE:
    Esta tabla se refinará y ampliará en la Fase 0 tras inspeccionar
    los archivos reales. Los alias actuales están basados en la
    especificación documentada del procedimiento manual.

    El sistema de normalización aplica normalize_for_comparison()
    antes de buscar en esta tabla, por lo que las claves deben estar
    en formato normalizado (minúsculas, sin tildes, sin puntuación extra).
"""

# ============================================================
# ALIAS DE ENCABEZADOS
# ============================================================
# Formato: "encabezado normalizado" → "nombre_canonico_interno"
# El texto normalizado elimina tildes, puntuación y usa minúsculas.

HEADER_ALIASES: dict[str, str] = {

    # ============================================================
    # ARCHIVO DE NOTAS — Identificadores de estudiante
    # ============================================================
    "orgdefinedid": "org_defined_id",
    "org defined id": "org_defined_id",

    "username": "username",
    "user name": "username",

    "last name": "apellidos",
    "lastname": "apellidos",
    "apellido": "apellidos",
    "apellidos": "apellidos",

    "first name": "nombres",
    "firstname": "nombres",
    "nombre": "nombres",
    "nombres": "nombres",

    "seccion": "seccion",
    "sección": "seccion",  # por si no se normalizó la tilde
    "section": "seccion",

    # ============================================================
    # ARCHIVO DE NOTAS — Nota final (variantes conocidas)
    # ============================================================
    # Nota: la detección principal de nota final se hace por regex
    # en headers.py. Estos aliases son respaldo para matches exactos.
    "calculated final grade": "nota_final",
    "calculated final": "nota_final",
    "final grade": "nota_final",
    "nota final": "nota_final",
    "promedio final": "nota_final",

    # ============================================================
    # ARCHIVO DE SISTEMATIZACIÓN — Campos principales
    # ============================================================
    "modalidad": "modalidad",
    "plataforma": "plataforma",
    "aula presencial": "aula_presencial",
    "tipo de grupo": "tipo_de_grupo",
    "aula virtual solicitud": "aula_virtual",
    "aula virtual  solicitud": "aula_virtual",
    "solicitante": "solicitante",
    "tipo de usuario": "tipo_de_usuario",
    "programa academico departamento academico dependencia": "programa_academico_sist",
    "curso presencial virtual": "curso_presencial_virtual",
    "modulos": "modulos",
    "dia inicial": "dia_inicial",
    "mes inicial": "mes_inicial",
    "ano inicial": "ano_inicial",
    "año inicial": "ano_inicial",
    "semestre": "semestre",
    "fecha final dd mm aaaa": "fecha_final",
    "fecha final ddmmaaaa": "fecha_final",
    "fecha final": "fecha_final",

    "hora inicial presencial o seminario web": "hora_inicial",
    "hora final presencial o seminario web": "hora_final",
    "total horas presencial o seminario web": "total_horas_presencial",
    "total horas virtual": "total_horas_virtual",

    "tipo de documento de identidad": "tipo_documento",
    "numero de documento de identidad": "documento",
    "correo electronico": "correo",
    "correo electrónico": "correo",

    "tipo de usuario 2": "tipo_usuario_2",
    "departamento academico solo para docentes": "departamento_academico",
    "programa academico dependencia": "programa_academico_dependencia",
    "escuela": "escuela",
    "nombre materia ciclo v p eafit externo": "nombre_materia_ciclo",
    "catalogo": "catalogo",
    "clase": "clase",
    "formador lider": "formador_lider",
    "formador acompanante": "formador_acompanante",

    "autorizacion tratamiento de datos": "autorizacion_datos",
    "ciudad de residencia": "ciudad_residencia",
    "departamento de residencia": "departamento_residencia",
    "pais de residencia": "pais_residencia",
    "institucion externos": "institucion_externos",
    "programa academico dependencia externos": "programa_externos",

    "calificacion modulo 1": "modulo_1",
    "calificacion modulo 2": "modulo_2",
    "calificacion modulo 3": "modulo_3",
    "calificacion modulo 4": "modulo_4",
    "promedio curso completo": "promedio_curso",

    "estado para certificacion": "estado_certificacion",
    "se elaboro certificado si no": "elaboro_certificado",
    "se elaboro certificado sino": "elaboro_certificado",
    "codigo del certificado": "codigo_certificado",
    "se envio certificado si no": "envio_certificado",
    "se envio certificado sino": "envio_certificado",
    "fecha de envio": "fecha_envio",

    # ============================================================
    # ARCHIVO DE CONTROL DE AULAS — Campos principales
    # ============================================================
    "ano": "ano",
    "año": "ano",
    "tipo de aula": "tipo_aula",
    "curso que realiza": "curso_que_realiza",
    "programa academico": "programa_academico_ctrl",
    "materia donde se cargan los estudiantes": "materia_carga",
    "nombre de la seccion": "nombre_seccion",
    "profesor a solicitante": "profesor_solicitante",
    "profesora solicitante": "profesor_solicitante",
    "profesor a al que se le reporta": "profesor_reporta",
    "profesora al que se le reporta": "profesor_reporta",
    "materia a la que se reporta": "materia_reporta",
    "agregar como estudiantes": "agregar_estudiantes",
    "porcentaje evaluativo": "porcentaje_evaluativo",
    "semana": "semana",
    "fecha de inicio del curso": "fecha_inicio_ctrl",
    "fecha de finalizacion del curso": "fecha_fin_ctrl",
    "fecha 2 de finalizacion del curso": "fecha_fin_2_ctrl",
    "fecha desactivacion del aula": "fecha_desactivacion",
    "siguiente aula para asignar": "siguiente_aula",
    "docente principal": "docente_principal",
    "certificados y cierre de curso": "certificados_cierre",
    "certificados": "certificados",
    "envio de notas a profesores": "envio_notas_profesores",
    "envío de notas a profesores": "envio_notas_profesores",
    "reporte final de notas formadores": "reporte_notas_formadores",

    # ============================================================
    # ARCHIVO DE CERTIFICADOS — Campos principales
    # ============================================================
    "n": "numero_certificado",
    "no": "numero_certificado",
    "n°": "numero_certificado",
    "ciclo": "ciclo",
    "docente coin": "docente_coin",
    "documento de identidad": "documento_cert",
    "correo electronico cert": "correo_cert",
    "codigo del certificado cert": "codigo_certificado_cert",
    "fecha de envio cert": "fecha_envio_cert",
    "ano de certificacion": "ano_certificacion",
    "año de certificacion": "ano_certificacion",
    "total certificados": "total_certificados",
    "fecha de totalizacion": "fecha_totalizacion",
}
