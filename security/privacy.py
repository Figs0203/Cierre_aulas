"""
Enmascaramiento de información personal (PII) para el Asistente de Cierre de Aulas COIN.

Proporciona funciones para enmascarar datos personales en mensajes de consola
y logs, garantizando que la información sensible de estudiantes no quede
expuesta innecesariamente fuera del archivo auxiliar de trabajo.

Principio de diseño:
    - El archivo Excel auxiliar SÍ contiene los datos completos (es documento de trabajo).
    - Los mensajes de consola, logs y reportes de error SIEMPRE enmascaran PII.
    - Nunca se registran documentos completos, correos completos ni nombres
      completos en la salida estándar.
"""

from __future__ import annotations

import re


def mask_email(email: str | None) -> str:
    """Enmascara un correo electrónico preservando parte del usuario y el dominio.

    Ejemplos:
        "juan.perez@eafit.edu.co" → "juan.p***@eafit.edu.co"
        "a@b.co" → "a***@b.co"
        None → ""
        "" → ""

    Args:
        email: Dirección de correo electrónico.

    Returns:
        Correo enmascarado.
    """
    if not email or not isinstance(email, str):
        return ""

    email = email.strip()

    if "@" not in email:
        return _mask_generic(email)

    local, domain = email.rsplit("@", 1)

    if len(local) <= 1:
        masked_local = local + "***"
    elif len(local) <= 4:
        masked_local = local[:2] + "***"
    else:
        # Preservar los primeros 5 caracteres o hasta el primer punto + 1 carácter
        dot_pos = local.find(".")
        if dot_pos > 0 and dot_pos < len(local) - 1:
            visible = min(dot_pos + 2, 6)
        else:
            visible = min(5, len(local) - 1)
        masked_local = local[:visible] + "***"

    return f"{masked_local}@{domain}"


def mask_document(document: str | None) -> str:
    """Enmascara un número de documento preservando los últimos 4 dígitos.

    Ejemplos:
        "1020304050" → "****4050"
        "12345678" → "****5678"
        "1234" → "****1234"
        "123" → "****123"
        None → ""

    Args:
        document: Número de documento de identidad.

    Returns:
        Documento enmascarado.
    """
    if not document or not isinstance(document, str):
        return ""

    # Limpiar: remover puntos, guiones, espacios
    clean = re.sub(r"[\s.\-]", "", document.strip())

    if not clean:
        return ""

    if len(clean) <= 4:
        return f"****{clean}"

    return f"****{clean[-4:]}"


def mask_name(name: str | None) -> str:
    """Enmascara un nombre preservando el primer nombre y la inicial del resto.

    Ejemplos:
        "Carlos Alberto Restrepo" → "Carlos R."
        "Juan Pérez" → "Juan P."
        "María" → "María"
        None → ""

    Args:
        name: Nombre completo.

    Returns:
        Nombre enmascarado.
    """
    if not name or not isinstance(name, str):
        return ""

    parts = name.strip().split()

    if not parts:
        return ""

    if len(parts) == 1:
        return parts[0]

    # Conservar el primer nombre y la inicial del último componente
    first = parts[0]
    last_initial = parts[-1][0].upper() + "." if parts[-1] else ""

    return f"{first} {last_initial}"


def mask_full_name(first_name: str | None, last_name: str | None) -> str:
    """Enmascara nombre y apellido combinados.

    Ejemplos:
        ("Juan", "Pérez García") → "Juan P."
        ("María", None) → "María"

    Args:
        first_name: Primer nombre o nombres.
        last_name: Apellido(s).

    Returns:
        Nombre enmascarado.
    """
    parts = []
    if first_name and isinstance(first_name, str):
        parts.append(first_name.strip().split()[0])
    if last_name and isinstance(last_name, str):
        last_parts = last_name.strip().split()
        if last_parts:
            parts.append(last_parts[0][0].upper() + ".")

    return " ".join(parts) if parts else ""


def _mask_generic(text: str) -> str:
    """Enmascaramiento genérico para texto no categorizado.

    Preserva los primeros 3 caracteres y reemplaza el resto con asteriscos.
    """
    if len(text) <= 3:
        return text + "***"
    return text[:3] + "***"


def mask_for_log(
    *,
    email: str | None = None,
    document: str | None = None,
    first_name: str | None = None,
    last_name: str | None = None,
) -> str:
    """Genera un identificador enmascarado para usar en logs.

    Combina los identificadores disponibles en un formato seguro.
    Prioriza documento > correo > nombre.

    Returns:
        String como "Estudiante (Doc: ****4050)" o
        "Estudiante (juan.p***@eafit.edu.co)" o
        "Estudiante (Juan P.)"
    """
    if document:
        return f"Estudiante (Doc: {mask_document(document)})"
    if email:
        return f"Estudiante ({mask_email(email)})"
    if first_name or last_name:
        name = mask_full_name(first_name, last_name)
        return f"Estudiante ({name})"
    return "Estudiante (sin identificador)"
