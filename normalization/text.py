"""
Normalización general de texto para el Asistente de Cierre de Aulas COIN.

Proporciona funciones para limpiar y normalizar cadenas de texto:
- Eliminación de tildes y diacríticos.
- Colapso de espacios múltiples.
- Remoción de signos de puntuación irrelevantes.
- Normalización de caracteres especiales.

Estas funciones son la base del sistema de detección flexible de encabezados
y de la normalización de nombres de estudiantes.
"""

from __future__ import annotations

import re
import unicodedata


def remove_accents(text: str) -> str:
    """Elimina tildes y diacríticos de un texto, preservando la ñ.

    Ejemplos:
        "Búsqueda ética" → "Busqueda etica"
        "módulo" → "modulo"
        "Año" → "Ano"  (la ñ se trata como caso especial si se necesita)

    Args:
        text: Texto con posibles tildes.

    Returns:
        Texto sin tildes.
    """
    # Preservar ñ/Ñ
    text = text.replace("ñ", "\x00n").replace("Ñ", "\x00N")

    # Descomponer caracteres Unicode y eliminar marcas de combinación
    nfkd = unicodedata.normalize("NFKD", text)
    result = "".join(c for c in nfkd if unicodedata.category(c) != "Mn")

    # Restaurar ñ/Ñ
    result = result.replace("\x00n", "ñ").replace("\x00N", "Ñ")

    return result


def normalize_whitespace(text: str) -> str:
    """Colapsa espacios múltiples, tabs y newlines en un solo espacio.

    Ejemplos:
        "  Nombre   Apellido  " → "Nombre Apellido"
        "Hoja\\t1" → "Hoja 1"

    Args:
        text: Texto con posibles espacios irregulares.

    Returns:
        Texto con espacios normalizados y sin espacios iniciales/finales.
    """
    return re.sub(r"\s+", " ", text).strip()


def normalize_text(text: str | None) -> str:
    """Normalización estándar completa de un texto.

    Aplica:
    1. Strip de espacios.
    2. Conversión a minúsculas.
    3. Eliminación de tildes.
    4. Colapso de espacios múltiples.

    Args:
        text: Texto a normalizar (puede ser None).

    Returns:
        Texto normalizado. String vacío si la entrada es None.
    """
    if text is None:
        return ""

    if not isinstance(text, str):
        text = str(text)

    text = text.strip()
    text = text.lower()
    text = remove_accents(text)
    text = normalize_whitespace(text)

    return text


def normalize_for_comparison(text: str | None) -> str:
    """Normalización agresiva para comparaciones de equivalencia.

    Además de la normalización estándar, elimina:
    - Signos de puntuación.
    - Paréntesis.
    - Barras.
    - Guiones.
    - Caracteres especiales.

    Útil para comparar encabezados de columnas.

    Args:
        text: Texto a normalizar.

    Returns:
        Texto normalizado para comparación.
    """
    text = normalize_text(text)
    # Eliminar caracteres no alfanuméricos excepto espacios y ñ
    text = re.sub(r"[^\w\sñ]", " ", text)
    text = normalize_whitespace(text)
    return text


def normalize_document(document: str | None) -> str:
    """Normaliza un número de documento de identidad para matching.

    Elimina puntos, guiones, espacios y otros caracteres no numéricos.

    Ejemplos:
        "1.020.304.050" → "1020304050"
        "10-203-040-50" → "1020304050"
        " 1020304050 " → "1020304050"
        None → ""

    Args:
        document: Número de documento.

    Returns:
        Documento normalizado (solo dígitos).
    """
    if document is None:
        return ""

    if not isinstance(document, str):
        document = str(document)

    # Conservar únicamente dígitos
    return re.sub(r"[^\d]", "", document.strip())


def normalize_email(email: str | None) -> str:
    """Normaliza un correo electrónico para matching.

    Convierte a minúsculas y elimina espacios.

    Ejemplos:
        " Juan.Perez@EAFIT.edu.co " → "juan.perez@eafit.edu.co"
        None → ""

    Args:
        email: Dirección de correo.

    Returns:
        Correo normalizado.
    """
    if email is None:
        return ""

    if not isinstance(email, str):
        email = str(email)

    return email.strip().lower()


def normalize_name_component(name: str | None) -> str:
    """Normaliza un componente de nombre (nombre o apellido) para matching.

    Aplica normalización de texto + elimina caracteres especiales.
    Preserva la ñ que es relevante en nombres colombianos.

    Ejemplos:
        "  Juan Carlos  " → "juan carlos"
        "PÉREZ GARCÍA" → "perez garcia"
        "María José" → "maria jose"

    Args:
        name: Componente del nombre.

    Returns:
        Nombre normalizado.
    """
    if name is None:
        return ""

    if not isinstance(name, str):
        name = str(name)

    name = name.strip().lower()
    name = remove_accents(name)
    name = normalize_whitespace(name)

    return name
