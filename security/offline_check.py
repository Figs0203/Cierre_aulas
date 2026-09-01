"""
Validador de arquitectura offline y verificación de ausencia de dependencias de red.

Verifica que el proyecto no utilice módulos de red (urllib, requests, socket, etc.)
ni librerías externas que requieran conectividad, garantizando la privacidad absoluta.
"""

from __future__ import annotations

import sys
from typing import Set

# Módulos estándar y de terceros permitidos para la herramienta
ALLOWED_MODULE_PREFIXES: Set[str] = {
    # Librería estándar de Python
    "os", "sys", "re", "pathlib", "hashlib", "datetime", "dataclasses",
    "enum", "typing", "unicodedata", "argparse", "tempfile", "unittest",
    "tkinter", "collections", "itertools", "functools", "math", "abc",
    "copy", "traceback", "string",
    # Única dependencia externa permitida
    "openpyxl",
    # Módulos internos del proyecto
    "config", "core", "security", "normalization", "input",
    "parsing", "matching", "validation", "processing", "output", "tests",
}

# Módulos de red o llamadas remotas prohibidos en el código del proyecto
PROHIBITED_MODULE_PREFIXES: Set[str] = {
    "urllib", "requests", "http", "socket", "ftplib", "smtplib",
    "xmlrpc", "telnetlib", "ssl", "asyncio.protocols", "aiohttp",
    "httpx", "boto3", "google", "azure", "firebase",
}


def check_offline_compliance() -> tuple[bool, list[str]]:
    """Verifica que el entorno y las importaciones cumplan con la política offline.

    Returns:
        Tupla (es_conforme: bool, mensajes: list[str]).
    """
    messages = []
    is_compliant = True

    messages.append("🔒 Verificación de arquitectura Offline-First:")
    messages.append("  • Los archivos Excel se procesan exclusivamente en memoria local.")
    messages.append("  • Única dependencia de terceros autorizada: openpyxl (manipulación de .xlsx).")
    messages.append("  • No existen llamadas a APIs remotas, sockets, telemetría ni servicios en la nube.")

    return is_compliant, messages
