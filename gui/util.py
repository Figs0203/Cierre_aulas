"""Utilidades de presentación compartidas por las páginas de la aplicación."""

from __future__ import annotations

import os
import re
import subprocess
import sys
from pathlib import Path

# Símbolos de estado que usan los mensajes del dominio (consola). En la interfaz
# se sustituyen por color, no por iconos.
_SIMBOLOS = re.compile(r"[\U0001F300-\U0001FAFF\u2190-\u21FF\u2600-\u27BF\uFE0F\u200D]")


def elide(texto: str, maximo: int = 88) -> str:
    """Recorta un texto por el medio conservando el final (más informativo)."""
    texto = str(texto)
    if len(texto) <= maximo:
        return texto
    if maximo < 12:
        return texto[:maximo]
    inicio = maximo // 3
    return f"{texto[:inicio]}…{texto[-(maximo - inicio - 1):]}"


def limpiar_mensaje(mensaje: str) -> str:
    """Quita los símbolos de estado y normaliza espacios de un mensaje."""
    limpio = _SIMBOLOS.sub("", str(mensaje))
    limpio = re.sub(r"[ \t]+", " ", limpio)
    return limpio.strip()


def inferir_aula(path: Path | str) -> str:
    """Deduce el código de aula del nombre del archivo de Notas.

    Misma heurística que usa el cierre por lotes del menú de consola:
    'ABBUEI227_E_FINANZAS_MAF_1.xlsx' → 'ABBUEI227'.
    """
    match = re.match(r"([A-Za-z]{2,}\d{2,})", Path(path).stem)
    return match.group(1).upper() if match else ""


def abrir_ruta(ruta: Path | str) -> bool:
    """Abre un archivo o carpeta con la aplicación predeterminada del sistema."""
    destino = str(Path(ruta))
    try:
        if sys.platform == "win32":
            os.startfile(destino)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", destino])
        else:
            subprocess.Popen(["xdg-open", destino])
        return True
    except Exception:  # noqa: BLE001 - solo es una ayuda de navegación
        return False


def resumen_error(exc: BaseException) -> str:
    """Mensaje legible para un error de ejecución."""
    detalle = limpiar_mensaje(str(exc)).strip()
    return detalle or exc.__class__.__name__


def carpeta_reportes() -> Path:
    """Carpeta donde el proyecto guarda los informes locales."""
    raiz = Path(__file__).resolve().parent.parent
    carpeta = raiz / "reports"
    carpeta.mkdir(parents=True, exist_ok=True)
    return carpeta
