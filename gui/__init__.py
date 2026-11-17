"""Interfaz gráfica de escritorio del Asistente de Cierre de Aulas COIN.

La aplicación es una capa de presentación sobre los mismos módulos que usa el
menú de consola (`main.py`): no reimplementa reglas, parsers ni escritura de
archivos. Cada acción invoca exactamente las mismas funciones del proyecto
(`run_cierre_pipeline`, `apply_direct_cierre`, `restore_last_backup`,
`diagnose_xlsx`, `repair_xlsx`, `run_discovery`, `run_diagnostico`).

Sigue siendo 100% local: no abre conexiones de red, no usa telemetría y no
depende de ninguna librería adicional a `openpyxl` (la interfaz usa `tkinter`,
incluido en Python).
"""

from __future__ import annotations

__all__ = ["CierreApp"]


def run() -> None:
    """Abre la aplicación de escritorio."""
    from gui.app import CierreApp

    CierreApp().mainloop()
