"""Inicio de la aplicación de escritorio del Asistente de Cierre de Aulas COIN.

Uso:
    python app.py

También se puede abrir con doble clic en «Cierre de Aulas.bat», que evita la
ventana de consola. El menú de consola original (`python main.py`) sigue
disponible sin cambios.
"""

from __future__ import annotations

import sys
import warnings
from pathlib import Path

_RAIZ = Path(__file__).resolve().parent
if str(_RAIZ) not in sys.path:
    sys.path.insert(0, str(_RAIZ))

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")


def main() -> int:
    try:
        import tkinter  # noqa: F401
    except Exception:  # noqa: BLE001 - instalación de Python sin Tcl/Tk
        print(
            "No se encontró tkinter en esta instalación de Python.\n"
            "En Windows, reinstale Python marcando la opción «tcl/tk and IDLE»."
        )
        return 1

    from gui.app import CierreApp

    try:
        app = CierreApp()
    except Exception as exc:  # noqa: BLE001 - por ejemplo, sin entorno gráfico
        print(f"No se pudo abrir la ventana de la aplicación: {exc}")
        return 1

    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
