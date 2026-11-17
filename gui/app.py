"""Ventana principal de la aplicación de escritorio.

Capa de presentación: navega entre páginas, ejecuta las operaciones del proyecto
en un hilo secundario (para que la ventana nunca se congele) y muestra el
registro de actividad. Toda la lógica de cierre sigue viviendo en los módulos
originales del proyecto.
"""

from __future__ import annotations

import contextlib
import importlib
import queue
import sys
import threading
import traceback
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk

from gui import theme
from gui.util import limpiar_mensaje, resumen_error
from gui.widgets import Console, Pill

_PROJECT_ROOT = Path(__file__).resolve().parent.parent
if str(_PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(_PROJECT_ROOT))

# Nitidez en pantallas de Windows con escalado distinto de 100 %.
if sys.platform == "win32":
    try:
        import ctypes

        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:  # noqa: BLE001 - versiones antiguas de Windows
            ctypes.windll.user32.SetProcessDPIAware()
    except Exception:  # noqa: BLE001 - nunca impide arrancar
        pass


# ---------------------------------------------------------------------------
# Definición de páginas
# ---------------------------------------------------------------------------
DEFINICIONES: dict[str, dict[str, str]] = {
    "cierre": {
        "label": "Cierre de aula",
        "titulo": "Cierre de aula",
        "subtitulo": "Cruza las notas con la sistematización y deja el aula lista para certificar.",
        "modulo": "gui.pages.cierre_aula",
        "clase": "CierreAulaPage",
    },
    "lote": {
        "label": "Cierre por lotes",
        "titulo": "Cierre por lotes",
        "subtitulo": "Cierra varias aulas en una sola operación compartiendo Sistematización y Control.",
        "modulo": "gui.pages.cierre_lote",
        "clase": "CierreLotePage",
    },
    "restaurar": {
        "label": "Restaurar respaldo",
        "titulo": "Restaurar respaldo",
        "subtitulo": "Devuelve los archivos oficiales al estado previo al último cierre.",
        "modulo": "gui.pages.restaurar",
        "clase": "RestaurarPage",
    },
    "diagnostico": {
        "label": "Diagnóstico",
        "titulo": "Diagnóstico",
        "subtitulo": "Informes técnicos locales, sin datos personales, para entender un cruce o un archivo.",
        "modulo": "gui.pages.diagnostico",
        "clase": "DiagnosticoPage",
    },
    "reparar": {
        "label": "Reparar archivo",
        "titulo": "Reparar archivo",
        "subtitulo": "Diagnostica un libro .xlsx dañado y genera una copia reparada sin tocar el original.",
        "modulo": "gui.pages.reparar",
        "clase": "RepararPage",
    },
    "seguridad": {
        "label": "Seguridad y privacidad",
        "titulo": "Seguridad y privacidad",
        "subtitulo": "Cómo protege la herramienta los datos del cierre de aulas.",
        "modulo": "gui.pages.seguridad",
        "clase": "SeguridadPage",
    },
}

SECCIONES: tuple[tuple[str, tuple[str, ...]], ...] = (
    ("OPERACIÓN", ("cierre", "lote")),
    ("MANTENIMIENTO", ("restaurar", "diagnostico", "reparar")),
    ("SISTEMA", ("seguridad",)),
)


# ---------------------------------------------------------------------------
# Utilidades internas
# ---------------------------------------------------------------------------
class _QueueStream:
    """Sustituto de `sys.stdout` que envía cada línea a la cola de la interfaz."""

    def __init__(self, cola: "queue.Queue"):
        self._cola = cola
        self._buffer = ""

    def write(self, texto: str) -> int:
        if not texto:
            return 0
        self._buffer += texto
        while "\n" in self._buffer:
            linea, self._buffer = self._buffer.split("\n", 1)
            self._cola.put(("texto", linea))
        return len(texto)

    def flush(self) -> None:
        if self._buffer:
            self._cola.put(("texto", self._buffer))
            self._buffer = ""

    def isatty(self) -> bool:
        return False


def _clasificar(linea: str) -> str:
    """Asigna un color del registro según el prefijo del mensaje del dominio."""
    limpio = linea.strip()
    if limpio.startswith("✅"):
        return "ok"
    if limpio.startswith(("⚠", "↳ ⚠", "!")):
        return "warn"
    if limpio.startswith("🔴"):
        return "error"
    if not limpio or set(limpio) <= {"=", "-"}:
        return "muted"
    return "info"


class NavItem(tk.Frame):
    """Elemento de navegación de la barra lateral."""

    def __init__(self, master, texto: str, command):
        super().__init__(master, bg=theme.SIDEBAR, cursor="hand2")
        self._command = command
        self._activo = False

        self._barra = tk.Frame(self, width=3, bg=theme.SIDEBAR)
        self._barra.pack(side="left", fill="y")
        self._label = tk.Label(
            self,
            text=texto,
            bg=theme.SIDEBAR,
            fg=theme.SIDEBAR_TEXT,
            font=theme.F_NAV,
            anchor="w",
            padx=17,
            pady=9,
            cursor="hand2",
        )
        self._label.pack(side="left", fill="x", expand=True)

        for widget in (self, self._label):
            widget.bind("<Button-1>", lambda _e: self._command())
            widget.bind("<Enter>", self._entrar)
            widget.bind("<Leave>", self._salir)

    def _entrar(self, _event=None):
        if not self._activo:
            self._pintar(theme.SIDEBAR_HOVER, theme.SIDEBAR_TEXT, theme.SIDEBAR_HOVER)

    def _salir(self, _event=None):
        if not self._activo:
            self._pintar(theme.SIDEBAR, theme.SIDEBAR_TEXT, theme.SIDEBAR)

    def _pintar(self, fondo: str, color: str, barra: str) -> None:
        self.configure(bg=fondo)
        self._label.configure(bg=fondo, fg=color)
        self._barra.configure(bg=barra)

    def set_activo(self, activo: bool) -> None:
        self._activo = activo
        if activo:
            self._pintar(theme.SIDEBAR_ACTIVE, "#FFFFFF", theme.ACCENT)
        else:
            self._pintar(theme.SIDEBAR, theme.SIDEBAR_TEXT, theme.SIDEBAR)


# ---------------------------------------------------------------------------
# Ventana principal
# ---------------------------------------------------------------------------
class CierreApp(tk.Tk):
    """Aplicación de escritorio del cierre de aulas."""

    VERSION = "1.0"

    def __init__(self):
        super().__init__()
        theme.init(self)

        self.title("Cierre de Aulas · Cursos COIN")
        # Tk informa píxeles físicos; en pantallas con escalado, las medidas del
        # gestor de ventanas están en unidades escaladas. Se normaliza para que la
        # ventana siempre quepa centrada, sea cual sea el factor de escala.
        factor = max(1.0, float(self.winfo_fpixels("1i")) / 96.0)
        ancho_pantalla = self.winfo_screenwidth() / factor
        alto_pantalla = self.winfo_screenheight() / factor
        ancho = int(max(960, min(1180, ancho_pantalla - 100)))
        alto = int(max(620, min(790, alto_pantalla - 90)))
        x = int(max(0, (ancho_pantalla - ancho) // 2))
        y = int(max(0, (alto_pantalla - alto) // 3))
        self.geometry(f"{ancho}x{alto}+{x}+{y}")
        self.minsize(940, 600)
        self.configure(bg=theme.PAGE)

        self._busy = False
        self._queue: "queue.Queue" = queue.Queue()
        self._pages: dict[str, tk.Frame] = {}
        self._nav: dict[str, NavItem] = {}
        self._active = ""
        self._on_success = None
        self._on_error = None
        self._log_visible = True

        self._build()
        self._bind_shortcuts()

        self.go("cierre")
        self.set_status("Listo")
        self.log(
            "Aplicación iniciada. Todo el procesamiento ocurre en este equipo: "
            "sin conexión a Internet y sin telemetría.",
            "muted",
        )
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ------------------------------------------------------------------
    # Construcción de la ventana
    # ------------------------------------------------------------------
    def _build(self) -> None:
        self._build_status_bar()

        contenedor = tk.Frame(self, bg=theme.PAGE)
        contenedor.pack(fill="both", expand=True)

        barra_lateral = tk.Frame(contenedor, bg=theme.SIDEBAR, width=252)
        barra_lateral.pack(side="left", fill="y")
        barra_lateral.pack_propagate(False)
        self._build_sidebar(barra_lateral)

        self._main = tk.Frame(contenedor, bg=theme.PAGE)
        self._main.pack(side="left", fill="both", expand=True)

        self._build_header()
        self._host = tk.Frame(self._main, bg=theme.PAGE)
        self._host.pack(fill="both", expand=True)
        self._build_log_dock()

    def _build_sidebar(self, sidebar: tk.Frame) -> None:
        marca = tk.Frame(sidebar, bg=theme.SIDEBAR, padx=22, pady=24)
        marca.pack(fill="x")
        tk.Label(
            marca,
            text="CIERRE DE AULAS",
            bg=theme.SIDEBAR,
            fg="#FFFFFF",
            font=(theme.F_H2[0], 12, "bold"),
            anchor="w",
        ).pack(fill="x")
        tk.Label(
            marca,
            text="Cursos COIN · Biblioteca",
            bg=theme.SIDEBAR,
            fg=theme.SIDEBAR_SECTION,
            font=theme.F_SMALL,
            anchor="w",
        ).pack(fill="x", pady=(4, 0))

        for seccion, claves in SECCIONES:
            tk.Label(
                sidebar,
                text=seccion,
                bg=theme.SIDEBAR,
                fg=theme.SIDEBAR_SECTION,
                font=theme.F_STEP,
                anchor="w",
            ).pack(fill="x", padx=24, pady=(16, 6))
            for clave in claves:
                item = NavItem(sidebar, DEFINICIONES[clave]["label"], lambda c=clave: self.go(c))
                item.pack(fill="x")
                self._nav[clave] = item

        pie = tk.Frame(sidebar, bg=theme.SIDEBAR, padx=22, pady=18)
        pie.pack(side="bottom", fill="x")
        tk.Frame(pie, bg="#26394B", height=1).pack(fill="x", pady=(0, 12))
        tk.Label(
            pie,
            text="Procesamiento 100 % local",
            bg=theme.SIDEBAR,
            fg=theme.SIDEBAR_TEXT,
            font=theme.F_SMALL,
            anchor="w",
        ).pack(fill="x")
        tk.Label(
            pie,
            text=f"Versión {self.VERSION}",
            bg=theme.SIDEBAR,
            fg=theme.SIDEBAR_SECTION,
            font=theme.F_SMALL,
            anchor="w",
        ).pack(fill="x", pady=(3, 0))

    def _build_header(self) -> None:
        cabecera = tk.Frame(self._main, bg=theme.PAGE)
        cabecera.pack(fill="x")

        fila = tk.Frame(cabecera, bg=theme.PAGE, padx=26, pady=18)
        fila.pack(fill="x", pady=(0, 14))

        self._pill = Pill(fila, "LOCAL · SIN CONEXIÓN")
        self._pill.pack(side="right", anchor="n", pady=(6, 0))

        self._title_var = tk.StringVar(value="Cierre de aula")
        self._subtitle_var = tk.StringVar(value="")
        tk.Label(
            fila,
            textvariable=self._title_var,
            bg=theme.PAGE,
            fg=theme.TEXT,
            font=theme.F_H1,
            anchor="w",
        ).pack(anchor="w")
        tk.Label(
            fila,
            textvariable=self._subtitle_var,
            bg=theme.PAGE,
            fg=theme.TEXT_MUTED,
            font=theme.F_SMALL,
            anchor="w",
        ).pack(anchor="w", pady=(5, 0))

        tk.Frame(cabecera, bg=theme.BORDER, height=1).pack(fill="x")

    def _build_log_dock(self) -> None:
        self._dock = tk.Frame(self._main, bg=theme.SURFACE)
        self._dock.pack(side="bottom", fill="x")
        tk.Frame(self._dock, bg=theme.BORDER, height=1).pack(fill="x", side="top")

        cabecera = tk.Frame(self._dock, bg=theme.SURFACE, padx=26, pady=8)
        cabecera.pack(fill="x", pady=(0, 4))

        tk.Label(
            cabecera,
            text="REGISTRO DE ACTIVIDAD",
            bg=theme.SURFACE,
            fg=theme.TEXT_MUTED,
            font=theme.F_STEP,
        ).pack(side="left")

        ttk.Button(
            cabecera, text="Limpiar", style="DarkLink.TButton", command=self._clear_log
        ).pack(side="right", padx=(6, 0))
        self._toggle_button = ttk.Button(
            cabecera, text="Ocultar", style="DarkLink.TButton", command=self._toggle_log
        )
        self._toggle_button.pack(side="right")

        self.console = Console(self._dock, height=5)
        self.console.pack(fill="x", padx=20, pady=(2, 14))

    def _build_status_bar(self) -> None:
        barra = tk.Frame(self, bg="#EDF1F5")
        barra.pack(side="bottom", fill="x")
        tk.Frame(barra, bg=theme.BORDER, height=1).pack(fill="x", side="top")

        self._status_var = tk.StringVar(value="Listo")
        tk.Label(
            barra,
            textvariable=self._status_var,
            bg="#EDF1F5",
            fg=theme.TEXT_MUTED,
            font=theme.F_SMALL,
            anchor="w",
        ).pack(side="left", padx=(14, 10), pady=6)

        self.progress = ttk.Progressbar(
            barra, mode="indeterminate", length=110, style="App.Horizontal.TProgressbar"
        )

        tk.Label(
            barra,
            text=f"Sesión local · v{self.VERSION}",
            bg="#EDF1F5",
            fg=theme.TEXT_FAINT,
            font=theme.F_SMALL,
        ).pack(side="right", padx=14, pady=6)

    def _bind_shortcuts(self) -> None:
        for indice, clave in enumerate(DEFINICIONES, start=1):
            self.bind_all(f"<Control-Key-{indice}>", lambda _e, c=clave: self.go(c))

    # ------------------------------------------------------------------
    # Navegación
    # ------------------------------------------------------------------
    def go(self, clave: str) -> None:
        """Muestra la página indicada."""
        definicion = DEFINICIONES[clave]

        if clave not in self._pages:
            modulo = importlib.import_module(definicion["modulo"])
            clase = getattr(modulo, definicion["clase"])
            self._pages[clave] = clase(self._host, self)

        for pagina in self._pages.values():
            pagina.pack_forget()
        self._pages[clave].pack(fill="both", expand=True)

        self._title_var.set(definicion["titulo"])
        self._subtitle_var.set(definicion["subtitulo"])
        for otra, item in self._nav.items():
            item.set_activo(otra == clave)
        self._active = clave

    # ------------------------------------------------------------------
    # Registro y estado
    # ------------------------------------------------------------------
    def log(self, texto: str, level: str = "info") -> None:
        """Escribe una línea en el registro de actividad."""
        if not texto:
            return
        self.console.write(texto.rstrip("\n") + "\n", level)

    def set_status(self, texto: str) -> None:
        self._status_var.set(texto)

    def _clear_log(self) -> None:
        self.console.clear()

    def _toggle_log(self) -> None:
        if self._log_visible:
            self.console.pack_forget()
            self._toggle_button.configure(text="Mostrar")
        else:
            self.console.pack(fill="x", padx=20, pady=(2, 14))
            self._toggle_button.configure(text="Ocultar")
        self._log_visible = not self._log_visible

    # ------------------------------------------------------------------
    # Diálogos
    # ------------------------------------------------------------------
    def inform(self, mensaje: str, titulo: str = "Aviso", tipo: str = "info") -> None:
        """Muestra un mensaje al usuario (sin símbolos de consola)."""
        texto = limpiar_mensaje(mensaje)
        if tipo == "error":
            messagebox.showerror(titulo, texto, parent=self)
        elif tipo == "warn":
            messagebox.showwarning(titulo, texto, parent=self)
        else:
            messagebox.showinfo(titulo, texto, parent=self)

    def confirmar(self, mensaje: str, titulo: str = "Confirmar", peligro: bool = False) -> bool:
        """Pide confirmación explícita antes de una acción con consecuencias."""
        texto = limpiar_mensaje(mensaje)
        if peligro:
            return messagebox.askyesno(titulo, texto, parent=self, icon="warning", default="no")
        return messagebox.askyesno(titulo, texto, parent=self, default="no")

    # ------------------------------------------------------------------
    # Ejecución de tareas
    # ------------------------------------------------------------------
    def run_task(self, mensaje: str, trabajo, al_terminar=None, al_fallar=None) -> bool:
        """Ejecuta `trabajo` en segundo plano mostrando progreso y registro."""
        if self._busy:
            self.log("Hay una operación en curso. Espere a que termine.", "warn")
            self.set_status("Operación en curso…")
            return False

        self._busy = True
        self._on_success = al_terminar
        self._on_error = al_fallar

        self.set_status(mensaje)
        self.progress.pack(side="right", padx=(0, 12), pady=9)
        self.progress.start(12)
        self.configure(cursor="watch")
        self.log(f"— {mensaje}", "title")

        hilo = threading.Thread(target=self._worker, args=(trabajo,), daemon=True)
        hilo.start()
        self.after(60, self._poll)
        return True

    def _worker(self, trabajo) -> None:
        stream = _QueueStream(self._queue)
        with contextlib.redirect_stdout(stream), contextlib.redirect_stderr(stream):
            try:
                resultado = trabajo()
            except BaseException as exc:  # noqa: BLE001 - se informa en la interfaz
                stream.flush()
                self._queue.put(("error", exc, traceback.format_exc()))
            else:
                stream.flush()
                self._queue.put(("ok", resultado))

    def _poll(self) -> None:
        final: tuple | None = None

        while True:
            try:
                item = self._queue.get_nowait()
            except queue.Empty:
                break
            if item[0] == "texto":
                self.log(item[1], _clasificar(item[1]))
            else:
                final = item

        if final is None:
            self.after(60, self._poll)
            return

        self.progress.stop()
        self.progress.pack_forget()
        self.configure(cursor="")
        self._busy = False

        if final[0] == "ok":
            self.set_status("Operación completada")
            callback = self._on_success
            self._on_success = None
            self._on_error = None
            if callback:
                callback(final[1])
        else:
            exc, detalle = final[1], final[2]
            self.set_status("La operación no se pudo completar")
            self.log(detalle.rstrip(), "muted")
            callback = self._on_error
            self._on_success = None
            self._on_error = None
            if callback:
                callback(exc)
            else:
                self.inform(resumen_error(exc), "No se pudo completar la operación", "error")

    # ------------------------------------------------------------------
    # Cierre de la aplicación
    # ------------------------------------------------------------------
    def _on_close(self) -> None:
        if self._busy:
            seguir = messagebox.askyesno(
                "Operación en curso",
                "Hay una operación en ejecución. Si cierra ahora, puede quedar a medias.\n\n"
                "Se recomienda esperar a que termine. ¿Desea cerrar de todos modos?",
                parent=self,
                icon="warning",
                default="no",
            )
            if not seguir:
                return
        self.destroy()
