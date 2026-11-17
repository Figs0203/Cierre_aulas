"""Tema visual de la aplicación (paleta, tipografías y estilos ttk).

La estética es deliberadamente sobria: fondo neutro, superficies blancas con
borde de 1 px, una única familia tipográfica y un solo color de acento. No se
usan degradados, sombras ni iconos decorativos.
"""

from __future__ import annotations

from tkinter import font as tkfont
from tkinter import ttk

# ---------------------------------------------------------------------------
# Paleta
# ---------------------------------------------------------------------------
PAGE = "#F4F6F8"
SURFACE = "#FFFFFF"
SURFACE_ALT = "#F8FAFC"
BORDER = "#DCE3EA"
BORDER_SOFT = "#E9EEF3"
TEXT = "#1B2733"
TEXT_MUTED = "#68757F"
TEXT_FAINT = "#93A0AC"

SIDEBAR = "#152230"
SIDEBAR_HOVER = "#1D2E3E"
SIDEBAR_ACTIVE = "#22374B"
SIDEBAR_TEXT = "#B9C7D4"
SIDEBAR_SECTION = "#6C7F92"

ACCENT = "#1F6FB2"
ACCENT_HOVER = "#1B5F98"
ACCENT_PRESSED = "#164E7E"
ACCENT_SOFT = "#EAF2F9"

SUCCESS = "#2C7A52"
SUCCESS_SOFT = "#E9F5EE"
WARNING = "#9C6100"
WARNING_SOFT = "#FCF2E2"
DANGER = "#B02B23"
DANGER_SOFT = "#FBEDEC"

CONSOLE_BG = "#101B25"
CONSOLE_TEXT = "#C7D4E0"
CONSOLE_MUTED = "#7C8FA1"
CONSOLE_OK = "#7FCB9C"
CONSOLE_WARN = "#E7B36B"
CONSOLE_ERROR = "#EE9086"
CONSOLE_TITLE = "#FFFFFF"

# ---------------------------------------------------------------------------
# Tipografías (se resuelven en `init` según las fuentes instaladas)
# ---------------------------------------------------------------------------
F_H1 = ("Segoe UI", 18, "bold")
F_H2 = ("Segoe UI", 11, "bold")
F_BODY = ("Segoe UI", 10)
F_SMALL = ("Segoe UI", 9)
F_STEP = ("Segoe UI", 8, "bold")
F_NAV = ("Segoe UI", 10)
F_VALUE = ("Segoe UI", 17, "bold")
F_MONO = ("Consolas", 9)


def _pick(root, candidatas: tuple[str, ...], respaldo: str) -> str:
    """Devuelve la primera fuente disponible del listado."""
    try:
        instaladas = set(tkfont.families(root))
    except Exception:  # noqa: BLE001 - nunca impide arrancar
        instaladas = set()
    for nombre in candidatas:
        if nombre in instaladas:
            return nombre
    return respaldo


def init(root) -> ttk.Style:
    """Aplica tipografías y estilos a la ventana indicada."""
    global F_H1, F_H2, F_BODY, F_SMALL, F_STEP, F_NAV, F_VALUE, F_MONO

    familia = _pick(root, ("Segoe UI", "Inter", "Helvetica Neue", "DejaVu Sans", "Arial"), "TkDefaultFont")
    mono = _pick(root, ("Cascadia Mono", "Consolas", "DejaVu Sans Mono", "Courier New"), "TkFixedFont")

    F_H1 = (familia, 18, "bold")
    F_H2 = (familia, 11, "bold")
    F_BODY = (familia, 10)
    F_SMALL = (familia, 9)
    F_STEP = (familia, 8, "bold")
    F_NAV = (familia, 10)
    F_VALUE = (familia, 17, "bold")
    F_MONO = (mono, 9)

    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except Exception:  # noqa: BLE001 - tema no disponible
        pass

    style.configure(
        ".",
        background=PAGE,
        foreground=TEXT,
        font=F_BODY,
        borderwidth=0,
        focuscolor=PAGE,
    )

    # --- Botones ---------------------------------------------------------
    style.configure(
        "Primary.TButton",
        background=ACCENT,
        foreground="#FFFFFF",
        font=(familia, 10, "bold"),
        padding=(18, 9),
        borderwidth=0,
        relief="flat",
    )
    style.map(
        "Primary.TButton",
        background=[
            ("disabled", "#A9C3D8"),
            ("pressed", ACCENT_PRESSED),
            ("active", ACCENT_HOVER),
        ],
        foreground=[("disabled", "#F1F5F9")],
    )

    style.configure(
        "Secondary.TButton",
        background=SURFACE_ALT,
        foreground=TEXT,
        font=F_BODY,
        padding=(16, 9),
        bordercolor="#C5D1DC",
        lightcolor="#C5D1DC",
        darkcolor="#C5D1DC",
        borderwidth=1,
        relief="solid",
    )
    style.map(
        "Secondary.TButton",
        background=[("disabled", "#F2F5F8"), ("pressed", "#E8EEF4"), ("active", "#EFF4F9")],
        foreground=[("disabled", TEXT_FAINT)],
        bordercolor=[("active", "#A9BAC9")],
        lightcolor=[("active", "#A9BAC9")],
        darkcolor=[("active", "#A9BAC9")],
    )

    style.configure(
        "Danger.TButton",
        background=DANGER,
        foreground="#FFFFFF",
        font=(familia, 10, "bold"),
        padding=(18, 9),
        borderwidth=0,
        relief="flat",
    )
    style.map(
        "Danger.TButton",
        background=[("disabled", "#D8B0AC"), ("pressed", "#8F221C"), ("active", "#9C251E")],
    )

    style.configure(
        "Link.TButton",
        background=SURFACE,
        foreground=ACCENT,
        font=F_SMALL,
        padding=(6, 4),
        borderwidth=0,
        relief="flat",
    )
    style.map(
        "Link.TButton",
        background=[("active", ACCENT_SOFT), ("pressed", ACCENT_SOFT)],
        foreground=[("disabled", TEXT_FAINT)],
    )

    style.configure(
        "DarkLink.TButton",
        background=SURFACE,
        foreground=TEXT_MUTED,
        font=F_SMALL,
        padding=(8, 4),
        borderwidth=0,
        relief="flat",
    )
    style.map(
        "DarkLink.TButton",
        background=[("active", BORDER_SOFT), ("pressed", BORDER_SOFT)],
        foreground=[("active", TEXT)],
    )

    # --- Campos de texto -------------------------------------------------
    style.configure(
        "App.TEntry",
        fieldbackground=SURFACE,
        foreground=TEXT,
        insertcolor=TEXT,
        bordercolor=BORDER,
        lightcolor=BORDER,
        darkcolor=BORDER,
        padding=(9, 7),
        relief="flat",
    )
    style.map(
        "App.TEntry",
        bordercolor=[("focus", ACCENT)],
        lightcolor=[("focus", ACCENT)],
        darkcolor=[("focus", ACCENT)],
    )

    # --- Casillas de verificación ---------------------------------------
    style.configure(
        "Card.TCheckbutton",
        background=SURFACE,
        foreground=TEXT,
        font=F_BODY,
        focuscolor=SURFACE,
    )
    style.map(
        "Card.TCheckbutton",
        background=[("active", SURFACE)],
        foreground=[("active", TEXT)],
        indicatorcolor=[("selected", ACCENT), ("!selected", SURFACE)],
    )

    # --- Barra de progreso y desplazamiento ------------------------------
    style.configure(
        "App.Horizontal.TProgressbar",
        troughcolor=BORDER_SOFT,
        background=ACCENT,
        borderwidth=0,
        thickness=4,
    )
    style.configure(
        "Vertical.TScrollbar",
        background="#CBD5DE",
        troughcolor=PAGE,
        bordercolor=PAGE,
        lightcolor=PAGE,
        darkcolor=PAGE,
        arrowcolor=TEXT_MUTED,
        borderwidth=0,
        arrowsize=11,
    )
    style.map("Vertical.TScrollbar", background=[("active", "#B4C1CD")])

    return style
