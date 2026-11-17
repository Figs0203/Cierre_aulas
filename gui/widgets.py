"""Componentes visuales reutilizables de la aplicación."""

from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import ttk
from typing import Callable

from gui import theme
from gui.util import elide


# ---------------------------------------------------------------------------
# Contenedores
# ---------------------------------------------------------------------------
class Card(tk.Frame):
    """Tarjeta blanca con borde de 1 px y título opcional."""

    def __init__(
        self,
        master,
        title: str = "",
        subtitle: str = "",
        padding: int = 22,
        bg: str = theme.SURFACE,
    ):
        super().__init__(master, bg=theme.BORDER)
        self.bg = bg
        self.inner = tk.Frame(self, bg=bg, padx=padding, pady=padding - 2)
        self.inner.pack(fill="both", expand=True, padx=1, pady=1)

        if title:
            tk.Label(
                self.inner, text=title, bg=bg, fg=theme.TEXT, font=theme.F_H2, anchor="w"
            ).pack(fill="x")
        if subtitle:
            tk.Label(
                self.inner,
                text=subtitle,
                bg=bg,
                fg=theme.TEXT_MUTED,
                font=theme.F_SMALL,
                justify="left",
                anchor="w",
                wraplength=760,
            ).pack(fill="x", pady=(5, 0))

        self.body = tk.Frame(self.inner, bg=bg)
        self.body.pack(fill="both", expand=True, pady=(16, 0) if (title or subtitle) else 0)


class ScrollArea(tk.Frame):
    """Área desplazable vertical con barra que solo aparece cuando hace falta."""

    def __init__(self, master, bg: str = theme.PAGE):
        super().__init__(master, bg=bg)
        self.canvas = tk.Canvas(self, bg=bg, highlightthickness=0, bd=0, takefocus=0)
        self.vbar = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vbar.set)

        self.inner = tk.Frame(self.canvas, bg=bg)
        self._window = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")

        self.canvas.pack(side="left", fill="both", expand=True)
        self.inner.bind("<Configure>", self._on_inner)
        self.canvas.bind("<Configure>", self._on_canvas)
        self.canvas.bind("<Enter>", self._bind_wheel)
        self.canvas.bind("<Leave>", self._unbind_wheel)

    # -- eventos ---------------------------------------------------------
    def _on_inner(self, _event=None):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        self._sync_bar()

    def _on_canvas(self, event):
        self.canvas.itemconfigure(self._window, width=event.width)
        self._sync_bar()

    def _sync_bar(self):
        necesita = self.inner.winfo_reqheight() > self.canvas.winfo_height() + 2
        visible = bool(self.vbar.winfo_ismapped())
        if necesita and not visible:
            self.vbar.pack(side="right", fill="y", before=self.canvas)
        elif not necesita and visible:
            self.vbar.pack_forget()
            self.canvas.yview_moveto(0)

    def _bind_wheel(self, _event=None):
        self.canvas.bind_all("<MouseWheel>", self._on_wheel)
        self.canvas.bind_all("<Button-4>", self._on_wheel_linux)
        self.canvas.bind_all("<Button-5>", self._on_wheel_linux)

    def _unbind_wheel(self, _event=None):
        self.canvas.unbind_all("<MouseWheel>")
        self.canvas.unbind_all("<Button-4>")
        self.canvas.unbind_all("<Button-5>")

    def _on_wheel(self, event):
        primero, ultimo = self.canvas.yview()
        if event.delta > 0 and primero <= 0.0:
            return
        if event.delta < 0 and ultimo >= 1.0:
            return
        self.canvas.yview_scroll(int(-event.delta / 120), "units")

    def _on_wheel_linux(self, event):
        self.canvas.yview_scroll(-1 if event.num == 4 else 1, "units")

    def scroll_to(self, widget) -> None:
        """Desplaza la vista hasta dejar visible el widget indicado."""
        self.update_idletasks()
        alto = max(self.inner.winfo_height(), 1)
        posicion = widget.winfo_rooty() - self.inner.winfo_rooty()
        self.canvas.yview_moveto(max(0.0, min(1.0, posicion / alto)))


# ---------------------------------------------------------------------------
# Texto y datos
# ---------------------------------------------------------------------------
class StatTile(tk.Frame):
    """Cifra destacada con su etiqueta (para resúmenes de operación)."""

    def __init__(self, master, label: str, value: str = "—", bg: str = theme.SURFACE):
        super().__init__(master, bg=bg)
        tk.Label(
            self, text=label.upper(), bg=bg, fg=theme.TEXT_MUTED, font=theme.F_STEP, anchor="w"
        ).pack(fill="x")
        self._value = tk.Label(self, text=value, bg=bg, fg=theme.TEXT, font=theme.F_VALUE, anchor="w")
        self._value.pack(fill="x", pady=(3, 0))

    def set(self, value: str | int) -> None:
        self._value.configure(text=str(value))


class KeyValue(tk.Frame):
    """Línea de detalle: etiqueta fija a la izquierda y valor atenuado."""

    def __init__(self, master, label: str, value: str = "", bg: str = theme.SURFACE):
        super().__init__(master, bg=bg)
        tk.Label(self, text=label, bg=bg, fg=theme.TEXT_MUTED, font=theme.F_SMALL, width=22,
                 anchor="w").pack(side="left")
        self._value = tk.Label(self, text=value, bg=bg, fg=theme.TEXT, font=theme.F_SMALL,
                               anchor="w", justify="left", wraplength=580)
        self._value.pack(side="left", fill="x", expand=True)

    def set(self, value: str) -> None:
        self._value.configure(text=value)


class Pill(tk.Label):
    """Etiqueta compacta de estado."""

    def __init__(self, master, text: str, fg: str = theme.SUCCESS, bg: str = theme.SUCCESS_SOFT):
        super().__init__(master, text=text, fg=fg, bg=bg, font=theme.F_STEP, padx=9, pady=4)

    def update_state(self, text: str, fg: str, bg: str) -> None:
        self.configure(text=text, fg=fg, bg=bg)


def divider(master, pady=(14, 14)) -> tk.Frame:
    """Separador horizontal de 1 px."""
    linea = tk.Frame(master, bg=theme.BORDER_SOFT, height=1)
    linea.pack(fill="x", pady=pady)
    return linea


# ---------------------------------------------------------------------------
# Selector de archivos
# ---------------------------------------------------------------------------
class FilePicker(tk.Frame):
    """Fila para elegir un archivo .xlsx del equipo, con estado y explicación."""

    def __init__(
        self,
        master,
        titulo: str,
        ayuda: str,
        on_change: Callable[["FilePicker"], None] | None = None,
        bg: str = theme.SURFACE,
        filetypes: list[tuple[str, str]] | None = None,
    ):
        super().__init__(master, bg=bg)
        self.on_change = on_change
        self.path: Path | None = None
        self._bg = bg
        self._titulo = titulo
        self._filetypes = filetypes or [
            ("Archivos Excel", "*.xlsx"),
            ("Todos los archivos", "*.*"),
        ]

        self.columnconfigure(1, weight=1)

        self._dot = tk.Canvas(self, width=9, height=9, bg=bg, highlightthickness=0, bd=0)
        self._dot.grid(row=0, column=0, rowspan=2, sticky="w", padx=(0, 10), pady=(6, 0))
        self._dot_id = self._dot.create_oval(0, 0, 8, 8, fill=theme.BORDER, outline="")

        tk.Label(self, text=titulo, bg=bg, fg=theme.TEXT, font=theme.F_H2, anchor="w").grid(
            row=0, column=1, sticky="w"
        )

        self._detail = tk.Label(
            self,
            text=ayuda,
            bg=bg,
            fg=theme.TEXT_MUTED,
            font=theme.F_SMALL,
            anchor="w",
            justify="left",
            wraplength=640,
        )
        self._detail.grid(row=1, column=1, sticky="w", pady=(2, 0))

        self._button = ttk.Button(self, text="Seleccionar", style="Secondary.TButton", command=self.pick)
        self._button.grid(row=0, column=2, rowspan=2, sticky="e", padx=(14, 0))

        self._ayuda = ayuda

    # -- API -------------------------------------------------------------
    def pick(self) -> None:
        from tkinter import filedialog

        seleccion = filedialog.askopenfilename(
            parent=self.winfo_toplevel(),
            title=f"Seleccionar {self._titulo}",
            filetypes=self._filetypes,
        )
        if not seleccion:
            return
        self.set_path(Path(seleccion))

    def set_path(self, ruta: Path | str, nota: str | None = None) -> None:
        self.path = Path(ruta).resolve()
        self._dot.itemconfigure(self._dot_id, fill=theme.SUCCESS)
        self._detail.configure(
            text=nota if nota is not None else elide(str(self.path), 104),
            fg=theme.TEXT_MUTED,
        )
        self._button.configure(text="Cambiar")
        if self.on_change:
            self.on_change(self)

    def set_error(self, mensaje: str) -> None:
        self.path = None
        self._dot.itemconfigure(self._dot_id, fill=theme.DANGER)
        self._detail.configure(text=mensaje, fg=theme.DANGER)
        self._button.configure(text="Seleccionar")

    def set_note(self, nota: str, color: str | None = None) -> None:
        """Muestra una nota bajo el nombre sin cambiar la selección."""
        self._detail.configure(text=nota, fg=color or theme.TEXT_MUTED)

    def clear(self) -> None:
        self.path = None
        self._dot.itemconfigure(self._dot_id, fill=theme.BORDER)
        self._detail.configure(text=self._ayuda, fg=theme.TEXT_MUTED)
        self._button.configure(text="Seleccionar")


class EntryField(tk.Frame):
    """Etiqueta + campo de texto + ayuda opcional."""

    def __init__(
        self,
        master,
        label: str,
        var: tk.StringVar,
        hint: str = "",
        width: int = 22,
        bg: str = theme.SURFACE,
    ):
        super().__init__(master, bg=bg)
        tk.Label(self, text=label, bg=bg, fg=theme.TEXT_MUTED, font=theme.F_SMALL, anchor="w").pack(
            fill="x"
        )
        self.entry = ttk.Entry(self, textvariable=var, style="App.TEntry", width=width, font=theme.F_BODY)
        self.entry.pack(fill="x", pady=(5, 0))
        if hint:
            tk.Label(self, text=hint, bg=bg, fg=theme.TEXT_FAINT, font=theme.F_SMALL,
                     anchor="w", justify="left", wraplength=300).pack(fill="x", pady=(4, 0))


# ---------------------------------------------------------------------------
# Registro de actividad
# ---------------------------------------------------------------------------
class Console(tk.Frame):
    """Panel de registro de actividad (monoespaciado, solo lectura)."""

    MAX_LINEAS = 2000

    def __init__(self, master, height: int = 7):
        super().__init__(master, bg=theme.CONSOLE_BG)
        self.text = tk.Text(
            self,
            height=height,
            bg=theme.CONSOLE_BG,
            fg=theme.CONSOLE_TEXT,
            font=theme.F_MONO,
            bd=0,
            padx=16,
            pady=10,
            wrap="word",
            insertbackground=theme.CONSOLE_TEXT,
            selectbackground="#274055",
            highlightthickness=0,
        )
        self.text.pack(side="left", fill="both", expand=True)

        self.text.tag_configure("info", foreground=theme.CONSOLE_TEXT)
        self.text.tag_configure("muted", foreground=theme.CONSOLE_MUTED)
        self.text.tag_configure("ok", foreground=theme.CONSOLE_OK)
        self.text.tag_configure("warn", foreground=theme.CONSOLE_WARN)
        self.text.tag_configure("error", foreground=theme.CONSOLE_ERROR)
        self.text.tag_configure("title", foreground=theme.CONSOLE_TITLE)
        self.text.configure(state="disabled")

    def write(self, texto: str, level: str = "info") -> None:
        if not texto:
            return
        self.text.configure(state="normal")
        self.text.insert("end", texto, level)
        if int(self.text.index("end-1c").split(".")[0]) > self.MAX_LINEAS:
            self.text.delete("1.0", f"{self.MAX_LINEAS // 3}.0")
        self.text.see("end")
        self.text.configure(state="disabled")

    def clear(self) -> None:
        self.text.configure(state="normal")
        self.text.delete("1.0", "end")
        self.text.configure(state="disabled")
