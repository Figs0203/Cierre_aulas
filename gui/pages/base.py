"""Base común de las páginas de la aplicación.

Cada página tiene dos zonas: el contenido desplazable y una barra inferior fija
con las acciones principales, de modo que los botones siempre estén a la vista
aunque el contenido sea largo.
"""

from __future__ import annotations

import tkinter as tk

from gui import theme
from gui.widgets import ScrollArea

PAD_X = 26
PAD_Y = 18
GAP = 12


class Page(tk.Frame):
    """Página con contenido desplazable y barra de acciones fija."""

    con_barra_acciones = True

    def __init__(self, master, app):
        super().__init__(master, bg=theme.PAGE)
        self.app = app

        self.barra: tk.Frame | None = None
        if self.con_barra_acciones:
            self.barra = self._crear_barra_acciones()

        self.scroll = ScrollArea(self)
        self.scroll.pack(fill="both", expand=True)

        self.content = tk.Frame(self.scroll.inner, bg=theme.PAGE)
        self.content.pack(fill="both", expand=True, padx=PAD_X, pady=(PAD_Y, PAD_Y))

        self.build(self.content)

        if self.barra is not None:
            self.build_acciones(self.barra)

    # ------------------------------------------------------------------
    def _crear_barra_acciones(self) -> tk.Frame:
        contenedor = tk.Frame(self, bg=theme.SURFACE)
        contenedor.pack(side="bottom", fill="x")
        tk.Frame(contenedor, bg=theme.BORDER, height=1).pack(fill="x")
        barra = tk.Frame(contenedor, bg=theme.SURFACE, padx=PAD_X, pady=11)
        barra.pack(fill="x")
        return barra

    # -- puntos de extensión -------------------------------------------
    def build(self, parent) -> None:  # pragma: no cover - cada página lo implementa
        raise NotImplementedError

    def build_acciones(self, barra: tk.Frame) -> None:
        """Acciones principales de la página (se dibujan en la barra fija)."""

    # -- utilidades -----------------------------------------------------
    def nota_acciones(self, barra: tk.Frame, lineas: tuple[str, ...]) -> None:
        """Texto explicativo a la izquierda de los botones de la barra."""
        bloque = tk.Frame(barra, bg=theme.SURFACE)
        bloque.pack(side="left", fill="x", expand=True)
        for indice, linea in enumerate(lineas):
            tk.Label(
                bloque,
                text=linea,
                bg=theme.SURFACE,
                fg=theme.TEXT_MUTED if indice == 0 else theme.TEXT_FAINT,
                font=theme.F_SMALL,
                anchor="w",
                justify="left",
                wraplength=380,
            ).pack(anchor="w")

    @staticmethod
    def fila_botones(barra: tk.Frame) -> tk.Frame:
        fila = tk.Frame(barra, bg=theme.SURFACE)
        fila.pack(side="right")
        return fila

    def pack_card(self, card, pady=(0, GAP)):
        card.pack(fill="x", pady=pady)
        return card
