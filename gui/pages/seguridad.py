"""Página «Seguridad y privacidad»: garantías de la herramienta."""

from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import ttk

from gui import theme
from gui.pages.base import GAP, Page
from gui.util import abrir_ruta, carpeta_reportes, limpiar_mensaje
from gui.widgets import Card

from security.offline_check import check_offline_compliance

GARANTIAS: tuple[tuple[str, str], ...] = (
    (
        "Todo se procesa en este equipo",
        "Los cuatro libros se abren en memoria y no se envían a ningún servidor, modelo o servicio externo. "
        "La aplicación no realiza conexiones de red ni recolecta estadísticas de uso.",
    ),
    (
        "Las fuentes son inmutables",
        "Cada archivo se audita con SHA-256 antes y después de cada operación. Solo la acción explícita de "
        "«Aplicar cierre» escribe en los archivos oficiales, y siempre con respaldo previo.",
    ),
    (
        "Escritura quirúrgica en el libro",
        "El cierre edita únicamente las celdas necesarias y conserva intacto el resto del paquete .xlsx: "
        "listas de validación, formato condicional, fórmulas, tablas dinámicas, gráficos e imágenes.",
    ),
    (
        "Un único respaldo por archivo",
        "Antes de escribir se guarda el estado previo en _backups_cierre. La opción «Restaurar respaldo» "
        "devuelve los libros a ese punto y la opción «Reparar archivo» trabaja siempre sobre una copia.",
    ),
    (
        "Datos personales fuera de los informes",
        "Los diagnósticos enmascaran documentos, correos y nombres, de modo que el informe puede "
        "compartirse sin exponer información de los estudiantes.",
    ),
)


class SeguridadPage(Page):
    con_barra_acciones = False

    def build(self, parent) -> None:
        intro = Card(
            parent,
            "Privacidad por diseño",
            "La herramienta se construyó para trabajar con información académica confidencial sin salir "
            "del computador de la biblioteca.",
        )
        for titulo, texto in GARANTIAS:
            fila = tk.Frame(intro.body, bg=theme.SURFACE)
            fila.pack(fill="x", pady=(0, 12))
            marca = tk.Frame(fila, bg=theme.ACCENT, width=3)
            marca.pack(side="left", fill="y", padx=(0, 12))
            bloque = tk.Frame(fila, bg=theme.SURFACE)
            bloque.pack(side="left", fill="x", expand=True)
            tk.Label(
                bloque,
                text=titulo,
                bg=theme.SURFACE,
                fg=theme.TEXT,
                font=theme.F_H2,
                anchor="w",
            ).pack(fill="x")
            tk.Label(
                bloque,
                text=texto,
                bg=theme.SURFACE,
                fg=theme.TEXT_MUTED,
                font=theme.F_SMALL,
                anchor="w",
                justify="left",
                wraplength=760,
            ).pack(fill="x", pady=(3, 0))
        intro.pack(fill="x", pady=(0, GAP))

        verificacion = Card(
            parent,
            "Verificación de arquitectura offline",
            "Comprueba que el entorno de ejecución solo usa la librería estándar y openpyxl.",
        )
        self._salida = tk.Frame(verificacion.body, bg=theme.SURFACE)
        self._salida.pack(fill="x")
        tk.Label(
            self._salida,
            text="La verificación aún no se ha ejecutado.",
            bg=theme.SURFACE,
            fg=theme.TEXT_FAINT,
            font=theme.F_SMALL,
            anchor="w",
        ).pack(fill="x")

        fila = tk.Frame(verificacion.body, bg=theme.SURFACE)
        fila.pack(fill="x", pady=(14, 0))
        ttk.Button(
            fila,
            text="Ejecutar verificación",
            style="Secondary.TButton",
            command=self._verificar,
        ).pack(side="left")
        ttk.Button(
            fila,
            text="Abrir carpeta de informes",
            style="Link.TButton",
            command=lambda: abrir_ruta(carpeta_reportes()),
        ).pack(side="left", padx=(10, 0))
        verificacion.pack(fill="x", pady=(0, GAP))

        proyecto = Card(parent, "Ubicaciones")
        for etiqueta, ruta in (
            ("Carpeta del proyecto", Path(__file__).resolve().parents[2]),
            ("Informes y diagnósticos", carpeta_reportes()),
        ):
            fila = tk.Frame(proyecto.body, bg=theme.SURFACE)
            fila.pack(fill="x", pady=(0, 6))
            tk.Label(
                fila,
                text=etiqueta,
                bg=theme.SURFACE,
                fg=theme.TEXT_MUTED,
                font=theme.F_SMALL,
                width=26,
                anchor="w",
            ).pack(side="left")
            tk.Label(
                fila,
                text=str(ruta),
                bg=theme.SURFACE,
                fg=theme.TEXT,
                font=theme.F_SMALL,
                anchor="w",
            ).pack(side="left", padx=(0, 10))
            ttk.Button(
                fila,
                text="Abrir",
                style="Link.TButton",
                command=lambda r=ruta: abrir_ruta(r),
            ).pack(side="left")
        proyecto.pack(fill="x", pady=(0, GAP))

    def _verificar(self) -> None:
        for hijo in self._salida.winfo_children():
            hijo.destroy()
        conforme, mensajes = check_offline_compliance()
        for mensaje in mensajes:
            tk.Label(
                self._salida,
                text=limpiar_mensaje(mensaje),
                bg=theme.SURFACE,
                fg=theme.TEXT_MUTED if not conforme else theme.TEXT,
                font=theme.F_SMALL,
                anchor="w",
                justify="left",
                wraplength=780,
            ).pack(fill="x", pady=(0, 2))
            self.app.log(limpiar_mensaje(mensaje), "ok")
        tk.Label(
            self._salida,
            text=(
                "Resultado: arquitectura conforme, sin dependencias de red."
                if conforme
                else "Resultado: se encontraron dependencias no autorizadas."
            ),
            bg=theme.SURFACE,
            fg=theme.SUCCESS if conforme else theme.DANGER,
            font=theme.F_SMALL,
            anchor="w",
        ).pack(fill="x", pady=(8, 0))
        self.app.set_status("Verificación offline completada")
