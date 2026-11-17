"""Página «Diagnóstico»: informes técnicos locales sin datos personales.

Expone los mismos generadores del menú de consola: `run_diagnostico` (cruce de
identidades, Modo 7) y `run_discovery` (inspección técnica, Modo 1).
"""

from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import ttk

from gui import theme
from gui.pages.base import GAP, Page
from gui.util import abrir_ruta, elide, limpiar_mensaje, resumen_error
from gui.widgets import Card, EntryField, FilePicker

from security.validation_files import (
    validate_file_extension,
    validate_file_exists,
    validate_not_auxiliary_file,
)
from tests.discovery.diagnose_matching import run_diagnostico
from tests.discovery.inspect_real_files import run_discovery

ARCHIVOS: tuple[tuple[str, str, str], ...] = (
    ("control_aulas", "Control de Aulas de Formadores", "Estructura y metadatos de las clases."),
    ("sistematizacion", "Sistematización de Cursos COIN", "Encabezados y registros por aula."),
    ("notas", "Notas del aula", "Hojas, columnas de módulos y nota final."),
    ("certificados", "Control de Códigos de Certificados", "Ubicación de la tabla de certificados."),
)


class DiagnosticoPage(Page):
    def build(self, parent) -> None:
        self.files: dict[str, Path | None] = {rol: None for rol, _t, _a in ARCHIVOS}
        self.pickers: dict[str, FilePicker] = {}
        self.aula_var = tk.StringVar()

        card = Card(
            parent,
            "Archivos a analizar",
            "Los informes se guardan en la carpeta reports/ y no incluyen documentos, correos ni "
            "nombres completos.",
        )
        for indice, (rol, titulo, ayuda) in enumerate(ARCHIVOS):
            if indice:
                tk.Frame(card.body, bg=theme.BORDER_SOFT, height=1).pack(fill="x", pady=(12, 12))
            picker = FilePicker(card.body, titulo, ayuda, on_change=self._archivo_elegido)
            picker.pack(fill="x")
            self.pickers[rol] = picker
        card.pack(fill="x", pady=(0, GAP))

        datos = Card(parent, "Aula de referencia")
        EntryField(
            datos.body,
            "Código de aula",
            self.aula_var,
            "Obligatorio para el diagnóstico de cruce; opcional para la inspección técnica.",
            width=20,
        ).pack(anchor="w")
        datos.pack(fill="x", pady=(0, GAP))

        self.card_resultado = Card(parent, "Informe generado")
        self.resultado_body = self.card_resultado.body

    def build_acciones(self, barra) -> None:
        self.nota_acciones(
            barra,
            (
                "El cruce explica por qué los estudiantes no coinciden entre Notas y Sistematización.",
                "La inspección técnica documenta hojas, encabezados, fórmulas y colores.",
            ),
        )
        fila = self.fila_botones(barra)
        ttk.Button(
            fila, text="Diagnóstico de cruce", style="Primary.TButton", command=self._diagnostico
        ).pack(side="right")
        ttk.Button(
            fila, text="Inspección técnica", style="Secondary.TButton", command=self._inspeccion
        ).pack(side="right", padx=(10, 0))

    # ------------------------------------------------------------------
    def _archivo_elegido(self, picker: FilePicker) -> None:
        rol = next((r for r, p in self.pickers.items() if p is picker), "")
        if not rol or picker.path is None:
            return
        errores = (
            validate_file_exists(picker.path)
            + validate_file_extension(picker.path)
            + validate_not_auxiliary_file(picker.path)
        )
        if errores:
            picker.set_error(limpiar_mensaje(errores[0]))
            self.app.inform(" ".join(errores), "Archivo no válido", "warn")
            return
        self.files[rol] = picker.path
        self.app.log(f"Archivo cargado: {picker.path.name}", "ok")

    def _completos(self) -> bool:
        faltantes = [titulo for rol, titulo, _a in ARCHIVOS if not self.files.get(rol)]
        if faltantes:
            self.app.inform(
                "Falta seleccionar: " + ", ".join(faltantes) + ".",
                "Archivos incompletos",
                "warn",
            )
            return False
        return True

    def _diagnostico(self) -> None:
        if not self._completos():
            return
        aula = self.aula_var.get().strip().upper()
        if not aula:
            self.app.inform(
                "Indique el código de aula para el diagnóstico de cruce.",
                "Falta el código de aula",
                "warn",
            )
            return

        files = {rol: ruta for rol, ruta in self.files.items() if ruta}

        self.app.run_task(
            "Generando el diagnóstico de cruce…",
            lambda: run_diagnostico(files, aula),
            lambda ruta: self._mostrar_informe(ruta, "Diagnóstico de cruce"),
            lambda exc: self.app.inform(resumen_error(exc), "No se pudo generar el diagnóstico", "error"),
        )

    def _inspeccion(self) -> None:
        if not self._completos():
            return
        files = {rol: ruta for rol, ruta in self.files.items() if ruta}
        aula = self.aula_var.get().strip().upper() or None

        def trabajo():
            ruta, _texto = run_discovery(files, target_aula=aula)
            return ruta

        self.app.run_task(
            "Inspeccionando la estructura de los archivos…",
            trabajo,
            lambda ruta: self._mostrar_informe(ruta, "Inspección técnica (Fase 0)"),
            lambda exc: self.app.inform(resumen_error(exc), "No se pudo inspeccionar", "error"),
        )

    def _mostrar_informe(self, ruta: Path, titulo: str) -> None:
        self.card_resultado.pack(fill="x", pady=(0, GAP))
        self.after(60, lambda: self.scroll.scroll_to(self.card_resultado))
        for hijo in self.resultado_body.winfo_children():
            hijo.destroy()
        cuerpo = self.resultado_body

        tk.Label(
            cuerpo,
            text=titulo,
            bg=theme.SURFACE,
            fg=theme.TEXT,
            font=theme.F_H2,
            anchor="w",
        ).pack(fill="x")
        tk.Label(
            cuerpo,
            text=elide(str(ruta), 100),
            bg=theme.SURFACE,
            fg=theme.TEXT_MUTED,
            font=theme.F_SMALL,
            anchor="w",
            justify="left",
            wraplength=800,
        ).pack(fill="x", pady=(6, 12))

        fila = tk.Frame(cuerpo, bg=theme.SURFACE)
        fila.pack(fill="x")
        ttk.Button(
            fila,
            text="Abrir informe",
            style="Secondary.TButton",
            command=lambda: self._abrir(ruta),
        ).pack(side="left")
        ttk.Button(
            fila,
            text="Abrir carpeta",
            style="Link.TButton",
            command=lambda: self._abrir(ruta.parent),
        ).pack(side="left", padx=(10, 0))

        tk.Label(
            cuerpo,
            text="El informe no contiene datos personales y puede compartirse.",
            bg=theme.SURFACE,
            fg=theme.TEXT_FAINT,
            font=theme.F_SMALL,
            anchor="w",
        ).pack(fill="x", pady=(12, 0))

        self.app.log(f"Informe generado: {ruta}", "ok")
        self.app.set_status("Informe generado")

    def _abrir(self, ruta: Path) -> None:
        if not abrir_ruta(ruta):
            self.app.inform(f"No se pudo abrir:\n{ruta}", "No se pudo abrir", "warn")
