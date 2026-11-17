"""Página «Reparar archivo»: diagnóstico y copia reparada de un .xlsx dañado.

Usa `diagnose_xlsx` y `repair_xlsx`; la reparación siempre se escribe en una
copia (`<nombre>_reparado.xlsx`) y el original nunca se modifica.
"""

from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import ttk

from gui import theme
from gui.pages.base import GAP, Page
from gui.util import abrir_ruta, limpiar_mensaje, resumen_error
from gui.widgets import Card, FilePicker, KeyValue

from output.xlsx_recovery import diagnose_xlsx, repair_xlsx


class RepararPage(Page):
    def build(self, parent) -> None:
        self.ruta: Path | None = None
        self.informe = None

        card = Card(
            parent,
            "Archivo a revisar",
            "Se acepta .xlsx o .xlsm. El análisis es de solo lectura y nunca modifica el archivo.",
        )
        self.picker = FilePicker(
            card.body,
            "Archivo dañado",
            "Por ejemplo, un libro que Excel marca como dañado o que muestra «CARGA BLOQUEADA».",
            on_change=self._archivo_elegido,
            filetypes=[("Archivos Excel", "*.xlsx *.xlsm"), ("Todos los archivos", "*.*")],
        )
        self.picker.pack(fill="x")
        card.pack(fill="x", pady=(0, GAP))

        self.card_resultado = Card(parent, "Diagnóstico")
        self.resultado_body = self.card_resultado.body

    def build_acciones(self, barra) -> None:
        self.nota_acciones(
            barra,
            (
                "Si se detectan daños, se genera una copia corregida junto al original.",
                "El archivo original permanece intacto.",
            ),
        )
        fila = self.fila_botones(barra)
        self.boton_reparar = ttk.Button(
            fila,
            text="Generar copia reparada",
            style="Primary.TButton",
            command=self._reparar,
            state="disabled",
        )
        self.boton_reparar.pack(side="right")
        ttk.Button(
            fila, text="Analizar archivo", style="Secondary.TButton", command=self._analizar
        ).pack(side="right", padx=(10, 0))

    # ------------------------------------------------------------------
    def _archivo_elegido(self, picker: FilePicker) -> None:
        if picker.path is None:
            return
        if picker.path.suffix.lower() not in (".xlsx", ".xlsm"):
            picker.set_error("Debe ser un archivo .xlsx o .xlsm.")
            self.app.inform(
                f"El archivo «{picker.path.name}» no es un libro moderno de Excel (.xlsx o .xlsm).",
                "Archivo no válido",
                "warn",
            )
            return
        self.ruta = picker.path
        self.boton_reparar.configure(state="disabled")
        self.app.log(f"Archivo cargado: {picker.path.name}", "ok")

    def _analizar(self) -> None:
        if not self.ruta:
            self.app.inform("Seleccione el archivo que desea revisar.", "Sin archivo", "warn")
            return
        ruta = self.ruta

        self.app.run_task(
            "Analizando la estructura del libro…",
            lambda: diagnose_xlsx(ruta),
            self._render_diagnostico,
            lambda exc: self.app.inform(resumen_error(exc), "No se pudo analizar", "error"),
        )

    def _render_diagnostico(self, informe) -> None:
        self.informe = informe
        self.card_resultado.pack(fill="x", pady=(0, GAP))
        self.after(60, lambda: self.scroll.scroll_to(self.card_resultado))
        for hijo in self.resultado_body.winfo_children():
            hijo.destroy()
        cuerpo = self.resultado_body

        sano = not informe.hay_danos
        tk.Label(
            cuerpo,
            text="Sin daños estructurales" if sano else "Se detectaron daños",
            bg=theme.SURFACE,
            fg=theme.SUCCESS if sano else theme.WARNING,
            font=theme.F_H2,
            anchor="w",
        ).pack(fill="x", pady=(0, 10))

        KeyValue(cuerpo, "Archivo", informe.archivo).pack(fill="x", pady=(0, 3))
        KeyValue(cuerpo, "Partes del paquete", str(informe.partes)).pack(fill="x", pady=(0, 3))
        KeyValue(
            cuerpo,
            "Fórmulas huérfanas",
            str(informe.esclavas_huerfanas),
        ).pack(fill="x", pady=(0, 3))
        KeyValue(cuerpo, "Celdas duplicadas", str(informe.celdas_duplicadas)).pack(fill="x", pady=(0, 3))
        KeyValue(
            cuerpo, "Estilos fuera de rango", str(informe.estilos_fuera_de_rango)
        ).pack(fill="x", pady=(0, 3))

        if informe.problemas:
            bloque = tk.Frame(cuerpo, bg=theme.WARNING_SOFT, padx=14, pady=12)
            bloque.pack(fill="x", pady=(12, 0))
            tk.Label(
                bloque,
                text="Problemas encontrados",
                bg=theme.WARNING_SOFT,
                fg=theme.WARNING,
                font=theme.F_STEP,
                anchor="w",
            ).pack(fill="x")
            for problema in informe.problemas:
                tk.Label(
                    bloque,
                    text=f"· {limpiar_mensaje(problema)}",
                    bg=theme.WARNING_SOFT,
                    fg=theme.TEXT,
                    font=theme.F_SMALL,
                    anchor="w",
                    justify="left",
                    wraplength=790,
                ).pack(fill="x", pady=(3, 0))
                self.app.log(limpiar_mensaje(problema), "warn")

        if informe.avisos:
            bloque = tk.Frame(cuerpo, bg=theme.SURFACE_ALT, padx=14, pady=12)
            bloque.pack(fill="x", pady=(12, 0))
            tk.Label(
                bloque,
                text="Observaciones",
                bg=theme.SURFACE_ALT,
                fg=theme.TEXT_MUTED,
                font=theme.F_STEP,
                anchor="w",
            ).pack(fill="x")
            for aviso in informe.avisos:
                tk.Label(
                    bloque,
                    text=f"· {limpiar_mensaje(aviso)}",
                    bg=theme.SURFACE_ALT,
                    fg=theme.TEXT,
                    font=theme.F_SMALL,
                    anchor="w",
                    justify="left",
                    wraplength=790,
                ).pack(fill="x", pady=(3, 0))

        if sano:
            tk.Label(
                cuerpo,
                text=(
                    "Si Excel sigue mostrando el aviso, el problema puede ser de sincronización "
                    "(OneDrive/SharePoint) o de un archivo abierto. Copie el libro a una carpeta local "
                    "y vuelva a intentarlo."
                ),
                bg=theme.SURFACE,
                fg=theme.TEXT_MUTED,
                font=theme.F_SMALL,
                anchor="w",
                justify="left",
                wraplength=800,
            ).pack(fill="x", pady=(12, 0))

        self.boton_reparar.configure(state="normal" if informe.hay_danos else "disabled")
        self.app.set_status("Diagnóstico completado")

    def _reparar(self) -> None:
        if not self.ruta:
            return
        ruta = self.ruta

        self.app.run_task(
            "Generando la copia reparada…",
            lambda: repair_xlsx(ruta),
            self._render_reparacion,
            lambda exc: self.app.inform(resumen_error(exc), "No se pudo reparar", "error"),
        )

    def _render_reparacion(self, resultado) -> None:
        self.card_resultado.pack(fill="x", pady=(0, GAP))
        for hijo in self.resultado_body.winfo_children():
            hijo.destroy()
        cuerpo = self.resultado_body

        exito = bool(resultado.reparado and resultado.salida)
        tk.Label(
            cuerpo,
            text="Copia reparada generada" if exito else "No se pudo completar la reparación",
            bg=theme.SURFACE,
            fg=theme.SUCCESS if exito else theme.DANGER,
            font=theme.F_H2,
            anchor="w",
        ).pack(fill="x", pady=(0, 10))

        for linea in (
            ("Grupos de fórmulas reparados", resultado.formulas_compartidas_deshechas),
            ("Fórmulas convertidas a su valor", resultado.esclavas_convertidas),
            ("Celdas duplicadas eliminadas", resultado.celdas_duplicadas_eliminadas),
            ("Estilos corregidos", resultado.estilos_corregidos),
        ):
            KeyValue(cuerpo, linea[0], str(linea[1])).pack(fill="x", pady=(0, 3))

        if exito:
            KeyValue(cuerpo, "Copia generada", str(resultado.salida)).pack(fill="x", pady=(6, 0))
            fila = tk.Frame(cuerpo, bg=theme.SURFACE)
            fila.pack(fill="x", pady=(14, 0))
            ttk.Button(
                fila,
                text="Abrir carpeta",
                style="Secondary.TButton",
                command=lambda: abrir_ruta(Path(resultado.salida).parent),
            ).pack(side="left")
            tk.Label(
                fila,
                text="  Abra la copia en Excel para comprobar que ya no pide reparación.",
                bg=theme.SURFACE,
                fg=theme.TEXT_MUTED,
                font=theme.F_SMALL,
            ).pack(side="left")
            self.app.log(f"Copia reparada: {resultado.salida}", "ok")
        else:
            bloque = tk.Frame(cuerpo, bg=theme.DANGER_SOFT, padx=14, pady=12)
            bloque.pack(fill="x", pady=(12, 0))
            tk.Label(
                bloque,
                text="Motivos",
                bg=theme.DANGER_SOFT,
                fg=theme.DANGER,
                font=theme.F_STEP,
                anchor="w",
            ).pack(fill="x")
            for problema in resultado.problemas_restantes or ["No se pudo generar la copia."]:
                tk.Label(
                    bloque,
                    text=f"· {limpiar_mensaje(problema)}",
                    bg=theme.DANGER_SOFT,
                    fg=theme.TEXT,
                    font=theme.F_SMALL,
                    anchor="w",
                    justify="left",
                    wraplength=790,
                ).pack(fill="x", pady=(3, 0))
                self.app.log(limpiar_mensaje(problema), "error")

        self.app.set_status("Reparación finalizada")
