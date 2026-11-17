"""Página «Restaurar respaldo»: revierte los archivos oficiales al estado previo.

Usa `restore_last_backup` y, antes de restaurar, valida los respaldos con
`diagnose_xlsx` igual que el Modo 6 del menú de consola.
"""

from __future__ import annotations

import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import ttk

from gui import theme
from gui.pages.base import GAP, Page
from gui.util import abrir_ruta, limpiar_mensaje, resumen_error
from gui.widgets import Card, FilePicker, KeyValue

from output.direct_updater import find_latest_backup, restore_last_backup
from output.xlsx_recovery import diagnose_xlsx
from security.validation_files import (
    validate_file_extension,
    validate_file_exists,
    validate_not_auxiliary_file,
)

ROLES: tuple[tuple[str, str, str], ...] = (
    ("sistematizacion", "Sistematización de Cursos COIN", "Libro que recibe las notas y el estado."),
    ("certificados", "Control de Códigos de Certificados", "Libro donde se anexan los certificados."),
    ("control_aulas", "Control de Aulas de Formadores", "Libro donde se marca la fecha de cierre."),
)

ETIQUETAS = {rol: titulo for rol, titulo, _a in ROLES}


class RestaurarPage(Page):
    def build(self, parent) -> None:
        self.pickers: dict[str, FilePicker] = {}
        self.files: dict[str, Path | None] = {rol: None for rol, _t, _a in ROLES}

        aviso = Card(parent, "Antes de continuar")
        tk.Label(
            aviso.body,
            text=(
                "La restauración devuelve cada libro al estado que tenía justo antes del último cierre. "
                "Cualquier cambio hecho después de ese respaldo se pierde. El archivo de Notas nunca se "
                "restaura porque el cierre no lo modifica."
            ),
            bg=theme.SURFACE,
            fg=theme.TEXT_MUTED,
            font=theme.F_SMALL,
            anchor="w",
            justify="left",
            wraplength=800,
        ).pack(fill="x")
        aviso.pack(fill="x", pady=(0, GAP))

        card = Card(
            parent,
            "Archivos oficiales",
            "Seleccione los libros cuyo respaldo desea restaurar. Cada uno se busca en "
            "_backups_cierre junto al archivo.",
        )
        for indice, (rol, titulo, ayuda) in enumerate(ROLES):
            if indice:
                tk.Frame(card.body, bg=theme.BORDER_SOFT, height=1).pack(fill="x", pady=(12, 12))
            picker = FilePicker(card.body, titulo, ayuda, on_change=self._archivo_elegido)
            picker.pack(fill="x")
            self.pickers[rol] = picker
        card.pack(fill="x", pady=(0, GAP))

        self.card_resultado = Card(parent, "Resultado de la restauración")
        self.resultado_body = self.card_resultado.body

    def build_acciones(self, barra) -> None:
        self.nota_acciones(
            barra,
            (
                "Se verifica cada respaldo antes de restaurar y se pide confirmación.",
                "La restauración reemplaza el contenido actual de los archivos oficiales.",
            ),
        )
        fila = self.fila_botones(barra)
        ttk.Button(
            fila,
            text="Restaurar respaldo",
            style="Danger.TButton",
            command=self._restaurar,
        ).pack(side="right")

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
        respaldo = find_latest_backup(picker.path)
        if respaldo is None:
            picker.set_note("No se encontró respaldo en _backups_cierre.", theme.WARNING)
        else:
            fecha = datetime.fromtimestamp(respaldo.stat().st_mtime).strftime("%d/%m/%Y %H:%M")
            picker.set_note(f"Respaldo disponible: {respaldo.name} · {fecha}", theme.SUCCESS)
        self.app.log(f"Archivo cargado: {picker.path.name}", "ok")

    def _restaurar(self) -> None:
        seleccionados = {rol: ruta for rol, ruta in self.files.items() if ruta}
        if not seleccionados:
            self.app.inform(
                "Seleccione al menos un archivo oficial para restaurar.", "Sin archivos", "warn"
            )
            return

        def trabajo():
            danados: list[str] = []
            for _rol, ruta in seleccionados.items():
                respaldo = find_latest_backup(ruta)
                if respaldo is None:
                    continue
                try:
                    informe = diagnose_xlsx(respaldo)
                except Exception:  # noqa: BLE001 - nunca bloquea la restauración
                    continue
                if informe.tiene_danos_graves:
                    danados.append(respaldo.name)
            return danados

        self.app.run_task(
            "Verificando los respaldos…",
            trabajo,
            lambda danados: self._confirmar_restauracion(seleccionados, danados),
            lambda exc: self.app.inform(resumen_error(exc), "No se pudo verificar", "error"),
        )

    def _confirmar_restauracion(self, seleccionados: dict[str, Path], danados: list[str]) -> None:
        if danados:
            texto = (
                "El respaldo de estos archivos también presenta daños:\n\n• "
                + "\n• ".join(danados)
                + "\n\nRestaurarlo no reparará el libro. Puede deberse a que un cierre anterior ya "
                "había dejado el archivo dañado.\n\n¿Desea restaurar de todos modos?"
            )
            if not self.app.confirmar(texto, "Respaldo con daños", peligro=True):
                self.app.log("Restauración cancelada: el respaldo presenta daños.", "warn")
                return

        nombres = "\n".join(f"• {Path(ruta).name}" for ruta in seleccionados.values())
        mensaje = (
            f"Se reemplazará el contenido actual de:\n\n{nombres}\n\n"
            "con el respaldo previo al último cierre. Esta acción no se puede deshacer.\n\n"
            "¿Confirma la restauración?"
        )
        if not self.app.confirmar(mensaje, "Confirmar restauración", peligro=True):
            self.app.log("Restauración cancelada por el usuario.", "warn")
            return

        def restaurar():
            return restore_last_backup(file_paths=seleccionados)

        self.app.run_task(
            "Restaurando los archivos desde el respaldo…",
            restaurar,
            self._render_resultado,
            lambda exc: self.app.inform(resumen_error(exc), "Error al restaurar", "error"),
        )

    def _render_resultado(self, resultados: dict) -> None:
        self.card_resultado.pack(fill="x", pady=(0, GAP))
        self.after(60, lambda: self.scroll.scroll_to(self.card_resultado))
        for hijo in self.resultado_body.winfo_children():
            hijo.destroy()
        cuerpo = self.resultado_body

        todo_ok = all(info.get("status") == "ok" for info in resultados.values())
        tk.Label(
            cuerpo,
            text="Restauración completada" if todo_ok else "Restauración con advertencias",
            bg=theme.SURFACE,
            fg=theme.SUCCESS if todo_ok else theme.WARNING,
            font=theme.F_H2,
            anchor="w",
        ).pack(fill="x", pady=(0, 10))

        for rol, info in resultados.items():
            estado = info.get("status", "error")
            linea = KeyValue(cuerpo, ETIQUETAS.get(rol, rol), limpiar_mensaje(info.get("mensaje", "")))
            linea.pack(fill="x", pady=(0, 4))
            respaldo = info.get("backup")
            if estado == "ok" and respaldo:
                tk.Label(
                    cuerpo,
                    text=f"      Restaurado desde: {Path(respaldo).name}",
                    bg=theme.SURFACE,
                    fg=theme.TEXT_FAINT,
                    font=theme.F_SMALL,
                    anchor="w",
                ).pack(fill="x")
            self.app.log(limpiar_mensaje(info.get("mensaje", "")), "ok" if estado == "ok" else "warn")

        fila = tk.Frame(cuerpo, bg=theme.SURFACE)
        fila.pack(fill="x", pady=(14, 0))
        carpeta = next((Path(r).parent / "_backups_cierre" for r in self.files.values() if r), None)
        if carpeta:
            ttk.Button(
                fila,
                text="Abrir carpeta de respaldos",
                style="Link.TButton",
                command=lambda: abrir_ruta(carpeta),
            ).pack(side="left")

        self.app.set_status("Restauración finalizada")
