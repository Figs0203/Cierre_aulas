"""Página «Cierre por lotes»: varias aulas en una sola operación.

Replica la secuencia del Modo 8 del menú de consola: los archivos compartidos se
eligen una vez y cada aula se aplica como un cierre independiente usando
`run_cierre_pipeline` y `apply_direct_cierre`.
"""

from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, ttk

from gui import theme
from gui.pages.base import GAP, Page
from gui.util import abrir_ruta, inferir_aula, limpiar_mensaje, resumen_error
from gui.widgets import Card, FilePicker, StatTile, divider

from output.direct_updater import (
    apply_direct_cierre,
    cloud_sync_warnings,
    files_not_writable,
)
from processing.cierre_orchestrator import run_cierre_pipeline
from security.validation_files import (
    validate_file_extension,
    validate_file_exists,
    validate_not_auxiliary_file,
)


class FilaAula(tk.Frame):
    """Fila editable de una aula del lote."""

    def __init__(self, master, notas: Path, on_remove, on_certificado):
        super().__init__(master, bg=theme.SURFACE)
        self.notas = notas
        self.certificado: Path | None = None
        self._on_certificado = on_certificado

        self.aula_var = tk.StringVar(value=inferir_aula(notas))
        self.curso_var = tk.StringVar(value="")

        self.columnconfigure(0, weight=1)

        tk.Label(
            self,
            text=notas.name,
            bg=theme.SURFACE,
            fg=theme.TEXT,
            font=theme.F_SMALL,
            anchor="w",
        ).grid(row=0, column=0, sticky="w", padx=(0, 12))

        ttk.Entry(self, textvariable=self.aula_var, style="App.TEntry", width=13).grid(
            row=0, column=1, padx=(0, 10)
        )
        ttk.Entry(self, textvariable=self.curso_var, style="App.TEntry", width=13).grid(
            row=0, column=2, padx=(0, 10)
        )

        self._cert_label = tk.Label(
            self,
            text="Compartido",
            bg=theme.SURFACE,
            fg=theme.TEXT_FAINT,
            font=theme.F_SMALL,
            anchor="w",
            width=22,
        )
        self._cert_label.grid(row=0, column=3, sticky="w", padx=(0, 10))

        self._cert_boton = ttk.Button(
            self,
            text="Elegir",
            style="Link.TButton",
            command=self._elegir_certificado,
        )
        self._cert_boton.grid(row=0, column=4, padx=(0, 8))
        self._cert_boton.grid_remove()

        ttk.Button(
            self,
            text="Quitar",
            style="Link.TButton",
            command=lambda: on_remove(self),
        ).grid(row=0, column=5)

    # ------------------------------------------------------------------
    @property
    def aula(self) -> str:
        return self.aula_var.get().strip().upper()

    @property
    def curso(self) -> str:
        return self.curso_var.get().strip().upper() or self.aula

    def set_compartido(self, compartido: bool, comun: Path | None = None) -> None:
        if compartido:
            self.certificado = None
            self._cert_boton.grid_remove()
            self._cert_label.configure(text="Compartido", fg=theme.TEXT_FAINT)
        else:
            self._cert_boton.grid()
            if self.certificado:
                self._cert_label.configure(text=self.certificado.name, fg=theme.TEXT)
            elif comun:
                self._cert_label.configure(text="Usar el compartido", fg=theme.TEXT_FAINT)
            else:
                self._cert_label.configure(text="Sin archivo", fg=theme.DANGER)

    def _elegir_certificado(self) -> None:
        seleccion = filedialog.askopenfilename(
            parent=self.winfo_toplevel(),
            title="Certificados del aula",
            filetypes=[("Archivos Excel", "*.xlsx"), ("Todos los archivos", "*.*")],
        )
        if not seleccion:
            return
        ruta = Path(seleccion).resolve()
        errores = (
            validate_file_exists(ruta) + validate_file_extension(ruta) + validate_not_auxiliary_file(ruta)
        )
        if errores:
            from tkinter import messagebox

            messagebox.showwarning(
                "Archivo no válido", limpiar_mensaje(errores[0]), parent=self.winfo_toplevel()
            )
            return
        self.certificado = ruta
        self._cert_label.configure(text=ruta.name, fg=theme.TEXT)
        self._on_certificado(self)


class CierreLotePage(Page):
    def build(self, parent) -> None:
        self.control: Path | None = None
        self.sistematizacion: Path | None = None
        self.certificados_comun: Path | None = None
        self.filas: list[FilaAula] = []
        self.items: list[dict] = []
        self.compartir_certificados = tk.BooleanVar(value=True)
        self.nota_var = tk.BooleanVar(value=False)

        self._build_compartidos(parent)
        self._build_aulas(parent)
        self._build_certificados(parent)
        self._build_opciones(parent)

        self.card_resultado = Card(parent, "Resumen del lote")
        self.resultado_body = self.card_resultado.body

        self._refrescar_estado()

    # ------------------------------------------------------------------
    # Construcción
    # ------------------------------------------------------------------
    def _build_compartidos(self, parent) -> None:
        card = Card(
            parent,
            "Archivos compartidos",
            "Se eligen una sola vez y se reutilizan en todas las aulas del lote.",
        )
        self.picker_control = FilePicker(
            card.body,
            "Control de Aulas de Formadores",
            "Clases, docentes y fechas programadas.",
            on_change=lambda p: self._guardar_compartido("control", p),
        )
        self.picker_control.pack(fill="x")
        divider(card.body, pady=(12, 12))
        self.picker_sistematizacion = FilePicker(
            card.body,
            "Sistematización de Cursos COIN",
            "Registro maestro de estudiantes por clase.",
            on_change=lambda p: self._guardar_compartido("sistematizacion", p),
        )
        self.picker_sistematizacion.pack(fill="x")
        card.pack(fill="x", pady=(0, GAP))

    def _build_aulas(self, parent) -> None:
        card = Card(
            parent,
            "Aulas del lote",
            "Añada los archivos de Notas. El código de aula se deduce del nombre y puede corregirse.",
        )
        cabecera = tk.Frame(card.body, bg=theme.SURFACE)
        cabecera.pack(fill="x")
        ttk.Button(
            cabecera,
            text="Añadir archivos de Notas…",
            style="Secondary.TButton",
            command=self._agregar_notas,
        ).pack(side="left")
        self._contador = tk.Label(
            cabecera,
            text="Ninguna aula añadida",
            bg=theme.SURFACE,
            fg=theme.TEXT_MUTED,
            font=theme.F_SMALL,
        )
        self._contador.pack(side="left", padx=(12, 0))

        self._encabezados = tk.Frame(card.body, bg=theme.SURFACE)
        self._encabezados.pack(fill="x", pady=(16, 4))
        for columna, titulo, ancho in (
            (0, "ARCHIVO DE NOTAS", 0),
            (1, "AULA", 13),
            (2, "CURSO", 13),
            (3, "CERTIFICADOS", 22),
            (4, "", 6),
            (5, "", 5),
        ):
            tk.Label(
                self._encabezados,
                text=titulo,
                bg=theme.SURFACE,
                fg=theme.TEXT_FAINT,
                font=theme.F_STEP,
                anchor="w",
                width=ancho,
            ).grid(row=0, column=columna, sticky="w", padx=(0, 10))
        self._encabezados.columnconfigure(0, weight=1)

        tk.Frame(card.body, bg=theme.BORDER_SOFT, height=1).pack(fill="x")

        self._lista = tk.Frame(card.body, bg=theme.SURFACE)
        self._lista.pack(fill="x")

        self._vacio = tk.Label(
            self._lista,
            text="Aún no hay aulas en el lote.",
            bg=theme.SURFACE,
            fg=theme.TEXT_FAINT,
            font=theme.F_SMALL,
            anchor="w",
            pady=12,
        )
        self._vacio.pack(fill="x")

        card.pack(fill="x", pady=(0, GAP))

    def _build_certificados(self, parent) -> None:
        card = Card(parent, "Certificados")
        ttk.Checkbutton(
            card.body,
            text="Todas las aulas usan el mismo archivo de certificados",
            variable=self.compartir_certificados,
            style="Card.TCheckbutton",
            command=self._cambiar_modo_certificados,
        ).pack(anchor="w")

        self.picker_certificados = FilePicker(
            card.body,
            "Control de Códigos de Certificados",
            "Consecutivo y tabla donde se anexan los certificados.",
            on_change=lambda p: self._guardar_compartido("certificados", p),
        )
        self.picker_certificados.pack(fill="x", pady=(12, 0))

        tk.Label(
            card.body,
            text="Si se desactiva, cada aula puede usar su propio archivo (se elige en la lista).",
            bg=theme.SURFACE,
            fg=theme.TEXT_FAINT,
            font=theme.F_SMALL,
            anchor="w",
        ).pack(fill="x", pady=(10, 0))

        card.pack(fill="x", pady=(0, GAP))

    def _build_opciones(self, parent) -> None:
        card = Card(parent, "Opciones")
        ttk.Checkbutton(
            card.body,
            text="Corregir las notas finales ya escritas que no coincidan con Notas",
            variable=self.nota_var,
            style="Card.TCheckbutton",
        ).pack(anchor="w")
        tk.Label(
            card.body,
            text=(
                "Aplica a todas las aulas del lote. Desactivado, el cierre respeta las notas finales "
                "preexistentes en Sistematización."
            ),
            bg=theme.SURFACE,
            fg=theme.TEXT_MUTED,
            font=theme.F_SMALL,
            anchor="w",
            justify="left",
            wraplength=780,
        ).pack(fill="x", pady=(6, 0))
        card.pack(fill="x", pady=(0, GAP))

    def build_acciones(self, barra) -> None:
        self.nota_acciones(
            barra,
            (
                "Analizar lote valida cada aula sin escribir ningún archivo.",
                "Aplicar el lote escribe en los oficiales con respaldo por aula.",
            ),
        )
        fila = self.fila_botones(barra)
        ttk.Button(
            fila, text="Aplicar lote", style="Primary.TButton", command=self._aplicar
        ).pack(side="right")
        ttk.Button(
            fila, text="Analizar lote", style="Secondary.TButton", command=self._analizar
        ).pack(side="right", padx=(10, 0))

    # ------------------------------------------------------------------
    # Datos
    # ------------------------------------------------------------------
    def _guardar_compartido(self, rol: str, picker: FilePicker) -> None:
        ruta = picker.path
        if ruta is None:
            return
        errores = (
            validate_file_exists(ruta) + validate_file_extension(ruta) + validate_not_auxiliary_file(ruta)
        )
        if errores:
            picker.set_error(limpiar_mensaje(errores[0]))
            self.app.inform(" ".join(errores), "Archivo no válido", "warn")
            return

        if rol == "control":
            self.control = ruta
        elif rol == "sistematizacion":
            self.sistematizacion = ruta
        else:
            self.certificados_comun = ruta
            for fila in self.filas:
                if self.compartir_certificados.get():
                    fila.set_compartido(True, ruta)
        self.app.log(f"Archivo cargado: {ruta.name}", "ok")
        self._refrescar_estado()

    def _agregar_notas(self) -> None:
        seleccion = filedialog.askopenfilenames(
            parent=self.winfo_toplevel(),
            title="Seleccionar las Notas de las aulas a cerrar",
            filetypes=[("Archivos Excel", "*.xlsx"), ("Todos los archivos", "*.*")],
        )
        if not seleccion:
            return

        existentes = {fila.notas for fila in self.filas}
        agregadas = 0
        for ruta_texto in seleccion:
            ruta = Path(ruta_texto).resolve()
            errores = (
                validate_file_exists(ruta)
                + validate_file_extension(ruta)
                + validate_not_auxiliary_file(ruta)
            )
            if errores:
                self.app.inform(" ".join(errores), "Archivo no válido", "warn")
                continue
            if ruta in existentes:
                self.app.log(f"Notas repetidas, se omite: {ruta.name}", "warn")
                continue
            self._agregar_fila(ruta)
            existentes.add(ruta)
            agregadas += 1

        self._refrescar_estado()
        if agregadas:
            self.app.log(f"{agregadas} archivo(s) de Notas añadidos al lote.", "ok")

    def _agregar_fila(self, ruta: Path) -> None:
        fila = FilaAula(self._lista, ruta, self._quitar_fila, lambda _f: self._refrescar_estado())
        fila.pack(fill="x", pady=(8, 0))
        fila.set_compartido(
            bool(self.compartir_certificados.get()) or not self.certificados_comun,
            self.certificados_comun,
        )
        self.filas.append(fila)

    def _quitar_fila(self, fila: FilaAula) -> None:
        if fila in self.filas:
            self.filas.remove(fila)
        fila.destroy()
        self.items = []
        self._refrescar_estado()

    def _cambiar_modo_certificados(self) -> None:
        compartido = bool(self.compartir_certificados.get())
        if compartido:
            self.picker_certificados.pack(fill="x", pady=(12, 0))
        else:
            self.picker_certificados.pack_forget()
        for fila in self.filas:
            fila.set_compartido(compartido, self.certificados_comun)

    def _refrescar_estado(self) -> None:
        total = len(self.filas)
        if total:
            self._vacio.pack_forget()
            self._contador.configure(text=f"{total} aula(s) en el lote")
        else:
            self._vacio.pack(fill="x")
            self._contador.configure(text="Ninguna aula añadida")

    # ------------------------------------------------------------------
    # Validaciones
    # ------------------------------------------------------------------
    def _datos(self):
        if not self.control or not self.sistematizacion:
            self.app.inform(
                "Seleccione el Control de Aulas y la Sistematización (se comparten en todo el lote).",
                "Archivos compartidos incompletos",
                "warn",
            )
            return None
        if not self.filas:
            self.app.inform(
                "Añada al menos un archivo de Notas al lote.", "Lote vacío", "warn"
            )
            return None
        if self.compartir_certificados.get() and not self.certificados_comun:
            self.app.inform(
                "Seleccione el archivo de certificados compartido o desactive la casilla para "
                "elegirlo por aula.",
                "Falta el archivo de certificados",
                "warn",
            )
            return None

        aulas_vistas: set[str] = set()
        for fila in self.filas:
            if not fila.aula:
                self.app.inform(
                    f"Indique el código de aula de «{fila.notas.name}».", "Falta un código de aula", "warn"
                )
                return None
            if fila.aula in aulas_vistas:
                self.app.inform(
                    f"El aula «{fila.aula}» está repetida en el lote.", "Aula repetida", "warn"
                )
                return None
            aulas_vistas.add(fila.aula)
            if not self.compartir_certificados.get() and not fila.certificado:
                self.app.inform(
                    f"Elija el archivo de certificados de «{fila.notas.name}».",
                    "Falta el archivo de certificados",
                    "warn",
                )
                return None
        return True

    def _items_previos(self) -> list[dict]:
        items = []
        for fila in self.filas:
            certificado = (
                self.certificados_comun if self.compartir_certificados.get() else fila.certificado
            )
            files = {
                "control_aulas": self.control,
                "sistematizacion": self.sistematizacion,
                "notas": fila.notas,
                "certificados": certificado,
            }
            items.append({"fila": fila, "aula": fila.aula, "curso": fila.curso, "files": files})
        return items

    # ------------------------------------------------------------------
    # Acciones
    # ------------------------------------------------------------------
    def _analizar(self) -> None:
        if not self._datos():
            return
        items = self._items_previos()

        def trabajo():
            problemas = files_not_writable(
                {
                    "control_aulas": self.control,
                    "sistematizacion": self.sistematizacion,
                }
            )
            for item in items:
                if item["files"]["certificados"]:
                    problemas += files_not_writable({"certificados": item["files"]["certificados"]})
            if problemas:
                raise PermissionError("\n".join(dict.fromkeys(problemas)))

            listos: list[dict] = []
            errores: list[str] = []
            for indice, item in enumerate(items, start=1):
                print(f"Analizando {indice}/{len(items)}: aula {item['aula']} …")
                try:
                    res, _salida, _intacto, _msgs = run_cierre_pipeline(
                        file_paths=item["files"],
                        target_aula=item["aula"],
                        codigo_curso=item["curso"],
                        is_simulation=True,
                    )
                except Exception as exc:  # noqa: BLE001 - una aula no bloquea el lote
                    errores.append(f"{item['aula']}: {resumen_error(exc)}")
                    continue
                if res.total_estudiantes == 0:
                    errores.append(f"{item['aula']}: sin estudiantes en Sistematización.")
                    continue
                aprobados = sum(
                    1
                    for cg in res.clases
                    for m in cg.matches
                    if str(m.estudiante_sist.estado_calculado or "").strip() == "Aprobó"
                )
                item["result"] = res
                item["aprobados"] = aprobados
                listos.append(item)

            if not listos:
                detalle = "\n".join(errores) if errores else "Ninguna aula pudo analizarse."
                raise ValueError(detalle)
            return listos, errores

        self.app.run_task(
            "Analizando las aulas del lote…",
            trabajo,
            self._render_analisis,
            lambda exc: self.app.inform(resumen_error(exc), "No se pudo analizar el lote", "error"),
        )

    def _render_analisis(self, resultado) -> None:
        listos, errores = resultado
        self.items = listos

        self.card_resultado.pack(fill="x", pady=(0, GAP))
        self.after(60, lambda: self.scroll.scroll_to(self.card_resultado))
        self._limpiar(self.resultado_body)
        cuerpo = self.resultado_body

        total_certificados = sum(item["aprobados"] for item in listos)
        estadisticas = tk.Frame(cuerpo, bg=theme.SURFACE)
        estadisticas.pack(fill="x", pady=(0, 14))
        for columna, (etiqueta, valor) in enumerate(
            (
                ("Aulas listas", len(listos)),
                ("Estudiantes", sum(i["result"].total_estudiantes for i in listos)),
                ("Con notas", sum(i["result"].total_matched for i in listos)),
                ("Certificados", total_certificados),
            )
        ):
            StatTile(estadisticas, etiqueta, valor).grid(row=0, column=columna, sticky="w", padx=(0, 34))

        tabla = tk.Frame(cuerpo, bg=theme.SURFACE)
        tabla.pack(fill="x")
        encabezados = ("AULA", "CURSO", "CLASES", "ESTUDIANTES", "CON NOTAS", "APROBADOS")
        for columna, titulo in enumerate(encabezados):
            tk.Label(
                tabla,
                text=titulo,
                bg=theme.SURFACE,
                fg=theme.TEXT_FAINT,
                font=theme.F_STEP,
                anchor="w",
                width=13,
            ).grid(row=0, column=columna, sticky="w", padx=(0, 8), pady=(0, 6))

        for indice, item in enumerate(listos, start=1):
            res = item["result"]
            valores = (
                item["aula"],
                item["curso"],
                str(len(res.clases)),
                str(res.total_estudiantes),
                str(res.total_matched),
                str(item["aprobados"]),
            )
            for columna, valor in enumerate(valores):
                tk.Label(
                    tabla,
                    text=valor,
                    bg=theme.SURFACE,
                    fg=theme.TEXT,
                    font=theme.F_SMALL,
                    anchor="w",
                    width=13,
                ).grid(row=indice, column=columna, sticky="w", padx=(0, 8), pady=(0, 3))

        tk.Label(
            cuerpo,
            text=f"Se anexarán {total_certificados} certificados en total.",
            bg=theme.SURFACE,
            fg=theme.TEXT,
            font=theme.F_SMALL,
            anchor="w",
        ).pack(fill="x", pady=(12, 0))

        if errores:
            bloque = tk.Frame(cuerpo, bg=theme.WARNING_SOFT, padx=14, pady=12)
            bloque.pack(fill="x", pady=(12, 0))
            tk.Label(
                bloque,
                text="Aulas que se omitirán",
                bg=theme.WARNING_SOFT,
                fg=theme.WARNING,
                font=theme.F_STEP,
                anchor="w",
            ).pack(fill="x")
            for error in errores:
                tk.Label(
                    bloque,
                    text=f"· {limpiar_mensaje(error)}",
                    bg=theme.WARNING_SOFT,
                    fg=theme.TEXT,
                    font=theme.F_SMALL,
                    anchor="w",
                    justify="left",
                    wraplength=790,
                ).pack(fill="x", pady=(3, 0))
                self.app.log(limpiar_mensaje(error), "warn")

        self.app.set_status(f"Lote analizado · {len(listos)} aula(s) listas")

    def _aplicar(self) -> None:
        if not self.items:
            self.app.inform(
                "Analice el lote antes de aplicarlo.", "Falta el análisis previo", "warn"
            )
            return
        if not self._datos():
            return

        total_certificados = sum(item["aprobados"] for item in self.items)
        mensaje = (
            f"Se aplicará el cierre de {len(self.items)} aula(s) en los archivos oficiales.\n\n"
            f"• Certificados que se anexarán: {total_certificados}.\n"
            "• El respaldo único se actualizará antes de cada aula.\n"
            "• Si un aula falla de forma crítica, el lote se detiene e informa cuáles quedaron cerradas.\n\n"
            "¿Desea continuar?"
        )
        if not self.app.confirmar(mensaje, "Confirmar cierre del lote", peligro=True):
            self.app.log("Lote cancelado por el usuario. No se modificó ningún archivo.", "warn")
            return

        avisos: list[str] = []
        for item in self.items:
            for aviso in cloud_sync_warnings(item["files"]):
                if aviso not in avisos:
                    avisos.append(aviso)
        if avisos:
            texto = (
                "\n".join(limpiar_mensaje(a) for a in avisos)
                + "\n\nConfirme que los archivos están cerrados en Excel (escritorio y web) "
                "y sincronizados antes de continuar."
            )
            if not self.app.confirmar(texto, "Archivos abiertos o sincronizándose", peligro=True):
                self.app.log("Lote cancelado: verifique el cierre de los archivos.", "warn")
                return

        items = list(self.items)
        actualizar = bool(self.nota_var.get())

        def trabajo():
            exitosas: list[str] = []
            novedades: list[tuple[str, str]] = []
            total = 0
            respaldos: dict[str, Path] = {}
            for indice, item in enumerate(items, start=1):
                print(f"Aplicando {indice}/{len(items)}: aula {item['aula']} …")
                try:
                    reporte = apply_direct_cierre(
                        file_paths=item["files"],
                        result=item["result"],
                        codigo_curso=item["curso"],
                        metadatos_control=item["result"].metadatos_control,
                        start_consecutivo=item["result"].start_consecutivo,
                        actualizar_nota_final_existente=actualizar,
                    )
                except Exception as exc:  # noqa: BLE001 - se detiene el lote
                    novedades.append((item["aula"], resumen_error(exc)))
                    print(f"Error crítico en {item['aula']}: {resumen_error(exc)}")
                    print("Se detiene el lote para no seguir sobre archivos inconsistentes.")
                    break

                total += reporte.certificados_rows_appended
                for respaldo in reporte.backups_created:
                    respaldos[str(respaldo)] = respaldo
                for error in reporte.errors:
                    print(f"Error en {item['aula']}: {limpiar_mensaje(error)}")
                for aviso in reporte.warnings:
                    print(f"Advertencia en {item['aula']}: {limpiar_mensaje(aviso)}")

                if reporte.is_successful:
                    print(
                        f"Aula {item['aula']} cerrada. Certificados anexados: "
                        f"{reporte.certificados_rows_appended}"
                    )
                    exitosas.append(item["aula"])
                else:
                    novedades.append((item["aula"], "; ".join(reporte.errors) or "con errores"))
                    if reporte.fallo_critico:
                        print("Fallo crítico: se detiene el lote.")
                        break
            return exitosas, novedades, total, respaldos

        self.app.run_task(
            "Aplicando el cierre del lote…",
            trabajo,
            self._render_lote,
            lambda exc: self.app.inform(resumen_error(exc), "Error durante el lote", "error"),
        )

    def _render_lote(self, resultado) -> None:
        exitosas, novedades, total, respaldos = resultado
        self.card_resultado.pack(fill="x", pady=(0, GAP))
        self.after(60, lambda: self.scroll.scroll_to(self.card_resultado))
        self._limpiar(self.resultado_body)
        cuerpo = self.resultado_body

        orientado_bien = len(novedades) == 0
        tk.Label(
            cuerpo,
            text="Lote cerrado correctamente" if orientado_bien else "Lote finalizado con novedades",
            bg=theme.SURFACE,
            fg=theme.SUCCESS if orientado_bien else theme.WARNING,
            font=theme.F_H2,
            anchor="w",
        ).pack(fill="x")

        estadisticas = tk.Frame(cuerpo, bg=theme.SURFACE)
        estadisticas.pack(fill="x", pady=(12, 12))
        for columna, (etiqueta, valor) in enumerate(
            (
                ("Aulas cerradas", f"{len(exitosas)} de {len(exitosas) + len(novedades)}"),
                ("Certificados anexados", total),
                ("Respaldos actualizados", len(respaldos)),
            )
        ):
            StatTile(estadisticas, etiqueta, valor).grid(row=0, column=columna, sticky="w", padx=(0, 34))

        if exitosas:
            tk.Label(
                cuerpo,
                text="Aulas cerradas: " + ", ".join(exitosas),
                bg=theme.SURFACE,
                fg=theme.TEXT,
                font=theme.F_SMALL,
                anchor="w",
                justify="left",
                wraplength=800,
            ).pack(fill="x")

        if novedades:
            bloque = tk.Frame(cuerpo, bg=theme.WARNING_SOFT, padx=14, pady=12)
            bloque.pack(fill="x", pady=(12, 0))
            tk.Label(
                bloque,
                text="Aulas con novedades",
                bg=theme.WARNING_SOFT,
                fg=theme.WARNING,
                font=theme.F_STEP,
                anchor="w",
            ).pack(fill="x")
            for aula, detalle in novedades:
                tk.Label(
                    bloque,
                    text=f"· {aula}: {limpiar_mensaje(detalle)}",
                    bg=theme.WARNING_SOFT,
                    fg=theme.TEXT,
                    font=theme.F_SMALL,
                    anchor="w",
                    justify="left",
                    wraplength=790,
                ).pack(fill="x", pady=(3, 0))
                self.app.log(f"{aula}: {limpiar_mensaje(detalle)}", "warn")

        fila_final = tk.Frame(cuerpo, bg=theme.SURFACE)
        fila_final.pack(fill="x", pady=(14, 0))
        carpeta = Path(self.sistematizacion).parent / "_backups_cierre" if self.sistematizacion else None
        if carpeta:
            ttk.Button(
                fila_final,
                text="Abrir carpeta de respaldos",
                style="Link.TButton",
                command=lambda: abrir_ruta(carpeta),
            ).pack(side="left")

        self.app.set_status(f"Lote finalizado · {len(exitosas)} aula(s) cerradas")

    def _limpiar(self, contenedor: tk.Frame) -> None:
        for hijo in contenedor.winfo_children():
            hijo.destroy()
