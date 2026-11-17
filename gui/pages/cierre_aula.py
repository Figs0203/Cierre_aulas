"""Página «Cierre de aula»: flujo principal de la herramienta.

Reúne los modos 2, 3 y 5 del menú de consola sobre las mismas funciones:
`run_cierre_pipeline` (simulación y archivo auxiliar) y `apply_direct_cierre`
(escritura en los archivos oficiales, con respaldo y confirmación).
"""

from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import ttk

from gui.pages.base import GAP, Page
from gui import theme
from gui.util import abrir_ruta, limpiar_mensaje, resumen_error
from gui.widgets import Card, EntryField, FilePicker, StatTile, divider

from output.direct_updater import (
    apply_direct_cierre,
    cloud_sync_warnings,
    files_not_writable,
)
from processing.cierre_orchestrator import run_cierre_pipeline
from security.validation_files import (
    validate_file_extension,
    validate_file_exists,
    validate_no_duplicate_files,
    validate_not_auxiliary_file,
)

ARCHIVOS: tuple[tuple[str, str, str], ...] = (
    (
        "control_aulas",
        "Control de Aulas de Formadores",
        "Clases, docentes y fechas programadas de cada aula.",
    ),
    (
        "sistematizacion",
        "Sistematización de Cursos COIN",
        "Registro maestro de estudiantes por clase.",
    ),
    (
        "notas",
        "Notas del aula",
        "Calificaciones por módulo y nota final. Se lee en modo solo lectura.",
    ),
    (
        "certificados",
        "Control de Códigos de Certificados",
        "Consecutivo y tabla donde se anexan los certificados.",
    ),
)


class CierreAulaPage(Page):
    def build(self, parent) -> None:
        self.files: dict[str, Path | None] = {rol: None for rol, _t, _a in ARCHIVOS}
        self.pickers: dict[str, FilePicker] = {}
        self.resultado = None
        self.reporte = None
        self.salida: Path | None = None

        self.aula_var = tk.StringVar()
        self.curso_var = tk.StringVar()
        self.nota_var = tk.BooleanVar(value=False)

        self._build_archivos(parent)
        self._build_datos(parent)
        self._build_opciones(parent)

        self.card_resultado = Card(parent, "Resultado del análisis")
        self.resultado_body = self.card_resultado.body

    # ------------------------------------------------------------------
    # Construcción
    # ------------------------------------------------------------------
    def _build_archivos(self, parent) -> None:
        card = Card(
            parent,
            "Archivos fuente",
            "Seleccione los cuatro libros oficiales del cierre. Se procesan en memoria; "
            "el archivo de Notas nunca se modifica.",
        )
        for indice, (rol, titulo, ayuda) in enumerate(ARCHIVOS):
            if indice:
                divider(card.body, pady=(8, 8))
            picker = FilePicker(card.body, titulo, ayuda, on_change=self._archivo_elegido)
            picker.pack(fill="x")
            self.pickers[rol] = picker
        card.pack(fill="x", pady=(0, GAP))

    def _build_datos(self, parent) -> None:
        card = Card(
            parent,
            "Datos del aula",
            "El código de aula aparece en la hoja de Sistematización; el de curso se usa en los certificados.",
        )
        fila = tk.Frame(card.body, bg=theme.SURFACE)
        fila.pack(fill="x")

        EntryField(
            fila,
            "Código de aula",
            self.aula_var,
            "Ejemplo: ABBUEI227",
            width=20,
        ).grid(row=0, column=0, sticky="nw", padx=(0, 28))

        EntryField(
            fila,
            "Código del curso",
            self.curso_var,
            "Si se deja vacío se usa el código del aula.",
            width=20,
        ).grid(row=0, column=1, sticky="nw")

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
                "Desactivado, el cierre respeta cualquier nota final preexistente en Sistematización. "
                "Activado, solo se corrige esa columna cuando el valor difiere del archivo de Notas."
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
                "Simular valida sin escribir archivos.",
                "Aplicar el cierre escribe en los oficiales con respaldo previo.",
            ),
        )
        fila = self.fila_botones(barra)
        ttk.Button(
            fila, text="Aplicar cierre", style="Primary.TButton", command=self._aplicar
        ).pack(side="right")
        ttk.Button(
            fila, text="Generar auxiliar", style="Secondary.TButton", command=self._generar
        ).pack(side="right", padx=(10, 0))
        ttk.Button(
            fila, text="Simular", style="Secondary.TButton", command=self._simular
        ).pack(side="right", padx=(10, 0))

    # ------------------------------------------------------------------
    # Selección de archivos
    # ------------------------------------------------------------------
    def _archivo_elegido(self, picker: FilePicker) -> None:
        rol = next((r for r, p in self.pickers.items() if p is picker), "")
        if not rol or picker.path is None:
            return

        errores: list[str] = []
        errores += validate_file_exists(picker.path)
        errores += validate_file_extension(picker.path)
        errores += validate_not_auxiliary_file(picker.path)
        if errores:
            picker.set_error(limpiar_mensaje(errores[0]))
            self.app.inform(" ".join(errores), "Archivo no válido", "warn")
            return

        for otro, camino in self.files.items():
            if otro != rol and camino and Path(camino) == picker.path:
                picker.set_error("Ese archivo ya está asignado a otro campo.")
                self.app.inform(
                    f"El archivo «{picker.path.name}» ya fue seleccionado para otro de los cuatro campos. "
                    "Cada campo requiere un archivo distinto.",
                    "Archivo repetido",
                    "warn",
                )
                return

        self.files[rol] = picker.path
        self.app.log(f"Archivo cargado: {picker.path.name}", "ok")

        if rol == "notas" and not self.aula_var.get().strip():
            from gui.util import inferir_aula

            sugerido = inferir_aula(picker.path)
            if sugerido:
                self.aula_var.set(sugerido)
                self.app.log(f"Código de aula deducido del nombre de las Notas: {sugerido}")

    # ------------------------------------------------------------------
    # Datos del formulario
    # ------------------------------------------------------------------
    def _datos(self):
        faltantes = [titulo for rol, titulo, _a in ARCHIVOS if not self.files.get(rol)]
        if faltantes:
            self.app.inform(
                "Falta seleccionar: " + ", ".join(faltantes) + ".",
                "Archivos incompletos",
                "warn",
            )
            return None

        aula = self.aula_var.get().strip().upper()
        if not aula:
            self.app.inform(
                "Indique el código de aula virtual a cerrar (por ejemplo, ABBUEI227).",
                "Falta el código de aula",
                "warn",
            )
            return None

        self.aula_var.set(aula)
        curso = self.curso_var.get().strip().upper() or aula
        self.curso_var.set(curso)

        files = {rol: self.files[rol] for rol, _t, _a in ARCHIVOS}  # type: ignore[dict-item]
        return files, aula, curso

    @staticmethod
    def _validar_duplicados(files: dict[str, Path]) -> None:
        errores = validate_no_duplicate_files({k: str(v) for k, v in files.items()})
        if errores:
            raise ValueError(" ".join(limpiar_mensaje(e) for e in errores))

    # ------------------------------------------------------------------
    # Acciones
    # ------------------------------------------------------------------
    def _simular(self) -> None:
        self._ejecutar_pipeline(simulacion=True)

    def _generar(self) -> None:
        self._ejecutar_pipeline(simulacion=False)

    def _ejecutar_pipeline(self, simulacion: bool) -> None:
        datos = self._datos()
        if not datos:
            return
        files, aula, curso = datos

        def trabajo():
            self._validar_duplicados(files)
            return run_cierre_pipeline(
                file_paths=files,
                target_aula=aula,
                codigo_curso=curso,
                is_simulation=simulacion,
            )

        mensaje = (
            "Validando el aula en memoria…"
            if simulacion
            else "Generando el archivo auxiliar de cierre…"
        )

        def al_terminar(resultado):
            res, salida, intacto, mensajes = resultado
            self.resultado = res
            self.salida = salida
            self._render_resultado(res, salida, intacto, mensajes, simulacion)

        self.app.run_task(
            mensaje,
            trabajo,
            al_terminar,
            lambda exc: self.app.inform(resumen_error(exc), "No se pudo procesar el aula", "error"),
        )

    def _aplicar(self) -> None:
        datos = self._datos()
        if not datos:
            return
        files, aula, curso = datos

        def preflight():
            problemas = files_not_writable(files)
            if problemas:
                raise PermissionError("\n".join(problemas))
            self._validar_duplicados(files)
            return run_cierre_pipeline(
                file_paths=files,
                target_aula=aula,
                codigo_curso=curso,
                is_simulation=True,
            )

        self.app.run_task(
            "Analizando el aula antes de escribir…",
            preflight,
            lambda resultado: self._confirmar_aplicacion(files, aula, curso, resultado),
            lambda exc: self.app.inform(resumen_error(exc), "No se puede iniciar el cierre", "error"),
        )

    def _confirmar_aplicacion(self, files, aula, curso, resultado) -> None:
        res = resultado[0]

        aprobados = sum(
            1
            for cg in res.clases
            for m in cg.matches
            if str(m.estudiante_sist.estado_calculado or "").strip() == "Aprobó"
        )
        clases = len(res.clases)

        mensaje = (
            f"Se aplicará el cierre del aula {aula} (curso {curso}) en los archivos oficiales.\n\n"
            f"• Sistematización: {res.total_estudiantes} estudiantes, nombres en mayúsculas "
            f"y celdas de cierre completadas.\n"
            f"• Certificados: {aprobados} certificado(s) en {clases} clase(s).\n"
            f"• Control de Aulas: fecha de cierre de las clases detectadas.\n\n"
            "Antes de escribir se actualizará el respaldo único en _backups_cierre.\n"
            "Ninguna celda preexistente con datos se modifica.\n\n"
            "¿Desea continuar?"
        )
        if not self.app.confirmar(mensaje, "Confirmar cierre del aula", peligro=True):
            self.app.log("Cierre cancelado por el usuario. No se modificó ningún archivo.", "warn")
            self.app.set_status("Cierre cancelado")
            return

        avisos = cloud_sync_warnings(files)
        if avisos:
            texto = (
                "\n".join(limpiar_mensaje(a) for a in avisos)
                + "\n\nConfirme que los archivos están cerrados en Excel (escritorio y web) "
                "y sincronizados antes de continuar."
            )
            if not self.app.confirmar(texto, "Archivos abiertos o sincronizándose", peligro=True):
                self.app.log("Cierre cancelado: verifique el cierre de los archivos.", "warn")
                return

        actualizar = bool(self.nota_var.get())

        def aplicar():
            return apply_direct_cierre(
                file_paths=files,
                result=res,
                codigo_curso=curso,
                metadatos_control=res.metadatos_control,
                start_consecutivo=res.start_consecutivo,
                actualizar_nota_final_existente=actualizar,
            )

        self.app.run_task(
            "Aplicando el cierre en los archivos oficiales…",
            aplicar,
            lambda reporte: self._render_informe(reporte, files),
            lambda exc: self.app.inform(resumen_error(exc), "Error durante el cierre", "error"),
        )

    # ------------------------------------------------------------------
    # Presentación de resultados
    # ------------------------------------------------------------------
    def _limpiar(self, contenedor: tk.Frame) -> None:
        for hijo in contenedor.winfo_children():
            hijo.destroy()

    def _montar_resultado(self) -> None:
        if not self.card_resultado.winfo_ismapped():
            self.card_resultado.pack(fill="x", pady=(0, GAP))
        self.after(60, lambda: self.scroll.scroll_to(self.card_resultado))

    def _render_resultado(self, res, salida, intacto, mensajes, simulacion: bool) -> None:
        self._montar_resultado()
        self._limpiar(self.resultado_body)
        cuerpo = self.resultado_body

        aprobados = sum(
            1
            for cg in res.clases
            for m in cg.matches
            if str(m.estudiante_sist.estado_calculado or "").strip() == "Aprobó"
        )

        tk.Label(
            cuerpo,
            text=f"Aula {self.aula_var.get().strip().upper()} · {len(res.clases)} clase(s) detectadas",
            bg=theme.SURFACE,
            fg=theme.TEXT_MUTED,
            font=theme.F_SMALL,
            anchor="w",
        ).pack(fill="x")

        estadisticas = tk.Frame(cuerpo, bg=theme.SURFACE)
        estadisticas.pack(fill="x", pady=(10, 16))
        valores = (
            ("Clases", len(res.clases)),
            ("Estudiantes", res.total_estudiantes),
            ("Con notas", res.total_matched),
            ("Sin notas", res.total_unmatched_sist),
            ("Excluidos", res.total_excluded),
            ("Aprobados", aprobados),
        )
        for columna, (etiqueta, valor) in enumerate(valores):
            tile = StatTile(estadisticas, etiqueta, valor)
            tile.grid(row=0, column=columna, sticky="w", padx=(0, 34))

        # Desglose por clase
        for cg in res.clases:
            aprob = sum(1 for m in cg.matches if m.estudiante_sist.estado_calculado == "Aprobó")
            no_aprob = sum(1 for m in cg.matches if m.estudiante_sist.estado_calculado == "No aprobó")
            aband = sum(1 for m in cg.matches if m.estudiante_sist.estado_calculado == "Abandonó")
            tk.Label(
                cuerpo,
                text=(
                    f"Clase {cg.clase_id}  ·  {len(cg.estudiantes_sist)} estudiantes  ·  "
                    f"{len(cg.matches)} con notas  ·  {aprob} aprobados, "
                    f"{no_aprob} no aprobados, {aband} abandonaron"
                ),
                bg=theme.SURFACE,
                fg=theme.TEXT,
                font=theme.F_SMALL,
                anchor="w",
            ).pack(fill="x", pady=(0, 3))

        # Integridad de los archivos fuente
        color = theme.SUCCESS if intacto else theme.DANGER
        tk.Label(
            cuerpo,
            text=(
                "Los 4 archivos oficiales permanecieron intactos (SHA-256 verificado)."
                if intacto
                else "Se detectó un cambio en los hashes de los archivos fuente. Revise el registro."
            ),
            bg=theme.SURFACE,
            fg=color,
            font=theme.F_SMALL,
            anchor="w",
        ).pack(fill="x", pady=(14, 0))
        for mensaje in mensajes:
            self.app.log(limpiar_mensaje(mensaje), "ok" if intacto else "error")

        # Avisos de identidad del cruce
        avisos = [(m, w) for cg in res.clases for m in cg.matches for w in m.warnings]
        if avisos:
            tk.Label(
                cuerpo,
                text=f"Observaciones de identidad: {len(avisos)} (detalle en el registro)",
                bg=theme.SURFACE,
                fg=theme.WARNING,
                font=theme.F_SMALL,
                anchor="w",
            ).pack(fill="x", pady=(6, 0))
            for _m, aviso in avisos[:20]:
                self.app.log(limpiar_mensaje(aviso), "warn")

        self._render_deteccion_notas(res)

        if salida:
            fila = tk.Frame(cuerpo, bg=theme.SURFACE)
            fila.pack(fill="x", pady=(14, 0))
            tk.Label(
                fila,
                text="Archivo generado:",
                bg=theme.SURFACE,
                fg=theme.TEXT_MUTED,
                font=theme.F_SMALL,
            ).pack(side="left")
            tk.Label(
                fila,
                text=Path(salida).name,
                bg=theme.SURFACE,
                fg=theme.TEXT,
                font=theme.F_SMALL,
            ).pack(side="left", padx=(6, 12))
            ttk.Button(
                fila,
                text="Abrir carpeta",
                style="Link.TButton",
                command=lambda p=salida: self._abrir(p.parent),
            ).pack(side="left")
            self.app.log(f"Archivo auxiliar generado: {salida}", "ok")
        elif simulacion:
            tk.Label(
                cuerpo,
                text="Simulación completada. No se escribió ningún archivo.",
                bg=theme.SURFACE,
                fg=theme.TEXT_MUTED,
                font=theme.F_SMALL,
                anchor="w",
            ).pack(fill="x", pady=(14, 0))

        self.app.set_status(f"Aula analizada · {res.total_matched} estudiantes con notas")

    def _render_deteccion_notas(self, res) -> None:
        info = getattr(res, "notas_deteccion", None) or {}
        if not info:
            return

        cuerpo = self.resultado_body
        bloque = tk.Frame(cuerpo, bg=theme.SURFACE_ALT, padx=14, pady=12)
        bloque.pack(fill="x", pady=(16, 0))

        metodo = info.get("metodo_modulos", "encabezados")
        descripcion = {
            "posicional": "inferidos por posición",
            "mixto": "encabezados + posición",
        }.get(metodo, "detectados por encabezado")

        tk.Label(
            bloque,
            text=f"Lectura del archivo de Notas  ·  hoja «{info.get('hoja') or 'desconocida'}»",
            bg=theme.SURFACE_ALT,
            fg=theme.TEXT,
            font=theme.F_SMALL,
            anchor="w",
        ).pack(fill="x")

        modulos = info.get("modulos") or {}
        partes = []
        for numero in sorted(modulos, key=lambda x: int(x)):
            dato = modulos.get(numero)
            if dato:
                partes.append(f"M{numero} {dato.get('letra')}")
        final = info.get("nota_final")
        resumen_modulos = f"módulos {descripcion}: {', '.join(partes) or 'sin detectar'}"
        if final:
            resumen_modulos += f" · nota final {final.get('letra')}"

        tk.Label(
            bloque,
            text=resumen_modulos,
            bg=theme.SURFACE_ALT,
            fg=theme.TEXT_MUTED,
            font=theme.F_SMALL,
            anchor="w",
            justify="left",
            wraplength=780,
        ).pack(fill="x", pady=(4, 0))

        for aviso in info.get("avisos") or []:
            tk.Label(
                bloque,
                text=f"· {aviso}",
                bg=theme.SURFACE_ALT,
                fg=theme.WARNING,
                font=theme.F_SMALL,
                anchor="w",
                justify="left",
                wraplength=780,
            ).pack(fill="x", pady=(3, 0))
            self.app.log(aviso, "warn")

    def _render_informe(self, reporte, files) -> None:
        self._montar_resultado()
        self._limpiar(self.resultado_body)
        cuerpo = self.resultado_body
        self.reporte = reporte

        exito = reporte.is_successful
        fila_estado = tk.Frame(cuerpo, bg=theme.SURFACE)
        fila_estado.pack(fill="x")
        tk.Label(
            fila_estado,
            text="Cierre aplicado" if exito else "Cierre con novedades",
            bg=theme.SURFACE,
            fg=theme.SUCCESS if exito else theme.WARNING,
            font=theme.F_H2,
        ).pack(side="left")
        tk.Label(
            fila_estado,
            text=f"   {reporte.timestamp}",
            bg=theme.SURFACE,
            fg=theme.TEXT_FAINT,
            font=theme.F_SMALL,
        ).pack(side="left")

        estadisticas = tk.Frame(cuerpo, bg=theme.SURFACE)
        estadisticas.pack(fill="x", pady=(12, 14))
        valores = (
            ("Nombres en mayúsculas", reporte.sistematizacion_names_uppercased),
            ("Celdas completadas", reporte.sistematizacion_updated_cells),
            ("Certificados anexados", reporte.certificados_rows_appended),
            ("Clases marcadas", reporte.control_aulas_classes_marked),
        )
        for columna, (etiqueta, valor) in enumerate(valores):
            StatTile(estadisticas, etiqueta, valor).grid(
                row=0, column=columna, sticky="w", padx=(0, 34)
            )

        detalles: list[str] = [
            f"Celdas preexistentes protegidas (sin cambios): {reporte.sistematizacion_untouched_cells}"
        ]
        if reporte.sistematizacion_celdas_protegidas:
            detalles.append(
                f"Celdas no escritas por ser fórmulas compartidas o matriciales: "
                f"{reporte.sistematizacion_celdas_protegidas}"
            )
        if reporte.sistematizacion_promedios_no_copiados:
            detalles.append(
                f"Notas finales no copiadas porque Notas no las trae: "
                f"{reporte.sistematizacion_promedios_no_copiados}"
            )
        if reporte.sistematizacion_promedios_actualizados:
            detalles.append(
                f"Notas finales corregidas: {reporte.sistematizacion_promedios_actualizados}"
            )
        if reporte.certificados_fila_inicial:
            detalles.append(
                f"Certificados en la hoja «{reporte.certificados_hoja}», filas "
                f"{reporte.certificados_fila_inicial}-{reporte.certificados_fila_final} "
                f"(N° {reporte.certificados_n_inicial}-{reporte.certificados_n_final})"
            )
            detalles.append(
                f"Verificación por relectura del archivo: "
                f"{reporte.certificados_rows_verificados} de "
                f"{reporte.certificados_rows_appended} filas confirmadas"
            )
        if reporte.certificados_omitidos_duplicados:
            detalles.append(
                f"Certificados omitidos por ya estar registrados: "
                f"{reporte.certificados_omitidos_duplicados}"
            )
        if reporte.control_aulas_classes_already_marked:
            detalles.append(
                f"Clases que ya tenían fecha de cierre: "
                f"{reporte.control_aulas_classes_already_marked}"
            )
        if reporte.control_aulas_classes_not_found:
            detalles.append(
                "Clases no encontradas en el Control de Aulas: "
                + ", ".join(reporte.control_aulas_classes_not_found)
            )
        if reporte.backups_created:
            detalles.append("Respaldo previo en: " + ", ".join(p.name for p in reporte.backups_created))

        for linea in detalles:
            tk.Label(
                cuerpo,
                text=f"· {linea}",
                bg=theme.SURFACE,
                fg=theme.TEXT,
                font=theme.F_SMALL,
                anchor="w",
                justify="left",
                wraplength=800,
            ).pack(fill="x", pady=(0, 2))
            self.app.log(linea)

        if reporte.validaciones_preservadas:
            tk.Label(
                cuerpo,
                text="Integridad del libro original conservada (validaciones, formato condicional, "
                "tablas dinámicas, gráficos e imágenes):",
                bg=theme.SURFACE,
                fg=theme.TEXT_MUTED,
                font=theme.F_SMALL,
                anchor="w",
                justify="left",
                wraplength=800,
            ).pack(fill="x", pady=(12, 2))
            for resumen in reporte.validaciones_preservadas.values():
                tk.Label(
                    cuerpo,
                    text=f"· {resumen}",
                    bg=theme.SURFACE,
                    fg=theme.TEXT_MUTED,
                    font=theme.F_SMALL,
                    anchor="w",
                    justify="left",
                    wraplength=800,
                ).pack(fill="x")

        if reporte.warnings:
            self._bloque_mensajes(cuerpo, "Advertencias", reporte.warnings, theme.WARNING, theme.WARNING_SOFT)
            for aviso in reporte.warnings:
                self.app.log(limpiar_mensaje(aviso), "warn")
        if reporte.errors:
            self._bloque_mensajes(cuerpo, "Errores", reporte.errors, theme.DANGER, theme.DANGER_SOFT)
            for error in reporte.errors:
                self.app.log(limpiar_mensaje(error), "error")

        fila_final = tk.Frame(cuerpo, bg=theme.SURFACE)
        fila_final.pack(fill="x", pady=(14, 0))
        carpeta_respaldos = Path(files["sistematizacion"]).parent / "_backups_cierre"
        ttk.Button(
            fila_final,
            text="Abrir carpeta de respaldos",
            style="Link.TButton",
            command=lambda: self._abrir(carpeta_respaldos),
        ).pack(side="left")
        ttk.Button(
            fila_final,
            text="Abrir carpeta de los archivos oficiales",
            style="Link.TButton",
            command=lambda: self._abrir(Path(files["sistematizacion"]).parent),
        ).pack(side="left", padx=(10, 0))

        self.app.set_status(
            f"Cierre aplicado · {reporte.certificados_rows_appended} certificados anexados"
            if exito
            else "Cierre finalizado con novedades"
        )

    def _bloque_mensajes(self, padre, titulo, mensajes, color, fondo) -> None:
        bloque = tk.Frame(padre, bg=fondo, padx=14, pady=12)
        bloque.pack(fill="x", pady=(12, 0))
        tk.Label(
            bloque,
            text=titulo,
            bg=fondo,
            fg=color,
            font=theme.F_STEP,
            anchor="w",
        ).pack(fill="x")
        for mensaje in mensajes:
            tk.Label(
                bloque,
                text=f"· {limpiar_mensaje(mensaje)}",
                bg=fondo,
                fg=theme.TEXT,
                font=theme.F_SMALL,
                anchor="w",
                justify="left",
                wraplength=790,
            ).pack(fill="x", pady=(3, 0))

    def _abrir(self, ruta) -> None:
        if not abrir_ruta(ruta):
            self.app.inform(f"No se pudo abrir la carpeta:\n{ruta}", "No se pudo abrir", "warn")
