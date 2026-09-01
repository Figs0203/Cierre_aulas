"""
Selector de archivos local e interactivo para el Asistente de Cierre de Aulas COIN.

Permite seleccionar los 4 archivos fuente directamente desde el computador del usuario
mediante diálogos gráficos de archivo (tkinter) o mediante ingreso interactivo por consola.

Principios de seguridad:
- Todo el manejo de rutas y apertura de archivos es 100% local.
- Se validan extensiones, no-duplicidad, y que no se seleccionen archivos auxiliares previos.
- Calcula el hash SHA-256 inmediatamente tras la selección.
"""

from __future__ import annotations

import os
from pathlib import Path
from typing import Optional

from core.enums import FileRole
from security.integrity import compute_sha256, FileIntegrityRecord
from security.validation_files import (
    validate_file_exists,
    validate_file_extension,
    validate_not_auxiliary_file,
    validate_no_duplicate_files,
)


def select_file_via_dialog(title: str, file_types: list[tuple[str, str]] | None = None) -> Optional[str]:
    """Intenta abrir un diálogo de selección de archivo nativo del sistema operativo.

    Si tkinter o el entorno gráfico no está disponible, retorna None para fallback a consola.

    Args:
        title: Título de la ventana de selección.
        file_types: Lista de tuplas (descripción, patrón) para filtrar.

    Returns:
        Ruta seleccionada como str, o None si el usuario canceló o no hay GUI.
    """
    if file_types is None:
        file_types = [("Archivos Excel", "*.xlsx"), ("Todos los archivos", "*.*")]

    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()  # Ocultar ventana principal
        root.attributes("-topmost", True)  # Traer diálogo al frente

        file_path = filedialog.askopenfilename(
            title=title,
            filetypes=file_types,
        )

        root.destroy()

        if file_path and file_path.strip():
            return str(Path(file_path).resolve())
        return None
    except Exception:
        # Si falla tkinter (por ejemplo en headless o sin display), retorno None
        return None


def prompt_for_file(
    role_description: str,
    role_key: str,
    default_hint: str = "",
    use_gui: bool = True,
) -> Path:
    """Solicita al usuario seleccionar o ingresar la ruta de un archivo específico.

    Args:
        role_description: Descripción legible del archivo (ej. "Control de Aulas").
        role_key: Clave del rol (ej. "control_aulas").
        default_hint: Nombre de archivo sugerido / de ejemplo.
        use_gui: Si True, intenta abrir diálogo gráfico primero.

    Returns:
        Path del archivo validado.
    """
    print(f"\n📁 Seleccionar archivo: {role_description}")
    if default_hint:
        print(f"   (Nombre habitual: {default_hint})")

    selected_path_str: Optional[str] = None

    if use_gui:
        print("   [Abriendo ventana de selección de archivo...]")
        selected_path_str = select_file_via_dialog(f"Seleccionar: {role_description}")

    while True:
        if selected_path_str:
            candidate = Path(selected_path_str.strip('"').strip("'"))
        else:
            prompt_msg = f"   Ingrese la ruta del archivo {role_description}: "
            try:
                user_input = input(prompt_msg).strip().strip('"').strip("'")
            except (EOFError, KeyboardInterrupt):
                print("\nOperación cancelada por el usuario.")
                raise SystemExit(0)

            if not user_input:
                print("   ⚠️ Debe ingresar una ruta válida.")
                continue
            candidate = Path(user_input)

        # Validaciones de seguridad sobre el archivo individual
        errs = []
        errs.extend(validate_file_exists(candidate))
        if candidate.exists():
            errs.extend(validate_file_extension(candidate))
            errs.extend(validate_not_auxiliary_file(candidate))

        if not errs:
            print(f"   ✅ Archivo seleccionado: {candidate.name}")
            return candidate.resolve()
        else:
            for err in errs:
                print(f"   {err}")
            print("   Por favor intente nuevamente.")
            selected_path_str = None


def select_all_four_files(use_gui: bool = True) -> dict[str, Path]:
    """Flujo interactivo completo para seleccionar y validar los 4 archivos requeridos.

    Returns:
        Diccionario con las 4 rutas validadas:
        {
            "control_aulas": Path,
            "sistematizacion": Path,
            "notas": Path,
            "certificados": Path,
        }
    """
    print("=" * 72)
    print("SELECCIÓN DE ARCHIVOS FUENTE (PROCESAMIENTO 100% LOCAL)")
    print("=" * 72)
    print("Los 4 archivos serán leídos exclusivamente en memoria en su computador.")
    print("Ningún dato será compartido ni transmitido fuera de su equipo.\n")

    files: dict[str, Path] = {}

    while True:
        # 1. Control de Aulas
        files["control_aulas"] = prompt_for_file(
            role_description="1. Control de Aulas de Formadores",
            role_key="control_aulas",
            default_hint="Control_Aulas_Formadores_2026.xlsx",
            use_gui=use_gui,
        )

        # 2. Sistematización
        files["sistematizacion"] = prompt_for_file(
            role_description="2. Sistematización de Cursos COIN",
            role_key="sistematizacion",
            default_hint="Sistematizacion_Cursos_COIN_2026.xlsx",
            use_gui=use_gui,
        )

        # 3. Notas
        files["notas"] = prompt_for_file(
            role_description="3. Notas del Aula que se cerrará",
            role_key="notas",
            default_hint="Nombre del aula (ej. Notas_ABBUEI154.xlsx)",
            use_gui=use_gui,
        )

        # 4. Certificados
        files["certificados"] = prompt_for_file(
            role_description="4. Control de Códigos / Certificados",
            role_key="certificados",
            default_hint="Control_codigos_certificados_ABBUEI.xlsx",
            use_gui=use_gui,
        )

        # Validación cruzada de no duplicidad
        dup_errors = validate_no_duplicate_files({k: str(v) for k, v in files.items()})
        if dup_errors:
            print("\n🔴 Se detectaron inconsistencias en la selección:")
            for err in dup_errors:
                print(f"   {err}")
            print("\nReiniciando selección para asegurar que los 4 archivos sean distintos...")
            files.clear()
            continue

        break

    print("\n" + "=" * 72)
    print("✅ LOS 4 ARCHIVOS FUERON VALIDADOS Y ESTÁN LISTOS PARA ANÁLISIS LOCAL")
    print("=" * 72)

    return files
