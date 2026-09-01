"""
Validación de identidad y estructura de archivos de entrada.

Protege contra errores humanos de selección:
- Mismo archivo seleccionado para dos roles diferentes.
- Archivo auxiliar generado previamente seleccionado como fuente.
- Estructura interna incompatible con el rol asignado (validación estructural profunda).
- Archivo ilegible o corrupto.

Principio de diseño:
    Nunca confiar únicamente en el nombre del archivo.
    Validar tanto la identidad (hash, ruta) como la estructura (hojas, columnas).
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List, Set

from normalization.text import normalize_for_comparison
from security.integrity import compute_sha256


# Patrones de encabezados mínimos esperados por cada rol estructural
ROLE_EXPECTED_HEADER_PATTERNS: Dict[str, Set[str]] = {
    "control_aulas": {
        "tipo de aula", "docente principal", "certificados",
        "programa academico", "curso que realiza", "clase",
    },
    "sistematizacion": {
        "aula virtual solicitud", "nombres", "apellidos",
        "numero de documento de identidad", "correo electronico",
        "estado para certificacion", "promedio curso completo",
    },
    "notas": {
        "orgdefinedid", "username", "last name", "first name",
        "calculated final grade", "calculated final", "seccion",
    },
    "certificados": {
        "ciclo", "docente coin", "codigo del certificado",
        "fecha de envio", "ano de certificacion", "total certificados",
    },
}


def validate_no_duplicate_files(file_paths: dict[str, str | Path]) -> list[str]:
    """Verifica que no se haya seleccionado el mismo archivo para roles diferentes.

    Compara tanto las rutas resueltas como los hashes SHA-256.
    """
    errors = []

    resolved_paths: dict[str, str] = {}
    for role, path in file_paths.items():
        resolved = str(Path(path).resolve())
        for existing_role, existing_path in resolved_paths.items():
            if resolved == existing_path:
                errors.append(
                    f"🔴 ERROR: El archivo '{Path(path).name}' está seleccionado "
                    f"para los roles '{existing_role}' y '{role}'. "
                    f"Cada rol debe utilizar un archivo diferente."
                )
        resolved_paths[role] = resolved

    if not errors:
        hashes: dict[str, str] = {}
        for role, path in file_paths.items():
            try:
                record = compute_sha256(path)
                for existing_role, existing_hash in hashes.items():
                    if record.sha256_hash == existing_hash:
                        errors.append(
                            f"🔴 ERROR: Los archivos para '{existing_role}' "
                            f"y '{role}' son idénticos (mismo hash SHA-256). "
                            f"Asegúrese de seleccionar archivos distintos para cada rol."
                        )
                hashes[role] = record.sha256_hash
            except (FileNotFoundError, PermissionError, OSError) as e:
                errors.append(f"🔴 ERROR: No se puede verificar '{role}': {e}")

    return errors


def validate_not_auxiliary_file(file_path: str | Path) -> list[str]:
    """Verifica que el archivo seleccionado no sea un archivo auxiliar generado previamente."""
    errors = []
    name = Path(file_path).name

    if name.startswith("Cierre_Aula_"):
        errors.append(
            f"🔴 ERROR: El archivo '{name}' parece ser un archivo auxiliar "
            f"generado previamente por esta herramienta ('Cierre_Aula_*'). "
            f"No debe usarse como archivo fuente oficial."
        )

    return errors


def validate_file_extension(file_path: str | Path) -> list[str]:
    """Verifica que el archivo tenga extensión .xlsx."""
    errors = []
    path = Path(file_path)

    if path.suffix.lower() not in (".xlsx",):
        errors.append(
            f"🔴 ERROR: El archivo '{path.name}' tiene extensión "
            f"'{path.suffix}'. Se esperan archivos .xlsx de Excel moderno."
        )

    return errors


def validate_file_exists(file_path: str | Path) -> list[str]:
    """Verifica que el archivo existe y es un archivo (no directorio)."""
    errors = []
    path = Path(file_path)

    if not path.exists():
        errors.append(f"🔴 ERROR: Archivo no encontrado: '{path.name}'.")
    elif not path.is_file():
        errors.append(f"🔴 ERROR: La ruta '{path.name}' no es un archivo.")

    return errors


def validate_structural_role(file_path: str | Path, role: str) -> list[str]:
    """Verifica que la estructura interna del archivo corresponda al rol asignado.

    Abre el archivo con openpyxl e inspecciona los encabezados de las primeras filas
    para confirmar que contenga las columnas características del rol esperado.
    """
    errors = []
    path = Path(file_path)

    if not path.exists() or path.suffix.lower() != ".xlsx":
        return errors  # Ya cubierto por otras validaciones

    try:
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    except Exception as e:
        errors.append(f"🔴 ERROR: No se pudo abrir '{path.name}' con openpyxl: {e}")
        return errors

    if not wb.sheetnames:
        errors.append(f"🔴 ERROR: El archivo '{path.name}' no contiene hojas de cálculo.")
        wb.close()
        return errors

    # Leer encabezados de las primeras 5 filas de todas las hojas
    found_headers: Set[str] = set()
    for sname in wb.sheetnames[:3]:
        ws = wb[sname]
        for r in range(1, min(6, ws.max_row + 1 if ws.max_row else 6)):
            for c in range(1, min(40, ws.max_column + 1 if ws.max_column else 40)):
                val = ws.cell(r, c).value
                if val and isinstance(val, str):
                    found_headers.add(normalize_for_comparison(val))

    wb.close()

    expected_patterns = ROLE_EXPECTED_HEADER_PATTERNS.get(role, set())
    if expected_patterns:
        # Calcular cuántos patrones coinciden
        matches = sum(1 for exp in expected_patterns if any(exp in h or h in exp for h in found_headers))
        if matches == 0:
            errors.append(
                f"⚠️ ADVERTENCIA ESTRUCTURAL: El archivo '{path.name}' seleccionado para el rol '{role}' "
                f"no parece contener los encabezados esperados de este tipo de archivo. "
                f"Verifique que no haya intercambiado los archivos."
            )

    return errors


def validate_all_input_files(file_paths: dict[str, str | Path]) -> list[str]:
    """Ejecuta todas las validaciones de archivos de entrada (existencia, extensión, no-duplicidad, estructura)."""
    all_messages = []

    for role, path in file_paths.items():
        all_messages.extend(validate_file_exists(path))
        if Path(path).exists():
            all_messages.extend(validate_file_extension(path))
            all_messages.extend(validate_not_auxiliary_file(path))
            all_messages.extend(validate_structural_role(path, role))

    existence_errors = [m for m in all_messages if "no encontrado" in m or "no es un archivo" in m]
    if not existence_errors:
        all_messages.extend(validate_no_duplicate_files(file_paths))

    return all_messages
