"""
Verificación de integridad de archivos para el Asistente de Cierre de Aulas COIN.

Calcula y verifica hashes SHA-256 de los archivos fuente para garantizar
matemáticamente que no fueron alterados durante el procesamiento.

Principio de diseño:
    Los 4 archivos originales son fuentes INMUTABLES.
    El hash se calcula ANTES de abrir los archivos y se verifica DESPUÉS
    de finalizar el procesamiento para certificar que permanecieron intactos.
"""

from __future__ import annotations

import hashlib
from dataclasses import dataclass
from pathlib import Path


# Tamaño del bloque de lectura para archivos grandes (64 KB)
_BLOCK_SIZE = 65536


@dataclass(frozen=True)
class FileIntegrityRecord:
    """Registro de integridad de un archivo individual.

    Attributes:
        file_path: Ruta absoluta del archivo.
        file_name: Nombre del archivo (sin ruta, para reportes seguros).
        file_size: Tamaño en bytes.
        sha256_hash: Hash SHA-256 hexadecimal completo.
    """
    file_path: str
    file_name: str
    file_size: int
    sha256_hash: str

    def matches(self, other_hash: str) -> bool:
        """Compara este hash con otro hash hexadecimal."""
        return self.sha256_hash.lower() == other_hash.lower()


def compute_sha256(file_path: str | Path) -> FileIntegrityRecord:
    """Calcula el hash SHA-256 de un archivo.

    Lee el archivo en bloques para manejar archivos grandes sin
    consumir memoria excesiva.

    Args:
        file_path: Ruta al archivo (str o Path).

    Returns:
        FileIntegrityRecord con el hash calculado.

    Raises:
        FileNotFoundError: Si el archivo no existe.
        PermissionError: Si no hay permisos de lectura.
        OSError: Si ocurre un error de I/O.
    """
    path = Path(file_path)

    if not path.exists():
        raise FileNotFoundError(f"Archivo no encontrado: {path.name}")

    if not path.is_file():
        raise ValueError(f"La ruta no corresponde a un archivo: {path.name}")

    hasher = hashlib.sha256()
    file_size = path.stat().st_size

    try:
        with open(path, "rb") as f:
            while True:
                block = f.read(_BLOCK_SIZE)
                if not block:
                    break
                hasher.update(block)
    except PermissionError as e:
        raise PermissionError(
            f"Permiso denegado en '{path.name}'. El archivo está actualmente abierto en Microsoft Excel "
            f"o siendo sincronizado por OneDrive. Por favor ciérrelo en Excel e intente nuevamente."
        ) from e

    return FileIntegrityRecord(
        file_path=str(path.resolve()),
        file_name=path.name,
        file_size=file_size,
        sha256_hash=hasher.hexdigest(),
    )


def compute_all_hashes(file_paths: dict[str, str | Path]) -> dict[str, FileIntegrityRecord]:
    """Calcula hashes SHA-256 para un conjunto de archivos.

    Args:
        file_paths: Diccionario {rol: ruta} donde rol es el identificador
                    del tipo de archivo (ej. "control_aulas", "sistematizacion").

    Returns:
        Diccionario {rol: FileIntegrityRecord}.

    Raises:
        FileNotFoundError: Si algún archivo no existe.
    """
    records = {}
    for role, path in file_paths.items():
        records[role] = compute_sha256(path)
    return records


def verify_integrity(
    original_records: dict[str, FileIntegrityRecord],
    file_paths: dict[str, str | Path],
) -> tuple[bool, list[str]]:
    """Verifica que los archivos no hayan sido modificados.

    Recalcula los hashes y los compara con los registros originales.

    Args:
        original_records: Registros de integridad calculados antes del procesamiento.
        file_paths: Diccionario {rol: ruta} de los archivos a verificar.

    Returns:
        Tupla (todos_intactos: bool, mensajes: list[str]).
        - todos_intactos es True si todos los hashes coinciden.
        - mensajes contiene un mensaje por cada archivo verificado.
    """
    all_intact = True
    messages = []

    for role, path in file_paths.items():
        if role not in original_records:
            messages.append(
                f"⚠️ Rol '{role}' no tiene registro de integridad previo."
            )
            all_intact = False
            continue

        original = original_records[role]
        current = compute_sha256(path)

        if original.matches(current.sha256_hash):
            messages.append(
                f"✅ {original.file_name}: Integridad verificada (SHA-256 coincide)."
            )
        else:
            messages.append(
                f"🔴 {original.file_name}: ¡INTEGRIDAD COMPROMETIDA! "
                f"El hash SHA-256 ha cambiado. "
                f"Original: {original.sha256_hash[:16]}... "
                f"Actual: {current.sha256_hash[:16]}..."
            )
            all_intact = False

    return all_intact, messages


def format_integrity_report(records: dict[str, FileIntegrityRecord]) -> str:
    """Genera un reporte legible de los registros de integridad.

    Args:
        records: Diccionario {rol: FileIntegrityRecord}.

    Returns:
        String formateado con la tabla de integridad.
    """
    lines = [
        "=" * 72,
        "REGISTRO DE INTEGRIDAD DE ARCHIVOS (SHA-256)",
        "=" * 72,
    ]

    for role, record in records.items():
        size_kb = record.file_size / 1024
        lines.append(f"  Rol:     {role}")
        lines.append(f"  Archivo: {record.file_name}")
        lines.append(f"  Tamaño:  {size_kb:.1f} KB")
        lines.append(f"  SHA-256: {record.sha256_hash}")
        lines.append("-" * 72)

    return "\n".join(lines)
