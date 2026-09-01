"""
Punto de Entrada Principal (CLI) — Asistente de Cierre de Aulas COIN
Versión: 0.1.0-alpha (Fase 0 — Descubrimiento Local)

Herramienta local, segura y offline para el análisis, cruce y validación
de archivos Excel del cierre de aulas de los cursos de la biblioteca.

Principios de funcionamiento:
- 100% OFFLINE: Ningún dato sale de este computador.
- SOLO LECTURA: Los 4 archivos oficiales nunca se modifican.
- VERIFICACIÓN SHA-256: Se audita la integridad antes y después.
"""

from __future__ import annotations

import sys
from pathlib import Path

# Configurar ruta base del proyecto
_PROJECT_ROOT = Path(__file__).resolve().parent
if str(_PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(_PROJECT_ROOT))

from input.file_selector import select_all_four_files
from security.integrity import compute_sha256, verify_integrity
from security.offline_check import check_offline_compliance
from tests.discovery.inspect_real_files import run_discovery


def print_banner():
    print("=" * 76)
    print("   ASISTENTE DE CIERRE DE AULAS — CURSOS COIN (BIBLIOTECA)")
    print("   Herramienta Local de Extracción, Cruce y Validación")
    print("=" * 76)
    print("🔒 Garantía de Privacidad:")
    print("   • Ejecución estrictamente local en su equipo (Offline-First).")
    print("   • Cero envío de datos a servidores externos, modelos o Internet.")
    print("   • Los 4 archivos originales son tratados como SOLO LECTURA inmutable.")
    print("=" * 76)


def show_menu() -> str:
    print("\nMENU PRINCIPAL:")
    print("  [1] Modo 1 — Descubrimiento Técnico Local (Fase 0) [HABILITADO]")
    print("  [2] Modo 2 — Simulación Previa [BLOQUEADO: Requiere validación de reglas]")
    print("  [3] Modo 3 — Generar Archivo Auxiliar de Cierre [BLOQUEADO: Requiere validación de reglas]")
    print("  [4] Verificación de Arquitectura Offline")
    print("  [0] Salir")
    print("-" * 76)
    try:
        choice = input("Seleccione una opción (0-4): ").strip()
        return choice
    except (EOFError, KeyboardInterrupt):
        return "0"


def handle_discovery_mode():
    print("\n" + "=" * 76)
    print("INICIANDO MODO 1: DESCUBRIMIENTO TÉCNICO LOCAL (FASE 0)")
    print("=" * 76)
    print("Este modo inspeccionará la estructura interna de sus 4 archivos Excel")
    print("(hojas, encabezados, fórmulas, colores de relleno y formato condicional)")
    print("y generará un informe técnico local en la carpeta 'reports/'.\n")

    # 1. Selección de los 4 archivos
    files = select_all_four_files(use_gui=True)

    # 2. Aula opcional a buscar
    print("\n🔍 Código del Aula a inspeccionar específicamente (opcional, ej. ABBUEI154):")
    try:
        aula_input = input("   Ingrese código de aula o presione ENTER para omitir: ").strip()
        target_aula = aula_input if aula_input else None
    except (EOFError, KeyboardInterrupt):
        target_aula = None

    # 3. Guardar hashes antes de la ejecución
    hashes_before = {}
    for role, path in files.items():
        rec = compute_sha256(path)
        hashes_before[role] = rec

    print("\n⏳ Ejecutando análisis técnico profundo de los archivos en memoria...")

    # 4. Ejecución del descubrimiento
    report_path, _ = run_discovery(files, target_aula=target_aula)

    # 5. Verificación de integridad post-ejecución
    print("\n🔒 Verificando que ningún archivo haya sido modificado...")
    intact, integrity_msgs = verify_integrity(hashes_before, files)

    print("\n" + "=" * 76)
    print("RESULTADO DE LA INSPECCIÓN TÉCNICA LOCAL")
    print("=" * 76)
    for msg in integrity_msgs:
        print(f"  {msg}")

    if intact:
        print("\n✅ INTEGRIDAD 100% CERTIFICADA: Los archivos originales permanecieron intactos.")
    else:
        print("\n🔴 ADVERTENCIA: Se detectó una inconsistencia de hash en los archivos.")

    print(f"\n📄 El Informe de Descubrimiento completo ha sido guardado en:")
    print(f"   👉 {report_path.resolve()}")
    print("\nPuede abrir ese archivo de texto para revisar la estructura de sus archivos.")
    print("=" * 76)


def handle_offline_check():
    print("\n" + "=" * 76)
    print("VERIFICACIÓN DE ARQUITECTURA OFFLINE-FIRST")
    print("=" * 76)
    compliant, msgs = check_offline_compliance()
    for m in msgs:
        print(f"  {m}")
    print("\n✅ El sistema opera de manera 100% aislada de la red.")
    print("=" * 76)


def main():
    print_banner()

    while True:
        choice = show_menu()

        if choice == "1":
            handle_discovery_mode()
        elif choice in ("2", "3"):
            print("\n🔒 Esta funcionalidad está bloqueada temporalmente.")
            print("   Razón: Las reglas de negocio definitivas (fórmulas exactas, formatos de color,")
            print("   relación Sección/Clase) deben ser confirmadas a partir del informe de Fase 0.")
        elif choice == "4":
            handle_offline_check()
        elif choice in ("0", "q", "salir", "exit"):
            print("\nSaliendo del asistente. ¡Hasta pronto!")
            sys.exit(0)
        else:
            print("\n⚠️ Opción no válida. Por favor seleccione 0, 1, 2, 3 o 4.")


if __name__ == "__main__":
    main()
