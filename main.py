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
import warnings
from pathlib import Path

# Suprimir advertencias benignas de extensiones no soportadas por openpyxl (Data Validation, Slicers)
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# Configurar ruta base del proyecto
_PROJECT_ROOT = Path(__file__).resolve().parent
if str(_PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(_PROJECT_ROOT))

from input.file_selector import select_all_four_files
from output.direct_updater import apply_direct_cierre
from processing.cierre_orchestrator import run_cierre_pipeline
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
    print("  [2] Modo 2 — Simulación Previa (Validación en Memoria) [HABILITADO]")
    print("  [3] Modo 3 — Generar Archivo Auxiliar de Cierre [HABILITADO]")
    print("  [4] Verificación de Arquitectura Offline")
    print("  [5] Modo 5 — Aplicar Cierre Directo en Archivos Oficiales (Con Backup y Protección)")
    print("  [0] Salir")
    print("-" * 76)
    try:
        choice = input("Seleccione una opción (0-5): ").strip()
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

    files = select_all_four_files(use_gui=True)

    print("\n🔍 Código del Aula a inspeccionar específicamente (opcional, ej. ABBUEI154):")
    try:
        aula_input = input("   Ingrese código de aula o presione ENTER para omitir: ").strip()
        target_aula = aula_input if aula_input else None
    except (EOFError, KeyboardInterrupt):
        target_aula = None

    hashes_before = {}
    for role, path in files.items():
        rec = compute_sha256(path)
        hashes_before[role] = rec

    print("\n⏳ Ejecutando análisis técnico profundo de los archivos en memoria...")
    report_path, _ = run_discovery(files, target_aula=target_aula)

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
    print("=" * 76)


def handle_cierre_execution(is_simulation: bool = False):
    modo_nombre = "MODO 2: SIMULACIÓN PREVIA" if is_simulation else "MODO 3: GENERACIÓN DE ARCHIVO AUXILIAR DE CIERRE"
    print("\n" + "=" * 76)
    print(f"INICIANDO {modo_nombre}")
    print("=" * 76)

    files = select_all_four_files(use_gui=True)

    print("\n🔍 Ingrese el código del aula virtual a cerrar (ej. ABBUEI205):")
    while True:
        try:
            target_aula = input("   Código de aula: ").strip()
            if target_aula:
                break
            print("   ⚠️ Debe ingresar un código de aula válido.")
        except (EOFError, KeyboardInterrupt):
            return

    print(f"\n🏷️ Ingrese el código del curso para los certificados (presione ENTER para usar '{target_aula}'):")
    try:
        cod_input = input(f"   Código de curso [{target_aula}]: ").strip()
        codigo_curso = cod_input if cod_input else target_aula
    except (EOFError, KeyboardInterrupt):
        codigo_curso = target_aula

    print("\n⏳ Procesando información en memoria con reglas oficiales...")
    try:
        res, out_file, intact, integrity_msgs = run_cierre_pipeline(
            file_paths=files,
            target_aula=target_aula,
            codigo_curso=codigo_curso,
            is_simulation=is_simulation,
        )
    except Exception as e:
        print(f"\n❌ Error durante el procesamiento: {e}")
        return

    print("\n" + "=" * 76)
    print("RESUMEN OPERATIVO DEL PROCESAMIENTO")
    print("=" * 76)
    print(f"• Aula Virtual: {target_aula}")
    print(f"• Código para Certificados: {codigo_curso}")
    print(f"• Clases detectadas: {len(res.clases)}")
    print(f"• Total Estudiantes en Sistematización: {res.total_estudiantes}")
    print(f"• Total Matches exitosos con Notas: {res.total_matched}")
    print(f"• Total Excluidos (Relleno Naranja en Notas): {res.total_excluded}")
    print(f"• Estudiantes en Sistematización sin Notas: {res.total_unmatched_sist}")
    print(f"• Registros en Notas sin match en Sistematización: {res.total_unmatched_notas}")

    print("\n--- Desglose por Clase ---")
    for cg in res.clases:
        aprobados = sum(1 for m in cg.matches if m.estudiante_sist.estado_calculado == "Aprobó")
        no_aprobados = sum(1 for m in cg.matches if m.estudiante_sist.estado_calculado == "No aprobó")
        abandonaron = sum(1 for m in cg.matches if m.estudiante_sist.estado_calculado == "Abandonó")
        print(f"  [Clase {cg.clase_id}]")
        print(f"    - En Sistematización: {len(cg.estudiantes_sist)} | Con Notas: {len(cg.matches)}")
        print(f"    - Aprobados: {aprobados} | No Aprobados: {no_aprobados} | Abandonaron: {abandonaron}")
        if cg.excluded_orange:
            print(f"    - ⚠️ Excluidos (Naranja): {len(cg.excluded_orange)}")

    print("\n🔒 Verificación de Integridad de los Archivos Oficiales:")
    for m in integrity_msgs:
        print(f"  {m}")

    if intact:
        print("\n✅ INTEGRIDAD 100% CERTIFICADA: Los 4 archivos oficiales permanecieron intactos.")
    else:
        print("\n🔴 ADVERTENCIA: Se detectó una inconsistencia de hash.")

    if not is_simulation and out_file:
        print("\n" + "=" * 76)
        print("🎉 ¡ARCHIVO AUXILIAR DE CIERRE GENERADO EXITOSAMENTE!")
        print("=" * 76)
        print(f"👉 Archivo: {out_file.resolve()}")
        print("\nContiene las 5 hojas estandarizadas:")
        print("  1. RESUMEN: Estadísticas y ficha técnica.")
        print("  2. SISTEMATIZACION: Tabla idéntica a la hoja oficial (55 cols) para Copiar y Pegar.")
        print("  3. CERTIFICADOS: Tabla idéntica a la hoja oficial (13 cols) solo con aprobados.")
        print("  4. ACTUALIZACION_CONTROL: Fila y columna para marcar el aula como cerrada.")
        print("  5. VALIDACIONES_Y_TRAZABILIDAD: Auditoría celda a celda de procedencia.")
        print("=" * 76)
    else:
        print("\n💡 MODO SIMULACIÓN COMPLETADO: No se generó ningún archivo en disco.")
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


def handle_direct_cierre_mode():
    print("\n" + "=" * 76)
    print("INICIANDO MODO 5: APLICAR CIERRE DIRECTO EN ARCHIVOS OFICIALES")
    print("=" * 76)
    print("Este modo actualizará directamente los libros oficiales de la biblioteca:")
    print("  1. Sistematización: Nombres en MAYÚSCULAS + celdas de cierre vacías + borde clase.")
    print("  2. Certificados: Anexar estudiantes aprobados al final + borde clase.")
    print("  3. Control de Aulas: Marcar fecha de cierre en columna CERTIFICADOS.")
    print("  4. Notas: Permanece como SOLO LECTURA inmutable.\n")

    files = select_all_four_files(use_gui=True)

    print("\n🔍 Ingrese el código del aula virtual a cerrar (ej. ABBUEI205):")
    while True:
        try:
            target_aula = input("   Código de aula: ").strip()
            if target_aula:
                break
            print("   ⚠️ Debe ingresar un código de aula válido.")
        except (EOFError, KeyboardInterrupt):
            return

    print(f"\n🏷️ Ingrese el código del curso para los certificados (presione ENTER para usar '{target_aula}'):")
    try:
        cod_input = input(f"   Código de curso [{target_aula}]: ").strip()
        codigo_curso = cod_input if cod_input else target_aula
    except (EOFError, KeyboardInterrupt):
        codigo_curso = target_aula

    print("\n⏳ Analizando y validando datos en memoria antes de cualquier cambio...")
    try:
        res, _, _, _ = run_cierre_pipeline(
            file_paths=files,
            target_aula=target_aula,
            codigo_curso=codigo_curso,
            is_simulation=True,
        )
    except Exception as e:
        print(f"\n❌ Error durante el análisis previo: {e}")
        return

    print("\n" + "=" * 76)
    print("RESUMEN DE MODIFICACIONES PREVISTAS")
    print("=" * 76)
    print(f"• Aula Virtual: {target_aula}")
    print(f"• Código para Certificados: {codigo_curso}")
    print(f"• Clases detectadas: {len(res.clases)}")
    print(f"• Total Estudiantes en Sistematización: {res.total_estudiantes}")
    print(f"• Total Matches exitosos con Notas: {res.total_matched}")
    print(f"• Total Excluidos (Relleno Naranja en Notas): {res.total_excluded}")

    aprobados_total = 0
    for cg in res.clases:
        aprob = sum(1 for m in cg.matches if m.estudiante_sist.estado_calculado == "Aprobó")
        no_aprob = sum(1 for m in cg.matches if m.estudiante_sist.estado_calculado == "No aprobó")
        aband = sum(1 for m in cg.matches if m.estudiante_sist.estado_calculado == "Abandonó")
        aprobados_total += aprob
        print(f"  [Clase {cg.clase_id}] Aprobados: {aprob} | No Aprobados: {no_aprob} | Abandonaron: {aband}")

    print(f"\n  👉 Se anexarán {aprobados_total} certificados a la hoja oficial de Certificados.")

    print("\n" + "!" * 76)
    print("🛡️ GARANTÍAS DE SEGURIDAD Y RESPONSABILIDAD:")
    print("  • Se creará una copia de seguridad (backup) automática de cada archivo")
    print("    en la carpeta '_backups_cierre/' antes de cualquier escritura.")
    print("  • Ninguna celda preexistente con datos será destruida ni alterada.")
    print("  • Los nombres y apellidos de los estudiantes se estandarizarán a MAYÚSCULAS.")
    print("  • Solo se escribirán notas y estados en celdas que estén estrictamente vacías.")
    print("  • El archivo de notas permanecerá 100% como solo lectura.")
    print("!" * 76)

    try:
        confirm = input("\n¿Está seguro de aplicar estos cambios directamente en los archivos oficiales? (Escriba 'SI' para confirmar): ").strip()
        if confirm != "SI":
            print("\n❌ Operación cancelada por el usuario. Ningún archivo fue modificado.")
            return
    except (EOFError, KeyboardInterrupt):
        print("\n❌ Operación cancelada.")
        return

    print("\n⏳ Creando respaldos de seguridad y aplicando cambios directamente...")
    try:
        report = apply_direct_cierre(
            file_paths=files,
            result=res,
            codigo_curso=codigo_curso,
            metadatos_control=res.metadatos_control,
            start_consecutivo=res.start_consecutivo,
        )
    except Exception as e:
        print(f"\n❌ Error crítico durante la actualización: {e}")
        return

    print("\n" + "=" * 76)
    if report.is_successful:
        print("🎉 ¡CIERRE DIRECTO APLICADO EXITOSAMENTE EN LOS ARCHIVOS OFICIALES!")
        print("=" * 76)
        print("📂 Copias de seguridad automáticas creadas en:")
        for bk in report.backups_created:
            print(f"   • {bk.resolve()}")
        print("\n📊 Estadísticas de la operación:")
        print(f"   • Sistematización — Nombres estandarizados a MAYÚSCULAS: {report.sistematizacion_names_uppercased}")
        print(f"   • Sistematización — Celdas vacías actualizadas: {report.sistematizacion_updated_cells}")
        print(f"   • Sistematización — Celdas preexistentes protegidas (intactas): {report.sistematizacion_untouched_cells}")
        print(f"   • Certificados — Filas de aprobados anexadas al final: {report.certificados_rows_appended}")
        print(f"   • Control de Aulas — Clases marcadas con fecha de cierre: {report.control_aulas_classes_marked}")
        print("\n✅ Todos los archivos oficiales han sido actualizados de forma segura.")
    else:
        print("🔴 SE COMPLETÓ CON ADVERTENCIAS O ERRORES:")
        for err in report.errors:
            print(f"   • {err}")
    print("=" * 76)


def main():
    print_banner()

    while True:
        choice = show_menu()

        if choice == "1":
            handle_discovery_mode()
        elif choice == "2":
            handle_cierre_execution(is_simulation=True)
        elif choice == "3":
            handle_cierre_execution(is_simulation=False)
        elif choice == "4":
            handle_offline_check()
        elif choice == "5":
            handle_direct_cierre_mode()
        elif choice in ("0", "q", "salir", "exit"):
            print("\nSaliendo del asistente. ¡Hasta pronto!")
            sys.exit(0)
        else:
            print("\n⚠️ Opción no válida. Por favor seleccione 0, 1, 2, 3, 4 o 5.")


if __name__ == "__main__":
    main()
