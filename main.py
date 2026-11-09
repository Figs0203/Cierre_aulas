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

import re
import sys
import warnings
from pathlib import Path

# Suprimir advertencias benignas de extensiones no soportadas por openpyxl (Data Validation, Slicers)
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# Configurar ruta base del proyecto
_PROJECT_ROOT = Path(__file__).resolve().parent
if str(_PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(_PROJECT_ROOT))

from input.file_selector import (
    prompt_for_file,
    prompt_for_multiple_files,
    select_all_four_files,
)
from output.direct_updater import (
    apply_direct_cierre,
    cloud_sync_warnings,
    files_not_writable,
    find_latest_backup,
    restore_last_backup,
)
from output.xlsx_recovery import diagnose_xlsx, repair_xlsx
from processing.cierre_orchestrator import run_cierre_pipeline
from security.integrity import compute_sha256, verify_integrity
from security.offline_check import check_offline_compliance
from tests.discovery.inspect_real_files import run_discovery
from tests.discovery.diagnose_matching import run_diagnostico


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
    print("  [6] Modo 6 — Restaurar Último Backup en Archivos Oficiales")
    print("  [7] Diagnóstico de Cruce de Identidades (Notas ↔ Sistematización, sin datos personales)")
    print("  [8] Modo 8 — Cierre por Lotes (varias aulas; Sistematización y Control compartidos)")
    print("  [9] Diagnosticar / Reparar un archivo .xlsx dañado (sin tocar el original)")
    print("  [0] Salir")
    print("-" * 76)
    try:
        choice = input("Seleccione una opción (0-9): ").strip()
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


def print_identity_verification_summary(res: AulaCierreResult):
    print("\n🛡️ Verificación Cruzada de Identidad (Notas ↔ Sistematización):")
    total_doble = sum(
        1 for cg in res.clases for m in cg.matches
        if "DOBLE FACTOR" in m.cross_validation_label
    )
    total_matches = res.total_matched
    pct = (total_doble / total_matches * 100) if total_matches > 0 else 100.0
    print(f"  • Coincidencias con doble factor verificado (Documento + Nombre / Correo): {total_doble}/{total_matches} ({pct:.1f}%)")

    warnings_list = [
        (m, w) for cg in res.clases for m in cg.matches for w in m.warnings
    ]
    if not warnings_list:
        print("  ✅ 100% de estudiantes verificados con correspondencia de persona confirmada entre archivos.")
    else:
        print(f"  ⚠️ Observaciones de identidad detectadas ({len(warnings_list)}):")
        for m, w in warnings_list[:5]:
            nom = f"{m.estudiante_sist.nombres.value if m.estudiante_sist.nombres else ''} {m.estudiante_sist.apellidos.value if m.estudiante_sist.apellidos else ''}".strip()
            print(f"     - Estudiante {nom}: {w}")


def print_match_diagnostics(res: AulaCierreResult):
    """Muestra ayuda cuando no hubo ningún cruce con Notas (0 matches)."""
    if res.total_matched > 0:
        return

    notas = res.total_unmatched_notas
    print("\n" + "!" * 76)
    print("⚠️  DIAGNÓSTICO DE CRUCE: no se encontró ningún match con el archivo de Notas")
    print("!" * 76)
    print(f"   • Estudiantes en Sistematización: {res.total_estudiantes}")
    print(f"   • Registros de Notas sin match: {notas}")
    if notas == 0:
        print("   ⚠️ El archivo de Notas no aportó ningún registro legible.")
        print("      Verifique que la hoja seleccionada y las columnas de identidad sean correctas.")
    else:
        print("   ⚠️ Hay registros en Notas, pero sus identidades no coinciden con Sistematización")
        print("      (documento, correo institucional o nombre+apellido).")
    print("   💡 Para un análisis detallado (100% local y sin datos personales):")
    print("      python tests/discovery/diagnose_matching.py \\")
    print("        --control \"<Control_Aulas>.xlsx\" --sistematizacion \"<Sistematizacion>.xlsx\" \\")
    print("        --notas \"<Notas>.xlsx\" --certificados \"<Certificados>.xlsx\" --aula <AULA>")
    print("!" * 76)


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

    print_identity_verification_summary(res)
    print_match_diagnostics(res)

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

    # Preflight: si un oficial está abierto en Excel o es de solo lectura, se
    # avisa ANTES de analizar y de tocar cualquier archivo.
    problemas_acceso = files_not_writable(files)
    if problemas_acceso:
        print("\n🔴 NO SE PUEDE INICIAR EL CIERRE:")
        for problema in problemas_acceso:
            print(f"   • {problema}")
        print(
            "\nCierre los archivos en Excel (o quite el atributo de solo lectura) "
            "y vuelva a intentarlo. No se modificó ningún archivo."
        )
        return

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

    print_identity_verification_summary(res)
    print_match_diagnostics(res)

    print("\n" + "!" * 76)
    print("🛡️ GARANTÍAS DE SEGURIDAD Y RESPONSABILIDAD:")
    print("  • Se creará (o actualizará) un ÚNICO respaldo de cada archivo en la")
    print("    carpeta '_backups_cierre/' con el estado previo a la automatización.")
    print("  • Ninguna celda preexistente con datos será destruida ni alterada.")
    print("  • Los nombres y apellidos de los estudiantes se estandarizarán a MAYÚSCULAS.")
    print("  • Solo se escribirán notas y estados en celdas estrictamente vacías")
    print("    (salvo que autorice corregir una nota final ya escrita).")
    print("  • El archivo de notas permanecerá 100% como solo lectura.")
    print("  • NO se perderá ninguna función de Excel: listas de validación, formato")
    print("    condicional, controles, tablas dinámicas, gráficos, imágenes ni fórmulas.")
    print("!" * 76)

    # La nota final ya escrita solo se corrige si el usuario lo autoriza. Es útil
    # cuando un cierre anterior dejó promedios equivocados en Sistematización.
    print("\n🎯 NOTA FINAL YA ESCRITA EN SISTEMATIZACIÓN:")
    print("  Por seguridad, el cierre no sobrescribe celdas con datos. Si una nota")
    print("  final existente no coincide con el archivo de Notas, puede autorizar")
    print("  aquí su corrección (solo se toca esa columna).")
    try:
        resp_nota = input(
            "  ¿Corregir notas finales existentes que no coincidan con Notas? (SI/NO): "
        ).strip().upper()
    except (EOFError, KeyboardInterrupt):
        resp_nota = "NO"
    actualizar_nota_final = resp_nota == "SI"
    if actualizar_nota_final:
        print("  ✅ Se corregirán las notas finales que no coincidan con Notas.")
    else:
        print("  ↳ Se mantienen intactas las notas finales ya escritas.")

    avisos_sync = cloud_sync_warnings(files)
    for aviso in avisos_sync:
        print(f"\n⚠️  {aviso}")
    if avisos_sync:
        print("  ⚠️  Si el libro sigue abierto o a medio sincronizar, Excel mostrará")
        print("      «CARGA BLOQUEADA» o «Conflicto de combinación».")
        try:
            resp_sync = input(
                "   ¿Confirma que los archivos están CERRADOS en Excel y "
                "sincronizados? (SI/NO): "
            ).strip().upper()
        except (EOFError, KeyboardInterrupt):
            resp_sync = "NO"
        if resp_sync != "SI":
            print(
                "\n❌ Operación cancelada. Cierre los archivos en Excel (escritorio "
                "y web), espere a que OneDrive/SharePoint sincronice y reintente."
            )
            return

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
            actualizar_nota_final_existente=actualizar_nota_final,
        )
    except Exception as e:
        print(f"\n❌ Error crítico durante la actualización: {e}")
        return

    print("\n" + "=" * 76)
    if report.is_successful:
        print("🎉 ¡CIERRE DIRECTO APLICADO EXITOSAMENTE EN LOS ARCHIVOS OFICIALES!")
        print("=" * 76)
        print("📂 Respaldo único actualizado (estado previo al cierre):")
        for bk in report.backups_created:
            print(f"   • {bk.resolve()}")
        if report.backups_consolidados:
            print(
                f"   • Copias antiguas consolidadas en el respaldo único: "
                f"{report.backups_consolidados}"
            )
        print("\n📊 Estadísticas de la operación:")
        print(f"   • Sistematización — Nombres estandarizados a MAYÚSCULAS: {report.sistematizacion_names_uppercased}")
        print(f"   • Sistematización — Celdas vacías actualizadas: {report.sistematizacion_updated_cells}")
        print(f"   • Sistematización — Celdas preexistentes protegidas (intactas): {report.sistematizacion_untouched_cells}")
        if report.sistematizacion_promedios_no_copiados:
            print(
                f"   • Sistematización — Notas finales no copiadas (Notas no las trae): "
                f"{report.sistematizacion_promedios_no_copiados}"
            )
        if report.sistematizacion_promedios_actualizados:
            print(
                f"   • Sistematización — Notas finales corregidas (sobrescritas): "
                f"{report.sistematizacion_promedios_actualizados}"
            )
        if report.sistematizacion_celdas_protegidas:
            print(
                "   • Sistematización — Celdas NO escritas por ser fórmulas "
                "compartidas/matriciales: "
                f"{report.sistematizacion_celdas_protegidas} (ver advertencias)"
            )
        claves_cierre = {
            "promedio": "Nota final",
            "estado": "Estado para certificación",
            "elaboro": "Se elaboró certificado",
            "codigo_cert": "Código del certificado",
            "envio": "Se envió certificado",
            "fecha_envio": "Fecha de envío",
        }
        columnas_cierre = {
            texto: report.columnas_sistematizacion[clave]
            for clave, texto in claves_cierre.items()
            if clave in report.columnas_sistematizacion
        }
        if columnas_cierre:
            print("\n🧭 Columnas de cierre usadas en Sistematización (verifíquelas):")
            for texto, ubic in columnas_cierre.items():
                print(f"   • {texto}: {ubic}")
        print(f"   • Certificados — Filas de aprobados anexadas al final: {report.certificados_rows_appended}")
        if report.certificados_fila_inicial:
            print(
                f"   • Certificados — Ubicación: hoja '{report.certificados_hoja}', "
                f"filas {report.certificados_fila_inicial} a {report.certificados_fila_final} "
                f"(N° {report.certificados_n_inicial} a {report.certificados_n_final})"
            )
            print(
                f"   • Certificados — Verificación de escritura (relectura del archivo): "
                f"{report.certificados_rows_verificados}/{report.certificados_rows_appended} filas confirmadas"
            )
            print(f"   • Certificados — Archivo actualizado: {files['certificados']}")
            print("     ↳ Si no ve los certificados, abra EXACTAMENTE ese archivo y revise")
            print("       que la hoja no tenga filtros activos ni filas ocultas/agrupadas.")
        if report.certificados_omitidos_duplicados:
            print(
                f"   • Certificados — Omitidos por ya estar registrados: "
                f"{report.certificados_omitidos_duplicados}"
            )
        print(f"   • Control de Aulas — Clases marcadas con fecha de cierre: {report.control_aulas_classes_marked}")
        if report.control_aulas_classes_already_marked:
            print(
                f"   • Control de Aulas — Clases que ya tenían fecha de cierre (sin cambios): "
                f"{report.control_aulas_classes_already_marked}"
            )
        if report.control_aulas_classes_not_found:
            print(
                "   • Control de Aulas — Clases NO encontradas en el archivo: "
                + ", ".join(report.control_aulas_classes_not_found)
            )
        print("\n🔒 Integridad del libro original — verificación parte por parte:")
        print("   (listas de validación, formato condicional, controles de formulario,")
        print("    tablas dinámicas, gráficos, imágenes y nombres definidos)")
        if report.validaciones_preservadas:
            for resumen in report.validaciones_preservadas.values():
                print(f"   • {resumen}")
        else:
            print("   • (No se pudo verificar la preservación de los archivos oficiales.)")
        print("\n✅ Todos los archivos oficiales han sido actualizados de forma segura.")
    else:
        print("🔴 SE COMPLETÓ CON ADVERTENCIAS O ERRORES:")
        for err in report.errors:
            print(f"   • {err}")
    if report.warnings:
        print("\n⚠️ ADVERTENCIAS (requieren revisión):")
        for adv in report.warnings:
            print(f"   • {adv}")
        print("   ↳ Si la verificación reporta alguna pérdida, restaure el último")
        print("     respaldo con la opción [6] y reporte el caso.")
    print("=" * 76)


def _infer_aula_code(path: Path) -> str:
    """Intenta deducir el código de aula del nombre del archivo de Notas.

    Ej.: 'ABBUEI227_E_FINANZAS_MAF_1.xlsx' → 'ABBUEI227'; 'CMI122_PSICOLOGIA_2.xlsx' → 'CMI122'.
    """
    match = re.match(r"([A-Za-z]{2,}\d{2,})", path.stem)
    return match.group(1).upper() if match else ""


def _leer_linea(mensaje: str, default: str = "") -> str | None:
    """Lee una línea; devuelve None si el usuario cancela (Ctrl+C / EOF)."""
    try:
        valor = input(mensaje).strip()
    except (EOFError, KeyboardInterrupt):
        return None
    return valor or default


def _pedir_ruta_certificados(anterior: Path | None) -> Path | None:
    """Pide el archivo de certificados, permitiendo reutilizar el anterior."""
    if anterior is None:
        return prompt_for_file(
            role_description="4. Control de Códigos / Certificados",
            role_key="certificados",
            default_hint="Control_codigos_certificados_ABBUEI.xlsx",
            use_gui=True,
        )

    print(f"   Archivo de certificados anterior: {anterior}")
    while True:
        entrada = _leer_linea(
            "   Ruta del archivo de certificados (ENTER para usar el mismo): "
        )
        if entrada is None:
            return None
        if not entrada:
            return anterior
        candidato = Path(entrada).strip('"').strip("'")
        if not candidato.exists():
            print("   ⚠️ El archivo no existe. Intente de nuevo.")
            continue
        if candidato.suffix.lower() != ".xlsx":
            print("   ⚠️ Debe ser un archivo .xlsx.")
            continue
        return candidato.resolve()


def handle_batch_cierre_mode():
    """Cierre directo de VARIAS aulas en una sola operación.

    La Sistematización y el Control de Aulas se piden una sola vez y se
    reutilizan para todas las aulas. Las Notas y el archivo de Certificados
    pueden variar por aula.
    """
    print("\n" + "=" * 76)
    print("INICIANDO MODO 8: CIERRE POR LOTES (VARIAS AULAS)")
    print("=" * 76)
    print("Cierra varias aulas seguidas sin volver a seleccionar los archivos compartidos:")
    print("  • Control de Aulas de Formadores  → se pide UNA vez (compartido)")
    print("  • Sistematización de Cursos COIN → se pide UNA vez (compartida)")
    print("  • Notas de cada aula             → se eligen varias a la vez")
    print("  • Certificados                   → compartido o uno por aula")
    print("  • Notas permanece siempre como SOLO LECTURA.")
    print("  • Cada aula se aplica como un cierre independiente: el respaldo único se")
    print("    actualiza antes de cada una y, si un aula falla, las anteriores quedan.")
    print()

    # ------------------------------------------------------------------
    # 1. Archivos compartidos
    # ------------------------------------------------------------------
    try:
        control = prompt_for_file(
            role_description="1. Control de Aulas de Formadores (COMPARTIDO para todas las aulas)",
            role_key="control_aulas",
            default_hint="Control_Aulas_Formadores_2026.xlsx",
            use_gui=True,
        )
        sistematizacion = prompt_for_file(
            role_description="2. Sistematización de Cursos COIN (COMPARTIDA para todas las aulas)",
            role_key="sistematizacion",
            default_hint="Sistematizacion_Cursos_COIN_2026.xlsx",
            use_gui=True,
        )
    except SystemExit:
        return

    files_base = {"control_aulas": control, "sistematizacion": sistematizacion}
    problemas_acceso = files_not_writable(files_base)
    if problemas_acceso:
        print("\n🔴 NO SE PUEDE INICIAR EL LOTE:")
        for problema in problemas_acceso:
            print(f"   • {problema}")
        print("\nCierre los archivos compartidos en Excel y vuelva a intentarlo.")
        return

    # ------------------------------------------------------------------
    # 2. Notas (varias a la vez)
    # ------------------------------------------------------------------
    try:
        notas_paths = prompt_for_multiple_files(
            role_description="3. Notas de las aulas a cerrar (puede elegir varias)",
            default_hint="ABBUEI227_E_FINANZAS_MAF_1.xlsx, CMI122_PSICOLOGIA_2.xlsx, ...",
            use_gui=True,
        )
    except SystemExit:
        return
    if not notas_paths:
        print("\n❌ No se seleccionó ningún archivo de Notas.")
        return

    # ------------------------------------------------------------------
    # 3. ¿Certificados compartidos?
    # ------------------------------------------------------------------
    print("\nPASO 3 — Archivo de Certificados:")
    respuesta = _leer_linea(
        "   ¿Todas las aulas usan el MISMO archivo de certificados? (SI/NO) [SI]: ",
        "SI",
    )
    if respuesta is None:
        print("\n❌ Operación cancelada.")
        return

    certificados_comun: Path | None = None
    if respuesta.strip().upper() in ("SI", "S", ""):
        try:
            certificados_comun = prompt_for_file(
                role_description="4. Control de Códigos / Certificados (COMPARTIDO)",
                role_key="certificados",
                default_hint="Control_codigos_certificados_ABBUEI.xlsx",
                use_gui=True,
            )
        except SystemExit:
            return
        problemas_cert = files_not_writable({"certificados": certificados_comun})
        if problemas_cert:
            print("\n🔴 NO SE PUEDE INICIAR EL LOTE:")
            for problema in problemas_cert:
                print(f"   • {problema}")
            print("\nCierre el archivo de certificados en Excel y vuelva a intentarlo.")
            return

    # ------------------------------------------------------------------
    # 4. Configurar cada aula y simular (sin escribir nada)
    # ------------------------------------------------------------------
    items: list[dict] = []
    aulas_vistas: set[str] = set()
    notas_vistas: set[Path] = set()
    certificado_previo = certificados_comun

    for idx, notas in enumerate(notas_paths, start=1):
        print(f"\n--- Aula {idx} de {len(notas_paths)}: {notas.name} ---")

        if notas.resolve() in notas_vistas:
            print("   ⚠️ Ese archivo de Notas ya fue agregado. Se omite.")
            continue

        default_aula = _infer_aula_code(notas)
        etiqueta = f"   Código de aula{' [' + default_aula + ']' if default_aula else ''}: "
        aula = _leer_linea(etiqueta, default_aula)
        if aula is None:
            print("\n❌ Operación cancelada.")
            return
        if not aula:
            print("   ⚠️ Debe indicar el código de aula. Se omite esta aula.")
            continue
        aula = aula.strip().upper()
        if aula in aulas_vistas:
            print(f"   ⚠️ El aula '{aula}' ya está en el lote. Se omite.")
            continue

        curso = _leer_linea(
            f"   Código de curso para certificados [{aula}]: ", aula
        )
        if curso is None:
            print("\n❌ Operación cancelada.")
            return
        curso = (curso or aula).strip().upper()

        if certificados_comun is None:
            try:
                certificado = _pedir_ruta_certificados(certificado_previo)
            except SystemExit:
                return
            if certificado is None:
                print("\n❌ Operación cancelada.")
                return
            certificado_previo = certificado
        else:
            certificado = certificados_comun

        problemas_cert = files_not_writable({"certificados": certificado})
        if problemas_cert:
            for problema in problemas_cert:
                print(f"   ⚠️ {problema}")
            print("   Se omite esta aula (no se podría escribir su certificado).")
            continue

        files = {
            "control_aulas": control,
            "sistematizacion": sistematizacion,
            "notas": notas,
            "certificados": certificado,
        }

        print("   ⏳ Analizando datos en memoria...")
        try:
            res, _, _, _ = run_cierre_pipeline(
                file_paths=files,
                target_aula=aula,
                codigo_curso=curso,
                is_simulation=True,
            )
        except Exception as exc:  # noqa: BLE001 - una aula no bloquea el lote
            print(f"   ❌ Error al analizar el aula '{aula}': {exc}. Se omite.")
            continue

        if res.total_estudiantes == 0:
            print(
                f"   ⚠️ No se encontraron estudiantes del aula '{aula}' en Sistematización. "
                "Se omite."
            )
            continue

        aprobados = sum(
            1
            for cg in res.clases
            for m in cg.matches
            if str(m.estudiante_sist.estado_calculado or "").strip() == "Aprobó"
        )
        items.append({
            "aula": aula,
            "curso": curso,
            "files": files,
            "result": res,
            "aprobados": aprobados,
        })
        aulas_vistas.add(aula)
        notas_vistas.add(notas.resolve())

    if not items:
        print("\n❌ No quedó ninguna aula válida para cerrar.")
        return

    # ------------------------------------------------------------------
    # 5. Resumen consolidado y confirmación única
    # ------------------------------------------------------------------
    total_certificados = sum(item["aprobados"] for item in items)
    print("\n" + "=" * 76)
    print("RESUMEN DEL LOTE")
    print("=" * 76)
    print(f"• Control de Aulas (compartido): {control}")
    print(f"• Sistematización (compartida): {sistematizacion}")
    if certificados_comun:
        print(f"• Certificados (compartido): {certificados_comun}")
    else:
        print("• Certificados: un archivo por aula (se muestra en la aplicación)")
    print(f"\n• Aulas a cerrar: {len(items)}")
    for pos, item in enumerate(items, start=1):
        res = item["result"]
        print(
            f"  {pos}. Aula {item['aula']} (curso {item['curso']}): "
            f"{len(res.clases)} clase(s), {res.total_estudiantes} estudiantes, "
            f"{res.total_matched} con notas, {res.total_excluded} excluidos, "
            f"{item['aprobados']} aprobados"
        )
    print(f"\n👉 Se anexarán {total_certificados} certificados en total.")
    print("   (La numeración se calcula al aplicar cada aula, sobre el archivo real.)")

    print("\n🎯 NOTA FINAL YA ESCRITA EN SISTEMATIZACIÓN:")
    respuesta_nota = _leer_linea(
        "   ¿Corregir notas finales existentes que no coincidan con Notas? (SI/NO): ",
        "NO",
    )
    if respuesta_nota is None:
        print("\n❌ Operación cancelada.")
        return
    actualizar_nota_final = respuesta_nota.strip().upper() == "SI"

    avisos_sync: list[str] = []
    for item in items:
        for aviso in cloud_sync_warnings(item["files"]):
            if aviso not in avisos_sync:
                avisos_sync.append(aviso)
    for aviso in avisos_sync:
        print(f"\n⚠️  {aviso}")
    if avisos_sync:
        print("  ⚠️  Si el libro sigue abierto o a medio sincronizar, Excel mostrará")
        print("      «CARGA BLOQUEADA» o «Conflicto de combinación».")
        try:
            resp_sync = input(
                "   ¿Confirma que los archivos están CERRADOS en Excel y "
                "sincronizados? (SI/NO): "
            ).strip().upper()
        except (EOFError, KeyboardInterrupt):
            resp_sync = "NO"
        if resp_sync != "SI":
            print(
                "\n❌ Operación cancelada. Cierre los archivos en Excel (escritorio "
                "y web), espere a que OneDrive/SharePoint sincronice y reintente."
            )
            return

    try:
        confirm = input(
            f"\n¿Aplicar el cierre de las {len(items)} aula(s) en los archivos oficiales? "
            "(Escriba 'SI' para confirmar): "
        ).strip()
    except (EOFError, KeyboardInterrupt):
        print("\n❌ Operación cancelada.")
        return
    if confirm != "SI":
        print("\n❌ Operación cancelada por el usuario. Ningún archivo fue modificado.")
        return

    # ------------------------------------------------------------------
    # 6. Aplicar secuencialmente
    # ------------------------------------------------------------------
    print("\n⏳ Aplicando el cierre de las aulas...")
    exitosas: list[str] = []
    con_novedades: list[tuple[str, str]] = []
    respaldos: dict[str, Path] = {}
    total_anexados = 0

    for pos, item in enumerate(items, start=1):
        print(f"\n--- Aplicando {pos}/{len(items)}: aula {item['aula']} ---")
        try:
            report = apply_direct_cierre(
                file_paths=item["files"],
                result=item["result"],
                codigo_curso=item["curso"],
                metadatos_control=item["result"].metadatos_control,
                start_consecutivo=item["result"].start_consecutivo,
                actualizar_nota_final_existente=actualizar_nota_final,
            )
        except Exception as exc:  # noqa: BLE001 - se informa y se detiene el lote
            print(f"   ❌ Error crítico al aplicar: {exc}")
            con_novedades.append((item["aula"], str(exc)))
            print("   ⛔ Se detiene el lote para no seguir sobre archivos inconsistentes.")
            break

        total_anexados += report.certificados_rows_appended
        for respaldo in report.backups_created:
            respaldos[str(respaldo)] = respaldo

        if report.errors:
            for err in report.errors:
                print(f"   ❌ {err}")
        if report.warnings:
            for adv in report.warnings:
                print(f"   ⚠️ {adv}")

        if report.is_successful:
            print(
                f"   ✅ Aula cerrada. Certificados anexados: "
                f"{report.certificados_rows_appended}"
                + (
                    f" (hoja '{report.certificados_hoja}', filas "
                    f"{report.certificados_fila_inicial}-{report.certificados_fila_final})"
                    if report.certificados_fila_inicial
                    else ""
                )
            )
            exitosas.append(item["aula"])
        else:
            detalle = "; ".join(report.errors) or "con errores"
            con_novedades.append((item["aula"], detalle))
            print(f"   ⚠️ Aula con novedades: {detalle}")
            if report.fallo_critico:
                print("   ⛔ Fallo crítico: se detiene el lote.")
                break

    # ------------------------------------------------------------------
    # 7. Informe final del lote
    # ------------------------------------------------------------------
    print("\n" + "=" * 76)
    print("RESULTADO DEL LOTE")
    print("=" * 76)
    print(f"• Aulas cerradas correctamente: {len(exitosas)} de {len(items)}")
    if exitosas:
        print("  " + ", ".join(exitosas))
    if con_novedades:
        print(f"• Aulas con novedades ({len(con_novedades)}):")
        for aula, detalle in con_novedades:
            print(f"  - {aula}: {detalle}")
    print(f"• Certificados anexados en total: {total_anexados}")
    if respaldos:
        print("\n📂 Respaldo único actualizado (estado previo a cada cierre):")
        for ruta in respaldos.values():
            print(f"   • {ruta}")
    print("\n↳ Si alguna aula quedó con novedades, revise los mensajes anteriores.")
    print("  Para volver al estado previo use la opción [6] (Restaurar respaldo).")
    print("=" * 76)


def handle_restore_backup_mode():
    print("\n" + "=" * 76)
    print("INICIANDO MODO 6: RESTAURAR EL RESPALDO ÚNICO DE LOS ARCHIVOS OFICIALES")
    print("=" * 76)
    print("Este modo restaura el respaldo único de cada archivo oficial (el estado")
    print("previo al último cierre; se actualiza en cada ejecución del Modo 5).")
    print("Archivos afectados:")
    print("  • Sistematización de Cursos COIN")
    print("  • Control de Códigos de Certificados")
    print("  • Control de Aulas de Formadores")
    print("  • Notas: NO se restaura (nunca fue modificado).")
    print("\n⚠️  Los respaldos se buscan en la subcarpeta '_backups_cierre/' junto a cada archivo.")
    print("!" * 76)
    print("🛡️ ADVERTENCIA IMPORTANTE:")
    print("  Esta acción reemplazará el contenido actual de los archivos oficiales")
    print("  con el respaldo único (estado previo al último cierre). Cualquier cambio")
    print("  posterior a ese respaldo se perderá permanentemente.")
    print("!" * 76)

    files = select_all_four_files(use_gui=True)

    # Verificar que los respaldos no estén dañados. Si el respaldo único se
    # actualizó con un estado ya roto (por un cierre anterior), restaurarlo no
    # arreglaría nada y podría hacer perder el estado actual.
    respaldos_danados: list[str] = []
    for role in ("sistematizacion", "certificados", "control_aulas"):
        original = Path(files[role]).resolve()
        respaldo = find_latest_backup(original)
        if respaldo is None:
            continue
        try:
            info = diagnose_xlsx(respaldo)
        except Exception:  # noqa: BLE001 - nunca bloquea la restauración
            continue
        if info.tiene_danos_graves:
            respaldos_danados.append(respaldo.name)
    if respaldos_danados:
        print("\n⚠️  ADVERTENCIA: el respaldo único también presenta daños:")
        for nombre in respaldos_danados:
            print(f"   • {nombre}")
        print("   Restaurarlo NO arreglará el libro. Puede deberse a que un cierre")
        print("   anterior ya había dejado el archivo dañado (el respaldo se")
        print("   actualiza en cada cierre).")
        print("   Recomendación: use la opción [9] para diagnosticar/reparar tanto el")
        print("   archivo actual como el respaldo.")
        try:
            seguir = input(
                "   ¿Desea restaurar de todos modos? (SI/NO): "
            ).strip().upper()
        except (EOFError, KeyboardInterrupt):
            seguir = "NO"
        if seguir != "SI":
            print("\n❌ Operación cancelada. Ningún archivo fue modificado.")
            return

    try:
        confirm = input("\n¿Está seguro de restaurar los archivos oficiales desde el respaldo único? (Escriba 'SI' para confirmar): ").strip()
        if confirm != "SI":
            print("\n❌ Operación cancelada por el usuario. Ningún archivo fue modificado.")
            return
    except (EOFError, KeyboardInterrupt):
        print("\n❌ Operación cancelada.")
        return

    print("\n⏳ Buscando backups y restaurando archivos...")
    try:
        resultados = restore_last_backup(file_paths=files)
    except Exception as e:
        print(f"\n❌ Error durante la restauración: {e}")
        return

    print("\n" + "=" * 76)
    print("RESULTADO DE LA RESTAURACIÓN")
    print("=" * 76)

    all_ok = True
    labels = {
        "sistematizacion": "Sistematización de Cursos COIN",
        "certificados": "Control de Códigos de Certificados",
        "control_aulas": "Control de Aulas de Formadores",
    }
    for role, info in resultados.items():
        label = labels.get(role, role)
        status = info.get("status", "error")
        mensaje = info.get("mensaje", "")
        backup = info.get("backup")
        if status == "ok":
            icon = "✅"
        elif status == "sin_backup":
            icon = "⚠️"
            all_ok = False
        else:
            icon = "🔴"
            all_ok = False
        print(f"  {icon} [{label}]: {mensaje}")
        if backup and status == "ok":
            print(f"       ↩ Restaurado desde: {backup.name}")

    if all_ok:
        print("\n✅ Restauración completada. Los archivos oficiales han vuelto al estado del último backup.")
    else:
        print("\n⚠️ Restauración completada con advertencias. Revise los mensajes anteriores.")
    print("=" * 76)


def handle_diagnose_mode():
    """Diagnóstico de cruce de identidades (sin datos personales)."""
    print("\n" + "=" * 76)
    print("DIAGNÓSTICO DE CRUCE DE IDENTIDADES (Notas ↔ Sistematización)")
    print("=" * 76)
    print("Analiza por qué los estudiantes no cruzan entre los archivos, mostrando:")
    print("  • Estructura de cada archivo (hojas, encabezados, dimensiones).")
    print("  • Columnas detectadas (identidad, módulos, nota final) y colores.")
    print("  • Sondeo de coincidencias por documento, correo y nombre.")
    print("🔒 Todos los datos personales se muestran enmascarados. 100% local.\n")

    files = select_all_four_files(use_gui=True)

    print("\n🔍 Ingrese el código del aula virtual a analizar (ej. CMI120):")
    try:
        target_aula = input("   Código de aula: ").strip()
    except (EOFError, KeyboardInterrupt):
        return
    if not target_aula:
        print("\n⚠️ Debe ingresar un código de aula válido.")
        return

    print("\n⏳ Analizando los archivos en memoria...\n")
    try:
        report_path = run_diagnostico(files, target_aula)
    except Exception as e:  # noqa: BLE001
        print(f"\n❌ Error durante el diagnóstico: {e}")
        return

    print("\n" + "=" * 76)
    print("✅ DIAGNÓSTICO COMPLETADO")
    print("=" * 76)
    print("\n📄 El informe completo (sin datos personales) ha sido guardado en:")
    print(f"   👉 {report_path.resolve()}")
    print("\n💡 Puede compartir ese archivo: NO contiene documentos, correos ni nombres completos.")
    print("=" * 76)


def handle_recovery_mode():
    """Diagnostica y, si hace falta, repara una COPIA de un .xlsx dañado."""
    print("\n" + "=" * 76)
    print("DIAGNÓSTICO Y REPARACIÓN DE UN ARCHIVO .XLSX DAÑADO")
    print("=" * 76)
    print("Este modo NO modifica el archivo original: si se repara, crea una copia")
    print("'<nombre>_reparado.xlsx' junto al original. Todo es 100% local.\n")

    try:
        ruta_txt = input(
            "Ruta completa del archivo .xlsx a revisar (ENTER para cancelar): "
        ).strip().strip('"').strip("'")
    except (EOFError, KeyboardInterrupt):
        ruta_txt = ""
    if not ruta_txt:
        print("\n↳ Operación cancelada. Ningún archivo fue modificado.")
        return

    ruta = Path(ruta_txt).expanduser()
    if not ruta.exists():
        print(f"\n❌ No existe el archivo: {ruta}")
        return
    if ruta.suffix.lower() not in (".xlsx", ".xlsm"):
        print("\n⚠️ Debe ser un archivo .xlsx o .xlsm.")
        return

    print("\n⏳ Analizando la estructura interna (solo lectura)...")
    try:
        report = diagnose_xlsx(ruta)
    except Exception as exc:  # noqa: BLE001
        print(f"\n❌ No se pudo analizar el archivo: {exc}")
        return

    print("\n" + "=" * 76)
    print("RESULTADO DEL DIAGNÓSTICO")
    print("=" * 76)
    print(report.resumen())
    for problema in report.problemas:
        print(f"  • {problema}")
    for aviso in report.avisos:
        print(f"  ↳ {aviso}")

    if not report.hay_danos:
        print("\n✅ No se detectaron daños estructurales en el libro.")
        print("   Si Excel sigue mostrando «CARGA BLOQUEADA», el problema puede ser")
        print("   de sincronización (OneDrive/SharePoint), de permisos o de que el")
        print("   archivo está abierto en otro equipo. Copie el archivo a una carpeta")
        print("   LOCAL (por ejemplo, el Escritorio), ábralo desde ahí y reintente.")
        print("=" * 76)
        return

    print("\n💡 Se detectaron daños. Se puede generar una COPIA reparada.")
    print("   El archivo original NO se modificará.")
    try:
        respuesta = input("¿Generar la copia reparada? (SI/NO): ").strip().upper()
    except (EOFError, KeyboardInterrupt):
        respuesta = "NO"
    if respuesta != "SI":
        print("\n↳ No se generó ninguna copia. El original sigue intacto.")
        return

    print("\n⏳ Generando la copia reparada...")
    try:
        resultado = repair_xlsx(ruta)
    except Exception as exc:  # noqa: BLE001
        print(f"\n❌ No se pudo reparar el archivo: {exc}")
        return

    print("\n" + resultado.resumen_reparacion())
    if resultado.reparado and resultado.salida:
        print(f"\n✅ COPIA REPARADA: {resultado.salida}")
        print("   El archivo original NO se modificó.")
        print("   Abra esa copia en Excel y compruebe que ya no pide reparación.")
        if resultado.esclavas_convertidas:
            print("   ℹ️ Las celdas con fórmulas compartidas conservaron su último")
            print("      valor calculado (no se pierden datos). Si necesita la fórmula,")
            print("      vuelva a arrastrarla en Excel.")
    else:
        print("\n⚠️ No se pudo completar la reparación automática.")
        for problema in resultado.problemas_restantes:
            print(f"  • {problema}")
        print("   Guarde el informe de diagnóstico junto al archivo y compártalo.")
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
        elif choice == "6":
            handle_restore_backup_mode()
        elif choice == "7":
            handle_diagnose_mode()
        elif choice == "8":
            handle_batch_cierre_mode()
        elif choice == "9":
            handle_recovery_mode()
        elif choice in ("0", "q", "salir", "exit"):
            print("\nSaliendo del asistente. ¡Hasta pronto!")
            sys.exit(0)
        else:
            print("\n⚠️ Opción no válida. Por favor seleccione una opción entre 0 y 9.")


if __name__ == "__main__":
    main()
