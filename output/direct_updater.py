"""
Módulo de Actualización Directa sobre Archivos Oficiales (Biblioteca EAFIT).

Este módulo ejecuta el cierre directamente sobre los libros Excel oficiales con
máxima responsabilidad y mecanismos de seguridad de nivel institucional:

1. Respaldo Automático Previo (Zero-Data-Loss):
   Copia íntegra del estado actual en la subcarpeta '_backups_cierre/' antes de
   tocar cualquier archivo. Se mantiene **un único respaldo por archivo**: en
   cada cierre se actualiza (sobrescribe) con el estado previo a la operación y
   se eliminan las copias antiguas con marca de tiempo del mismo archivo.

2. Regla Estricta de No-Sobreescritura:
   - Celdas con datos preexistentes: Se respetan intactas y no se tocan.
   - Excepción autorizada: Nombres y apellidos de los estudiantes procesados se
     estandarizan a mayúsculas sostenidas (.upper()).
   - Celdas de cierre (módulos, promedio, estado, certificados): Solo se escriben
     si la celda está estrictamente vacía.

3. Separadores de Clase:
   Borde inferior grueso ('medium', color negro) al final de cada clase en
   Sistematización y en Certificados.

4. Archivo de Notas:
   Se mantiene estrictamente como SOLO LECTURA (no se modifica).

5. Preservación total del libro original (Zero-Loss):
   El cierre NO vuelve a escribir el libro con openpyxl. Se usa `XlsxPatcher`
   (`output/xlsx_patcher.py`), que modifica únicamente las celdas concretas dentro
   de `xl/worksheets/sheetN.xml` y copia byte a byte el resto del paquete. Así se
   conservan las listas de validación (estándar y `x14`), el formato condicional,
   los nombres definidos, los controles de formulario, las tablas dinámicas, los
   gráficos, las imágenes y los vínculos externos. Al final se verifica parte por
   parte (`verify_preservation`) que no se haya perdido nada.
"""

from __future__ import annotations

import os
import re
import shutil
import zipfile
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from xml.etree import ElementTree

import openpyxl
from openpyxl.styles import Alignment, Border, Font, Side

from config.settings import (
    CERTIFICADO_NO,
    CERTIFICADO_SI,
    CONTROL_AULAS_ALL_SHEETS,
    CONTROL_AULAS_IGNORED_SHEET,
    VALOR_NO_APLICA,
)
from core.models import AulaCierreResult, AulaMetadata, ClaseGroup, StudentMatch
from normalization.dates import (
    format_certificate_date,
    format_ciclo_date,
    format_semester_label,
    parse_date,
)
from normalization.text import format_ciclo, format_docente_coin, normalize_document
from output.xlsx_patcher import BorderSpec, PreservationReport, XlsxPatcher, verify_preservation

# Estilos de bordes oficiales.
# Todos los bordes se dibujan en negro (mismo color en los cuatro lados) para que
# sean visibles y consistentes con el formato de los archivos oficiales.
_BORDER_COLOR = "FF000000"

# La celda normal lleva la caja completa con bordes finos.
_BORDER_THIN = BorderSpec(
    left="thin", right="thin", top="thin", bottom="thin", color=_BORDER_COLOR
)

# La celda de fin de clase conserva los cuatro bordes: los tres lados con el mismo
# estilo fino que el resto de la tabla y solo el borde inferior en 'medium' (grueso).
_BORDER_CLASS_SEPARATOR = BorderSpec(
    left="thin", right="thin", top="thin", bottom="medium", color=_BORDER_COLOR
)

# Fuente oficial de la tabla de Certificados (confirmado: Zurich Cn BT 11).
_FONT_CERTIFICADO_NAME = "Zurich Cn BT"
_FONT_CERTIFICADO_SIZE = 11

# Toda la información de la tabla de Certificados va centrada.
_ALIGN_CENTER_CERT = '<alignment horizontal="center" vertical="center"/>'


@dataclass
class DirectUpdateReport:
    """Informe detallado de la actualización directa en archivos oficiales."""
    timestamp: str
    backups_created: List[Path] = field(default_factory=list)
    # Respaldos antiguos (con marca de tiempo) consolidados en el backup único.
    backups_consolidados: int = 0
    sistematizacion_updated_cells: int = 0
    sistematizacion_names_uppercased: int = 0
    sistematizacion_untouched_cells: int = 0
    # Estudiantes cuyo archivo de Notas no traía nota final (no se inventa promedio)
    sistematizacion_promedios_no_copiados: int = 0
    certificados_rows_appended: int = 0
    # Ubicación exacta de lo anexado, para que el usuario pueda comprobarlo:
    certificados_hoja: str = ""
    certificados_fila_inicial: int = 0
    certificados_fila_final: int = 0
    certificados_n_inicial: int = 0
    certificados_n_final: int = 0
    # Relectura del archivo recién escrito: cuántas filas se confirmaron.
    certificados_rows_verificados: int = 0
    # Certificados del mismo curso ya presentes (posible cierre repetido).
    certificados_previos: int = 0
    # Aprobados que ya estaban registrados y por tanto NO se duplicaron.
    certificados_omitidos_duplicados: int = 0
    control_aulas_classes_marked: int = 0
    # Clases que ya tenían fecha de cierre (no se tocaron) o que no se hallaron.
    control_aulas_classes_already_marked: int = 0
    control_aulas_classes_not_found: List[str] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    # Verificación de preservación por archivo (PreservationReport.resumen()):
    # listas de validación antes/después + integridad estructural del paquete.
    validaciones_preservadas: Dict[str, str] = field(default_factory=dict)
    # Hojas modificadas a propósito en cada archivo (clave: etiqueta del informe).
    hojas_modificadas: Dict[str, set] = field(default_factory=dict)

    @property
    def is_successful(self) -> bool:
        return len(self.errors) == 0


def _strip_pivot_tables(wb: openpyxl.Workbook) -> int:
    """Elimina todas las tablas dinámicas (PivotTables) de un workbook en memoria.

    **Utilidad de legado.** El cierre directo ya no la usa: al no reserializar el
    libro con openpyxl, las tablas dinámicas se conservan intactas. Se mantiene
    porque hay utilidades/tests que la usan sobre libros propios.

    openpyxl no soporta PivotTables y produce XML inválido al guardar archivos que
    las contienen, lo que provoca el mensaje de "reparación" de Excel.

    Args:
        wb: El workbook a sanear (se modifica en lugar).

    Returns:
        Número de tablas dinámicas eliminadas.
    """
    count = 0
    for ws in wb.worksheets:
        if hasattr(ws, '_pivots') and ws._pivots:
            count += len(ws._pivots)
            ws._pivots.clear()
    return count


# Detecta referencias a libros externos en fórmulas (ej. "[1]ABBUEI!$G:$G").
_EXTERNAL_REF_RE = re.compile(r"\[\d+\]")


def _strip_external_links(wb: openpyxl.Workbook) -> int:
    """Elimina las referencias a fórmulas externas (externalLinks) de un workbook.

    **Utilidad de legado.** El cierre directo ya no la usa: como no se reserializa
    el libro, los vínculos externos se conservan tal como están en el original.

    openpyxl no preserva correctamente la caché de valores de las referencias
    externas (`externalLink*.xml`), lo que al guardar produce un XML inconsistente
    y Excel muestra el aviso de "reparación". Al eliminar las referencias externas,
    dichos archivos dejan de escribirse y el libro abre limpio.

    Las fórmulas que apuntaban a otros libros pueden quedar sin referencia (Excel
    las marca como no disponibles), pero se conservan los valores ya calculados en
    las celdas que sí se escriben. El backup previo garantiza el original íntegro.

    Args:
        wb: El workbook a sanear (se modifica en el lugar).

    Returns:
        Número de referencias externas eliminadas.
    """
    count = 0
    if hasattr(wb, "_external_links") and wb._external_links:
        count = len(wb._external_links)
        wb._external_links.clear()

    # Eliminar nombres definidos que apunten a libros externos (ej. "[1]Hoja!$A:$A")
    # para que no queden referencias colgantes tras quitar los externalLinks.
    try:
        for name, defined in list(wb.defined_names.items()):
            value = getattr(defined, "value", None) or getattr(defined, "attr_text", None) or ""
            if isinstance(value, str) and _EXTERNAL_REF_RE.search(value):
                wb.defined_names.pop(name, None)
                count += 1
    except Exception:
        pass

    return count


def _resolve_external_link_formulas(wb: openpyxl.Workbook, source_path: Path | str) -> int:
    """Sustituye fórmulas que referencian libros externos por su valor en caché.

    **Utilidad de legado.** El cierre directo ya no la usa: las fórmulas con
    referencias externas se conservan intactas.

    Al guardar con openpyxl y eliminar las referencias externas, las fórmulas que
    apuntaban a otros libros (ej. `=+_xlfn.XLOOKUP(AJ79,[1]ABBUEI!$G:$G,...)`)
    quedan con referencias colgantes y Excel las elimina al "reparar" el libro.

    Para evitarlo, se carga el mismo archivo con `data_only=True` (que devuelve el
    último valor calculado por Excel) y se reemplaza cada fórmula con referencia
    externa por dicho valor.

    Args:
        wb: Workbook cargado con `data_only=False` (se modifica en el lugar).
        source_path: Ruta del archivo fuente para leer los valores en caché.

    Returns:
        Número de celdas con fórmula externa sustituidas.
    """
    try:
        wb_values = openpyxl.load_workbook(source_path, data_only=True, read_only=True)
    except Exception:
        return 0

    replaced = 0
    try:
        for ws in wb.worksheets:
            if ws.title not in wb_values.sheetnames:
                continue
            ws_values = wb_values[ws.title]
            for row in ws.iter_rows():
                for cell in row:
                    val = cell.value
                    if isinstance(val, str) and val.startswith("=") and _EXTERNAL_REF_RE.search(val):
                        cached = ws_values.cell(cell.row, cell.column).value
                        cell.value = cached
                        replaced += 1
    finally:
        wb_values.close()

    return replaced


# ---------------------------------------------------------------------------
# PRESERVACIÓN DE LISTAS DE VALIDACIÓN DE DATOS Y EXTENSIONES DE HOJA
# ---------------------------------------------------------------------------
# openpyxl no soporta (y por tanto descarta al guardar):
#   • Las extensiones de hoja (`extLst`), donde Excel 2010+ guarda las listas de
#     validación modernas (`x14:dataValidations`) y el formato condicional moderno.
#   • En algunos escenarios, el bloque estándar `<dataValidations>`.
#   • Nombres definidos del libro (`definedNames`) que son el origen de muchas
#     listas desplegables (p. ej. `formula1="ListaSI"`).
#
# Estrategia (a prueba de pérdidas):
#   1. Se extraen los bloques del respaldo íntegro del archivo original, que es la
#      fuente de verdad: openpyxl sólo refleja lo que leyó, nunca añade listas.
#   2. Se REESCRIBEN en el archivo guardado (reemplazando lo que haya escrito
#      openpyxl, que normaliza atributos). No se combinan por texto: hacerlo
#      duplicaba las listas (3→6, 7→14) porque openpyxl reescribe cada
#      `<dataValidation>` (agrega `showDropDown="0"`, reordena atributos, etc.).
#   3. Se reinyectan respetando el orden estricto del esquema OOXML y declarando
#      los namespaces necesarios en la raíz del worksheet (incluidas las URIs
#      conocidas de Excel si la declaración se perdió).
#   4. Se VERIFICA huella por huella (tipo + rango + fórmula) que ninguna validación
#      del original se haya perdido; si algo no cuadra se informa explícitamente
#      al usuario en lugar de fallar en silencio.

# Bloque <extLst> de una hoja de cálculo (no anidado). Acepta la forma autocerrada.
_SHEET_EXTLST_RE = re.compile(r"<extLst\b[^>]*>.*?</extLst>|<extLst\b[^>]*/>", re.DOTALL)
# Bloque <ext ...>...</ext> individual dentro de un extLst.
_EXT_BLOCK_RE = re.compile(r"<ext\b[^>]*>.*?</ext>", re.DOTALL)
# Atributo uri="..." de un <ext ...>
_EXT_URI_RE = re.compile(r'<ext\b[^>]*\buri="([^"]+)"')
# Declaración de namespace con prefijo: xmlns:prefijo="uri"
_NS_DECL_RE = re.compile(r'\bxmlns:([A-Za-z_][\w.\-]*)\s*=\s*["\']([^"\']+)["\']')
# Bloque <dataValidations>...</dataValidations> (listas de validación estándar).
# Se admite también la forma autocerrada `<dataValidations count="0"/>`.
_DATA_VALIDATIONS_RE = re.compile(
    r"<dataValidations\b[^>]*>.*?</dataValidations>|<dataValidations\b[^>]*/>", re.DOTALL
)
# Elemento <dataValidation> individual (contenido o autocerrado).
# No confunde `x14:dataValidation` porque exige `<dataValidation` tras el '<'.
_DATA_VALIDATION_RE = re.compile(
    r"<dataValidation\b[^>]*/>|<dataValidation\b[^>]*>.*?</dataValidation>", re.DOTALL
)
# Equivalente moderno dentro de `extLst` (`x14:dataValidation`).
_X14_VALIDATION_RE = re.compile(
    r"<x14:dataValidation\b[^>]*/>|<x14:dataValidation\b[^>]*>.*?</x14:dataValidation>",
    re.DOTALL,
)
# Elemento <definedName> individual del libro.
_DEFINED_NAME_RE = re.compile(
    r"<definedName\b[^>]*/>|<definedName\b[^>]*>.*?</definedName>", re.DOTALL
)
# Bloque <definedNames>...</definedNames> del libro.
_DEFINED_NAMES_BLOCK_RE = re.compile(
    r"<definedNames\b[^>]*>.*?</definedNames>|<definedNames\b[^>]*/>", re.DOTALL
)

# Elementos que, según el esquema CT_Worksheet, van DESPUÉS de <dataValidations>.
# IMPORTANTE: no agregar elementos que vayan ANTES (mergeCells, conditionalFormatting,
# autoFilter, etc.); hacerlo produciría un XML fuera de orden de esquema.
_AFTER_DATA_VALIDATIONS = (
    "<hyperlinks", "<printOptions", "<pageMargins", "<pageSetup", "<headerFooter",
    "<rowBreaks", "<colBreaks", "<customProperties", "<cellWatches", "<ignoredErrors",
    "<smartTags", "<drawing", "<legacyDrawing", "<legacyDrawingHF", "<picture",
    "<oleObjects", "<controls", "<webPublishItems", "<tableParts", "<extLst",
)

# Elementos que, según el esquema CT_Workbook, van DESPUÉS de <definedNames>.
_AFTER_DEFINED_NAMES = (
    "<calcPr", "<oleSize", "<customWorkbookViews", "<pivotCaches", "<smartTagPr",
    "<smartTagTypes", "<webPublishing", "<fileRecoveryPr", "<webPublishObjects", "<extLst",
)


def _replace_or_insert_block(sheet_xml: str, regex: "re.Pattern", block: str, after_tags) -> str:
    """Reemplaza un bloque XML si existe, o lo inserta en la posición correcta del esquema.

    Si el bloque no existe, se inserta antes del primero de `after_tags` (elementos
    que van después según el esquema OOXML); si ninguno existe, antes de </worksheet>.
    """
    m = regex.search(sheet_xml)
    if m:
        return sheet_xml[:m.start()] + block + sheet_xml[m.end():]
    pos = len(sheet_xml)
    for tag in after_tags:
        i = sheet_xml.find(tag)
        if i != -1 and i < pos:
            pos = i
    if pos == len(sheet_xml):
        i = sheet_xml.rfind("</worksheet>")
        if i != -1:
            pos = i
    return sheet_xml[:pos] + block + sheet_xml[pos:]


def _collect_ns_decls(xml: str) -> Dict[str, str]:
    """Recolecta las declaraciones {prefijo: uri} de un fragmento XML."""
    return {m.group(1): m.group(2) for m in _NS_DECL_RE.finditer(xml)}


# Prefijos usados en nombres de elementos (<x14:foo>) y de atributos (x14:foo="...").
# Reglas para no confundir datos con prefijos:
#   • Un elemento con prefijo siempre empieza con '<prefijo:'.
#   • Un atributo con prefijo va precedido de espacio y SIEMPRE lleva '=' tras el
#     nombre local. Sin exigir el '=' se confundían rangos de celdas de un `sqref`
#     multirrango (p. ej. "B2:B9 AD2:AF2") con prefijos de namespace, lo que hacía
#     que se descartara el `extLst` completo y se perdieran las listas modernas.
_ELEM_PREFIX_RE = re.compile(r"</?([A-Za-z_][\w.\-]*):[A-Za-z_]")
_ATTR_PREFIX_RE = re.compile(r"\s([A-Za-z_][\w.\-]*):([A-Za-z_][\w.\-]*)\s*=")

# Prefijos implícitos del estándar XML: siempre están declarados.
_IMPLICIT_PREFIXES = frozenset({"xml", "xmlns"})

# URIs de los prefijos usados por Excel cuando las validaciones modernas no traen
# la declaración `xmlns:` (o cuando se perdió al guardar). Sirven de red de
# seguridad para no descartar un bloque por "prefijo no declarado".
_KNOWN_NS_URIS = {
    "x14": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
    "x14ac": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
    "x15": "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main",
    "x15ac": "http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac",
    "xm": "http://schemas.microsoft.com/office/excel/2006/main",
    "xr": "http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
    "xr2": "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2",
    "xr3": "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "a14": "http://schemas.microsoft.com/office/drawing/2010/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def _prefixes_used(xml: str) -> set:
    """Devuelve el conjunto de prefijos de namespace usados en un fragmento XML."""
    prefixes = {m.group(1) for m in _ELEM_PREFIX_RE.finditer(xml)}
    prefixes |= {m.group(1) for m in _ATTR_PREFIX_RE.finditer(xml)}
    return prefixes - _IMPLICIT_PREFIXES


def _complete_ns_pool(ns_pool: Dict[str, str], required: set) -> Dict[str, str]:
    """Completa el pool de namespaces con las URIs conocidas de Excel.

    Permite reinyectar un bloque cuyos prefijos (`x14`, `xm`, `xr`…) no vinieran
    declarados en el archivo original, evitando descartar la lista por completo.
    """
    completed = dict(ns_pool)
    for prefix in required:
        if prefix not in completed and prefix in _KNOWN_NS_URIS:
            completed[prefix] = _KNOWN_NS_URIS[prefix]
    return completed


def _ensure_ns_decls(sheet_xml: str, ns_decls: Dict[str, str]) -> str:
    """Agrega a la etiqueta <worksheet> las declaraciones de namespace faltantes.

    Necesario cuando el `extLst` reinyectado usa prefijos (p. ej. `x14:`, `xm:`)
    declarados en la raíz del worksheet original, que openpyxl no conserva. Sin
    esto, Excel reporta "Prefijo no declarado".
    """
    if not ns_decls:
        return sheet_xml
    m = re.search(r"<worksheet\b[^>]*>", sheet_xml)
    if not m:
        return sheet_xml
    root_tag = m.group(0)
    existing = _collect_ns_decls(root_tag)
    missing = {p: u for p, u in ns_decls.items() if p not in existing}
    if not missing:
        return sheet_xml
    decls = "".join(f' xmlns:{p}="{u}"' for p, u in missing.items())
    new_tag = root_tag[:-1] + decls + ">"
    return sheet_xml.replace(root_tag, new_tag, 1)


def _read_sheet_names_by_path(zf: "zipfile.ZipFile") -> Dict[str, str]:
    """Mapea el nombre de cada hoja a su ruta interna (p. ej. 'Formato' -> 'xl/worksheets/sheet2.xml').

    El mapeo se obtiene de `xl/workbook.xml` (nombre + r:id) y de
    `xl/_rels/workbook.xml.rels` (r:id -> destino).

    Args:
        zf: Archivo .xlsx abierto como zip en modo lectura.

    Returns:
        Diccionario {nombre_hoja: ruta_dentro_del_zip}.
    """
    mapping: Dict[str, str] = {}
    try:
        workbook_xml = zf.read("xl/workbook.xml").decode("utf-8", "ignore")
        rels_xml = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8", "ignore")
    except KeyError:
        return mapping

    # rId -> Target (ruta relativa a xl/)
    rel_targets: Dict[str, str] = {}
    for m in re.finditer(r'<Relationship\b[^>]*>', rels_xml):
        tag = m.group(0)
        rid_m = re.search(r'Id="([^"]+)"', tag)
        target_m = re.search(r'Target="([^"]+)"', tag)
        if rid_m and target_m:
            target = target_m.group(1).lstrip("/")
            if not target.startswith("xl/"):
                target = "xl/" + target
            rel_targets[rid_m.group(1)] = target

    # sheet name + r:id
    for m in re.finditer(r'<sheet\b[^>]*>', workbook_xml):
        tag = m.group(0)
        name_m = re.search(r'name="([^"]+)"', tag)
        rid_m = re.search(r'r:id="([^"]+)"', tag)
        if name_m and rid_m and rid_m.group(1) in rel_targets:
            mapping[name_m.group(1)] = rel_targets[rid_m.group(1)]

    return mapping


@dataclass
class SheetBlocksReport:
    """Resultado de la preservación de listas de validación de un archivo oficial.

    Permite informar al usuario, de forma verificable, cuántas listas desplegables
    (validaciones de datos) existían antes y cuántas quedaron después de aplicar el
    cierre, además de cualquier problema detectado durante la restauración.
    """

    archivo: str = ""
    validaciones_original: int = 0
    validaciones_resultado: int = 0
    validaciones_x14_original: int = 0
    validaciones_x14_resultado: int = 0
    hojas_con_validaciones: int = 0
    nombres_definidos_restaurados: int = 0
    problemas: List[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        """True si no se detectó ninguna pérdida de listas de validación."""
        return not self.problemas

    def resumen(self) -> str:
        """Resumen de una línea, apto para la consola del Modo 5."""
        if not self.validaciones_original and not self.validaciones_x14_original:
            return f"✅ {self.archivo}: sin listas de validación de datos que preservar."
        icono = "✅" if self.ok else "⚠️"
        extra = ""
        if self.nombres_definidos_restaurados:
            extra = f" | nombres definidos restaurados: {self.nombres_definidos_restaurados}"
        return (
            f"{icono} {self.archivo}: listas de validación estándar "
            f"{self.validaciones_original}→{self.validaciones_resultado}, "
            f"modernas (x14) {self.validaciones_x14_original}→{self.validaciones_x14_resultado}"
            f"{extra}"
        )


class _MemZip:
    """Adaptador mínimo para reutilizar `_read_sheet_names_by_path` sobre bytes en memoria."""

    def __init__(self, parts: Dict[str, bytes]):
        self._parts = parts

    def read(self, name: str) -> bytes:
        return self._parts[name]


def _read_sheet_names_by_path_dict(members: Dict[str, bytes]) -> Dict[str, str]:
    """Igual que `_read_sheet_names_by_path` pero sobre un dict de partes ya leídas."""
    return _read_sheet_names_by_path(_MemZip(members))


def _normalize_xml_fragment(fragment: str) -> str:
    """Huella estable de un fragmento XML (insensible a espacios y a `xr:uid`)."""
    txt = re.sub(r'\s+xr:uid="[^"]*"', "", fragment)
    return re.sub(r"\s+", " ", txt).strip()


def _validation_fingerprints(sheet_xml: str) -> set:
    """Conjunto de validaciones de datos presentes en el XML de una hoja.

    Incluye tanto las estándar (`<dataValidation>`) como las modernas
    (`<x14:dataValidation>` dentro de `extLst`). Cada huella conserva el rango
    (`sqref`), el tipo y las fórmulas, de modo que dos listas distintas nunca se
    confunden entre sí.
    """
    items = _DATA_VALIDATION_RE.findall(sheet_xml) + _X14_VALIDATION_RE.findall(sheet_xml)
    return {_normalize_xml_fragment(item) for item in items}


def _is_well_formed(xml: str) -> bool:
    """True si el XML se puede parsear (sin prefijos sin declarar, sin errores)."""
    try:
        ElementTree.fromstring(xml)
        return True
    except Exception:  # noqa: BLE001 - cualquier error de parseo invalida el parche
        return False


def _build_data_validations_block(source_block: Optional[str]) -> Optional[str]:
    """Devuelve el bloque `<dataValidations>` íntegro del archivo original (respaldo).

    Se escribe el bloque del respaldo **tal cual** (es la fuente de verdad) y se
    reemplaza cualquier `<dataValidations>` que haya escrito openpyxl.

    Importante: NO se combinan las validaciones del archivo guardado con las del
    respaldo. openpyxl reescribe los atributos de cada `<dataValidation>` (añade
    `showDropDown="0"`, `showInputMessage="0"`, reordena, etc.), por lo que
    compararlas por texto las consideraba distintas y se duplicaban (3→6, 7→14).
    openpyxl nunca añade validaciones nuevas, sólo refleja las que leyó, de modo
    que el bloque del original siempre es suficiente y exacto.

    Args:
        source_block: Bloque `<dataValidations>` del archivo original (respaldo).

    Returns:
        El bloque XML a escribir, o None si el original no tenía validaciones.
    """
    if not source_block:
        return None
    if not _DATA_VALIDATION_RE.search(source_block):
        # Bloques vacíos (`<dataValidations count="0"/>`): nada que restaurar.
        return None
    return source_block.strip()


def _apply_preserved_blocks(
    sheet_xml: str,
    dv_block: Optional[str],
    ext_block: Optional[str],
    ns_pool: Dict[str, str],
) -> Tuple[str, List[str]]:
    """Reinyecta en el XML de una hoja los bloques de validación y extensiones.

    Args:
        sheet_xml: XML de la hoja guardada por openpyxl.
        dv_block: Bloque `<dataValidations>` del original (o None).
        ext_block: Bloque `<extLst>` del original (o None).
        ns_pool: Declaraciones `{prefijo: uri}` disponibles para resolver prefijos.

    Returns:
        Tupla (XML resultante, lista de problemas detectados).
    """
    problemas: List[str] = []
    new_xml = sheet_xml

    if dv_block:
        block = _build_data_validations_block(dv_block)
        if block:
            new_xml = _replace_or_insert_block(
                new_xml, _DATA_VALIDATIONS_RE, block, _AFTER_DATA_VALIDATIONS
            )

    if ext_block:
        # Se completan los namespaces faltantes con las URIs conocidas de Excel
        # para no descartar el bloque por un prefijo no declarado.
        ns_pool = _complete_ns_pool(ns_pool, _prefixes_used(ext_block))
        unknown = _prefixes_used(ext_block) - set(ns_pool)
        if unknown:
            problemas.append(
                "no se pudieron declarar los prefijos de namespace "
                + ", ".join(sorted(unknown))
            )
        else:
            candidate = _merge_extlst_into_sheet_xml(new_xml, ext_block)
            if candidate is None:
                problemas.append("no se pudo reinyectar el bloque extLst")
            else:
                new_xml = candidate

    if new_xml != sheet_xml:
        new_xml = _ensure_ns_decls(new_xml, ns_pool)

    return new_xml, problemas


def _defined_name_key(item: str) -> Tuple[str, str]:
    """Clave de identidad de un `<definedName>`: (nombre en mayúsculas, hoja local)."""
    name_m = re.search(r'\bname="([^"]*)"', item)
    sheet_m = re.search(r'\blocalSheetId="([^"]*)"', item)
    return (
        (name_m.group(1) if name_m else "").upper(),
        sheet_m.group(1) if sheet_m else "",
    )


def _insert_workbook_block(workbook_xml: str, block: str, after_tags) -> str:
    """Inserta un bloque en `xl/workbook.xml` respetando el orden del esquema OOXML."""
    pos = len(workbook_xml)
    for tag in after_tags:
        i = workbook_xml.find(tag)
        if i != -1 and i < pos:
            pos = i
    if pos == len(workbook_xml):
        i = workbook_xml.rfind("</workbook>")
        if i != -1:
            pos = i
    return workbook_xml[:pos] + block + workbook_xml[pos:]


def _restore_defined_names(
    output_workbook_xml: Optional[str],
    source_workbook_xml: Optional[str],
) -> Tuple[Optional[str], int, List[str]]:
    """Reincorpora al libro guardado los nombres definidos del original que falten.

    Muchas listas desplegables usan un nombre definido como origen
    (`<formula1>ListaSI</formula1>`). Si el nombre desaparece del libro, Excel deja
    de mostrar la lista, por lo que se restauran los que falten sin eliminar ni
    duplicar ninguno y sin reintroducir referencias a libros externos.

    Returns:
        Tupla (XML resultante, cantidad restaurada, lista de problemas).
    """
    if not output_workbook_xml or not source_workbook_xml:
        return output_workbook_xml, 0, []

    problemas: List[str] = []
    try:
        source_block_m = _DEFINED_NAMES_BLOCK_RE.search(source_workbook_xml)
        if not source_block_m:
            return output_workbook_xml, 0, []
        source_items = _DEFINED_NAME_RE.findall(source_block_m.group(0))
        if not source_items:
            return output_workbook_xml, 0, []

        target_block_m = _DEFINED_NAMES_BLOCK_RE.search(output_workbook_xml)
        target_block = target_block_m.group(0) if target_block_m else ""
        target_items = _DEFINED_NAME_RE.findall(target_block)

        existing = {_defined_name_key(item) for item in target_items}
        restored: List[str] = []
        for item in source_items:
            if _EXTERNAL_REF_RE.search(item):
                # Nunca se reintroducen nombres apuntando a libros externos:
                # quedarían colgantes y Excel pediría "reparar" el libro.
                continue
            key = _defined_name_key(item)
            if not key[0] or key in existing:
                continue
            existing.add(key)
            restored.append(item)

        if not restored:
            return output_workbook_xml, 0, problemas

        if target_block and not target_block.endswith("/>"):
            merged_block = target_block.replace(
                "</definedNames>", "".join(restored) + "</definedNames>", 1
            )
            return output_workbook_xml.replace(target_block, merged_block, 1), len(restored), problemas

        merged_block = "<definedNames>%s</definedNames>" % "".join(target_items + restored)
        if target_block:
            return output_workbook_xml.replace(target_block, merged_block, 1), len(restored), problemas

        return _insert_workbook_block(output_workbook_xml, merged_block, _AFTER_DEFINED_NAMES), len(restored), problemas
    except Exception as exc:  # noqa: BLE001 - nunca debe romper el cierre
        problemas.append(f"no se pudieron restaurar los nombres definidos: {exc}")
        return output_workbook_xml, 0, problemas


def _read_zip_parts(path: Path) -> Tuple[Dict[str, bytes], Dict[str, "zipfile.ZipInfo"]]:
    """Lee todas las partes de un .xlsx (contenido + metadatos de cada entrada)."""
    with zipfile.ZipFile(path, "r") as zf:
        members = {name: zf.read(name) for name in zf.namelist()}
        infos = {info.filename: info for info in zf.infolist()}
    return members, infos


def _write_zip_parts(
    path: Path,
    members: Dict[str, bytes],
    infos: Dict[str, "zipfile.ZipInfo"],
) -> None:
    """Reescribe un .xlsx preservando el orden y los metadatos de sus partes."""
    import os as _os
    import tempfile as _tempfile

    fd, tmp_name = _tempfile.mkstemp(suffix=".xlsx", dir=str(path.parent))
    _os.close(fd)
    tmp_path = Path(tmp_name)
    try:
        with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zf_new:
            for name, data in members.items():
                info = infos.get(name)
                if info is not None:
                    zf_new.writestr(info, data)
                else:
                    zf_new.writestr(name, data)
        _os.replace(tmp_path, path)
    finally:
        if tmp_path.exists():
            try:
                tmp_path.unlink()
            except OSError:
                pass


def preserve_sheet_data_blocks(
    output_path: Path | str,
    source_path: Path | str,
) -> SheetBlocksReport:
    """Restaura y VERIFICA las listas de validación de datos del archivo original.

    **Utilidad de legado.** El cierre directo (Modo 5) ya no la necesita: en lugar
    de guardar con openpyxl y reponer bloques a posteriori, `XlsxPatcher` edita
    solo las celdas y conserva el paquete completo. Se mantiene disponible para
    reparar libros que sí hayan pasado por openpyxl.

    Debe invocarse inmediatamente después de `wb.save(...)`, usando como
    `source_path` la copia de seguridad íntegra creada antes de la modificación.

    Restaura, hoja por hoja:
      • El bloque `<dataValidations>` (listas de validación estándar).
      • El bloque `<extLst>` (validaciones modernas `x14:dataValidations`).
    Y a nivel de libro:
      • Los `<definedNames>` que sirven de origen a algunas listas.

    Finalmente compara huella por huella (rango + tipo + fórmula) que ninguna
    validación del original se haya perdido. Si detecta pérdidas reintenta con una
    estrategia "dura" y, si aun así falla, lo reporta en `problemas`.

    Args:
        output_path: Ruta del archivo recién guardado por openpyxl (se reescribe).
        source_path: Ruta del original íntegro (respaldo) con las validaciones.

    Returns:
        SheetBlocksReport con los conteos antes/después y los problemas detectados.
    """
    out_path = Path(output_path)
    src_path = Path(source_path)
    report = SheetBlocksReport(archivo=out_path.name)

    if not out_path.exists() or not src_path.exists():
        report.problemas.append(
            "No se pudo verificar la preservación de listas de validación "
            "(falta el archivo guardado o su respaldo)."
        )
        return report

    # ---------------------------------------------------------
    # 1. Extraer bloques y huellas del original (respaldo íntegro)
    # ---------------------------------------------------------
    try:
        with zipfile.ZipFile(src_path, "r") as zf_src:
            source_workbook_xml = zf_src.read("xl/workbook.xml").decode("utf-8", "ignore")
            source_map = _read_sheet_names_by_path(zf_src)
            # Declaraciones de namespace del libro (algunas extensiones usan prefijos
            # declarados sólo en la raíz de xl/workbook.xml, p. ej. `xr`).
            workbook_ns = _collect_ns_decls(source_workbook_xml)
            blocks: Dict[str, Dict[str, Any]] = {}
            for sheet_name, sheet_path in source_map.items():
                try:
                    src_xml = zf_src.read(sheet_path).decode("utf-8", "ignore")
                except KeyError:
                    continue
                dv_m = _DATA_VALIDATIONS_RE.search(src_xml)
                ext_m = _SHEET_EXTLST_RE.search(src_xml)
                if not dv_m and not ext_m:
                    continue
                ns_pool = dict(workbook_ns)
                ns_pool.update(_collect_ns_decls(src_xml))
                blocks[sheet_name] = {
                    "dv": dv_m.group(0) if dv_m else None,
                    "ext": ext_m.group(0) if ext_m else None,
                    "ns": ns_pool,
                    "fingerprints": _validation_fingerprints(src_xml),
                }
    except Exception as exc:  # noqa: BLE001
        report.problemas.append(f"No se pudo leer el respaldo '{src_path.name}': {exc}")
        return report

    report.validaciones_original = sum(
        len(_DATA_VALIDATION_RE.findall(blk["dv"] or "")) for blk in blocks.values()
    )
    report.validaciones_x14_original = sum(
        len(_X14_VALIDATION_RE.findall(blk["ext"] or "")) for blk in blocks.values()
    )
    report.hojas_con_validaciones = sum(1 for blk in blocks.values() if blk["fingerprints"])

    if not blocks:
        return report

    # ---------------------------------------------------------
    # 2. Aplicar la restauración sobre el archivo guardado
    # ---------------------------------------------------------
    try:
        out_members, out_infos = _read_zip_parts(out_path)
    except Exception as exc:  # noqa: BLE001
        report.problemas.append(f"No se pudo leer el archivo guardado '{out_path.name}': {exc}")
        return report

    out_map = _read_sheet_names_by_path_dict(out_members)

    def _build_candidate() -> Tuple[Dict[str, bytes], List[str], int]:
        """Genera una versión candidata del archivo con los bloques reinyectados."""
        candidate = dict(out_members)
        problemas: List[str] = []
        restored_names = 0

        for sheet_name, blk in blocks.items():
            sheet_path = out_map.get(sheet_name)
            if not sheet_path or sheet_path not in candidate:
                problemas.append(
                    f"'{sheet_name}': la hoja no existe en el archivo guardado; "
                    "no se pudieron restaurar sus listas de validación"
                )
                continue
            current = candidate[sheet_path].decode("utf-8", "ignore")
            new_xml, sheet_problems = _apply_preserved_blocks(
                current, blk["dv"], blk["ext"], blk["ns"]
            )
            problemas.extend(f"'{sheet_name}': {p}" for p in sheet_problems)
            if new_xml != current:
                if _is_well_formed(new_xml):
                    candidate[sheet_path] = new_xml.encode("utf-8")
                else:
                    problemas.append(
                        f"'{sheet_name}': el XML resultante no es válido; "
                        "se conserva la versión guardada por openpyxl"
                    )

        wb_key = "xl/workbook.xml"
        if wb_key in candidate:
            current_wb = candidate[wb_key].decode("utf-8", "ignore")
            new_wb, restored_names, wb_problems = _restore_defined_names(
                current_wb, source_workbook_xml
            )
            problemas.extend(wb_problems)
            if new_wb is not None and new_wb != current_wb:
                if _is_well_formed(new_wb):
                    candidate[wb_key] = new_wb.encode("utf-8")
                else:
                    problemas.append(
                        "'xl/workbook.xml': los nombres definidos no se pudieron restaurar "
                        "(XML resultante inválido)"
                    )
                    restored_names = 0

        # Verificación huella por huella: ninguna validación debe faltar.
        for sheet_name, blk in blocks.items():
            sheet_path = out_map.get(sheet_name)
            if not sheet_path or sheet_path not in candidate:
                continue
            result_xml = candidate[sheet_path].decode("utf-8", "ignore")
            missing = blk["fingerprints"] - _validation_fingerprints(result_xml)
            if missing:
                problemas.append(
                    f"'{sheet_name}': {len(missing)} validación(es) de datos "
                    "no quedaron en el archivo final"
                )
        return candidate, problemas, restored_names

    candidate, problemas, nombres_restaurados = _build_candidate()

    report.nombres_definidos_restaurados = nombres_restaurados
    # Se eliminan mensajes repetidos conservando el orden original.
    report.problemas.extend(dict.fromkeys(problemas))

    # ---------------------------------------------------------
    # 3. Escribir el resultado (solo si hubo cambios) y devolver conteos
    # ---------------------------------------------------------
    sheets_involved = {out_map[name] for name in blocks if name in out_map}
    report.validaciones_resultado = sum(
        len(_DATA_VALIDATION_RE.findall(candidate[p].decode("utf-8", "ignore")))
        for p in sheets_involved
        if p in candidate
    )
    report.validaciones_x14_resultado = sum(
        len(_X14_VALIDATION_RE.findall(candidate[p].decode("utf-8", "ignore")))
        for p in sheets_involved
        if p in candidate
    )

    if any(candidate.get(k) != out_members.get(k) for k in set(candidate) | set(out_members)):
        _write_zip_parts(out_path, candidate, out_infos)

    return report


def _restore_sheet_blocks(
    output_path: Path | str,
    source_path: Path | str,
) -> tuple:
    """Compatibilidad histórica: devuelve (validaciones estándar, validaciones x14).

    Mantiene la firma anterior para no romper llamadas existentes; toda la lógica e
    información detallada vive en :func:`preserve_sheet_data_blocks`.
    """
    report = preserve_sheet_data_blocks(output_path, source_path)
    return (report.validaciones_resultado, report.validaciones_x14_resultado)


def _merge_extlst_into_sheet_xml(sheet_xml: str, source_extlst: str) -> Optional[str]:
    """Combina el `extLst` original dentro del XML de una hoja guardada por openpyxl.

    - Si la hoja guardada no tiene `extLst`, se inserta el original antes de `</worksheet>`.
    - Si ya tiene `extLst`, se añaden dentro los `<ext>` del original que no estén
      ya presentes (comparando por `uri`), para evitar duplicados.
    - Se admite la forma autocerrada `<extLst/>` escrita por algunos generadores.

    Returns:
        El XML resultante, o None si no se pudo procesar.
    """
    try:
        existing_m = _SHEET_EXTLST_RE.search(sheet_xml)
        source_exts = _EXT_BLOCK_RE.findall(source_extlst)

        if not existing_m:
            if "</worksheet>" not in sheet_xml:
                return None
            return sheet_xml.replace("</worksheet>", source_extlst + "</worksheet>", 1)

        if not source_exts:
            return sheet_xml

        existing_extlst = existing_m.group(0)
        existing_uris = {u.upper() for u in _EXT_URI_RE.findall(existing_extlst)}

        # Bloques <ext> del origen que aún no están presentes
        missing_exts = [
            block for block in source_exts
            if not any(u.upper() in existing_uris for u in _EXT_URI_RE.findall(block))
        ]
        if not missing_exts:
            return sheet_xml

        if existing_extlst.endswith("/>"):
            # <extLst/> vacío (autocerrado): se abre y se le añaden las extensiones.
            opened = existing_extlst[:-2].rstrip() + ">"
            merged_extlst = opened + "".join(missing_exts) + "</extLst>"
        else:
            merged_extlst = existing_extlst.replace(
                "</extLst>", "".join(missing_exts) + "</extLst>", 1
            )
        return sheet_xml.replace(existing_extlst, merged_extlst, 1)
    except Exception:  # noqa: BLE001
        return None


def backup_path_for(file_path: Path | str, backup_dir: Optional[Path | str] = None) -> Path:
    """Ruta del respaldo ÚNICO de un archivo oficial (sin marca de tiempo).

    El respaldo vive en ``_backups_cierre/<nombre>_backup.<ext>`` y se
    sobrescribe en cada cierre con el estado del archivo justo antes de la
    automatización: siempre hay exactamente un respaldo por archivo.
    """
    path = Path(file_path).resolve()
    target_dir = Path(backup_dir).resolve() if backup_dir else path.parent / "_backups_cierre"
    return target_dir / f"{path.stem}_backup{path.suffix}"


def find_legacy_backups(
    file_path: Path | str, backup_dir: Optional[Path | str] = None
) -> List[Path]:
    """Respaldos antiguos con marca de tiempo del mismo archivo (formato previo)."""
    path = Path(file_path).resolve()
    target_dir = Path(backup_dir).resolve() if backup_dir else path.parent / "_backups_cierre"
    if not target_dir.exists():
        return []
    pattern = f"{path.stem}_backup_*{path.suffix}"
    return sorted(target_dir.glob(pattern))


def create_timestamped_backup(file_path: Path | str, backup_dir: Optional[Path | str] = None) -> Path:
    """Crea o ACTUALIZA el respaldo único del archivo oficial.

    En cada cierre se reemplaza el respaldo existente por el estado actual del
    archivo (antes de la automatización), de modo que ``_backups_cierre/``
    conserva **un solo respaldo por archivo**. Se eliminan los respaldos
    antiguos con marca de tiempo del mismo archivo (formato previo), para no
    acumular copias.

    Args:
        file_path: Ruta del archivo original a respaldar.
        backup_dir: Carpeta de destino (si es None, crea '_backups_cierre' junto al archivo).

    Returns:
        Ruta al archivo de respaldo único.
    """
    path = Path(file_path).resolve()
    if not path.exists():
        raise FileNotFoundError(f"No se puede respaldar un archivo inexistente: {path}")

    backup_path = backup_path_for(path, backup_dir)
    backup_path.parent.mkdir(parents=True, exist_ok=True)

    # Respaldo atómico: si la copia falla, el respaldo anterior no se destruye.
    temporal = backup_path.with_name(backup_path.name + ".tmp")
    try:
        shutil.copy2(path, temporal)
        os.replace(temporal, backup_path)
    finally:
        if temporal.exists():
            try:
                temporal.unlink()
            except OSError:
                pass

    # Consolidar: se eliminan las copias con marca de tiempo del mismo archivo.
    for legacy in find_legacy_backups(path, backup_dir):
        if legacy == backup_path:
            continue
        try:
            legacy.unlink()
        except OSError:
            pass

    return backup_path


def find_latest_backup(original_path: Path | str, backup_dir: Optional[Path | str] = None) -> Optional[Path]:
    """Devuelve el respaldo único actual del archivo oficial.

    Si todavía no existe el respaldo único pero hay copias antiguas con marca de
    tiempo, devuelve la más reciente (migración transparente).

    Args:
        original_path: Ruta al archivo oficial original.
        backup_dir: Carpeta de backups (si es None, usa '_backups_cierre' junto al archivo).

    Returns:
        Ruta al backup, o None si no hay ninguno.
    """
    path = Path(original_path).resolve()
    target_dir = Path(backup_dir).resolve() if backup_dir else path.parent / "_backups_cierre"

    if not target_dir.exists():
        return None

    unico = target_dir / f"{path.stem}_backup{path.suffix}"
    if unico.exists():
        return unico

    # Compatibilidad con respaldos del formato anterior (con marca de tiempo).
    pattern = f"{path.stem}_backup_*{path.suffix}"
    candidates = sorted(target_dir.glob(pattern), reverse=True)
    return candidates[0] if candidates else None


def restore_last_backup(
    file_paths: Dict[str, Path | str],
    backup_dir: Optional[Path | str] = None,
) -> Dict[str, Any]:
    """Restaura el respaldo único de cada archivo oficial.

    El respaldo único contiene el estado del archivo **justo antes del último
    cierre** (se actualiza en cada ejecución). Solo restaura los archivos que
    tienen respaldo (Sistematización, Certificados y Control de Aulas). El
    archivo de Notas nunca se modifica, así que no tiene backup y se omite.

    Args:
        file_paths: Diccionario con rutas {'control_aulas', 'sistematizacion', 'certificados', ...}.
        backup_dir: Carpeta donde buscar los backups (por defecto '_backups_cierre').

    Returns:
        Diccionario con los resultados por archivo:
            {'role': {'status': 'ok'|'sin_backup'|'error', 'backup': Path|None, 'mensaje': str}}
    """
    resultados: Dict[str, Any] = {}
    roles_a_restaurar = ["sistematizacion", "certificados", "control_aulas"]

    for role in roles_a_restaurar:
        if role not in file_paths:
            resultados[role] = {"status": "omitido", "backup": None, "mensaje": "Archivo no especificado."}
            continue

        original = Path(file_paths[role]).resolve()
        # Inferir carpeta de backup desde el archivo original si no se especificó
        inferred_backup_dir = backup_dir if backup_dir else original.parent / "_backups_cierre"
        latest = find_latest_backup(original, backup_dir=inferred_backup_dir)

        if latest is None:
            resultados[role] = {
                "status": "sin_backup",
                "backup": None,
                "mensaje": f"No se encontró ningún backup para '{original.name}' en '{inferred_backup_dir}'.",
            }
            continue

        try:
            shutil.copy2(latest, original)
            resultados[role] = {
                "status": "ok",
                "backup": latest,
                "mensaje": f"'{original.name}' restaurado exitosamente desde '{latest.name}'.",
            }
        except Exception as exc:
            resultados[role] = {
                "status": "error",
                "backup": latest,
                "mensaje": f"Error al restaurar '{original.name}': {exc}",
            }

    return resultados


def _verify_preservation_and_report(
    report: "DirectUpdateReport",
    etiqueta: str,
    output_path: Path | str,
    source_path: Path | str,
    modified_sheets: Optional[set] = None,
) -> "PreservationReport":
    """Verifica que el cierre no haya perdido NINGUNA parte del libro oficial.

    Nunca lanza excepción: cualquier fallo se registra como advertencia en el
    informe para que el usuario lo vea en consola (y pueda restaurar el respaldo).

    Args:
        report: Informe de la actualización directa (se completa en el lugar).
        etiqueta: Nombre legible del archivo para el informe (ej. 'Sistematización').
        output_path: Archivo ya actualizado.
        source_path: Respaldo íntegro creado antes de la modificación.
        modified_sheets: Hojas que el cierre modificó a propósito.

    Returns:
        El `PreservationReport` con el detalle de la verificación.
    """
    try:
        info = verify_preservation(
            source_path,
            output_path,
            expected_modified_sheets=modified_sheets or set(),
        )
    except Exception as exc:  # noqa: BLE001 - la verificación nunca debe romper el cierre
        mensaje = (
            f"{etiqueta}: no se pudo verificar que el libro original quedara intacto ({exc})."
        )
        report.warnings.append(mensaje)
        return PreservationReport(archivo=Path(output_path).name, problemas=[mensaje])

    report.validaciones_preservadas[etiqueta] = info.resumen()
    report.hojas_modificadas[etiqueta] = set(modified_sheets or set())
    for problema in info.problemas:
        report.warnings.append(f"{etiqueta}: {problema}")
    return info


# ---------------------------------------------------------------------------
# Localización y verificación de la tabla de Certificados
# ---------------------------------------------------------------------------
# Tope defensivo de filas a inspeccionar (evita bucles eternos en hojas que
# declaran la columna completa como usada).
_MAX_FILAS_CERTIFICADOS = 200_000


def _inspect_certificate_table(ws) -> Tuple[int, int, bool]:
    """Localiza el final real de la tabla de Certificados y el último N° usado.

    **Por qué no basta con "última fila con algo en A o B":** las hojas
    institucionales suelen tener notas, títulos, totales o leyendas sueltas muy
    por debajo de la tabla. Si se anexa después de ellas, los certificados
    quedan a cientos o miles de filas del final visible de la tabla y el usuario
    concluye —con razón— que "no se guardaron".

    Un registro real de certificados se reconoce porque tiene **código del
    certificado (columna H)** o, al menos, **nombres y apellidos (D y E)**. Las
    notas y totales no cumplen esa condición.

    Args:
        ws: Hoja de openpyxl de Certificados.

    Returns:
        Tupla ``(ultima_fila_tabla, maximo_n, uso_respaldo)``. Si la hoja no
        tiene el formato esperado, ``uso_respaldo`` es True y se devuelve la
        última fila con contenido en A/B (comportamiento histórico).
    """
    max_r = min(ws.max_row or 1, _MAX_FILAS_CERTIFICADOS)
    last_row = 0
    max_n = 0
    last_row_any = 1

    for r in range(2, max_r + 1):
        val_n = ws.cell(r, 1).value
        val_b = ws.cell(r, 2).value
        nombre = str(ws.cell(r, 4).value or "").strip()
        apellido = str(ws.cell(r, 5).value or "").strip()
        codigo = str(ws.cell(r, 8).value or "").strip()

        if val_n is not None or val_b is not None:
            last_row_any = r

        es_registro = bool(codigo) or (bool(nombre) and bool(apellido))
        if not es_registro:
            continue

        last_row = r
        if val_n is not None:
            try:
                n_int = int(float(str(val_n).strip()))
                if n_int > max_n:
                    max_n = n_int
            except (ValueError, TypeError):
                pass

    if last_row:
        return last_row, max_n, False
    return last_row_any, max_n, last_row_any > 1


def _certificate_rows_by_document(ws, codigo_curso: str) -> Dict[str, int]:
    """Certificados del mismo curso ya registrados, indexados por documento.

    Sirve para que un segundo cierre del mismo curso (o una repetición por
    error) no duplique filas.
    """
    objetivo = str(codigo_curso or "").strip().upper()
    if not objetivo:
        return {}

    encontrados: Dict[str, int] = {}
    max_r = min(ws.max_row or 1, _MAX_FILAS_CERTIFICADOS)
    for r in range(2, max_r + 1):
        codigo = str(ws.cell(r, 8).value or "").strip().upper()
        if codigo != objetivo:
            continue
        doc = normalize_document(ws.cell(r, 6).value)
        if doc and doc not in encontrados:
            encontrados[doc] = r
    return encontrados


def _fila_certificados_vacia(ws, row: int, ncols: int = 13) -> bool:
    """True si ninguna de las columnas de la tabla tiene contenido en la fila."""
    for c in range(1, ncols + 1):
        val = ws.cell(row, c).value
        if val is not None and str(val).strip() != "":
            return False
    return True


def _merges_covering_rows(ws, filas: List[int]) -> List[str]:
    """Rangos combinados preexistentes que cubren filas anexadas.

    Escribir dentro de una celda combinada es invisible en Excel (solo se ve el
    valor de la celda superior izquierda del rango), así que conviene avisarlo:
    el certificado queda escrito pero parte de sus datos no se muestra.
    """
    cubiertos: List[str] = []
    try:
        rangos = list(ws.merged_cells.ranges)
    except Exception:  # noqa: BLE001 - la comprobación nunca debe romper el cierre
        return cubiertos
    for rango in rangos:
        if any(rango.min_row <= fila <= rango.max_row for fila in filas):
            cubiertos.append(str(rango))
    return cubiertos


def _verify_certificate_rows(
    path: Path | str,
    sheet_name: str,
    esperadas: List[Tuple[int, int, str, str]],
) -> Tuple[int, List[int]]:
    """Relee el archivo recién escrito y confirma que las filas quedaron ahí.

    Es la comprobación que faltaba: antes el cierre informaba "6 filas anexadas"
    solo porque las había enviado al parcheador, sin volver a leer el archivo.

    Args:
        path: Archivo de Certificados ya escrito.
        sheet_name: Hoja donde se anexó.
        esperadas: Lista de ``(fila, N°, documento, código del certificado)``.

    Returns:
        Tupla ``(filas_confirmadas, filas_fallidas)``.
    """
    if not esperadas:
        return 0, []

    todas = [fila for fila, *_ in esperadas]
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except Exception:  # noqa: BLE001 - se informa como verificación fallida
        return 0, todas

    try:
        if sheet_name not in wb.sheetnames:
            return 0, todas
        ws = wb[sheet_name]
        confirmadas = 0
        fallidas: List[int] = []
        for fila, numero, documento, codigo in esperadas:
            val_n = ws.cell(fila, 1).value
            val_doc = ws.cell(fila, 6).value
            val_cod = ws.cell(fila, 8).value

            try:
                n_ok = int(float(str(val_n).strip())) == int(numero)
            except (TypeError, ValueError):
                n_ok = str(val_n).strip() == str(numero).strip()
            doc_ok = normalize_document(val_doc) == normalize_document(documento)
            cod_ok = str(val_cod or "").strip().upper() == str(codigo).strip().upper()

            if n_ok and doc_ok and cod_ok:
                confirmadas += 1
            else:
                fallidas.append(fila)
        return confirmadas, fallidas
    finally:
        wb.close()


def apply_direct_cierre(
    file_paths: Dict[str, Path | str],
    result: AulaCierreResult,
    codigo_curso: str,
    metadatos_control: Optional[Dict[str, AulaMetadata]] = None,
    start_consecutivo: int = 1,
    backup_dir: Optional[Path | str] = None,
) -> DirectUpdateReport:
    """Aplica el cierre directamente sobre los archivos oficiales con respaldo previo y seguridad de celdas.

    Args:
        file_paths: Diccionario con rutas {'control_aulas', 'sistematizacion', 'notas', 'certificados'}.
        result: Resultado consolidado del cruce y motor de reglas.
        codigo_curso: Código oficial del curso (ej. 'ABBUEI205').
        metadatos_control: Diccionario opcional {clase_id: AulaMetadata}.
        start_consecutivo: Número correlativo inicial sugerido para Certificados.
        backup_dir: Directorio para las copias de seguridad (opcional).

    Returns:
        DirectUpdateReport con las estadísticas de actualización.
    """
    report = DirectUpdateReport(timestamp=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

    path_sist = Path(file_paths["sistematizacion"]).resolve()
    path_cert = Path(file_paths["certificados"]).resolve()
    path_ctrl = Path(file_paths["control_aulas"]).resolve()

    # ---------------------------------------------------------
    # PASO 1: GENERACIÓN DE COPIAS DE SEGURIDAD (BACKUP)
    # ---------------------------------------------------------
    # Los backups también sirven como fuente de las extensiones de hoja
    # (`extLst`: listas de validación modernas, formato condicional moderno, etc.)
    # que openpyxl descarta al guardar.
    backup_by_role: Dict[str, Path] = {}
    for role, p in [("sistematizacion", path_sist), ("certificados", path_cert), ("control_aulas", path_ctrl)]:
        try:
            # Respaldo único por archivo: se actualiza con el estado actual y se
            # consolidan las copias antiguas con marca de tiempo.
            legacy = find_legacy_backups(p, backup_dir=backup_dir)
            bk_path = create_timestamped_backup(p, backup_dir=backup_dir)
            report.backups_consolidados += len(legacy)
            report.backups_created.append(bk_path)
            backup_by_role[role] = bk_path
        except Exception as e:
            report.errors.append(f"Error crítico al crear copia de seguridad de {p.name}: {e}")
            return report

    # ---------------------------------------------------------
    # PASO 2: ACTUALIZACIÓN EN SISTEMATIZACIÓN
    # ---------------------------------------------------------
    try:
        # openpyxl se usa SOLO para leer valores y estructura: nunca se guarda con
        # él. La escritura la hace `XlsxPatcher`, que toca únicamente las celdas
        # necesarias y conserva byte a byte el resto del paquete (listas de
        # validación, formato condicional, controles de formulario, tablas
        # dinámicas, gráficos, imágenes, vínculos externos, etc.).
        wb_sist = openpyxl.load_workbook(path_sist, data_only=False)
        ws_sist = wb_sist["Formato"] if "Formato" in wb_sist.sheetnames else wb_sist.active
        sheet_name_sist = ws_sist.title
        patcher_sist = XlsxPatcher(path_sist)
        sheet_sist = patcher_sist.sheet(sheet_name_sist)
        border_thin_sist = patcher_sist.styles.border_id(_BORDER_THIN)
        border_class_sist = patcher_sist.styles.border_id(_BORDER_CLASS_SEPARATOR)

        # Recopilar todos los estudiantes ordenados por fila original
        all_matches: List[Tuple[StudentMatch, ClaseGroup]] = []
        for cg in result.clases:
            for m in cg.matches:
                all_matches.append((m, cg))
        all_matches.sort(key=lambda x: x[0].estudiante_sist.row_number)

        # Mapeo de columnas relevantes desde el primer estudiante
        first_es = all_matches[0][0].estudiante_sist if all_matches else None
        col_map = first_es.column_indices if first_es and first_es.column_indices else {}

        col_nombres = col_map.get("nombres")
        col_apellidos = col_map.get("apellidos")
        col_documento = col_map.get("documento")
        max_col_sist = ws_sist.max_column or 55

        for idx, (m, cg) in enumerate(all_matches):
            es = m.estudiante_sist
            r_idx = es.row_number

            # 0. VERIFICACIÓN DE SEGURIDAD DE IDENTIDAD EN FILA DESTINO
            # Comprobar que la fila r_idx corresponda efectivamente a la persona correcta
            row_doc_raw = str(ws_sist.cell(r_idx, col_documento).value or "") if col_documento else ""
            row_doc_norm = normalize_document(row_doc_raw)
            es_doc_norm = normalize_document(es.documento.value if es.documento else "")

            row_nom = str(ws_sist.cell(r_idx, col_nombres).value or "").strip().upper() if col_nombres else ""
            row_ape = str(ws_sist.cell(r_idx, col_apellidos).value or "").strip().upper() if col_apellidos else ""
            es_nom = str(es.nombres.value or "").strip().upper() if es.nombres else ""
            es_ape = str(es.apellidos.value or "").strip().upper() if es.apellidos else ""

            doc_matches = bool(row_doc_norm and es_doc_norm and row_doc_norm == es_doc_norm)
            name_matches = bool(
                (row_nom and es_nom and (row_nom in es_nom or es_nom in row_nom))
                or (row_ape and es_ape and (row_ape in es_ape or es_ape in row_ape))
            )

            if not (doc_matches or name_matches):
                report.errors.append(
                    f"DISCREPANCIA EN FILA {r_idx} DE SISTEMATIZACIÓN: La celda contiene '{row_nom} {row_ape}' "
                    f"(Doc: '{row_doc_raw}'), pero se esperaba calificar a '{es_nom} {es_ape}' (Doc: '{es_doc_norm}'). "
                    f"Fila protegida contra escritura errónea."
                )
                continue

            # 1. Estandarizar nombres y apellidos a MAYÚSCULAS
            if col_nombres:
                val_nom = ws_sist.cell(r_idx, col_nombres).value
                if isinstance(val_nom, str) and val_nom:
                    val_upper = val_nom.strip().upper()
                    if val_nom != val_upper:
                        sheet_sist.set_cell(r_idx, col_nombres, val_upper, has_value=True)
                        report.sistematizacion_names_uppercased += 1

            if col_apellidos:
                val_ape = ws_sist.cell(r_idx, col_apellidos).value
                if isinstance(val_ape, str) and val_ape:
                    val_upper = val_ape.strip().upper()
                    if val_ape != val_upper:
                        sheet_sist.set_cell(r_idx, col_apellidos, val_upper, has_value=True)
                        report.sistematizacion_names_uppercased += 1

            # 2. Celdas de cierre: solo si están estrictamente VACÍAS
            #    Se aceptan alias de claves porque el parser oficial usa
            #    ('estado', 'elaboro', 'codigo_cert', 'envio') mientras que otros
            #    orígenes/datos de prueba usan nombres más explícitos.
            target_values: Dict[int, Any] = {}
            col_keys = [
                ("modulo_1", "modulo_1"),
                ("modulo_2", "modulo_2"),
                ("modulo_3", "modulo_3"),
                ("modulo_4", "modulo_4"),
                ("modulo_5", "modulo_5"),
                ("promedio", "promedio"),
                ("estado", "estado_certificacion"),
                ("elaboro", "elaboro_certificado"),
                ("codigo_cert", "codigo_certificado"),
                ("envio", "envio_certificado"),
                ("fecha_envio", "fecha_envio"),
            ]
            for primary_key, alias_key in col_keys:
                c_idx = col_map.get(primary_key) or col_map.get(alias_key)
                if c_idx and c_idx in es.raw_row_data:
                    target_values[c_idx] = es.raw_row_data[c_idx]

            # La nota final se copia TAL CUAL de Notas. Si Notas no la trae, la
            # celda queda vacía (nunca se escribe un promedio calculado).
            c_prom = col_map.get("promedio")
            if c_prom and es.raw_row_data.get(c_prom) is None:
                report.sistematizacion_promedios_no_copiados += 1

            for c_idx, val_to_write in target_values.items():
                if val_to_write is None:
                    continue
                cur_val = ws_sist.cell(r_idx, c_idx).value
                if cur_val is None or str(cur_val).strip() == "":
                    sheet_sist.set_cell(r_idx, c_idx, val_to_write, has_value=True)
                    report.sistematizacion_updated_cells += 1
                else:
                    # Celda con datos preexistentes: NO TOCAR
                    report.sistematizacion_untouched_cells += 1

            # 3. Bordes en TODAS las celdas de la fila procesada.
            #    - Fila normal: caja completa con bordes finos.
            #    - Fila de fin de clase: los mismos tres lados finos + borde inferior grueso.
            current_clase = str(es.clase.value if es.clase and es.clase.value else cg.clase_id).strip()
            is_last_of_class = False
            if idx == len(all_matches) - 1:
                is_last_of_class = True
            else:
                next_m, next_cg = all_matches[idx + 1]
                next_es = next_m.estudiante_sist
                next_clase = str(next_es.clase.value if next_es.clase and next_es.clase.value else next_cg.clase_id).strip()
                if current_clase != next_clase:
                    is_last_of_class = True

            row_border = border_class_sist if is_last_of_class else border_thin_sist
            sheet_sist.set_row_border(
                r_idx, 1, max_col_sist, row_border, patcher_sist.styles
            )

        if report.sistematizacion_promedios_no_copiados:
            report.warnings.append(
                f"Sistematización: {report.sistematizacion_promedios_no_copiados} "
                "estudiante(s) no tenían nota final en el archivo de Notas. Esa columna "
                "se dejó vacía: NO se calculó ningún promedio. Revise que Notas traiga "
                "la columna de nota final ('Calculated Final', 'Nota final', "
                "'Promedio curso completo', etc.)."
            )

        wb_sist.close()
        patcher_sist.save()
        _verify_preservation_and_report(
            report,
            "Sistematización",
            path_sist,
            backup_by_role.get("sistematizacion", path_sist),
            {sheet_name_sist},
        )
    except Exception as e:
        report.errors.append(f"Error al actualizar Sistematización: {e}")

    # ---------------------------------------------------------
    # PASO 3: ANEXAR EN CERTIFICADOS
    # ---------------------------------------------------------
    try:
        # Igual que en Sistematización: openpyxl solo LEE. La escritura quirúrgica
        # la hace XlsxPatcher (conserva tablas dinámicas, gráficos, controles…).
        wb_cert = openpyxl.load_workbook(path_cert, data_only=False)
        if "Códigos" in wb_cert.sheetnames:
            ws_cert = wb_cert["Códigos"]
        elif "Certificados" in wb_cert.sheetnames:
            ws_cert = wb_cert["Certificados"]
        elif "CERTIFICADOS" in wb_cert.sheetnames:
            ws_cert = wb_cert["CERTIFICADOS"]
        else:
            ws_cert = wb_cert.active
        sheet_name_cert = ws_cert.title
        patcher_cert = XlsxPatcher(path_cert)
        sheet_cert = patcher_cert.sheet(sheet_name_cert)
        font_cert_id = patcher_cert.styles.font_id(
            _FONT_CERTIFICADO_NAME, _FONT_CERTIFICADO_SIZE
        )
        border_thin_cert = patcher_cert.styles.border_id(_BORDER_THIN)
        border_class_cert = patcher_cert.styles.border_id(_BORDER_CLASS_SEPARATOR)

        # --- Localizar el FINAL REAL de la tabla de Certificados ----------
        # No se usa "última fila con algo en A o B": una nota, un total o un
        # título suelto debajo de la tabla movía el punto de anexado cientos de
        # filas hacia abajo y los certificados quedaban fuera de la vista del
        # usuario (parecía que no se habían escrito).
        last_row, max_n, uso_respaldo = _inspect_certificate_table(ws_cert)
        if uso_respaldo:
            report.warnings.append(
                "Certificados: la hoja no presenta el formato esperado de la tabla; "
                "se anexará después de la última fila con contenido en las columnas A/B. "
                "Verifique que las filas nuevas hayan quedado dentro de la tabla."
            )

        previos_por_doc = _certificate_rows_by_document(ws_cert, codigo_curso)
        report.certificados_previos = len(previos_por_doc)
        report.certificados_hoja = sheet_name_cert

        # Determinar consecutivo inicial
        current_n = max(max_n + 1, start_consecutivo)

        # Filtrar solo estudiantes aprobados
        all_approved: List[Tuple[StudentMatch, ClaseGroup]] = []
        for cg in result.clases:
            for m in cg.matches:
                st = str(m.estudiante_sist.estado_calculado or "").strip().lower()
                if "aprob" in st and "no" not in st:
                    all_approved.append((m, cg))
        all_approved.sort(key=lambda item: item[0].estudiante_sist.row_number)

        # No duplicar certificados ya registrados del mismo curso (mismo documento).
        pendientes: List[Tuple[StudentMatch, ClaseGroup]] = []
        for m, cg in all_approved:
            es = m.estudiante_sist
            doc_norm = normalize_document(es.documento.value if es.documento else "")
            if doc_norm and doc_norm in previos_por_doc:
                report.certificados_omitidos_duplicados += 1
                continue
            pendientes.append((m, cg))

        # Primera fila libre a partir del final de la tabla (nunca sobre datos).
        curr_cert_r = (last_row + 1) if last_row else 2
        max_cert_r = min(ws_cert.max_row or 1, _MAX_FILAS_CERTIFICADOS)
        fila_ajustada = False
        while curr_cert_r <= max_cert_r and not _fila_certificados_vacia(ws_cert, curr_cert_r):
            curr_cert_r += 1
            fila_ajustada = True
        if fila_ajustada:
            report.warnings.append(
                "Certificados: la fila inmediatamente posterior a la tabla ya tenía "
                f"contenido ajeno; los certificados se anexaron desde la fila {curr_cert_r}."
            )

        fila_inicial = curr_cert_r
        n_inicial = current_n
        filas_esperadas: List[Tuple[int, int, str, str]] = []
        filas_ocultas_restauradas = 0

        for idx, (m, cg) in enumerate(pendientes):
            es = m.estudiante_sist
            meta = (metadatos_control or {}).get(cg.clase_id) or AulaMetadata()
            programa_str = meta.programa_academico.value if meta.programa_academico else ""
            catalogo_str = meta.catalogo.value if meta.catalogo else ""
            f_ini_dt = parse_date(meta.fecha_inicio.value) if meta.fecha_inicio else None
            f_fin_dt = parse_date(meta.fecha_fin.value) if meta.fecha_fin else None

            f_ini_str = format_ciclo_date(f_ini_dt) if f_ini_dt else ""
            f_fin_str = format_ciclo_date(f_fin_dt) if f_fin_dt else ""
            f_envio_cert_str = format_certificate_date(f_fin_dt) if f_fin_dt else ""

            ciclo_val = format_ciclo(cg.clase_id, catalogo_str, programa_str, f_ini_str, f_fin_str).upper()

            formador_lider = ""
            if es.formador_lider and es.formador_lider.value:
                formador_lider = str(es.formador_lider.value).strip()
            elif meta.docente_principal and meta.docente_principal.value:
                formador_lider = str(meta.docente_principal.value).strip()

            docente_coin_val = format_docente_coin(codigo_curso, formador_lider).strip().upper()
            # Columna 'Semestre': periodo académico con formato oficial 'AAAA-S' (ej. '2026-1').
            semestre_val = format_semester_label(f_fin_dt) if f_fin_dt else ""
            ano_val = f_fin_dt.year if isinstance(f_fin_dt, (date, datetime)) else 2026

            nom = str(es.nombres.value if es.nombres else "").strip().upper()
            ape = str(es.apellidos.value if es.apellidos else "").strip().upper()
            doc = es.documento.value if es.documento else ""
            cor = str(es.correo.value if es.correo else "").strip()

            row_cert = [
                current_n,
                ciclo_val,
                docente_coin_val,
                nom,
                ape,
                doc,
                cor,
                str(codigo_curso).strip().upper(),
                f_envio_cert_str,
                ano_val,
                semestre_val,
                1,  # Total certificados = 1
                None,
            ]

            # Separador con borde inferior grueso al terminar cada clase distinta
            current_clase = str(es.clase.value if es.clase and es.clase.value else cg.clase_id).strip()
            is_last_of_class = False
            if idx == len(pendientes) - 1:
                is_last_of_class = True
            else:
                next_m, next_cg = pendientes[idx + 1]
                next_es = next_m.estudiante_sist
                next_clase = str(next_es.clase.value if next_es.clase and next_es.clase.value else next_cg.clase_id).strip()
                if current_clase != next_clase:
                    is_last_of_class = True

            border_cell = border_class_cert if is_last_of_class else border_thin_cert

            # Si la fila destino existía como plantilla oculta, se hace visible:
            # de lo contrario el certificado se escribe pero no se ve en Excel.
            if sheet_cert.show_row(curr_cert_r):
                filas_ocultas_restauradas += 1

            for c_idx, val in enumerate(row_cert, start=1):
                # La columna N° (1) es la única que NO lleva borde inferior grueso
                # en la fila de fin de clase; el resto sí.
                cell_border = border_thin_cert if (is_last_of_class and c_idx == 1) else border_cell
                style_id = patcher_cert.styles.xf_id(
                    base=0,
                    border=cell_border,
                    font=font_cert_id,
                    alignment=_ALIGN_CENTER_CERT,
                )
                sheet_cert.set_cell(
                    curr_cert_r, c_idx, val, has_value=True, style=style_id
                )

            filas_esperadas.append(
                (curr_cert_r, current_n, str(doc), str(codigo_curso).strip().upper())
            )
            current_n += 1
            curr_cert_r += 1
            report.certificados_rows_appended += 1

        if filas_esperadas:
            report.certificados_fila_inicial = fila_inicial
            report.certificados_fila_final = curr_cert_r - 1
            report.certificados_n_inicial = n_inicial
            report.certificados_n_final = current_n - 1

        if filas_ocultas_restauradas:
            report.warnings.append(
                f"Certificados: {filas_ocultas_restauradas} fila(s) destino estaban "
                "ocultas o agrupadas y se hicieron visibles para que los certificados "
                "se puedan ver en Excel."
            )

        if report.certificados_omitidos_duplicados:
            report.warnings.append(
                f"Certificados: se omitieron {report.certificados_omitidos_duplicados} "
                f"estudiante(s) que ya tenían certificado registrado del curso "
                f"'{str(codigo_curso).strip().upper()}' en la hoja '{sheet_name_cert}' "
                "(no se duplicaron). Si el cierre anterior no se vio reflejado, revise "
                "esa hoja antes de repetir el proceso."
            )

        merges_nombres = _merges_covering_rows(
            ws_cert, [fila for fila, *_ in filas_esperadas]
        )
        if merges_nombres:
            report.warnings.append(
                "Certificados: las filas anexadas caen dentro de celdas combinadas "
                f"preexistentes ({', '.join(merges_nombres)}). Excel solo muestra el "
                "valor de la celda superior izquierda, así que parte de los datos "
                "podría no verse; revise el formato de esas filas."
            )

        wb_cert.close()
        patcher_cert.save()

        # Relectura del archivo: confirma que lo anexado quedó realmente escrito.
        verificadas, fallidas = _verify_certificate_rows(
            path_cert, sheet_name_cert, filas_esperadas
        )
        report.certificados_rows_verificados = verificadas
        if fallidas:
            report.errors.append(
                "Certificados: la relectura del archivo no encontró las filas recién "
                f"anexadas {fallidas}. Restaure el último respaldo (opción 6) y revise "
                "el archivo antes de continuar."
            )

        _verify_preservation_and_report(
            report,
            "Certificados",
            path_cert,
            backup_by_role.get("certificados", path_cert),
            {sheet_name_cert},
        )
    except Exception as e:
        report.errors.append(f"Error al anexar en Certificados: {e}")

    # ---------------------------------------------------------
    # PASO 4: ACTUALIZAR EN CONTROL DE AULAS
    # ---------------------------------------------------------
    try:
        # openpyxl solo LEE; la escritura quirúrgica la hace XlsxPatcher.
        wb_ctrl = openpyxl.load_workbook(path_ctrl, data_only=False)
        patcher_ctrl = XlsxPatcher(path_ctrl)
        fecha_cierre_hoy = datetime.now().strftime("%d/%m/%Y")

        sheets_to_check = [
            s for s in wb_ctrl.sheetnames
            if s != CONTROL_AULAS_IGNORED_SHEET and (s in CONTROL_AULAS_ALL_SHEETS or s.startswith("Tabla_"))
        ]

        clases_pendientes = {str(cg.clase_id).strip() for cg in result.clases}
        modified_control_sheets: set = set()

        for sheet_name in sheets_to_check:
            ws = wb_ctrl[sheet_name]
            max_col = ws.max_column or 1
            max_row = ws.max_row or 1

            col_clase = None
            col_certificados = None

            for c in range(1, max_col + 1):
                h_val = str(ws.cell(1, c).value or "").strip().upper()
                if h_val == "CLASE":
                    col_clase = c
                elif "CERTIFICADOS" in h_val:
                    col_certificados = c

            if not col_clase or not col_certificados:
                continue

            sheet_ctrl = patcher_ctrl.sheet(sheet_name)

            for r in range(2, max_row + 1):
                c_val = ws.cell(r, col_clase).value
                if c_val is None:
                    continue
                c_str = str(int(c_val)) if isinstance(c_val, float) and c_val.is_integer() else str(c_val).strip()

                if c_str in clases_pendientes:
                    cur_val = ws.cell(r, col_certificados).value
                    if cur_val is None or str(cur_val).strip() == "":
                        sheet_ctrl.set_cell(
                            r, col_certificados, fecha_cierre_hoy, has_value=True
                        )
                        report.control_aulas_classes_marked += 1
                        modified_control_sheets.add(sheet_name)
                    else:
                        # La clase ya tenía fecha de cierre: no se toca ni se cuenta
                        # como error, pero se informa para que el usuario lo sepa.
                        report.control_aulas_classes_already_marked += 1
                    clases_pendientes.discard(c_str)

        report.control_aulas_classes_not_found = sorted(clases_pendientes)

        wb_ctrl.close()
        patcher_ctrl.save()
        _verify_preservation_and_report(
            report,
            "Control de Aulas",
            path_ctrl,
            backup_by_role.get("control_aulas", path_ctrl),
            modified_control_sheets,
        )
    except Exception as e:
        report.errors.append(f"Error al actualizar Control de Aulas: {e}")

    return report
