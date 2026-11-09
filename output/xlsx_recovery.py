"""
Diagnóstico y reparación LOCAL de libros `.xlsx` dañados por una edición de celdas.

Motivación (reporte real)
-------------------------
Si se reemplaza el valor de la **celda maestra de una fórmula compartida** de
Excel, las demás celdas del grupo quedan con `si="N"` apuntando a un maestro que
ya no existe. Excel detecta el libro como dañado, lo "repara" (eliminando
fórmulas) y después puede mostrar el aviso **«CARGA BLOQUEADA — Se ha producido
un problema con este archivo…»**, incluso al guardar una copia.

Este módulo:

1. **Diagnostica** el paquete `.xlsx` parte por parte (ZIP, XML bien formado,
   fórmulas compartidas huérfanas, celdas duplicadas, estilos fuera de rango,
   `calcPr` duplicado, índices inválidos de `sharedStrings`, caracteres de
   control prohibidos).
2. **Repara** esos daños en una **copia nueva** (`<nombre>_reparado.xlsx`); el
   archivo original nunca se modifica.
3. Nunca requiere Internet ni envía datos: todo es local.

Uso:
    python -m output.xlsx_recovery "C:\\ruta\\Sistematizacion.xlsx"            # solo diagnóstico
    python -m output.xlsx_recovery "C:\\ruta\\Sistematizacion.xlsx" --reparar  # genera copia reparada

También está disponible en el menú de `main.py` (opción 9).
"""

from __future__ import annotations

import argparse
import re
import shutil
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from xml.parsers import expat

from output.xlsx_patcher import sheet_paths

__all__ = [
    "RecoveryReport",
    "diagnose_xlsx",
    "repair_xlsx",
]


# Caracteres que XML 1.0 prohíbe (misma regla que usa el patcher al escribir).
_INVALID_XML_CHARS_RE = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]")

# Ampersand sin escapar (causa habitual de XML mal formado tras ediciones manuales).
_UNESCAPED_AMP_RE = re.compile(
    r"&(?!(?:[A-Za-z_][A-Za-z0-9_.\-]*|#[0-9]+|#x[0-9A-Fa-f]+);)"
)

# Etiquetas con atributos que nos interesan.
_ATTR_RE = r"""([\w:.\-]+)\s*=\s*("([^"]*)"|'([^']*)')"""
_CELL_RE = re.compile(r"<c\b([^>]*?)(/>|>(.*?)</c>)", re.DOTALL)
_ROW_RE = re.compile(r"(<row\b[^>]*?)(/>|>(.*?)</row>)", re.DOTALL)
_FORMULA_RE = re.compile(r"<f\b([^>]*?)(/>|>(.*?)</f>)", re.DOTALL)
_CALCPR_RE = re.compile(r"<calcPr\b[^>]*/>|<calcPr\b[^>]*>.*?</calcPr>", re.DOTALL)
_SI_RE = re.compile(r"<si\b[^>]*?>.*?</si>|<si\b[^>]*?/>", re.DOTALL)
_SHARED_STR_CELL_RE = re.compile(
    r'<c\b[^>]*\bt="s"[^>]*>\s*<v>(\d+)</v>\s*</c>', re.DOTALL
)


def _attr(attrs: str, name: str) -> Optional[str]:
    """Lee un atributo XML dentro del texto de atributos de una etiqueta."""
    m = re.search(
        r"\b%s\s*=\s*(\"([^\"]*)\"|'([^']*)')" % re.escape(name), attrs or ""
    )
    if not m:
        return None
    return m.group(2) if m.group(2) is not None else m.group(3)


def _del_attr(attrs: str, name: str) -> str:
    """Elimina un atributo (si existe) conservando el resto del texto."""
    return re.sub(r"\s%s\s*=\s*(\"[^\"]*\"|'[^']*')" % re.escape(name), "", attrs or "")


def _xml_error(data: bytes) -> Optional[str]:
    """Devuelve el error de parseo XML, o None si la parte es válida."""
    parser = expat.ParserCreate(namespace_separator="|")
    try:
        parser.Parse(data, True)
    except expat.ExpatError as exc:
        return str(exc)
    return None


def _is_xml_part(name: str) -> bool:
    return name.endswith(".xml") or name.endswith(".rels") or name.endswith(".vml")


def _read_parts(path: Path) -> Tuple[Dict[str, bytes], List[zipfile.ZipInfo], Optional[str]]:
    """Lee todas las partes de un `.xlsx` conservando el orden y los metadatos."""
    with zipfile.ZipFile(path) as zf:
        bad = zf.testzip()
        infos = zf.infolist()
        parts = {info.filename: zf.read(info.filename) for info in infos}
    return parts, infos, bad


@dataclass
class RecoveryReport:
    """Resultado del diagnóstico o de la reparación de un libro."""

    archivo: str = ""
    partes: int = 0
    zip_danado: Optional[str] = None
    xml_invalidos: List[str] = field(default_factory=list)
    problemas: List[str] = field(default_factory=list)
    avisos: List[str] = field(default_factory=list)

    # Contadores de daños detectados (y reparados, si aplica).
    caracteres_invalidos: int = 0
    formulas_compartidas: int = 0
    esclavas_huerfanas: int = 0
    celdas_duplicadas: int = 0
    estilos_fuera_de_rango: int = 0
    calcpr_duplicados: int = 0
    strings_fuera_de_rango: int = 0
    referencias_incoherentes: int = 0

    # Contadores de lo que efectivamente se corrigió en la copia reparada.
    formulas_compartidas_deshechas: int = 0
    esclavas_convertidas: int = 0
    celdas_duplicadas_eliminadas: int = 0
    estilos_corregidos: int = 0
    strings_agregados: int = 0

    reparado: bool = False
    salida: Optional[Path] = None
    problemas_restantes: List[str] = field(default_factory=list)

    @property
    def hay_danos(self) -> bool:
        return bool(
            self.zip_danado
            or self.xml_invalidos
            or self.problemas
            or self.caracteres_invalidos
            or self.esclavas_huerfanas
            or self.celdas_duplicadas
            or self.estilos_fuera_de_rango
            or self.calcpr_duplicados
            or self.strings_fuera_de_rango
            or self.referencias_incoherentes
        )

    @property
    def tiene_danos_graves(self) -> bool:
        """Daños que obligan a Excel a "reparar" el libro (o bloquear el guardado).

        Se excluyen los avisos menores (referencias incoherentes) para no bloquear
        un cierre por algo que Excel tolera.
        """
        return bool(
            self.zip_danado
            or self.xml_invalidos
            or self.esclavas_huerfanas
            or self.celdas_duplicadas
            or self.estilos_fuera_de_rango
            or self.strings_fuera_de_rango
            or self.calcpr_duplicados
        )

    @property
    def ok(self) -> bool:
        """True si el libro no presenta daños estructurales detectables."""
        return not self.hay_danos and not self.problemas_restantes

    def resumen(self) -> str:
        icono = "✅" if self.ok else "⚠️"
        return (
            f"{icono} {self.archivo}: {self.partes} parte(s), "
            f"XML inválidos {len(self.xml_invalidos)}, "
            f"fórmulas compartidas huérfanas {self.esclavas_huerfanas}, "
            f"celdas duplicadas {self.celdas_duplicadas}, "
            f"estilos fuera de rango {self.estilos_fuera_de_rango}, "
            f"calcPr duplicados {self.calcpr_duplicados}"
        )

    def resumen_reparacion(self) -> str:
        """Resumen de lo que se corrigió (o del diagnóstico si no se reparó)."""
        if not self.reparado:
            return self.resumen()
        return (
            f"✅ {self.archivo}: copia reparada — "
            f"{self.formulas_compartidas_deshechas} grupo(s) de fórmulas compartidas "
            f"huérfanas reparados, {self.esclavas_convertidas} esclava(s) convertidas a su "
            f"valor, {self.celdas_duplicadas_eliminadas} celda(s) duplicada(s) "
            f"eliminada(s), {self.estilos_corregidos} estilo(s) corregido(s), "
            f"{self.strings_agregados} texto(s) compartido(s) añadido(s)"
        )

    def detalle(self) -> str:
        """Informe en texto (sin datos personales) para guardar junto al archivo."""
        lineas = [
            "DIAGNÓSTICO DE INTEGRIDAD DE LIBRO .XLSX",
            "=" * 60,
            f"Archivo: {self.archivo}",
            f"Partes del paquete: {self.partes}",
            f"Resultado: {'SIN DAÑOS DETECTADOS' if self.ok else 'DAÑOS DETECTADOS'}",
            "",
        ]
        if self.zip_danado:
            lineas.append(f"• ZIP dañado en: {self.zip_danado}")
        if self.xml_invalidos:
            lineas.append("• Partes XML no válidas:")
            lineas.extend(f"    - {item}" for item in self.xml_invalidos)
        if self.problemas:
            lineas.append("• Problemas:")
            lineas.extend(f"    - {item}" for item in self.problemas)
        contadores = [
            ("Caracteres de control no válidos", self.caracteres_invalidos),
            ("Fórmulas compartidas (maestras)", self.formulas_compartidas),
            ("Fórmulas compartidas huérfanas", self.esclavas_huerfanas),
            ("Celdas duplicadas en una fila", self.celdas_duplicadas),
            ("Estilos (s) fuera de rango", self.estilos_fuera_de_rango),
            ("Elementos <calcPr> duplicados", self.calcpr_duplicados),
            ("Índices de texto compartido fuera de rango", self.strings_fuera_de_rango),
            ("Referencias de celda incoherentes", self.referencias_incoherentes),
        ]
        for etiqueta, valor in contadores:
            if valor:
                lineas.append(f"• {etiqueta}: {valor}")
        if self.avisos:
            lineas.append("• Avisos:")
            lineas.extend(f"    - {item}" for item in self.avisos)
        if self.reparado:
            lineas.extend(
                [
                    "",
                    "REPARACIÓN",
                    "-" * 60,
                    f"Copia reparada: {self.salida}",
                    "El archivo original NO se modificó.",
                    f"• Grupos de fórmulas compartidas huérfanas reparados: {self.formulas_compartidas_deshechas}",
                    f"• Fórmulas esclavas convertidas a su valor: {self.esclavas_convertidas}",
                    f"• Celdas duplicadas eliminadas: {self.celdas_duplicadas_eliminadas}",
                    f"• Estilos inexistentes corregidos: {self.estilos_corregidos}",
                    f"• Textos compartidos añadidos: {self.strings_agregados}",
                ]
            )
            if self.problemas_restantes:
                lineas.append("• No se pudieron corregir:")
                lineas.extend(f"    - {item}" for item in self.problemas_restantes)
        lineas.append("=" * 60)
        return "\n".join(lineas)


# ---------------------------------------------------------------------------
# Diagnóstico
# ---------------------------------------------------------------------------

def _cellxfs_count(styles_xml: Optional[bytes]) -> Optional[int]:
    """Número de estilos de celda (`<xf>`) declarados en `xl/styles.xml`."""
    if not styles_xml:
        return None
    xml = styles_xml.decode("utf-8", "ignore")
    m = re.search(r"<cellXfs\b[^>]*>(.*?)</cellXfs>|<cellXfs\b[^>]*/>", xml, re.DOTALL)
    if not m:
        return None
    if m.group(0).endswith("/>"):
        return 0
    return len(re.findall(r"<xf\b", m.group(1) or ""))


def _analyze_sheet_xml(
    xml: str, cellxfs: Optional[int], strings_count: Optional[int]
) -> Dict[str, int]:
    """Cuenta los daños estructurales de una hoja (sin modificar nada)."""
    datos = {
        "formulas_compartidas": 0,
        "esclavas_huerfanas": 0,
        "celdas_duplicadas": 0,
        "estilos_fuera_de_rango": 0,
        "strings_fuera_de_rango": 0,
        "referencias_incoherentes": 0,
        "caracteres_invalidos": len(_INVALID_XML_CHARS_RE.findall(xml)),
    }

    # Fórmulas compartidas: maestro = <f t="shared" ... ref="...">texto</f>
    maestros = set()
    esclavas: List[str] = []
    for m in _FORMULA_RE.finditer(xml):
        attrs = m.group(1) or ""
        if _attr(attrs, "t") != "shared":
            continue
        si = _attr(attrs, "si")
        if si is None:
            continue
        if _attr(attrs, "ref"):
            maestros.add(si)
        elif m.group(2) == "/>" or not (m.group(3) or "").strip():
            esclavas.append(si)
        else:
            # shared con texto pero sin ref: no es un maestro válido.
            esclavas.append(si)
    datos["formulas_compartidas"] = len(maestros)
    datos["esclavas_huerfanas"] = sum(1 for si in esclavas if si not in maestros)

    # Celdas duplicadas y coherencia de referencias.
    for row_m in _ROW_RE.finditer(xml):
        if row_m.group(2) == "/>":
            continue
        numero = _attr(row_m.group(1) or "", "r")
        vistos = set()
        for cm in _CELL_RE.finditer(row_m.group(3) or ""):
            attrs = cm.group(1) or ""
            ref = _attr(attrs, "r")
            if not ref:
                continue
            if ref in vistos:
                datos["celdas_duplicadas"] += 1
            vistos.add(ref)
            rm = re.match(r"^([A-Z]+)(\d+)$", ref)
            if rm and numero and rm.group(2) != str(numero):
                datos["referencias_incoherentes"] += 1
            s = _attr(attrs, "s")
            if s is not None and cellxfs is not None:
                try:
                    if int(s) >= cellxfs or int(s) < 0:
                        datos["estilos_fuera_de_rango"] += 1
                except ValueError:
                    datos["estilos_fuera_de_rango"] += 1
            if _attr(attrs, "t") == "s" and strings_count is not None and cm.group(2) != "/>":
                vm = re.search(r"<v>\s*(\d+)\s*</v>", cm.group(3) or "")
                if vm and int(vm.group(1)) >= strings_count:
                    datos["strings_fuera_de_rango"] += 1
    return datos


def diagnose_xlsx(path: Path | str) -> RecoveryReport:
    """Analiza un `.xlsx` y devuelve un informe de daños (no escribe nada)."""
    ruta = Path(path).resolve()
    report = RecoveryReport(archivo=ruta.name)
    if not ruta.exists():
        report.problemas.append("El archivo no existe.")
        return report

    try:
        parts, _infos, bad = _read_parts(ruta)
    except (zipfile.BadZipFile, OSError) as exc:
        report.problemas.append(f"No se pudo abrir el archivo como .xlsx: {exc}")
        return report

    report.partes = len(parts)
    report.zip_danado = bad
    if bad:
        report.problemas.append(f"El contenedor ZIP está dañado en la parte '{bad}'.")

    for obligatoria in ("[Content_Types].xml", "_rels/.rels", "xl/workbook.xml"):
        if obligatoria not in parts:
            report.problemas.append(f"Falta la parte obligatoria '{obligatoria}'.")

    for name, data in parts.items():
        if not _is_xml_part(name):
            continue
        error = _xml_error(data)
        if error:
            report.xml_invalidos.append(f"{name}: {error}")

    # Partes de hoja: daños que Excel intenta "reparar".
    hojas = sheet_paths(parts)
    cellxfs = _cellxfs_count(parts.get("xl/styles.xml"))
    strings_count = None
    if "xl/sharedStrings.xml" in parts:
        strings_count = len(
            _SI_RE.findall(parts["xl/sharedStrings.xml"].decode("utf-8", "ignore"))
        )

    totales: Dict[str, int] = {}
    for _hoja, parte in hojas.items():
        data = parts.get(parte)
        if data is None:
            report.problemas.append(f"La hoja apunta a una parte inexistente: {parte}")
            continue
        datos = _analyze_sheet_xml(
            data.decode("utf-8", "ignore"), cellxfs, strings_count
        )
        for clave, valor in datos.items():
            totales[clave] = totales.get(clave, 0) + valor
        if datos["esclavas_huerfanas"]:
            report.problemas.append(
                f"'{_hoja}' ({parte}): {datos['esclavas_huerfanas']} fórmula(s) "
                "compartida(s) huérfana(s). Es la causa habitual del aviso de Excel."
            )
        if datos["celdas_duplicadas"]:
            report.problemas.append(
                f"'{_hoja}' ({parte}): {datos['celdas_duplicadas']} celda(s) duplicada(s)."
            )
        if datos["estilos_fuera_de_rango"]:
            report.problemas.append(
                f"'{_hoja}' ({parte}): {datos['estilos_fuera_de_rango']} celda(s) con "
                "estilo inexistente."
            )
        if datos["strings_fuera_de_rango"]:
            report.problemas.append(
                f"'{_hoja}' ({parte}): {datos['strings_fuera_de_rango']} celda(s) "
                "apuntan a un texto compartido inexistente."
            )
        if datos["referencias_incoherentes"]:
            report.avisos.append(
                f"'{_hoja}' ({parte}): {datos['referencias_incoherentes']} celda(s) "
                "con referencia que no coincide con su fila."
            )

    for clave, valor in totales.items():
        setattr(report, clave, valor)

    wb = parts.get("xl/workbook.xml", b"").decode("utf-8", "ignore")
    report.calcpr_duplicados = max(0, len(_CALCPR_RE.findall(wb)) - 1)
    if report.calcpr_duplicados:
        report.problemas.append(
            f"'xl/workbook.xml': {report.calcpr_duplicados} elemento(s) <calcPr> "
            "duplicado(s) (el esquema admite uno solo)."
        )

    return report


# ---------------------------------------------------------------------------
# Reparación
# ---------------------------------------------------------------------------

def _fix_calcpr(xml: str) -> Tuple[str, int]:
    """Deja un único `<calcPr>` en `xl/workbook.xml`."""
    encontrados = list(_CALCPR_RE.finditer(xml))
    if len(encontrados) <= 1:
        return xml, 0
    for m in reversed(encontrados[1:]):
        xml = xml[: m.start()] + xml[m.end():]
    return xml, len(encontrados) - 1


def _fix_shared_formulas(xml: str) -> Tuple[str, int, int]:
    """Repara SOLO los grupos de fórmulas compartidas que quedaron huérfanos.

    Un grupo es huérfano cuando tiene celdas esclavas (`<f t="shared" si="N"/>`)
    pero ya no tiene maestro con fórmula (`<f t="shared" ref="..." si="N">…</f>`):
    es exactamente el daño que deja Excel sin poder abrir el libro. Las esclavas
    se convierten en su valor constante (se elimina el `<f>` y se conserva el
    `<v>` con el último valor calculado). Los grupos **sanos** no se tocan.

    Returns:
        (XML resultante, grupos huérfanos reparados, esclavas convertidas).
    """
    grupos: Dict[str, Dict[str, object]] = {}
    for m in _FORMULA_RE.finditer(xml):
        attrs = m.group(1) or ""
        if _attr(attrs, "t") != "shared":
            continue
        si = _attr(attrs, "si")
        if si is None:
            continue
        grupo = grupos.setdefault(si, {"maestro": False, "esclavas": 0})
        if _attr(attrs, "ref") and (m.group(3) or "").strip():
            grupo["maestro"] = True
        else:
            grupo["esclavas"] = int(grupo["esclavas"]) + 1

    rotos = {
        si for si, grupo in grupos.items()
        if grupo["esclavas"] and not grupo["maestro"]
    }
    if not rotos:
        return xml, 0, 0

    esclavas = 0

    def _reemplazar(m: "re.Match") -> str:
        nonlocal esclavas
        attrs = m.group(1) or ""
        if _attr(attrs, "t") != "shared" or _attr(attrs, "si") not in rotos:
            return m.group(0)
        esclavas += 1
        return ""

    nuevo = _FORMULA_RE.sub(_reemplazar, xml)
    # Celdas que quedaron vacías tras quitar la fórmula: se autocierran.
    nuevo = re.sub(r"<c\b([^>]*?)></c>", r"<c\1/>", nuevo)
    return nuevo, len(rotos), esclavas


def _dedupe_cells(xml: str) -> Tuple[str, int]:
    """Elimina celdas duplicadas dentro de una misma fila (conserva la primera)."""
    eliminadas = 0

    def _row(m: "re.Match") -> str:
        nonlocal eliminadas
        if m.group(2) == "/>":
            return m.group(0)
        contenido = m.group(3) or ""
        vistos = set()
        piezas: List[str] = []
        ultimo = 0
        for cm in _CELL_RE.finditer(contenido):
            ref = _attr(cm.group(1) or "", "r")
            if ref and ref in vistos:
                piezas.append(contenido[ultimo: cm.start()])
                ultimo = cm.end()
                eliminadas += 1
            vistos.add(ref)
        piezas.append(contenido[ultimo:])
        return "%s>%s</row>" % (m.group(1), "".join(piezas))

    return _ROW_RE.sub(_row, xml), eliminadas


def _fix_style_refs(xml: str, cellxfs: Optional[int]) -> Tuple[str, int]:
    """Quita el atributo `s` cuando apunta a un estilo que no existe."""
    if cellxfs is None:
        return xml, 0
    corregidas = 0

    def _cell(m: "re.Match") -> str:
        nonlocal corregidas
        attrs = m.group(1) or ""
        s = _attr(attrs, "s")
        if s is None:
            return m.group(0)
        try:
            valido = 0 <= int(s) < cellxfs
        except ValueError:
            valido = False
        if valido:
            return m.group(0)
        nuevos = _del_attr(attrs, "s")
        corregidas += 1
        cierre = "/>" if m.group(2) == "/>" else ">%s</c>" % (m.group(3) or "")
        return "<c%s%s" % (nuevos, cierre)

    return _CELL_RE.sub(_cell, xml), corregidas


def _fix_shared_strings_index(
    sheet_xml: str, strings_xml: Optional[bytes]
) -> Tuple[str, Optional[bytes], int]:
    """Amplía `sharedStrings.xml` si alguna celda apunta más allá del final."""
    if strings_xml is None:
        return sheet_xml, strings_xml, 0
    texto = strings_xml.decode("utf-8", "ignore")
    existentes = len(_SI_RE.findall(texto))
    indices = [int(n) for n in _SHARED_STR_CELL_RE.findall(sheet_xml)]
    maximo = max(indices, default=-1)
    faltan = max(0, maximo + 1 - existentes)
    if not faltan:
        return sheet_xml, strings_xml, 0

    relleno = "<si><t xml:space=\"preserve\"></t></si>" * faltan
    if "</sst>" in texto:
        texto = texto.replace("</sst>", relleno + "</sst>", 1)
    else:
        return sheet_xml, strings_xml, 0

    # Actualizar count y uniqueCount si están presentes.
    def _actualizar(attrs: str) -> str:
        nuevos = attrs
        for nombre in ("count", "uniqueCount"):
            if _attr(nuevos, nombre) is not None:
                nuevos = re.sub(
                    r"(\s%s\s*=\s*)(\"[^\"]*\"|'[^']*')" % nombre,
                    lambda mm: f'{mm.group(1)}"{existentes + faltan}"',
                    nuevos,
                    count=1,
                )
        return nuevos

    raiz = re.search(r"<sst\b([^>]*)>", texto)
    if raiz:
        texto = texto[: raiz.start()] + "<sst%s>" % _actualizar(raiz.group(1)) + texto[raiz.end():]
    return sheet_xml, texto.encode("utf-8"), faltan


def repair_xlsx(
    path: Path | str, output_path: Path | str | None = None
) -> RecoveryReport:
    """Genera una copia reparada del libro. El original NUNCA se modifica.

    Args:
        path: Ruta del `.xlsx` dañado.
        output_path: Ruta de salida. Por defecto `<nombre>_reparado.xlsx` junto
            al original (se añade un sufijo si ya existe).

    Returns:
        `RecoveryReport` con el diagnóstico previo y el resultado de la reparación.
    """
    ruta = Path(path).resolve()
    report = diagnose_xlsx(ruta)
    if not ruta.exists():
        return report

    if output_path is None:
        destino = ruta.with_name(f"{ruta.stem}_reparado{ruta.suffix}")
    else:
        destino = Path(output_path).resolve()

    if destino == ruta:
        raise ValueError("La reparación nunca escribe sobre el archivo original.")
    if destino.exists():
        n = 2
        while True:
            candidato = destino.with_name(f"{destino.stem}_{n}{destino.suffix}")
            if not candidato.exists():
                destino = candidato
                break
            n += 1

    try:
        parts, infos, _bad = _read_parts(ruta)
    except (zipfile.BadZipFile, OSError) as exc:
        report.problemas_restantes.append(f"No se pudo leer el paquete: {exc}")
        return report

    nuevos = dict(parts)

    # 1. Caracteres de control prohibidos en cualquier parte XML.
    for name, data in list(nuevos.items()):
        if not _is_xml_part(name):
            continue
        texto = data.decode("utf-8", "ignore")
        limpio = _INVALID_XML_CHARS_RE.sub("", texto)
        if limpio != texto:
            nuevos[name] = limpio.encode("utf-8")

    # 2. XML mal formado: intentar escapar ampersands sueltos.
    for name in list(report.xml_invalidos):
        parte = name.split(":", 1)[0]
        data = nuevos.get(parte)
        if data is None:
            continue
        texto = data.decode("utf-8", "ignore")
        arreglado = _UNESCAPED_AMP_RE.sub("&amp;", texto)
        if arreglado != texto and _xml_error(arreglado.encode("utf-8")) is None:
            nuevos[parte] = arreglado.encode("utf-8")
            report.avisos.append(
                f"'{parte}': se escaparon ampersands sueltos para que el XML fuera válido."
            )

    # 3. <calcPr> duplicado en el libro.
    if "xl/workbook.xml" in nuevos:
        texto = nuevos["xl/workbook.xml"].decode("utf-8", "ignore")
        texto, eliminados = _fix_calcpr(texto)
        if eliminados:
            nuevos["xl/workbook.xml"] = texto.encode("utf-8")

    # 4. Estilos y textos compartidos.
    cellxfs = _cellxfs_count(nuevos.get("xl/styles.xml"))
    hojas = sheet_paths(nuevos)
    total_deshechos = 0
    total_esclavas = 0
    total_duplicadas = 0
    total_estilos = 0
    total_strings = 0

    for _hoja, parte in hojas.items():
        data = nuevos.get(parte)
        if data is None:
            continue
        texto = data.decode("utf-8", "ignore")

        texto, deshechos, esclavas = _fix_shared_formulas(texto)
        texto, duplicadas = _dedupe_cells(texto)
        texto, estilos = _fix_style_refs(texto, cellxfs)
        texto, strings_data, strings = _fix_shared_strings_index(
            texto, nuevos.get("xl/sharedStrings.xml")
        )
        if strings_data is not None:
            nuevos["xl/sharedStrings.xml"] = strings_data

        nuevos[parte] = texto.encode("utf-8")
        total_deshechos += deshechos
        total_esclavas += esclavas
        total_duplicadas += duplicadas
        total_estilos += estilos
        total_strings += strings

    report.formulas_compartidas_deshechas = total_deshechos
    report.esclavas_convertidas = total_esclavas
    report.celdas_duplicadas_eliminadas = total_duplicadas
    report.estilos_corregidos = total_estilos
    report.strings_agregados = total_strings

    # 5. Verificación final: ninguna parte XML puede seguir mal formada.
    for name, data in nuevos.items():
        if not _is_xml_part(name):
            continue
        error = _xml_error(data)
        if error:
            report.problemas_restantes.append(f"{name}: {error}")

    if report.problemas_restantes:
        report.avisos.append(
            "No se generó la copia reparada porque el XML sigue siendo inválido. "
            "Envíe el informe de diagnóstico para revisar el caso."
        )
        return report

    # 6. Escritura de la copia reparada (nunca sobre el original).
    try:
        with zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED) as zf_out:
            escritas = set()
            for info in infos:
                zf_out.writestr(info, nuevos.get(info.filename, parts[info.filename]))
                escritas.add(info.filename)
            for name, data in nuevos.items():
                if name not in escritas:
                    zf_out.writestr(name, data)
        with zipfile.ZipFile(destino) as zf:
            if zf.testzip() is not None:
                raise zipfile.BadZipFile("la copia reparada no pasó la prueba del ZIP")
    except (OSError, zipfile.BadZipFile) as exc:
        report.problemas_restantes.append(f"No se pudo escribir la copia reparada: {exc}")
        return report

    # 7. Releer la copia y confirmar que ya no quedan daños graves.
    final = diagnose_xlsx(destino)
    report.salida = destino
    if final.tiene_danos_graves:
        report.problemas_restantes.extend(final.problemas)
        return report

    report.reparado = True
    try:
        informe = destino.with_name(f"{destino.stem}_diagnostico.txt")
        informe.write_text(report.detalle(), encoding="utf-8")
    except OSError:
        informe = None
    if informe:
        report.avisos.append(f"Informe guardado en: {informe}")
    return report


def _main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Diagnostica y repara (en una copia) un libro .xlsx dañado. "
            "100% local: no usa Internet."
        )
    )
    parser.add_argument("archivo", help="Ruta del archivo .xlsx a revisar")
    parser.add_argument(
        "--reparar",
        action="store_true",
        help="Genera '<nombre>_reparado.xlsx' con las correcciones",
    )
    parser.add_argument(
        "--salida",
        default=None,
        help="Ruta de la copia reparada (por defecto, junto al original)",
    )
    args = parser.parse_args(argv)

    ruta = Path(args.archivo).expanduser()
    if not ruta.exists():
        print(f"❌ No existe el archivo: {ruta}")
        return 2

    report = diagnose_xlsx(ruta)
    print("=" * 70)
    print("DIAGNÓSTICO DE INTEGRIDAD")
    print("=" * 70)
    print(report.resumen())
    for problema in report.problemas:
        print(f"  • {problema}")
    for aviso in report.avisos:
        print(f"  ↳ {aviso}")

    if not report.hay_danos:
        print("\n✅ No se detectaron daños estructurales. Si Excel sigue mostrando")
        print("   el aviso, el problema puede ser de sincronización (OneDrive) o de")
        print("   un archivo abierto/bloqueado. Pruebe copiarlo a una carpeta local.")
        return 0

    if not args.reparar:
        print("\n💡 Ejecute el mismo comando con --reparar para crear una copia")
        print("   corregida (el original no se toca).")
        return 1

    print("\n⏳ Generando copia reparada...")
    resultado = repair_xlsx(ruta, output_path=args.salida)
    print(resultado.resumen_reparacion())
    if resultado.reparado and resultado.salida:
        print(f"\n✅ Copia reparada: {resultado.salida}")
        print("   El archivo original NO se modificó.")
        print("   Abra la COPIA en Excel y compruebe que ya no pide reparación.")
    else:
        print("\n⚠️ No se pudo completar la reparación automática.")
        for problema in resultado.problemas_restantes:
            print(f"  • {problema}")
        print("   Guarde el informe de diagnóstico y compártalo para revisarlo.")
    return 0 if resultado.reparado else 1


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(_main())
