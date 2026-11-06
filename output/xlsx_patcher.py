"""
Edición quirúrgica de celdas sobre libros .xlsx **sin reserializar el libro**.

Motivación
----------
`openpyxl` no es un editor *round-trip* fiel: al cargar y volver a guardar un
libro descarta todo lo que no sabe representar. En los archivos oficiales de la
biblioteca eso incluye, entre otros:

  * Las listas de validación modernas (`x14:dataValidations`, dentro de `extLst`).
  * Las extensiones de hoja y de libro (`extLst`): formato condicional moderno,
    minigráficos (*sparklines*), segmentaciones (*slicers*) y líneas de tiempo.
  * Los controles de formulario / botones / cuadros de texto (`vmlDrawing`).
  * Las tablas dinámicas, los gráficos con `extLst` y parte del formato
    condicional y de los estilos `dxf`.
  * Los vínculos externos y los `definedNames` que apuntan a otros libros.

Este módulo evita el problema de raíz: **no vuelve a escribir el libro**. Abre el
`.xlsx` como lo que es (un ZIP de partes XML), localiza únicamente las celdas que
hay que modificar y las reescribe dentro de la hoja correspondiente. El resto de
partes del paquete se copian byte a byte, de modo que todo lo que no se toca queda
exactamente igual que en el original.

Solo se modifican estas partes, y únicamente cuando hace falta:

  * `xl/worksheets/sheetN.xml` → celdas concretas (valor, estilo o borde).
  * `xl/styles.xml`           → se AÑADEN fuentes/bordes/estilos; no se reescribe nada.
  * `xl/workbook.xml`         → `fullCalcOnLoad="1"` para que Excel recalcule las
                                fórmulas que dependan de las celdas modificadas.

Nunca se elimina ni se reordena ninguna parte del paquete.
"""

from __future__ import annotations

import os
import re
import tempfile
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple
from xml.sax.saxutils import escape

__all__ = [
    "XlsxPatcher",
    "BorderSpec",
    "PreservationReport",
    "verify_preservation",
    "sheet_paths",
    "column_letter",
    "column_index",
    "cell_reference",
]


# ---------------------------------------------------------------------------
# Utilidades de referencias y expresiones regulares
# ---------------------------------------------------------------------------

_REF_RE = re.compile(r"^([A-Z]+)(\d+)$")


def column_letter(index: int) -> str:
    """Convierte un índice de columna 1-based a letras (1 -> 'A', 27 -> 'AA')."""
    if index < 1:
        raise ValueError(f"Índice de columna inválido: {index}")
    letters = []
    n = index
    while n:
        n, rem = divmod(n - 1, 26)
        letters.append(chr(ord("A") + rem))
    return "".join(reversed(letters))


def column_index(letter: str) -> int:
    """Convierte letras de columna a índice 1-based ('A' -> 1, 'AA' -> 27)."""
    n = 0
    for ch in letter.strip().upper():
        if not ("A" <= ch <= "Z"):
            raise ValueError(f"Columna inválida: {letter!r}")
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n


def cell_reference(row: int, col: int) -> str:
    """Devuelve la referencia A1 de una celda (fila/columna 1-based)."""
    return f"{column_letter(col)}{row}"


# Atributos de una etiqueta XML: nombre="valor" | nombre='valor'
_ATTR_RE = r"""([\w:.\-]+)\s*=\s*("([^"]*)"|'([^']*)')"""


def _attr(attrs: str, name: str) -> Optional[str]:
    """Lee el valor de un atributo dentro del texto de atributos de una etiqueta."""
    m = re.search(r"\b%s\s*=\s*(\"([^\"]*)\"|'([^']*)')" % re.escape(name), attrs)
    if not m:
        return None
    return m.group(2) if m.group(2) is not None else m.group(3)


def _set_attr(attrs: str, name: str, value: str) -> str:
    """Fija (o añade) un atributo conservando el resto del texto de atributos."""
    pattern = re.compile(r"(\s)%s\s*=\s*(\"[^\"]*\"|'[^']*')" % re.escape(name))
    if pattern.search(attrs):
        return pattern.sub(lambda m: f'{m.group(1)}{name}="{value}"', attrs, count=1)
    return f'{attrs.rstrip()} {name}="{value}"'


def _del_attr(attrs: str, name: str) -> str:
    """Elimina un atributo si existe."""
    return re.sub(r"\s%s\s*=\s*(\"[^\"]*\"|'[^']*')" % re.escape(name), "", attrs)


# Bloques y elementos. `\b` evita confundir <c> con <cols> o <row> con <rowBreaks>.
def _block_pattern(name: str) -> "re.Pattern":
    return re.compile(
        r"<%s\b([^>]*?)(/>|>(.*?)</%s>)" % (re.escape(name), re.escape(name)), re.DOTALL
    )


_ROW_RE = re.compile(r"<row\b([^>]*?)(/>|>(.*?)</row>)", re.DOTALL)
_CELL_RE = re.compile(r"<c\b([^>]*?)(/>|>(.*?)</c>)", re.DOTALL)


# ---------------------------------------------------------------------------
# Mapa de hojas del paquete
# ---------------------------------------------------------------------------

def sheet_paths(members: Dict[str, bytes]) -> Dict[str, str]:
    """Mapea {nombre de hoja -> ruta interna} a partir de workbook.xml y sus rels."""
    mapping: Dict[str, str] = {}
    try:
        workbook_xml = members["xl/workbook.xml"].decode("utf-8", "ignore")
        rels_xml = members["xl/_rels/workbook.xml.rels"].decode("utf-8", "ignore")
    except KeyError:
        return mapping

    rel_targets: Dict[str, str] = {}
    for m in re.finditer(r"<Relationship\b[^>]*>", rels_xml):
        tag = m.group(0)
        rid = _attr(tag[1:-1], "Id")
        target = _attr(tag[1:-1], "Target")
        if not rid or not target:
            continue
        target = target.lstrip("/")
        if not target.startswith("xl/"):
            target = "xl/" + target
        rel_targets[rid] = target

    for m in re.finditer(r"<sheet\b[^>]*>", workbook_xml):
        tag = m.group(0)[1:-1]
        name = _attr(tag, "name")
        rid = _attr(tag, "r:id")
        if name and rid and rid in rel_targets:
            mapping[name] = rel_targets[rid]
    return mapping


# ---------------------------------------------------------------------------
# Gestión de estilos (solo añade, nunca reescribe lo existente)
# ---------------------------------------------------------------------------

# Orden de los cuatro lados de un borde en el esquema CT_Border.
_BORDER_SIDES = ("left", "right", "top", "bottom")


@dataclass(frozen=True)
class BorderSpec:
    """Especificación de un borde: estilos de los cuatro lados y color ARGB."""

    left: Optional[str] = "thin"
    right: Optional[str] = "thin"
    top: Optional[str] = "thin"
    bottom: Optional[str] = "thin"
    color: str = "FF000000"

    @property
    def signature(self) -> Tuple[str, ...]:
        return (self.left or "", self.right or "", self.top or "", self.bottom or "", self.color)


class _StyleBook:
    """Administra `xl/styles.xml` añadiendo fuentes, bordes y estilos de celda.

    Mantiene los índices existentes intactos (los `cellXfs` del libro original no
    se reordenan) y sólo agrega entradas nuevas al final de cada colección.
    """

    def __init__(self, xml: Optional[str]):
        self.xml = xml or ""
        self.changed = False
        self._cache_border: Dict[Tuple[str, ...], int] = {}
        self._cache_font: Dict[Tuple[str, int], int] = {}
        self._cache_xf: Dict[Tuple[int, Optional[int], Optional[int], Optional[str]], int] = {}

        self._fonts_block = _block_pattern("fonts").search(self.xml)
        self._borders_block = _block_pattern("borders").search(self.xml)
        self._xfs_block = _block_pattern("cellXfs").search(self.xml)

        self.fonts: List[str] = self._children(self._fonts_block, "font")
        self.borders: List[str] = self._children(self._borders_block, "border")
        self.xfs: List[str] = self._children(self._xfs_block, "xf")

    # -- lectura ---------------------------------------------------------
    @staticmethod
    def _children(block: Optional["re.Match"], tag: str) -> List[str]:
        if not block:
            return []
        inner = block.group(3) or ""
        pattern = re.compile(
            r"<%s\b[^>]*?/>|<%s\b[^>]*?>.*?</%s>" % (tag, tag, tag), re.DOTALL
        )
        return pattern.findall(inner)

    @property
    def available(self) -> bool:
        return bool(self.xml) and self._xfs_block is not None

    # -- alta de colecciones --------------------------------------------
    def _append(self, block_name: str, element: str, collection: List[str], block) -> int:
        """Añade un elemento al final de una colección y devuelve su índice."""
        index = len(collection)
        collection.append(element)
        self.changed = True

        if block is None:
            # No existía la colección: se crea respetando el orden del esquema.
            return index

        start, end = block.span()
        block_xml = block.group(0)
        attrs = block.group(1)
        count = _attr(attrs, "count")
        if count is not None:
            new_attrs = _set_attr(attrs, "count", str(index + 1))
            block_xml = block_xml.replace("<%s%s" % (block_name, attrs), "<%s%s" % (block_name, new_attrs), 1)
            attrs = new_attrs

        if block.group(2) == "/>":
            opened = block_xml[:-2].rstrip()
            new_block = "%s>%s</%s>" % (opened, element, block_name)
        else:
            new_block = block_xml.replace("</%s>" % block_name, element + "</%s>" % block_name, 1)

        self.xml = self.xml[:start] + new_block + self.xml[end:]
        return index

    def _ensure_collection(self, block_name: str, element_name: str) -> None:
        """Crea la colección si el libro no la tuviera (muy poco habitual)."""
        if _block_pattern(block_name).search(self.xml):
            return
        # Orden del esquema CT_Stylesheet.
        order = ["numFmts", "fonts", "fills", "borders", "cellStyleXfs", "cellXfs",
                 "cellStyles", "dxfs", "tableStyles", "colors", "extLst"]
        pos = len(self.xml)
        for tag in order[order.index(block_name) + 1:]:
            i = self.xml.find("<%s" % tag)
            if i != -1 and i < pos:
                pos = i
        if pos == len(self.xml):
            i = self.xml.rfind("</styleSheet>")
            if i != -1:
                pos = i
        new_block = '<%s count="0"></%s>' % (block_name, block_name)
        self.xml = self.xml[:pos] + new_block + self.xml[pos:]
        self.changed = True

    # -- bordes ----------------------------------------------------------
    def border_id(self, spec: BorderSpec) -> int:
        """Devuelve el índice del borde indicado, creándolo si no existe."""
        if spec.signature in self._cache_border:
            return self._cache_border[spec.signature]

        sides = []
        for side, style in zip(_BORDER_SIDES, (spec.left, spec.right, spec.top, spec.bottom)):
            if style:
                sides.append(
                    '<%s style="%s"><color rgb="%s"/></%s>' % (side, style, spec.color, side)
                )
            else:
                sides.append("<%s/>" % side)
        element = "<border>%s<diagonal/></border>" % "".join(sides)

        existing = self._find_element(self.borders, element)
        if existing is not None:
            self._cache_border[spec.signature] = existing
            return existing

        self._ensure_collection("borders", "border")
        self._borders_block = _block_pattern("borders").search(self.xml)
        index = self._append("borders", element, self.borders, self._borders_block)
        self._cache_border[spec.signature] = index
        return index

    # -- fuentes ---------------------------------------------------------
    def font_id(self, name: str, size: float) -> int:
        """Devuelve el índice de la fuente (nombre+tamaño), creándola si falta."""
        key = (name, int(size) if float(size).is_integer() else size)
        if key in self._cache_font:
            return self._cache_font[key]

        size_txt = "%g" % float(size)
        for idx, font in enumerate(self.fonts):
            name_m = re.search(r'<name[^>]*\bval="([^"]*)"', font)
            size_m = re.search(r'<sz[^>]*\bval="([^"]*)"', font)
            if not name_m or not size_m:
                continue
            if name_m.group(1) == name and size_m.group(1) in (size_txt, str(size)):
                self._cache_font[key] = idx
                return idx

        element = '<font><sz val="%s"/><name val="%s"/><family val="2"/></font>' % (
            size_txt, escape(name, {'"': "&quot;"})
        )
        self._ensure_collection("fonts", "font")
        self._fonts_block = _block_pattern("fonts").search(self.xml)
        index = self._append("fonts", element, self.fonts, self._fonts_block)
        self._cache_font[key] = index
        return index

    # -- estilos de celda ------------------------------------------------
    def xf_id(
        self,
        base: int = 0,
        border: Optional[int] = None,
        font: Optional[int] = None,
        alignment: Optional[str] = None,
    ) -> int:
        """Crea (o reutiliza) un `xf` derivado de otro, cambiando solo lo indicado.

        Clonar el `xf` de la celda conserva su fuente, relleno, alineación y
        formato numérico originales; así el cierre no altera el aspecto de la
        hoja más allá del borde solicitado.
        """
        key = (base, border, font, alignment)
        if key in self._cache_xf:
            return self._cache_xf[key]

        raw = self.xfs[base] if 0 <= base < len(self.xfs) else (
            '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        )
        new_xf = self._clone_xf(raw, border, font, alignment)

        existing = self._find_element(self.xfs, new_xf)
        if existing is not None:
            self._cache_xf[key] = existing
            return existing

        self._ensure_collection("cellXfs", "xf")
        self._xfs_block = _block_pattern("cellXfs").search(self.xml)
        index = self._append("cellXfs", new_xf, self.xfs, self._xfs_block)
        self._cache_xf[key] = index
        return index

    @staticmethod
    def _clone_xf(
        raw: str,
        border: Optional[int],
        font: Optional[int],
        alignment: Optional[str],
    ) -> str:
        m = re.match(r"<xf\b([^>]*?)(/>|>(.*?)</xf>)", raw, re.DOTALL)
        if not m:
            return '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        attrs = m.group(1)
        self_closing = m.group(2) == "/>"
        inner = "" if self_closing else (m.group(3) or "")

        if border is not None:
            attrs = _set_attr(attrs, "borderId", str(border))
            attrs = _set_attr(attrs, "applyBorder", "1")
        if font is not None:
            attrs = _set_attr(attrs, "fontId", str(font))
            attrs = _set_attr(attrs, "applyFont", "1")
        if alignment is not None:
            inner = re.sub(r"<alignment\b[^>]*/>", "", inner, flags=re.DOTALL)
            inner = re.sub(r"<alignment\b[^>]*>.*?</alignment>", "", inner, flags=re.DOTALL)
            inner = alignment + inner
            attrs = _set_attr(attrs, "applyAlignment", "1")

        if inner.strip():
            return "<xf%s>%s</xf>" % (attrs, inner)
        return "<xf%s/>" % attrs

    @staticmethod
    def _find_element(collection: Sequence[str], element: str) -> Optional[int]:
        """Busca un elemento equivalente (ignorando espacios) y devuelve su índice."""
        target = re.sub(r"\s+", "", element)
        for idx, raw in enumerate(collection):
            if re.sub(r"\s+", "", raw) == target:
                return idx
        return None


# ---------------------------------------------------------------------------
# Editor de una hoja
# ---------------------------------------------------------------------------

@dataclass
class _Cell:
    ref: str
    col: int
    attrs: str
    inner: Optional[str]          # None => etiqueta autocerrada
    raw: Optional[str]            # texto original (None => hay que re-renderizar)


@dataclass
class _Row:
    number: int
    attrs: str
    cells: Dict[int, _Cell] = field(default_factory=dict)
    raw: Optional[str] = None     # texto original (None => hay que re-renderizar)

    @property
    def dirty(self) -> bool:
        return self.raw is None


class _SheetEditor:
    """Representa una hoja y aplica cambios celda a celda conservando el resto."""

    def __init__(self, name: str, xml: str):
        self.name = name
        self.original_xml = xml
        self.xml = xml
        self.dirty = False
        self.changed_cells = 0

        block = _block_pattern("sheetData").search(xml)
        if block is None:
            raise ValueError(f"La hoja '{name}' no contiene <sheetData>")
        self._sd_span = block.span()
        self._sd_attrs = block.group(1)
        self._sd_self = block.group(2) == "/>"

        self.rows: Dict[int, _Row] = {}
        self._parse_rows(block.group(3) or "")

        # Estilos por columna (`<col style="N">`), para heredarlos al crear celdas.
        self._col_styles = self._parse_col_styles(xml)

    # -- parsing ---------------------------------------------------------
    def _parse_rows(self, inner: str) -> None:
        for m in _ROW_RE.finditer(inner):
            attrs = m.group(1)
            number = int(_attr(attrs, "r") or 0)
            if number <= 0:
                continue
            raw = m.group(0)
            row = _Row(number=number, attrs=attrs, raw=raw)
            content = "" if m.group(2) == "/>" else (m.group(3) or "")
            for cm in _CELL_RE.finditer(content):
                cattrs = cm.group(1)
                ref = _attr(cattrs, "r") or ""
                ref_m = _REF_RE.match(ref)
                if not ref_m:
                    continue
                cell = _Cell(
                    ref=ref,
                    col=column_index(ref_m.group(1)),
                    attrs=cattrs,
                    inner=None if cm.group(2) == "/>" else (cm.group(3) or ""),
                    raw=cm.group(0),
                )
                row.cells[cell.col] = cell
            # Mientras no se toque la fila se reutiliza su texto original tal cual.
            row.raw = raw
            self.rows[number] = row

    @staticmethod
    def _parse_col_styles(xml: str) -> Dict[int, int]:
        styles: Dict[int, int] = {}
        block = _block_pattern("cols").search(xml)
        if not block:
            return styles
        for m in re.finditer(r"<col\b[^>]*>", block.group(0)):
            tag = m.group(0)[1:-1]
            style = _attr(tag, "style")
            if style is None:
                continue
            try:
                lo = int(_attr(tag, "min") or 0)
                hi = int(_attr(tag, "max") or lo)
                sid = int(style)
            except ValueError:
                continue
            for c in range(max(1, lo), min(hi, 16384) + 1):
                styles[c] = sid
        return styles

    # -- edición ---------------------------------------------------------
    def _row(self, number: int) -> _Row:
        row = self.rows.get(number)
        if row is None:
            row = _Row(number=number, attrs=' r="%d"' % number, raw=None)
            self.rows[number] = row
            self.dirty = True
        return row

    def show_row(self, number: int) -> bool:
        """Quita los atributos que ocultan una fila (`hidden` y `outlineLevel`).

        Se usa al anexar registros nuevos (p. ej. certificados) en una hoja que
        tenía filas de plantilla ocultas o agrupadas justo debajo de la tabla:
        sin esto el registro se escribe en el XML, pero Excel lo muestra oculto
        y parece que "no se guardó".

        Returns:
            True si la fila estaba oculta y se hizo visible.
        """
        row = self._row(number)
        new_attrs = _del_attr(row.attrs, "hidden")
        new_attrs = _del_attr(new_attrs, "outlineLevel")
        if new_attrs == row.attrs:
            return False
        row.attrs = new_attrs
        row.raw = None
        self.dirty = True
        return True

    def _cell(self, number: int, col: int) -> _Cell:
        row = self._row(number)
        cell = row.cells.get(col)
        if cell is None:
            cell = _Cell(
                ref=cell_reference(number, col),
                col=col,
                attrs=' r="%s"' % cell_reference(number, col),
                inner=None,
                raw=None,
            )
            row.cells[col] = cell
            row.raw = None
            self.dirty = True
        return cell

    def base_style(self, row_number: int, col: int) -> int:
        """Estilo de referencia de una celda: el propio o el de su columna.

        Solo se usa para **clonar** el estilo (fuente, relleno, formato numérico…)
        y cambiar únicamente el borde. Una celda inexistente hereda el estilo de la
        columna (`<col style>`), igual que la ve Excel; nunca el de una vecina.
        """
        row = self.rows.get(row_number)
        if row is not None:
            cell = row.cells.get(col)
            if cell is not None:
                s = _attr(cell.attrs, "s")
                if s is not None:
                    try:
                        return int(s)
                    except ValueError:
                        pass
        return self._col_styles.get(col, 0)

    def set_cell(
        self,
        row: int,
        col: int,
        value: Any = None,
        has_value: bool = False,
        style: Optional[int] = None,
    ) -> None:
        """Escribe un valor y/o un estilo en una celda, creándola si no existe."""
        cell = self._cell(row, col)
        touched = False

        if has_value:
            cell.attrs = _del_attr(cell.attrs, "t")
            cell.attrs = _del_attr(cell.attrs, "cm")
            cell.attrs = _del_attr(cell.attrs, "vm")
            type_attr, cell.inner = self._serialize_value(value)
            if type_attr:
                cell.attrs = _set_attr(cell.attrs, "t", type_attr)
            touched = True

        if style is not None:
            cell.attrs = _set_attr(cell.attrs, "s", str(style))
            touched = True

        if touched:
            cell.raw = None
            self.rows[row].raw = None
            self.dirty = True
            self.changed_cells += 1

    @staticmethod
    def _serialize_value(value: Any) -> Tuple[Optional[str], Optional[str]]:
        """Devuelve (atributo `t`, contenido interno) de una celda.

        Las cadenas se escriben como `inlineStr`, de modo que no hay que tocar
        `xl/sharedStrings.xml` (se mantiene byte a byte igual que el original).
        """
        if value is None:
            return None, None
        if isinstance(value, bool):
            return "b", "<v>%d</v>" % (1 if value else 0)
        if isinstance(value, (int, float)):
            return None, "<v>%s</v>" % _format_number(value)
        text = str(value)
        if text.startswith("="):
            return None, "<f>%s</f>" % escape(text[1:])
        return "inlineStr", '<is><t xml:space="preserve">%s</t></is>' % escape(text)

    def set_row_border(self, row: int, first_col: int, last_col: int, border_id: int,
                       style_book: _StyleBook) -> None:
        """Aplica el mismo borde a todas las celdas de un rango de una fila."""
        for col in range(first_col, last_col + 1):
            base = self.base_style(row, col)
            style = style_book.xf_id(base=base, border=border_id)
            self.set_cell(row, col, style=style)

    # -- render ----------------------------------------------------------
    def render(self) -> str:
        if not self.dirty:
            return self.original_xml

        rendered_rows: List[str] = []
        for number in sorted(self.rows):
            row = self.rows[number]
            if row.raw is not None:
                rendered_rows.append(row.raw)
                continue

            attrs = row.attrs
            if row.cells:
                cols = sorted(row.cells)
                attrs = _set_attr(attrs, "spans", "%d:%d" % (cols[0], cols[-1]))
                body = "".join(self._render_cell(row.cells[c]) for c in cols)
                rendered_rows.append("<row%s>%s</row>" % (attrs, body))
            else:
                rendered_rows.append("<row%s/>" % attrs)

        if self._sd_self:
            new_sd = "<sheetData%s>%s</sheetData>" % (self._sd_attrs, "".join(rendered_rows))
        else:
            new_sd = "<sheetData%s>%s</sheetData>" % (self._sd_attrs, "".join(rendered_rows))

        start, end = self._sd_span
        xml = self.original_xml[:start] + new_sd + self.original_xml[end:]
        return self._expand_dimension(xml)

    @staticmethod
    def _render_cell(cell: _Cell) -> str:
        if cell.raw is not None:
            return cell.raw
        if cell.inner is None:
            return "<c%s/>" % cell.attrs
        return "<c%s>%s</c>" % (cell.attrs, cell.inner)

    def _expand_dimension(self, xml: str) -> str:
        """Amplía `<dimension>` si las celdas escritas superan el rango declarado."""
        if not self.rows:
            return xml
        max_row_seen = max(self.rows)
        max_col_seen = max((max(r.cells) for r in self.rows.values() if r.cells), default=1)
        m = re.search(r"<dimension\b([^>]*?)/>", xml)
        if not m:
            return xml
        ref = _attr(m.group(1), "ref")
        if not ref:
            return xml
        parts = ref.split(":")
        first = parts[0]
        last = parts[-1]
        fm = _REF_RE.match(first)
        lm = _REF_RE.match(last)
        if not fm or not lm:
            return xml
        old_max_row = int(lm.group(2))
        old_max_col = column_index(lm.group(1))
        new_max_row = max(old_max_row, max_row_seen)
        new_max_col = max(old_max_col, max_col_seen)
        if new_max_row == old_max_row and new_max_col == old_max_col:
            return xml
        new_ref = "%s:%s" % (first, cell_reference(new_max_row, new_max_col))
        attrs = _set_attr(m.group(1), "ref", new_ref)
        return xml[:m.start()] + "<dimension%s/>" % attrs + xml[m.end():]


def _format_number(value: Any) -> str:
    """Serializa un número sin notación científica ni ruido de coma flotante."""
    if isinstance(value, int):
        return str(value)
    text = format(float(value), ".15g")
    return text


# ---------------------------------------------------------------------------
# Patcher principal
# ---------------------------------------------------------------------------

class XlsxPatcher:
    """Edita celdas de un `.xlsx` conservando intacto todo lo demás del paquete."""

    def __init__(self, path: Path | str):
        self.path = Path(path).resolve()
        with zipfile.ZipFile(self.path) as zf:
            self._infos: Dict[str, zipfile.ZipInfo] = {}
            self._members: Dict[str, bytes] = {}
            self._order: List[str] = []
            for info in zf.infolist():
                self._order.append(info.filename)
                self._infos[info.filename] = info
                self._members[info.filename] = zf.read(info.filename)

        self._sheet_map = sheet_paths(self._members)
        styles_xml = self._members.get("xl/styles.xml")
        self.styles = _StyleBook(
            styles_xml.decode("utf-8", "ignore") if styles_xml is not None else None
        )
        self._editors: Dict[str, _SheetEditor] = {}
        self._recalc_needed = False
        self._recalc_modified = False

    # -- hojas -----------------------------------------------------------
    @property
    def sheet_names(self) -> List[str]:
        return list(self._sheet_map)

    def has_sheet(self, name: str) -> bool:
        return name in self._sheet_map

    def sheet(self, name: str) -> _SheetEditor:
        """Devuelve el editor de la hoja indicada (por nombre visible)."""
        if name in self._editors:
            return self._editors[name]
        part = self._sheet_map.get(name)
        if part is None:
            raise KeyError(f"La hoja '{name}' no existe en '{self.path.name}'")
        xml = self._members[part].decode("utf-8", "ignore")
        editor = _SheetEditor(name, xml)
        self._editors[name] = editor
        return editor

    # -- guardado --------------------------------------------------------
    def _apply_recalc(self) -> None:
        """Activa el recálculo al abrir para que las fórmulas reflejen los cambios."""
        if not self._recalc_needed:
            return
        wb_key = "xl/workbook.xml"
        raw = self._members.get(wb_key)
        if raw is None:
            return
        xml = raw.decode("utf-8", "ignore")
        if 'fullCalcOnLoad="1"' in xml:
            return

        m = re.search(r"<calcPr\b([^>]*?)/>", xml)
        if m:
            attrs = _set_attr(m.group(1), "fullCalcOnLoad", "1")
            xml = xml[:m.start()] + "<calcPr%s/>" % attrs + xml[m.end():]
        else:
            attrs = ' calcId="191029" fullCalcOnLoad="1"'
            after = ("<oleSize", "<customWorkbookViews", "<pivotCaches", "<smartTagPr",
                     "<smartTagTypes", "<webPublishing", "<fileRecoveryPr",
                     "<webPublishObjects", "<extLst")
            pos = len(xml)
            for tag in after:
                i = xml.find(tag)
                if i != -1 and i < pos:
                    pos = i
            if pos == len(xml):
                i = xml.rfind("</workbook>")
                if i != -1:
                    pos = i
            xml = xml[:pos] + "<calcPr%s/>" % attrs + xml[pos:]

        self._members[wb_key] = xml.encode("utf-8")
        self._recalc_modified = True

    def _workbook_has_formulas(self) -> bool:
        for name, data in self._members.items():
            if name.startswith("xl/worksheets/") and name.endswith(".xml"):
                if b"<f>" in data or b"<f " in data:
                    return True
        return False

    def save(self, force_recalc: bool = True) -> None:
        """Escribe el libro con las celdas modificadas (resto del paquete intacto)."""
        changed_parts: Dict[str, bytes] = {}

        for name, editor in self._editors.items():
            if editor.dirty:
                part = self._sheet_map[name]
                changed_parts[part] = editor.render().encode("utf-8")

        if self.styles.changed:
            changed_parts["xl/styles.xml"] = self.styles.xml.encode("utf-8")

        self._recalc_modified = False
        touched_sheets = any(
            part.startswith("xl/worksheets/") for part in changed_parts
        )
        if force_recalc and touched_sheets and self._workbook_has_formulas():
            self._recalc_needed = True
        self._apply_recalc()

        if not changed_parts and not self._recalc_modified:
            # Nada que escribir: el archivo queda exactamente como estaba.
            return

        self._write(changed_parts)

    def _write(self, changed_parts: Dict[str, bytes]) -> None:
        fd, tmp_name = tempfile.mkstemp(suffix=".xlsx", dir=str(self.path.parent))
        os.close(fd)
        tmp_path = Path(tmp_name)
        try:
            with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zf_out:
                written = set()
                for name in self._order:
                    data = changed_parts.get(name, self._members.get(name, b""))
                    info = self._infos.get(name)
                    if info is not None:
                        zf_out.writestr(info, data)
                    else:
                        zf_out.writestr(name, data)
                    written.add(name)
                for name, data in changed_parts.items():
                    if name not in written:
                        zf_out.writestr(name, data)
            try:
                os.replace(tmp_path, self.path)
            except PermissionError as exc:
                raise PermissionError(
                    f"No se pudo escribir '{self.path.name}': el archivo está abierto "
                    "en Excel (u otro programa), es de solo lectura o está bloqueado "
                    "por OneDrive/antivirus. Ciérrelo y vuelva a intentar el cierre."
                ) from exc
        finally:
            if tmp_path.exists():
                try:
                    tmp_path.unlink()
                except OSError:
                    pass


# ---------------------------------------------------------------------------
# Verificación de pérdida cero
# ---------------------------------------------------------------------------

@dataclass
class PreservationReport:
    """Verifica que aplicar el cierre no haya perdido NINGUNA parte del original."""

    archivo: str = ""
    partes_original: int = 0
    partes_resultado: int = 0
    partes_perdidas: List[str] = field(default_factory=list)
    partes_alteradas: List[str] = field(default_factory=list)
    partes_alteradas_esperadas: List[str] = field(default_factory=list)
    hojas_modificadas: List[str] = field(default_factory=list)
    validaciones_original: int = 0
    validaciones_resultado: int = 0
    validaciones_x14_original: int = 0
    validaciones_x14_resultado: int = 0
    formato_condicional_original: int = 0
    formato_condicional_resultado: int = 0
    tablas_dinamicas: int = 0
    graficos: int = 0
    imagenes: int = 0
    controles_formulario: int = 0
    nombres_definidos_original: int = 0
    nombres_definidos_resultado: int = 0
    problemas: List[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return not self.problemas

    def resumen(self) -> str:
        """Resumen de una línea apto para la consola del Modo 5."""
        icono = "✅" if self.ok else "⚠️"
        partes = (
            f"partes del libro {self.partes_resultado}/{self.partes_original} conservadas, "
            f"{len(self.partes_alteradas)} modificada(s) a propósito"
        )
        if not self.validaciones_original and not self.validaciones_x14_original:
            extra = f", formato condicional {self.formato_condicional_original}→{self.formato_condicional_resultado}"
            return f"{icono} {self.archivo}: sin listas de validación que preservar ({partes}{extra})."
        return (
            f"{icono} {self.archivo}: listas de validación estándar "
            f"{self.validaciones_original}→{self.validaciones_resultado}, "
            f"modernas (x14) {self.validaciones_x14_original}→{self.validaciones_x14_resultado}, "
            f"formato condicional {self.formato_condicional_original}→{self.formato_condicional_resultado} "
            f"({partes})"
        )


_DV_RE = re.compile(r"<dataValidation\b")
_DV_X14_RE = re.compile(r"<x14:dataValidation\b")
_CF_RE = re.compile(r"<cfRule\b")
_DEFINED_NAME_RE = re.compile(r"<definedName\b")
_VALIDATION_FP_RE = re.compile(
    r"<dataValidation\b.*?(?:/>|</dataValidation>)|<x14:dataValidation\b.*?(?:/>|</x14:dataValidation>)",
    re.DOTALL,
)
_CF_FP_RE = re.compile(r"<cfRule\b.*?(?:/>|</cfRule>)", re.DOTALL)
_DEFINED_NAME_FP_RE = re.compile(
    r"<definedName\b.*?(?:/>|</definedName>)", re.DOTALL
)


def _read_parts(path: Path) -> Dict[str, bytes]:
    with zipfile.ZipFile(path) as zf:
        return {name: zf.read(name) for name in zf.namelist()}


def _fingerprints(xml: str, pattern: "re.Pattern") -> set:
    """Huellas normalizadas de los elementos que captura `pattern`."""
    return {re.sub(r"\s+", " ", fp).strip() for fp in pattern.findall(xml)}


def _validation_fingerprints(xml: str) -> set:
    return _fingerprints(xml, _VALIDATION_FP_RE)


def verify_preservation(
    original_path: Path | str,
    output_path: Path | str,
    expected_modified_sheets: Optional[Iterable[str]] = None,
    allow_styles: bool = True,
    allow_workbook_calc: bool = True,
) -> PreservationReport:
    """Comprueba, parte por parte, que el libro resultante conserva el original.

    Compara los dos paquetes `.xlsx`:

      * Ninguna parte del original puede faltar en el resultado.
      * Ninguna parte puede cambiar salvo las hojas indicadas en
        `expected_modified_sheets` (y `styles.xml` / `workbook.xml` si se permite).
      * Cada lista de validación del original debe seguir presente (huella por
        huella), y los conteos de formato condicional y nombres definidos deben
        coincidir.

    Args:
        original_path: Copia íntegra del archivo antes del cierre (respaldo).
        output_path: Archivo ya actualizado.
        expected_modified_sheets: Nombres de hoja que el cierre modifica a propósito.
        allow_styles: Permite que cambie `xl/styles.xml` (solo se añaden estilos).
        allow_workbook_calc: Permite que cambie `xl/workbook.xml` (recálculo al abrir).

    Returns:
        `PreservationReport` con el detalle de la verificación.
    """
    original = Path(original_path)
    output = Path(output_path)
    report = PreservationReport(archivo=output.name)

    if not original.exists() or not output.exists():
        report.problemas.append(
            "No se pudo verificar la preservación (falta el archivo guardado o su respaldo)."
        )
        return report

    try:
        before = _read_parts(original)
        after = _read_parts(output)
    except Exception as exc:  # noqa: BLE001
        report.problemas.append(f"No se pudieron leer los archivos para verificar: {exc}")
        return report

    report.partes_original = len(before)
    report.partes_resultado = len(after)

    # 1. Ninguna parte puede desaparecer.
    report.partes_perdidas = sorted(set(before) - set(after))
    if report.partes_perdidas:
        report.problemas.append(
            "se perdieron partes del libro original: " + ", ".join(report.partes_perdidas)
        )

    # 2. Solo pueden cambiar las partes previstas.
    allowed = set()
    names = sheet_paths(before)
    for sheet in expected_modified_sheets or []:
        part = names.get(sheet)
        if part:
            allowed.add(part)
            report.hojas_modificadas.append(sheet)
    if allow_styles:
        allowed.add("xl/styles.xml")
    if allow_workbook_calc:
        allowed.add("xl/workbook.xml")

    changed = [name for name in before if name in after and before[name] != after[name]]
    report.partes_alteradas = sorted(changed)
    report.partes_alteradas_esperadas = sorted(n for n in changed if n in allowed)
    unexpected = sorted(n for n in changed if n not in allowed)
    if unexpected:
        report.problemas.append(
            "se alteraron partes que debían quedar intactas: " + ", ".join(unexpected)
        )

    # 3. Listas de validación: conteo global y huella por hoja.
    for name, data in before.items():
        if not (name.startswith("xl/worksheets/") and name.endswith(".xml")):
            continue
        xml_before = data.decode("utf-8", "ignore")
        xml_after = after.get(name, b"").decode("utf-8", "ignore")
        report.validaciones_original += len(_DV_RE.findall(xml_before))
        report.validaciones_x14_original += len(_DV_X14_RE.findall(xml_before))
        report.formato_condicional_original += len(_CF_RE.findall(xml_before))
        report.validaciones_resultado += len(_DV_RE.findall(xml_after))
        report.validaciones_x14_resultado += len(_DV_X14_RE.findall(xml_after))
        report.formato_condicional_resultado += len(_CF_RE.findall(xml_after))

        missing = _validation_fingerprints(xml_before) - _validation_fingerprints(xml_after)
        if missing:
            report.problemas.append(
                f"'{name}': {len(missing)} lista(s) de validación de datos no están en el resultado"
            )

        missing_cf = _fingerprints(xml_before, _CF_FP_RE) - _fingerprints(xml_after, _CF_FP_RE)
        if missing_cf:
            report.problemas.append(
                f"'{name}': {len(missing_cf)} regla(s) de formato condicional no están en el resultado"
            )

    # 4. Nombres definidos del libro (origen de muchas listas desplegables).
    wb_before = before.get("xl/workbook.xml", b"").decode("utf-8", "ignore")
    wb_after = after.get("xl/workbook.xml", b"").decode("utf-8", "ignore")
    report.nombres_definidos_original = len(_DEFINED_NAME_RE.findall(wb_before))
    report.nombres_definidos_resultado = len(_DEFINED_NAME_RE.findall(wb_after))
    missing_names = _fingerprints(wb_before, _DEFINED_NAME_FP_RE) - _fingerprints(
        wb_after, _DEFINED_NAME_FP_RE
    )
    if missing_names:
        report.problemas.append(
            f"se perdieron nombres definidos: {report.nombres_definidos_original}"
            f"→{report.nombres_definidos_resultado}"
        )

    # 5. Inventario de elementos "frágiles" (deben sobrevivir al 100%).
    report.tablas_dinamicas = sum(
        1 for n in after if n.startswith("xl/pivotTables/") or n.startswith("xl/pivotCache/")
    )
    report.graficos = sum(1 for n in after if n.startswith("xl/charts/") and n.endswith(".xml"))
    report.imagenes = sum(1 for n in after if n.startswith("xl/media/"))
    report.controles_formulario = sum(
        1 for n in after if n.startswith("xl/worksheets/") and n.endswith(".vml")
    )
    for label, prefix in ((
        "tablas dinámicas", ("xl/pivotTables/", "xl/pivotCache/"),
    ), (
        "gráficos", ("xl/charts/",),
    ), (
        "imágenes", ("xl/media/",),
    ), (
        "controles de formulario", ("xl/worksheets/",),
    )):
        if label == "controles de formulario":
            lost = [
                n for n in before
                if n.startswith(prefix) and n.endswith(".vml") and n not in after
            ]
        else:
            lost = [n for n in before if n.startswith(prefix) and n not in after]
        if lost:
            report.problemas.append(f"se perdieron {label}: " + ", ".join(sorted(lost)))

    return report
