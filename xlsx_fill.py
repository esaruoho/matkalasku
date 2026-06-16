"""convey.xlsx_fill — surgical, dependency-free XLSX cell editor.

Fill specific cells in an .xlsx WITHOUT rewriting the whole workbook. We edit only
the target sheet's XML inside the zip and copy every other part byte-for-byte. This
preserves the things a full rewrite (e.g. openpyxl) silently drops: embedded images
(logos), data-validation dropdowns (the `x14` extension), cell comments, themes, and
the exact formula cells the form recomputes from. "Ready to send to the organizer"
demands that fidelity.

stdlib only (zipfile + xml.etree). The unit of work is:

    set_cells(in_path, out_path, sheet_name, {"B3": "Esa Ruoho", "E10": 280, "F10": None})

A value of None clears the cell (keeping its style). Strings are written as inline
strings so we never have to touch sharedStrings. Numbers are written bare. Cells that
already exist keep their style (`s=`); newly-created cells inherit the style of another
cell in the same column so the template's yellow input shading carries over.

FEATURE-CARD >> convey-matkalasku.feature
"""
from __future__ import annotations

import re
import shutil
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

_COL_RE = re.compile(r"^([A-Z]+)(\d+)$")


def read_cell_text(in_path, sheet_name: str, coords) -> dict:
    """Return {coord: visible text} for the given cells (shared OR inline strings).

    Read-only; used to self-check that a template matches its profile before filling.
    Numbers/blank cells come back as "" — we only care about label text here.
    """
    import html
    coords = set(coords)
    with zipfile.ZipFile(in_path) as z:
        members = {n: z.read(n) for n in z.namelist()}
    shared = []
    if "xl/sharedStrings.xml" in members:
        ss = members["xl/sharedStrings.xml"].decode("utf-8", "replace")
        for si in re.findall(r"<si>(.*?)</si>", ss, re.S):
            shared.append(html.unescape("".join(re.findall(r"<t[^>]*>(.*?)</t>", si, re.S))))
    path = _sheet_path_for(members, sheet_name)
    xml = members[path].decode("utf-8", "replace")
    out = {c: "" for c in coords}
    # match BOTH self-closing (<c r=".."/>) and full (<c r="..">..</c>) cells — otherwise
    # an empty cell earlier in a row swallows the next real cell up to its </c>
    for m in re.finditer(r'<c r="([A-Z]+\d+)"([^>]*?)(?:/>|>(.*?)</c>)', xml, re.S):
        ref, attrs, inner = m.group(1), m.group(2), (m.group(3) or "")
        if ref not in coords:
            continue
        typ = (re.search(r't="([^"]+)"', attrs) or [None, ""])[1] if 't="' in attrs else ""
        vm = re.search(r"<v>(.*?)</v>", inner, re.S)
        if typ == "s" and vm:
            try:
                out[ref] = shared[int(vm.group(1))]
            except (IndexError, ValueError):
                out[ref] = ""
        elif typ == "inlineStr":
            out[ref] = html.unescape("".join(re.findall(r"<t[^>]*>(.*?)</t>", inner, re.S)))
        elif vm:
            out[ref] = html.unescape(vm.group(1))
    return out


def read_row_heights(in_path, sheet_name: str, rows) -> dict:
    """Return {row_number: height_in_points} for the given 1-based rows (default 15)."""
    rows = set(int(r) for r in rows)
    with zipfile.ZipFile(in_path) as z:
        members = {n: z.read(n) for n in z.namelist()}
    path = _sheet_path_for(members, sheet_name)
    xml = members[path].decode("utf-8", "replace")
    out = {r: 15.0 for r in rows}
    for m in re.finditer(r'<row[^>]*r="(\d+)"([^>]*)>', xml):
        rn = int(m.group(1))
        if rn in rows:
            hm = re.search(r'ht="([0-9.]+)"', m.group(2))
            if hm:
                out[rn] = float(hm.group(1))
    return out


def col_to_index(col: str) -> int:
    """'A' -> 1, 'B' -> 2, 'AA' -> 27."""
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n


def split_coord(coord: str) -> tuple[str, int]:
    m = _COL_RE.match(coord.upper())
    if not m:
        raise ValueError(f"bad cell coordinate: {coord!r}")
    return m.group(1), int(m.group(2))


def _register_root_namespaces(xml_bytes: bytes) -> None:
    """Register every xmlns on the root element so ET round-trips prefixes intact.

    Without this, ElementTree reassigns ns0/ns1 prefixes on serialize and Excel
    rejects the file.
    """
    head = xml_bytes[:4000].decode("utf-8", "replace")
    for prefix, uri in re.findall(r'xmlns:([A-Za-z0-9_]+)="([^"]+)"', head):
        ET.register_namespace(prefix, uri)
    m = re.search(r'xmlns="([^"]+)"', head)
    if m:
        ET.register_namespace("", m.group(1))


def _q(tag: str) -> str:
    return f"{{{MAIN_NS}}}{tag}"


def _sheet_path_for(members: dict[str, bytes], sheet_name: str) -> str:
    """Resolve the worksheet XML path inside the zip for a sheet by its tab name."""
    wb = members["xl/workbook.xml"].decode("utf-8")
    rid = None
    for m in re.finditer(r"<sheet\b[^>]*/>", wb):
        tag = m.group(0)
        name_m = re.search(r'name="([^"]*)"', tag)
        rid_m = re.search(r'r:id="([^"]+)"', tag)
        if name_m and rid_m and name_m.group(1) == sheet_name:
            rid = rid_m.group(1)
            break
    if rid is None:
        raise KeyError(f"sheet {sheet_name!r} not found in workbook")
    rels = members["xl/_rels/workbook.xml.rels"].decode("utf-8")
    for m in re.finditer(r"<Relationship\b[^>]*/>", rels):
        tag = m.group(0)
        id_m = re.search(r'Id="([^"]+)"', tag)
        tgt_m = re.search(r'Target="([^"]+)"', tag)
        if id_m and tgt_m and id_m.group(1) == rid:
            target = tgt_m.group(1)
            if target.startswith("/"):
                return target.lstrip("/")
            return "xl/" + target.replace("../", "")
    raise KeyError(f"no relationship target for {rid!r}")


def _find_or_make_row(sheet_data: ET.Element, rownum: int) -> ET.Element:
    rows = sheet_data.findall(_q("row"))
    for row in rows:
        r = row.get("r")
        if r and int(r) == rownum:
            return row
    # insert in ascending row order
    new = ET.Element(_q("row"), {"r": str(rownum)})
    insert_at = len(sheet_data)
    for i, row in enumerate(rows):
        if int(row.get("r", "0")) > rownum:
            insert_at = list(sheet_data).index(row)
            break
    sheet_data.insert(insert_at, new)
    return new


def _style_for_column(sheet_data: ET.Element, col: str) -> str | None:
    """Find an existing cell's style index in the same column, to inherit shading."""
    want = col_to_index(col)
    for row in sheet_data.findall(_q("row")):
        for c in row.findall(_q("c")):
            ref = c.get("r", "")
            mm = _COL_RE.match(ref)
            if mm and col_to_index(mm.group(1)) == want and c.get("s"):
                return c.get("s")
    return None


def _find_or_make_cell(row: ET.Element, sheet_data: ET.Element, coord: str) -> ET.Element:
    col, _rn = split_coord(coord)
    want = col_to_index(col)
    cells = row.findall(_q("c"))
    for c in cells:
        if c.get("r") == coord:
            return c
    new = ET.Element(_q("c"), {"r": coord})
    style = _style_for_column(sheet_data, col)
    if style is not None:
        new.set("s", style)
    insert_at = len(row)
    for c in cells:
        cm = _COL_RE.match(c.get("r", ""))
        if cm and col_to_index(cm.group(1)) > want:
            insert_at = list(row).index(c)
            break
    row.insert(insert_at, new)
    return new


def _set_value(cell: ET.Element, value) -> None:
    # clear existing children + type
    for child in list(cell):
        cell.remove(child)
    if "t" in cell.attrib:
        del cell.attrib["t"]
    if value is None or value == "":
        return
    if isinstance(value, bool):  # avoid bool-as-int surprises
        value = int(value)
    if isinstance(value, (int, float)):
        v = ET.SubElement(cell, _q("v"))
        # render ints without trailing .0
        v.text = str(int(value)) if float(value).is_integer() else repr(value)
    else:
        cell.set("t", "inlineStr")
        is_el = ET.SubElement(cell, _q("is"))
        t_el = ET.SubElement(is_el, _q("t"))
        t_el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t_el.text = str(value)


def _strip_formula_caches(sheet_data: ET.Element) -> None:
    """Drop cached <v> results from every formula cell so apps must recompute.

    A template ships cached results from its own demo data; once we change the
    inputs those caches are stale. Removing them (paired with fullCalcOnLoad)
    guarantees no tool displays an old number.
    """
    for row in sheet_data.findall(_q("row")):
        for c in row.findall(_q("c")):
            if c.find(_q("f")) is not None:
                v = c.find(_q("v"))
                if v is not None:
                    c.remove(v)


def _set_full_calc_on_load(members: dict[str, bytes]) -> None:
    wb = members.get("xl/workbook.xml")
    if wb is None:
        return
    txt = wb.decode("utf-8")
    if "fullCalcOnLoad" in txt:
        return
    m = re.search(r"<calcPr\b[^>]*?/?>", txt)
    if m:
        tag = m.group(0)
        new = tag[:-2] + ' fullCalcOnLoad="1"/>' if tag.endswith("/>") else tag[:-1] + ' fullCalcOnLoad="1">'
        txt = txt[: m.start()] + new + txt[m.end():]
    else:
        txt = txt.replace("</workbook>", '<calcPr calcId="0" fullCalcOnLoad="1"/></workbook>')
    members["xl/workbook.xml"] = txt.encode("utf-8")


def _sheet_index(members: dict, sheet_name: str) -> int:
    wb = members["xl/workbook.xml"].decode("utf-8")
    tags = re.findall(r"<sheet\b[^>]*/>", wb)
    for i, tag in enumerate(tags):
        m = re.search(r'name="([^"]*)"', tag)
        if m and m.group(1) == sheet_name:
            return i
    return -1


def hide_sheet(members: dict, sheet_name: str) -> None:
    """Mark a sheet hidden in workbook.xml. Hidden sheets still calculate (so formula
    references survive) but are NOT exported to PDF — used to drop the country-list sheet.
    """
    wb = members["xl/workbook.xml"].decode("utf-8")

    def repl(m):
        tag = m.group(0)
        if f'name="{sheet_name}"' not in tag:
            return tag
        if "state=" in tag:
            return tag
        return tag[:-2] + ' state="hidden"/>'

    members["xl/workbook.xml"] = re.sub(r"<sheet\b[^>]*/>", repl, wb).encode("utf-8")


def set_print_area(members: dict, sheet_name: str, ref: str) -> None:
    """Define _xlnm.Print_Area for a sheet so only `ref` (e.g. $A$1:$G$40) prints —
    excludes the foreign per-diem block and the helper columns from the PDF."""
    wb = members["xl/workbook.xml"].decode("utf-8")
    idx = _sheet_index(members, sheet_name)
    if idx < 0:
        return
    quoted = sheet_name.replace("'", "''")
    new_ref = f"'{quoted}'!{ref}"
    entry = f'<definedName name="_xlnm.Print_Area" localSheetId="{idx}">{new_ref}</definedName>'
    # If a Print_Area already exists for this sheet (templates often define one for the
    # whole page), REPLACE its range — appending a second would conflict and be ignored.
    existing = re.compile(
        r'<definedName name="_xlnm\.Print_Area" localSheetId="%d">.*?</definedName>' % idx)
    if existing.search(wb):
        wb = existing.sub(entry, wb, count=1)
    elif "<definedNames>" in wb:
        wb = wb.replace("<definedNames>", "<definedNames>" + entry, 1)
    else:
        block = f"<definedNames>{entry}</definedNames>"
        if "<calcPr" in wb:
            wb = re.sub(r"(<calcPr\b)", block + r"\1", wb, count=1)
        else:
            wb = wb.replace("</workbook>", block + "</workbook>")
    members["xl/workbook.xml"] = wb.encode("utf-8")


def _fit_to_one_page(root: ET.Element) -> None:
    """Make the worksheet print on a single page (fit width & height to 1).

    Requires both <sheetPr><pageSetUpPr fitToPage="1"/> and fitToWidth/Height on
    <pageSetup>; otherwise a wide form spills onto extra pages sideways.
    """
    ns = MAIN_NS
    sheetpr = root.find(_q("sheetPr"))
    if sheetpr is None:
        sheetpr = ET.Element(_q("sheetPr"))
        root.insert(0, sheetpr)  # sheetPr must be the first child of <worksheet>
    psup = sheetpr.find(_q("pageSetUpPr"))
    if psup is None:
        psup = ET.SubElement(sheetpr, _q("pageSetUpPr"))
    psup.set("fitToPage", "1")
    ps = root.find(_q("pageSetup"))
    if ps is None:
        # insert after pageMargins if present, else append
        ps = ET.Element(_q("pageSetup"))
        margins = root.find(_q("pageMargins"))
        if margins is not None:
            root.insert(list(root).index(margins) + 1, ps)
        else:
            root.append(ps)
    ps.attrib.pop("scale", None)
    ps.set("fitToWidth", "1")
    ps.set("fitToHeight", "1")
    # drop manual page breaks (templates put one between sections → blank trailing page)
    for tag in ("rowBreaks", "colBreaks"):
        el = root.find(_q(tag))
        if el is not None:
            root.remove(el)


def _remove_drawing(root: ET.Element) -> None:
    """Drop the DrawingML <drawing> ref (the logo), keeping <legacyDrawing> (comments).
    The image part stays in the zip, just unreferenced."""
    d = root.find(_q("drawing"))
    if d is not None:
        root.remove(d)


def _strip_print_footer(root: ET.Element) -> None:
    """Remove the print <headerFooter> — it carries the 'Matkalaskupohjan sinulle tarjosi
    Virtuaaliassari.fi' branding (which is a print footer, not a cell)."""
    hf = root.find(_q("headerFooter"))
    if hf is not None:
        root.remove(hf)


_SIG_DRAWING_TMPL = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"'
    ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
    '<xdr:oneCellAnchor><xdr:from><xdr:col>{col}</xdr:col><xdr:colOff>{coloff}</xdr:colOff>'
    '<xdr:row>{row}</xdr:row><xdr:rowOff>{rowoff}</xdr:rowOff></xdr:from>'
    '<xdr:ext cx="{cx}" cy="{cy}"/><xdr:pic><xdr:nvPicPr>'
    '<xdr:cNvPr id="2" name="signature"/><xdr:cNvPicPr/></xdr:nvPicPr>'
    '<xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    ' r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>'
    '<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>'
    '<xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>'
)


def apply_signature(members: dict, sig_bytes: bytes, col: int, row: int, cx: int, cy: int,
                    coloff: int = 0, rowoff: int = 0) -> None:
    """Place a signature image by REPURPOSING the template's existing drawing plumbing:
    overwrite the logo image bytes with the signature and re-anchor drawing1 at (col,row)
    — 0-based. Reuses the sheet→drawing rel (rId3) and the drawing→image embed (rId2), so
    no new parts/content-types are needed. The old hyperlink (virtuaaliassari.fi) is dropped
    by rewriting the drawing. Caller must NOT also remove the <drawing> element.
    """
    if "xl/media/image1.png" in members:
        members["xl/media/image1.png"] = sig_bytes
    if "xl/drawings/drawing1.xml" in members:
        members["xl/drawings/drawing1.xml"] = _SIG_DRAWING_TMPL.format(
            col=col, row=row, cx=int(cx), cy=int(cy),
            coloff=int(coloff), rowoff=int(rowoff)).encode("utf-8")


def _apply_cells_to_sheet(members: dict, sheet_name: str, cells: dict) -> None:
    """Set cells on a non-primary sheet (e.g. the rate on the hidden sheet2)."""
    path = _sheet_path_for(members, sheet_name)
    _register_root_namespaces(members[path])
    root = ET.fromstring(members[path])
    sheet_data = root.find(_q("sheetData"))
    if sheet_data is None:
        return
    for coord, value in cells.items():
        _, rownum = split_coord(coord)
        row = _find_or_make_row(sheet_data, rownum)
        cell = _find_or_make_cell(row, sheet_data, coord)
        _set_value(cell, value)
    body = ET.tostring(root, encoding="unicode")
    members[path] = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
                     + body).encode("utf-8")


def set_cells(in_path, out_path, sheet_name: str, cells: dict,
              force_recalc: bool = True, hide_sheets=None, print_area=None,
              fit_to_page: bool = False, remove_drawing: bool = False,
              strip_footer: bool = False, signature=None, extra_cells=None) -> Path:
    """Set/clear `cells` (coord -> value|None) in `sheet_name`, preserving all else.

    With force_recalc (default), cached formula results in the edited sheet are
    stripped and the workbook is marked fullCalcOnLoad, so every app recomputes
    totals on open instead of showing the template's stale demo numbers.

    Returns the output Path. in_path and out_path may be the same.
    """
    in_path = Path(in_path)
    out_path = Path(out_path)
    members: dict[str, bytes] = {}
    order: list[str] = []
    with zipfile.ZipFile(in_path) as z:
        for info in z.infolist():
            order.append(info.filename)
            members[info.filename] = z.read(info.filename)

    sheet_path = _sheet_path_for(members, sheet_name)
    _register_root_namespaces(members[sheet_path])
    root = ET.fromstring(members[sheet_path])
    sheet_data = root.find(_q("sheetData"))
    if sheet_data is None:
        raise ValueError("worksheet has no <sheetData>")

    for coord, value in cells.items():
        _, rownum = split_coord(coord)
        row = _find_or_make_row(sheet_data, rownum)
        cell = _find_or_make_cell(row, sheet_data, coord)
        _set_value(cell, value)

    if force_recalc:
        _strip_formula_caches(sheet_data)
        _set_full_calc_on_load(members)

    if fit_to_page:
        _fit_to_one_page(root)
    if strip_footer:
        _strip_print_footer(root)
    if signature:
        sig_path, col, row, cx, cy = signature[:5]
        rowoff = signature[5] if len(signature) > 5 else 0
        coloff = signature[6] if len(signature) > 6 else 0
        apply_signature(members, Path(sig_path).read_bytes(), col, row, cx, cy,
                        coloff=coloff, rowoff=rowoff)
    elif remove_drawing:
        _remove_drawing(root)
    for s in (hide_sheets or []):
        hide_sheet(members, s)
    if print_area:
        set_print_area(members, print_area[0], print_area[1])

    body = ET.tostring(root, encoding="unicode")
    xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + body
    members[sheet_path] = xml.encode("utf-8")

    for sname, scells in (extra_cells or {}).items():
        _apply_cells_to_sheet(members, sname, scells)

    # write a fresh zip preserving member order
    if out_path != in_path:
        shutil.copyfile(in_path, out_path)  # not strictly needed; we rewrite below
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        for name in order:
            z.writestr(name, members[name])
    return out_path
