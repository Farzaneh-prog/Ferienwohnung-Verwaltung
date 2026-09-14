# -*- coding: utf-8 -*-
"""
Strips pure-formatting row bloat (empty styled rows extending to ~row
1,048,576 — usually from a "select whole column, apply fill/border" edit
in Excel) out of a GästeListe_*.xlsx file, without touching any real data,
formulas, or styles. This is what made the files 28-40MB and slow to read
even though the actual reservation data is only ~30-95 rows per sheet.

Usage (always writes to a NEW file — never overwrites in place):
    python tools/shrink_xlsx.py "GästeListe_2027K.xlsx" "GästeListe_2027K.shrunk.xlsx"

Then verify the output (e.g. open it, or diff a few rows) before replacing
the original. Recommended whenever a new year's file starts feeling slow.
"""
import re
import shutil
import zipfile
import datetime

ROW_RE = re.compile(rb'<row r="(\d+)"[^>]*?(?:/>|>(.*?)</row>)', re.DOTALL)
VALUE_RE = re.compile(rb'<v>|<f[ >]')
# Only these two element *kinds* legitimately reference a row further down than
# the data itself (e.g. a merged caption, a hyperlink anchor). Deliberately NOT
# a generic ref="..." scan, which would also match <dimension ref="A1:BG.../>"
# and defeat the whole trim.
MERGECELL_RE = re.compile(rb'<mergeCell ref="[A-Z]+(\d+)(?::[A-Z]+(\d+))?"')
HYPERLINK_RE = re.compile(rb'<hyperlink ref="[A-Z]+(\d+)(?::[A-Z]+(\d+))?"')
BLOAT_THRESHOLD = 200  # only trim sheets where junk rows exceed this many rows past real data
SAFETY_BUFFER = 100    # keep this many extra empty rows past the last real one, just in case


def find_cutoff(sheet_xml: bytes) -> int:
    """Last row number that contains any real value/formula, plus a safety buffer,
    raised further if any mergeCell/hyperlink references a row past that."""
    last_real = 0
    for m in ROW_RE.finditer(sheet_xml):
        r = int(m.group(1))
        inner = m.group(2) or b""
        if VALUE_RE.search(inner):
            last_real = max(last_real, r)
    cutoff = last_real + SAFETY_BUFFER
    # Don't cut through any mergeCell / hyperlink that points further down.
    for pattern in (MERGECELL_RE, HYPERLINK_RE):
        for m in pattern.finditer(sheet_xml):
            for g in m.groups():
                if g:
                    cutoff = max(cutoff, int(g) + SAFETY_BUFFER)
    return cutoff, last_real


def trim_sheet(sheet_xml: bytes):
    cutoff, last_real = find_cutoff(sheet_xml)

    total_rows = len(list(ROW_RE.finditer(sheet_xml)))
    if total_rows == 0:
        return sheet_xml, False, last_real, cutoff

    max_row_present = max(int(m.group(1)) for m in ROW_RE.finditer(sheet_xml))
    if max_row_present - last_real < BLOAT_THRESHOLD:
        return sheet_xml, False, last_real, cutoff  # not meaningfully bloated, leave untouched

    # Rebuild sheetData with only rows <= cutoff
    def repl(m):
        r = int(m.group(1))
        return m.group(0) if r <= cutoff else b""

    new_xml = ROW_RE.sub(repl, sheet_xml)

    # Fix <dimension ref="A1:XXnnnn"/> -> cap the row number at cutoff
    def fix_dim(m):
        full = m.group(0)
        return re.sub(rb':([A-Z]+)(\d+)"', lambda mm: b':' + mm.group(1) + str(cutoff).encode() + b'"', full)

    new_xml = re.sub(rb'<dimension ref="[^"]+"/>', fix_dim, new_xml, count=1)

    return new_xml, True, last_real, cutoff


def strip_calc_chain(names, read):
    """Return (new_names_set_to_skip, content_types_bytes_patched, rels_bytes_patched) if calcChain exists."""
    has_calc = "xl/calcChain.xml" in names
    ct = read("[Content_Types].xml")
    ct = re.sub(rb'<Override PartName="/xl/calcChain\.xml"[^/]*/>', b'', ct)
    rels = read("xl/_rels/workbook.xml.rels")
    rels = re.sub(rb'<Relationship [^>]*Target="calcChain\.xml"[^>]*/>', b'', rels)
    return has_calc, ct, rels


def shrink(src_path: str, dst_path: str):
    zin = zipfile.ZipFile(src_path, "r")
    names = zin.namelist()
    read = zin.read

    has_calc, ct_patched, rels_patched = strip_calc_chain(names, read)

    report = []
    zout = zipfile.ZipFile(dst_path, "w", zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        name = item.filename
        if name == "xl/calcChain.xml" and has_calc:
            continue  # drop it; Excel/openpyxl regenerate on save/open
        data = zin.read(item.filename)
        if name == "[Content_Types].xml" and has_calc:
            data = ct_patched
        elif name == "xl/_rels/workbook.xml.rels" and has_calc:
            data = rels_patched
        elif re.match(r"xl/worksheets/sheet\d+\.xml$", name):
            new_data, trimmed, last_real, cutoff = trim_sheet(data)
            if trimmed:
                report.append((name, len(data), len(new_data), last_real, cutoff))
            data = new_data
        zout.writestr(item, data)
    zout.close()
    zin.close()
    return report


if __name__ == "__main__":
    import sys

    src = sys.argv[1]
    dst = sys.argv[2]
    report = shrink(src, dst)
    for name, old_size, new_size, last_real, cutoff in report:
        print(f"{name}: {old_size/1024/1024:.2f}MB -> {new_size/1024/1024:.2f}MB "
              f"(last_real_row={last_real}, cutoff={cutoff})")
