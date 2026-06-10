from __future__ import annotations

import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
ET.register_namespace("", MAIN_NS)
ET.register_namespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")


@dataclass(slots=True)
class WorkbookCell:
    value: str | int | float | None
    style: str = "body"


STYLE_IDS = {
    "body": 0,
    "header": 1,
    "meta": 2,
    "number": 3,
}


def write_simple_workbook(
    output_path: str | Path,
    sheet_name: str,
    rows: list[list[WorkbookCell]],
    title: str = "Export",
) -> Path:
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)

    sheet_xml = _build_sheet_xml(rows)

    with ZipFile(output, "w", ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", _content_types())
        archive.writestr("_rels/.rels", _root_rels())
        archive.writestr("xl/workbook.xml", _workbook(sheet_name))
        archive.writestr("xl/_rels/workbook.xml.rels", _workbook_rels())
        archive.writestr("xl/worksheets/sheet1.xml", sheet_xml)
        archive.writestr("xl/styles.xml", _styles())
        archive.writestr("docProps/core.xml", _core_props(title))
        archive.writestr("docProps/app.xml", _app_props(sheet_name))

    return output


def _build_sheet_xml(rows: list[list[WorkbookCell]]) -> bytes:
    worksheet = ET.Element(f"{{{MAIN_NS}}}worksheet")
    last_row = max(1, len(rows))
    last_col = max((len(row) for row in rows), default=1)
    ET.SubElement(worksheet, f"{{{MAIN_NS}}}dimension", {"ref": f"A1:{_column_letter(last_col)}{last_row}"})
    sheet_data = ET.SubElement(worksheet, f"{{{MAIN_NS}}}sheetData")

    for row_index, row_values in enumerate(rows, start=1):
        row_node = ET.SubElement(sheet_data, f"{{{MAIN_NS}}}row", {"r": str(row_index)})
        for col_index, cell_data in enumerate(row_values, start=1):
            cell_ref = f"{_column_letter(col_index)}{row_index}"
            style_id = str(STYLE_IDS.get(cell_data.style, 0))
            cell = ET.SubElement(row_node, f"{{{MAIN_NS}}}c", {"r": cell_ref, "s": style_id})
            if cell_data.value is None:
                continue
            if isinstance(cell_data.value, (int, float)):
                ET.SubElement(cell, f"{{{MAIN_NS}}}v").text = _format_number(cell_data.value)
            else:
                cell.set("t", "inlineStr")
                is_node = ET.SubElement(cell, f"{{{MAIN_NS}}}is")
                text_node = ET.SubElement(is_node, f"{{{MAIN_NS}}}t")
                if cell_data.value.startswith(" ") or cell_data.value.endswith(" "):
                    text_node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                text_node.text = str(cell_data.value)

    return ET.tostring(worksheet, encoding="utf-8", xml_declaration=True)


def _format_number(value: int | float) -> str:
    numeric = float(value)
    if numeric.is_integer():
        return str(int(numeric))
    return f"{numeric:.2f}".rstrip("0").rstrip(".")


def _column_letter(index: int) -> str:
    result = ""
    current = index
    while current > 0:
        current, remainder = divmod(current - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _content_types() -> bytes:
    return b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>"""


def _root_rels() -> bytes:
    return b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>"""


def _workbook(sheet_name: str) -> bytes:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="{MAIN_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="{_escape_xml(sheet_name)}" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>""".encode("utf-8")


def _workbook_rels() -> bytes:
    return b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""


def _styles() -> bytes:
    return b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="3">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="12"/><name val="Calibri"/></font>
  </fonts>
  <fills count="4">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFF5ECDF"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFEEE3D2"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/><diagonal/></border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="4">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
    <xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>
    <xf numFmtId="1" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"""


def _core_props(title: str) -> bytes:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>{_escape_xml(title)}</dc:title>
</cp:coreProperties>""".encode("utf-8")


def _app_props(sheet_name: str) -> bytes:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Codex</Application>
  <TitlesOfParts>
    <vt:vector size="1" baseType="lpstr">
      <vt:lpstr>{_escape_xml(sheet_name)}</vt:lpstr>
    </vt:vector>
  </TitlesOfParts>
</Properties>""".encode("utf-8")


def _escape_xml(value: str) -> str:
    return (
        str(value)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )
