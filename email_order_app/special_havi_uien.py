from __future__ import annotations

import re
import xml.etree.ElementTree as ET
import unicodedata
from difflib import SequenceMatcher
from dataclasses import dataclass
from email import policy
from email.parser import BytesParser
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile


@dataclass(slots=True)
class HaviUienDestination:
    destination: str
    reference: str
    company: str
    slices: int
    cases: int

    @property
    def picking_columns(self) -> list[int]:
        if self.cases <= 0:
            return []

        chunk_count = max(1, -(-self.cases // 40))
        base_value = self.cases // chunk_count
        remainder = self.cases % chunk_count
        columns = [base_value] * chunk_count
        for index in range(remainder):
            columns[-(index + 1)] += 1
        return columns


@dataclass(slots=True)
class HaviUienOrder:
    source_file: str
    subject: str
    order_date: str
    delivery_date: str
    ve_per_pallet: str
    product_name: str
    unit_label: str
    case_label: str
    destinations: list[HaviUienDestination]
    total_slices: int
    total_cases: int
    manual_article: str = ""
    manual_description: str = ""


def parse_havi_uien_email(eml_path: str | Path) -> HaviUienOrder | None:
    path = Path(eml_path)
    with path.open("rb") as handle:
        message = BytesParser(policy=policy.default).parse(handle)

    subject = str(message.get("subject", "")).strip()
    if "HAVI DE UIEN" not in subject.upper():
        return None

    text_parts: list[str] = []
    for part in message.walk():
        if part.get_content_type() == "text/plain" and str(part.get_content_disposition()) != "attachment":
            text_parts.append(part.get_content())
    body = "\n".join(text_parts).replace("\r", "")
    lines = [line.strip() for line in body.split("\n") if line.strip()]

    destination_refs: list[tuple[str, str, str]] = []
    for line in lines:
        match = re.match(r"^([^:]+):\s*(VCSO\d+)\s*:\s*(.+)$", line)
        if match:
            destination_refs.append((match.group(1).strip(), match.group(2).strip(), match.group(3).strip()))
            continue
        no_order_match = re.match(r"^([^:]+):\s*NO ORDER$", line, re.IGNORECASE)
        if no_order_match:
            destination_refs.append((no_order_match.group(1).strip(), "NO ORDER", "Havi Logistics GmbH"))

    if not destination_refs:
        return None

    order_date = _value_after(lines, "Bestelldatum")
    ve_per_pallet = _value_after(lines, "VE/ Palette")
    delivery_date = _value_after(lines, "Lieferdatum")
    product_name = next((line for line in lines if "WRIN:" in line), "Splitting Onions")

    destination_blocks = _extract_destination_blocks(lines)
    block_map = {_normalize_destination_name(name): (name, slices, unit_label, cases, case_label) for name, slices, unit_label, cases, case_label in destination_blocks}
    destinations: list[HaviUienDestination] = []
    for destination, reference, company in destination_refs:
        block = _find_destination_block(destination, block_map)
        if block is None:
            continue
        block_name, slices, unit_label, cases, case_label = block
        destinations.append(
            HaviUienDestination(
                destination=block_name.rstrip(":") or destination.rstrip(":"),
                reference=reference,
                company=company,
                slices=slices,
                cases=cases,
            )
        )

    total_slices = 0
    total_cases = 0
    total_match = _parse_destination_block(lines, "Total")
    if total_match is not None:
        total_slices, _, total_cases, case_label = total_match
    else:
        case_label = "cases"
        total_slices = sum(item.slices for item in destinations)
        total_cases = sum(item.cases for item in destinations)

    unit_label = total_match[1] if total_match is not None else "slices"

    return HaviUienOrder(
        source_file=path.name,
        subject=subject,
        order_date=order_date,
        delivery_date=delivery_date,
        ve_per_pallet=ve_per_pallet,
        product_name=product_name,
        unit_label=unit_label,
        case_label=case_label,
        destinations=destinations,
        total_slices=total_slices,
        total_cases=total_cases,
    )


def _value_after(lines: list[str], marker: str) -> str:
    for index, line in enumerate(lines):
        if line == marker and index + 1 < len(lines):
            return lines[index + 1]
    return ""


def _parse_destination_block(lines: list[str], name: str) -> tuple[int, str, int, str] | None:
    candidates = [f"{name}:", name.rstrip(":")]
    for index, line in enumerate(lines):
        if line not in candidates:
            continue
        try:
            slices = _parse_int(lines[index + 1])
            unit_label = lines[index + 2]
            cases = _parse_int(lines[index + 3])
            case_label = lines[index + 4]
            return slices, unit_label, cases, case_label
        except (IndexError, ValueError):
            continue
    return None


def _extract_destination_blocks(lines: list[str]) -> list[tuple[str, int, str, int, str]]:
    blocks: list[tuple[str, int, str, int, str]] = []
    for index, line in enumerate(lines):
        if "VCSO" in line or line.startswith("Total"):
            continue
        if not re.match(r"^[A-Za-zÀ-ÿ\- ]+:?$", line):
            continue
        try:
            slices = _parse_int(lines[index + 1])
            unit_label = lines[index + 2]
            cases = _parse_int(lines[index + 3])
            case_label = lines[index + 4]
            if unit_label.lower() != "slices":
                continue
            if not case_label.lower().startswith("case"):
                continue
            blocks.append((line.rstrip(":"), slices, unit_label, cases, case_label))
        except (IndexError, ValueError):
            continue
    return blocks


def _normalize_destination_name(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", value)
    ascii_only = normalized.encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^a-z0-9]+", "", ascii_only.lower())


def _find_destination_block(
    destination: str,
    block_map: dict[str, tuple[str, int, str, int, str]],
) -> tuple[str, int, str, int, str] | None:
    normalized = _normalize_destination_name(destination)
    exact = block_map.get(normalized)
    if exact is not None:
        return exact

    best_key = None
    best_score = 0.0
    for key in block_map:
        score = SequenceMatcher(a=normalized, b=key).ratio()
        if score > best_score:
            best_score = score
            best_key = key

    if best_key is not None and best_score >= 0.8:
        return block_map[best_key]
    return None


def _parse_int(value: str) -> int:
    cleaned = value.replace(".", "").replace(",", "").strip()
    return int(cleaned)


def write_havi_uien_excel(output_path: str | Path, data: HaviUienOrder) -> Path:
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)

    sheet_xml = _build_sheet_xml(data)
    styles_xml = _build_styles_xml()

    with ZipFile(output, "w", ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", _content_types())
        archive.writestr("_rels/.rels", _root_rels())
        archive.writestr("xl/workbook.xml", _workbook())
        archive.writestr("xl/_rels/workbook.xml.rels", _workbook_rels())
        archive.writestr("xl/worksheets/sheet1.xml", sheet_xml)
        archive.writestr("xl/styles.xml", styles_xml)
        archive.writestr("docProps/core.xml", _core_props())
        archive.writestr("docProps/app.xml", _app_props())

    return output


def _build_sheet_xml(data: HaviUienOrder) -> bytes:
    ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    ET.register_namespace("", ns)
    worksheet = ET.Element(f"{{{ns}}}worksheet")
    max_picking_columns = max((len(item.picking_columns) for item in data.destinations), default=1)
    final_column = _column_letter(7 + max_picking_columns)
    header_row_number = 7
    first_data_row = header_row_number + 1
    last_row = header_row_number + len(data.destinations)
    ET.SubElement(worksheet, f"{{{ns}}}dimension", {"ref": f"A1:{final_column}{last_row}"})
    cols = ET.SubElement(worksheet, f"{{{ns}}}cols")
    column_widths = [18, 22, 18, 38, 18, 14, 12] + [12] * max_picking_columns
    for index, width in enumerate(column_widths, start=1):
        ET.SubElement(
            cols,
            f"{{{ns}}}col",
            {"min": str(index), "max": str(index), "width": str(width), "customWidth": "1"},
        )
    sheet_data = ET.SubElement(worksheet, f"{{{ns}}}sheetData")

    header_row: list[tuple[str, object, str]] = [
        (f"A{header_row_number}", "DC", "header"),
        (f"B{header_row_number}", "Order Number", "header"),
        (f"C{header_row_number}", "Article", "header"),
        (f"D{header_row_number}", "Description", "header"),
        (f"E{header_row_number}", "slices Quantity", "header"),
        (f"F{header_row_number}", "Quantity", "header"),
        (f"G{header_row_number}", "Unit", "header"),
    ]
    for index in range(max_picking_columns):
        header_row.append((f"{_column_letter(8 + index)}{header_row_number}", "Picking", "header"))

    rows: list[tuple[int, list[tuple[str, object, str]]]] = [
        (2, [("A2", "Customer:", "meta_label"), ("B2", "HAVI DE UIEN", "meta_value")]),
        (3, [("A3", "Delivery Date:", "meta_label"), ("B3", data.delivery_date, "meta_value")]),
        (header_row_number, header_row),
    ]

    current_row = first_data_row
    for item in data.destinations:
        row_entries: list[tuple[str, object, str]] = [
            (f"A{current_row}", item.destination, "text"),
            (f"B{current_row}", item.reference, "text"),
            (f"C{current_row}", data.manual_article, "text"),
            (f"D{current_row}", data.manual_description, "text"),
            (f"E{current_row}", item.slices, "number"),
            (f"F{current_row}", item.cases, "number"),
            (f"G{current_row}", data.case_label, "text"),
        ]
        for index in range(max_picking_columns):
            value = item.picking_columns[index] if index < len(item.picking_columns) else None
            row_entries.append((f"{_column_letter(8 + index)}{current_row}", value, "number_optional"))
        rows.append((current_row, row_entries))
        current_row += 1

    style_map = {"header": 1, "text": 2, "number": 3, "number_optional": 3, "meta_label": 4, "meta_value": 5}

    for row_number, row_entries in rows:
        if not row_entries:
            continue
        attributes = {"r": str(row_number)}
        if row_number == header_row_number:
            attributes["ht"] = "20"
            attributes["customHeight"] = "1"
        row = ET.SubElement(sheet_data, f"{{{ns}}}row", attributes)
        for cell_ref, value, style_name in row_entries:
            cell = ET.SubElement(row, f"{{{ns}}}c", {"r": cell_ref, "s": str(style_map[style_name])})
            if value is None:
                continue
            if isinstance(value, (int, float)):
                ET.SubElement(cell, f"{{{ns}}}v").text = str(value)
            else:
                cell.set("t", "inlineStr")
                is_node = ET.SubElement(cell, f"{{{ns}}}is")
                ET.SubElement(is_node, f"{{{ns}}}t").text = str(value)

    return ET.tostring(worksheet, encoding="utf-8", xml_declaration=True)


def _build_styles_xml() -> bytes:
    return b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="4">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="12"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><color rgb="FF064E2F"/><name val="Calibri"/></font>
    <font><b/><sz val="12"/><color rgb="FF062B1F"/><name val="Calibri"/></font>
  </fonts>
  <fills count="4">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFEFE1CD"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFE9F5EE"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/><diagonal/></border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="6">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>
    <xf numFmtId="1" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
    <xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment vertical="center"/></xf>
    <xf numFmtId="0" fontId="3" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment vertical="center"/></xf>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"""


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


def _workbook() -> bytes:
    return b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="HAVI DE UIEN" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"""


def _workbook_rels() -> bytes:
    return b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""


def _core_props() -> bytes:
    return b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>HAVI DE UIEN Export</dc:title>
</cp:coreProperties>"""


def _app_props() -> bytes:
    return b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Codex</Application>
</Properties>"""


def _column_letter(index: int) -> str:
    result = ""
    current = index
    while current > 0:
        current, remainder = divmod(current - 1, 26)
        result = chr(65 + remainder) + result
    return result
