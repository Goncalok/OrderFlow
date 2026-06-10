from __future__ import annotations

import copy
import re
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any
from zipfile import ZIP_DEFLATED, ZipFile

from .models import CustomerOrder, ParsedOrderEmail
from .simple_excel import WorkbookCell, write_simple_workbook


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS = {"x": MAIN_NS}
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
WORKBOOK_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
ET.register_namespace("", MAIN_NS)
ET.register_namespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
ET.register_namespace("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006")
ET.register_namespace("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac")
ET.register_namespace("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision")
ET.register_namespace("xr2", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2")
ET.register_namespace("xr3", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3")

OPTIONAL_ATTR_NAMESPACES = {
    "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
    "http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
    "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2",
    "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3",
}


def fill_template(
    template_path: str | Path,
    output_path: str | Path,
    email_data: ParsedOrderEmail,
    customer_order: CustomerOrder,
    worksheet_name: str | None = None,
) -> Path:
    template = Path(template_path)
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)

    if not template.exists():
        return _write_fallback_workbook(output, email_data, customer_order)

    with ZipFile(template, "r") as source_archive:
        file_map = {name: source_archive.read(name) for name in source_archive.namelist()}

    shared_strings_root = ET.fromstring(file_map["xl/sharedStrings.xml"])
    worksheet_path = _resolve_worksheet_path(file_map, worksheet_name)
    sheet_root = ET.fromstring(file_map[worksheet_path])
    if worksheet_name:
        _keep_only_selected_worksheet(file_map, worksheet_name)

    start_row = 4
    # Auto-detect how many item rows this sheet has (e.g. 18 for Paderborn,
    # 20 for Rosbach, 21 for Venlo) so the footer is never placed in the
    # middle of the item area.
    base_capacity = _find_sheet_capacity(sheet_root, shared_strings_root, start_row)

    ensure_item_capacity(sheet_root, len(customer_order.items), base_capacity)
    normalize_item_grid(sheet_root, max(base_capacity, len(customer_order.items)))

    # ── Row 1 / Row 2 header ────────────────────────────────────────────────
    # Template layout (ORDER REWE PENNY Print In Color.xlsx):
    #   A1  "Delivery Date to DC: ..."  ← full label+value in one cell
    #   C1  "Fatrans Dispo: Paderborn"  ← pre-filled per sheet, do NOT touch
    #   E1  "Customer DC:"              ← static label already in template
    #   F1  <customer label>             ← we write the customer label here
    #   G1  "Leaving Date Venlo:"       ← static label already in template
    #   J1  <date value>                ← we write just the date string here
    #   G2  "Leaving Time Venlo:"       ← static label already in template
    #   J2  <time value>                ← we write just the time string here
    set_string_cell(sheet_root, shared_strings_root, "A1", f"Delivery Date to DC: {email_data.delivery_date_to_dc}")
    set_string_cell(sheet_root, shared_strings_root, "F1", _build_customer_cell_label(customer_order))
    set_string_cell(sheet_root, shared_strings_root, "D3", "Unit")

    try:
        from datetime import datetime
        d = datetime.strptime(email_data.leaving_date_venlo.strip(), "%d-%m-%Y")
        excel_date = (d - datetime(1899, 12, 30)).days
        set_numeric_cell(sheet_root, "J1", excel_date)
    except Exception:
        set_string_cell(sheet_root, shared_strings_root, "J1", email_data.leaving_date_venlo)

    j2_cell = _find_cell(sheet_root, "J2")
    j2_is_empty = True
    if j2_cell is not None:
        v = j2_cell.find("x:v", NS)
        if v is not None and v.text:
            j2_is_empty = False

    if j2_is_empty:
        try:
            from datetime import datetime
            t = datetime.strptime(email_data.leaving_time_venlo.strip(), "%H:%M")
            excel_time = t.hour / 24.0 + t.minute / 1440.0
            set_numeric_cell(sheet_root, "J2", excel_time)
        except Exception:
            set_string_cell(sheet_root, shared_strings_root, "J2", email_data.leaving_time_venlo)

    max_rows = max(base_capacity, len(customer_order.items))
    footer_row = start_row + max_rows

    # ── Footer labels — columns A / E / G (matching the template) ──────────
    set_string_cell(sheet_root, shared_strings_root, f"A{footer_row}", "Picker Name:")
    set_string_cell(sheet_root, shared_strings_root, f"E{footer_row}", "Checker Name:")
    set_string_cell(sheet_root, shared_strings_root, f"G{footer_row}", "Loader Name :")
    set_string_cell(sheet_root, shared_strings_root, f"A{footer_row + 1}", "Pallets:")
    set_string_cell(sheet_root, shared_strings_root, f"E{footer_row + 1}", "Pallets:")
    set_string_cell(sheet_root, shared_strings_root, f"G{footer_row + 1}", "Pallets:")
    set_string_cell(sheet_root, shared_strings_root, f"A{footer_row + 2}", "Pallets Places:")
    set_string_cell(sheet_root, shared_strings_root, f"E{footer_row + 2}", "Pallets Places:")
    set_string_cell(sheet_root, shared_strings_root, f"G{footer_row + 2}", "Pallets Places:")

    for offset in range(max_rows):
        row_number = start_row + offset
        if offset < len(customer_order.items):
            item = customer_order.items[offset]
            set_string_cell(sheet_root, shared_strings_root, f"A{row_number}", item.article_number)
            set_string_cell(sheet_root, shared_strings_root, f"B{row_number}", item.description)
            set_numeric_cell(sheet_root, f"C{row_number}", item.quantity_boxes)
            set_string_cell(sheet_root, shared_strings_root, f"D{row_number}", item.unit)
        else:
            clear_cell_value(sheet_root, f"A{row_number}")
            clear_cell_value(sheet_root, f"B{row_number}")
            clear_cell_value(sheet_root, f"C{row_number}")
            clear_cell_value(sheet_root, f"D{row_number}")

    sanitize_optional_namespace_attributes(sheet_root)
    file_map[worksheet_path] = ET.tostring(sheet_root, encoding="utf-8", xml_declaration=False)
    
    sanitize_optional_namespace_attributes(shared_strings_root)
    shared_strings_root.set("count", str(_count_shared_string_references_in_workbook(file_map)))
    shared_strings_root.set("uniqueCount", str(len(shared_strings_root.findall("x:si", NS))))
    file_map["xl/sharedStrings.xml"] = ET.tostring(shared_strings_root, encoding="utf-8", xml_declaration=False)

    with ZipFile(output, "w", ZIP_DEFLATED) as target_archive:
        for name, content in file_map.items():
            target_archive.writestr(name, content)

    return output


def _write_fallback_workbook(output: Path, email_data: ParsedOrderEmail, customer_order: CustomerOrder) -> Path:
    rows: list[list[WorkbookCell]] = [
        [WorkbookCell("Delivery Date to DC", "meta"), WorkbookCell(email_data.delivery_date_to_dc, "meta")],
        [WorkbookCell("DC", "meta"), WorkbookCell(customer_order.fatrans_dc, "meta")],
        [WorkbookCell("Customer DC", "meta"), WorkbookCell(customer_order.customer_dc, "meta")],
        [WorkbookCell("Leaving Date Venlo", "meta"), WorkbookCell(email_data.leaving_date_venlo, "meta")],
        [WorkbookCell("Leaving Time Venlo", "meta"), WorkbookCell(email_data.leaving_time_venlo, "meta")],
        [],
        [
            WorkbookCell("Article", "header"),
            WorkbookCell("Description", "header"),
            WorkbookCell("Quantity", "header"),
            WorkbookCell("Unit", "header"),
        ],
    ]

    for item in customer_order.items:
        rows.append(
            [
                WorkbookCell(item.article_number),
                WorkbookCell(item.description),
                WorkbookCell(item.quantity_boxes, "number"),
                WorkbookCell(item.unit),
            ]
        )

    return write_simple_workbook(output, "Order Export", rows, title="Order Export")


def _build_customer_cell_label(order: CustomerOrder) -> str:
    customer_label = str(order.customer_dc or "").strip()
    client = str(order.fatrans_dc or "").strip().lower()
    if client in {"penny", "rewe"}:
        cleaned_customer = re.sub(rf'^{client}\s+', '', customer_label, flags=re.IGNORECASE).strip()
        if cleaned_customer:
            return f"{client.capitalize()} {cleaned_customer}"
        return client.capitalize()
    return customer_label


def _resolve_worksheet_path(file_map: dict[str, bytes], worksheet_name: str | None) -> str:
    if not worksheet_name:
        return "xl/worksheets/sheet1.xml"

    workbook_root = ET.fromstring(file_map["xl/workbook.xml"])
    rels_root = ET.fromstring(file_map["xl/_rels/workbook.xml.rels"])
    rel_map = {
        rel.get("Id"): rel.get("Target")
        for rel in list(rels_root)
        if rel.tag == f"{{{REL_NS}}}Relationship"
    }

    for sheet in workbook_root.findall(f".//{{{MAIN_NS}}}sheet"):
        if sheet.get("name") != worksheet_name:
            continue
        rel_id = sheet.get(f"{{{WORKBOOK_REL_NS}}}id")
        target = rel_map.get(rel_id)
        if not target:
            break
        if target.startswith("/"):
            return target.lstrip("/")
        if target.startswith("xl/"):
            return target
        return f"xl/{target}"

    raise ValueError(f"Worksheet {worksheet_name} was not found in the template.")


def _keep_only_selected_worksheet(file_map: dict[str, bytes], worksheet_name: str) -> None:
    workbook_root = ET.fromstring(file_map["xl/workbook.xml"])
    sheets_node = workbook_root.find(f".//{{{MAIN_NS}}}sheets")
    selected_index = 0
    if sheets_node is not None:
        for index, sheet in enumerate(list(sheets_node)):
            is_selected = sheet.get("name") == worksheet_name
            if is_selected:
                selected_index = index
                sheet.attrib.pop("state", None)
            else:
                sheet.set("state", "hidden")

    workbook_view = workbook_root.find(f".//{{{MAIN_NS}}}workbookView")
    if workbook_view is not None:
        workbook_view.set("activeTab", str(selected_index))

    target_path = _resolve_worksheet_path(file_map, worksheet_name)
    for worksheet_path in [name for name in file_map if name.startswith("xl/worksheets/sheet") and name.endswith(".xml")]:
        sheet_root = ET.fromstring(file_map[worksheet_path])
        sheet_view = sheet_root.find(f".//{{{MAIN_NS}}}sheetView")
        if sheet_view is not None:
            if worksheet_path == target_path:
                sheet_view.set("tabSelected", "1")
            else:
                sheet_view.attrib.pop("tabSelected", None)
        file_map[worksheet_path] = ET.tostring(sheet_root, encoding="utf-8", xml_declaration=True)

    file_map["xl/workbook.xml"] = ET.tostring(workbook_root, encoding="utf-8", xml_declaration=True)


def _find_sheet_capacity(sheet_root: ET.Element, shared_strings_root: ET.Element, start_row: int = 4) -> int:
    """Return the number of item rows in the template sheet by locating the
    first 'Picker Name:' (or similar) cell in column A after *start_row*.
    Falls back to 18 when the footer cannot be detected."""
    shared_strings: list[str] = [
        "".join((t.text or "") for t in item.iterfind(".//x:t", NS))
        for item in shared_strings_root.findall("x:si", NS)
    ]
    for row in sheet_root.findall(".//x:sheetData/x:row", NS):
        rn = int(row.get("r", "0"))
        if rn <= start_row:
            continue
        cell = _find_cell(sheet_root, f"A{rn}")
        if cell is None:
            continue
        if cell.get("t") != "s":
            continue
        v = cell.find("x:v", NS)
        if v is None or v.text is None:
            continue
        try:
            val = shared_strings[int(v.text)]
        except (ValueError, IndexError):
            continue
        if "picker" in val.lower():
            return rn - start_row
    return 18


def ensure_item_capacity(sheet_root: ET.Element, item_count: int, base_capacity: int = 18) -> None:
    threshold = 4 + base_capacity  # first row after the item area (= footer row in the blank template)
    if item_count <= base_capacity:
        return

    extra_rows = item_count - base_capacity
    sheet_data = sheet_root.find("x:sheetData", NS)
    if sheet_data is None:
        raise ValueError("The template is missing sheet data.")

    rows = sheet_data.findall("x:row", NS)
    template_item_row = next((row for row in rows if row.get("r") == "4"), None)
    if template_item_row is None:
        raise ValueError("The template item row could not be found.")

    new_rows: list[ET.Element] = []
    for row in rows:
        row_number = int(row.get("r"))
        if row_number >= threshold:
            _shift_row(row, extra_rows)
        new_rows.append(row)

    insertion_index = next((index for index, row in enumerate(new_rows) if int(row.get("r")) >= threshold + extra_rows), len(new_rows))
    clones = [_clone_row(template_item_row, threshold + offset) for offset in range(extra_rows)]
    for offset, clone in enumerate(clones):
        new_rows.insert(insertion_index + offset, clone)

    for row in list(sheet_data):
        sheet_data.remove(row)
    for row in new_rows:
        sheet_data.append(row)

    _shift_merged_ranges(sheet_root, extra_rows, threshold)
    _update_dimension(sheet_root, extra_rows)


def _clone_row(template_row: ET.Element, new_row_number: int) -> ET.Element:
    cloned = copy.deepcopy(template_row)
    cloned.set("r", str(new_row_number))
    for cell in cloned.findall("x:c", NS):
        cell_ref = cell.get("r")
        if not cell_ref:
            continue
        column = re.match(r"[A-Z]+", cell_ref).group(0)
        cell.set("r", f"{column}{new_row_number}")
    return cloned


def normalize_item_grid(sheet_root: ET.Element, total_rows: int) -> None:
    sheet_data = sheet_root.find("x:sheetData", NS)
    if sheet_data is None:
        return

    template_row = next((row for row in sheet_data.findall("x:row", NS) if row.get("r") == "4"), None)
    if template_row is None:
        return

    template_cells = {cell.get("r"): cell for cell in template_row.findall("x:c", NS)}
    style_a = template_cells.get("A4").get("s")
    style_b = template_cells.get("B4").get("s")
    style_c = template_cells.get("C4").get("s")
    style_d = template_cells.get("D4").get("s")
    row_height = template_row.get("ht")
    custom_height = template_row.get("customHeight")

    for row_number in range(4, 4 + total_rows):
        row = next((item for item in sheet_data.findall("x:row", NS) if item.get("r") == str(row_number)), None)
        if row is None:
            continue

        if row_height:
            row.set("ht", row_height)
        if custom_height:
            row.set("customHeight", custom_height)

        for cell in row.findall("x:c", NS):
            cell_ref = cell.get("r", "")
            column = re.match(r"[A-Z]+", cell_ref)
            if not column:
                continue
            if column.group(0) == "A" and style_a is not None:
                cell.set("s", style_a)
            elif column.group(0) == "B" and style_b is not None:
                cell.set("s", style_b)
            elif column.group(0) == "C" and style_c is not None:
                cell.set("s", style_c)
            elif column.group(0) == "D" and style_d is not None:
                cell.set("s", style_d)


def _shift_row(row: ET.Element, offset: int) -> None:
    new_row_number = int(row.get("r")) + offset
    row.set("r", str(new_row_number))
    for cell in row.findall("x:c", NS):
        cell_ref = cell.get("r")
        if not cell_ref:
            continue
        column = re.match(r"[A-Z]+", cell_ref).group(0)
        cell.set("r", f"{column}{new_row_number}")


# CMR cell layout per template variant:
# "fif": references start at E36 (2 rows), emballage items E39-E42, PP at E43
# "kdc": references start at E35 (2 rows), emballage items E38-E41, PP at E41 (but row 41 is last emballage item)
_CMR_VARIANT_MAP: dict[str, dict[str, object]] = {
    "fif": {"ref_start_row": 36, "pp_cell": "E43", "emballage_start_row": 39, "emballage_end_row": 42},
    "kdc": {"ref_start_row": 35, "pp_cell": "E41", "emballage_start_row": 38, "emballage_end_row": 41},
}


def fill_cmr_template(
    template_path: str | Path,
    output_path: str | Path,
    email_data: ParsedOrderEmail,
    customer_order: CustomerOrder,
    references: list[str] | None = None,
    pallet_places: str | None = None,
    variant: str = "fif",
    pakbon_items: list[dict[str, Any]] | None = None,
) -> Path:
    template = Path(template_path)
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)

    if not template.exists():
        return _write_fallback_cmr_workbook(output, email_data, customer_order)

    with ZipFile(template, "r") as source_archive:
        file_map = {name: source_archive.read(name) for name in source_archive.namelist()}

    # Find CMR sheet path
    shared_strings_root = ET.fromstring(file_map["xl/sharedStrings.xml"])
    try:
        cmr_sheet_path = _resolve_worksheet_path(file_map, "CMR")
    except ValueError:
        return _write_fallback_cmr_workbook(output, email_data, customer_order)

    sheet_root = ET.fromstring(file_map[cmr_sheet_path])

    # Resolve cell layout for this variant
    layout = _CMR_VARIANT_MAP.get(variant, _CMR_VARIANT_MAP["fif"])
    ref_start_row: int = layout["ref_start_row"]
    pp_cell: str = layout["pp_cell"]
    emballage_start_row: int = layout.get("emballage_start_row", 39)
    emballage_end_row: int = layout.get("emballage_end_row", 42)

    # Fill delivery date
    set_string_cell(sheet_root, shared_strings_root, "C25", email_data.delivery_date_to_dc)

    # Fill references (up to 2)
    reference_values = _normalize_cmr_references(references, customer_order.reference)
    for row_offset in range(2):
        clear_cell_value(sheet_root, f"E{ref_start_row + row_offset}")
    for offset, reference in enumerate(reference_values[:2]):
        set_string_cell(sheet_root, shared_strings_root, f"E{ref_start_row + offset}", reference)

    # Fill pakbon items if provided (DELIVERED EMBALLAGE section)
    # The template has pre-filled descriptions in column F
    # We need to match pakbon items with template descriptions and fill quantities in column E
    # Note: For FIF, row 43 is PP. For KDC, E41 is both last emballage item AND PP cell
    if pakbon_items:
        # Build a map of pakbon items by description (case-insensitive)
        pakbon_map = {}
        for item in pakbon_items:
            desc = item.get("description", "").strip().lower()
            if desc:
                pakbon_map[desc] = item.get("quantity", 0)
        
        # Read template descriptions from column F and match with pakbon items
        # Use the configured emballage row range for this variant
        for row in range(emballage_start_row, emballage_end_row + 1):
            # Get the description from column F using shared strings
            cell = _find_cell(sheet_root, f"F{row}")
            if cell is None:
                continue
            
            # Extract description from shared strings
            template_desc = ""
            if cell.get("t") == "s":  # Shared string
                v = cell.find("x:v", NS)
                if v is not None and v.text is not None:
                    try:
                        shared_strings = [
                            "".join((t.text or "") for t in item.iterfind(".//x:t", NS))
                            for item in shared_strings_root.findall("x:si", NS)
                        ]
                        template_desc = shared_strings[int(v.text)]
                    except (ValueError, IndexError):
                        pass
            else:  # Inline string
                v = cell.find("x:v", NS)
                if v is not None and v.text:
                    template_desc = v.text
            
            # Match with pakbon items (case-insensitive)
            template_desc_lower = template_desc.strip().lower()
            if template_desc_lower in pakbon_map:
                quantity = pakbon_map[template_desc_lower]
                try:
                    set_numeric_cell(sheet_root, f"E{row}", float(quantity))
                except (ValueError, TypeError):
                    set_string_cell(sheet_root, shared_strings_root, f"E{row}", str(quantity))
    
    # Fill pallet places AFTER pakbon items (to ensure PP cell is not overwritten by pakbon data)
    # This is especially important for KDC where E41 is both last emballage row and PP cell
    if pallet_places:
        try:
            set_numeric_cell(sheet_root, pp_cell, float(str(pallet_places).replace(",", ".")))
        except ValueError:
            set_string_cell(sheet_root, shared_strings_root, pp_cell, str(pallet_places))

    sanitize_optional_namespace_attributes(sheet_root)
    file_map[cmr_sheet_path] = ET.tostring(sheet_root, encoding="utf-8", xml_declaration=False)
    
    sanitize_optional_namespace_attributes(shared_strings_root)
    shared_strings_root.set("count", str(_count_shared_string_references_in_workbook(file_map)))
    shared_strings_root.set("uniqueCount", str(len(shared_strings_root.findall("x:si", NS))))
    file_map["xl/sharedStrings.xml"] = ET.tostring(shared_strings_root, encoding="utf-8", xml_declaration=False)

    with ZipFile(output, "w", ZIP_DEFLATED) as target_archive:
        for name, content in file_map.items():
            target_archive.writestr(name, content)

    return output


_COLRUYT_CMR_DATE_CELL = "C27"
_COLRUYT_CMR_REF_CELL = "E37"
_COLRUYT_CMR_PP_CELL = "E40"
_COLRUYT_CMR_GOEDEREN_TOTAL_CELL = "E41"  # Total from pakbon Goederen section
_COLRUYT_CMR_SLOTBOEKING_CELL = "E58"
_COLRUYT_CMR_REF_PREFIX = "HESSING ORDERNO. "


def fill_colruyt_cmr_template(
    template_path: str | Path,
    output_path: str | Path,
    email_data: ParsedOrderEmail,
    customer_order: CustomerOrder,
    references: list[str] | None = None,
    pallet_places: str | None = None,
    goederen_total: float | None = None,
) -> Path:
    """Fill the Colruyt CMR xlsx template (C27=date, E37=ref, E40=PP, E41=Goederen total)."""
    template = Path(template_path)
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)

    if not template.exists():
        return _write_fallback_cmr_workbook(output, email_data, customer_order)

    with ZipFile(template, "r") as source_archive:
        file_map = {name: source_archive.read(name) for name in source_archive.namelist()}

    # Build/load shared strings (create empty root if missing)
    if "xl/sharedStrings.xml" in file_map:
        shared_strings_root = ET.fromstring(file_map["xl/sharedStrings.xml"])
    else:
        shared_strings_root = ET.Element(f"{{{MAIN_NS}}}sst")

    # Dynamically resolve the CMR sheet path
    try:
        sheet_path = _resolve_worksheet_path(file_map, "CMR")
    except ValueError:
        # Fallback if no sheet named "CMR" exists
        sheet_path = "xl/worksheets/sheet1.xml"

    if sheet_path not in file_map:
        return _write_fallback_cmr_workbook(output, email_data, customer_order)

    sheet_root = ET.fromstring(file_map[sheet_path])

    # Fill delivery date
    set_string_cell(sheet_root, shared_strings_root, _COLRUYT_CMR_DATE_CELL, email_data.delivery_date_to_dc)

    # Fill order reference with "HESSING ORDERNO." prefix
    ref_values = _normalize_cmr_references(references, customer_order.reference)
    first_ref = ref_values[0] if ref_values else (customer_order.reference or "")
    if first_ref:
        upper = first_ref.strip().upper()
        if not upper.startswith("HESSING"):
            formatted_ref = f"{_COLRUYT_CMR_REF_PREFIX}{first_ref.strip()}"
        else:
            formatted_ref = first_ref.strip()
        set_string_cell(sheet_root, shared_strings_root, _COLRUYT_CMR_REF_CELL, formatted_ref)

    # Fill pallet places (E40)
    if pallet_places:
        try:
            set_numeric_cell(sheet_root, _COLRUYT_CMR_PP_CELL, float(str(pallet_places).replace(",", ".")))
        except ValueError:
            set_string_cell(sheet_root, shared_strings_root, _COLRUYT_CMR_PP_CELL, str(pallet_places))

    # Fill Goederen total from pakbon (E41)
    if goederen_total is not None:
        try:
            set_numeric_cell(sheet_root, _COLRUYT_CMR_GOEDEREN_TOTAL_CELL, float(goederen_total))
        except (ValueError, TypeError):
            set_string_cell(sheet_root, shared_strings_root, _COLRUYT_CMR_GOEDEREN_TOTAL_CELL, str(goederen_total))

    # Fill slotboeking ID (e.g., L22293033) - clear first, then fill if available
    clear_cell_value(sheet_root, _COLRUYT_CMR_SLOTBOEKING_CELL)
    if customer_order.slotboeking_id:
        set_string_cell(sheet_root, shared_strings_root, _COLRUYT_CMR_SLOTBOEKING_CELL, customer_order.slotboeking_id)

    sanitize_optional_namespace_attributes(sheet_root)
    file_map[sheet_path] = ET.tostring(sheet_root, encoding="utf-8", xml_declaration=False)
    
    sanitize_optional_namespace_attributes(shared_strings_root)
    shared_strings_root.set("count", str(_count_shared_string_references_in_workbook(file_map)))
    shared_strings_root.set("uniqueCount", str(len(shared_strings_root.findall("x:si", NS))))
    file_map["xl/sharedStrings.xml"] = ET.tostring(shared_strings_root, encoding="utf-8", xml_declaration=False)

    with ZipFile(output, "w", ZIP_DEFLATED) as target_archive:
        for name, content in file_map.items():
            target_archive.writestr(name, content)

    return output


def fill_denemark_cmr_template(
    template_path: str | Path,
    output_path: str | Path,
    email_data: ParsedOrderEmail,
    customer_order: CustomerOrder,
) -> Path:
    """Fill the Denemark CMR template (no order data needed - just copy template)."""
    template = Path(template_path)
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)

    if not template.exists():
        return _write_fallback_cmr_workbook(output, email_data, customer_order)

    # For Denemark, simply copy the template as-is (no data to fill)
    output.write_bytes(template.read_bytes())
    return output


def fill_edeka_laatzen_cmr_template(
    template_path: str | Path,
    output_path: str | Path,
    email_data: ParsedOrderEmail,
    customer_order: CustomerOrder,
) -> Path:
    """Fill the Edeka Laatzen CMR template - put order reference in merged cells A26:AG27."""
    template = Path(template_path)
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)

    if not template.exists():
        return _write_fallback_cmr_workbook(output, email_data, customer_order)

    with ZipFile(template, "r") as source_archive:
        file_map = {name: source_archive.read(name) for name in source_archive.namelist()}

    # Build/load shared strings (create empty root if missing)
    if "xl/sharedStrings.xml" in file_map:
        shared_strings_root = ET.fromstring(file_map["xl/sharedStrings.xml"])
    else:
        shared_strings_root = ET.Element(f"{{{MAIN_NS}}}sst")

    # Dynamically resolve the CMR sheet path
    try:
        sheet_path = _resolve_worksheet_path(file_map, "CMR")
    except ValueError:
        # Fallback if no sheet named "CMR" exists
        sheet_path = "xl/worksheets/sheet1.xml"

    if sheet_path not in file_map:
        return _write_fallback_cmr_workbook(output, email_data, customer_order)

    sheet_root = ET.fromstring(file_map[sheet_path])

    # Fill order reference in merged cells A26:AG27
    # Always clear A26 first (remove template example data)
    clear_cell_value(sheet_root, "A26")
    
    # Only fill with order reference if there's one
    if customer_order.reference:
        set_string_cell(sheet_root, shared_strings_root, "A26", customer_order.reference)

    sanitize_optional_namespace_attributes(sheet_root)
    file_map[sheet_path] = ET.tostring(sheet_root, encoding="utf-8", xml_declaration=False)
    
    sanitize_optional_namespace_attributes(shared_strings_root)
    shared_strings_root.set("count", str(_count_shared_string_references_in_workbook(file_map)))
    shared_strings_root.set("uniqueCount", str(len(shared_strings_root.findall("x:si", NS))))
    file_map["xl/sharedStrings.xml"] = ET.tostring(shared_strings_root, encoding="utf-8", xml_declaration=False)

    with ZipFile(output, "w", ZIP_DEFLATED) as target_archive:
        for name, content in file_map.items():
            target_archive.writestr(name, content)

    return output


def fill_edeka_mochmuhl_cmr_template(
    template_path: str | Path,
    output_path: str | Path,
    email_data: ParsedOrderEmail,
    customer_order: CustomerOrder,
) -> Path:
    """Fill the Edeka Mochmuhl CMR template - put order reference in merged cells A26:AG27."""
    template = Path(template_path)
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)

    if not template.exists():
        return _write_fallback_cmr_workbook(output, email_data, customer_order)

    with ZipFile(template, "r") as source_archive:
        file_map = {name: source_archive.read(name) for name in source_archive.namelist()}

    # Build/load shared strings (create empty root if missing)
    if "xl/sharedStrings.xml" in file_map:
        shared_strings_root = ET.fromstring(file_map["xl/sharedStrings.xml"])
    else:
        shared_strings_root = ET.Element(f"{{{MAIN_NS}}}sst")

    # Dynamically resolve the CMR sheet path
    try:
        sheet_path = _resolve_worksheet_path(file_map, "CMR")
    except ValueError:
        # Fallback if no sheet named "CMR" exists
        sheet_path = "xl/worksheets/sheet1.xml"

    if sheet_path not in file_map:
        return _write_fallback_cmr_workbook(output, email_data, customer_order)

    sheet_root = ET.fromstring(file_map[sheet_path])

    # Fill order reference in merged cells A26:AG27
    # Always clear A26 first (remove template example data)
    clear_cell_value(sheet_root, "A26")
    
    # Only fill with order reference if there's one
    if customer_order.reference:
        set_string_cell(sheet_root, shared_strings_root, "A26", customer_order.reference)

    sanitize_optional_namespace_attributes(sheet_root)
    file_map[sheet_path] = ET.tostring(sheet_root, encoding="utf-8", xml_declaration=False)
    
    sanitize_optional_namespace_attributes(shared_strings_root)
    shared_strings_root.set("count", str(_count_shared_string_references_in_workbook(file_map)))
    shared_strings_root.set("uniqueCount", str(len(shared_strings_root.findall("x:si", NS))))
    file_map["xl/sharedStrings.xml"] = ET.tostring(shared_strings_root, encoding="utf-8", xml_declaration=False)

    with ZipFile(output, "w", ZIP_DEFLATED) as target_archive:
        for name, content in file_map.items():
            target_archive.writestr(name, content)

    return output


def fill_havi_duisburg_cmr_template(
    template_path: str | Path,
    output_path: str | Path,
    email_data: ParsedOrderEmail,
    customer_order: CustomerOrder,
    references: list[str] | None = None,
) -> Path:
    """Fill the Havi Duisburg CMR template - put order reference in F40 with bold and size 12."""
    template = Path(template_path)
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)

    if not template.exists():
        return _write_fallback_cmr_workbook(output, email_data, customer_order)

    with ZipFile(template, "r") as source_archive:
        file_map = {name: source_archive.read(name) for name in source_archive.namelist()}

    # Build/load shared strings (create empty root if missing)
    if "xl/sharedStrings.xml" in file_map:
        shared_strings_root = ET.fromstring(file_map["xl/sharedStrings.xml"])
    else:
        shared_strings_root = ET.Element(f"{{{MAIN_NS}}}sst")

    # Dynamically resolve the CMR sheet path
    try:
        sheet_path = _resolve_worksheet_path(file_map, "CMR")
    except ValueError:
        # Fallback if no sheet named "CMR" exists
        sheet_path = "xl/worksheets/sheet1.xml"

    if sheet_path not in file_map:
        return _write_fallback_cmr_workbook(output, email_data, customer_order)

    sheet_root = ET.fromstring(file_map[sheet_path])

    # Fill order reference in F40 with bold and font size 12
    clear_cell_value(sheet_root, "F40")
    
    # Determine what reference to use
    reference_to_use = ""
    if references:
        # Use provided references (from Leverschema selection), join with "/"
        cleaned_refs = [ref.strip() for ref in references if ref and ref.strip()]
        if cleaned_refs:
            reference_to_use = " / ".join(cleaned_refs)
    elif customer_order.reference:
        # Fallback to single order reference
        reference_to_use = customer_order.reference
    
    # Fill with reference if we have one (bold, size 12)
    if reference_to_use:
        set_formatted_string_cell(sheet_root, shared_strings_root, "F40", reference_to_use, bold=True, font_size=12)

    sanitize_optional_namespace_attributes(sheet_root)
    file_map[sheet_path] = ET.tostring(sheet_root, encoding="utf-8", xml_declaration=False)
    
    sanitize_optional_namespace_attributes(shared_strings_root)
    shared_strings_root.set("count", str(_count_shared_string_references_in_workbook(file_map)))
    shared_strings_root.set("uniqueCount", str(len(shared_strings_root.findall("x:si", NS))))
    file_map["xl/sharedStrings.xml"] = ET.tostring(shared_strings_root, encoding="utf-8", xml_declaration=False)

    with ZipFile(output, "w", ZIP_DEFLATED) as target_archive:
        for name, content in file_map.items():
            target_archive.writestr(name, content)

    return output


def fill_havi_wunstorf_cmr_template(
    template_path: str | Path,
    output_path: str | Path,
    email_data: ParsedOrderEmail,
    customer_order: CustomerOrder,
    references: list[str] | None = None,
) -> Path:
    """Fill the Havi Wunstorf CMR template - put order reference in F40 with bold and size 12."""
    template = Path(template_path)
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)

    if not template.exists():
        return _write_fallback_cmr_workbook(output, email_data, customer_order)

    with ZipFile(template, "r") as source_archive:
        file_map = {name: source_archive.read(name) for name in source_archive.namelist()}

    # Build/load shared strings (create empty root if missing)
    if "xl/sharedStrings.xml" in file_map:
        shared_strings_root = ET.fromstring(file_map["xl/sharedStrings.xml"])
    else:
        shared_strings_root = ET.Element(f"{{{MAIN_NS}}}sst")

    # Dynamically resolve the CMR sheet path
    try:
        sheet_path = _resolve_worksheet_path(file_map, "CMR")
    except ValueError:
        # Fallback if no sheet named "CMR" exists
        sheet_path = "xl/worksheets/sheet1.xml"

    if sheet_path not in file_map:
        return _write_fallback_cmr_workbook(output, email_data, customer_order)

    sheet_root = ET.fromstring(file_map[sheet_path])

    # Fill order reference in F40 with bold and font size 12
    clear_cell_value(sheet_root, "F40")
    
    # Determine what reference to use
    reference_to_use = ""
    if references:
        # Use provided references (from Leverschema selection), join with "/"
        cleaned_refs = [ref.strip() for ref in references if ref and ref.strip()]
        if cleaned_refs:
            reference_to_use = " / ".join(cleaned_refs)
    elif customer_order.reference:
        # Fallback to single order reference
        reference_to_use = customer_order.reference
    
    # Fill with reference if we have one (bold, size 12)
    if reference_to_use:
        set_formatted_string_cell(sheet_root, shared_strings_root, "F40", reference_to_use, bold=True, font_size=12)

    sanitize_optional_namespace_attributes(sheet_root)
    file_map[sheet_path] = ET.tostring(sheet_root, encoding="utf-8", xml_declaration=False)
    
    sanitize_optional_namespace_attributes(shared_strings_root)
    shared_strings_root.set("count", str(_count_shared_string_references_in_workbook(file_map)))
    shared_strings_root.set("uniqueCount", str(len(shared_strings_root.findall("x:si", NS))))
    file_map["xl/sharedStrings.xml"] = ET.tostring(shared_strings_root, encoding="utf-8", xml_declaration=False)

    with ZipFile(output, "w", ZIP_DEFLATED) as target_archive:
        for name, content in file_map.items():
            target_archive.writestr(name, content)

    return output


def fill_havi_neu_wulmstorf_cmr_template(
    template_path: str | Path,
    output_path: str | Path,
    email_data: ParsedOrderEmail,
    customer_order: CustomerOrder,
    references: list[str] | None = None,
) -> Path:
    """Fill the Havi Neu Wulmstorf CMR template - put order reference in F40 with bold and size 12."""
    template = Path(template_path)
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)

    if not template.exists():
        return _write_fallback_cmr_workbook(output, email_data, customer_order)

    with ZipFile(template, "r") as source_archive:
        file_map = {name: source_archive.read(name) for name in source_archive.namelist()}

    # Build/load shared strings (create empty root if missing)
    if "xl/sharedStrings.xml" in file_map:
        shared_strings_root = ET.fromstring(file_map["xl/sharedStrings.xml"])
    else:
        shared_strings_root = ET.Element(f"{{{MAIN_NS}}}sst")

    # Dynamically resolve the CMR sheet path
    try:
        sheet_path = _resolve_worksheet_path(file_map, "CMR")
    except ValueError:
        # Fallback if no sheet named "CMR" exists
        sheet_path = "xl/worksheets/sheet1.xml"

    if sheet_path not in file_map:
        return _write_fallback_cmr_workbook(output, email_data, customer_order)

    sheet_root = ET.fromstring(file_map[sheet_path])

    # Fill order reference in F40 with bold and font size 12
    clear_cell_value(sheet_root, "F40")
    
    # Determine what reference to use
    reference_to_use = ""
    if references:
        # Use provided references (from Leverschema selection), join with "/"
        cleaned_refs = [ref.strip() for ref in references if ref and ref.strip()]
        if cleaned_refs:
            reference_to_use = " / ".join(cleaned_refs)
    elif customer_order.reference:
        # Fallback to single order reference
        reference_to_use = customer_order.reference
    
    # Fill with reference if we have one (bold, size 12)
    if reference_to_use:
        set_formatted_string_cell(sheet_root, shared_strings_root, "F40", reference_to_use, bold=True, font_size=12)

    sanitize_optional_namespace_attributes(sheet_root)
    file_map[sheet_path] = ET.tostring(sheet_root, encoding="utf-8", xml_declaration=False)
    
    sanitize_optional_namespace_attributes(shared_strings_root)
    shared_strings_root.set("count", str(_count_shared_string_references_in_workbook(file_map)))
    shared_strings_root.set("uniqueCount", str(len(shared_strings_root.findall("x:si", NS))))
    file_map["xl/sharedStrings.xml"] = ET.tostring(shared_strings_root, encoding="utf-8", xml_declaration=False)

    with ZipFile(output, "w", ZIP_DEFLATED) as target_archive:
        for name, content in file_map.items():
            target_archive.writestr(name, content)

    return output


def fill_havi_de_cmr_template(
    template_path: str | Path,
    output_path: str | Path,
    email_data: ParsedOrderEmail,
    customer_order: CustomerOrder,
) -> Path:
    """Fill the Havi DE CMR template (no order data needed - just copy template like Denemark)."""
    template = Path(template_path)
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)

    if not template.exists():
        return _write_fallback_cmr_workbook(output, email_data, customer_order)

    # For Havi DE, simply copy the template as-is (no data to fill)
    output.write_bytes(template.read_bytes())
    return output


def fill_havi_be_cmr_template(
    template_path: str | Path,
    output_path: str | Path,
    email_data: ParsedOrderEmail,
    customer_order: CustomerOrder,
    references: list[str] | None = None,
) -> Path:
    """Fill the Havi BE CMR template - put order reference in cells C29 and C30.
    If multiple references provided, join them with '/'."""
    template = Path(template_path)
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)

    if not template.exists():
        return _write_fallback_cmr_workbook(output, email_data, customer_order)

    with ZipFile(template, "r") as source_archive:
        file_map = {name: source_archive.read(name) for name in source_archive.namelist()}

    # Build/load shared strings (create empty root if missing)
    if "xl/sharedStrings.xml" in file_map:
        shared_strings_root = ET.fromstring(file_map["xl/sharedStrings.xml"])
    else:
        shared_strings_root = ET.Element(f"{{{MAIN_NS}}}sst")

    # Dynamically resolve the CMR sheet path
    try:
        sheet_path = _resolve_worksheet_path(file_map, "CMR")
    except ValueError:
        # Fallback if no sheet named "CMR" exists
        sheet_path = "xl/worksheets/sheet1.xml"

    if sheet_path not in file_map:
        return _write_fallback_cmr_workbook(output, email_data, customer_order)

    sheet_root = ET.fromstring(file_map[sheet_path])

    # Fill order reference in C29 and C30 (lines 29 and 30, column C)
    # Always clear C29 and C30 first (remove template example data)
    clear_cell_value(sheet_root, "C29")
    clear_cell_value(sheet_root, "C30")
    
    # Determine what reference to use
    reference_to_use = ""
    if references:
        # Use provided references (from Leverschema selection), join with "/"
        cleaned_refs = [ref.strip() for ref in references if ref and ref.strip()]
        if cleaned_refs:
            reference_to_use = " / ".join(cleaned_refs)
    elif customer_order.reference:
        # Fallback to single order reference
        reference_to_use = customer_order.reference
    
    # Fill with reference if we have one
    # Put the reference in C29, and if it's very long, continue in C30
    if reference_to_use:
        # If reference is short, put it all in C29
        if len(reference_to_use) <= 50:
            set_string_cell(sheet_root, shared_strings_root, "C29", reference_to_use)
        else:
            # If reference is long, split between C29 and C30
            mid_point = len(reference_to_use) // 2
            # Find a good split point (after a space or "/")
            split_point = mid_point
            for i in range(mid_point - 10, mid_point + 10):
                if i < len(reference_to_use) and reference_to_use[i] in [' ', '/']:
                    split_point = i + 1
                    break
            
            part1 = reference_to_use[:split_point].strip()
            part2 = reference_to_use[split_point:].strip()
            
            set_string_cell(sheet_root, shared_strings_root, "C29", part1)
            if part2:
                set_string_cell(sheet_root, shared_strings_root, "C30", part2)

    sanitize_optional_namespace_attributes(sheet_root)
    file_map[sheet_path] = ET.tostring(sheet_root, encoding="utf-8", xml_declaration=False)
    
    sanitize_optional_namespace_attributes(shared_strings_root)
    shared_strings_root.set("count", str(_count_shared_string_references_in_workbook(file_map)))
    shared_strings_root.set("uniqueCount", str(len(shared_strings_root.findall("x:si", NS))))
    file_map["xl/sharedStrings.xml"] = ET.tostring(shared_strings_root, encoding="utf-8", xml_declaration=False)

    with ZipFile(output, "w", ZIP_DEFLATED) as target_archive:
        for name, content in file_map.items():
            target_archive.writestr(name, content)

    return output


def fill_havi_nl_cmr_template(
    template_path: str | Path,
    output_path: str | Path,
    email_data: ParsedOrderEmail,
    customer_order: CustomerOrder,
    references: list[str] | None = None,
) -> Path:
    """Fill the Havi NL CMR template - put order reference in cell F38.
    If multiple references provided, join them with '/'."""
    template = Path(template_path)
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)

    if not template.exists():
        return _write_fallback_cmr_workbook(output, email_data, customer_order)

    with ZipFile(template, "r") as source_archive:
        file_map = {name: source_archive.read(name) for name in source_archive.namelist()}

    # Build/load shared strings (create empty root if missing)
    if "xl/sharedStrings.xml" in file_map:
        shared_strings_root = ET.fromstring(file_map["xl/sharedStrings.xml"])
    else:
        shared_strings_root = ET.Element(f"{{{MAIN_NS}}}sst")

    # Dynamically resolve the CMR sheet path
    try:
        sheet_path = _resolve_worksheet_path(file_map, "CMR")
    except ValueError:
        # Fallback if no sheet named "CMR" exists
        sheet_path = "xl/worksheets/sheet1.xml"

    if sheet_path not in file_map:
        return _write_fallback_cmr_workbook(output, email_data, customer_order)

    sheet_root = ET.fromstring(file_map[sheet_path])

    # Fill order reference in F38 (E38 has label "Order:")
    # Always clear F38 first (remove template example data)
    clear_cell_value(sheet_root, "F38")
    
    # Determine what reference to use
    reference_to_use = ""
    if references:
        # Use provided references (from Leverschema selection), join with "/"
        cleaned_refs = [ref.strip() for ref in references if ref and ref.strip()]
        if cleaned_refs:
            reference_to_use = " / ".join(cleaned_refs)
    elif customer_order.reference:
        # Fallback to single order reference
        reference_to_use = customer_order.reference
    
    # Fill with reference if we have one
    if reference_to_use:
        set_string_cell(sheet_root, shared_strings_root, "F38", reference_to_use)

    sanitize_optional_namespace_attributes(sheet_root)
    file_map[sheet_path] = ET.tostring(sheet_root, encoding="utf-8", xml_declaration=False)
    
    sanitize_optional_namespace_attributes(shared_strings_root)
    shared_strings_root.set("count", str(_count_shared_string_references_in_workbook(file_map)))
    shared_strings_root.set("uniqueCount", str(len(shared_strings_root.findall("x:si", NS))))
    file_map["xl/sharedStrings.xml"] = ET.tostring(shared_strings_root, encoding="utf-8", xml_declaration=False)

    with ZipFile(output, "w", ZIP_DEFLATED) as target_archive:
        for name, content in file_map.items():
            target_archive.writestr(name, content)

    return output

    with ZipFile(output, "w", ZIP_DEFLATED) as target_archive:
        for name, content in file_map.items():
            target_archive.writestr(name, content)

    return output


def fill_heeren_cmr_template(
    template_path: str | Path,
    output_path: str | Path,
    email_data: ParsedOrderEmail,
    customer_order: CustomerOrder,
    references: list[str] | None = None,
) -> Path:
    """Fill the Heeren CMR template - put order reference in appropriate cell."""
    template = Path(template_path)
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)

    if not template.exists():
        return _write_fallback_cmr_workbook(output, email_data, customer_order)

    with ZipFile(template, "r") as source_archive:
        file_map = {name: source_archive.read(name) for name in source_archive.namelist()}

    # Build/load shared strings (create empty root if missing)
    if "xl/sharedStrings.xml" in file_map:
        shared_strings_root = ET.fromstring(file_map["xl/sharedStrings.xml"])
    else:
        shared_strings_root = ET.Element(f"{{{MAIN_NS}}}sst")

    # Dynamically resolve the CMR sheet path
    try:
        sheet_path = _resolve_worksheet_path(file_map, "CMR")
    except ValueError:
        # Fallback if no sheet named "CMR" exists
        sheet_path = "xl/worksheets/sheet1.xml"

    if sheet_path not in file_map:
        return _write_fallback_cmr_workbook(output, email_data, customer_order)

    sheet_root = ET.fromstring(file_map[sheet_path])

    # Fill order reference (using same cell as Havi NL: F38)
    clear_cell_value(sheet_root, "F38")
    
    # Determine what reference to use
    reference_to_use = ""
    if references:
        # Use provided references (from Leverschema selection), join with "/"
        cleaned_refs = [ref.strip() for ref in references if ref and ref.strip()]
        if cleaned_refs:
            reference_to_use = " / ".join(cleaned_refs)
    elif customer_order.reference:
        # Fallback to single order reference
        reference_to_use = customer_order.reference
    
    # Fill with reference if we have one
    if reference_to_use:
        set_string_cell(sheet_root, shared_strings_root, "F38", reference_to_use)

    sanitize_optional_namespace_attributes(sheet_root)
    file_map[sheet_path] = ET.tostring(sheet_root, encoding="utf-8", xml_declaration=False)
    
    sanitize_optional_namespace_attributes(shared_strings_root)
    shared_strings_root.set("count", str(_count_shared_string_references_in_workbook(file_map)))
    shared_strings_root.set("uniqueCount", str(len(shared_strings_root.findall("x:si", NS))))
    file_map["xl/sharedStrings.xml"] = ET.tostring(shared_strings_root, encoding="utf-8", xml_declaration=False)

    with ZipFile(output, "w", ZIP_DEFLATED) as target_archive:
        for name, content in file_map.items():
            target_archive.writestr(name, content)

    return output


def fill_nettomd_cmr_template(
    template_path: str | Path,
    output_path: str | Path,
    email_data: ParsedOrderEmail,
    customer_order: CustomerOrder,
) -> Path:
    """Fill the NettoMD CMR template - just copy template (no data to fill)."""
    template = Path(template_path)
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)

    if not template.exists():
        return _write_fallback_cmr_workbook(output, email_data, customer_order)

    # For NettoMD, simply copy the template as-is (no data to fill)
    output.write_bytes(template.read_bytes())
    return output


def fill_rewe_penny_cmr_template(
    template_path: str | Path,
    output_path: str | Path,
    email_data: ParsedOrderEmail,
    customer_order: CustomerOrder,
) -> Path:
    """Fill the Rewe/Penny CMR template - just copy template (no data to fill)."""
    template = Path(template_path)
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)

    if not template.exists():
        return _write_fallback_cmr_workbook(output, email_data, customer_order)

    # For Rewe/Penny, simply copy the template as-is (no data to fill)
    output.write_bytes(template.read_bytes())
    return output


def fill_hanos_cmr_template(
    template_path: str | Path,
    output_path: str | Path,
    email_data: ParsedOrderEmail,
    customer_order: CustomerOrder,
) -> Path:
    """Fill the Hanos CMR template - put order reference (VSCO...) in E36."""
    template = Path(template_path)
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)

    if not template.exists():
        return _write_fallback_cmr_workbook(output, email_data, customer_order)

    with ZipFile(template, "r") as source_archive:
        file_map = {name: source_archive.read(name) for name in source_archive.namelist()}

    # Build/load shared strings (create empty root if missing)
    if "xl/sharedStrings.xml" in file_map:
        shared_strings_root = ET.fromstring(file_map["xl/sharedStrings.xml"])
    else:
        shared_strings_root = ET.Element(f"{{{MAIN_NS}}}sst")

    # Dynamically resolve the CMR sheet path
    try:
        sheet_path = _resolve_worksheet_path(file_map, "CMR")
    except ValueError:
        # Fallback if no sheet named "CMR" exists
        sheet_path = "xl/worksheets/sheet1.xml"

    if sheet_path not in file_map:
        return _write_fallback_cmr_workbook(output, email_data, customer_order)

    sheet_root = ET.fromstring(file_map[sheet_path])

    # Fill order reference in E36 (only if there's a reference)
    clear_cell_value(sheet_root, "E36")
    if customer_order.reference:
        set_string_cell(sheet_root, shared_strings_root, "E36", customer_order.reference)

    sanitize_optional_namespace_attributes(sheet_root)
    file_map[sheet_path] = ET.tostring(sheet_root, encoding="utf-8", xml_declaration=False)
    
    sanitize_optional_namespace_attributes(shared_strings_root)
    shared_strings_root.set("count", str(_count_shared_string_references_in_workbook(file_map)))
    shared_strings_root.set("uniqueCount", str(len(shared_strings_root.findall("x:si", NS))))
    file_map["xl/sharedStrings.xml"] = ET.tostring(shared_strings_root, encoding="utf-8", xml_declaration=False)

    with ZipFile(output, "w", ZIP_DEFLATED) as target_archive:
        for name, content in file_map.items():
            target_archive.writestr(name, content)

    return output


def _write_fallback_cmr_workbook(output: Path, email_data: ParsedOrderEmail, customer_order: CustomerOrder) -> Path:
    rows: list[list[WorkbookCell]] = [
        [WorkbookCell("Delivery Date", "header"), WorkbookCell(email_data.delivery_date_to_dc)],
        [WorkbookCell("Customer", "header"), WorkbookCell(customer_order.customer_dc)],
        [WorkbookCell("Reference", "header"), WorkbookCell(customer_order.reference)],
        [],
        [WorkbookCell("Quantity", "header"), WorkbookCell("Description", "header")],
    ]

    for item in customer_order.items:
        rows.append([
            WorkbookCell(item.quantity_boxes, "number"),
            WorkbookCell(item.description),
        ])

    return write_simple_workbook(output, "CMR", rows, title="CMR Document")


def _normalize_cmr_references(references: list[str] | None, fallback_reference: str) -> list[str]:
    values = references or []
    if not values and fallback_reference:
        values = [part.strip() for part in str(fallback_reference).split("+")]

    seen: set[str] = set()
    ordered: list[str] = []
    for value in values:
        cleaned = str(value or "").strip()
        if not cleaned:
            continue
        if cleaned in seen:
            continue
        seen.add(cleaned)
        ordered.append(cleaned)
    return ordered


def _shift_merged_ranges(sheet_root: ET.Element, offset: int, threshold: int = 22) -> None:
    merge_cells = sheet_root.find("x:mergeCells", NS)
    if merge_cells is None:
        return

    for merge_cell in merge_cells.findall("x:mergeCell", NS):
        ref = merge_cell.get("ref", "")
        if not ref or ":" not in ref:
            continue
        start, end = ref.split(":")
        start_col, start_row = _split_cell_ref(start)
        end_col, end_row = _split_cell_ref(end)
        if start_row >= threshold:
            start_row += offset
            end_row += offset
            merge_cell.set("ref", f"{start_col}{start_row}:{end_col}{end_row}")


def _update_dimension(sheet_root: ET.Element, offset: int) -> None:
    dimension = sheet_root.find("x:dimension", NS)
    if dimension is None:
        return
    ref = dimension.get("ref", "")
    if ":" not in ref:
        return
    start, end = ref.split(":")
    end_col, end_row = _split_cell_ref(end)
    dimension.set("ref", f"{start}:{end_col}{end_row + offset}")


def _split_cell_ref(cell_ref: str) -> tuple[str, int]:
    match = re.fullmatch(r"([A-Z]+)(\d+)", cell_ref)
    if not match:
        raise ValueError(f"Invalid cell reference: {cell_ref}")
    return match.group(1), int(match.group(2))


def sanitize_optional_namespace_attributes(root: ET.Element) -> None:
    for element in root.iter():
        to_delete = []
        for attribute_name in element.attrib:
            if attribute_name.startswith("{"):
                namespace = attribute_name[1:].split("}", 1)[0]
                if namespace in OPTIONAL_ATTR_NAMESPACES:
                    to_delete.append(attribute_name)
        for attribute_name in to_delete:
            del element.attrib[attribute_name]


def set_string_cell(sheet_root: ET.Element, shared_strings_root: ET.Element, cell_ref: str, value: str) -> None:
    cell = _find_cell(sheet_root, cell_ref)
    if cell is None:
        return

    cell.set("t", "s")
    value_node = cell.find("x:v", NS)
    if value_node is None:
        value_node = ET.SubElement(cell, f"{{{MAIN_NS}}}v")
    value_node.text = str(_get_shared_string_index(shared_strings_root, value))


def set_formatted_string_cell(sheet_root: ET.Element, shared_strings_root: ET.Element, cell_ref: str, value: str, bold: bool = False, font_size: int = None) -> None:
    """Set a string cell with specific formatting (bold, font size)."""
    cell = _find_cell(sheet_root, cell_ref)
    if cell is None:
        return

    cell.set("t", "s")
    
    # Create or update the shared string with formatting
    string_index = _get_formatted_shared_string_index(shared_strings_root, value, bold, font_size)
    
    value_node = cell.find("x:v", NS)
    if value_node is None:
        value_node = ET.SubElement(cell, f"{{{MAIN_NS}}}v")
    value_node.text = str(string_index)


def _get_formatted_shared_string_index(shared_strings_root: ET.Element, value: str, bold: bool = False, font_size: int = None) -> int:
    """Get or create a shared string index with specific formatting."""
    # Check if this formatted string already exists
    for index, si in enumerate(shared_strings_root.findall("x:si", NS)):
        t_elem = si.find("x:t", NS)
        if t_elem is not None and t_elem.text == value:
            # Check if formatting matches
            r_elem = si.find("x:r", NS)
            if r_elem is not None:
                rpr_elem = r_elem.find("x:rPr", NS)
                if rpr_elem is not None:
                    has_bold = rpr_elem.find("x:b", NS) is not None
                    sz_elem = rpr_elem.find("x:sz", NS)
                    has_size = sz_elem is not None and sz_elem.get("val") == str(font_size) if font_size else sz_elem is None
                    if has_bold == bold and has_size:
                        return index
            elif not bold and not font_size:
                # Plain text match
                return index
    
    # Create new formatted shared string
    si = ET.SubElement(shared_strings_root, f"{{{MAIN_NS}}}si")
    
    if bold or font_size:
        # Create rich text with formatting
        r = ET.SubElement(si, f"{{{MAIN_NS}}}r")
        rpr = ET.SubElement(r, f"{{{MAIN_NS}}}rPr")
        
        if bold:
            ET.SubElement(rpr, f"{{{MAIN_NS}}}b")
        
        if font_size:
            sz = ET.SubElement(rpr, f"{{{MAIN_NS}}}sz")
            sz.set("val", str(font_size))
        
        t = ET.SubElement(r, f"{{{MAIN_NS}}}t")
        t.text = value
    else:
        # Plain text
        t = ET.SubElement(si, f"{{{MAIN_NS}}}t")
        t.text = value
    
    return len(shared_strings_root.findall("x:si", NS)) - 1


def set_numeric_cell(sheet_root: ET.Element, cell_ref: str, value: float) -> None:
    cell = _find_cell(sheet_root, cell_ref)
    if cell is None:
        return

    cell.attrib.pop("t", None)
    value_node = cell.find("x:v", NS)
    if value_node is None:
        value_node = ET.SubElement(cell, f"{{{MAIN_NS}}}v")
    value_node.text = _format_number(value)


def clear_cell_value(sheet_root: ET.Element, cell_ref: str) -> None:
    cell = _find_cell(sheet_root, cell_ref)
    if cell is None:
        return

    cell.attrib.pop("t", None)
    value_node = cell.find("x:v", NS)
    if value_node is not None:
        cell.remove(value_node)


def _find_cell(sheet_root: ET.Element, cell_ref: str) -> ET.Element | None:
    return sheet_root.find(f".//x:c[@r='{cell_ref}']", NS)


def _get_shared_string_index(shared_strings_root: ET.Element, value: str) -> int:
    items = shared_strings_root.findall("x:si", NS)
    for index, item in enumerate(items):
        text_node = item.find("x:t", NS)
        if text_node is not None and (text_node.text or "") == value:
            return index

    si = ET.SubElement(shared_strings_root, f"{{{MAIN_NS}}}si")
    text_node = ET.SubElement(si, f"{{{MAIN_NS}}}t")
    if value.startswith(" ") or value.endswith(" "):
        text_node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    text_node.text = value
    return len(items)


def _count_shared_string_references(sheet_root: ET.Element) -> int:
    count = 0
    for cell in sheet_root.findall(".//x:c", NS):
        if cell.get("t") == "s":
            count += 1
    return count


def _count_shared_string_references_in_workbook(file_map: dict[str, bytes]) -> int:
    total = 0
    for name, content in file_map.items():
        if not name.startswith("xl/worksheets/") or not name.endswith(".xml"):
            continue
        try:
            sheet_root = ET.fromstring(content)
        except ET.ParseError:
            continue
        total += _count_shared_string_references(sheet_root)
    return total


def _format_number(value: float) -> str:
    if value.is_integer():
        return str(int(value))
    return f"{value:.2f}".rstrip("0").rstrip(".")


def fill_edeka_laatzen_cmr_template(
    template_path: str | Path,
    output_path: str | Path,
    email_data: ParsedOrderEmail,
    customer_order: CustomerOrder,
) -> Path:
    """Fill the Edeka Laatzen CMR template (A26/A27=reference)."""
    template = Path(template_path)
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)

    if not template.exists():
        return _write_fallback_cmr_workbook(output, email_data, customer_order)

    with ZipFile(template, "r") as source_archive:
        file_map = {name: source_archive.read(name) for name in source_archive.namelist()}

    # build/load shared strings
    if "xl/sharedStrings.xml" in file_map:
        shared_strings_root = ET.fromstring(file_map["xl/sharedStrings.xml"])
    else:
        shared_strings_root = ET.Element(f"{{{MAIN_NS}}}sst")

    # Resolve worksheet
    try:
        sheet_path = _resolve_worksheet_path(file_map, "CMR")
    except ValueError:
        sheet_path = "xl/worksheets/sheet1.xml"

    if sheet_path not in file_map:
        return _write_fallback_cmr_workbook(output, email_data, customer_order)

    sheet_root = ET.fromstring(file_map[sheet_path])

    # Fill reference in A26 (merged with A27)
    reference = customer_order.reference or ""
    set_string_cell(sheet_root, shared_strings_root, "A26", reference)

    # Sanitize and serialize
    sanitize_optional_namespace_attributes(sheet_root)
    file_map[sheet_path] = ET.tostring(sheet_root, encoding="utf-8", xml_declaration=False)

    sanitize_optional_namespace_attributes(shared_strings_root)
    shared_strings_root.set("count", str(_count_shared_string_references_in_workbook(file_map)))
    shared_strings_root.set("uniqueCount", str(len(shared_strings_root.findall("x:si", NS))))
    file_map["xl/sharedStrings.xml"] = ET.tostring(shared_strings_root, encoding="utf-8", xml_declaration=False)

    with ZipFile(output, "w", ZIP_DEFLATED) as target_archive:
        for name, content in file_map.items():
            target_archive.writestr(name, content)

    return output
