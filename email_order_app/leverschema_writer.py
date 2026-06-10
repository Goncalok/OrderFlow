from __future__ import annotations

import copy
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from .simple_excel import WorkbookCell, write_simple_workbook

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS = {"x": MAIN_NS}
APP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
APP_VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
ET.register_namespace("", MAIN_NS)
ET.register_namespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")

MASTER_SHEETS = {
    "Monday-Thursday": {
        "sheet_path": "xl/worksheets/sheet1.xml",
        "max_row": 9,
        "rows": [
            {"key": "havi_duisburg", "cell": "B3", "date_cell": "C3", "date_offset": 1},
            {"key": "carrefour_fif", "cell": "B4", "date_cell": "C4", "date_offset": 0},
            {"key": "colruyt", "cell": "B6", "date_cell": "C6", "date_offset": 2},
            {"key": "heeren", "cell": "B7", "date_cell": "C7", "date_offset": 2},
            {"key": "carrefour_kdc", "cell": "B8", "date_cell": "C8", "date_offset": 1},
        ],
    },
    "Friday": {
        "sheet_path": "xl/worksheets/sheet2.xml",
        "max_row": 11,
        "rows": [
            {"key": "havi_duisburg", "cell": "B3", "date_cell": "C3", "date_offset": 2},
            {"key": "havi_duisburg_saturday", "cell": "B4", "date_cell": "C4", "date_offset": 3},
            {"key": "carrefour_fif", "cell": "B5", "date_cell": "C5", "date_offset": 0},
            {"key": "colruyt", "cell": "B7", "date_cell": "C7", "date_offset": 3},
            {"key": "colruyt_saturday", "cell": "B8", "date_cell": "C8", "date_offset": 4},
            {"key": "heeren", "cell": "B9", "date_cell": "C9", "date_offset": 3},
            {"key": "carrefour_kdc", "cell": "B10", "date_cell": "C10", "date_offset": 1},
        ],
    },
    "Saturday": {
        "sheet_path": "xl/worksheets/sheet3.xml",
        "max_row": 9,
        "rows": [
            {"key": "havi_duisburg", "cell": "B3", "date_cell": "C3", "date_offset": 2},
            {"key": "carrefour_fif", "cell": "B4", "date_cell": "C4", "date_offset": 0},
            {"key": "colruyt", "cell": "B6", "date_cell": "C6", "date_offset": 3},
            {"key": "heeren", "cell": "B7", "date_cell": "C7", "date_offset": 3},
            {"key": "carrefour_kdc", "cell": "B8", "date_cell": "C8", "date_offset": 2},
        ],
    },
}


def export_leverschema_workbook(
    template_path: str | Path,
    output_path: str | Path,
    leverschema_results: dict,
    selected_sheet: str,
    session_date: str | None = None,
) -> Path:
    template = Path(template_path)
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)

    if selected_sheet not in MASTER_SHEETS:
        raise ValueError(f"Unsupported Leverschema sheet: {selected_sheet}")
    if not template.exists():
        return _export_leverschema_fallback(output, leverschema_results, selected_sheet, session_date)

    with ZipFile(template, "r") as source_archive:
        file_map = {name: source_archive.read(name) for name in source_archive.namelist()}

    grouped = _group_results_by_sheet_and_key(leverschema_results)
    config = MASTER_SHEETS[selected_sheet]
    sheet_root = ET.fromstring(file_map[config["sheet_path"]])
    grouped_entries = grouped.get(selected_sheet, {})

    for row_config in config["rows"]:
        customer_key = row_config["key"]
        cell_ref = row_config["cell"]
        queue = grouped_entries.get(customer_key, [])
        value = None
        target_date = _format_workbook_date(row_config.get("date_offset", 0), session_date)
        if queue:
            match_index = next(
                (index for index, item in enumerate(queue) if _normalize_workbook_date(item.get("deliveryDate")) == _normalize_workbook_date(target_date)),
                -1,
            )
            selected_entry = queue.pop(match_index) if match_index >= 0 else queue.pop(0)
            value = selected_entry.get("totalPalletPlaces")
        set_text_cell(sheet_root, row_config["date_cell"], target_date)
        time_cell = f"D{_extract_row_number(row_config['date_cell'])}"
        set_text_cell(sheet_root, time_cell, _fallback_time(selected_sheet, customer_key))
        if value in (None, "", 0):
            clear_cell_value(sheet_root, cell_ref)
        else:
            set_numeric_cell(sheet_root, cell_ref, float(value))

    trim_export_sheet(sheet_root, config["max_row"])
    file_map[config["sheet_path"]] = ET.tostring(sheet_root, encoding="utf-8", xml_declaration=True)
    file_map["xl/workbook.xml"] = ET.tostring(_build_single_sheet_workbook(file_map["xl/workbook.xml"], selected_sheet), encoding="utf-8", xml_declaration=True)
    file_map["docProps/app.xml"] = ET.tostring(_build_single_sheet_app_properties(selected_sheet), encoding="utf-8", xml_declaration=True)

    with ZipFile(output, "w", ZIP_DEFLATED) as target_archive:
        for name, content in file_map.items():
            target_archive.writestr(name, content)

    return output


def _export_leverschema_fallback(
    output: Path,
    leverschema_results: dict,
    selected_sheet: str,
    session_date: str | None,
) -> Path:
    grouped = _group_results_by_sheet_and_key(leverschema_results)
    config = MASTER_SHEETS[selected_sheet]
    grouped_entries = grouped.get(selected_sheet, {})

    rows: list[list[WorkbookCell]] = [
        [WorkbookCell("Klant", "header"), WorkbookCell("Palletplaatsen (PP)", "header"), WorkbookCell("Losdatum", "header"), WorkbookCell("Levertijd", "header"), WorkbookCell("Vervoerder", "header"), WorkbookCell("Opmerking", "header")]
    ]

    for row_config in config["rows"]:
        customer_key = row_config["key"]
        queue = grouped_entries.get(customer_key, [])
        target_date = _format_workbook_date(row_config.get("date_offset", 0), session_date)
        value = ""
        if queue:
            match_index = next(
                (index for index, item in enumerate(queue) if _normalize_workbook_date(item.get("deliveryDate")) == _normalize_workbook_date(target_date)),
                -1,
            )
            selected_entry = queue.pop(match_index) if match_index >= 0 else queue.pop(0)
            value = selected_entry.get("totalPalletPlaces") or ""

        rows.append(
            [
                WorkbookCell(_fallback_klant_name(customer_key)),
                WorkbookCell(value, "number" if value != "" else "body"),
                WorkbookCell(target_date),
                WorkbookCell(_fallback_time(selected_sheet, customer_key)),
                WorkbookCell(_fallback_carrier(selected_sheet, customer_key)),
                WorkbookCell("prognoose"),
            ]
        )

    return write_simple_workbook(output, selected_sheet, rows, title=f"Leverschema {selected_sheet}")


def _fallback_klant_name(customer_key: str) -> str:
    mapping = {
        "havi_duisburg": "Havi Duisburg",
        "havi_duisburg_saturday": "Havi Duisburg",
        "carrefour_fif": "Carrefour FIF",
        "colruyt": "Colruyt",
        "colruyt_saturday": "Colruyt",
        "heeren": "Heeren",
        "carrefour_kdc": "Carrefour KDC",
    }
    return mapping.get(customer_key, customer_key)


def _fallback_time(sheet_name: str, customer_key: str) -> str:
    for row in MASTER_SHEETS[sheet_name]["rows"]:
        if row["key"] == customer_key:
            if customer_key in {"havi_duisburg", "havi_duisburg_saturday"}:
                return "15:30"
            if customer_key == "carrefour_fif":
                return "11:00"
            if customer_key in {"colruyt", "colruyt_saturday"}:
                return "01:30"
            if customer_key == "heeren":
                return "05:15"
            if customer_key == "carrefour_kdc":
                return "03:15"
    return ""


def _fallback_carrier(sheet_name: str, customer_key: str) -> str:
    if customer_key in {"havi_duisburg", "havi_duisburg_saturday"}:
        return "Hendrikx/SVZ"
    if customer_key in {"carrefour_fif", "carrefour_kdc"}:
        return "Van Tilburg"
    if customer_key in {"colruyt", "colruyt_saturday"}:
        return "Hendrikx"
    if customer_key == "heeren":
        return "SVZ"
    return ""


def _group_results_by_sheet_and_key(results: dict) -> dict[str, dict[str, list[dict]]]:
    grouped: dict[str, dict[str, list[dict]]] = {}
    for entry in results.values():
        sheet = entry.get("sheet")
        customer_key = entry.get("masterKey") or ""
        if not sheet or not customer_key:
            continue
        grouped.setdefault(sheet, {}).setdefault(customer_key, []).append(copy.deepcopy(entry))

    for sheet_entries in grouped.values():
        for items in sheet_entries.values():
            items.sort(key=lambda item: item.get("savedAt", ""))

    return grouped


def _format_workbook_date(offset: int, base_date_value: str | None = None) -> str:
    base_date = _parse_base_date(base_date_value)
    target = base_date + timedelta(days=int(offset or 0))
    return target.strftime("%d/%m/%Y")


def _parse_base_date(value: str | None) -> datetime:
    if value:
        text = str(value).strip()
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
            try:
                return datetime.strptime(text, fmt)
            except ValueError:
                continue
    return datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)


def _normalize_workbook_date(value: str | None) -> str:
    text = str(value or "").strip().replace(".", "/").replace("-", "/")
    parts = [part for part in text.split("/") if part]
    if len(parts) != 3:
        return ""
    day, month, year = parts
    return f"{day.zfill(2)}/{month.zfill(2)}/{year}"


def trim_export_sheet(sheet_root: ET.Element, max_row: int) -> None:
    sheet_data = sheet_root.find("x:sheetData", NS)
    if sheet_data is not None:
        for row in list(sheet_data.findall("x:row", NS)):
            row_number = int(row.get("r", "0"))
            if row_number > max_row:
                sheet_data.remove(row)
                continue
            for cell in list(row.findall("x:c", NS)):
                if _extract_column_letters(cell.get("r", "")) > "F":
                    row.remove(cell)

    merge_cells = sheet_root.find("x:mergeCells", NS)
    if merge_cells is not None:
        for merge_cell in list(merge_cells.findall("x:mergeCell", NS)):
            start_ref = merge_cell.get("ref", "").split(":", 1)[0]
            start_row = _extract_row_number(start_ref)
            start_col = _extract_column_letters(start_ref)
            if start_row > max_row or start_col > "F":
                merge_cells.remove(merge_cell)
        merge_cells.set("count", str(len(merge_cells.findall("x:mergeCell", NS))))
        if len(merge_cells.findall("x:mergeCell", NS)) == 0:
            sheet_root.remove(merge_cells)

    drawing = sheet_root.find("x:drawing", NS)
    if drawing is not None:
        sheet_root.remove(drawing)

    dimension = sheet_root.find("x:dimension", NS)
    if dimension is not None:
        dimension.set("ref", f"A1:F{max_row}")


def _build_single_sheet_workbook(workbook_bytes: bytes, selected_sheet: str) -> ET.Element:
    workbook_root = ET.fromstring(workbook_bytes)
    sheets = workbook_root.find("x:sheets", NS)
    if sheets is None:
        raise ValueError("Workbook sheets could not be found in Leverschema template.")

    selected_sheet_element = None
    for sheet in sheets.findall("x:sheet", NS):
        if sheet.get("name") == selected_sheet:
            selected_sheet_element = copy.deepcopy(sheet)
            break
    if selected_sheet_element is None:
        raise ValueError(f"Sheet {selected_sheet} could not be found in Leverschema template.")

    for sheet in list(sheets):
        sheets.remove(sheet)
    sheets.append(selected_sheet_element)

    defined_names = workbook_root.find("x:definedNames", NS)
    if defined_names is not None:
        workbook_root.remove(defined_names)

    custom_views = workbook_root.find("x:customWorkbookViews", NS)
    if custom_views is not None:
        workbook_root.remove(custom_views)

    return workbook_root


def _build_single_sheet_app_properties(selected_sheet: str) -> ET.Element:
    root = ET.Element(f"{{{APP_NS}}}Properties")
    ET.SubElement(root, f"{{{APP_NS}}}TotalTime").text = "0"
    ET.SubElement(root, f"{{{APP_NS}}}Application").text = "Microsoft Excel"
    ET.SubElement(root, f"{{{APP_NS}}}DocSecurity").text = "0"
    ET.SubElement(root, f"{{{APP_NS}}}ScaleCrop").text = "false"

    heading_pairs = ET.SubElement(root, f"{{{APP_NS}}}HeadingPairs")
    heading_vector = ET.SubElement(heading_pairs, f"{{{APP_VT_NS}}}vector", size="2", baseType="variant")
    variant_1 = ET.SubElement(heading_vector, f"{{{APP_VT_NS}}}variant")
    ET.SubElement(variant_1, f"{{{APP_VT_NS}}}lpstr").text = "Worksheets"
    variant_2 = ET.SubElement(heading_vector, f"{{{APP_VT_NS}}}variant")
    ET.SubElement(variant_2, f"{{{APP_VT_NS}}}i4").text = "1"

    titles = ET.SubElement(root, f"{{{APP_NS}}}TitlesOfParts")
    titles_vector = ET.SubElement(titles, f"{{{APP_VT_NS}}}vector", size="1", baseType="lpstr")
    ET.SubElement(titles_vector, f"{{{APP_VT_NS}}}lpstr").text = selected_sheet

    ET.SubElement(root, f"{{{APP_NS}}}Manager").text = ""
    ET.SubElement(root, f"{{{APP_NS}}}Company").text = "Hessing Supervers"
    ET.SubElement(root, f"{{{APP_NS}}}LinksUpToDate").text = "false"
    ET.SubElement(root, f"{{{APP_NS}}}SharedDoc").text = "false"
    ET.SubElement(root, f"{{{APP_NS}}}HyperlinkBase").text = ""
    ET.SubElement(root, f"{{{APP_NS}}}HyperlinksChanged").text = "false"
    ET.SubElement(root, f"{{{APP_NS}}}AppVersion").text = "16.0300"
    return root


def _extract_row_number(cell_ref: str) -> int:
    digits = "".join(character for character in cell_ref if character.isdigit())
    return int(digits) if digits else 0


def _extract_column_letters(cell_ref: str) -> str:
    return "".join(character for character in cell_ref if character.isalpha())


def set_numeric_cell(sheet_root: ET.Element, cell_ref: str, value: float) -> None:
    cell = _find_cell(sheet_root, cell_ref)
    if cell is None:
        raise ValueError(f"Cell {cell_ref} was not found in the leverschema template.")

    formula_node = cell.find("x:f", NS)
    if formula_node is not None:
        cell.remove(formula_node)
    cell.attrib.pop("t", None)

    value_node = cell.find("x:v", NS)
    if value_node is None:
        value_node = ET.SubElement(cell, f"{{{MAIN_NS}}}v")
    value_node.text = _format_number(value)


def set_text_cell(sheet_root: ET.Element, cell_ref: str, value: str) -> None:
    cell = _find_cell(sheet_root, cell_ref)
    if cell is None:
        raise ValueError(f"Cell {cell_ref} was not found in the leverschema template.")

    formula_node = cell.find("x:f", NS)
    if formula_node is not None:
        cell.remove(formula_node)

    value_node = cell.find("x:v", NS)
    if value_node is not None:
        cell.remove(value_node)

    inline_node = cell.find("x:is", NS)
    if inline_node is not None:
        cell.remove(inline_node)

    cell.set("t", "inlineStr")
    inline_node = ET.SubElement(cell, f"{{{MAIN_NS}}}is")
    text_node = ET.SubElement(inline_node, f"{{{MAIN_NS}}}t")
    text_node.text = value


def clear_cell_value(sheet_root: ET.Element, cell_ref: str) -> None:
    cell = _find_cell(sheet_root, cell_ref)
    if cell is None:
        return

    formula_node = cell.find("x:f", NS)
    if formula_node is not None:
        cell.remove(formula_node)
    cell.attrib.pop("t", None)

    inline_node = cell.find("x:is", NS)
    if inline_node is not None:
        cell.remove(inline_node)

    value_node = cell.find("x:v", NS)
    if value_node is not None:
        cell.remove(value_node)


def _find_cell(sheet_root: ET.Element, cell_ref: str) -> ET.Element | None:
    return sheet_root.find(f".//x:c[@r='{cell_ref}']", NS)


def _format_number(value: float) -> str:
    rounded = int(value) if float(value).is_integer() else int(value)
    return str(rounded)
