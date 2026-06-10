from __future__ import annotations

import re
import unicodedata
from typing import Literal
import xml.etree.ElementTree as ET
from html import unescape
from dataclasses import dataclass
from email import policy
from email.parser import BytesParser
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from .simple_excel import WorkbookCell, write_simple_workbook

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
NS = {"x": MAIN_NS, "rel": REL_NS, "ct": CONTENT_TYPES_NS}
ET.register_namespace("", MAIN_NS)
ET.register_namespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
ET.register_namespace("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006")
ET.register_namespace("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac")
ET.register_namespace("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision")
ET.register_namespace("xr2", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2")
ET.register_namespace("xr3", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3")
ET.register_namespace("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main")
ET.register_namespace("x15ac", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac")
ET.register_namespace("xr6", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision6")
ET.register_namespace("xr10", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10")
ET.register_namespace("xcalcf", "http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures")


@dataclass(slots=True)
class NettoMdItem:
    article_number: str
    description: str
    quantity: float
    unit: str


@dataclass(slots=True)
class NettoMdOrder:
    customer_name: str
    delivery_name: str
    sales_order: str
    reference: str
    items: list[NettoMdItem]


@dataclass(slots=True)
class NettoMdEmail:
    source_file: str
    subject: str
    delivery_date: str
    dc_name: str
    orders: list[NettoMdOrder]


def _xl_target_path(target: str) -> str:
    target = (target or "").replace("\\", "/").strip().lstrip("/")
    if target.startswith("xl/"):
        return target
    return f"xl/{target}"


def _workbook_rel_id_to_target(file_map: dict[str, bytes]) -> dict[str, str]:
    rel_root = ET.fromstring(file_map["xl/_rels/workbook.xml.rels"])
    out: dict[str, str] = {}
    for rel in rel_root.findall(f"{{{REL_NS}}}Relationship"):
        rid = rel.get("Id")
        if rid:
            out[rid] = rel.get("Target") or ""
    return out


def _sheet_xml_key_for_name(file_map: dict[str, bytes], sheet_name: str) -> str | None:
    wb_root = ET.fromstring(file_map["xl/workbook.xml"])
    rel_map = _workbook_rel_id_to_target(file_map)
    want = sheet_name.strip().lower()
    sheets_el = wb_root.find(f"{{{MAIN_NS}}}sheets")
    if sheets_el is None:
        return None
    for sheet in sheets_el.findall(f"{{{MAIN_NS}}}sheet"):
        if (sheet.get("name") or "").strip().lower() != want:
            continue
        rid = sheet.get(f"{{{OFFICE_REL_NS}}}id")
        if not rid or rid not in rel_map:
            continue
        key = _xl_target_path(rel_map[rid])
        if key in file_map:
            return key
    return None


def _summary_sheet_xml_key(file_map: dict[str, bytes]) -> str | None:
    for key in sorted(file_map):
        if not re.fullmatch(r"xl/worksheets/sheet\d+\.xml", key):
            continue
        try:
            root = ET.fromstring(file_map[key])
        except ET.ParseError:
            continue
        cell = root.find(".//x:c[@r='C4']", NS)
        if cell is None:
            continue
        formula = cell.find("x:f", NS)
        if formula is not None and formula.text and "COUNTIF" in formula.text.upper():
            return key
    if "xl/worksheets/sheet2.xml" in file_map:
        return "xl/worksheets/sheet2.xml"
    return None


def parse_netto_md_email(eml_path: str | Path) -> NettoMdEmail | None:
    path = Path(eml_path)
    with path.open("rb") as handle:
        message = BytesParser(policy=policy.default).parse(handle)

    subject = str(message.get("subject", "")).strip()
    # Relaxed subject check: allow layout detection to proceed unless it's obviously NOT a Netto email
    # We remove the hard return None here and rely on _detect_netto_md_layout

    body = _extract_plain_text(message).replace("\r", "")
    lines = [line.strip() for line in body.split("\n") if line.strip()]
    layout_config = _detect_netto_md_layout(lines)
    if layout_config is None:
        return None
    header_index = layout_config.header_index

    reference_map = _extract_reference_map(lines[:header_index])
    orders = _extract_orders(lines, layout_config, reference_map)
    if not orders:
        return None

    return NettoMdEmail(
        source_file=path.name,
        subject=subject,
        delivery_date=_extract_delivery_date(lines, subject),
        dc_name="NettoMD",
        orders=orders,
    )


def write_netto_md_excel(template_path: str | Path, output_path: str | Path, data: NettoMdEmail) -> Path:
    template = Path(template_path)
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)

    if not template.is_file():
        return _write_netto_md_fallback(output, data)

    with ZipFile(template, "r") as source_archive:
        file_map = {name: source_archive.read(name) for name in source_archive.namelist()}

    required = ("xl/workbook.xml", "xl/sharedStrings.xml", "xl/_rels/workbook.xml.rels")
    if any(key not in file_map for key in required):
        return _write_netto_md_fallback(output, data)

    overview_key = _sheet_xml_key_for_name(file_map, "Overview")
    if overview_key is None or overview_key not in file_map:
        overview_key = "xl/worksheets/sheet1.xml"
    if overview_key not in file_map:
        return _write_netto_md_fallback(output, data)

    shared_strings_root = ET.fromstring(file_map["xl/sharedStrings.xml"])
    sheet_root = ET.fromstring(file_map[overview_key])
    workbook_root = ET.fromstring(file_map["xl/workbook.xml"])

    shared_strings = _read_shared_strings(shared_strings_root)
    block_map = _extract_template_blocks(sheet_root, shared_strings)
    _clear_amount_cells(sheet_root, block_map)
    _write_order_quantities(sheet_root, block_map, data.orders)

    summary_key = _summary_sheet_xml_key(file_map)
    if summary_key and summary_key in file_map:
        summary_root = ET.fromstring(file_map[summary_key])
        _fix_summary_formulas(summary_root)
        _normalize_sheet_ignorable(summary_root)
        file_map[summary_key] = ET.tostring(summary_root, encoding="utf-8", xml_declaration=True)

    _normalize_sheet_ignorable(sheet_root)
    _enable_full_recalculation(workbook_root)

    file_map[overview_key] = ET.tostring(sheet_root, encoding="utf-8", xml_declaration=True)
    file_map["xl/workbook.xml"] = ET.tostring(workbook_root, encoding="utf-8", xml_declaration=True)
    _remove_calc_chain(file_map)

    with ZipFile(output, "w", ZIP_DEFLATED) as target_archive:
        for name, content in file_map.items():
            target_archive.writestr(name, content)

    return output


def _write_netto_md_fallback(output: Path, data: NettoMdEmail) -> Path:
    rows: list[list[WorkbookCell]] = [
        [WorkbookCell("DC", "meta"), WorkbookCell(data.dc_name, "meta")],
        [WorkbookCell("Delivery Date", "meta"), WorkbookCell(data.delivery_date, "meta")],
        [],
        [
            WorkbookCell("Customer", "header"),
            WorkbookCell("Reference", "header"),
            WorkbookCell("Sales Order", "header"),
            WorkbookCell("Article", "header"),
            WorkbookCell("Description", "header"),
            WorkbookCell("Amount", "header"),
            WorkbookCell("Unit", "header"),
        ],
    ]

    for order in data.orders:
        for item in order.items:
            rows.append(
                [
                    WorkbookCell(order.customer_name),
                    WorkbookCell(order.reference),
                    WorkbookCell(order.sales_order),
                    WorkbookCell(item.article_number),
                    WorkbookCell(item.description),
                    WorkbookCell(item.quantity, "number"),
                    WorkbookCell(item.unit),
                ]
            )

    return write_simple_workbook(output, "Overview", rows, title="NettoMD Orderpicking")


def _extract_plain_text(message) -> str:
    text_parts: list[str] = []
    for part in message.walk():
        if part.get_content_type() == "text/plain" and str(part.get_content_disposition()) != "attachment":
            text_parts.append(part.get_content())
    if text_parts:
        return "\n".join(text_parts)

    html_parts: list[str] = []
    for part in message.walk():
        if part.get_content_type() == "text/html" and str(part.get_content_disposition()) != "attachment":
            html_parts.append(part.get_content())
    if html_parts:
        return _html_table_to_line_oriented_text("\n".join(html_parts))

    content = message.get_body(preferencelist=("plain", "html"))
    if content is None:
        return ""
    payload = content.get_content()
    if content.get_content_type() == "text/html":
        return _html_table_to_line_oriented_text(str(payload))
    return str(payload)


def _html_table_to_line_oriented_text(html: str) -> str:
    """Approximate Outlook plain-text export: one table cell per line (for NettoMD layout)."""
    text = unescape(html.replace("\r", "")).replace("\xa0", " ")
    text = re.sub(r"(?i)</t[dh]>", "\n", text)
    text = re.sub(r"(?i)<br\s*/?>", "\n", text)
    text = re.sub(r"<[^>]+>", "", text)
    lines: list[str] = [" ".join(line.split()) for line in text.split("\n")]
    return "\n".join(lines)


NettoMdLayout = Literal["full", "compact"]


@dataclass(slots=True)
class NettoMdConfig:
    header_index: int
    kind: NettoMdLayout
    offsets: dict[str, int]
    step: int


def _detect_netto_md_layout(lines: list[str]) -> NettoMdConfig | None:
    """Detection based on a cluster of expected headers. Returns a config with offsets."""
    
    # ULTRA-STRICT: Only proceed if we have VERY specific NettoMD indicators
    content_text = " ".join(lines[:50]).lower()
    
    # Must have BOTH NettoMD-specific indicators
    required_indicators = [
        "netto lager",  # NettoMD uses "Netto lager [Location]"
        "netto md",     # NettoMD specific product names
    ]
    
    # If we don't have BOTH indicators, this is NOT NettoMD
    if not all(indicator in content_text for indicator in required_indicators):
        return None
    
    # Additional verification: must have NettoMD location names
    netto_locations = [
        "bottrop", "ganderkesee", "hamm", "henstedt", "hodenhagen", 
        "kerpen", "krefeld"
    ]
    
    has_netto_location = any(location in content_text for location in netto_locations)
    if not has_netto_location:
        return None
    
    # CRITICAL: Reject if we see REWE/Penny indicators
    if "rewe" in content_text or "penny" in content_text:
        return None
    
    # Possible header names mapped to our internal keys
    # Order matters: check longer ones first to avoid partial matches (e.g. Sales orders vs Sales order)
    header_items = [
        ("sales orders", "sales_order_extra"),
        ("sales order", "sales_order"),
        ("delivery name", "delivery_name"),
        ("customer", "customer_id"),
        ("item number", "article_number"),
        ("article number", "article_number"),
        ("artikelnummer", "article_number"),  # Dutch variant
        ("item description", "description"),
        ("item desciption", "description"),  # Typo in some emails
        ("productnaam", "description"),  # Dutch variant
        ("quantity", "quantity"),
        ("hoeveelheid", "quantity"),  # Dutch variant
        ("unit", "unit"),
        ("eenheid", "unit"),  # Dutch variant
        ("delivery type", "delivery_type"),
    ]

    for index in range(len(lines) - 8):
        # Look at a slightly larger header window so layouts with extra columns still match.
        window = lines[index:index + 16]
        found_offsets: dict[str, int] = {}
        
        for offset, line in enumerate(window):
            clean_line = line.lower().strip()
            for h_text, h_key in header_items:
                if h_text in clean_line and h_key not in found_offsets:
                    found_offsets[h_key] = offset
                    break
        
        # We need at least these 3 for a valid parse (more lenient - removed delivery_name requirement)
        required = {"article_number", "quantity", "unit"}
        if all(r in found_offsets for r in required):
            header_start_offset = min(found_offsets.values())
            actual_offsets = {key: offset - header_start_offset for key, offset in found_offsets.items()}
            header_line = index + header_start_offset
            
            # If delivery_name is missing, use description as fallback
            if "delivery_name" not in actual_offsets and "description" in actual_offsets:
                # Set delivery_name to be before description (we'll extract it differently)
                actual_offsets["delivery_name"] = max(0, actual_offsets["description"] - 1)
            
            step = _choose_netto_md_step(lines, header_line, actual_offsets)
            header_index = header_line + step - 1
            kind: NettoMdLayout = "full" if "customer_id" in found_offsets else "compact"

            return NettoMdConfig(header_index=header_index, kind=kind, offsets=actual_offsets, step=step)
            
    return None


def _choose_netto_md_step(lines: list[str], header_line: int, offsets: dict[str, int]) -> int:
    minimum_step = max(offsets.values()) + 1
    best_step = minimum_step
    best_score = -1

    for step in range(minimum_step, minimum_step + 12):
        score = _score_netto_md_step(lines, header_line, offsets, step)
        if score > best_score:
            best_score = score
            best_step = step

    return best_step


def _score_netto_md_step(lines: list[str], header_line: int, offsets: dict[str, int], step: int) -> int:
    score = 0
    start = header_line + step

    for row_index in range(3):
        row_start = start + (row_index * step)
        if row_start + max(offsets.values()) >= len(lines):
            break
        if _looks_like_netto_md_row(lines, row_start, offsets):
            score += 1

    return score


def _looks_like_netto_md_row(lines: list[str], row_start: int, offsets: dict[str, int]) -> bool:
    # More lenient check - just need article number and quantity
    article_number = lines[row_start + offsets["article_number"]].strip() if row_start + offsets["article_number"] < len(lines) else ""
    quantity_text = lines[row_start + offsets["quantity"]].strip() if row_start + offsets["quantity"] < len(lines) else ""
    unit = lines[row_start + offsets["unit"]].strip() if row_start + offsets["unit"] < len(lines) else ""

    # Check if article number looks valid (at least 3 digits)
    if not re.search(r"\d{3,}", article_number):
        return False
    
    # Check if quantity has a number
    if not re.search(r"\d", quantity_text):
        return False
    
    # Check if unit is not a header keyword
    if unit.lower() in {"unit", "eenheid", "delivery type", "sales order", "requested receipt date", "artikelnummer", "productnaam", "hoeveelheid"}:
        return False
    
    # Optional: check delivery_name if it exists
    if "delivery_name" in offsets and row_start + offsets["delivery_name"] < len(lines):
        delivery_name = lines[row_start + offsets["delivery_name"]].strip()
        if delivery_name.lower() in {"delivery name", "customer", "item number", "article number", "artikelnummer", "productnaam"}:
            return False
    
    return True


def _extract_reference_map(lines: list[str]) -> dict[str, str]:
    mapping: dict[str, str] = {}
    for line in lines:
        match = re.match(r"^(.+?)\s*:\s*(VCSO\d+)\s*:", line, re.IGNORECASE)
        if not match:
            continue
        mapping[_normalize_name(match.group(1))] = match.group(2).upper()
    return mapping


def _extract_orders(
    lines: list[str],
    config: NettoMdConfig,
    reference_map: dict[str, str],
) -> list[NettoMdOrder]:
    grouped: dict[str, NettoMdOrder] = {}
    
    # Extract relative offsets from config
    o_name = config.offsets.get("delivery_name", -1)
    o_art = config.offsets.get("article_number", 0)
    o_desc = config.offsets.get("description", 1)
    o_qty = config.offsets.get("quantity", 2)
    o_unit = config.offsets.get("unit", 3)
    o_id = config.offsets.get("customer_id", -1) # only for full
    o_so = config.offsets.get("sales_order", -1)
    o_type = config.offsets.get("delivery_type", -1)

    index = config.header_index + 1
    max_offset = max(config.offsets.values())
    
    # Track current customer from VCSO lines or header
    current_customer = "NettoMD"
    current_sales_order = ""
    
    # Check if this is a line-by-line format (step == number of fields)
    # In this case, each field is on a separate line
    is_line_by_line = config.step == len(config.offsets)
    
    while index + max_offset < len(lines):
        # Check if this line contains a VCSO reference (new customer section)
        line_text = lines[index] if index < len(lines) else ""
        vcso_match = re.search(r'(VCSO\d+)\s*:\s*(.+)', line_text, re.IGNORECASE)
        if vcso_match:
            current_sales_order = vcso_match.group(1).upper()
            current_customer = vcso_match.group(2).strip()
            index += 1
            continue
        
        # Skip lines that look like sales order numbers alone (SSCO format)
        if re.fullmatch(r'SSCO\d+', line_text.strip(), re.IGNORECASE):
            index += 1
            continue
        
        # Get field values based on offsets
        article_number = lines[index + o_art] if index + o_art < len(lines) else ""
        description = lines[index + o_desc] if o_desc >= 0 and index + o_desc < len(lines) else ""
        quantity_text = lines[index + o_qty] if index + o_qty < len(lines) else ""
        unit = lines[index + o_unit] if index + o_unit < len(lines) else ""
        
        # More lenient validation: 3+ digits for articles
        if not re.search(r"\d{3,}", article_number):
            index += 1
            continue
            
        # Lenient quantity: allow empty or invalid, but skip if definitely not a number
        if not re.search(r"\d", quantity_text):
            index += 1
            continue
        
        # Skip if unit looks like a header
        if unit.lower() in {"unit", "eenheid", "delivery type", "sales order", "requested receipt date"}:
            index += 1
            continue
        
        # Determine delivery name
        if o_name >= 0 and index + o_name < len(lines):
            delivery_name = lines[index + o_name]
            # Skip if it's a digit-only line (likely a sales order number)
            if delivery_name.isdigit() and len(delivery_name) < 3:
                index += 1
                continue
            # Skip if it starts with SSCO or VCSO
            if re.match(r'(SSCO|VCSO)\d+', delivery_name, re.IGNORECASE):
                index += 1
                continue
        else:
            delivery_name = current_customer

        customer_name = _extract_customer_name(delivery_name) if delivery_name else current_customer
        
        # Skip if customer name is empty or invalid after cleaning
        if not customer_name or len(customer_name.strip()) < 3:
            index += 1
            continue
        
        key = _normalize_name(customer_name)
        
        # Skip if normalized key is too short (likely invalid)
        if len(key) < 3:
            index += 1
            continue
        
        sales_order = current_sales_order
        if o_so >= 0 and index + o_so < len(lines):
            so_value = lines[index + o_so]
            if so_value and not so_value.lower() in {"sales order", "sales orders"}:
                sales_order = so_value
            
        if o_type >= 0 and index + o_type < len(lines):
            delivery_type = lines[index + o_type]
            if not _is_netto_md_direct_delivery(delivery_type):
                # If we detected a valid article but it's not direct delivery, 
                # we skip this BLOCK to avoid partial matches
                index += 1
                continue

        order = grouped.get(key)
        if order is None:
            order = NettoMdOrder(
                customer_name=customer_name,
                delivery_name=delivery_name if delivery_name else customer_name,
                sales_order=sales_order,
                reference=reference_map.get(key, sales_order),
                items=[],
            )
            grouped[key] = order

        try:
            order.items.append(
                NettoMdItem(
                    article_number=article_number.strip(),
                    description=description.strip(),
                    quantity=_parse_quantity(quantity_text),
                    unit=unit.strip(),
                )
            )
        except (ValueError, AttributeError) as e:
            # Skip invalid items but continue processing
            pass
        
        # Move forward by step if line-by-line, otherwise just 1
        index += config.step if is_line_by_line else 1

    return list(grouped.values())


def _extract_customer_name(delivery_name: str) -> str:
    """Extract clean customer name from delivery name, removing extra data."""
    # Skip if it's a Content-ID or other technical identifier
    if re.match(r'\[cid:', delivery_name, re.IGNORECASE):
        return ""
    
    # Skip if it's mostly special characters or looks like a UUID/hash
    if re.search(r'[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}', delivery_name, re.IGNORECASE):
        return ""
    
    # Skip if it contains too many special characters (likely not a real name)
    special_char_count = len(re.findall(r'[^a-zA-Z0-9\s\-]', delivery_name))
    if special_char_count > 3:
        return ""
    
    # Remove "Netto lager" prefix
    cleaned = re.sub(r"^Netto\s+lager\s+", "", delivery_name, flags=re.IGNORECASE)
    cleaned = re.sub(r"^Netto\s+Lager\s+", "", cleaned, flags=re.IGNORECASE)
    
    # Remove article numbers (5+ digits)
    cleaned = re.sub(r'\b\d{5,}\b', '', cleaned)
    
    # Remove quantities (numbers with units like "kg", "st", "stk", etc.)
    cleaned = re.sub(r'\b\d+[.,]?\d*\s*(kg|st|stk|pc|pcs|piece|pieces|unit|units|eenheid)\b', '', cleaned, flags=re.IGNORECASE)
    
    # Remove standalone numbers that might be quantities
    cleaned = re.sub(r'\b\d+[.,]?\d*\b', '', cleaned)
    
    # Clean up extra whitespace
    cleaned = ' '.join(cleaned.split())
    
    # List of non-customer terms to filter out (German/Dutch/English)
    non_customer_terms = [
        'vertriebsinnendienst',  # Inside Sales Department
        'vertrieb',              # Sales
        'innendienst',           # Inside Service
        'collo',                 # Package/Parcel
        'collos',                # Packages
        'delivery type',         # Header field
        'sales order',           # Header field
        'customer',              # Header field
        'item number',           # Header field
        'article number',        # Header field
        'quantity',              # Header field
        'requested receipt',     # Header field
        'receipt date',          # Header field
        'item description',      # Header field
        'item desciption',       # Header field (typo variant)
        'unit',                  # Header field
        'eenheid',              # Unit (Dutch)
        'hoeveelheid',          # Quantity (Dutch)
        'artikelnummer',        # Article number (Dutch)
        'productnaam',          # Product name (Dutch)
    ]
    
    # Check if cleaned name matches any non-customer term
    cleaned_lower = cleaned.lower().strip()
    for term in non_customer_terms:
        if cleaned_lower == term or term in cleaned_lower:
            return ""
    
    # Final validation: must have at least 3 letters
    if len(re.findall(r'[a-zA-Z]', cleaned)) < 3:
        return ""
    
    return cleaned.strip()


def _parse_quantity(value: str) -> float:
    cleaned = value.strip().replace("\xa0", "").replace(" ", "")
    if "." in cleaned and "," in cleaned:
        if cleaned.rfind(",") > cleaned.rfind("."):
            cleaned = cleaned.replace(".", "").replace(",", ".")
        else:
            cleaned = cleaned.replace(",", "")
    elif "," in cleaned:
        cleaned = cleaned.replace(",", ".")
    return float(cleaned)


def _is_netto_md_direct_delivery(value: str) -> bool:
    normalized = " ".join(value.replace("\xa0", " ").lower().split())
    # Be extremely lenient: unless it says "LAGER" or "STOCK", consider it direct
    # especially since filtered emails are usually the ones intended for us.
    if not normalized: 
        return True
    if "lager" in normalized or "stock" in normalized:
        return False
    return True


def _extract_delivery_date(lines: list[str], subject: str) -> str:
    for index, line in enumerate(lines):
        if line.lower() == "requested receipt date" and index + 1 < len(lines):
            value = lines[index + 1]
            match = re.fullmatch(r"(\d{2})\.(\d{2})\.(\d{4})", value)
            if match:
                day, month, year = match.groups()
                return f"{day}/{month}/{year}"

    match = re.search(r"(\d{2})-(\d{2})(?:-(\d{4}))?", subject)
    if match:
        day, month, year = match.groups()
        return f"{day}/{month}/{year or '2026'}"
    return ""


def _normalize_name(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", value)
    ascii_only = normalized.encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^a-z0-9]+", "", ascii_only.lower())


def _extract_template_blocks(sheet_root: ET.Element, shared_strings: list[str]) -> dict[str, dict[str, object]]:
    rows = sheet_root.findall(".//x:sheetData/x:row", NS)
    blocks: dict[str, dict[str, object]] = {}
    current_key: str | None = None

    for row in rows:
        row_number = int(row.get("r", "0"))
        label = _cell_text(sheet_root, f"B{row_number}", shared_strings)
        article = _cell_text(sheet_root, f"C{row_number}", shared_strings)
        if label:
            current_key = _normalize_name(label)
            blocks[current_key] = {"label": label, "rows": {}}
            continue
        if current_key and article and re.fullmatch(r"\d{5,}", article):
            blocks[current_key]["rows"][article] = row_number

    return blocks


def _clear_amount_cells(sheet_root: ET.Element, block_map: dict[str, dict[str, object]]) -> None:
    for block in block_map.values():
        for row_number in block["rows"].values():
            _set_numeric_cell(sheet_root, f"E{row_number}", 0)


def _netto_md_block_for_order(block_map: dict[str, dict[str, object]], order: NettoMdOrder) -> dict[str, object] | None:
    for key in (_normalize_name(order.customer_name), _normalize_name(order.delivery_name)):
        block = block_map.get(key)
        if block is not None:
            return block
    return None


def _write_order_quantities(sheet_root: ET.Element, block_map: dict[str, dict[str, object]], orders: list[NettoMdOrder]) -> None:
    for order in orders:
        block = _netto_md_block_for_order(block_map, order)
        if block is None:
            continue
        row_map: dict[str, int] = block["rows"]
        for item in order.items:
            row_number = row_map.get(item.article_number)
            if row_number is None:
                continue
            _set_numeric_cell(sheet_root, f"E{row_number}", item.quantity)


def _enable_full_recalculation(workbook_root: ET.Element) -> None:
    calc_pr = workbook_root.find("x:calcPr", NS)
    if calc_pr is None:
        calc_pr = ET.SubElement(workbook_root, f"{{{MAIN_NS}}}calcPr")
    calc_pr.set("calcMode", "auto")
    calc_pr.set("fullCalcOnLoad", "1")
    calc_pr.set("forceFullCalc", "1")


def _fix_summary_formulas(summary_root: ET.Element) -> None:
    cell = summary_root.find(".//x:c[@r='C4']", NS)
    if cell is None:
        return
    formula = cell.find("x:f", NS)
    if formula is None:
        formula = ET.SubElement(cell, f"{{{MAIN_NS}}}f")
    formula.text = 'COUNTIF(Overview!F13:K21,"<>0")'


def _normalize_sheet_ignorable(sheet_root: ET.Element) -> None:
    ignorable_key = "{http://schemas.openxmlformats.org/markup-compatibility/2006}Ignorable"
    if ignorable_key in sheet_root.attrib:
        sheet_root.set(ignorable_key, "x14ac xr")


def _remove_calc_chain(file_map: dict[str, bytes]) -> None:
    file_map.pop("xl/calcChain.xml", None)

    workbook_rels_xml = file_map.get("xl/_rels/workbook.xml.rels")
    if workbook_rels_xml:
        rels_root = ET.fromstring(workbook_rels_xml)
        for rel in list(rels_root.findall("rel:Relationship", NS)):
            if rel.get("Type", "").endswith("/calcChain"):
                rels_root.remove(rel)
        file_map["xl/_rels/workbook.xml.rels"] = ET.tostring(rels_root, encoding="utf-8", xml_declaration=True)

    content_types_xml = file_map.get("[Content_Types].xml")
    if content_types_xml:
        content_root = ET.fromstring(content_types_xml)
        for node in list(content_root.findall("ct:Override", NS)):
            if node.get("PartName") == "/xl/calcChain.xml":
                content_root.remove(node)
        file_map["[Content_Types].xml"] = ET.tostring(content_root, encoding="utf-8", xml_declaration=True)


def _set_numeric_cell(sheet_root: ET.Element, cell_ref: str, value: float) -> None:
    cell = sheet_root.find(f".//x:c[@r='{cell_ref}']", NS)
    if cell is None:
        return
    cell.attrib.pop("t", None)
    formula = cell.find("x:f", NS)
    if formula is not None:
        cell.remove(formula)
    value_node = cell.find("x:v", NS)
    if value_node is None:
        value_node = ET.SubElement(cell, f"{{{MAIN_NS}}}v")
    value_node.text = _format_number(value)


def _format_number(value: float) -> str:
    return str(int(value)) if float(value).is_integer() else f"{value:.2f}".rstrip("0").rstrip(".")


def _read_shared_strings(shared_strings_root: ET.Element) -> list[str]:
    shared_strings: list[str] = []
    for item in shared_strings_root.findall("x:si", NS):
        shared_strings.append("".join((text.text or "") for text in item.iterfind(".//x:t", NS)))
    return shared_strings


def _cell_text(sheet_root: ET.Element, cell_ref: str, shared_strings: list[str]) -> str:
    cell = sheet_root.find(f".//x:c[@r='{cell_ref}']", NS)
    if cell is None:
        return ""
    inline = cell.find("x:is", NS)
    if inline is not None:
        return "".join((t.text or "") for t in inline.iterfind(".//x:t", NS))
    value_node = cell.find("x:v", NS)
    if value_node is None:
        return ""
    value = value_node.text or ""
    if cell.get("t") == "s":
        try:
            return shared_strings[int(value)]
        except (ValueError, IndexError):
            return ""
    return value
