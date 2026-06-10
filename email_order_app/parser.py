from __future__ import annotations

import re
import unicodedata
from datetime import datetime
from email import policy
from email.parser import BytesParser
from email.utils import parsedate_to_datetime
from pathlib import Path

from .models import CustomerOrder, OrderItem, ParsedOrderEmail


HEADER_MARKERS = {
    "item number",
    "artikelnummer",
}

ALLOWED_UNITS = {
    "collo",
    "box",
    "boxes",
    "stuk",
    "stuck",
}


def parse_order_email(eml_path: str | Path) -> ParsedOrderEmail:
    path = Path(eml_path)
    with path.open("rb") as handle:
        message = BytesParser(policy=policy.default).parse(handle)

    subject = str(message.get("subject", "")).strip()
    sender = str(message.get("from", "")).strip()
    received_at = _parse_email_datetime(message.get("date"))
    body_text = _extract_plain_text(message)
    lines = _clean_lines(body_text)
    sections = _split_order_sections(lines, subject)
    orders = [_parse_section(section, subject) for section in sections]
    orders = [order for order in orders if order.items]

    if not orders:
        raise ValueError("No order blocks were found in the email body.")

    delivery_date = _extract_delivery_date(subject, lines, received_at)
    leaving_date = received_at.strftime("%d/%m/%Y") if received_at else ""
    leaving_time = received_at.strftime("%H:%M") if received_at else ""

    return ParsedOrderEmail(
        source_file=path.name,
        subject=subject,
        sender=sender,
        received_at=received_at,
        delivery_date_to_dc=delivery_date,
        leaving_date_venlo=leaving_date,
        leaving_time_venlo=leaving_time,
        orders=orders,
    )


def _parse_email_datetime(value: str | None) -> datetime | None:
    if not value:
        return None
    try:
        return parsedate_to_datetime(value)
    except (TypeError, ValueError, IndexError):
        return None


def _extract_plain_text(message) -> str:
    text_parts: list[str] = []
    for part in message.walk():
        if part.get_content_type() == "text/plain" and str(part.get_content_disposition()) != "attachment":
            text_parts.append(part.get_content())

    if text_parts:
        return "\n".join(text_parts)

    content = message.get_body(preferencelist=("plain", "html"))
    if content:
        return content.get_content()
    return ""


def _clean_lines(body_text: str) -> list[str]:
    normalized = body_text.replace("\r", "")
    raw_lines = [line.strip() for line in normalized.split("\n")]
    return [line for line in raw_lines if line]


def _is_rewe_penny_email(lines: list[str], subject: str) -> bool:
    """Centralized detection for REWE/Penny emails."""
    # Check subject first (most reliable)
    subject_lower = subject.lower()
    if "rewe" in subject_lower or "penny" in subject_lower:
        return True
    
    # Check first 20 lines of content (more comprehensive)
    content_text = " ".join(lines[:20]).lower()
    if "rewe" in content_text or "penny" in content_text:
        return True
    
    return False


def _split_order_sections(lines: list[str], subject: str = "") -> list[list[str]]:
    # Check if this is a REWE/Penny email using centralized detection
    if _is_rewe_penny_email(lines, subject):
        # For REWE/Penny emails, treat all content as one section
        return [lines]
    
    # Original logic for other emails
    sections: list[list[str]] = []
    current: list[str] = []

    for line in lines:
        if "vcso" in line.lower():
            if current:
                sections.append(current)
            current = [line]
            continue

        if current:
            current.append(line)

    if current:
        sections.append(current)

    return sections


def _parse_section(section_lines: list[str], subject: str) -> CustomerOrder:
    reference_line = section_lines[0]
    items = _extract_items(section_lines, subject)
    
    # Use centralized detection
    if _is_rewe_penny_email(section_lines, subject):
        # For REWE/Penny emails, extract customer from subject
        subject_customer = _extract_customer_from_subject(subject)
        
        # Extract just the location part (remove REWE prefix)
        location_only = subject_customer
        if location_only.lower().startswith("rewe "):
            location_only = location_only[5:].strip()  # Remove "REWE " prefix
        elif location_only.lower().startswith("penny "):
            location_only = location_only[6:].strip()  # Remove "PENNY " prefix

        if _is_generic_rewe_penny_location(location_only):
            location_only = _extract_rewe_penny_delivery_point(section_lines)
        
        customer_dc = location_only or "REWE/Penny"
        
        content_text = " ".join(section_lines[:20]).lower()
        fatrans_dc = "Penny" if "penny" in subject.lower() or "penny" in content_text else "Rewe"
        
        # Find VCSO reference in the email content
        reference = ""
        for line in section_lines:
            vcso_match = re.search(r'(VCSO\d+)', line, re.IGNORECASE)
            if vcso_match:
                reference = vcso_match.group(1).upper()
                break
        
        # If no VCSO found, try to extract from first line
        if not reference:
            reference = _extract_reference(reference_line)
    else:
        # Original logic for other emails
        customer = _extract_customer_from_reference(reference_line)
        subject_customer = _extract_customer_from_subject(subject)
        customer_dc = customer or subject_customer or "Unknown customer"
        fatrans_dc = _resolve_fatrans_dc(subject_customer, customer_dc, reference_line)
        reference = _extract_reference(reference_line)
    
    slotboeking_id = _extract_slotboeking_id(section_lines)

    return CustomerOrder(
        customer_dc=customer_dc,
        fatrans_dc=fatrans_dc,
        reference=reference,
        items=items,
        slotboeking_id=slotboeking_id,
    )


def _extract_items(lines: list[str], subject: str = "") -> list[OrderItem]:
    # Use centralized detection
    if _is_rewe_penny_email(lines, subject):
        return _extract_rewe_penny_items(lines)
    else:
        return _extract_standard_items(lines)


def _extract_rewe_penny_items(lines: list[str]) -> list[OrderItem]:
    """Extract items from REWE/Penny email format with intelligent parsing.
    
    This parser is designed to be flexible and handle various table formats:
    - Different number of columns (3, 4, 5, or more)
    - Different column orders
    - With or without headers
    - Tabular or line-by-line format
    - With or without explicit unit column (defaults to "Collo")
    """
    items: list[OrderItem] = []
    
    # Find where the table starts - look for header row with "Item number" or "Quantity"
    table_start_index = None
    has_header = False
    
    for i, line in enumerate(lines):
        line_lower = line.lower()
        # Look for common header keywords
        if any(keyword in line_lower for keyword in [
            "item number", "artikelnummer", "artikel", "quantity", "menge",
            "description", "beschreibung", "desciption", "unit", "einheit"
        ]):
            # Table starts after the header
            table_start_index = i + 1
            has_header = True
            break
    
    if table_start_index is None:
        # Fallback: look for lines that look like article numbers (6 digits starting with 32 or 28)
        for i, line in enumerate(lines):
            line_clean = line.strip()
            if re.match(r'^(32|28)\d{4}$', line_clean):
                table_start_index = i
                break
    
    if table_start_index is None:
        return []
    
    # Strategy 1: Try to parse as structured table (all data in consecutive lines)
    i = table_start_index
    consecutive_failures = 0
    
    while i < len(lines) and consecutive_failures < 5:
        current_line = lines[i].strip()
        
        # Skip empty lines
        if not current_line:
            i += 1
            continue
        
        # Check if this looks like an article number (32xxxx or 28xxxx format)
        if re.match(r'^(32|28)\d{4}$', current_line):
            article_number = current_line
            
            # Collect the next several lines as potential data
            potential_data = []
            for j in range(1, 10):  # Look up to 9 lines ahead
                if i + j >= len(lines):
                    break
                next_line = lines[i + j].strip()
                if not next_line:
                    continue
                # Stop if we hit another article number
                if re.match(r'^(32|28)\d{4}$', next_line):
                    break
                potential_data.append(next_line)
            
            # Now intelligently extract description, quantity, and unit from potential_data
            description = None
            quantity = None
            unit = None
            
            for data_line in potential_data:
                # Skip reference codes (SSCO, VCSO, etc.)
                if re.match(r'^(SSCO|VCSO|REF)\d+', data_line, re.IGNORECASE):
                    continue
                
                # Check if it's a quantity (number with optional comma/dot)
                if quantity is None and _looks_like_quantity(data_line):
                    quantity = data_line
                    continue
                
                # Check if it's a unit
                if unit is None and data_line.lower() in ALLOWED_UNITS:
                    unit = data_line
                    continue
                
                # Check if it's a description (text that's not quantity or unit)
                if (description is None and 
                    not _looks_like_quantity(data_line) and 
                    data_line.lower() not in ALLOWED_UNITS and
                    len(data_line) > 2):  # Reasonable description length
                    description = data_line
                    continue
            
            # If we found description and quantity but no unit, default to "Collo"
            if description and quantity and not unit:
                unit = "Collo"
            
            # If we found all required data, add the item
            if description and quantity and unit:
                try:
                    items.append(
                        OrderItem(
                            article_number=article_number,
                            description=description,
                            quantity_boxes=_parse_quantity(quantity),
                            unit=unit,
                        )
                    )
                    consecutive_failures = 0  # Reset failure counter
                except:
                    consecutive_failures += 1
            else:
                consecutive_failures += 1
            
            # Move past this article number
            i += 1
        else:
            i += 1
        
        # Safety limit
        if len(items) > 100:
            break
    
    # Strategy 2: If we didn't find many items, try a more aggressive search
    # Look for patterns where article number, description, quantity, and unit might be on the same line
    # or in a different order
    if len(items) < 3:
        items_alt = []
        for i, line in enumerate(lines[table_start_index:], start=table_start_index):
            # Look for article numbers in the line
            article_match = re.search(r'\b(32|28)\d{4}\b', line)
            if article_match:
                article_number = article_match.group(0)
                
                # Try to extract other data from the same line or nearby lines
                search_lines = [line]
                # Add next 5 lines to search
                for j in range(1, 6):
                    if i + j < len(lines):
                        search_lines.append(lines[i + j])
                
                combined_text = " ".join(search_lines)
                
                # Look for quantity pattern (number followed by optional unit)
                quantity_match = re.search(r'\b(\d+[.,]?\d*)\s*(collo|box|pallet|stuk|pieces?|kg|stuks?)\b', 
                                          combined_text, re.IGNORECASE)
                
                quantity_text = None
                unit_text = None
                
                if quantity_match:
                    quantity_text = quantity_match.group(1)
                    unit_text = quantity_match.group(2)
                else:
                    # Look for standalone quantity (just a number)
                    for search_line in search_lines[1:]:  # Skip first line (has article number)
                        if _looks_like_quantity(search_line.strip()):
                            quantity_text = search_line.strip()
                            unit_text = "Collo"  # Default unit
                            break
                
                if quantity_text:
                    # Try to find description (text between article number and quantity)
                    desc_pattern = rf'{article_number}\s+(.+?)\s+{re.escape(quantity_text)}'
                    desc_match = re.search(desc_pattern, combined_text, re.IGNORECASE)
                    
                    description = None
                    if desc_match:
                        description = desc_match.group(1).strip()
                    else:
                        # Fallback: look for any text that looks like a description
                        for search_line in search_lines:
                            search_line_stripped = search_line.strip()
                            if (article_number not in search_line_stripped and
                                not _looks_like_quantity(search_line_stripped) and
                                search_line_stripped.lower() not in ALLOWED_UNITS and
                                len(search_line_stripped) > 5):
                                description = search_line_stripped
                                break
                    
                    if description and unit_text:
                        try:
                            items_alt.append(
                                OrderItem(
                                    article_number=article_number,
                                    description=description,
                                    quantity_boxes=_parse_quantity(quantity_text),
                                    unit=unit_text,
                                )
                            )
                        except:
                            pass
        
        # Use alternative results if they're better
        if len(items_alt) > len(items):
            items = items_alt
    
    return items



def _extract_standard_items(lines: list[str]) -> list[OrderItem]:
    """Extract items using the original logic for non-REWE/Penny emails."""
    header_index = _find_header_index(lines)
    if header_index is None:
        return []

    items: list[OrderItem] = []
    step = 5 if _has_requested_receipt_date_column(lines, header_index) else 4
    index = header_index + step
    while index + 3 < len(lines):
        article_number = lines[index]
        description = lines[index + 1]
        quantity_text = lines[index + 2]
        unit = lines[index + 3]
        block_size = step

        if _is_ssco_line(article_number) and index + 4 < len(lines):
            article_number = lines[index + 1]
            description = lines[index + 2]
            quantity_text = lines[index + 3]
            unit = lines[index + 4]
            block_size = 5

        if not _looks_like_article_number(article_number):
            break
        if not _looks_like_quantity(quantity_text):
            break
        if unit.lower() not in ALLOWED_UNITS:
            break

        items.append(
            OrderItem(
                article_number=article_number,
                description=description,
                quantity_boxes=_parse_quantity(quantity_text),
                unit=unit,
            )
        )
        index += block_size

    return items


def _find_header_index(lines: list[str]) -> int | None:
    for index, line in enumerate(lines):
        if line.lower() in HEADER_MARKERS and index + 3 < len(lines):
            return index
    return None


def _has_requested_receipt_date_column(lines: list[str], header_index: int) -> bool:
    return header_index + 4 < len(lines) and lines[header_index + 4].lower() == "requested receipt date"


def _looks_like_article_number(value: str) -> bool:
    return bool(re.fullmatch(r"\d{5,}", value))


def _looks_like_quantity(value: str) -> bool:
    return bool(re.fullmatch(r"\d{1,3}(?:[.,]\d{3})*(?:[.,]\d+)?", value))


def _is_ssco_line(value: str) -> bool:
    return bool(re.fullmatch(r"SSCO\d+", value, flags=re.IGNORECASE))


def _parse_quantity(value: str) -> float:
    import re
    cleaned = re.sub(r'[^\d.,]', '', value.strip())
    if not cleaned:
        raise ValueError("No number found")
    if '.' in cleaned and ',' in cleaned:
        if cleaned.rfind(',') > cleaned.rfind('.'):
            cleaned = cleaned.replace('.', '').replace(',', '.')
        else:
            cleaned = cleaned.replace(',', '')
    elif ',' in cleaned:
        cleaned = cleaned.replace(',', '.')
    return float(cleaned)


def _extract_reference(reference_line: str) -> str:
    match = re.search(r"(VCSO\d+)", reference_line, re.IGNORECASE)
    return match.group(1).upper() if match else ""


def _extract_slotboeking_id(section_lines: list[str]) -> str | None:
    """Extract ID slotboeking (e.g., L22293033) from Colruyt order section.
    
    First tries to find "ID slotboeking: L..." pattern.
    If not found, searches for any standalone reference starting with L followed by digits.
    Returns None if no valid reference is found.
    """
    # First try: explicit "ID slotboeking:" pattern
    for line in section_lines:
        match = re.search(r"ID\s+slotboeking\s*:\s*(L\d+)", line, re.IGNORECASE)
        if match:
            return match.group(1)
    
    # Second try: search for any L followed by 8+ digits (typical format: L22293033)
    for line in section_lines:
        # Look for L followed by at least 8 digits, not part of a larger word
        match = re.search(r"\b(L\d{8,})\b", line, re.IGNORECASE)
        if match:
            return match.group(1).upper()
    
    return None


def _extract_fatrans_dc(reference_line: str) -> str:
    if ":" in reference_line:
        return reference_line.split(":", 1)[1].strip()
    return reference_line.strip()


def _extract_customer_from_reference(reference_line: str) -> str:
    return _extract_fatrans_dc(reference_line)


def _extract_customer_from_subject(subject: str) -> str:
    subject_clean = subject.replace("->", "|").replace("_", " ")
    subject_clean = re.sub(r"\blevering\b", "", subject_clean, flags=re.IGNORECASE)
    subject_clean = re.sub(r"\bbestelling binnen\b", "", subject_clean, flags=re.IGNORECASE)
    subject_clean = re.sub(r"\b(?:LD|DD)\b.*$", "", subject_clean, flags=re.IGNORECASE)
    subject_clean = re.sub(r"\s{2,}", " ", subject_clean)
    parts = [part.strip(" -") for part in subject_clean.split("|") if part.strip(" -")]
    return parts[0] if parts else ""


def _extract_rewe_penny_delivery_point(lines: list[str]) -> str:
    label_patterns = [
        r"(?:delivery\s*point|delivery\s*name|ship[-\s]*to|ship\s*to\s*party|location|warehouse)\s*[:\-]\s*(.+)",
        r"(?:lieferstelle|lieferadresse|warenempf[aä]nger|empf[aä]nger|ablade(?:stelle)?|standort)\s*[:\-]\s*(.+)",
    ]
    ignored = {
        "delivery point",
        "delivery name",
        "ship-to",
        "ship to",
        "ship to party",
        "lieferstelle",
        "lieferadresse",
        "warenempfänger",
        "warenempfaenger",
        "empfänger",
        "empfaenger",
        "location",
    }

    for index, line in enumerate(lines[:60]):
        cleaned = line.strip()
        normalized = _normalize_text(cleaned)
        for pattern in label_patterns:
            match = re.search(pattern, cleaned, re.IGNORECASE)
            if match:
                candidate = _clean_rewe_penny_delivery_point(match.group(1))
                if candidate:
                    return candidate
        if normalized in ignored and index + 1 < len(lines):
            candidate = _clean_rewe_penny_delivery_point(lines[index + 1])
            if candidate:
                return candidate

    return ""


def _is_generic_rewe_penny_location(value: str) -> bool:
    normalized = _normalize_text(value)
    return normalized in {
        "",
        "rewe",
        "penny",
        "rewe penny",
        "rewe/penny",
        "order",
        "orders",
        "bestellung",
        "bestellungen",
        "purchase order",
        "po",
    }


def _clean_rewe_penny_delivery_point(value: str) -> str:
    cleaned = re.sub(r"\s+", " ", str(value or "").strip(" -:\t"))
    normalized = _normalize_text(cleaned)
    if not cleaned or normalized in {"rewe", "penny", "rewe penny", "rewe/penny"}:
        return ""
    if re.search(r"\b(VCSO|SSCO)\d+\b", cleaned, re.IGNORECASE):
        return ""
    if re.search(r"\b(item number|artikelnummer|quantity|menge|description|beschreibung|unit|einheit)\b", cleaned, re.IGNORECASE):
        return ""
    if re.match(r"^(32|28)\d{4}$", cleaned):
        return ""
    if len(cleaned) > 80:
        return ""
    return re.sub(r"^(REWE|PENNY)\s+", "", cleaned, flags=re.IGNORECASE).strip()


def _extract_dc_name(value: str) -> str:
    cleaned = value.strip()
    cleaned = re.sub(
        r"^(REWE|PENNY|NETTO|EDEKA|CARREFOUR KDC|CARREFOUR|HAVI BELGIUM|HAVI LOGISTICS GMBH|HAVI)\s+",
        "",
        cleaned,
        flags=re.IGNORECASE,
    )
    return cleaned.strip(" -")


def _resolve_fatrans_dc(subject_customer: str, customer_dc: str, reference_line: str) -> str:
    if _is_netto_denmark_subject(subject_customer, customer_dc):
        return "Denemark"

    subject_dc = _extract_dc_name(subject_customer)
    return subject_dc or _extract_fatrans_dc(reference_line)


def _is_netto_denmark_subject(subject_customer: str, customer_dc: str) -> bool:
    subject_text = _normalize_text(subject_customer)
    customer_text = _normalize_text(customer_dc)
    return (
        "netto" in subject_text
        and "koge" in subject_text
        and "brabrand" in subject_text
        and customer_text.startswith("netto ")
    )


def _normalize_text(value: str) -> str:
    text = str(value or "")
    replacements = {
        "ø": "o",
        "Ø": "O",
        "å": "a",
        "Å": "A",
        "æ": "ae",
        "Æ": "Ae",
    }
    for source, target in replacements.items():
        text = text.replace(source, target)
    normalized = unicodedata.normalize("NFKD", text)
    ascii_only = normalized.encode("ascii", "ignore").decode("ascii")
    return ascii_only.lower()


def _extract_delivery_date(subject: str, lines: list[str], received_at: datetime | None) -> str:
    subject_match = re.search(r"\b(\d{1,2})-(\d{1,2})(?:-(\d{4}))?\b", subject)
    if subject_match:
        day, month, year = subject_match.groups()
        resolved_year = int(year) if year else (received_at.year if received_at else datetime.now().year)
        return f"{int(day):02d}/{int(month):02d}/{resolved_year:04d}"

    for line in lines:
        afspraak_match = re.search(r"Afspraak.*?:\s*(\d{4})-(\d{2})-(\d{2})", line, re.IGNORECASE)
        if afspraak_match:
            year, month, day = afspraak_match.groups()
            return f"{int(day):02d}/{int(month):02d}/{int(year):04d}"

    return received_at.strftime("%d/%m/%Y") if received_at else ""
