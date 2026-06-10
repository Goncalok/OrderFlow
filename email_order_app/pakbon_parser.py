"""Parser for pakbon (invoice) PDF files to extract item quantities."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

try:
    import pdfplumber
    HAS_PDFPLUMBER = True
except ImportError:
    HAS_PDFPLUMBER = False

try:
    import PyPDF2
    HAS_PYPDF2 = True
except ImportError:
    HAS_PYPDF2 = False


class PakbonItem:
    """Represents an item from a pakbon."""
    
    def __init__(self, article_number: str, description: str, quantity: float, unit: str = "Stuk"):
        self.article_number = article_number
        self.description = description
        self.quantity = quantity
        self.unit = unit
    
    def __repr__(self):
        return f"PakbonItem({self.article_number}, {self.description}, {self.quantity} {self.unit})"


def parse_pakbon_pdf(pdf_path: str | Path, section: str = "Emballage") -> list[PakbonItem]:
    """Parse a pakbon PDF file and extract items with quantities.
    
    Args:
        pdf_path: Path to the pakbon PDF file
        section: Section to parse - "Emballage" or "Goederen" (default: "Emballage")
        
    Returns:
        List of PakbonItem objects
    """
    pdf_path = Path(pdf_path)
    
    if not pdf_path.exists():
        raise FileNotFoundError(f"Pakbon file not found: {pdf_path}")
    
    # Try pdfplumber first (better table extraction)
    if HAS_PDFPLUMBER:
        return _parse_with_pdfplumber(pdf_path, section)
    elif HAS_PYPDF2:
        return _parse_with_pypdf2(pdf_path, section)
    else:
        raise ImportError("Neither pdfplumber nor PyPDF2 is installed. Please install one of them.")


def _parse_with_pdfplumber(pdf_path: Path, section: str = "Emballage") -> list[PakbonItem]:
    """Parse pakbon using pdfplumber (preferred method)."""
    import pdfplumber
    
    items = []
    
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            # Extract text to find the specified section
            text = page.extract_text()
            if not text:
                continue
            
            lines = text.split('\n')
            
            # Find the section (Emballage or Goederen)
            section_start = None
            for i, line in enumerate(lines):
                if section in line:
                    section_start = i
                    break
            
            if section_start is None:
                continue
            
            # Find the header row after section name
            header_idx = None
            for i in range(section_start, min(section_start + 5, len(lines))):
                if 'ArtNr' in lines[i] and 'Omschrijving' in lines[i]:
                    header_idx = i
                    break
            
            if header_idx is None:
                continue
            
            # Parse items after the header until we hit "Totaal", "Emballage", or another section
            for i in range(header_idx + 1, len(lines)):
                line = lines[i].strip()
                
                # Stop at "Totaal", "Emballage" (if parsing Goederen), or empty lines
                if not line or 'Totaal' in line or 'doc.nr' in line:
                    break
                # If parsing Goederen, stop when we hit Emballage section
                if section == "Goederen" and 'Emballage' in line:
                    break
                
                # Split the line by whitespace
                parts = line.split()
                
                if len(parts) < 4:
                    continue
                
                # Try to parse: ArtNr Description... Unit Besteld Geleverd
                # Article number is usually the first part (numeric)
                if not parts[0].isdigit():
                    continue
                
                article_number = parts[0]
                
                # The last part should be the quantity (Geleverd)
                try:
                    quantity = float(parts[-1].replace(',', '.'))
                except (ValueError, IndexError):
                    continue
                
                # Find unit and numbers from the end
                # Format: ... Unit Besteld Geleverd
                # or: ... Unit - Geleverd
                unit = "Stuk"
                desc_end_idx = len(parts) - 1  # Start from the end
                
                # Skip the last number (Geleverd)
                desc_end_idx -= 1
                
                # Check if second-to-last is a number or "-" (Besteld)
                if desc_end_idx > 0 and (parts[desc_end_idx].replace('-', '').replace(',', '.').replace('.', '').isdigit() or parts[desc_end_idx] == '-'):
                    desc_end_idx -= 1
                
                # The part before numbers should be the unit
                if desc_end_idx > 0 and parts[desc_end_idx].lower() in ['stuk', 'collo', 'box', 'pallet']:
                    unit = parts[desc_end_idx]
                    desc_end_idx -= 1
                
                # Everything between article number and unit is the description
                description = ' '.join(parts[1:desc_end_idx + 1])
                
                if quantity > 0 and description:
                    items.append(PakbonItem(article_number, description, quantity, unit))
    
    return items


def _parse_with_pypdf2(pdf_path: Path, section: str = "Emballage") -> list[PakbonItem]:
    """Parse pakbon using PyPDF2 (fallback method)."""
    import PyPDF2
    
    items = []
    
    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        
        for page in reader.pages:
            text = page.extract_text()
            lines = text.split('\n')
            
            # Look for lines that match the pattern: ArtNr Description Quantity
            for i, line in enumerate(lines):
                # Try to find article number (usually numeric)
                match = re.search(r'\b(\d{3,})\b', line)
                if match:
                    article_number = match.group(1)
                    
                    # Look for quantity in the same line or nearby lines
                    quantity_match = re.search(r'\b(\d+(?:[.,]\d+)?)\s*(?:Stuk|Collo|Box)?\b', line)
                    if quantity_match:
                        try:
                            quantity = float(quantity_match.group(1).replace(',', '.'))
                            
                            # Try to extract description (text between article number and quantity)
                            desc_pattern = rf'{article_number}\s+(.+?)\s+{re.escape(quantity_match.group(1))}'
                            desc_match = re.search(desc_pattern, line)
                            description = desc_match.group(1).strip() if desc_match else ""
                            
                            if quantity > 0:
                                items.append(PakbonItem(article_number, description, quantity, "Stuk"))
                        except ValueError:
                            continue
    
    return items


def merge_pakbon_items(items: list[PakbonItem]) -> dict[str, PakbonItem]:
    """Merge items with the same article number, summing their quantities.
    
    Args:
        items: List of PakbonItem objects
        
    Returns:
        Dictionary mapping article_number to merged PakbonItem
    """
    merged = {}
    
    for item in items:
        if item.article_number in merged:
            # Sum quantities for same article number
            merged[item.article_number].quantity += item.quantity
        else:
            # Create a copy to avoid modifying original
            merged[item.article_number] = PakbonItem(
                item.article_number,
                item.description,
                item.quantity,
                item.unit
            )
    
    return merged


def parse_multiple_pakbons(pdf_paths: list[str | Path], section: str = "Emballage") -> dict[str, PakbonItem]:
    """Parse multiple pakbon PDFs and merge items with same article numbers.
    
    Args:
        pdf_paths: List of paths to pakbon PDF files
        section: Section to parse - "Emballage" or "Goederen" (default: "Emballage")
        
    Returns:
        Dictionary mapping article_number to merged PakbonItem
    """
    all_items = []
    
    for pdf_path in pdf_paths:
        try:
            items = parse_pakbon_pdf(pdf_path, section)
            all_items.extend(items)
        except Exception as e:
            print(f"Warning: Failed to parse {pdf_path}: {e}")
            continue
    
    return merge_pakbon_items(all_items)


def calculate_goederen_total(pdf_paths: list[str | Path]) -> float:
    """Calculate the total quantity from the Goederen section of pakbon PDFs.
    
    This is specifically for Colruyt CMR where we need to sum all Goederen quantities
    and insert the total in the PP cell.
    
    Args:
        pdf_paths: List of paths to pakbon PDF files
        
    Returns:
        Total quantity from all Goederen items
    """
    goederen_items = parse_multiple_pakbons(pdf_paths, section="Goederen")
    total = sum(item.quantity for item in goederen_items.values())
    return total
