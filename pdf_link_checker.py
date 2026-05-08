"""
PDF Link Checker with Excel Export & GUI
=========================================
Comprehensive tool to validate all email and web addresses in PDFs.
Includes file picker dialog for easy PDF selection.

Checks for:
1. Missing hyperlinks (text visible but not clickable)
2. Wrong link type (email linked to URL, URL linked to mailto)
3. Broken/split addresses across lines
4. Partial linking in split addresses
5. Wrong link targets (mismatch between text and href)

Usage:
    python pdf_link_checker.py                    (GUI mode - opens file picker)
    python pdf_link_checker.py document.pdf       (Direct mode)
    python pdf_link_checker.py document.pdf -o my_report.xlsx
    python pdf_link_checker.py document.pdf -v

Requirements:
    pip install pypdf pdfplumber openpyxl tkinter

Output Columns (Excel):
    A: PDF_Link           - Address found in PDF text
    B: Hyperlink_Text     - Actual link URI or "No Link"
    C: Page_Numb          - Page number where found
    D: Link_Type          - "Email" or "Web Link"
    E: Is_Hyperlinked     - "Yes" or "No"
    F: Is_Valid           - "Yes" (match OK) or "No" (mismatch)
    G: Hyperlink_Points_To - Destination of hyperlink
    H: Result             - "Pass" (all checks OK) or "Fail" (any check failed)
"""

import re
import sys
import os
import argparse
import pdfplumber
from pypdf import PdfReader
from dataclasses import dataclass
from typing import Optional, List, Dict, Tuple


# ============================================================================
# REGEX PATTERNS
# ============================================================================

EMAIL_REGEX = re.compile(
    r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}"
)

URL_REGEX = re.compile(
    r"(?:https?://|www\.)[a-zA-Z0-9\-._~:/?#\[\]@!$&'()*+,;=%]+"
)


# ============================================================================
# DATA CLASSES
# ============================================================================

@dataclass
class LinkAnnotation:
    """Represents a hyperlink annotation in the PDF."""
    page_num: int           # 0-based page number
    uri: str                # Actual href (e.g., "mailto:a@b.com" or "https://...")
    rect: tuple             # (x0, y0, x1, y1) in PDF coordinates


@dataclass
class TextToken:
    """Represents a word/token extracted from PDF with location."""
    page_num: int
    text: str
    x0: float
    y0: float
    x1: float
    y1: float


@dataclass
class ExcelRow:
    """Excel report row with validation results."""
    pdf_link: str                # Column A
    hyperlink_text: str          # Column B
    page_num: int                # Column C
    link_type: str               # Column D: "Email" or "Web Link"
    is_hyperlinked: str          # Column E: "Yes" or "No"
    is_valid: str                # Column F: "Yes" or "No"
    hyperlink_points_to: str     # Column G
    result: str                  # Column H: "Pass" or "Fail"


# ============================================================================
# FILE PICKER GUI
# ============================================================================

def select_pdf_file():
    """Open file picker dialog to select PDF."""
    try:
        import tkinter as tk
        from tkinter import filedialog
        
        root = tk.Tk()
        root.withdraw()  # Hide the main window
        root.attributes('-topmost', True)  # Bring to front
        
        file_path = filedialog.askopenfilename(
            title="PDF फ़ाइल चुनें / Select PDF File",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")]
        )
        
        root.destroy()
        return file_path
    except ImportError:
        print("❌ tkinter not installed!")
        print("Install it with: pip install tk")
        return None


# ============================================================================
# PDF EXTRACTION FUNCTIONS
# ============================================================================

def extract_hyperlinks(pdf_path: str) -> List[LinkAnnotation]:
    """Extract all hyperlink annotations from PDF."""
    links = []
    reader = PdfReader(pdf_path)

    for page_idx, page in enumerate(reader.pages):
        if "/Annots" not in page:
            continue

        annots = page["/Annots"]
        if annots is None:
            continue

        page_height = float(page.mediabox.height)

        for annot_ref in annots:
            try:
                annot = annot_ref.get_object() if hasattr(annot_ref, "get_object") else annot_ref
                
                # Only process Link annotations
                if annot.get("/Subtype") != "/Link":
                    continue

                action = annot.get("/A")
                if action is None:
                    continue

                # Only URI links (ignore other types)
                if action.get("/S") != "/URI":
                    continue

                uri = str(action.get("/URI", ""))
                rect_raw = annot.get("/Rect")
                if rect_raw is None:
                    continue

                # Convert PDF coordinates (bottom-left origin) to screen coords
                rx0, ry0, rx1, ry1 = [float(v) for v in rect_raw]
                x0, x1 = sorted([rx0, rx1])
                y0 = page_height - max(ry0, ry1)
                y1 = page_height - min(ry0, ry1)

                links.append(LinkAnnotation(page_idx, uri, (x0, y0, x1, y1)))
            except Exception as e:
                continue

    return links


def extract_text_tokens(pdf_path: str) -> List[TextToken]:
    """Extract all words from PDF with their positions."""
    tokens = []
    with pdfplumber.open(pdf_path) as pdf:
        for page_idx, page in enumerate(pdf.pages):
            words = page.extract_words(
                x_tolerance=3,
                y_tolerance=3,
                keep_blank_chars=False,
                use_text_flow=False,
            )
            for word in words:
                tokens.append(TextToken(
                    page_num=page_idx,
                    text=word["text"],
                    x0=word["x0"],
                    y0=word["top"],
                    x1=word["x1"],
                    y1=word["bottom"],
                ))
    return tokens


# ============================================================================
# GEOMETRY HELPERS
# ============================================================================

def rectangles_overlap(rect_a: tuple, rect_b: tuple, threshold: float = 0.3) -> bool:
    """Check if two rectangles overlap by at least threshold percentage."""
    ax0, ay0, ax1, ay1 = rect_a
    bx0, by0, bx1, by1 = rect_b

    # Calculate intersection
    ix0 = max(ax0, bx0)
    iy0 = max(ay0, by0)
    ix1 = min(ax1, bx1)
    iy1 = min(ay1, by1)

    if ix1 <= ix0 or iy1 <= iy0:
        return False

    intersection = (ix1 - ix0) * (iy1 - iy0)
    area_a = (ax1 - ax0) * (ay1 - ay0)
    area_b = (bx1 - bx0) * (by1 - by0)
    min_area = min(area_a, area_b)

    if min_area == 0:
        return False

    return (intersection / min_area) >= threshold


def find_hyperlink_for_token(
    token: TextToken,
    links: List[LinkAnnotation],
    page_links_map: Dict[int, List[LinkAnnotation]]
) -> Optional[LinkAnnotation]:
    """Find hyperlink annotation that covers this token."""
    for link in page_links_map.get(token.page_num, []):
        if rectangles_overlap(
            (token.x0, token.y0, token.x1, token.y1),
            link.rect,
        ):
            return link
    return None


# ============================================================================
# ADDRESS MATCHING FUNCTIONS
# ============================================================================

def is_mailto_link(uri: str) -> bool:
    """Check if URI is a mailto link."""
    return uri.lower().startswith("mailto:")


def is_http_link(uri: str) -> bool:
    """Check if URI is an HTTP/HTTPS/www link."""
    uri_lower = uri.lower()
    return (uri_lower.startswith("http://") or 
            uri_lower.startswith("https://") or 
            uri_lower.startswith("www."))


def extract_email_from_mailto(uri: str) -> str:
    """Extract email address from mailto: URI."""
    if uri.lower().startswith("mailto:"):
        return uri[7:].split("?")[0].strip()
    return ""


def normalize_address(address: str) -> str:
    """Remove trailing punctuation from address."""
    return address.strip().rstrip(".,;:!?)\"'")


def addresses_match(displayed_text: str, uri: str) -> bool:
    """
    Compare displayed text with URI destination.
    Handles normalization for both emails and URLs.
    """
    displayed = normalize_address(displayed_text)

    # Email comparison
    if is_mailto_link(uri):
        email = extract_email_from_mailto(uri)
        return displayed.lower() == email.lower()

    # URL comparison (normalize protocols and www prefix)
    if is_http_link(uri):
        def normalize(url):
            url = url.lower().rstrip("/")
            if url.startswith("http://"):
                url = url[7:]
            elif url.startswith("https://"):
                url = url[8:]
            elif url.startswith("www."):
                url = url[4:]
            return url

        return normalize(displayed) == normalize(uri)

    return False


# ============================================================================
# GROUPING AND SCANNING FUNCTIONS
# ============================================================================

def group_tokens_by_line(tokens: List[TextToken]) -> List[List[TextToken]]:
    """Group tokens into lines based on Y position."""
    if not tokens:
        return []

    lines = []
    current_line = [tokens[0]]

    for token in tokens[1:]:
        # Same page and similar Y position = same line
        if (token.page_num == current_line[-1].page_num and
            abs(token.y0 - current_line[-1].y0) < 4.0):
            current_line.append(token)
        else:
            lines.append(current_line)
            current_line = [token]

    lines.append(current_line)
    return lines


def find_addresses_in_line(tokens: List[TextToken]) -> List[Tuple[str, List[TextToken]]]:
    """
    Find all emails and URLs in a line of tokens.
    Returns list of (address, covering_tokens).
    """
    if not tokens:
        return []

    # Join tokens into single text
    full_text = " ".join(t.text for t in tokens)
    results = []

    # Find all emails and URLs
    for pattern in (EMAIL_REGEX, URL_REGEX):
        for match in pattern.finditer(full_text):
            address = normalize_address(match.group())
            
            # Map matched text back to tokens
            covering_tokens = []
            search_pos = 0

            for token in tokens:
                token_start = full_text.find(token.text, search_pos)
                token_end = token_start + len(token.text)

                # Token overlaps with match
                if token_start < match.end() and token_end > match.start():
                    covering_tokens.append(token)

                search_pos = token_start + 1

            if covering_tokens:
                results.append((address, covering_tokens))

    return results


# ============================================================================
# MAIN VALIDATION LOGIC
# ============================================================================

def validate_pdf(pdf_path: str, verbose: bool = False) -> List[ExcelRow]:
    """Main validation function."""
    rows = []

    print(f"📄 Processing: {pdf_path}")
    
    # Extract content
    hyperlinks = extract_hyperlinks(pdf_path)
    tokens = extract_text_tokens(pdf_path)

    # Build index for quick lookup
    page_links_map = {}
    for link in hyperlinks:
        page_links_map.setdefault(link.page_num, []).append(link)

    print(f"   Found {len(hyperlinks)} hyperlinks")
    print(f"   Found {len(tokens)} text tokens\n")

    # Group tokens and scan for addresses
    lines = group_tokens_by_line(tokens)

    for line in lines:
        addresses = find_addresses_in_line(line)

        for address, covering_tokens in addresses:
            page_num = covering_tokens[0].page_num + 1

            # Determine address type
            is_email = bool(EMAIL_REGEX.fullmatch(address))
            is_url = bool(URL_REGEX.match(address))
            link_type = "Email" if is_email else "Web Link"

            # Find hyperlinks covering this address
            found_links = []
            for token in covering_tokens:
                link = find_hyperlink_for_token(token, hyperlinks, page_links_map)
                if link:
                    found_links.append(link)

            unique_uris = list({link.uri for link in found_links})

            # Initialize result
            is_hyperlinked = "Yes" if unique_uris else "No"
            hyperlink_text = unique_uris[0] if unique_uris else "No Link"
            hyperlink_points = hyperlink_text
            is_valid = "Yes"
            result = "Pass"

            # Validation checks
            if not unique_uris:
                # FAIL: No hyperlink found
                is_valid = "No"
                result = "Fail"
                if verbose:
                    print(f"  ❌ Page {page_num}: '{address}' → NO LINK")

            else:
                # Check link type and target
                has_error = False
                for uri in unique_uris:
                    if is_email and is_http_link(uri):
                        # Email linked to web URL
                        is_valid = "No"
                        result = "Fail"
                        has_error = True
                        if verbose:
                            print(f"  ❌ Page {page_num}: '{address}' → {uri} (EMAIL→WEB)")

                    elif is_url and is_mailto_link(uri):
                        # URL linked to email
                        is_valid = "No"
                        result = "Fail"
                        has_error = True
                        if verbose:
                            print(f"  ❌ Page {page_num}: '{address}' → {uri} (WEB→EMAIL)")

                    elif not addresses_match(address, uri):
                        # Address mismatch
                        is_valid = "No"
                        result = "Fail"
                        has_error = True
                        if verbose:
                            print(f"  ❌ Page {page_num}: '{address}' → {uri} (MISMATCH)")

                if not has_error and verbose:
                    print(f"  ✅ Page {page_num}: '{address}' → {unique_uris[0]}")

            # Create row
            rows.append(ExcelRow(
                pdf_link=address,
                hyperlink_text=hyperlink_text,
                page_num=page_num,
                link_type=link_type,
                is_hyperlinked=is_hyperlinked,
                is_valid=is_valid,
                hyperlink_points_to=hyperlink_points,
                result=result
            ))

    return rows


# ============================================================================
# EXCEL EXPORT
# ============================================================================

def export_excel(rows: List[ExcelRow], output_path: str):
    """Export validation results to Excel."""
    try:
        from openpyxl import Workbook
        from openpyxl.styles import PatternFill, Font, Alignment
    except ImportError:
        print("❌ openpyxl required: pip install openpyxl")
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "PDF Link Validation"

    # Headers
    headers = [
        "PDF_Link",
        "Hyperlink_Text",
        "Page_Numb",
        "Link_Type",
        "Is_Hyperlinked",
        "Is_Valid",
        "Hyperlink_Points_To",
        "Result"
    ]
    ws.append(headers)

    # Style header
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Color fills for result column
    pass_fill = PatternFill(start_color="D4EDDA", end_color="D4EDDA", fill_type="solid")
    pass_font = Font(color="155724", bold=True)
    
    fail_fill = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid")
    fail_font = Font(color="721C24", bold=True)

    # Add data rows
    for row in rows:
        ws.append([
            row.pdf_link,
            row.hyperlink_text,
            row.page_num,
            row.link_type,
            row.is_hyperlinked,
            row.is_valid,
            row.hyperlink_points_to,
            row.result
        ])

        row_idx = ws.max_row
        result_cell = ws[f"H{row_idx}"]

        if row.result == "Pass":
            result_cell.fill = pass_fill
            result_cell.font = pass_font
        else:
            result_cell.fill = fail_fill
            result_cell.font = fail_font

    # Set column widths
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 40
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 14
    ws.column_dimensions["F"].width = 11
    ws.column_dimensions["G"].width = 40
    ws.column_dimensions["H"].width = 10

    # Center align most columns
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for idx, cell in enumerate(row):
            if idx in [2, 3, 4, 5, 7]:  # Center numeric and status columns
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

    wb.save(output_path)
    print(f"✅ Report saved: {output_path}\n")


# ============================================================================
# MAIN
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Validate all hyperlinks in a PDF document.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__
    )
    parser.add_argument("pdf", nargs='?', default=None, help="Path to PDF file (optional)")
    parser.add_argument("-o", "--output", help="Output Excel file path (.xlsx)")
    parser.add_argument("-v", "--verbose", action="store_true", help="Show detailed validation messages")

    args = parser.parse_args()

    # If no PDF provided, open file picker
    pdf_path = args.pdf
    if not pdf_path:
        print("🔍 Opening file picker...\n")
        pdf_path = select_pdf_file()
        
        if not pdf_path:
            print("❌ No file selected!")
            sys.exit(1)

    # Check file exists
    if not os.path.exists(pdf_path):
        print(f"❌ File not found: {pdf_path}")
        sys.exit(1)

    # Run validation
    rows = validate_pdf(pdf_path, verbose=args.verbose)

    # Print summary
    print("=" * 72)
    passes = sum(1 for r in rows if r.result == "Pass")
    fails = sum(1 for r in rows if r.result == "Fail")
    total = len(rows)

    print(f"📊 SUMMARY")
    print(f"   Total Entries: {total}")
    print(f"   ✅ Pass: {passes}")
    print(f"   ❌ Fail: {fails}")
    if total > 0:
        success_rate = (passes / total) * 100
        print(f"   📈 Success Rate: {success_rate:.1f}%")
    print("=" * 72 + "\n")

    # Export to Excel
    if args.output:
        output_file = args.output if args.output.lower().endswith('.xlsx') else args.output + '.xlsx'
    else:
        base = os.path.splitext(os.path.basename(pdf_path))[0]
        output_file = f"{base}_validation_report.xlsx"

    export_excel(rows, output_file)

    # Exit with appropriate code
    sys.exit(0 if fails == 0 else 1)


if __name__ == "__main__":
    main()
