"""
Lampiran 7C ASPI Automation Script
===================================
Script untuk mengotomasi pembuatan dokumen Lampiran 7C ASPI (format Word)
dari sumber data UAT Script (format Excel).

Fungsi utama:
- Membaca file UAT Script.xlsx dari folder input/
- Parsing kolom Remarks untuk mengambil URL, Headers, Request Body, dan Response
- Membuat dokumen Word Lampiran 7C dari scratch dengan format yang sesuai
- Menyimpan hasil ke folder output/
"""

import os
import re
import sys
from pathlib import Path

from openpyxl import load_workbook
from docx import Document
from docx.shared import Inches, Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn


# ============================================================
# KONFIGURASI & KONSTANTA
# ============================================================

# Path konfigurasi
INPUT_DIR = Path("input")
OUTPUT_DIR = Path("output")
UAT_SCRIPT_FILENAME = "UAT Script.xlsx"
OUTPUT_FILENAME = "Lampiran 7C - Hasil UAT.docx"

# Sheet name di file UAT Script
UAT_SHEET_NAME = "UAT Script"

# Kolom index di UAT Script (0-based, dengan kolom A kosong)
# Kolom A (index 0): KOSONG
# Kolom B (index 1): Kategori Tes
# Kolom C (index 2): Nama Modul
# Kolom D (index 3): Nomor Skenario
# Kolom E (index 4): Nomor Kasus Tes (contoh: 1.1, 2.3, 7.15)
# Kolom F (index 5): Langkah Tes
# Kolom G (index 6): Hasil yang diharapkan
# Kolom H (index 7): Hasil Aktual
# Kolom I (index 8): Remarks (berisi log URL, Headers, Request Body, Response)
# Kolom J (index 9): Tanggal Pelaksanaan
# Kolom K (index 10): Jenis Script
# Kolom L (index 11): Pelaksana

COL_KATEGORI_TES = 1
COL_NAMA_MODUL = 2
COL_NOMOR_SKENARIO = 3
COL_NOMOR_KASUS_TES = 4
COL_LANGKAH_TES = 5
COL_HASIL_DIHARAPKAN = 6
COL_HASIL_AKTUAL = 7
COL_REMARKS = 8
COL_TANGGAL = 9
COL_JENIS_SCRIPT = 10
COL_PELAKSANA = 11

# Definisi 6 section API di Lampiran 7C
# Format: (nama_lampiran, jumlah_skenario)
LAMPIRAN_SECTIONS = [
    ("API Balance Inquiry", 11),
    ("Intrabank Transfer", 12),
    ("Interbank Transfer", 13),
    ("API RTGS Transfer", 13),
    ("API SKNBI Transfer", 13),
    ("API Virtual Account", 33),
]

# Mapping dari section header di UAT Script ke section di Lampiran 7C
# Format: (header_uat, skenario_prefix, target_lampiran_section, is_fill_empty_only)
UAT_TO_LAMPIRAN_MAPPING = [
    ("Balance Services", "1", "API Balance Inquiry", False),
    ("Intrabank Transfer", "2", "Intrabank Transfer", False),
    ("Interbank Transfer", "3", "Interbank Transfer", False),
    ("Interbank Transfer via BI FAST", "4", "Interbank Transfer", True),
    ("RTGS Transfer", "5", "API RTGS Transfer", False),
    ("SKNBI Transfer", "6", "API SKNBI Transfer", False),
    ("Transfer VA", "7", "API Virtual Account", False),
    ("Transfer VA Prima", "8", "API Virtual Account", True),
    ("Transfer VA BI FAST", "9", "API Virtual Account", True),
]

# Skenario yang menggunakan format "BSS YANG HIT"
BSS_HIT_SCENARIOS = ["7", "8", "9"]

# Skenario yang menggunakan format "Mitra HIT"
MITRA_HIT_SCENARIOS = ["2", "3", "4", "5", "6"]


# ============================================================
# PARSER KOLOM REMARKS
# ============================================================

# Header-like keys yang umum ditemukan di HTTP headers
_HEADER_KEYS = [
    "x-timestamp", "x-partner-id", "x-signature", "authorization",
    "x-external-id", "channel-id", "content-type", "x-forwarded-for",
    "x-forwarded-host", "x-forwarded-proto", "user-agent", "accept",
    "host", "connection", "cache-control", "pragma", "x-request-id",
    "x-api-key", "x-client-key", "x-client-secret",
]


def _is_header_like_json(text):
    """
    Check if a JSON block looks like HTTP headers (contains header-like keys).
    Returns True if the JSON contains keys commonly found in HTTP headers.
    """
    text_lower = text.lower()
    header_indicators = 0
    for key in _HEADER_KEYS:
        if key in text_lower:
            header_indicators += 1
    # If at least 2 header-like keys found, it's likely headers
    return header_indicators >= 2


def _is_response_like_json(text):
    """
    Check if a JSON block looks like an API response.
    """
    text_lower = text.lower()
    return "responsecode" in text_lower or "responsemessage" in text_lower or \
           '"responsecode"' in text_lower or '"responsemessage"' in text_lower


def _is_header_key_value_line(line):
    """
    Check if a line looks like a header in Key: Value format (not JSON).
    E.g., "X-TIMESTAMP: 2025-07-30T14:28:56+07:00"
    """
    stripped = line.strip()
    if not stripped or stripped.startswith('{') or stripped.startswith('['):
        return False
    # Must have a colon with a key part that looks like a header name
    if ':' not in stripped:
        return False
    key_part = stripped.split(':', 1)[0].strip()
    # Header names are typically alphanumeric with hyphens, no spaces
    if not key_part:
        return False
    # Check if it matches common header pattern (letters, digits, hyphens)
    if re.match(r'^[A-Za-z][A-Za-z0-9\-]*$', key_part):
        key_lower = key_part.lower()
        for hk in _HEADER_KEYS:
            if key_lower == hk:
                return True
        # Even if not in the known list, if it looks like a header name
        # (starts with X- or common patterns)
        if key_lower.startswith('x-') or key_lower in [
            'authorization', 'content-type', 'accept', 'host',
            'connection', 'pragma', 'cookie', 'referer', 'origin'
        ]:
            return True
    return False


def _extract_url_from_text(text):
    """
    Extract URL from text by finding lines containing http:// or https://.
    Strips any label prefix like "URL:", "- url:", etc.
    Returns the first URL found, or empty string.
    """
    for line in text.split('\n'):
        stripped = line.strip()
        # Find http:// or https:// in the line
        match = re.search(r'(https?://\S+)', stripped)
        if match:
            return match.group(1).rstrip(',').rstrip('"').rstrip("'")
    return ""


def _extract_json_blocks(text):
    """
    Extract all JSON blocks from text (objects starting with { and ending with }).
    Returns list of tuples: (start_pos, end_pos, json_text)
    Handles nested braces.
    """
    blocks = []
    i = 0
    while i < len(text):
        if text[i] == '{':
            depth = 0
            start = i
            while i < len(text):
                if text[i] == '{':
                    depth += 1
                elif text[i] == '}':
                    depth -= 1
                    if depth == 0:
                        block_text = text[start:i + 1]
                        blocks.append((start, i + 1, block_text))
                        break
                i += 1
        i += 1
    return blocks


def _parse_with_labels(text):
    """
    Try to parse text using label/keyword detection.
    Supports both Format A (Mitra HIT) and Format B (BSS HIT) labels.
    Returns dict or None if no labels found.
    """
    result = {
        "url": "",
        "headers": "",
        "request_body": "",
        "response": ""
    }

    # Combined patterns for both formats:
    # Format A: "URL:", "Header:", "Headers:", "RequestBody:", "Request Body:", "Response:"
    # Format B: "- url:", "- headers:", "- Request Body:", "- body:", "- Response:"
    url_pattern = r'(?i)(?:^|\n)\s*[-*]?\s*url\s*:\s*'
    header_pattern = r'(?i)(?:^|\n)\s*[-*]?\s*headers?\s*:\s*'
    request_body_pattern = r'(?i)(?:^|\n)\s*[-*]?\s*(?:Request\s*Body|RequestBody|body)\s*:\s*'
    response_pattern = r'(?i)(?:^|\n)\s*[-*]?\s*Response\s*:\s*'

    url_match = re.search(url_pattern, text)
    header_match = re.search(header_pattern, text)
    request_body_match = re.search(request_body_pattern, text)
    response_match = re.search(response_pattern, text)

    sections = []
    if url_match:
        sections.append(("url", url_match.end(), url_match.start()))
    if header_match:
        sections.append(("headers", header_match.end(), header_match.start()))
    if request_body_match:
        sections.append(("request_body", request_body_match.end(), request_body_match.start()))
    if response_match:
        sections.append(("response", response_match.end(), response_match.start()))

    if not sections:
        return None

    # If no URL label found, check if text starts with a URL (before any label)
    if not url_match:
        first_label_start = min(s[2] for s in sections)
        # Check lines before the first label for a URL
        text_before_labels = text[:first_label_start]
        url_found = _extract_url_from_text(text_before_labels)
        if url_found:
            result["url"] = url_found

    # Sort by position
    sections.sort(key=lambda x: x[1])

    # Extract content for each section
    for i, (section_name, start_pos, label_start) in enumerate(sections):
        if i + 1 < len(sections):
            next_label_start = sections[i + 1][2]
            content = text[start_pos:next_label_start].strip()
        else:
            content = text[start_pos:].strip()
        result[section_name] = content

    # For URL field, extract just the URL if there's extra text
    if result["url"]:
        url_extracted = _extract_url_from_text(result["url"])
        if url_extracted:
            result["url"] = url_extracted
        else:
            # Maybe the URL is on the same line after the label
            first_line = result["url"].split('\n')[0].strip()
            if first_line.startswith("http://") or first_line.startswith("https://"):
                result["url"] = first_line

    return result


def _parse_by_content_detection(text):
    """
    Parse text by detecting content patterns (URLs, headers, JSON blocks)
    without relying on labels/keywords.

    Strategy:
    1. Find URL (line with http:// or https://)
    2. Find all JSON blocks
    3. Classify JSON blocks as headers vs request body vs response
    4. Find header key:value lines between URL and first non-header JSON block
    """
    result = {
        "url": "",
        "headers": "",
        "request_body": "",
        "response": ""
    }

    # Step 1: Extract URL
    url = _extract_url_from_text(text)
    result["url"] = url

    # Step 2: Find all JSON blocks
    json_blocks = _extract_json_blocks(text)

    if not json_blocks:
        # No JSON blocks - try to find headers as key:value lines
        lines = text.split('\n')
        header_lines = []
        for line in lines:
            stripped = line.strip()
            # Skip URL line
            if 'http://' in stripped or 'https://' in stripped:
                continue
            # Skip empty lines and label lines
            if not stripped:
                continue
            if _is_header_key_value_line(line):
                header_lines.append(stripped)
        if header_lines:
            result["headers"] = '\n'.join(header_lines)
        return result

    # Step 3: Classify JSON blocks
    header_json_blocks = []
    body_json_blocks = []
    response_json_blocks = []

    for start, end, block_text in json_blocks:
        if _is_header_like_json(block_text):
            header_json_blocks.append((start, end, block_text))
        elif _is_response_like_json(block_text):
            response_json_blocks.append((start, end, block_text))
        else:
            body_json_blocks.append((start, end, block_text))

    # If we have unclassified blocks, use position-based heuristic:
    # - First unclassified after headers = request body
    # - Last block in text = response (if not already classified)
    if not response_json_blocks and body_json_blocks and len(body_json_blocks) >= 2:
        # Last body block is likely the response
        response_json_blocks.append(body_json_blocks.pop())
    elif not response_json_blocks and body_json_blocks and len(body_json_blocks) == 1:
        # Single unclassified JSON - if no header blocks either, it might be a response
        # Check position: if it's the last thing in text, treat as response
        if not header_json_blocks:
            response_json_blocks.append(body_json_blocks.pop())

    # Assign results
    if header_json_blocks:
        result["headers"] = header_json_blocks[0][2]
    if body_json_blocks:
        result["request_body"] = body_json_blocks[0][2]
    if response_json_blocks:
        result["response"] = response_json_blocks[-1][2]

    # Step 4: If no JSON headers found, look for key:value header lines
    if not header_json_blocks:
        lines = text.split('\n')
        header_lines = []
        url_line_found = False
        # Find headers between URL and first JSON block
        first_json_start = json_blocks[0][0] if json_blocks else len(text)
        for line in lines:
            stripped = line.strip()
            if not stripped:
                continue
            # Get the approximate position of this line in text
            line_start = text.find(stripped)
            if line_start >= first_json_start:
                break
            if 'http://' in stripped or 'https://' in stripped:
                url_line_found = True
                continue
            if _is_header_key_value_line(line):
                header_lines.append(stripped)
        if header_lines:
            result["headers"] = '\n'.join(header_lines)

    return result


def parse_remarks_mitra_hit(remarks_text):
    """
    Parse format Remarks dari 'Mitra HIT' (Skenario 2, 3, 4, 5, 6).

    Uses a content-detection approach:
    1. First tries label-based parsing (URL:, Header:, RequestBody:, Response:)
    2. Falls back to content detection (finds URLs, headers, JSON blocks by pattern)

    Supported formats:
        FORMAT A (with labels):
            URL:
            https://...

            Header:
            X-TIMESTAMP: ...
            Authorization: Bearer ...

            RequestBody:
            {"key": "value"}

            Response:
            {"responseCode": "..."}

        FORMAT (no labels or partial labels):
            Detects URL by http/https pattern
            Detects headers by Key: Value pattern
            Detects request body and response by JSON content analysis

    Returns:
        dict with keys: url, headers, request_body, response
    """
    if not remarks_text or not remarks_text.strip():
        return None

    text = remarks_text.strip()
    text = text.replace('\r\n', '\n').replace('\r', '\n')

    # Strategy 1: Try label-based parsing first
    result = _parse_with_labels(text)
    if result and any(result.values()):
        # Validate: if we got at least a URL or some meaningful content, return it
        if result["url"] or result["headers"] or result["request_body"] or result["response"]:
            return result

    # Strategy 2: Content-based detection (fallback)
    result = _parse_by_content_detection(text)
    if result and any(result.values()):
        return result

    # Strategy 3: If text starts with URL directly
    if text.startswith("http://") or text.startswith("https://"):
        lines = text.split('\n')
        fallback_result = {
            "url": lines[0].strip(),
            "headers": "",
            "request_body": "",
            "response": ""
        }
        remaining = '\n'.join(lines[1:]).strip()
        if remaining:
            fallback_result["response"] = remaining
        return fallback_result

    return None


def parse_remarks_bss_hit(remarks_text):
    """
    Parse format Remarks dari 'BSS YANG HIT' (Skenario 7, 8, 9).

    Uses a content-detection approach:
    1. Handles multiple sections (#validate, #commit, etc.)
    2. First tries label-based parsing (- url:, - headers:, - Request Body:, - Response:)
    3. Falls back to content detection (finds URLs, headers, JSON blocks by pattern)

    Supported formats:
        FORMAT B (with bullet markers):
            - url: https://...
            - headers: {"key": "value", ...}
            - Request Body: {"key": "value"}
            - Response: {"responseCode": "..."}

        FORMAT (no labels):
            Detects URL by http/https pattern
            Detects headers by JSON with header-like keys
            Detects request body and response by JSON content analysis

    Returns:
        dict with keys: url, headers, request_body, response
    """
    if not remarks_text or not remarks_text.strip():
        return None

    text = remarks_text.strip()
    text = text.replace('\r\n', '\n').replace('\r', '\n')

    # Cek apakah ada multiple sections (#validate, #commit, #1, #2, dll)
    section_marker_pattern = r'(?:^|\n)\s*#\s*\w+'
    section_markers = list(re.finditer(section_marker_pattern, text))

    if len(section_markers) >= 2:
        # Multiple sections - gabungkan semua
        return _parse_bss_multi_section(text, section_markers)

    # Single section (mungkin ada satu # marker di awal yang di-skip)
    # Hapus baris yang dimulai dengan # (section markers)
    lines = text.split('\n')
    cleaned_lines = []
    for line in lines:
        if re.match(r'^\s*#', line):
            continue
        cleaned_lines.append(line)
    cleaned_text = '\n'.join(cleaned_lines).strip()

    if not cleaned_text:
        return None

    return _parse_bss_single_section(cleaned_text)


def _parse_bss_single_section(text):
    """
    Parse a single BSS HIT section.
    Uses label-based parsing first, then falls back to content detection.
    """
    # Strategy 1: Try label-based parsing
    result = _parse_with_labels(text)
    if result and any(result.values()):
        if result["url"] or result["headers"] or result["request_body"] or result["response"]:
            return result

    # Strategy 2: Content-based detection
    result = _parse_by_content_detection(text)
    if result and any(result.values()):
        return result

    return None


def _parse_bss_multi_section(text, section_markers):
    """
    Parse multiple BSS HIT sections (e.g., #validate and #commit).
    Gabungkan semua sections menjadi satu output dengan separator.
    """
    # Split text into sections based on markers
    section_texts = []
    section_names = []

    for i, marker in enumerate(section_markers):
        # Extract section name from marker
        marker_text = text[marker.start():marker.end()].strip()
        # Remove the # prefix
        section_name = marker_text.lstrip('#').strip()
        section_names.append(section_name)

        # Get content between this marker and the next
        start = marker.end()
        if i + 1 < len(section_markers):
            end = section_markers[i + 1].start()
        else:
            end = len(text)

        section_content = text[start:end].strip()
        section_texts.append(section_content)

    # Parse each section individually
    parsed_sections = []
    valid_section_names = []
    for idx, section_content in enumerate(section_texts):
        # Remove any # lines within the section content
        lines = section_content.split('\n')
        cleaned_lines = [l for l in lines if not re.match(r'^\s*#', l)]
        cleaned = '\n'.join(cleaned_lines).strip()
        if cleaned:
            parsed = _parse_bss_single_section(cleaned)
            if parsed and any(parsed.values()):
                parsed_sections.append(parsed)
                valid_section_names.append(section_names[idx])

    if not parsed_sections:
        return None

    if len(parsed_sections) == 1:
        return parsed_sections[0]

    # Gabungkan multiple parsed sections with separators
    result = {
        "url": "",
        "headers": "",
        "request_body": "",
        "response": ""
    }

    # Collect values per field, paired with their section names
    urls = [(valid_section_names[i], p.get("url", ""))
            for i, p in enumerate(parsed_sections) if p.get("url")]
    headers_list = [(valid_section_names[i], p.get("headers", ""))
                    for i, p in enumerate(parsed_sections) if p.get("headers")]
    bodies = [(valid_section_names[i], p.get("request_body", ""))
              for i, p in enumerate(parsed_sections) if p.get("request_body")]
    responses = [(valid_section_names[i], p.get("response", ""))
                 for i, p in enumerate(parsed_sections) if p.get("response")]

    def combine_field(items):
        """Combine field values with section separators."""
        if not items:
            return ""
        if len(items) == 1:
            return items[0][1]
        return "\n\n".join(
            f"--- {name.upper()} ---\n{value}"
            for name, value in items
        )

    result["url"] = combine_field(urls)
    result["headers"] = combine_field(headers_list)
    result["request_body"] = combine_field(bodies)
    result["response"] = combine_field(responses)

    return result


def parse_remarks(remarks_text, scenario_prefix):
    """
    Parse kolom Remarks berdasarkan format yang sesuai dengan skenario.

    Uses content-detection approach. The scenario_prefix determines which
    parser to try first, but both parsers use the same underlying content
    detection logic.

    Args:
        remarks_text: Teks dari kolom Remarks
        scenario_prefix: Prefix skenario (e.g., "2", "3", "7")

    Returns:
        dict with keys: url, headers, request_body, response
        atau None jika tidak bisa di-parse
    """
    if not remarks_text or not remarks_text.strip():
        return None

    if scenario_prefix in BSS_HIT_SCENARIOS:
        result = parse_remarks_bss_hit(remarks_text)
        if result and any(result.values()):
            return result
        # Fallback to mitra parser
        return parse_remarks_mitra_hit(remarks_text)
    elif scenario_prefix in MITRA_HIT_SCENARIOS:
        result = parse_remarks_mitra_hit(remarks_text)
        if result and any(result.values()):
            return result
        # Fallback to bss parser
        return parse_remarks_bss_hit(remarks_text)
    else:
        # Skenario 1 (Balance) - coba kedua format
        result = parse_remarks_mitra_hit(remarks_text)
        if result and any(result.values()):
            return result
        result = parse_remarks_bss_hit(remarks_text)
        if result and any(result.values()):
            return result
        return None


# ============================================================
# EXCEL READER
# ============================================================

def get_cell_value(row, index):
    """Safely get cell value from a row tuple by index."""
    if index < len(row) and row[index] is not None:
        return str(row[index]).strip()
    return ""


def detect_section_header(row):
    """
    Detect if a row is a section header.

    Section header rules:
    - Kolom E (Nomor Kasus Tes, index 4) HARUS KOSONG
    - Kolom F (Langkah Tes, index 5) HARUS KOSONG
    - Teks section muncul di kolom B (index 1) atau kolom C (index 2)

    Returns:
        str: Detected section name, or None if not a section header
    """
    # Kolom E (Nomor Kasus Tes) harus kosong
    nomor_kasus = get_cell_value(row, COL_NOMOR_KASUS_TES)
    if nomor_kasus:
        return None

    # Kolom F (Langkah Tes) harus kosong
    langkah_tes = get_cell_value(row, COL_LANGKAH_TES)
    if langkah_tes:
        return None

    # Cek teks di kolom B atau C
    kategori_tes = get_cell_value(row, COL_KATEGORI_TES)
    nama_modul = get_cell_value(row, COL_NAMA_MODUL)

    # Gabungkan kedua kolom untuk pengecekan
    text_to_check = kategori_tes + " " + nama_modul

    # Section keywords - ordered from most specific to least specific
    # to ensure exact matching (e.g., "Transfer VA Prima" before "Transfer VA")
    section_keywords = [
        "Interbank Transfer via BI FAST",
        "Transfer VA Prima",
        "Transfer VA BI FAST",
        "Balance Services",
        "Intrabank Transfer",
        "Interbank Transfer",
        "RTGS Transfer",
        "SKNBI Transfer",
        "Transfer VA",
    ]

    for keyword in section_keywords:
        # Check in both kolom B and kolom C
        if keyword.lower() in kategori_tes.lower() or keyword.lower() in nama_modul.lower():
            # For "Transfer VA" (without Prima/BI FAST), we need exact matching
            # to avoid matching "Transfer VA Prima" or "Transfer VA BI FAST"
            if keyword == "Transfer VA":
                # Make sure neither kolom B nor kolom C contains "Prima" or "BI FAST"
                combined_lower = text_to_check.lower()
                if "prima" in combined_lower or "bi fast" in combined_lower:
                    continue
            elif keyword == "Interbank Transfer":
                # Make sure it's not "Interbank Transfer via BI FAST"
                combined_lower = text_to_check.lower()
                if "bi fast" in combined_lower or "via bi" in combined_lower:
                    continue
            return keyword

    return None


def read_uat_script(filepath):
    """
    Membaca file UAT Script.xlsx dan mengembalikan data terstruktur.

    Kolom di UAT Script dimulai dari kolom B (kolom A kosong):
    - B (1): Kategori Tes
    - C (2): Nama Modul
    - D (3): Nomor Skenario
    - E (4): Nomor Kasus Tes
    - F (5): Langkah Tes
    - G (6): Hasil yang diharapkan
    - H (7): Hasil Aktual
    - I (8): Remarks
    - J (9): Tanggal Pelaksanaan
    - K (10): Jenis Script
    - L (11): Pelaksana

    Returns:
        dict: {
            section_header: [
                {
                    'nama_modul': str,
                    'nomor_skenario': str,
                    'nomor_kasus_tes': str,
                    'langkah_tes': str,
                    'hasil_diharapkan': str,
                    'hasil_aktual': str,
                    'remarks': str,
                }
            ]
        }
    """
    wb = load_workbook(filepath, read_only=True, data_only=True)

    if UAT_SHEET_NAME not in wb.sheetnames:
        # Coba cari sheet dengan nama yang mirip
        for sheet_name in wb.sheetnames:
            if "uat" in sheet_name.lower() or "script" in sheet_name.lower():
                ws = wb[sheet_name]
                break
        else:
            ws = wb.active
            print(f"  [WARNING] Sheet '{UAT_SHEET_NAME}' tidak ditemukan. Menggunakan sheet: {ws.title}")
    else:
        ws = wb[UAT_SHEET_NAME]

    data = {}
    current_section = None
    rows_list = list(ws.iter_rows(min_row=1, values_only=True))

    # Identifikasi header row (baris dengan label kolom)
    # Cek kolom B (index 1) untuk label "Kategori Tes" atau similar
    header_row_idx = 0
    for idx, row in enumerate(rows_list):
        if not row or len(row) <= COL_KATEGORI_TES:
            continue
        cell_val = get_cell_value(row, COL_KATEGORI_TES)
        if cell_val.lower() in ['kategori tes', 'kategori', 'no']:
            header_row_idx = idx
            break

    # Process data rows (after header)
    for idx in range(header_row_idx + 1, len(rows_list)):
        row = rows_list[idx]
        if not row or all(cell is None or str(cell).strip() == '' for cell in row):
            continue

        # Cek apakah ini section header
        detected_section = detect_section_header(row)
        if detected_section:
            current_section = detected_section
            if current_section not in data:
                data[current_section] = []
            continue

        # Jika belum ada section, skip
        if current_section is None:
            continue

        # Parse data row - harus punya Nomor Kasus Tes (kolom E, index 4)
        nomor_kasus = get_cell_value(row, COL_NOMOR_KASUS_TES)
        if not nomor_kasus:
            continue

        row_data = {
            'nama_modul': get_cell_value(row, COL_NAMA_MODUL),
            'nomor_skenario': get_cell_value(row, COL_NOMOR_SKENARIO),
            'nomor_kasus_tes': nomor_kasus,
            'langkah_tes': get_cell_value(row, COL_LANGKAH_TES),
            'hasil_diharapkan': get_cell_value(row, COL_HASIL_DIHARAPKAN),
            'hasil_aktual': get_cell_value(row, COL_HASIL_AKTUAL),
            'remarks': get_cell_value(row, COL_REMARKS),
        }

        data[current_section].append(row_data)

    wb.close()
    return data


# ============================================================
# MAPPING LOGIC
# ============================================================

def extract_sub_number(nomor_kasus_tes):
    """
    Extract sub-number dari Nomor Kasus Tes.
    Contoh: "2.3" -> 3, "7.15" -> 15, "1.1" -> 1

    Args:
        nomor_kasus_tes: String nomor kasus (e.g., "2.3")

    Returns:
        int: Sub-number (1-based index untuk baris di Lampiran 7C)
    """
    parts = nomor_kasus_tes.split('.')
    if len(parts) >= 2:
        try:
            return int(parts[1])
        except ValueError:
            return None
    return None


def map_result_value(hasil_aktual):
    """
    Map Hasil Aktual ke Result value di Lampiran 7C.

    - "Berhasil" -> "PASS"
    - "Tidak dites" -> "N/A"
    - "Gagal" -> "NOT PASS"
    - Lainnya -> nilai asli
    """
    if not hasil_aktual:
        return ""
    lower = hasil_aktual.lower().strip()
    if lower == "berhasil":
        return "PASS"
    elif lower == "tidak dites":
        return "N/A"
    elif lower == "gagal":
        return "NOT PASS"
    return hasil_aktual


def map_uat_to_lampiran(uat_data):
    """
    Map data dari UAT Script ke struktur Lampiran 7C.

    Mengimplementasikan logika "fill empty only" untuk section yang di-share.

    Args:
        uat_data: Dict hasil dari read_uat_script()

    Returns:
        dict: {
            section_name: [
                {
                    'no': int,
                    'service': str,
                    'scenario': str,
                    'expected_result': str,
                    'request': str,
                    'response': str,
                    'result': str,
                    'notes': str,
                }
                atau None (jika row belum terisi)
            ]
        }
    """
    # Inisialisasi struktur Lampiran 7C dengan None untuk setiap row
    lampiran_data = {}
    for section_name, count in LAMPIRAN_SECTIONS:
        lampiran_data[section_name] = [None] * count

    # Proses mapping berdasarkan urutan prioritas
    for uat_section, scenario_prefix, target_section, fill_empty_only in UAT_TO_LAMPIRAN_MAPPING:
        if uat_section not in uat_data:
            continue

        rows = uat_data[uat_section]
        target_count = None
        for sec_name, sec_count in LAMPIRAN_SECTIONS:
            if sec_name == target_section:
                target_count = sec_count
                break

        if target_count is None:
            continue

        for row_data in rows:
            sub_num = extract_sub_number(row_data['nomor_kasus_tes'])
            if sub_num is None or sub_num < 1 or sub_num > target_count:
                continue

            row_idx = sub_num - 1  # Convert ke 0-based index

            # Cek apakah harus fill empty only
            if fill_empty_only and lampiran_data[target_section][row_idx] is not None:
                continue

            # Parse remarks
            hasil_aktual = row_data['hasil_aktual']
            remarks_text = row_data['remarks']

            request_content = ""
            response_content = ""
            notes = ""
            result_value = map_result_value(hasil_aktual)

            # Service diambil dari kolom Nama Modul
            service = row_data.get('nama_modul', '') or target_section

            # Cek kondisi khusus
            if hasil_aktual.lower().strip() == "tidak dites":
                # Tidak dites: Request dan Response kosong, Notes dari Remarks
                notes = remarks_text if remarks_text and remarks_text.lower() != "none" else "Tidak dites"
                request_content = ""
                response_content = ""
            elif not remarks_text or remarks_text.lower() == "none":
                # Remarks kosong: Request dan Response kosong
                notes = ""
            else:
                # Parse remarks untuk mengambil data
                parsed = parse_remarks(remarks_text, scenario_prefix)
                if parsed:
                    url = parsed.get('url', '')
                    headers = parsed.get('headers', '')
                    request_body = parsed.get('request_body', '')
                    response = parsed.get('response', '')

                    # Format Request: URL + Headers + Request Body
                    request_content = (
                        f"URL Endpoint:\n{url}\n\n"
                        f"Header Request:\n{headers}\n\n"
                        f"Request Body:\n{request_body}"
                    )

                    # Format Response: Response Body saja
                    response_content = f"Response Body:\n{response}"
                else:
                    # Remarks ada tapi tidak bisa di-parse
                    notes = remarks_text

            lampiran_data[target_section][row_idx] = {
                'no': sub_num,
                'service': service,
                'scenario': row_data.get('langkah_tes', '') or f"Skenario {row_data['nomor_kasus_tes']}",
                'expected_result': row_data.get('hasil_diharapkan', ''),
                'request': request_content,
                'response': response_content,
                'result': result_value,
                'notes': notes,
            }

    return lampiran_data


# ============================================================
# WORD DOCUMENT GENERATOR
# ============================================================

def create_lampiran_document(lampiran_data, output_path):
    """
    Membuat dokumen Word Lampiran 7C dari scratch.

    Args:
        lampiran_data: Dict hasil dari map_uat_to_lampiran()
        output_path: Path untuk menyimpan file output
    """
    doc = Document()

    # Set default font
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(9)

    # Buat setiap section
    for section_idx, (section_name, section_count) in enumerate(LAMPIRAN_SECTIONS):
        # Tambah page break sebelum section (kecuali section pertama)
        if section_idx > 0:
            doc.add_page_break()

        # Header section
        heading = doc.add_heading(level=2)
        heading_run = heading.add_run(f"Lampiran 7C - {section_name}")
        heading_run.font.size = Pt(12)
        heading_run.font.bold = True

        # Info tambahan
        info_para = doc.add_paragraph()
        info_para.add_run(f"Nama Layanan API: {section_name}").bold = True
        info_para.space_after = Pt(6)

        # Buat tabel
        # Kolom: No, Service, Scenario, Expected Result, Request, Response, Result, Notes
        table = doc.add_table(rows=1, cols=8)
        table.style = 'Table Grid'
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        # Header row
        header_cells = table.rows[0].cells
        headers = ['No', 'Service', 'Scenario', 'Expected Result',
                   'Request', 'Response', 'Result', 'Notes']
        for i, header_text in enumerate(headers):
            header_cells[i].text = header_text
            # Bold header
            for paragraph in header_cells[i].paragraphs:
                for run in paragraph.runs:
                    run.bold = True
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Data rows
        section_data = lampiran_data.get(section_name, [])
        for row_idx in range(section_count):
            row_cells = table.add_row().cells

            if row_idx < len(section_data) and section_data[row_idx] is not None:
                row_data = section_data[row_idx]
                row_cells[0].text = str(row_data['no'])
                row_cells[1].text = row_data['service']
                row_cells[2].text = row_data['scenario']
                row_cells[3].text = row_data['expected_result']
                row_cells[4].text = row_data['request']
                row_cells[5].text = row_data['response']
                row_cells[6].text = row_data['result']
                row_cells[7].text = row_data['notes']
            else:
                # Row kosong - isi nomor saja
                row_cells[0].text = str(row_idx + 1)
                row_cells[1].text = ""
                row_cells[2].text = ""
                row_cells[3].text = ""
                row_cells[4].text = ""
                row_cells[5].text = ""
                row_cells[6].text = ""
                row_cells[7].text = ""

        # Set column widths (approximate)
        for row in table.rows:
            row.cells[0].width = Cm(1.0)    # No
            row.cells[1].width = Cm(2.5)    # Service
            row.cells[2].width = Cm(3.0)    # Scenario
            row.cells[3].width = Cm(2.5)    # Expected Result
            row.cells[4].width = Cm(5.0)    # Request
            row.cells[5].width = Cm(5.0)    # Response
            row.cells[6].width = Cm(2.0)    # Result
            row.cells[7].width = Cm(2.5)    # Notes

    # Simpan dokumen
    doc.save(str(output_path))
    print(f"  [OK] Dokumen berhasil disimpan: {output_path}")


# ============================================================
# MAIN ENTRY POINT
# ============================================================

def main():
    """Fungsi utama untuk menjalankan automation Lampiran 7C."""
    print("=" * 60)
    print("  AUTOMATION LAMPIRAN 7C ASPI")
    print("  Konversi UAT Script (Excel) -> Lampiran 7C (Word)")
    print("=" * 60)
    print()

    # Cek file input
    uat_script_path = INPUT_DIR / UAT_SCRIPT_FILENAME
    if not uat_script_path.exists():
        print(f"  [ERROR] File UAT Script tidak ditemukan!")
        print(f"  Lokasi yang dicari: {uat_script_path.resolve()}")
        print()
        print("  Langkah yang harus dilakukan:")
        print(f"  1. Copy file UAT Script Anda ke folder: {INPUT_DIR.resolve()}")
        print(f"  2. Pastikan nama file: {UAT_SCRIPT_FILENAME}")
        print("  3. Jalankan ulang script ini")
        print()
        sys.exit(1)

    # Pastikan folder output ada
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    # Baca UAT Script
    print(f"  [1/3] Membaca file UAT Script: {uat_script_path}")
    try:
        uat_data = read_uat_script(uat_script_path)
    except Exception as e:
        print(f"  [ERROR] Gagal membaca file UAT Script: {e}")
        sys.exit(1)

    print(f"        Ditemukan {len(uat_data)} section:")
    for section_name, rows in uat_data.items():
        print(f"        - {section_name}: {len(rows)} baris data")
    print()

    # Mapping data
    print("  [2/3] Memproses mapping UAT Script -> Lampiran 7C...")
    lampiran_data = map_uat_to_lampiran(uat_data)

    # Hitung statistik
    total_filled = 0
    for section_name, rows in lampiran_data.items():
        filled = sum(1 for r in rows if r is not None)
        total_filled += filled
        section_count = next(c for n, c in LAMPIRAN_SECTIONS if n == section_name)
        print(f"        - {section_name}: {filled}/{section_count} baris terisi")
    print()

    # Generate dokumen Word
    output_path = OUTPUT_DIR / OUTPUT_FILENAME
    print(f"  [3/3] Membuat dokumen Lampiran 7C: {output_path}")
    try:
        create_lampiran_document(lampiran_data, output_path)
    except Exception as e:
        print(f"  [ERROR] Gagal membuat dokumen Word: {e}")
        sys.exit(1)

    print()
    print("  " + "=" * 56)
    print(f"  SELESAI! Total {total_filled} baris data berhasil dipindahkan.")
    print(f"  File output: {output_path.resolve()}")
    print("  " + "=" * 56)


if __name__ == "__main__":
    main()
