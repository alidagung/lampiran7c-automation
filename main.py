"""
Lampiran 7C ASPI Automation Script
===================================
Script untuk mengotomasi pembuatan dokumen Lampiran 7C ASPI (format Word)
dari sumber data UAT Script (format Excel).

Fungsi utama:
- Membaca file UAT Script.xlsx dari folder input/
- Menyalin seluruh isi kolom Remarks langsung ke kolom Request di Lampiran 7C
- Membuat dokumen Word Lampiran 7C dari scratch dengan format yang sesuai
- Menyimpan hasil ke folder output/
"""

import os
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

# Info header dokumen Lampiran 7.C
NAMA_PENYEDIA_LAYANAN = "Bank Sahabat Sampoerna"
NAMA_PENGGUNA_LAYANAN = ""
TANGGAL_PENGUJIAN = ""

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
    # Catatan: "Interbank Transfer via BI FAST" (skenario 4.x) SENGAJA TIDAK
    # dipindahkan ke Lampiran 7C sesuai keputusan bisnis.
    ("RTGS Transfer", "5", "API RTGS Transfer", False),
    ("SKNBI Transfer", "6", "API SKNBI Transfer", False),
    ("Transfer VA", "7", "API Virtual Account", False),
    ("Transfer VA Prima", "8", "API Virtual Account", True),
    ("Transfer VA BI FAST", "9", "API Virtual Account", True),
]




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
# REMARKS PARSER
# ============================================================

import json
import re


# Penanda blok yang dikenali di kolom Remarks. Mendukung 2 format sumber:
#   Format A (mis. Intrabank): "URL:", "Headers:", "Request Body:", "Response:"
#   Format B (mis. Balance):   "Request URL:", "Request headers:",
#                              "Request body:", "Response body:"
# Setiap entri: (nama_kanonik, regex penanda)
_BLOCK_MARKERS = [
    ("url", r"(?:request\s+)?url(?:\s+endpoint)?\s*:"),
    ("headers", r"(?:request\s+)?head(?:er|ers)(?:\s+request)?\s*:"),
    ("request_body", r"request\s*body\s*:"),
    ("response", r"response(?:\s*body)?\s*:"),
]

# Regex gabungan untuk menemukan SEMUA penanda di mana saja (termasuk yang
# menempel di tengah baris, contoh: "}Response body:").
_MARKER_REGEX = re.compile(
    r"(?im)(" + "|".join(m[1] for m in _BLOCK_MARKERS) + r")"
)


def _classify_marker(marker_text):
    """Kembalikan nama kanonik blok dari teks penanda yang cocok."""
    t = marker_text.strip().lower()
    # Urutan pengecekan penting: 'request body' & 'response body' sebelum yang umum
    if re.match(r"request\s*body\s*:", t):
        return "request_body"
    if re.match(r"response(?:\s*body)?\s*:", t):
        return "response"
    if re.match(r"(?:request\s+)?url(?:\s+endpoint)?\s*:", t):
        return "url"
    if re.match(r"(?:request\s+)?head(?:er|ers)(?:\s+request)?\s*:", t):
        return "headers"
    return None


def _parse_remarks_blocks(text):
    """
    Pecah isi Remarks menjadi dict blok: {url, headers, request_body, response}.

    Menggunakan posisi tiap penanda (di mana saja dalam teks) sehingga tahan
    terhadap variasi format, termasuk penanda yang menempel di akhir baris
    sebelumnya (mis. "}Response body:").
    """
    blocks = {}
    matches = list(_MARKER_REGEX.finditer(text))
    if not matches:
        return blocks

    for i, m in enumerate(matches):
        name = _classify_marker(m.group(0))
        if name is None:
            continue
        start = m.end()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        content = text[start:end].strip()
        # Jangan timpa blok yang sudah terisi (ambil kemunculan pertama)
        if name not in blocks and content:
            blocks[name] = content

    return blocks


def _format_headers(headers_raw):
    """
    Format blok Headers dari UAT Script ke format array Lampiran 7C.

    Mendukung 2 bentuk input:
      Bentuk A (list baris polos):
        [
        Content-Type: application/json
        Authorization: Bearer xxx
        ]
      Bentuk B (JSON object, nilai berupa array):
        {
          "Authorization": ["Bearer xxx"],
          "Content-Type": ["application/json"]
        }

    Output (format Lampiran 7C):
        [
          "Content-Type=application/json",
          "Authorization=Bearer xxx"
        ]
    """
    pairs = []
    raw = headers_raw.strip()

    # Coba parse sebagai JSON object dulu (Bentuk B)
    parsed = None
    try:
        parsed = json.loads(raw)
    except (ValueError, TypeError):
        parsed = None

    if isinstance(parsed, dict):
        for key, val in parsed.items():
            if isinstance(val, list):
                val = ", ".join(str(v) for v in val)
            pairs.append(f'{key}={val}')
    else:
        # Bentuk A: parsing per baris
        for line in raw.split("\n"):
            s = line.strip().rstrip(",")
            if not s or s in ("[", "]", "{", "}"):
                continue
            if ":" in s:
                key, val = s.split(":", 1)
                pairs.append(f'{key.strip().strip(chr(34))}={val.strip().strip(chr(34))}')
            elif "=" in s:
                key, val = s.split("=", 1)
                pairs.append(f'{key.strip()}={val.strip()}')
            else:
                pairs.append(s)

    if not pairs:
        return ""

    out = ["["]
    for i, p in enumerate(pairs):
        comma = "," if i < len(pairs) - 1 else ""
        out.append(f'  "{p}"{comma}')
    out.append("]")
    return "\n".join(out)


def _pretty_json(raw):
    """
    Pretty-print JSON dengan indentasi 2 spasi jika bisa di-parse.
    Jika gagal parse, kembalikan teks asli apa adanya (sudah di-strip).
    """
    raw = raw.strip()
    if not raw:
        return ""
    try:
        obj = json.loads(raw)
        return json.dumps(obj, indent=2, ensure_ascii=False)
    except (ValueError, TypeError):
        return raw


def split_request_response(remarks_text):
    """
    Memisahkan DAN memformat isi kolom Remarks menjadi bagian Request dan
    Response sesuai format Lampiran 7C.

    Format Remarks (sumber, dari UAT Script):
        URL:
        <url>

        Headers:
        [ Key: Value ... ]

        Request Body:
        { ... }

        Response:
        { ... }

    Format hasil (Lampiran 7C):
        Request:
            URL Endpoint:
            <url>

            Header Request:
            [
              "Key=Value",
              ...
            ]

            Request Body:
            { ...pretty... }

        Response:
            Response Body:
            { ... }

    Returns:
        tuple: (request_content, response_content) keduanya sudah di-strip.
    """
    if not remarks_text or remarks_text.strip().lower() == "none":
        return "", ""

    text = remarks_text.replace("\r\n", "\n").replace("\r", "\n")

    blocks = _parse_remarks_blocks(text)

    url_raw = blocks.get("url", "")
    headers_raw = blocks.get("headers", "")
    body_raw = blocks.get("request_body", "")
    response_raw = blocks.get("response", "")

    # URL kadang diawali HTTP method (contoh: "POST https://..."). Ambil apa
    # adanya, hanya bersihkan whitespace.
    url_raw = url_raw.strip()

    # Susun bagian Request
    request_blocks = []
    if url_raw:
        request_blocks.append(f"URL Endpoint:\n{url_raw}")
    if headers_raw:
        request_blocks.append(f"Header Request:\n{_format_headers(headers_raw)}")
    if body_raw:
        request_blocks.append(f"Request Body:\n{_pretty_json(body_raw)}")

    request_part = "\n\n".join(request_blocks).strip()

    # Susun bagian Response
    if response_raw:
        response_part = f"Response Body:\n{_pretty_json(response_raw)}"
    else:
        response_part = ""

    # Fallback: jika format tidak dikenali sama sekali, pisahkan sederhana
    # pada penanda "Response" agar data tidak hilang.
    if not request_part and not response_part:
        m = re.search(r"(?i)response(?:\s*body)?\s*:", text)
        if m:
            request_part = text[:m.start()].strip()
            response_part = text[m.end():].strip()
        else:
            request_part = text.strip()

    return request_part, response_part


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

    Mengacu pada legend "KET" di UAT Script:
    - "Berhasil"          -> "PASS"     (sesuai hasil yang diharapkan)
    - "Gagal"             -> "NOT PASS" (tidak sesuai hasil yang diharapkan)
    - "Tidak dites"       -> "N/A"      (di luar scope / tidak ada alat pendukung)
    - "Belum dites"       -> "N/A"      (script belum waktunya dites)
    - "Siap dites"        -> "N/A"      (menunggu update fixing dari vendor)
    - "Butuh Konfirmasi"  -> "N/A"      (butuh konfirmasi user/vendor)
    - Lainnya             -> nilai asli
    """
    if not hasil_aktual:
        return ""
    lower = hasil_aktual.lower().strip()
    if lower == "berhasil":
        return "PASS"
    elif lower == "gagal":
        return "NOT PASS"
    elif lower in ("tidak dites", "belum dites", "siap dites", "butuh konfirmasi"):
        return "N/A"
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
    for uat_section, _, target_section, fill_empty_only in UAT_TO_LAMPIRAN_MAPPING:
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
            if hasil_aktual.lower().strip() in ("tidak dites", "belum dites"):
                # Tidak dites / Belum dites: baris tetap ditampilkan (Result N/A),
                # tetapi kolom Request, Response, dan Notes dikosongkan.
                notes = ""
                request_content = ""
                response_content = ""
            elif not remarks_text or remarks_text.lower() == "none":
                # Remarks kosong: Request dan Response kosong
                notes = ""
            else:
                # Pisahkan Remarks menjadi Request (URL+Headers+Request Body)
                # dan Response (isi setelah penanda "Response:")
                request_content, response_content = split_request_response(remarks_text)

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

    # Hanya render section yang benar-benar punya data hasil tes.
    # Section yang seluruh barisnya kosong (None) tidak ditampilkan.
    sections_to_render = []
    for section_name, section_count in LAMPIRAN_SECTIONS:
        section_data = lampiran_data.get(section_name, [])
        filled_rows = [r for r in section_data if r is not None]
        if filled_rows:
            sections_to_render.append((section_name, filled_rows))

    if not sections_to_render:
        # Tidak ada satupun layanan yang dites
        doc.add_paragraph("Tidak ada data hasil UAT yang ditemukan pada UAT Script.")
        doc.save(str(output_path))
        print(f"  [OK] Dokumen berhasil disimpan: {output_path}")
        return

    # Buat setiap section yang ada datanya
    for section_idx, (section_name, filled_rows) in enumerate(sections_to_render):
        # Tambah page break sebelum section (kecuali section pertama)
        if section_idx > 0:
            doc.add_page_break()

        # Judul dokumen (format resmi Lampiran 7.C)
        title1 = doc.add_paragraph()
        run1 = title1.add_run("Lampiran 7.C")
        run1.font.size = Pt(12)
        run1.font.bold = True

        title2 = doc.add_paragraph()
        run2 = title2.add_run("Skenario dan Hasil Uji Fungsionalitas")
        run2.font.size = Pt(11)
        run2.font.bold = True

        doc.add_paragraph()

        # Info penyedia & layanan
        doc.add_paragraph(f"Nama Penyedia Layanan : {NAMA_PENYEDIA_LAYANAN}")
        doc.add_paragraph(f"Nama Pengguna Layanan : {NAMA_PENGGUNA_LAYANAN}")
        doc.add_paragraph(f"Nama Layanan API      : {section_name}")
        doc.add_paragraph(f"Tanggal Pengujian     : {TANGGAL_PENGUJIAN}")

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

        # Data rows - hanya baris yang benar-benar dites (skip baris kosong)
        for display_no, row_data in enumerate(filled_rows, start=1):
            row_cells = table.add_row().cells
            row_cells[0].text = str(display_no)
            row_cells[1].text = row_data['service']
            row_cells[2].text = row_data['scenario']
            row_cells[3].text = row_data['expected_result']
            row_cells[4].text = row_data['request']
            row_cells[5].text = row_data['response']
            row_cells[6].text = row_data['result']
            row_cells[7].text = row_data['notes']

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
