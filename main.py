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

def parse_remarks_mitra_hit(remarks_text):
    """
    Parse format Remarks dari 'Mitra HIT' (Skenario 2, 3, 4, 5, 6).

    Format yang diharapkan:
        URL:
        [url]

        Header:
        [headers]

        RequestBody:
        [json body]

        Response:
        [json response]

    Kadang URL langsung dimulai dengan https://... tanpa label "URL:"

    Returns:
        dict with keys: url, headers, request_body, response
    """
    if not remarks_text or not remarks_text.strip():
        return None

    result = {
        "url": "",
        "headers": "",
        "request_body": "",
        "response": ""
    }

    text = remarks_text.strip()

    # Normalisasi line endings
    text = text.replace('\r\n', '\n').replace('\r', '\n')

    # Coba parse dengan label-label yang dikenal
    # Pattern: mencari section berdasarkan label
    # Labels yang mungkin: URL:, Header:, RequestBody:, Response:
    # Variasi: url:, Headers:, Request Body:, header:

    # Split berdasarkan pattern label
    # Gunakan regex untuk menangkap berbagai variasi
    url_pattern = r'(?i)(?:^|\n)\s*(?:URL)\s*:\s*\n?'
    header_pattern = r'(?i)(?:^|\n)\s*(?:Headers?)\s*:\s*\n?'
    request_body_pattern = r'(?i)(?:^|\n)\s*(?:Request\s*Body|RequestBody)\s*:\s*\n?'
    response_pattern = r'(?i)(?:^|\n)\s*(?:Response)\s*:\s*\n?'

    # Cari posisi masing-masing section
    url_match = re.search(url_pattern, text)
    header_match = re.search(header_pattern, text)
    request_body_match = re.search(request_body_pattern, text)
    response_match = re.search(response_pattern, text)

    # Kumpulkan semua section yang ditemukan beserta posisinya
    sections = []
    if url_match:
        sections.append(("url", url_match.end()))
    if header_match:
        sections.append(("headers", header_match.end()))
    if request_body_match:
        sections.append(("request_body", request_body_match.end()))
    if response_match:
        sections.append(("response", response_match.end()))

    # Jika tidak ada label yang ditemukan, cek apakah dimulai dengan URL langsung
    if not sections:
        # Cek apakah dimulai dengan http:// atau https://
        if text.startswith("http://") or text.startswith("https://"):
            # Anggap baris pertama adalah URL
            lines = text.split('\n')
            result["url"] = lines[0].strip()
            # Sisanya mungkin response atau body
            remaining = '\n'.join(lines[1:]).strip()
            if remaining:
                result["response"] = remaining
            return result
        return None

    # Sort berdasarkan posisi
    sections.sort(key=lambda x: x[1])

    # Extract konten masing-masing section
    for i, (section_name, start_pos) in enumerate(sections):
        if i + 1 < len(sections):
            # Ada section berikutnya, ambil sampai awal label berikutnya
            next_section_start = None
            # Cari posisi label berikutnya (termasuk label-nya)
            if sections[i + 1][0] == "url":
                next_match = url_match
            elif sections[i + 1][0] == "headers":
                next_match = header_match
            elif sections[i + 1][0] == "request_body":
                next_match = request_body_match
            elif sections[i + 1][0] == "response":
                next_match = response_match
            next_section_start = next_match.start()
            content = text[start_pos:next_section_start].strip()
        else:
            # Section terakhir, ambil sampai akhir
            content = text[start_pos:].strip()

        result[section_name] = content

    return result


def parse_remarks_bss_hit(remarks_text):
    """
    Parse format Remarks dari 'BSS YANG HIT' (Skenario 7, 8, 9).

    Format yang diharapkan:
        - url: [url]
        - headers: {json}
        - Request Body: {json}
        - Response: {json}

    Variasi:
        - body: {json}  (pengganti Request Body)
        - response: {json}  (huruf kecil)

    Returns:
        dict with keys: url, headers, request_body, response
    """
    if not remarks_text or not remarks_text.strip():
        return None

    result = {
        "url": "",
        "headers": "",
        "request_body": "",
        "response": ""
    }

    text = remarks_text.strip()
    text = text.replace('\r\n', '\n').replace('\r', '\n')

    # Pattern untuk format "- key: value" (bisa multiline value)
    # Kita cari semua keys yang dikenal
    url_pattern = r'(?i)-\s*url\s*:\s*'
    headers_pattern = r'(?i)-\s*headers?\s*:\s*'
    request_body_pattern = r'(?i)-\s*(?:Request\s*Body|body)\s*:\s*'
    response_pattern = r'(?i)-\s*(?:Response|response)\s*:\s*'

    # Cari posisi masing-masing
    url_match = re.search(url_pattern, text)
    headers_match = re.search(headers_pattern, text)
    request_body_match = re.search(request_body_pattern, text)
    response_match = re.search(response_pattern, text)

    # Kumpulkan sections yang ditemukan
    sections = []
    if url_match:
        sections.append(("url", url_match.end()))
    if headers_match:
        sections.append(("headers", headers_match.end()))
    if request_body_match:
        sections.append(("request_body", request_body_match.end()))
    if response_match:
        sections.append(("response", response_match.end()))

    if not sections:
        return None

    # Sort berdasarkan posisi
    sections.sort(key=lambda x: x[1])

    # Extract konten
    for i, (section_name, start_pos) in enumerate(sections):
        if i + 1 < len(sections):
            # Ambil sampai awal pattern berikutnya
            next_key = sections[i + 1][0]
            if next_key == "url":
                next_match = url_match
            elif next_key == "headers":
                next_match = headers_match
            elif next_key == "request_body":
                next_match = request_body_match
            elif next_key == "response":
                next_match = response_match
            content = text[start_pos:next_match.start()].strip()
        else:
            content = text[start_pos:].strip()

        result[section_name] = content

    return result


def parse_remarks(remarks_text, scenario_prefix):
    """
    Parse kolom Remarks berdasarkan format yang sesuai dengan skenario.

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
        return parse_remarks_bss_hit(remarks_text)
    elif scenario_prefix in MITRA_HIT_SCENARIOS:
        return parse_remarks_mitra_hit(remarks_text)
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

def read_uat_script(filepath):
    """
    Membaca file UAT Script.xlsx dan mengembalikan data terstruktur.

    Returns:
        dict: {
            section_header: [
                {
                    'nomor_skenario': str,  # e.g., "1.1", "2.3"
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

    # Identifikasi header row (baris pertama biasanya header kolom)
    # Kolom: Kategori Tes, Nama Modul, Nomor Skenario, Nomor Kasus Tes,
    #         Langkah Tes, Hasil yang diharapkan, Hasil Aktual, Remarks,
    #         Tanggal Pelaksanaan, Jenis Script, Pelaksana
    header_row_idx = 0

    for idx, row in enumerate(rows_list):
        if row and row[0] and str(row[0]).strip().lower() in ['kategori tes', 'kategori', 'no']:
            header_row_idx = idx
            break

    # Process data rows
    for idx in range(header_row_idx + 1, len(rows_list)):
        row = rows_list[idx]
        if not row or all(cell is None or str(cell).strip() == '' for cell in row):
            continue

        # Cek apakah ini section header
        # Section header biasanya hanya punya nilai di kolom pertama atau kedua
        first_cell = str(row[0]).strip() if row[0] else ""
        second_cell = str(row[1]).strip() if len(row) > 1 and row[1] else ""

        # Deteksi section header berdasarkan keyword yang dikenal
        section_keywords = [
            "Balance Services", "Intrabank Transfer", "Interbank Transfer",
            "Interbank Transfer via BI FAST", "RTGS Transfer", "SKNBI Transfer",
            "Transfer VA", "Transfer VA Prima", "Transfer VA BI FAST"
        ]

        is_section_header = False
        detected_section = None

        for keyword in section_keywords:
            if keyword.lower() in first_cell.lower() or keyword.lower() in second_cell.lower():
                # Pastikan ini bukan data row biasa
                # Section header biasanya tidak punya nomor skenario (kolom ke-3 atau ke-4 kosong)
                nomor_col = str(row[3]).strip() if len(row) > 3 and row[3] else ""
                if not nomor_col or not re.match(r'\d+\.\d+', nomor_col):
                    is_section_header = True
                    detected_section = keyword
                    break

        if is_section_header:
            current_section = detected_section
            if current_section not in data:
                data[current_section] = []
            continue

        # Jika belum ada section, skip
        if current_section is None:
            continue

        # Parse data row
        # Index kolom (0-based):
        # 0: Kategori Tes
        # 1: Nama Modul
        # 2: Nomor Skenario
        # 3: Nomor Kasus Tes
        # 4: Langkah Tes
        # 5: Hasil yang diharapkan
        # 6: Hasil Aktual
        # 7: Remarks
        nomor_kasus = str(row[3]).strip() if len(row) > 3 and row[3] else ""
        if not nomor_kasus:
            continue

        row_data = {
            'nomor_skenario': str(row[2]).strip() if len(row) > 2 and row[2] else "",
            'nomor_kasus_tes': nomor_kasus,
            'langkah_tes': str(row[4]).strip() if len(row) > 4 and row[4] else "",
            'hasil_diharapkan': str(row[5]).strip() if len(row) > 5 and row[5] else "",
            'hasil_aktual': str(row[6]).strip() if len(row) > 6 and row[6] else "",
            'remarks': str(row[7]).strip() if len(row) > 7 and row[7] else "",
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
            result_value = hasil_aktual

            # Cek kondisi khusus
            if hasil_aktual.lower() == "tidak dites":
                notes = "Tidak dites"
                result_value = "Tidak dites"
            elif not remarks_text or remarks_text.lower() == "none":
                notes = remarks_text if remarks_text and remarks_text.lower() != "none" else ""
                # Request dan Response kosong
            elif hasil_aktual.lower() == "berhasil" or remarks_text:
                # Parse remarks untuk mengambil data
                parsed = parse_remarks(remarks_text, scenario_prefix)
                if parsed:
                    url = parsed.get('url', '')
                    headers = parsed.get('headers', '')
                    request_body = parsed.get('request_body', '')
                    response = parsed.get('response', '')

                    request_content = f"URL Endpoint:\n{url}\nHeader Request:\n{headers}\nRequest Body:\n{request_body}"
                    response_content = f"Response Body:\n{response}"
                else:
                    # Remarks ada tapi tidak bisa di-parse
                    notes = remarks_text

            lampiran_data[target_section][row_idx] = {
                'no': sub_num,
                'service': target_section,
                'scenario': row_data.get('langkah_tes', f"Skenario {row_data['nomor_kasus_tes']}"),
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
                row_cells[1].text = section_name
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
