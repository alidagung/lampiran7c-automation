"""
Diagnostik cepat file UAT Script (TIDAK mengubah file apa pun).

Pemakaian (di folder repo):
    python cek_file.py "C:\\lokasi\\file\\UAT Script ....xlsx"

Menampilkan: sheet yang dipakai, baris header tabel, isi kolom per indeks,
section yang terbaca, dan contoh beberapa case (status + awal Remarks).
"""

import sys

from openpyxl import load_workbook

import main as m


def short(v, n=40):
    s = "" if v is None else str(v).replace("\n", " ").strip()
    return s[:n]


def main():
    if len(sys.argv) < 2:
        print('Pemakaian: python cek_file.py "path\\ke\\file.xlsx"')
        return
    path = sys.argv[1]
    wb = load_workbook(path, read_only=True, data_only=True)
    print("Sheet:", wb.sheetnames)
    ws = wb[m.UAT_SHEET_NAME] if m.UAT_SHEET_NAME in wb.sheetnames else wb.active
    print("Sheet dipakai:", ws.title)

    rows = list(ws.iter_rows(min_row=1, max_row=60, values_only=True))
    header_idx = None
    for i, r in enumerate(rows):
        cells = [short(c, 20).lower() for c in r]
        if any("nomor kasus" in c for c in cells) and any("langkah" in c for c in cells):
            header_idx = i
            break
    print("\nBaris header tabel (0-based):", header_idx)
    if header_idx is not None:
        for ci, v in enumerate(rows[header_idx]):
            if v is not None and str(v).strip():
                huruf = chr(ord("A") + ci) if ci < 26 else str(ci)
                print(f"  kolom {huruf} (index {ci}): {short(v)}")
        print("\nYang DIHARAPKAN aplikasi: E(4)=Nomor Kasus Tes, F(5)=Langkah Tes, "
              "H(7)=Hasil Aktual, I(8)=Remarks")

    data = m.read_uat_script(path)
    print("\nSection terbaca:")
    for sec, rs in data.items():
        print(f"  {sec}: {len(rs)} case")
    print("\nContoh case:")
    shown = 0
    for sec, rs in data.items():
        for r in rs:
            print(f"  [{sec}] {r['nomor_kasus_tes']} | hasil={short(r['hasil_aktual'], 15)!r} "
                  f"| remarks={short(r['remarks'], 50)!r}")
            shown += 1
            if shown >= 8:
                return


if __name__ == "__main__":
    main()
