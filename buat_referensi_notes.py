"""
Buat / perbarui referensi Notes Lampiran 7C dari UAT Script acuan (Veloka /
PT Integrasi Sinaptiq). Dipakai bila UAT Script mitra lain berstatus tidak
dites tetapi kolom Remarks-nya kosong.

Pemakaian:
    python buat_referensi_notes.py "input/UAT Script ... Sinaptiq ....xlsx"

Hasil: referensi/notes_veloka.json
"""
import json
import sys
from pathlib import Path

from main import read_uat_script, status_tidak_dites, _alasan_level_api

TUJUAN = Path(__file__).resolve().parent / "referensi" / "notes_veloka.json"


def buat(sumber):
    data = read_uat_script(sumber)
    hasil = {"sumber": Path(sumber).name, "section": {}}
    for section, rows in data.items():
        alasan_api = _alasan_level_api(rows)
        kasus = []
        for r in rows:
            remarks = str(r.get("remarks") or "").strip()
            kasus.append({
                "no": r["nomor_kasus_tes"],
                "langkah": r.get("langkah_tes", ""),
                # Remarks case itu sendiri (hanya bila case tidak dites)
                "remarks": remarks if status_tidak_dites(r.get("hasil_aktual")) else "",
            })
        hasil["section"][section] = {"alasan_api": alasan_api, "kasus": kasus}
    TUJUAN.parent.mkdir(exist_ok=True)
    TUJUAN.write_text(json.dumps(hasil, ensure_ascii=False, indent=1), encoding="utf-8")
    return hasil


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print(__doc__)
        sys.exit(1)
    h = buat(sys.argv[1])
    n = sum(1 for s in h["section"].values() for k in s["kasus"] if k["remarks"])
    print(f"Tersimpan: {TUJUAN} ({len(h['section'])} section, {n} Remarks)")
