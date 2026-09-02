"""
Aplikasi Web Lampiran 7C Automation
===================================
Aplikasi web sederhana berbasis Streamlit untuk mengubah file UAT Script
(Excel) menjadi dokumen Lampiran 7C ASPI (Word).

Cara pakai:
    1. Install dependency:  pip install -r requirements.txt
    2. Jalankan aplikasi:   streamlit run app.py
    3. Browser akan terbuka otomatis. Upload file UAT Script.xlsx,
       klik "Proses", lalu unduh hasilnya.
"""

from datetime import datetime

import streamlit as st

from main import (
    convert_uat_to_lampiran_bytes,
    LAMPIRAN_SECTIONS,
    OUTPUT_FILENAME,
)


# ============================================================
# KONFIGURASI HALAMAN
# ============================================================

st.set_page_config(
    page_title="Automation Lampiran 7C ASPI",
    page_icon="📄",
    layout="centered",
)


# ============================================================
# HEADER
# ============================================================

st.title("📄 Automation Lampiran 7C ASPI")
st.caption(
    "Ubah file **UAT Script** (Excel) menjadi dokumen **Lampiran 7C** (Word) "
    "secara otomatis. Cukup unggah file, klik proses, lalu unduh hasilnya."
)

st.divider()


# ============================================================
# LANGKAH 1 — UNGGAH FILE
# ============================================================

st.subheader("1. Unggah file UAT Script")

uploaded_file = st.file_uploader(
    "Pilih file UAT Script (.xlsx)",
    type=["xlsx"],
    help="File Excel hasil UAT. Harus memiliki sheet berisi data UAT Script.",
)

if uploaded_file is not None:
    st.success(f"File terpilih: **{uploaded_file.name}** "
               f"({uploaded_file.size / 1024:.1f} KB)")


# ============================================================
# LANGKAH 2 — PROSES
# ============================================================

st.subheader("2. Proses")

process_clicked = st.button(
    "🚀 Proses & Buat Lampiran 7C",
    type="primary",
    disabled=(uploaded_file is None),
    use_container_width=True,
)

# Simpan hasil di session_state agar tombol unduh tetap muncul setelah rerun
if "result_bytes" not in st.session_state:
    st.session_state.result_bytes = None
    st.session_state.result_stats = None
    st.session_state.result_warnings = None

if process_clicked and uploaded_file is not None:
    try:
        with st.spinner("Memproses file, mohon tunggu..."):
            file_bytes = uploaded_file.getvalue()
            docx_bytes, stats, warnings = convert_uat_to_lampiran_bytes(file_bytes)

        st.session_state.result_bytes = docx_bytes
        st.session_state.result_stats = stats
        st.session_state.result_warnings = warnings

        total = sum(stats.values())
        if total == 0:
            st.warning(
                "File berhasil diproses, tetapi **tidak ada data hasil UAT** "
                "yang ditemukan. Periksa kembali isi file UAT Script Anda."
            )
        else:
            st.success(f"Berhasil! Total **{total} baris** data dipindahkan "
                       f"ke Lampiran 7C.")
    except Exception as e:  # noqa: BLE001 - tampilkan pesan error ramah pengguna
        st.session_state.result_bytes = None
        st.session_state.result_stats = None
        st.session_state.result_warnings = None
        st.error(
            "Gagal memproses file. Pastikan file yang diunggah adalah "
            "UAT Script (.xlsx) dengan format yang benar."
        )
        with st.expander("Detail teknis (untuk troubleshooting)"):
            st.code(str(e))


# ============================================================
# PANEL PERINGATAN (DATA ABNORMAL)
# ============================================================

if st.session_state.get("result_warnings"):
    _warns = st.session_state.result_warnings
    st.warning(
        f"⚠️ Ditemukan **{len(_warns)} hal yang perlu dicek** pada data. "
        "Data TIDAK diubah — parameter & value tetap apa adanya. "
        "Berikut baris yang sebaiknya diperiksa manual:"
    )
    with st.expander(f"Lihat {len(_warns)} peringatan", expanded=True):
        for _w in _warns:
            st.markdown(f"- {_w}")
elif st.session_state.get("result_stats") is not None:
    st.info("✅ Tidak ada data abnormal terdeteksi. Parameter & value aman (tidak diubah).")


# ============================================================
# LANGKAH 3 — RINGKASAN & UNDUH
# ============================================================

if st.session_state.result_stats is not None:
    st.divider()
    st.subheader("3. Ringkasan hasil")

    stats = st.session_state.result_stats

    # Tabel ringkasan: layanan yang dites (baris terisi) vs total baris template
    rows = []
    for section_name, total_baris in LAMPIRAN_SECTIONS:
        terisi = stats.get(section_name, 0)
        status = "✅ Ditampilkan" if terisi > 0 else "— Tidak dites"
        rows.append(
            {
                "Layanan API": section_name,
                "Baris terisi": terisi,
                "Status": status,
            }
        )

    st.dataframe(rows, use_container_width=True, hide_index=True)

    total = sum(stats.values())
    jumlah_layanan = sum(1 for v in stats.values() if v > 0)
    col1, col2 = st.columns(2)
    col1.metric("Total baris data", total)
    col2.metric("Layanan ditampilkan", jumlah_layanan)

    # Tombol unduh
    if st.session_state.result_bytes:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        download_name = OUTPUT_FILENAME.replace(
            ".docx", f" {timestamp}.docx"
        )
        st.download_button(
            "⬇️ Unduh Lampiran 7C (.docx)",
            data=st.session_state.result_bytes,
            file_name=download_name,
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            ),
            type="primary",
            use_container_width=True,
        )


# ============================================================
# FOOTER
# ============================================================

st.divider()
with st.expander("ℹ️ Tentang aplikasi ini"):
    st.markdown(
        """
        Aplikasi ini membaca file **UAT Script** lalu menghasilkan dokumen
        **Lampiran 7C** dengan aturan berikut:

        - Hanya layanan yang **benar-benar dites** yang ditampilkan.
        - Kolom **Scenario** diambil dari kolom *Langkah Tes*.
        - Kolom **Request** & **Response** dipisah dan diformat
          (Header Request, Request Body, dan Response Body di-*compress*
          menjadi satu baris).
        - Case **Interbank via BI FAST** tidak dipindahkan.
        - Baris **Tidak dites / Belum dites** ditampilkan dengan hasil `N/A`
          dan kolom Request/Response/Notes dikosongkan.
        """
    )
