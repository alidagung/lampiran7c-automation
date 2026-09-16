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
    convert_uat_to_result_bytes,
    PRODUCT_PROFILES,
    OUTPUT_FILENAME,
)

# Label produk untuk dropdown (tampilan -> kunci di PRODUCT_PROFILES / "AUTO")
PRODUCT_OPTIONS = {
    "Deteksi otomatis": "AUTO",
    "Fund Transfer & Virtual Account (VA/FT)": "FT_VA",
    "QRIS": "QRIS",
}
PRODUCT_LABEL = {"FT_VA": "Fund Transfer & Virtual Account (VA/FT)", "QRIS": "QRIS"}


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

st.subheader("1. Pilih produk & unggah file UAT Script")

product_label = st.selectbox(
    "Jenis produk UAT Script",
    options=list(PRODUCT_OPTIONS.keys()),
    index=0,
    help=(
        "Pilih 'Fund Transfer & Virtual Account' untuk UAT VA/FT, atau 'QRIS' "
        "untuk UAT QRIS. Pilihan ini MEMAKSA konfigurasi produk yang dipakai. "
        "Gunakan 'Deteksi otomatis' bila ingin aplikasi menebak sendiri dari isi file."
    ),
)
selected_product = PRODUCT_OPTIONS[product_label]

if selected_product == "FT_VA":
    st.caption("📦 Mode **VA/FT** — section: Balance Inquiry, Intrabank, Interbank, RTGS, SKNBI, Virtual Account.")
elif selected_product == "QRIS":
    st.caption("📦 Mode **QRIS** — section: Balance Inquiry, API Transaction History List, QR MPM.")
else:
    st.caption("🔎 Mode **Otomatis** — produk ditentukan dari isi file.")

# Jenis dokumen keluaran
output_type = st.radio(
    "Jenis dokumen yang dibuat",
    options=["Lampiran 7C", "UAT Result"],
    index=0,
    horizontal=True,
    help=(
        "**Lampiran 7C**: tabel skenario hasil uji fungsional (untuk laporan ASPI). "
        "**UAT Result**: dokumen naratif per-case (semua case & semua section) "
        "lengkap dengan Daftar Isi."
    ),
)

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
    f"🚀 Proses & Buat {output_type}",
    type="primary",
    disabled=(uploaded_file is None),
    use_container_width=True,
)

# Simpan hasil di session_state agar tombol unduh tetap muncul setelah rerun
if "result_bytes" not in st.session_state:
    st.session_state.result_bytes = None
    st.session_state.result_stats = None
    st.session_state.result_warnings = None
    st.session_state.result_product = None
    st.session_state.result_output_name = None
    st.session_state.result_kind = None

if process_clicked and uploaded_file is not None:
    try:
        with st.spinner("Memproses file, mohon tunggu..."):
            file_bytes = uploaded_file.getvalue()
            if output_type == "UAT Result":
                docx_bytes, product_key, output_name = convert_uat_to_result_bytes(
                    file_bytes, product=selected_product
                )
                stats, warnings = None, None
            else:
                docx_bytes, stats, warnings, product_key, output_name = convert_uat_to_lampiran_bytes(
                    file_bytes, product=selected_product
                )

        st.session_state.result_bytes = docx_bytes
        st.session_state.result_stats = stats
        st.session_state.result_warnings = warnings
        st.session_state.result_product = product_key
        st.session_state.result_output_name = output_name
        st.session_state.result_kind = output_type

        # Info produk yang dipakai (berguna terutama pada mode Otomatis)
        st.info(f"Produk yang diproses: **{PRODUCT_LABEL.get(product_key, product_key)}** "
                f"— dokumen: **{output_type}**")

        if output_type == "UAT Result":
            st.success("Berhasil! Dokumen **UAT Result** dibuat. "
                       "Buka di Word lalu klik kanan pada Daftar Isi → *Update Field* "
                       "untuk menampilkan nomor halaman.")
        else:
            total = sum(stats.values())
            if total == 0:
                st.warning(
                    "File berhasil diproses, tetapi **tidak ada data hasil UAT** "
                    "yang ditemukan. Periksa kembali isi file UAT Script Anda "
                    "atau pastikan jenis produk yang dipilih sudah tepat."
                )
            else:
                st.success(f"Berhasil! Total **{total} baris** data dipindahkan "
                           f"ke Lampiran 7C.")
    except Exception as e:  # noqa: BLE001 - tampilkan pesan error ramah pengguna
        st.session_state.result_bytes = None
        st.session_state.result_stats = None
        st.session_state.result_warnings = None
        st.session_state.result_product = None
        st.session_state.result_kind = None
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

if st.session_state.result_bytes is not None:
    st.divider()
    st.subheader("3. Ringkasan & unduh")

    # Ringkasan section hanya relevan untuk Lampiran 7C (punya stats).
    if st.session_state.result_stats is not None:
        stats = st.session_state.result_stats
        used_product = st.session_state.get("result_product") or "FT_VA"
        used_sections = PRODUCT_PROFILES[used_product]["sections"]

        rows = []
        for section_name, total_baris in used_sections:
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

    # Tombol unduh (untuk Lampiran 7C maupun UAT Result)
    kind = st.session_state.get("result_kind") or "Lampiran 7C"
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    base_name = st.session_state.get("result_output_name") or OUTPUT_FILENAME
    download_name = base_name.replace(".docx", f" {timestamp}.docx")
    st.download_button(
        f"⬇️ Unduh {kind} (.docx)",
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
