# app.py
import streamlit as st
import pandas as pd
from pathlib import Path
from utils import export_to_template_excel, export_to_pdf

st.set_page_config(page_title="Perencanaan Anggaran - Export per Unit", layout="wide")

st.title("📘 Aplikasi Perencanaan & Penganggaran per Unit (Input Excel)")

# === PANEL SAMPING ===
st.sidebar.header("⚙️ Pengaturan")

uploaded_excel = st.sidebar.file_uploader(
    "Upload file Excel data anggaran", 
    type=["xlsx"], 
    accept_multiple_files=False
)
template_file = st.sidebar.file_uploader(
    "Upload template Excel (opsional, default Matriks FEBI.xlsx)", 
    type=["xlsx"], 
    accept_multiple_files=False
)

# Template default
DEFAULT_TEMPLATE = "Matriks FEBI.xlsx"
template_path = template_file if template_file else DEFAULT_TEMPLATE

# Jika belum upload, stop
if not uploaded_excel:
    st.info("⬆️ Silakan upload file Excel (misal: Data_Anggaran_FEBI.xlsx)")
    st.stop()

# === BACA FILE EXCEL ===
try:
    xls = pd.ExcelFile(uploaded_excel)
    st.sidebar.write(f"Sheet ditemukan: {xls.sheet_names}")
    sheet_selected = st.sidebar.selectbox("Pilih sheet data", xls.sheet_names)
    df = pd.read_excel(xls, sheet_name=sheet_selected)
except Exception as e:
    st.error(f"Gagal membaca Excel: {e}")
    st.stop()

st.subheader("📋 Preview Data (10 baris)")
st.dataframe(df.head(10))

# === TENTUKAN KOLOM UNIT ===
st.sidebar.markdown("---")
unit_col = st.sidebar.selectbox("Pilih kolom Unit", df.columns)
st.sidebar.write(f"Kolom Unit terpilih: **{unit_col}**")

# === PILIH UNIT ===
units = df[unit_col].dropna().unique().tolist()
units_sorted = sorted(units)
selected_unit = st.sidebar.selectbox("Pilih Unit", units_sorted)

# === FILTER DATA ===
df_filtered = df[df[unit_col] == selected_unit].copy()
st.subheader(f"📊 Data Unit: {selected_unit} — {len(df_filtered)} baris")
st.dataframe(df_filtered)

# === REKAP PER AKUN (opsional) ===
akun_cols = [c for c in df.columns if "akun" in c.lower()]
nilai_cols = [c for c in df.columns if any(x in c.lower() for x in ["jumlah", "nilai", "pagu"])]

if akun_cols and nilai_cols:
    akun_col = akun_cols[0]
    nilai_col = nilai_cols[0]
    st.markdown("---")
    st.subheader("💰 Rekapitulasi per Akun")
    pivot = df_filtered.groupby(akun_col)[nilai_col].sum().reset_index()
    pivot.columns = ["Kode Akun", "Total Anggaran"]
    st.dataframe(pivot)

# === EKSPOR HASIL ===
st.markdown("---")
st.subheader("📦 Ekspor Laporan")

col1, col2 = st.columns(2)

# --- EKSPOR EXCEL ---
with col1:
    out_excel_name = st.text_input(
        "Nama file output Excel", 
        value=f"Matriks_{selected_unit}.xlsx"
    )
    if st.button("📤 Ekspor ke Excel (pakai template)"):
        try:
            mapping = {
                "KODE_KEGIATAN": "KODE_KEGIATAN" if "KODE_KEGIATAN" in df.columns else None,
                "URAIAN": "URAIAN_SUBKOMPONEN" if "URAIAN_SUBKOMPONEN" in df.columns else None,
                "KODE_AKUN": "KODE_AKUN" if "KODE_AKUN" in df.columns else None,
                "JUMLAH": "JUMLAH1" if "JUMLAH1" in df.columns else None,
                "TOTAL": "TOTAL" if "TOTAL" in df.columns else None,
            }
            # bersihkan None
            mapping = {k: v for k, v in mapping.items() if v}

            tmp_template_path = template_path
            if not isinstance(template_path, str):
                tmp_template_path = "uploaded_template.xlsx"
                with open(tmp_template_path, "wb") as f:
                    f.write(template_path.getbuffer())

            out_path = export_to_template_excel(
                df_filtered, tmp_template_path, out_excel_name, 
                sheet_name="FEBI", mapping=mapping
            )
            st.success(f"✅ File Excel berhasil dibuat: {out_path}")
            with open(out_path, "rb") as f:
                st.download_button("⬇️ Unduh Excel", f, file_name=out_excel_name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"Gagal ekspor ke Excel: {e}")

# --- EKSPOR PDF ---
with col2:
    out_pdf_name = st.text_input(
        "Nama file output PDF", 
        value=f"Laporan_{selected_unit}.pdf"
    )
    if st.button("🧾 Ekspor ke PDF"):
        try:
            printable_df = df_filtered.copy()
            # batasi kolom agar muat di A4 landscape
            if printable_df.shape[1] > 10:
                printable_df = printable_df.iloc[:, :10]
            pdf_path = export_to_pdf(printable_df, out_pdf_name, title=f"LAPORAN {selected_unit}")
            st.success(f"✅ File PDF berhasil dibuat: {pdf_path}")
            with open(pdf_path, "rb") as f:
                st.download_button("⬇️ Unduh PDF", f, file_name=out_pdf_name, mime="application/pdf")
        except Exception as e:
            st.error(f"Gagal ekspor ke PDF: {e}")
