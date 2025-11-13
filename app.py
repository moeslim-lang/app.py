# app.py
import streamlit as st
import pandas as pd
from pathlib import Path
from utils import export_to_template_excel, export_to_pdf

st.set_page_config(page_title="Perencanaan Anggaran - Export per Unit", layout="wide")

st.title("Perencanaan Anggaran — Export per Unit (Template Matriks FEBI)")

# Paths (ubah jika perlu)
DEFAULT_TEMPLATE = "Matriks FEBI.xlsx"

st.sidebar.header("Pengaturan")
uploaded_csv = st.sidebar.file_uploader("Upload file CSV data anggaran", type=["csv"], accept_multiple_files=False)
template_path = st.sidebar.file_uploader("Upload template Excel (opsional)", type=["xlsx"], accept_multiple_files=False)
if template_path:
    template_path_value = template_path
else:
    template_path_value = DEFAULT_TEMPLATE

st.sidebar.write("Masukkan tahun & kolom unit jika berbeda")
default_unit_col = st.sidebar.text_input("Nama kolom Unit di CSV", value="KDSATKER")
year_col = st.sidebar.text_input("Nama kolom Tahun di CSV (opsional)", value="TAHUN")
sheet_name = st.sidebar.text_input("Nama sheet template", value="FEBI")

if not uploaded_csv:
    st.info("Silakan upload file CSV (misal: Data Anggaran Satker ... .csv)")
    st.stop()

# read CSV robustly (try auto-detect sep)
try:
    df = pd.read_csv(uploaded_csv, sep=None, engine="python")
except Exception as e:
    st.error(f"Gagal membaca CSV: {e}")
    st.stop()

st.subheader("Preview data (10 baris)")
st.dataframe(df.head(10))

# Normalize unit column name
unit_col = default_unit_col
if unit_col not in df.columns:
    # try case-insensitive find
    found = None
    for c in df.columns:
        if c.strip().lower() == unit_col.strip().lower():
            found = c
            break
    if found:
        unit_col = found
    else:
        st.error(f"Kolom unit '{default_unit_col}' tidak ditemukan di CSV. Kolom tersedia: {list(df.columns)}")
        st.stop()

units = df[unit_col].dropna().unique().tolist()
units_sorted = sorted(units)
st.sidebar.subheader("Filter Unit")
selected_unit = st.sidebar.selectbox("Pilih Unit", options=units_sorted)

# Optional year filter
year_value = None
if year_col in df.columns:
    years = df[year_col].dropna().unique().tolist()
    if years:
        year_value = st.sidebar.selectbox("Filter Tahun (opsional)", options=["(Semua)"] + sorted([str(y) for y in years]), index=0)
    else:
        year_value = "(Semua)"

# Apply filtering
df_filtered = df[df[unit_col] == selected_unit].copy()
if year_value and year_value != "(Semua)" and year_col in df.columns:
    df_filtered = df_filtered[df_filtered[year_col].astype(str) == str(year_value)]

st.subheader(f"Hasil filter untuk unit: {selected_unit} — {len(df_filtered)} baris")
st.dataframe(df_filtered)

# small aggregation
agg_col_candidates = ["KODE_AKUN", "KODE_KEGIATAN", "URAIAN_SUBKOMPONEN", "JUMLAH1", "1763627000.1"]
# show basic pivot by akun if present
if "KODE_AKUN" in df_filtered.columns:
    pivot = df_filtered.groupby("KODE_AKUN").agg({"1763627000.1": "sum"}).reset_index().rename(columns={"1763627000.1": "PAGU_SUM"})
    st.subheader("Rekap per KODE_AKUN")
    st.dataframe(pivot)

# Export actions
st.markdown("---")
st.write("**Ekspor hasil**")

col1, col2 = st.columns(2)

with col1:
    out_excel_name = st.text_input("Nama file output Excel", value=f"Matriks_{selected_unit}.xlsx")
    if st.button("📤 Ekspor ke Excel (pakai template)"):
        # mapping: here you may want to tune template headers -> df columns
        # basic automatic mapping: try to fill columns by matching substring
        # The mapping dict keys are template header text; values are df column names
        # You can adjust mapping manually if needed.
        mapping = {
            # contoh mapping, sesuaikan bila template menggunakan header berbeda
            "KODE_KEGIATAN": "KODE_KEGIATAN",
            "URAIAN": "URAIAN_SUBKOMPONEN",
            "KODE_AKUN": "KODE_AKUN",
            "JUMLAH": "JUMLAH1",
            "HARGA": "HARGA1",
            "TOTAL": "1763627000.1",
            # tambahkan mapping lain sesuai template
        }
        # create a simplified df for export (reorder columns by mapping values)
        export_df = df_filtered.copy()
        # Ensure all mapping source columns exist
        # If some mapping targets not present, they will be filled with empty string
        # We'll pass df_filtered as-is and utils will try matching template headers
        try:
            # save temp to disk then use utils to copy template and insert
            template_file_path = template_path_value if isinstance(template_path_value, str) else template_path_value.name
            # handle case when user uploaded template via uploader: write to temporary path
            if not isinstance(template_path_value, str):
                tmp_template = Path("uploaded_template.xlsx")
                with open(tmp_template, "wb") as f:
                    f.write(template_path_value.getbuffer())
                template_file_path = str(tmp_template)
            out_path = export_to_template_excel(export_df, template_file_path, out_excel_name, sheet_name=sheet_name, mapping=mapping)
            st.success(f"File Excel dibuat: {out_path}")
            st.download_button("⬇️ Unduh Excel", data=open(out_path, "rb").read(), file_name=out_excel_name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"Gagal ekspor Excel: {e}")

with col2:
    out_pdf_name = st.text_input("Nama file output PDF", value=f"Laporan_{selected_unit}.pdf")
    if st.button("📤 Ekspor ke PDF"):
        try:
            pdf_path = f"{out_pdf_name}"
            # build a printable dataframe: choose subset / rename columns to fit
            printable_df = df_filtered.copy()
            # optionally select/rename columns to be more readable
            # If too many columns, keep a selection
            max_cols = 10
            if printable_df.shape[1] > max_cols:
                printable_df = printable_df.iloc[:, :max_cols]
            export_to_pdf(printable_df, pdf_path, title=f"LAPORAN {selected_unit}")
            st.success(f"File PDF dibuat: {pdf_path}")
            st.download_button("⬇️ Unduh PDF", data=open(pdf_path, "rb").read(), file_name=out_pdf_name, mime="application/pdf")
        except Exception as e:
            st.error(f"Gagal ekspor PDF: {e}")

st.markdown("---")
st.write("Tips: Jika template Excel memiliki header berbeda, buka file template dan sesuaikan dict `mapping` di app.py agar kolom template diisi dari kolom CSV yang benar.")
