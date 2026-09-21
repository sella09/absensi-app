import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime, time
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Laporan Absensi", layout="wide")
st.title("🕒 Konversi Absensi ke Laporan Harian")
st.caption("Upload Excel absensi mentah → otomatis jadi laporan Nama, Tanggal, Jam Masuk, Jam Pulang, Status, Menit Terlambat.")


JAM_MASUK_NORMAL = time(8, 30)
JAM_PULANG_NORMAL = time(16, 0)
BATAS_HALFDAY = time(13, 0)
HARI_LIBUR = [5, 6]  # Sabtu, Minggu


def parse_datetime(s):
    if pd.isna(s):
        return None
    if isinstance(s, datetime):
        return s
    s = str(s).strip()
    for fmt in ("%d/%m/%Y %H:%M:%S", "%Y-%m-%d %H:%M:%S",
                "%d-%m-%Y %H:%M:%S", "%m/%d/%Y %H:%M:%S"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def jam_str(t):
    if t is None:
        return ""
    return t.strftime("%H:%M:%S")


def hitung_menit_terlambat(jam_masuk):
    if jam_masuk is None:
        return 0
    normal_dt = datetime.combine(datetime.today(), JAM_MASUK_NORMAL)
    masuk_dt = datetime.combine(datetime.today(), jam_masuk)
    selisih = (masuk_dt - normal_dt).total_seconds() / 60
    return max(0, int(selisih))


def konversi_absensi(df_raw):
    col_nama = None
    col_waktu = None
    for c in df_raw.columns:
        cl = str(c).lower()
        if "nama" in cl:
            col_nama = c
        if "tgl" in cl or "waktu" in cl or "tanggal" in cl:
            col_waktu = c

    if col_nama is None or col_waktu is None:
        st.error(f"Kolom 'Nama' atau 'Tgl/Waktu' tidak ditemukan. Kolom ada: {list(df_raw.columns)}")
        return pd.DataFrame()

    df = df_raw[[col_nama, col_waktu]].copy()
    df.columns = ["Nama", "TglWaktu"]
    df["TglWaktu"] = df["TglWaktu"].apply(parse_datetime)
    df = df.dropna(subset=["TglWaktu"])

    df["Tanggal"] = df["TglWaktu"].dt.date
    df["Jam"] = df["TglWaktu"].dt.time
    df["Hari"] = df["TglWaktu"].dt.weekday

    hasil = []
    for (nama, tanggal), group in df.groupby(["Nama", "Tanggal"]):
        hari = group["Hari"].iloc[0]
        jam_list = sorted(group["Jam"].tolist())

        jam_masuk = None
        for j in jam_list:
            if j < time(12, 0):
                jam_masuk = j
                break

        jam_pulang = None
        for j in reversed(jam_list):
            if j >= time(12, 0):
                jam_pulang = j
                break

        if hari in HARI_LIBUR:
            status = "Libur"
            menit_terlambat = 0
        elif jam_masuk is None and jam_pulang is None:
            status = "Tanpa Keterangan"
            menit_terlambat = 0
        elif jam_pulang is not None and jam_pulang <= BATAS_HALFDAY:
            status = "Half-day"
            menit_terlambat = 0
        elif jam_masuk is not None and jam_masuk > JAM_MASUK_NORMAL:
            status = "Terlambat"
            menit_terlambat = hitung_menit_terlambat(jam_masuk)
        elif jam_masuk is None:
            status = "Tanpa Keterangan"
            menit_terlambat = 0
        else:
            status = "Tepat Waktu"
            menit_terlambat = 0

        hasil.append({
            "Nama": nama,
            "Tanggal": tanggal.strftime("%d/%m/%Y"),
            "Jam Masuk": jam_str(jam_masuk),
            "Jam Pulang": jam_str(jam_pulang),
            "Status": status,
            "Menit Terlambat": menit_terlambat,
        })

    df_hasil = pd.DataFrame(hasil)
    if not df_hasil.empty:
        df_hasil = df_hasil.sort_values(["Nama", "Tanggal"]).reset_index(drop=True)
    return df_hasil


def tulis_df(writer, df, sheet, startrow=0, kolom_angka=None):
    df.to_excel(writer, sheet_name=sheet, index=False, startrow=startrow)
    ws = writer.sheets[sheet]
    kolom_angka = kolom_angka or []

    for col_name in kolom_angka:
        if col_name not in df.columns:
            continue
        idx = df.columns.get_loc(col_name) + 1
        letter = get_column_letter(idx)
        for row in range(startrow + 2, startrow + 2 + len(df)):
            ws[f"{letter}{row}"].number_format = "0"

    for col in range(1, 1 + len(df.columns)):
        c = ws.cell(row=startrow + 1, column=col)
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="FFE699")
        c.alignment = Alignment(horizontal="center", vertical="center")

    for col_cells in ws.columns:
        max_len = 0
        letter = col_cells[0].column_letter
        for cell in col_cells:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[letter].width = min(max_len + 2, 30)


uploaded = st.file_uploader("Upload file Excel absensi (.xls / .xlsx)", type=["xls", "xlsx"])

if uploaded:
    try:
        df_raw = pd.read_excel(uploaded, header=0)
    except Exception as e:
        st.error(f"Gagal baca file: {e}")
        st.stop()

    st.markdown("### 📋 Data Mentah")
    st.dataframe(df_raw.head(20), use_container_width=True)

    if st.button("🚀 Konversi ke Laporan", type="primary"):
        df_hasil = konversi_absensi(df_raw)

        if df_hasil.empty:
            st.warning("Tidak ada data yang bisa dikonversi.")
        else:
            st.markdown("### ✅ Hasil Laporan")
            st.dataframe(df_hasil, use_container_width=True)

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Total Baris", len(df_hasil))
            c2.metric("Terlambat", (df_hasil["Status"] == "Terlambat").sum())
            c3.metric("Tanpa Keterangan", (df_hasil["Status"] == "Tanpa Keterangan").sum())
            c4.metric("Half-day", (df_hasil["Status"] == "Half-day").sum())

            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                tulis_df(writer, df_hasil, "Laporan Absensi",
                         kolom_angka=["Menit Terlambat"])

                for nama in df_hasil["Nama"].unique():
                    sub = df_hasil[df_hasil["Nama"] == nama].reset_index(drop=True)
                    sheet_name = str(nama)[:28]
                    tulis_df(writer, sub, sheet_name, kolom_angka=["Menit Terlambat"])

            st.download_button(
                label="⬇️ Download Excel Laporan",
                data=buffer.getvalue(),
                file_name="laporan_absensi.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

else:
    st.info("Silakan upload file Excel absensi terlebih dahulu.")
