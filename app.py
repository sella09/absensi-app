import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime, time, timedelta, date
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Laporan Absensi", layout="wide")
st.title("🕒 Konversi Absensi ke Laporan Harian")
st.caption("Upload Excel absensi → input rentang tanggal → otomatis jadi laporan lengkap.")


# === ATURAN ===
JAM_MASUK_NORMAL = time(9, 0, 0)      # ≥ 09:00:00 = terlambat
BATAS_HALFDAY = time(13, 0)           # pulang ≤ 13:00 = half-day (khusus Sabtu)
HARI_LIBUR = [6]                      # hanya Minggu


def parse_datetime(s):
    if pd.isna(s):
        return None
    if isinstance(s, datetime):
        return s
    s = str(s).strip()
    for fmt in ("%Y-%m-%d %H:%M:%S", "%d/%m/%Y %H:%M:%S",
                "%d-%m-%Y %H:%M:%S", "%m/%d/%Y %H:%M:%S",
                "%Y-%m-%d %H:%M", "%d/%m/%Y %H:%M"):
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
    """≥ 09:00:00 = terlambat. Minimal 1 menit."""
    if jam_masuk is None:
        return 0
    batas = datetime.combine(datetime.today(), JAM_MASUK_NORMAL)
    masuk_dt = datetime.combine(datetime.today(), jam_masuk)
    selisih_detik = (masuk_dt - batas).total_seconds()
    if selisih_detik < 0:
        return 0
    menit = int(selisih_detik // 60)
    if selisih_detik > 0 and menit == 0:
        menit = 1
    return menit


def cari_kolom(df, kandidat):
    """Cari kolom dari daftar kandidat (prioritas urutan)."""
    for k in kandidat:
        for c in df.columns:
            if k in str(c).lower().strip():
                return c
    return None


def konversi_absensi(df_raw, tgl_mulai, tgl_selesai):
    """Generate laporan lengkap: Nama × Tanggal dalam rentang."""

    # Deteksi kolom: coba tiap kandidat sesuai prioritas
    col_nama = cari_kolom(df_raw, ["name", "nama"])
    col_waktu = cari_kolom(df_raw, ["date/time", "datetime", "tgl/waktu", "tanggal", "waktu"])

    if col_nama is None or col_waktu is None:
        st.error(f"Kolom Nama/Tanggal tidak ditemukan. Kolom ada: {list(df_raw.columns)}")
        return pd.DataFrame()

    st.info(f"📌 Kolom terdeteksi — Nama: `{col_nama}` | Waktu: `{col_waktu}`")

    df = df_raw[[col_nama, col_waktu]].copy()
    df.columns = ["Nama", "TglWaktu"]
    df["TglWaktu"] = df["TglWaktu"].apply(parse_datetime)
    df = df.dropna(subset=["TglWaktu"])

    df["Tanggal"] = df["TglWaktu"].dt.date
    df["Jam"] = df["TglWaktu"].dt.time
    df["Hari"] = df["TglWaktu"].dt.weekday

    semua_nama = sorted(df["Nama"].unique().tolist())

    semua_tanggal = []
    tgl = tgl_mulai
    while tgl <= tgl_selesai:
        semua_tanggal.append(tgl)
        tgl += timedelta(days=1)

    hasil = []
    for nama in semua_nama:
        df_nama = df[df["Nama"] == nama]

        for tanggal in semua_tanggal:
            df_hari = df_nama[df_nama["Tanggal"] == tanggal]
            hari = tanggal.weekday()

            if df_hari.empty:
                status = "Libur" if hari in HARI_LIBUR else "Tidak Ada Absen"
                hasil.append({
                    "Nama": nama,
                    "Tanggal": tanggal.strftime("%d/%m/%Y"),
                    "Jam Masuk": "",
                    "Jam Pulang": "",
                    "Status": status,
                    "Menit Terlambat": 0,
                })
                continue

            jam_list = sorted(df_hari["Jam"].tolist())

            jam_masuk = None
            for j in jam_list:
                if j < time(12, 0):
                    jam_masuk = j
                    break

            jam_pulang = None
            for j in jam_list:
                if j >= time(12, 0):
                    jam_pulang = j
                    break

            if hari in HARI_LIBUR:
                status = "Libur"
                menit_terlambat = 0
            elif jam_masuk is None and jam_pulang is None:
                status = "Tidak Ada Absen"
                menit_terlambat = 0
            elif jam_masuk is not None and jam_pulang is None:
                if jam_masuk >= JAM_MASUK_NORMAL:
                    status = "Terlambat"
                    menit_terlambat = hitung_menit_terlambat(jam_masuk)
                else:
                    status = "Hanya Absen Masuk"
                    menit_terlambat = 0
            elif jam_masuk is None and jam_pulang is not None:
                status = "Hanya Absen Pulang"
                menit_terlambat = 0
            elif hari == 5 and jam_pulang <= BATAS_HALFDAY:
                status = "Half-day"
                menit_terlambat = 0
            elif jam_masuk >= JAM_MASUK_NORMAL:
                status = "Terlambat"
                menit_terlambat = hitung_menit_terlambat(jam_masuk)
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


# =========================================================
# UI
# =========================================================
uploaded = st.file_uploader("Upload file Excel absensi (.xls / .xlsx)", type=["xls", "xlsx"])

if uploaded:
    df_raw = None
    errors = []

    try:
        uploaded.seek(0)
        df_raw = pd.read_excel(uploaded, header=0, engine="xlrd")
    except Exception as e:
        errors.append(f"xlrd: {e}")

    if df_raw is None:
        try:
            uploaded.seek(0)
            df_raw = pd.read_excel(uploaded, header=0, engine="openpyxl")
        except Exception as e:
            errors.append(f"openpyxl: {e}")

    if df_raw is None:
        try:
            uploaded.seek(0)
            df_raw = pd.read_excel(uploaded, header=0)
        except Exception as e:
            errors.append(f"default: {e}")

    if df_raw is None:
        st.error("Gagal baca file. Detail error:")
        for err in errors:
            st.code(err)
        st.stop()

    st.markdown("### 📋 Data Mentah")
    st.dataframe(df_raw.head(10), use_container_width=True)

    st.markdown("### 📅 Rentang Tanggal Laporan")
    c1, c2 = st.columns(2)
    with c1:
        tgl_mulai = st.date_input("Tanggal Mulai", value=date(2026, 8, 26))
    with c2:
        tgl_selesai = st.date_input("Tanggal Selesai", value=date(2026, 8, 29))

    if st.button("🚀 Konversi ke Laporan", type="primary"):
        if tgl_mulai > tgl_selesai:
            st.error("Tanggal Mulai harus ≤ Tanggal Selesai")
            st.stop()

        df_hasil = konversi_absensi(df_raw, tgl_mulai, tgl_selesai)

        if df_hasil.empty:
            st.warning("Tidak ada data yang bisa dikonversi.")
        else:
            st.markdown("### ✅ Hasil Laporan")
            st.dataframe(df_hasil, use_container_width=True)

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Total Baris", len(df_hasil))
            c2.metric("Terlambat", (df_hasil["Status"] == "Terlambat").sum())
            c3.metric("Tidak Ada Absen", (df_hasil["Status"] == "Tidak Ada Absen").sum())
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
