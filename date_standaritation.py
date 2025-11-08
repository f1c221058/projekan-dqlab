import pandas as pd
import re

def normalize_tanggal_transaksi(input_xlsx_path: str, output_csv_path: str) -> None:
    # Baca file Excel
    df = pd.read_excel(input_xlsx_path, dtype=str)

    # Peta bulan Indonesia ke Inggris
    bulan_map = {
        'januari': 'January', 'jan': 'January',
        'februari': 'February', 'feb': 'February',
        'maret': 'March', 'mar': 'March',
        'april': 'April', 'apr': 'April',
        'mei': 'May', 'mei.': 'May',
        'juni': 'June', 'jun': 'June',
        'juli': 'July', 'jul': 'July',
        'agustus': 'August', 'agu': 'August', 'aug': 'August',
        'september': 'September', 'sep': 'September',
        'oktober': 'October', 'okt': 'October', 'oct': 'October',
        'november': 'November', 'nov': 'November',
        'desember': 'December', 'des': 'December', 'dec': 'December'
    }

    def parse_date(s):
        if pd.isna(s):
            return s
        s = str(s).strip()
        s = re.sub(r"[,'’‘]", " ", s)
        s = re.sub(r"\s+", " ", s)
        s = " ".join([bulan_map.get(w.lower(), w) for w in s.split()])

        # Coba parse otomatis
        for dayfirst in (True, False):
            dt = pd.to_datetime(s, errors='coerce', dayfirst=dayfirst, infer_datetime_format=True)
            if pd.notna(dt):
                return dt.strftime("%d-%m-%Y")
        return s  # kembalikan jika tidak bisa diparse

    # Deteksi kolom tanggal otomatis berdasarkan sample
    for col in df.columns:
        sample = df[col].dropna().astype(str).head(10)
        detect = pd.to_datetime(sample, errors='coerce', dayfirst=True, infer_datetime_format=True)
        if detect.notna().sum() >= 3:
            df[col] = df[col].apply(parse_date)

    # Simpan hasil
    df.to_csv(output_csv_path, index=False)
    print(f"✅ File berhasil dinormalisasi dan disimpan ke: {output_csv_path}")

# Jalankan fungsi
normalize_tanggal_transaksi(
    "penjualan_dqmart_01-beta.xlsx",
    "penjualan_dqmart_01-beta_normalized.csv"
)
