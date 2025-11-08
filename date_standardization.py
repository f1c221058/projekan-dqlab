import pandas as pd

def normalize_tanggal_transaksi(input_xlsx_path: str, output_xlsx_path: str) -> None:
    # Baca file Excel
    df = pd.read_excel(input_xlsx_path, sheet_name='transaksi', dtype=str)

    # Pemetaan bulan Indonesia dan Inggris
    bulan = {
        'januari': 'January', 'jan': 'January',
        'februari': 'February', 'feb': 'February',
        'maret': 'March', 'mar': 'March',
        'april': 'April', 'apr': 'April',
        'mei': 'May', 'may': 'May',
        'juni': 'June', 'jun': 'June',
        'juli': 'July', 'jul': 'July',
        'agustus': 'August', 'agu': 'August', 'aug': 'August',
        'september': 'September', 'sep': 'September',
        'oktober': 'October', 'okt': 'October', 'oct': 'October',
        'november': 'November', 'nov': 'November',
        'desember': 'December', 'des': 'December', 'dec': 'December'
    }

    def fix_date(x):
        if pd.isna(x):
            return x

        s = str(x).strip()
        # Bersihkan tanda baca dan kutipan
        for ch in [",", ".", "'", "‘", "’", "–", "-", "/", "\\"]:
            s = s.replace(ch, " ")
        s = " ".join(s.split())

        # Ubah nama bulan ke Inggris
        s = " ".join([bulan.get(w.lower(), w) for w in s.split()])

        # Tangani format "YYYY DD MMM" atau "YYYY, DD MMM"
        parts = s.split()
        if len(parts) == 3 and parts[0].isdigit() and len(parts[0]) == 4:
            y, d, m = parts
            s = f"{d} {m} {y}"

        # Tangani format tahun dua digit misal '24
        s = s.replace("‘", "").replace("’", "").strip()
        s = s.replace("'", "")

        # Coba parsing berbagai kemungkinan format
        for dayfirst in (True, False):
            dt = pd.to_datetime(s, errors='coerce', dayfirst=dayfirst, infer_datetime_format=True)
            if pd.notna(dt):
                return dt.strftime("%d-%m-%Y")

        # Jika gagal, coba paksa perbaikan dengan urutan lain
        try:
            dt = pd.to_datetime(s, format="%Y %m %d", errors='coerce')
            if pd.notna(dt):
                return dt.strftime("%d-%m-%Y")
        except Exception:
            pass

        # Jika tetap gagal, kembalikan nilai asli agar tidak hilang
        return x

    # Cari kolom yang mengandung tanggal
    for col in df.columns:
        sample = df[col].dropna().astype(str).head(20)
        if sample.apply(lambda x: pd.to_datetime(x, errors='coerce')).notna().sum() >= 3:
            df[col] = df[col].apply(fix_date)

    # Simpan hasil ke file output dengan format sama
    df.to_excel(output_xlsx_path, index=False, sheet_name='transaksi')
