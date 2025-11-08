import pandas as pd
import re

# Fungsi utama untuk menstandarkan penulisan tanggal pada file Excel transaksi
def normalize_tanggal_transaksi(input_xlsx_path: str, output_xlsx_path: str) -> None:
    # Baca file Excel sesuai instruksi (sheet: transaksi)
    df = pd.read_excel(input_xlsx_path, sheet_name='transaksi', dtype=str)

    # Mapping nama bulan dalam Bahasa Indonesia ke Bahasa Inggris
    # supaya bisa dikenali oleh pandas.to_datetime()
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

    # Fungsi kecil untuk membersihkan dan menormalkan format tanggal
    def parse_date(val):
        if pd.isna(val):
            return val
        s = str(val).strip()

        # Hapus karakter yang sering bikin parsing gagal
        s = re.sub(r"[,’‘’]", " ", s)
        s = s.replace(",", " ")
        s = re.sub(r"\s+", " ", s)

        # Ubah nama bulan lokal (Indonesia) ke format Inggris
        s = " ".join([bulan_map.get(w.lower(), w) for w in s.split()])

        # Ubah tahun dua digit seperti '24 menjadi 2024
        s = re.sub(r"'\s*(\d{2})(?!\d)", r"20\1", s)

        # Tangani pola seperti "2024, 6 Nov" agar jadi "6 Nov 2024"
        match = re.match(r"^(\d{4})\s*,\s*(\d{1,2})\s+([A-Za-z]+)", s)
        if match:
            s = f"{match.group(2)} {match.group(3)} {match.group(1)}"

        # Coba parse dengan dua kemungkinan format (day first dan month first)
        for dayfirst in (True, False):
            dt = pd.to_datetime(s, errors='coerce', dayfirst=dayfirst, infer_datetime_format=True)
            if pd.notna(dt):
                return dt.strftime("%d-%m-%Y")

        # Kalau gagal diparse, kembalikan nilai aslinya
        return val

    # Deteksi otomatis kolom yang isinya tanggal
    # (dicek berdasarkan 10 sampel pertama di tiap kolom)
    for col in df.columns:
        sample = df[col].dropna().astype(str).head(10)
        if pd.to_datetime(sample, errors='coerce', dayfirst=True).notna().sum() >= 3:
            df[col] = df[col].apply(parse_date)

    # Simpan hasil ke Excel baru dengan struktur kolom tetap sama
    df.to_excel(output_xlsx_path, index=False, sheet_name='transaksi')

    print(f"✅ Standarisasi tanggal selesai. File hasil: {output_xlsx_path}")


# Eksekusi utama
if __name__ == "__main__":
    # File input dan output sesuai instruksi hackathon
    normalize_tanggal_transaksi(
        "penjualan_dqmart_01.xlsx",
        "penjualan_dqmart_01_normalized.xlsx"
    )
