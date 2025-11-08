import pandas as pd
from datetime import datetime
import re

def normalize_tanggal_transaksi(input_xlsx_path: str, output_xlsx_path: str) -> None:
    # Baca file Excel
    df = pd.read_excel(input_xlsx_path, sheet_name='transaksi')

    def parse_date(date_str):
        if pd.isna(date_str):
            return date_str
        
        date_str = str(date_str).strip()

        # Hapus tanda kutip atau koma
        date_str = date_str.replace("‘", "'").replace("’", "'").replace(",", " ").replace(".", " ")
        date_str = re.sub(r"\s+", " ", date_str).strip()

        # Ubah bulan bahasa Indonesia → Inggris
        bulan_map = {
            "januari": "January", "jan": "Jan",
            "februari": "February", "feb": "Feb",
            "maret": "March", "mar": "Mar",
            "april": "April", "apr": "Apr",
            "mei": "May",
            "juni": "June", "jun": "Jun",
            "juli": "July", "jul": "Jul",
            "agustus": "August", "agu": "Aug",
            "september": "September", "sep": "Sep",
            "oktober": "October", "okt": "Oct",
            "november": "November", "nov": "Nov",
            "desember": "December", "des": "Dec"
        }

        for indo, eng in bulan_map.items():
            pattern = re.compile(rf"\b{indo}\b", re.IGNORECASE)
            date_str = pattern.sub(eng, date_str)

        # Deteksi tahun dua digit dengan tanda kutip ('24 → 2024)
        if re.search(r"'\d{2}", date_str):
            date_str = re.sub(r"'(\d{2})", lambda m: "20" + m.group(1), date_str)

        # Format tanggal yang mungkin
        possible_formats = [
            "%d %B %Y", "%d %b %Y", "%Y-%m-%d", "%d-%m-%Y",
            "%Y %d %B", "%Y %d %b", "%Y, %d %B", "%Y, %d %b",
            "%d %B", "%d %b %Y", "%d %B '%y", "%d %b '%y"
        ]

        for fmt in possible_formats:
            try:
                parsed = datetime.strptime(date_str, fmt)
                return parsed.strftime("%d-%m-%Y")
            except ValueError:
                continue

        # Coba gunakan pandas untuk format yang tidak terduga
        try:
            parsed = pd.to_datetime(date_str, errors='coerce', dayfirst=True)
            if pd.notna(parsed):
                return parsed.strftime("%d-%m-%Y")
        except Exception:
            pass

        # Jika gagal total
        return date_str

    # Terapkan ke kolom tanggal transaksi
    df['tanggal transaksi'] = df['tanggal transaksi'].apply(parse_date)

    # Simpan kembali ke Excel dengan urutan kolom tetap
    df.to_excel(output_xlsx_path, index=False)

# Contoh penggunaan:
normalize_tanggal_transaksi("penjualan_dqmart_01.xlsx", "penjualan_dqmart_01_output.xlsx")

