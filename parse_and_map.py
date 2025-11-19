import pandas as pd
import re
from datetime import datetime

def extract_date_from_keterangan(keterangan):
    # Menyusun pola regex untuk mencocokkan tanggal dalam berbagai format, termasuk koma dan spasi
    date_pattern = r'(\d{4}),?\s*(\d{1,2})\s*([A-Za-z]+|[A-Za-z]{3})|(\d{1,2})\s([A-Za-z]+|[A-Za-z]{3})\s(\d{2,4}|\'\d{2})'
    
    # Mencocokkan regex
    match = re.search(date_pattern, keterangan)
    
    if match:
        # Jika ada tahun di awal
        if match.group(1):  # Format "2024, 27 Aug"
            year = match.group(1)
            day = match.group(2)
            month_name = match.group(3)
        # Jika tahun di belakang (misal format '25 Jun '24')
        else:  # Format '25 Jun '24' atau '7 August 2024'
            day = match.group(4)
            month_name = match.group(5)
            year = match.group(6)
            if year.startswith("'"):  # Jika tahun dua digit
                year = '20' + year[1:]
        
        # Menyusun nama bulan dengan angka
        month_dict = {
            "Januari": 1, "Februari": 2, "Maret":3, "April": 4,
            "Mei": 5, "June": 6, "Jul": 7, "Agustus": 8,
            "September": 9, "Oktober": 10, "November": 11, "Desember": 12,
            "Aug": 8, "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12
            ,"Okt":10,"Mar":3,"August":7,"Juni":6 # Untuk singkatan bulan
        }
        
        if month_name in month_dict:
            month = month_dict[month_name]
            # Mengonversi tanggal menjadi format dd-mm-yyyy
            date_string = f"{day}-{month:02d}-{year}"
            return date_string
    return None

def process_excel(input_path, output_path, sheet_name='transaksi'):
    # Baca file Excel input dan pilih sheet yang sesuai dengan sheet_name
    df = pd.read_excel(input_path, sheet_name=sheet_name)
    
    # Tambahkan kolom 'Tanggal Transaksi' dengan hasil ekstraksi dari kolom 'Keterangan'
    df['Tanggal Transaksi'] = df['Keterangan'].apply(extract_date_from_keterangan)
    
    # Simpan hasil ke file output dengan nama sheet yang ditentukan
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)

# Pemanggilan fungsi dengan sheet_name
if __name__ == "__main__":
    process_excel(input_path="transaksi-raw-partial.xlsx", output_path="transaksi-output.xlsx", sheet_name='transaksi')
