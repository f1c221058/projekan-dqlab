import pandas as pd
from mlxtend.frequent_patterns import apriori, association_rules


def run_analysis(input_xlsx_path: str, output_xlsx_path: str) -> None:
    """
    Fungsi utama yang menjalankan seluruh proses analisis pola pembelian.
    Fungsi ini membaca data transaksi dari Excel, melakukan analisis dengan algoritma Apriori,
    membentuk aturan asosiasi (association rules), dan menyimpan hasil akhirnya
    dalam file Excel bernama 'product_packaging.xlsx' sesuai ketentuan soal.
    """

    # =====================================================================
    # 1. Membaca data transaksi dari sheet "Transaksi"
    # =====================================================================
    # Kita mulai dengan membuka file Excel yang berisi data transaksi.
    # Soal menyebutkan bahwa file memiliki kolom: Kode Transaksi, Nama Produk, dan Jumlah.
    # Kolom "Jumlah" tidak digunakan dalam analisis, jadi akan diabaikan.
    # =====================================================================
    df = pd.read_excel(input_xlsx_path, sheet_name="Transaksi")

    # Pastikan dua kolom penting ada dalam data
    if not {"Kode Transaksi", "Nama Produk"}.issubset(df.columns):
        raise ValueError("Kolom 'Kode Transaksi' dan 'Nama Produk' wajib ada di sheet 'Transaksi'.")

    # Kita hanya ambil dua kolom yang diperlukan, lalu hapus baris kosong
    df = df[["Kode Transaksi", "Nama Produk"]].dropna()

    # Bersihkan data dari spasi dan pastikan formatnya string
    df["Kode Transaksi"] = df["Kode Transaksi"].astype(str).str.strip()
    df["Nama Produk"] = df["Nama Produk"].astype(str).str.strip()

    # Jika ternyata data kosong setelah dibersihkan, buat file kosong berheader
    if df.empty:
        pd.DataFrame(columns=["Packaging Set ID", "Products", "Maximum Lift", "Maximum Confidence"])\
            .to_excel(output_xlsx_path, index=False)
        return

    # =====================================================================
    # 2. Mengubah data transaksi menjadi format basket (matrix transaksi × produk)
    # =====================================================================
    # Setiap baris transaksi akan dikonversi ke bentuk "one-hot encoded"
    # di mana setiap kolom mewakili produk tertentu.
    # Nilai 1 berarti produk tersebut dibeli dalam transaksi itu, 0 berarti tidak.
    # =====================================================================
    basket = (
        df.groupby(["Kode Transaksi", "Nama Produk"])["Nama Produk"]
        .count().unstack(fill_value=0)
    )

    # Konversi ke tipe boolean (True/False) agar sesuai dengan format yang diharapkan oleh mlxtend
    basket = (basket > 0).astype(int)

    # Jika matrix basket kosong, buat file kosong saja agar tidak error
    if basket.empty:
        pd.DataFrame(columns=["Packaging Set ID", "Products", "Maximum Lift", "Maximum Confidence"])\
            .to_excel(output_xlsx_path, index=False)
        return

    # =====================================================================
    # 3. Menjalankan algoritma Apriori
    # =====================================================================
    # Apriori digunakan untuk mencari kombinasi produk yang sering muncul bersama.
    # Parameter:
    #   - min_support = 0.05 → artinya kombinasi harus muncul di minimal 5% transaksi.
    #   - use_colnames=True agar hasilnya tetap memakai nama produk asli, bukan indeks kolom.
    # =====================================================================
    frequent_itemsets = apriori(basket, min_support=0.05, use_colnames=True)

    # Jika tidak ada kombinasi yang memenuhi syarat support minimal, hasil kosong
    if frequent_itemsets.empty:
        pd.DataFrame(columns=["Packaging Set ID", "Products", "Maximum Lift", "Maximum Confidence"])\
            .to_excel(output_xlsx_path, index=False)
        return

    # =====================================================================
    # 4. Membentuk aturan asosiasi (Association Rules)
    # =====================================================================
    # Setelah kombinasi produk ditemukan, kita ingin tahu hubungan antarproduk.
    # Fungsi 'association_rules' menghasilkan metrik seperti lift dan confidence.
    # Kita gunakan metrik 'confidence' dengan ambang minimal 0.4 (40%) seperti di soal.
    # =====================================================================
    rules = association_rules(frequent_itemsets, metric="confidence", min_threshold=0.4)

    # Jika tidak ada aturan yang memenuhi, buat file kosong saja
    if rules.empty:
        pd.DataFrame(columns=["Packaging Set ID", "Products", "Maximum Lift", "Maximum Confidence"])\
            .to_excel(output_xlsx_path, index=False)
        return

    # =====================================================================
    # 5. Menyatukan antecedents dan consequents menjadi satu set produk
    # =====================================================================
    # Setiap aturan memiliki dua sisi:
    #   - antecedents: produk awal (misalnya Product A)
    #   - consequents: produk yang cenderung dibeli bersamaan (misalnya Product B)
    # Kita gabungkan keduanya agar merepresentasikan satu paket produk lengkap.
    # =====================================================================
    def union_sets(row):
        return row["antecedents"].union(row["consequents"])

    rules["product_set"] = rules.apply(union_sets, axis=1)

    # =====================================================================
    # 6. Menyaring kombinasi produk tunggal (karena paket minimal 2 produk)
    # =====================================================================
    rules = rules[rules["product_set"].apply(len) >= 2]

    # Jika hasil kosong, buat file kosong
    if rules.empty:
        pd.DataFrame(columns=["Packaging Set ID", "Products", "Maximum Lift", "Maximum Confidence"])\
            .to_excel(output_xlsx_path, index=False)
        return

    # =====================================================================
    # 7. Mengurutkan nama produk dalam setiap kombinasi dan membuat string gabungan
    # =====================================================================
    # Setiap kombinasi produk diurutkan berdasarkan abjad untuk menghindari duplikasi palsu.
    # Contoh: {Product 20, Product 4, Product 25} akan diubah menjadi "Product 4;Product 20;Product 25".
    # =====================================================================
    rules["Products"] = rules["product_set"].apply(lambda s: ";".join(sorted(s)))

    # =====================================================================
    # 8. Menggabungkan kombinasi duplikat
    # =====================================================================
    # Bisa jadi ada beberapa aturan berbeda yang menghasilkan kombinasi produk yang sama.
    # Maka, kita gabungkan berdasarkan kolom 'Products' dan ambil nilai lift & confidence tertinggi.
    # =====================================================================
    agg = (
        rules.groupby("Products")
        .agg(Maximum_Lift=("lift", "max"), Maximum_Confidence=("confidence", "max"))
        .reset_index()
    )

    # =====================================================================
    # 9. Mengurutkan hasil akhir berdasarkan lift tertinggi, lalu confidence tertinggi
    # =====================================================================
    # Kombinasi produk dengan lift lebih besar berarti asosiasinya lebih kuat,
    # sehingga ditampilkan paling atas dalam file hasil.
    # =====================================================================
    agg = agg.sort_values(by=["Maximum_Lift", "Maximum_Confidence"], ascending=[False, False]).reset_index(drop=True)

    # =====================================================================
    # 10. Menambahkan kolom Packaging Set ID
    # =====================================================================
    # Kolom ini adalah nomor urut (1, 2, 3, ...) untuk memudahkan pembaca non-teknis.
    # =====================================================================
    agg.insert(0, "Packaging Set ID", range(1, len(agg) + 1))

    # =====================================================================
    # 11. Menyimpan hasil akhir ke file Excel
    # =====================================================================
    # File akhir bernama "product_packaging.xlsx" seperti yang diminta soal.
    # =====================================================================
    agg.to_excel(output_xlsx_path, index=False)


# =====================================================================
# Pemanggilan fungsi utama (contoh penggunaan)
# =====================================================================
# Saat file dijalankan langsung (bukan diimpor), fungsi ini otomatis dijalankan.
# Dalam sistem penilaian, fungsi ini juga akan dipanggil dengan cara yang sama:
# run_analysis("transaksi_dqmart.xlsx", "product_packaging.xlsx")
# =====================================================================
if __name__ == "__main__":
    run_analysis("transaksi_dqmart.xlsx", "product_packaging.xlsx")
