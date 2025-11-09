import pandas as pd
from mlxtend.frequent_patterns import apriori, association_rules

def run_analysis(input_xlsx_path: str, output_xlsx_path: str) -> None:
    # Baca data transaksi
    df = pd.read_excel(input_xlsx_path, sheet_name="Transaksi")
    df.columns = df.columns.str.strip().str.lower()

    # Cek kolom
    if not {'kode transaksi', 'nama produk', 'jumlah'}.issubset(df.columns):
        raise ValueError("Kolom wajib: 'Kode Transaksi', 'Nama Produk', dan 'Jumlah'.")

    # Ubah jadi format basket (produk dibeli = 1)
    basket = (df
        .pivot_table(index='kode transaksi', columns='nama produk', values='jumlah', fill_value=0)
        .gt(0).astype(int)
    )

    # Apriori + association rules
    rules = association_rules(
        apriori(basket, min_support=0.05, use_colnames=True),
        metric="confidence", min_threshold=0.4
    )

    # Gabungkan antecedent + consequent → kombinasi produk unik
    rules['Products'] = rules.apply(
        lambda r: ';'.join(sorted(set(r['antecedents']) | set(r['consequents']))),
        axis=1
    )

    # Ambil kombinasi unik dengan lift & confidence maksimum
    result = (rules.groupby('Products', as_index=False)
        .agg(Maximum_Lift=('lift', 'max'), Maximum_Confidence=('confidence', 'max'))
        .sort_values(['Maximum_Lift', 'Maximum_Confidence'], ascending=[False, False])
        .reset_index(drop=True)
    )

    # Tambahkan Packaging Set ID
    result.insert(0, 'Packaging Set ID', range(1, len(result) + 1))

    # Simpan ke Excel
    result.to_excel(output_xlsx_path, sheet_name="Packaging", index=False)

if __name__ == "__main__":
    run_analysis("transaksi_dqmart.xlsx","product_packaging.xlsx")