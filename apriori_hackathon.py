# apriori_hackathon.py
import pandas as pd
from mlxtend.frequent_patterns import apriori, association_rules


def run_analysis(input_xlsx_path: str, output_xlsx_path: str) -> None:
    # 1. Baca data
    df = pd.read_excel(input_xlsx_path, sheet_name="Transaksi")
    df = df.rename(columns=lambda x: x.strip().lower())
    df = df.rename(columns={"kode transaksi": "Kode Transaksi", "nama produk": "Nama Produk"})
    df = df[["Kode Transaksi", "Nama Produk"]].dropna().drop_duplicates()
    df["Nama Produk"] = df["Nama Produk"].astype(str).str.strip()

    # 2. Buat basket (0/1)
    basket = pd.crosstab(df["Kode Transaksi"], df["Nama Produk"]).astype(bool).astype(int)

    # 3. Jalankan Apriori + Association Rules
    itemsets = apriori(basket, min_support=0.05, use_colnames=True)
    rules = association_rules(itemsets, metric="confidence", min_threshold=0.4)

    if rules.empty:
        print("Tidak ada kombinasi yang memenuhi kriteria.")
        return

    # 4. Gabungkan antecedents dan consequents
    rules["union_set"] = rules.apply(lambda r: frozenset(r["antecedents"].union(r["consequents"])), axis=1)
    grouped = (
        rules.groupby(rules["union_set"])
        .agg(Maximum_Lift=("lift", "max"), Maximum_Confidence=("confidence", "max"))
        .reset_index()
    )

    # 5. Format output
    grouped["Products"] = grouped["union_set"].apply(lambda x: ";".join(sorted(x)))
    grouped = grouped[["Products", "Maximum_Lift", "Maximum_Confidence"]]
    grouped = grouped.sort_values(by=["Maximum_Lift", "Maximum_Confidence"], ascending=[False, False])
    grouped.insert(0, "Packaging Set ID", range(1, len(grouped) + 1))
    grouped.to_excel(output_xlsx_path, sheet_name = 'Packaging', index=False)
    print(f"[OK] Hasil disimpan ke: {output_xlsx_path}")


# Contoh pemanggilan langsung:
run_analysis("transaksi_dqmart.xlsx", "product_packaging.xlsx")