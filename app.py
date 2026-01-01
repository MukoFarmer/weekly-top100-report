import pandas as pd

# ===== NORMALIZE FONKSİYONU =====
def normalize_text(x):
    if pd.isna(x):
        return None
    return str(x).strip().upper()


# ===== DOSYA YOLLARI =====
merchant_csv = r"C:\Users\mukocift\Desktop\merchant-name.csv"
report_excel = r"C:\Users\mukocift\Desktop\Report_OHL_ASIN_CHECK.xlsx"
output_excel = r"C:\Users\mukocift\Desktop\Report_OHL_ASIN_CHECK_OHL_FLAGGED.xlsx"

# ===== MERCHANT CSV OKU (;) =====
df_merch = pd.read_csv(merchant_csv, sep=";")
df_merch.columns = df_merch.columns.str.strip().str.upper()

# Merchant Name kolonu kontrol
if "MERCHANT NAME" not in df_merch.columns:
    raise ValueError(
        f"'MERCHANT NAME' sütunu bulunamadı. Mevcut kolonlar: {list(df_merch.columns)}"
    )

merchant_set = set(
    normalize_text(x)
    for x in df_merch["MERCHANT NAME"]
    if normalize_text(x) is not None
)

print(f"🔍 Merchant reference count: {len(merchant_set)}")

# ===== EXCEL TÜM SHEETLER =====
xls = pd.ExcelFile(report_excel)

with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(report_excel, sheet_name=sheet_name)

        # kolon normalize (eşleşme için)
        col_map = {c: c.strip().upper() for c in df.columns}

        # BRAND yoksa sheet'i aynen yaz
        if "BRAND" not in col_map.values():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            continue

        # gerçek BRAND kolon adını bul
        brand_col = [k for k, v in col_map.items() if v == "BRAND"][0]

        normalized_brands = df[brand_col].apply(normalize_text)

        status = normalized_brands.apply(
            lambda x: "OHL" if x in merchant_set else "NOT OHL"
        )

        # BRAND yanına ekle
        idx = df.columns.get_loc(brand_col)
        df.insert(idx + 1, "OHL_STATUS", status)

        df.to_excel(writer, sheet_name=sheet_name, index=False)

print("✅ Brand ↔ Merchant Name kontrolü tamamlandı")
print("📁 Çıktı:", output_excel)
