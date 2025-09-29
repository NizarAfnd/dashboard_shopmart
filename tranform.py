import pandas as pd

# 1. Load hasil cleaning
fact = pd.read_excel(r"D:\Magang\Keuangan UNUSA\konversi visual\cleaned_data.xlsx", sheet_name="fact_clean")
item = pd.read_excel(r"D:\Magang\Keuangan UNUSA\konversi visual\cleaned_data.xlsx", sheet_name="item_clean")
store = pd.read_excel(r"D:\Magang\Keuangan UNUSA\konversi visual\cleaned_data.xlsx", sheet_name="store_clean")
trans = pd.read_excel(r"D:\Magang\Keuangan UNUSA\konversi visual\cleaned_data.xlsx", sheet_name="trans_clean")
time = pd.read_excel(r"D:\Magang\Keuangan UNUSA\konversi visual\cleaned_data.xlsx", sheet_name="time_clean")

# 2. Join tabel
df = (fact
      .merge(item, on="item_key", how="left")
      .merge(store, on="store_key", how="left")
      .merge(trans, on="payment_key", how="left")
      .merge(time, on="time_key", how="left"))

# 3. Fungsi untuk memisahkan kategori produk
def split_category(cat):
    if pd.isna(cat):
        return None, None
    parts = cat.split("-")
    if len(parts) == 2:
        return parts[0].strip(), parts[1].strip()
    else:
        words = cat.split()
        if len(words) >= 2:
            return words[0].strip(), words[1].strip()
        else:
            return words[0].strip(), words[0].strip()

# Terapkan ke kolom 'desc'
df[["Kategori Produk Baru", "Nama Produk Baru"]] = df["desc"].apply(lambda x: pd.Series(split_category(x)))

# 4. Buat kolom nama bulan dari Tanggal Transaksi
df["Bulan"] = pd.to_datetime(df["date"], errors="coerce").dt.strftime('%B')

# 5. Pilih kolom sesuai kebutuhan analitik
df_final = pd.DataFrame({
    "Tanggal Transaksi": pd.to_datetime(df["date"], errors="coerce"),
    "Nama Produk": df["Nama Produk Baru"],
    "Kategori Produk": df["Kategori Produk Baru"],
    "Jumlah Item Terjual": pd.to_numeric(df["quantity"], errors="coerce"),
    "Harga Per Item": pd.to_numeric(df["unit_price_x"], errors="coerce"),
    "Negara": "Bangladesh",
    "Provinsi": df["division"],
    "Bulan": df["Bulan"],
    "Jam": df["hour"],
    "Hari": df["day"],
    "Metode Pembayaran": df["trans_type"]
})

# 6. Simpan ke 1 file Excel khusus hasil transformasi
df_final.to_excel(r"D:\Magang\Keuangan UNUSA\konversi visual\dashboard.xlsx", index=False)

print("✅ Transformasi selesai!")
