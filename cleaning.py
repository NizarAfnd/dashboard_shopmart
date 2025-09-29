import pandas as pd

# 1. Load data mentah
fact = pd.read_excel(r"D:\Magang\Keuangan UNUSA\konversi visual\fact_table.xlsx")
item = pd.read_excel(r"D:\Magang\Keuangan UNUSA\konversi visual\item_dim.xlsx")
store = pd.read_excel(r"D:\Magang\Keuangan UNUSA\konversi visual\store_dim.xlsx")
trans = pd.read_excel(r"D:\Magang\Keuangan UNUSA\konversi visual\Trans_dim.xlsx")
time = pd.read_excel(r"D:\Magang\Keuangan UNUSA\konversi visual\time_dim.xlsx")
customer = pd.read_excel(r"D:\Magang\Keuangan UNUSA\konversi visual\customer_dim.xlsx")

# =======================
# 2. CLEANING
# =======================

## a. Transaksi (trans)
# Isi nilai bank_name yang kosong (misalnya untuk cash) dengan 'N/A'
trans['bank_name'] = trans['bank_name'].fillna('N/A')
# Normalisasi huruf kecil ke huruf kapital untuk trans_type
trans['trans_type'] = trans['trans_type'].str.strip().str.capitalize()

## b. Customer
# Pastikan NID unik (hapus duplikat)
customer = customer.drop_duplicates(subset=['nid'])
# Hilangkan spasi berlebih pada nama
customer['name'] = customer['name'].str.strip().str.title()

## c. Fact Table
# Pastikan total_price benar = quantity * unit_price
fact['total_price_check'] = fact['quantity'] * fact['unit_price']
# Jika ada perbedaan, isi ulang dengan nilai yang benar
fact.loc[fact['total_price'] != fact['total_price_check'], 'total_price'] = fact['total_price_check']
fact = fact.drop(columns=['total_price_check'])

# Normalisasi unit (lowercase semua biar seragam)
fact['unit'] = fact['unit'].str.lower().str.strip()

## d. Item
# Bersihkan kategori produk (desc), hilangkan prefix "a. "
item['desc'] = item['desc'].str.replace(r'^[aA]\.\s*', '', regex=True).str.strip()
# Normalisasi negara: huruf pertama kapital
item['man_country'] = item['man_country'].str.title()
# Normalisasi supplier
item['supplier'] = item['supplier'].str.title()

## e. Store
# Normalisasi nama lokasi (title case)
store['division'] = store['division'].str.title()
store['district'] = store['district'].str.title()
store['upazila'] = store['upazila'].str.title()

## f. Time
# Pisahkan datetime jadi tanggal dan jam
time['date'] = pd.to_datetime(time['date'], errors='coerce', dayfirst=True)
time['hour'] = pd.to_numeric(time['hour'], errors='coerce')
time['day'] = pd.to_numeric(time['day'], errors='coerce')

# =======================
# 3. HASIL
# =======================

print("Cleaning selesai ✅")
print("Fact table preview:")
print(fact.head())

# Opsional: Simpan hasil cleaning ke file Excel baru
with pd.ExcelWriter(r"D:\Magang\Keuangan UNUSA\konversi visual\cleaned_data.xlsx") as writer:
    fact.to_excel(writer, sheet_name="fact_clean", index=False)
    item.to_excel(writer, sheet_name="item_clean", index=False)
    store.to_excel(writer, sheet_name="store_clean", index=False)
    trans.to_excel(writer, sheet_name="trans_clean", index=False)
    time.to_excel(writer, sheet_name="time_clean", index=False)
    customer.to_excel(writer, sheet_name="customer_clean", index=False)
