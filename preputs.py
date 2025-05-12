import pandas as pd
from datetime import datetime

# Buka file Excel
file_path = "D:\Semester 6\Data Mining\gcommerce\dataset.xlsx"
df = pd.read_excel(file_path, sheet_name="Laporan orders")

timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
file_name = f"dataset_bersih_{timestamp}.xlsx"

print(df.head())
print(df.info())

df.drop(columns=[
    'NO. RESI', 
    'STATUS PESANAN', 
    'INSTANT CHECKOUT', 
    'BIAYA PLATFORM'  # karena semua NaN
], inplace=True)

df.drop_duplicates(inplace=True)

harga_cols = [
    'HARGA PRODUK SATUAN', 
    'HARGA TOTAL PRODUK DIBELI', 
    'BIAYA ADMIN'
]

for col in harga_cols:
    df[col] = df[col].str.replace("Rp.", "", regex=False)
    df[col] = df[col].str.replace(".", "", regex=False)
    df[col] = pd.to_numeric(df[col], errors="coerce")

df['HARGA PRODUK SATUAN'] = pd.to_numeric(df['HARGA PRODUK SATUAN'], errors='coerce')
df['JUMLAH PRODUK DIBELI'] = pd.to_numeric(df['JUMLAH PRODUK DIBELI'], errors='coerce')
df['TANGGAL PEMBELIAN'] = pd.to_datetime(df['TANGGAL PEMBELIAN'], errors='coerce')
df['TANGGAL PROSES SELLER'] = pd.to_datetime(df['TANGGAL PROSES SELLER'], errors='coerce')

df[df['TANGGAL PROSES SELLER'].isnull()]

df.dropna(inplace=True)

print(df.isnull().sum())      # cek nilai kosong
print(df.dtypes)              # cek tipe data

df.to_excel(file_name, index=False)
