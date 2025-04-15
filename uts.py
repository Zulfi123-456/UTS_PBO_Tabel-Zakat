import pandas as pd
from datetime import datetime
import os

# Inisialisasi data
harga_beras = [15000, 20000, 25000]
EXCEL_FILE = "data_zakat.xlsx"

def load_data_from_excel():
    global data_zakat
    try:
        if os.path.exists(EXCEL_FILE):
            df = pd.read_excel(EXCEL_FILE)
            data_zakat = df.to_dict('records')
        else:
            data_zakat = [
                {
                    'NIK': '2322001',
                    'Nama': 'Mahdi',
                    'Tanggal': '03/03/2025',
                    'Jumlah beras': 2,
                    'Jumlah yang harus dibayar': 50000,
                    'Jumlah uang': 60000,
                    'kembalian': 10000
                }
            ]
            save_to_excel()
    except Exception as e:
        print(f"Error loading data: {str(e)}")
        data_zakat = []

def save_to_excel():
    try:
        df = pd.DataFrame(data_zakat)
        with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Data Zakat')
            
            # Mengatur lebar kolom
            worksheet = writer.sheets['Data Zakat']
            for idx, col in enumerate(df.columns):
                max_length = max(df[col].astype(str).apply(len).max(), len(col)) + 2
                worksheet.column_dimensions[chr(65 + idx)].width = max_length
        return True
    except Exception as e:
        print(f"Error saving data: {str(e)}")
        return False

def tampilkan_harga_beras():
    print("\nDaftar Harga Beras:")
    for i, harga in enumerate(harga_beras, 1):
        print(f"{i}. Rp {harga:,}")

def input_harga_beras():
    try:
        harga_baru = int(input("\nMasukkan harga beras baru: Rp "))
        harga_beras.append(harga_baru)
        print(f"Harga beras Rp {harga_baru:,} berhasil ditambahkan!")
    except ValueError:
        print("Error: Mohon masukkan angka yang valid!")

def tampilkan_data_zakat():
    if not data_zakat:
        print("\nBelum ada data zakat yang tersimpan.")
        return
    
    print("\nData Pembayaran Zakat:")
    print("NIK\t\tNama\t\tTanggal Bayar\tBeras/Liter\tJumlah\t\tBayar\t\tKembalian")
    print("-" * 100)
    for data in data_zakat:
        print(f"{data['NIK']}\t{data['Nama']}\t{data['Tanggal']}\t{data['Jumlah beras']}\t\tRp {data['Jumlah yang harus dibayar']:,}\tRp {data['Jumlah uang']:,}\tRp {data['kembalian']:,}")

def pembayaran_zakat():
    print("\nForm Pembayaran Zakat")
    nik = input("Masukkan NIK: ")
    nama = input("Masukkan Nama: ")
    tanggal = input("Masukkan Tanggal (DD/MM/YYYY): ")
    
    print("\nPilih harga beras:")
    tampilkan_harga_beras()
    try:
        pilihan = int(input("Pilih nomor harga beras: ")) - 1
        if 0 <= pilihan < len(harga_beras):
            harga = harga_beras[pilihan]
            beras = float(input("Masukkan jumlah beras (liter): "))
            jumlah = harga * beras
            
            print(f"\nTotal zakat yang harus dibayar: Rp {jumlah:,}")
            bayar = float(input("Masukkan jumlah pembayaran: Rp "))
            
            kembalian = bayar - jumlah
            if kembalian < 0:
                print(f"\nPembayaran kurang sebesar: Rp {abs(kembalian):,}")
                return
            
            data_baru = {
                'NIK': nik,
                'Nama': nama,
                'Tanggal': tanggal,
                'Jumlah beras': beras,
                'Jumlah yang harus dibayar': jumlah,
                'Jumlah uang': bayar,
                'kembalian': kembalian
            }
            
            data_zakat.append(data_baru)
            if save_to_excel():
                print(f"\nPembayaran zakat berhasil!")
                print(f"Kembalian: Rp {kembalian:,}")
                print(f"Data telah disimpan ke {EXCEL_FILE}")
            else:
                print("\nPembayaran berhasil tetapi gagal menyimpan ke Excel!")
        else:
            print("Pilihan tidak valid!")
    except ValueError:
        print("Error: Mohon masukkan angka yang valid!")

def export_excel():
    if not data_zakat:
        print("\nBelum ada data zakat yang tersimpan.")
        return
    
    if save_to_excel():
        print(f"\nData berhasil disimpan ke {EXCEL_FILE}")
    else:
        print("\nGagal menyimpan data ke Excel!")

def main():
    # Load data dari Excel saat program dimulai
    load_data_from_excel()
    
    while True:
        print("\n=== MENU ZAKAT ===")
        print("1. Tampilkan harga beras")
        print("2. Input harga beras")
        print("3. Tampilkan data zakat")
        print("4. Pembayaran zakat")
        print("5. Export Excel")
        print("6. Keluar")
        
        try:
            pilihan = int(input("\nPilih menu (1-6): "))
            
            if pilihan == 1:
                tampilkan_harga_beras()
            elif pilihan == 2:
                input_harga_beras()
            elif pilihan == 3:
                tampilkan_data_zakat()
            elif pilihan == 4:
                pembayaran_zakat()
            elif pilihan == 5:
                export_excel()
            elif pilihan == 6:
                print("\nTerima kasih telah menggunakan program ini!")
                break
            else:
                print("Pilihan tidak valid! Silakan pilih 1-6.")
        except ValueError:
            print("Error: Mohon masukkan angka yang valid!")

if __name__ == "_main_":
    main()