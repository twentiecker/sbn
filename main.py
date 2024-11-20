import pandas as pd
import os

# def process_and_merge_sheets(input_file, output_file):
#     try:
#         # Membaca semua sheet dalam file
#         excel_data = pd.ExcelFile(input_file)
        
#         # Dictionary untuk menyimpan data berdasarkan triwulan
#         triwulan_data = {
#             "triwulan1": [],
#             "triwulan2": [],
#             "triwulan3": [],
#             "triwulan4": []
#         }
        
#         for sheet_name in excel_data.sheet_names:
#             print(f"Memproses sheet: {sheet_name}")
            
#             # Ambil bagian sebelum tanda "-"
#             prefix = sheet_name.split("-")[0]
            
#             # Membaca kolom A-H dari setiap sheet
#             sheet_data = pd.read_excel(input_file, sheet_name=sheet_name, usecols='A:H')
            
#             try:
#                 # Menyeleksi baris di mana kolom 'Unnamed: 2' dan 'Unnamed: 3' tidak kosong
#                 filtered_data = sheet_data.dropna(subset=['Unnamed: 2', 'Unnamed: 3'])
#             except:
#                 continue

#             # Jika ada data setelah difilter
#             if not filtered_data.empty:
#                 # Menggunakan baris pertama sebagai header
#                 filtered_data.columns = filtered_data.iloc[0]
#                 filtered_data = filtered_data[1:]  # Menghapus baris pertama setelah menjadi header
#                 filtered_data.reset_index(drop=True, inplace=True)  # Reset indeks
                
#                 # Tentukan triwulan berdasarkan prefix
#                 if prefix in ('1', '2', '3', '01', '02', '03'):
#                     print(f"Sheet {sheet_name} masuk ke triwulan1")
#                     triwulan_data["triwulan1"].append(filtered_data)
#                 elif prefix in ('4', '5', '6', '04', '05', '06'):
#                     print(f"Sheet {sheet_name} masuk ke triwulan2")
#                     triwulan_data["triwulan2"].append(filtered_data)
#                 elif prefix in ('7', '8', '9', '07', '08', '09',):
#                     print(f"Sheet {sheet_name} masuk ke triwulan3")
#                     triwulan_data["triwulan3"].append(filtered_data)
#                 elif prefix in ('10', '11', '12'):
#                     print(f"Sheet {sheet_name} masuk ke triwulan4")
#                     triwulan_data["triwulan4"].append(filtered_data)
#                 else:
#                     print(f"Prefix {prefix} tidak dikenali untuk sheet {sheet_name}")
#             else:
#                 print(f"Sheet {sheet_name} kosong setelah difilter.")
        
#         # Menulis hasil gabungan ke file Excel baru
#         with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
#             for triwulan, data_frames in triwulan_data.items():
#                 if data_frames:
#                     # Menggabungkan semua DataFrame dalam satu triwulan
#                     merged_data = pd.concat(data_frames, ignore_index=True)
#                     merged_data.to_excel(writer, sheet_name=triwulan, index=False)
#                 else:
#                     print(f"Tidak ada data untuk {triwulan}.")
        
#         print(f"Data yang difilter dan digabungkan telah disimpan ke: {output_file}")
        
#     except Exception as e:
#         print(f"Terjadi kesalahan: {e}")

# def open_excel_file(file_base_name):
#     """
#     Mencoba membuka file dengan ekstensi .xlsx terlebih dahulu, jika gagal coba .xls.
#     """
#     # Tentukan nama file dengan kedua ekstensi
#     file_xlsx = f"{file_base_name}.xlsx"
#     file_xls = f"{file_base_name}.xls"
    
#     try:
#         # Coba buka file dengan ekstensi .xlsx
#         if os.path.exists(file_xlsx):
#             print(f"Mencoba membuka file: {file_xlsx}")
#             return pd.ExcelFile(file_xlsx), file_xlsx
#         elif os.path.exists(file_xls):
#             # Jika file .xlsx tidak ada, coba .xls
#             print(f"Mencoba membuka file: {file_xls}")
#             return pd.ExcelFile(file_xls), file_xls
#         else:
#             # Jika kedua file tidak ditemukan
#             raise FileNotFoundError(f"Tidak ditemukan file dengan nama {file_base_name}.xlsx atau {file_base_name}.xls")
#     except Exception as e:
#         # Tangani kesalahan saat membuka file
#         print(f"Terjadi kesalahan saat membuka file: {e}")
#         raise

# # File input dan output
# tahun = '2024'
# input_file, full_file_name = open_excel_file(f"sbn-{tahun}")
# output_file = f"sbn-{tahun}-triwulan.xlsx"  # Ganti dengan path file output yang diinginkan

# process_and_merge_sheets(input_file, output_file)

# =====================================================================================================
# =====================================================================================================

# def process_and_merge_sheets(input_file, output_file):
#     try:
#         # Membaca semua sheet dalam file
#         excel_data = pd.ExcelFile(input_file)
        
#         # Dictionary untuk menyimpan data berdasarkan triwulan
#         triwulan_data = {
#             "triwulan1": [],
#             "triwulan2": [],
#             "triwulan3": [],
#             "triwulan4": []
#         }
        
#         for sheet_name in excel_data.sheet_names:
#             print(f"Memproses sheet: {sheet_name}")
            
#             # Ambil bagian sebelum tanda "-"
#             prefix = sheet_name.split("-")[0]
            
#             # Membaca kolom A-H dari setiap sheet
#             sheet_data = pd.read_excel(input_file, sheet_name=sheet_name, usecols='A:H')
            
#             try:
#                 # Menyeleksi baris di mana kolom 'Unnamed: 2' dan 'Unnamed: 3' tidak kosong
#                 filtered_data = sheet_data.dropna(subset=['Unnamed: 2', 'Unnamed: 3'])
#             except:
#                 continue

#             # Jika ada data setelah difilter
#             if not filtered_data.empty:
#                 # Menggunakan baris pertama sebagai header
#                 filtered_data.columns = filtered_data.iloc[0]
#                 filtered_data = filtered_data[1:]  # Menghapus baris pertama setelah menjadi header
#                 filtered_data.reset_index(drop=True, inplace=True)  # Reset indeks
                
#                 # Tentukan triwulan berdasarkan prefix
#                 if prefix in ('1', '2', '3', '01', '02', '03'):
#                     print(f"Sheet {sheet_name} masuk ke triwulan1")
#                     triwulan_data["triwulan1"].append(filtered_data)
#                 elif prefix in ('4', '5', '6', '04', '05', '06'):
#                     print(f"Sheet {sheet_name} masuk ke triwulan2")
#                     triwulan_data["triwulan2"].append(filtered_data)
#                 elif prefix in ('7', '8', '9', '07', '08', '09'):
#                     print(f"Sheet {sheet_name} masuk ke triwulan3")
#                     triwulan_data["triwulan3"].append(filtered_data)
#                 elif prefix in ('10', '11', '12'):
#                     print(f"Sheet {sheet_name} masuk ke triwulan4")
#                     triwulan_data["triwulan4"].append(filtered_data)
#                 else:
#                     print(f"Prefix {prefix} tidak dikenali untuk sheet {sheet_name}")
#             else:
#                 print(f"Sheet {sheet_name} kosong setelah difilter.")
        
#         # print(triwulan_data['triwulan1'])
#         # Menulis hasil gabungan ke file Excel baru
#         with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
#             for triwulan, data_frames in triwulan_data.items():
#                 if data_frames:
#                     # Menggabungkan semua DataFrame dalam satu triwulan
#                     merged_data = pd.concat(data_frames, ignore_index=True)
                    
#                     # Menghilangkan duplikasi berdasarkan kolom B, C, D
#                     merged_data = merged_data.drop_duplicates(subset=['Series', 'First Issue Date', 'Maturity Date'])
                    
#                     # Menulis data yang telah difilter dan tanpa duplikasi ke sheet
#                     merged_data.to_excel(writer, sheet_name=triwulan, index=False)
#                 else:
#                     print(f"Tidak ada data untuk {triwulan}.")
        
#         print(f"Data yang difilter, digabungkan, dan dihapus duplikasi telah disimpan ke: {output_file}")
        
#     except Exception as e:
#         print(f"Terjadi kesalahan: {e}")

# def open_excel_file(file_base_name):
#     """
#     Mencoba membuka file dengan ekstensi .xlsx terlebih dahulu, jika gagal coba .xls.
#     """
#     # Tentukan nama file dengan kedua ekstensi
#     file_xlsx = f"{file_base_name}.xlsx"
#     file_xls = f"{file_base_name}.xls"
    
#     try:
#         # Coba buka file dengan ekstensi .xlsx
#         if os.path.exists(file_xlsx):
#             print(f"Mencoba membuka file: {file_xlsx}")
#             return pd.ExcelFile(file_xlsx), file_xlsx
#         elif os.path.exists(file_xls):
#             # Jika file .xlsx tidak ada, coba .xls
#             print(f"Mencoba membuka file: {file_xls}")
#             return pd.ExcelFile(file_xls), file_xls
#         else:
#             # Jika kedua file tidak ditemukan
#             raise FileNotFoundError(f"Tidak ditemukan file dengan nama {file_base_name}.xlsx atau {file_base_name}.xls")
#     except Exception as e:
#         # Tangani kesalahan saat membuka file
#         print(f"Terjadi kesalahan saat membuka file: {e}")
#         raise

# # File input dan output
# tahun = '2023'
# input_file, full_file_name = open_excel_file(f"sbn-{tahun}")
# output_file = f"sbn-{tahun}-triwulan.xlsx"  # Ganti dengan path file output yang diinginkan

# process_and_merge_sheets(input_file, output_file)

# =====================================================================================================
# =====================================================================================================

def process_and_merge_sheets(input_file, output_file):
    try:
        # Membaca semua sheet dalam file
        excel_data = pd.ExcelFile(input_file)
        
        # Dictionary untuk menyimpan data berdasarkan triwulan
        triwulan_data = {
            "triwulan1": [],
            "triwulan2": [],
            "triwulan3": [],
            "triwulan4": []
        }
        
        for sheet_name in excel_data.sheet_names:
            print(f"Memproses sheet: {sheet_name}")
            
            # Ambil bagian sebelum tanda "-"
            prefix = sheet_name.split("-")[0]
            
            # Membaca kolom A-H dari setiap sheet
            sheet_data = pd.read_excel(input_file, sheet_name=sheet_name, usecols='A:H')
            
            try:
                # Menyeleksi baris di mana kolom 'Unnamed: 2' dan 'Unnamed: 3' tidak kosong
                filtered_data = sheet_data.dropna(subset=['Unnamed: 2', 'Unnamed: 3'])
            except:
                continue

            # Jika ada data setelah difilter
            if not filtered_data.empty:
                # Menggunakan baris pertama sebagai header
                filtered_data.columns = filtered_data.iloc[0]
                filtered_data = filtered_data[1:]  # Menghapus baris pertama setelah menjadi header
                filtered_data.reset_index(drop=True, inplace=True)  # Reset indeks
                
                # Tentukan triwulan berdasarkan prefix
                if prefix in ('1', '2', '3', '01', '02', '03'):
                    print(f"Sheet {sheet_name} masuk ke triwulan1")
                    triwulan_data["triwulan1"].append(filtered_data)
                elif prefix in ('4', '5', '6', '04', '05', '06'):
                    print(f"Sheet {sheet_name} masuk ke triwulan2")
                    triwulan_data["triwulan2"].append(filtered_data)
                elif prefix in ('7', '8', '9', '07', '08', '09'):
                    print(f"Sheet {sheet_name} masuk ke triwulan3")
                    triwulan_data["triwulan3"].append(filtered_data)
                elif prefix in ('10', '11', '12'):
                    print(f"Sheet {sheet_name} masuk ke triwulan4")
                    triwulan_data["triwulan4"].append(filtered_data)
                else:
                    print(f"Prefix {prefix} tidak dikenali untuk sheet {sheet_name}")
            else:
                print(f"Sheet {sheet_name} kosong setelah difilter.")
        
        # Dictionary untuk menyimpan data gabungan tahunan
        yearly_data = []

        # Menulis hasil gabungan ke file Excel baru
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            for triwulan, data_frames in triwulan_data.items():
                if data_frames:
                    # Menggabungkan semua DataFrame dalam satu triwulan
                    merged_data = pd.concat(data_frames, ignore_index=True)
                    
                    # Menghilangkan duplikasi berdasarkan kolom B, C, D
                    merged_data = merged_data.drop_duplicates(subset=['Series', 'First Issue Date', 'Maturity Date'])
                    
                    # Menulis data yang telah difilter dan tanpa duplikasi ke sheet triwulan
                    merged_data.to_excel(writer, sheet_name=triwulan, index=False)
                    
                    # Menambahkan data ke yearly_data
                    yearly_data.append(merged_data)
                else:
                    print(f"Tidak ada data untuk {triwulan}.")
            
            if yearly_data:
                # Menggabungkan semua data triwulan menjadi satu sheet tahunan
                merged_yearly_data = pd.concat(yearly_data, ignore_index=True)
                
                # Menghilangkan duplikasi berdasarkan kolom 'Series', 'First Issue Date', 'Maturity Date'
                merged_yearly_data = merged_yearly_data.drop_duplicates(
                    subset=['Series', 'First Issue Date', 'Maturity Date']
                )
                
                # Menulis data tahunan ke sheet "Tahunan"
                merged_yearly_data.to_excel(writer, sheet_name="Tahunan", index=False)
                print("Sheet tahunan telah digabungkan dan ditulis ke file.")
        
        print(f"Data yang difilter, digabungkan, dan dihapus duplikasi telah disimpan ke: {output_file}")
        
    except Exception as e:
        print(f"Terjadi kesalahan: {e}")

def open_excel_file(file_base_name):
    """
    Mencoba membuka file dengan ekstensi .xlsx terlebih dahulu, jika gagal coba .xls.
    """
    # Tentukan nama file dengan kedua ekstensi
    file_xlsx = f"{file_base_name}.xlsx"
    file_xls = f"{file_base_name}.xls"
    
    try:
        # Coba buka file dengan ekstensi .xlsx
        if os.path.exists(file_xlsx):
            print(f"Mencoba membuka file: {file_xlsx}")
            return pd.ExcelFile(file_xlsx), file_xlsx
        elif os.path.exists(file_xls):
            # Jika file .xlsx tidak ada, coba .xls
            print(f"Mencoba membuka file: {file_xls}")
            return pd.ExcelFile(file_xls), file_xls
        else:
            # Jika kedua file tidak ditemukan
            raise FileNotFoundError(f"Tidak ditemukan file dengan nama {file_base_name}.xlsx atau {file_base_name}.xls")
    except Exception as e:
        # Tangani kesalahan saat membuka file
        print(f"Terjadi kesalahan saat membuka file: {e}")
        raise

# File input dan output
tahun = input('Masukkan tahun: ')
input_file, full_file_name = open_excel_file(f"sbn-{tahun}")
output_file = f"sbn-{tahun}-triwulan-tahunan.xlsx"  # Ganti dengan path file output yang diinginkan

process_and_merge_sheets(input_file, output_file)
