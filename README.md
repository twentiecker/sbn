# Data Rekapitulasi SBN

Script Python ini dirancang untuk menggabungkan data triwulanan dan tahunan dari file Excel (_.xls atau _.xlsx\*) yang memiliki format nama tertentu, serta memastikan data yang dihasilkan unik (tidak ada duplikasi).

## Fitur Utama

- Membaca file input dengan format nama `sbn-<nama tahun>`.
- Mendukung format file Excel: **.xls** dan **.xlsx**.
- Menggabungkan data dari berbagai periode (triwulanan dan tahunan).
- Menghapus data duplikat secara otomatis.
- Memproses file yang ada dalam **satu root directory** dengan file utama (`main.py`).

## Persyaratan

### Python Version

- Python **3.10.11**

### Modul yang Dibutuhkan

Sebelum menjalankan script, pastikan modul berikut sudah diinstal:

- **pandas**: untuk manipulasi dan analisis data.
- **openpyxl**: untuk mendukung file Excel dengan format .xlsx.
- **xlrd**: untuk mendukung file Excel dengan format .xls.

Anda dapat menginstalnya dengan perintah berikut:

```bash
pip install pandas openpyxl xlrd
```

## Format File Input

File yang akan diproses harus:

1. Berformat **.xls** atau **.xlsx**.
2. Memiliki nama dengan format:  
   `sbn-<nama tahun>`  
   Contoh: `sbn-2023.xlsx`

---

## Cara Penggunaan

1. **Letakkan File di Folder yang Sama**  
   Pastikan semua file input berada dalam satu folder yang sama dengan `main.py`.

2. **Jalankan Program**  
   Saat menjalankan script, program akan meminta Anda untuk memasukkan tahun yang ingin diproses. Contohnya:
   ```bash
   python main.py
   ```
3. **Hasil Output**  
   Data yang sudah digabungkan dan dihapus duplikasinya akan disimpan dalam file baru dengan format `sbn-<tahun>-triwulan-tahunan.xlsx`

## Struktur Proyek

Berikut adalah contoh struktur direktori:
/project-folder
├── main.py
├── sbn-2023.xlsx
├── sbn-2022.xlsx
└── README.md

## Catatan

- Pastikan file input menggunakan format nama yang benar, atau program tidak dapat mengenalinya.
- Data yang digabungkan akan otomatis menghapus duplikasi berdasarkan kriteria unik dalam data.

---

## Kontribusi

Jika Anda ingin mengembangkan atau melaporkan bug, silakan ajukan issue atau pull request ke repository ini.
