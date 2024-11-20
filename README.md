# Data Rekapitulasi SBN

Script Python ini dirancang untuk menggabungkan data triwulanan dan tahunan dari file Excel (_.xls atau _.xlsx\*) yang memiliki format nama tertentu, serta memastikan data yang dihasilkan unik (tidak ada duplikasi).

## Fitur Utama

- Membaca file input dengan format nama `sbn-[nama tahun]`.
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
