# Instruksi Penggunaan Program

## Deskripsi
Repositori ini berisi dua skrip Python untuk mengelola dokumen:
1. `convert_word_to_pdf.py`: Mengonversi file Word (.docx) di direktori saat ini menjadi file PDF menggunakan Microsoft Word melalui COM automation.
2. `generate_proposal_docs.py`: Membuat dokumen proposal sponsorship yang disesuaikan untuk daftar sponsor berdasarkan template Word.

## Prasyarat
- **Python 3.x** terinstal.
- **Microsoft Word** terinstal (untuk `convert_word_to_pdf.py`).
- **Modul Python** yang diperlukan:
  - `pywin32` (untuk `convert_word_to_pdf.py`): Instal dengan `pip install pywin32`.
  - `python-docx` (untuk `generate_proposal_docs.py`): Instal dengan `pip install python-docx`.
- Sistem operasi: **Windows** (karena `convert_word_to_pdf.py` menggunakan COM automation yang spesifik untuk Windows).
- File template `PROPOSAL_SPONSORSHIP_KESENIAN_BATHIN_ALAM.docx` harus ada di direktori yang sama dengan `generate_proposal_docs.py`.

## Cara Penggunaan

### 1. Konversi File Word ke PDF (`convert_word_to_pdf.py`)
Skrip ini mengonversi semua file `.docx` di direktori saat ini menjadi file PDF.

#### Langkah-langkah:
1. Pastikan Microsoft Word terinstal di komputer Anda.
2. Instal modul `pywin32`:
   ```
   pip install pywin32
   ```
3. Tempatkan file `.docx` yang ingin dikonversi ke PDF di direktori yang sama dengan skrip `convert_word_to_pdf.py`.
4. Jalankan skrip:
   ```
   python convert_word_to_pdf.py
   ```
5. Skrip akan:
   - Membuka setiap file `.docx` di direktori.
   - Mengonversinya ke PDF dengan nama yang sama (ekstensi `.pdf`).
   - Menampilkan pesan konfirmasi untuk setiap file yang berhasil dikonversi atau pesan kesalahan jika gagal.
6. File PDF akan disimpan di direktori yang sama.

#### Catatan:
- Pastikan tidak ada file `.docx` yang sedang dibuka di Microsoft Word saat menjalankan skrip.
- Skrip akan membuka dan menutup Microsoft Word secara otomatis.

### 2. Membuat Dokumen Proposal Sponsorship (`generate_proposal_docs.py`)
Skrip ini membuat dokumen Word baru untuk setiap sponsor berdasarkan template, mengganti nama "Pemuda Grafika" dengan nama sponsor.

#### Langkah-langkah:
1. Instal modul `python-docx`:
   ```
   pip install python-docx
   ```
2. Pastikan file template `PROPOSAL_SPONSORSHIP_KESENIAN_BATHIN_ALAM.docx` ada di direktori yang sama dengan skrip.
3. Jalankan skrip:
   ```
   python generate_proposal_docs.py
   ```
4. Skrip akan:
   - Membaca daftar sponsor yang sudah ditentukan di dalam skrip.
   - Membuat dokumen Word baru untuk setiap sponsor dengan nama file `Proposal_Sponsorship_<nama_sponsor>.docx`.
   - Mengganti teks "Pemuda Grafika" di template dengan nama sponsor, mempertahankan format font Times New Roman, ukuran 12.
   - Menyimpan dokumen baru di direktori yang sama.
   - Menampilkan pesan konfirmasi untuk setiap file yang dibuat.

#### Catatan:
- Pastikan template `PROPOSAL_SPONSORSHIP_KESENIAN_BATHIN_ALAM.docx` memiliki teks "Pemuda Grafika" di bagian yang ingin diganti.
- Nama sponsor dalam daftar di skrip dapat diubah sesuai kebutuhan dengan mengedit variabel `sponsors` di dalam kode.

## Contoh Output
- Untuk `convert_word_to_pdf.py`:
  ```
  File PDF dibuat: ./NamaDokumen.pdf
  ```
- Untuk `generate_proposal_docs.py`:
  ```
  File dibuat: Proposal_Sponsorship_Pemuda_Grafika.docx
  File dibuat: Proposal_Sponsorship_Sejarah_Baru.docx
  ...
  ```

## Pemecahan Masalah
- **Error `pywin32` tidak ditemukan**: Pastikan modul `pywin32` terinstal.
- **Error `python-docx` tidak ditemukan**: Instal modul `python-docx`.
- **File template tidak ditemukan**: Pastikan file `PROPOSAL_SPONSORSHIP_KESENIAN_BATHIN_ALAM.docx` ada di direktori yang sama.
- **Microsoft Word tidak terdeteksi**: Pastikan Microsoft Word terinstal dan skrip dijalankan di Windows.
- Jika terjadi error lain, periksa pesan error di output konsol untuk detail.

## Kontribusi
Jika ingin menambahkan fitur atau memperbaiki bug, silakan buat pull request di repositori ini.

## Lisensi
Skrip ini dibagikan tanpa lisensi resmi. Gunakan dengan tanggung jawab sendiri.