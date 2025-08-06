import os
from win32com import client

# Direktori tempat file Word berada
directory = "."

# Inisialisasi aplikasi Word
word = client.Dispatch("Word.Application")
word.Visible = True  # Jalankan Word di latar belakang

# Loop melalui semua file Word di direktori
for filename in os.listdir(directory):
    if filename.endswith(".docx"):
        try:
            # Buka dokumen Word
            doc_path = os.path.abspath(os.path.join(directory, filename))
            doc = word.Documents.Open(doc_path)
            
            # Simpan sebagai PDF
            pdf_path = os.path.splitext(doc_path)[0] + ".pdf"
            doc.SaveAs(pdf_path, FileFormat=17)  # 17 adalah format PDF di Word
            doc.Close()
            
            print(f"File PDF dibuat: {pdf_path}")
        except Exception as e:
            print(f"Error mengonversi {filename}: {e}")

# Tutup aplikasi Word
word.Quit()