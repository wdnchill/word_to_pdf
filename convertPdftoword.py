from spire.doc import *
from spire.doc.common import *
import easygui

# Fungsi untuk memilih file Word
def pilih_file():
    file_path = easygui.fileopenbox(
        title="Pilih file Word", 
        filetypes=["*.docx", "*.doc"]
    )
    return file_path

# Memilih file Word
file_word = pilih_file()

# Pastikan file dipilih
if file_word:
    # Membuat objek Document
    document = Document()
    
    # Memuat file Word yang dipilih
    document.LoadFromFile(file_word)
    
    # Menyimpan file dalam format PDF
    output_pdf = file_word.rsplit('.', 1)[0] + ".pdf"
    document.SaveToFile(output_pdf, FileFormat.PDF)
    
    # Menutup dokumen
    document.Close()

    print(f"File berhasil dikonversi menjadi {output_pdf}")
else:
    print("Tidak ada file yang dipilih.")

