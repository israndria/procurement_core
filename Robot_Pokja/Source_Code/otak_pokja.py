import google.generativeai as genai
import os
import glob
import pdfplumber
from dotenv import load_dotenv

# Load Environment Variables
load_dotenv()

# Setup API Key
API_KEY = os.getenv("GOOGLE_API_KEY")

class RobotPokja:
    def __init__(self):
        if not API_KEY:
            print("❌ API Key belum diset! Silakan buat file .env berisi GOOGLE_API_KEY=...")
            self.model = None
            return
            
        genai.configure(api_key=API_KEY)
        self.model = genai.GenerativeModel('gemini-pro')
        self.context = ""
        print("✅ Otak Gemini Siap!")

    def ingest_peraturan(self, folder_path):
        print(f"📖 Membaca peraturan dari: {folder_path}")
        pdf_files = glob.glob(os.path.join(folder_path, "*.pdf"))
        
        full_text = ""
        for pdf_file in pdf_files:
            print(f"   - Memproses: {os.path.basename(pdf_file)}...")
            try:
                with pdfplumber.open(pdf_file) as pdf:
                    for page in pdf.pages:
                        extracted = page.extract_text()
                        if extracted:
                            full_text += extracted + "\n"
            except Exception as e:
                print(f"     ⚠️ Gagal baca PDF: {e}")
        
        self.context = full_text
        print(f"✅ Selesai! Bot sekarang mengingat {len(full_text)} karakter dari peraturan.")

    def tanya_hakim(self, pertanyaan):
        if not self.model:
            return "Bot belum siap (No API Key)."
        
        if not self.context:
            return "Bot belum membaca peraturan apapun. Jalankan ingest_peraturan dulu."

        prompt = f"""
        Kamu adalah Asisten Pokja (Kelompok Kerja Pemilihan) yang ahli dalam Pengadaan Barang/Jasa Pemerintah.
        
        Berikut adalah konteks peraturan yang kamu ketahui:
        {self.context[:30000]}  # Limit sementara agar tidak overload token (bisa diimprove dgn RAG)
        
        Pertanyaan User: {pertanyaan}
        
        Jawablah berdasarkan konteks di atas. Jika tidak ada di konteks, katakan bahwa tidak ditemukan di peraturan yang diberikan.
        Sertakan referensi (misal: Sesuai Pasal X...) jika memungkinkan.
        """
        
        try:
            print("🤔 Sedang berpikir...")
            response = self.model.generate_content(prompt)
            return response.text
        except Exception as e:
            return f"Error: {e}"

if __name__ == "__main__":
    bot = RobotPokja()
    
    # Path Default
    base_path = os.path.dirname(os.path.abspath(__file__))
    gudang_path = os.path.join(os.path.dirname(base_path), "Gudang_Peraturan")
    
    # Cek folder
    if not os.path.exists(gudang_path):
        os.makedirs(gudang_path)
        print("📁 Folder Gudang_Peraturan dibuat. Silakan isi PDF!")
    
    # Menu Loop
    while True:
        print("\n--- ROBOT POKJA ---")
        print("1. Baca Ulang Peraturan (Ingest)")
        print("2. Tanya Aturan")
        print("3. Keluar")
        
        pilihan = input("Pilih: ")
        
        if pilihan == "1":
            bot.ingest_peraturan(gudang_path)
        elif pilihan == "2":
            q = input("Pertanyaan: ")
            jawaban = bot.tanya_hakim(q)
            print("\n🤖 JAWABAN:")
            print(jawaban)
        elif pilihan == "3":
            break
