import sys
import os
import subprocess
import json
import base64
import tempfile

def compress_pdf(input_file, output_file, quality="printer"):
    """
    Ghostscript kullanarak PDF dosyasını sıkıştır
    
    Args:
        input_file: Giriş PDF dosya yolu
        output_file: Çıkış PDF dosya yolu
        quality: Sıkıştırma kalitesi (screen, ebook, printer, prepress)
    """
    command = [
        "gs",
        "-sDEVICE=pdfwrite",
        "-dCompatibilityLevel=1.4",
        f"-dPDFSETTINGS=/{quality}",
        "-dNOPAUSE",
        "-dQUIET",
        "-dBATCH",
        f"-sOutputFile={output_file}",
        input_file
    ]
    
    subprocess.run(command, check=True)
    print(f"{output_file} başarıyla sıkıştırıldı.", file=sys.stderr)

def main():
    """
    Script komut satırından çağrıldığında çalışır.
    
    2 parametre alır:
    1. Base64 formatında PDF içeriği
    2. Sıkıştırma seviyesi (light, medium, high)
    """
    if len(sys.argv) < 3:
        print("Kullanım: python compress_pdf.py <base64_data> <compression_level>", file=sys.stderr)
        sys.exit(1)
    
    # Parametreleri al
    base64_data = sys.argv[1]
    compression_level = sys.argv[2]
    
    # Sıkıştırma kalitesini belirle
    if compression_level == "light":
        quality = "printer"    # En yüksek kalite - hafif sıkıştırma
    elif compression_level == "medium":
        quality = "prepress"   # Yüksek kalite - orta sıkıştırma
    elif compression_level == "high":
        quality = "ebook"      # Orta kalite - yüksek sıkıştırma
    else:
        quality = "prepress"   # Varsayılan
    
    try:
        # Geçici dosyaları oluştur
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as temp_input:
            temp_input_path = temp_input.name
        
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as temp_output:
            temp_output_path = temp_output.name
        
        try:
            # Base64 içeriği çöz ve giriş dosyasına yaz
            pdf_data = base64.b64decode(base64_data)
            with open(temp_input_path, 'wb') as f:
                f.write(pdf_data)
            
            # Orijinal boyutu al
            original_size = len(pdf_data)
            
            # PDF'i sıkıştır
            compress_pdf(temp_input_path, temp_output_path, quality)
            
            # Sıkıştırılmış PDF'i oku
            with open(temp_output_path, 'rb') as f:
                compressed_data = f.read()
            
            # Sıkıştırılmış boyutu al
            compressed_size = len(compressed_data)
            
            # Sonucu JSON olarak döndür
            result = {
                "compressed_pdf": base64.b64encode(compressed_data).decode('utf-8'),
                "original_size": original_size,
                "compressed_size": compressed_size
            }
            
            print(json.dumps(result))
            
        finally:
            # Geçici dosyaları temizle
            if os.path.exists(temp_input_path):
                os.unlink(temp_input_path)
            if os.path.exists(temp_output_path):
                os.unlink(temp_output_path)
    
    except Exception as e:
        print(f"Hata: {str(e)}", file=sys.stderr)
        error_result = {"error": str(e)}
        print(json.dumps(error_result))
        sys.exit(1)

if __name__ == "__main__":
    main()