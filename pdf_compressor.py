import sys
import base64
import json
import os
import tempfile
import subprocess

def compress_pdf(input_file, output_file, compression_level="medium"):
    """
    Ghostscript kullanarak PDF'i sıkıştır
    
    Args:
        input_file: Giriş PDF dosya yolu
        output_file: Çıkış PDF dosya yolu
        compression_level: Sıkıştırma seviyesi
    """
    # PDFSETTINGS değerini belirle
    if compression_level == "light":
        pdfsettings = "/prepress"  # 300 dpi yüksek kalite
    elif compression_level == "medium": 
        pdfsettings = "/ebook"     # 150 dpi orta kalite
    elif compression_level == "high":
        pdfsettings = "/screen"    # 72 dpi düşük kalite
    else:
        pdfsettings = "/ebook"     # varsayılan
    
    # Ghostscript komutu
    cmd = [
        'gs',
        '-sDEVICE=pdfwrite',
        '-dCompatibilityLevel=1.4',
        f'-dPDFSETTINGS={pdfsettings}',
        '-dNOPAUSE',
        '-dQUIET',
        '-dBATCH',
        f'-sOutputFile={output_file}',
        input_file
    ]
    
    # Komutu çalıştır
    subprocess.run(cmd, check=True)

def main():
    """
    Komut satırından çağrıldığında çalışır.
    """
    try:
        # Base64 kodlanmış PDF içeriği ve sıkıştırma seviyesi
        encoded_pdf = sys.argv[1]
        compression_level = sys.argv[2] if len(sys.argv) > 2 else "medium"
        
        # Giriş için geçici dosya oluştur
        temp_input = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf')
        temp_input_path = temp_input.name
        temp_input.close()
        
        # Çıkış için geçici dosya yolu
        temp_output_path = tempfile.NamedTemporaryFile(delete=False, suffix='.pdf').name
        
        try:
            # Base64'ten bytes'a çevir ve dosyaya yaz
            pdf_bytes = base64.b64decode(encoded_pdf)
            with open(temp_input_path, 'wb') as f:
                f.write(pdf_bytes)
            
            # Orijinal boyut
            original_size = len(pdf_bytes)
            
            # PDF'i sıkıştır
            compress_pdf(temp_input_path, temp_output_path, compression_level)
            
            # Sıkıştırılmış PDF'i oku
            with open(temp_output_path, 'rb') as f:
                compressed_bytes = f.read()
            
            # Sıkıştırılmış boyut
            compressed_size = len(compressed_bytes)
            
            # Base64'e kodla
            compressed_base64 = base64.b64encode(compressed_bytes).decode('utf-8')
            
            # Sonucu JSON olarak döndür
            result = {
                "compressed_pdf": compressed_base64,
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
        error_result = {
            "error": str(e)
        }
        print(json.dumps(error_result), file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()