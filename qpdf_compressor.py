#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys
import base64
import json
import tempfile
import subprocess
from pathlib import Path

def compress_pdf_with_qpdf(input_data, compression_level="medium"):
    """
    QPDF kullanarak PDF dosyasını sıkıştırır
    
    Args:
        input_data: Base64 olarak kodlanmış PDF içeriği
        compression_level: "light", "medium", "high" sıkıştırma seviyesi
    
    Returns:
        Base64 olarak kodlanmış sıkıştırılmış PDF içeriği ve boyut bilgileri
    """
    try:
        # Temp dizini oluştur
        temp_dir = tempfile.mkdtemp()
        input_path = os.path.join(temp_dir, "input.pdf")
        optimized_path = os.path.join(temp_dir, "optimized.pdf")
        output_path = os.path.join(temp_dir, "output.pdf")
        
        # Base64'ten normal veri formatına dönüştür
        pdf_data = base64.b64decode(input_data)
        
        # Geçici dosyaya yaz
        with open(input_path, "wb") as f:
            f.write(pdf_data)
        
        # Orijinal boyutu kaydet
        original_size = os.path.getsize(input_path)
        
        # 1. Adım: QPDF ile optimize et
        subprocess.run([
            "qpdf",
            "--linearize",           # Web optimizasyonu
            "--compress-streams=y",  # Tüm akışları sıkıştır
            "--object-streams=generate", # Nesne akışlarını oluştur
            input_path,
            optimized_path
        ], check=True)
        
        # 2. Adım: Sıkıştırma seviyesine göre ek optimizasyon
        if compression_level == "light":
            # Hafif sıkıştırma - minimum değişiklik, içerik korunur
            subprocess.run([
                "qpdf",
                "--linearize",
                "--compress-streams=y",
                "--preserve-unreferenced=y",  # Referanssız nesneleri koru
                optimized_path,
                output_path
            ], check=True)
        elif compression_level == "medium":
            # Orta sıkıştırma - daha fazla optimizasyon
            subprocess.run([
                "qpdf",
                "--linearize",
                "--compress-streams=y",
                "--object-streams=generate",
                "--recompress-flate",  # Flate sıkıştırmasını yeniden uygula
                optimized_path,
                output_path
            ], check=True)
        elif compression_level == "high":
            # Yüksek sıkıştırma - en agresif ayarlar
            subprocess.run([
                "qpdf",
                "--linearize",
                "--compress-streams=y",
                "--object-streams=generate",
                "--recompress-flate",
                "--compression-level=9",  # En yüksek sıkıştırma seviyesi
                "--min-version=1.5",     # En düşük PDF versiyonu
                "--remove-unreferenced", # Referanssız nesneleri temizle
                optimized_path,
                output_path
            ], check=True)
        else:
            # Varsayılan orta sıkıştırma
            subprocess.run([
                "qpdf",
                "--linearize",
                "--compress-streams=y",
                "--object-streams=generate",
                optimized_path,
                output_path
            ], check=True)
        
        # Sıkıştırılmış dosya boyutunu al
        compressed_size = os.path.getsize(output_path)
        
        # Sıkıştırılmış PDF'i oku ve base64'e çevir
        with open(output_path, "rb") as f:
            compressed_data = f.read()
            compressed_pdf_base64 = base64.b64encode(compressed_data).decode("utf-8")
        
        # Geçici dosyaları temizle
        try:
            os.remove(input_path)
            os.remove(optimized_path)
            os.remove(output_path)
            os.rmdir(temp_dir)
        except:
            pass
        
        # Sonuçları döndür
        return {
            "original_size": original_size,
            "compressed_size": compressed_size,
            "compressed_pdf": compressed_pdf_base64,
            "error": None
        }
    
    except Exception as e:
        return {
            "error": str(e),
            "original_size": 0,
            "compressed_size": 0,
            "compressed_pdf": ""
        }

def main():
    """
    Komut satırından çağrıldığında çalışır.
    Beklenen argümanlar:
    1. Base64 formatında PDF içeriği
    2. Sıkıştırma seviyesi (light, medium, high)
    """
    if len(sys.argv) < 3:
        result = {
            "error": "Geçersiz argüman sayısı. Base64 PDF ve sıkıştırma seviyesi gerekli.",
            "original_size": 0,
            "compressed_size": 0,
            "compressed_pdf": ""
        }
        print(json.dumps(result))
        sys.exit(1)
    
    # Komut satırı argümanlarını al
    pdf_base64 = sys.argv[1]
    compression_level = sys.argv[2]
    
    # PDF'i sıkıştır
    result = compress_pdf_with_qpdf(pdf_base64, compression_level)
    
    # JSON formatında sonucu yazdır
    print(json.dumps(result))

if __name__ == "__main__":
    main()