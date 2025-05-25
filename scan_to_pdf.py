#!/usr/bin/env python3
"""
Görsel dosyasını PDF'e dönüştürme scripti
Kullanım: python3 scan_to_pdf.py input_image.jpg output.pdf
"""

import sys
import os
from fpdf import FPDF
from PIL import Image

def image_to_pdf(image_path, pdf_path):
    """
    Görsel dosyasını PDF'e dönüştürür
    
    Args:
        image_path: Giriş görsel dosyası yolu
        pdf_path: Çıkış PDF dosyası yolu
    """
    try:
        # Görseli aç ve boyutlarını al
        with Image.open(image_path) as img:
            # RGBA modunda ise RGB'ye dönüştür
            if img.mode == 'RGBA':
                img = img.convert('RGB')
            
            img_width, img_height = img.size
            
            # PDF oluştur
            pdf = FPDF()
            
            # A4 sayfa boyutları (mm)
            page_width = 210
            page_height = 297
            
            # Görsel oranını koru ve sayfaya sığdır
            aspect_ratio = img_width / img_height
            
            if aspect_ratio > page_width / page_height:
                # Görsel daha geniş, genişliğe göre ayarla
                pdf_img_width = page_width - 20  # 10mm kenar boşluğu
                pdf_img_height = pdf_img_width / aspect_ratio
            else:
                # Görsel daha uzun, yüksekliğe göre ayarla
                pdf_img_height = page_height - 20  # 10mm kenar boşluğu
                pdf_img_width = pdf_img_height * aspect_ratio
            
            # Ortalamak için pozisyon hesapla
            x = (page_width - pdf_img_width) / 2
            y = (page_height - pdf_img_height) / 2
            
            # Sayfa ekle ve görseli yerleştir
            pdf.add_page()
            pdf.image(image_path, x=x, y=y, w=pdf_img_width, h=pdf_img_height)
            
            # PDF'i kaydet
            pdf.output(pdf_path)
            
            print(f"PDF başarıyla oluşturuldu: {pdf_path}")
            return True
            
    except Exception as e:
        print(f"Hata: {str(e)}", file=sys.stderr)
        return False

def main():
    """
    Ana fonksiyon - komut satırı argümanlarını işle
    """
    if len(sys.argv) != 3:
        print("Kullanım: python3 scan_to_pdf.py input_image.jpg output.pdf", file=sys.stderr)
        sys.exit(1)
    
    image_path = sys.argv[1]
    pdf_path = sys.argv[2]
    
    # Giriş dosyasının varlığını kontrol et
    if not os.path.exists(image_path):
        print(f"Hata: Görsel dosyası bulunamadı: {image_path}", file=sys.stderr)
        sys.exit(1)
    
    # Çıkış klasörünü oluştur
    output_dir = os.path.dirname(pdf_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Dönüştürme işlemini gerçekleştir
    success = image_to_pdf(image_path, pdf_path)
    
    if success:
        sys.exit(0)
    else:
        sys.exit(1)

if __name__ == "__main__":
    main()