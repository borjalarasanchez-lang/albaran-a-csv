"""Vuelca el texto extraído de ambos PDFs para diseñar el parser real."""
import sys
sys.path.insert(0, r'c:\App para pasar formato')
from extractor import extract_text_and_ocr

for pdf_name in ['muestras/albaran_SALTOKI.pdf', 'muestras/Cmr_Getafe.pdf']:
    print(f"\n{'='*60}")
    print(f"ARCHIVO: {pdf_name}")
    print('='*60)
    text, used_ocr = extract_text_and_ocr(pdf_name)
    print(text[:6000])
    print(f"\n[OCR: {used_ocr}, TOTAL CHARS: {len(text)}]")
