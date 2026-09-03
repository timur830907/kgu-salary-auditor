import io
import fitz  # PyMuPDF
import pytesseract
from PIL import Image
import pandas as pd

def parse_excel_accruals(file_bytes):
    """Парсинг расчетной ведомости Excel."""
    df_raw = pd.read_excel(file_bytes, header=None)
    
    header_idx = None
    for idx, row in df_raw.iterrows():
        row_str = " ".join(row.dropna().astype(str)).lower()
        if any(keyword in row_str for keyword in ["фио", "фамилия", "иин", "начислено", "к выплате", "всего"]):
            header_idx = idx
            break
            
    if header_idx is not None:
        df = pd.read_excel(file_bytes, header=header_idx)
        df.columns = [str(c).strip().replace('\n', ' ') for c in df.columns]
        df = df.dropna(how='all')
        return df
    
    return df_raw

# Алиас для обеспечения совместимости
parse_excel_payroll = parse_excel_accruals

def extract_text_from_pdf(pdf_bytes):
    """Извлечение текста из PDF с поддержкой OCR через Tesseract."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    full_text = ""

    for page in doc:
        text = page.get_text()
        if text and len(text.strip()) > 10:
            full_text += text + "\n"
        else:
            pix = page.get_pixmap(dpi=150)
            img = Image.open(io.BytesIO(pix.tobytes("png")))
            try:
                ocr_text = pytesseract.image_to_string(img, lang="rus+eng")
                full_text += ocr_text + "\n"
            except Exception:
                ocr_text = pytesseract.image_to_string(img)
                full_text += ocr_text + "\n"

    return full_text

def parse_image_5_15a(image_bytes):
    """Распознавание текста со скана/изображения 5-15А."""
    img = Image.open(io.BytesIO(image_bytes))
    try:
        return pytesseract.image_to_string(img, lang="rus+eng")
    except Exception:
        return pytesseract.image_to_string(img)

def parse_pdf_5_15a(pdf_bytes):
    """Парсинг PDF выписки 5-15А."""
    return extract_text_from_pdf(pdf_bytes)