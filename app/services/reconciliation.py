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
    return df_raw