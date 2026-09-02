import io
import re
import fitz  # PyMuPDF
import numpy as np
import pandas as pd
from PIL import Image

# Инициализация EasyOCR для русского и английского языков
try:
    import easyocr

    ocr_reader = easyocr.Reader(["ru", "en"], gpu=False)
except Exception:
    ocr_reader = None


def extract_text_from_image_bytes(image_bytes: bytes) -> str:
    """Извлечение текста из массива байт изображения через EasyOCR."""
    if not ocr_reader:
        return ""
    try:
        image = Image.open(io.BytesIO(image_bytes)).convert("RGB")
        img_np = np.array(image)
        results = ocr_reader.readtext(img_np, detail=0)
        return "\n".join(results)
    except Exception:
        return ""


def parse_excel_accruals(file) -> pd.DataFrame:
    """Парсинг расчетно-платежных ведомостей Excel."""
    try:
        file.seek(0)
        df_raw = pd.read_excel(file, header=None)

        fio_col, net_col = None, None

        # Поиск ключевых столбцов по заголовкам
        for r in range(min(15, len(df_raw))):
            row_vals = [str(val).strip().lower() for val in df_raw.iloc[r]]
            for c, val in enumerate(row_vals):
                if "фио" in val or "фамилия" in val or "работник" in val:
                    fio_col = c
                if "к выдаче" in val or "на руки" in val or "сумма к выплате" in val:
                    net_col = c

        if fio_col is None:
            fio_col = 1
        if net_col is None:
            net_col = df_raw.shape[1] - 1

        records = []
        for r in range(0, len(df_raw)):
            fio_val = str(df_raw.iloc[r, fio_col]).strip()
            net_val = df_raw.iloc[r, net_col]

            if (
                fio_val
                and fio_val.lower() != "nan"
                and len(fio_val.split()) >= 2
            ):
                try:
                    net_float = float(
                        str(net_val)
                        .replace(" ", "")
                        .replace(",", ".")
                        .replace("\xa0", "")
                    )
                    if net_float > 0:
                        records.append({"ФИО": fio_val, "К выдаче (Ведомость)": net_float})
                except ValueError:
                    continue

        return pd.DataFrame(records)
    except Exception as e:
        print(f"Ошибка парсинга Excel: {e}")
        return pd.DataFrame()


def parse_pdf_5_15a(file) -> pd.DataFrame:
    """Гибридный парсер PDF (Текстовый слой + OCR Vision для сканов)."""
    try:
        file.seek(0)
        file_bytes = file.read()
        doc = fitz.open(stream=file_bytes, filetype="pdf")

        full_text = ""
        is_scanned = True

        for page in doc:
            text = page.get_text()
            if len(text.strip()) > 50:
                full_text += text + "\n"
                is_scanned = False

        if is_scanned and ocr_reader:
            ocr_text_list = []
            for page_num in range(len(doc)):
                page = doc[page_num]
                pix = page.get_pixmap(dpi=300)
                img_bytes = pix.tobytes("png")
                page_text = extract_text_from_image_bytes(img_bytes)
                ocr_text_list.append(page_text)
            full_text = "\n".join(ocr_text_list)

        records = []
        lines = full_text.split("\n")

        fio_pattern = re.compile(
            r"([А-ЯӘҒҚҢӨҰҮҺІ][а-яәғқңөұүһі]+\s+[А-ЯӘҒҚҢӨҰҮҺІ][а-яәғқңөұүһі]+(?:\s+[А-ЯӘҒҚҢӨҰҮҺІ][а-яәғқңөұүһі]+)?)"
        )
        amount_pattern = re.compile(r"(\d{1,3}(?:\s?\d{3})*(?:[.,]\d{2}))")

        for line in lines:
            fio_match = fio_pattern.search(line)
            amount_matches = amount_pattern.findall(line)

            if fio_match and amount_matches:
                fio = fio_match.group(1).strip()
                raw_amount = amount_matches[-1]
                try:
                    net_float = float(
                        raw_amount.replace(" ", "").replace(",", ".").replace("\xa0", "")
                    )
                    if net_float > 0:
                        records.append({"ФИО": fio, "Выплачено (5-15А)": net_float})
                except ValueError:
                    continue

        return pd.DataFrame(records)
    except Exception as e:
        print(f"Ошибка парсинга PDF: {e}")
        return pd.DataFrame()


def parse_image_5_15a(file) -> pd.DataFrame:
    """Парсинг прямых файлов изображений (PNG, JPG, JPEG)."""
    try:
        file.seek(0)
        img_bytes = file.read()
        text = extract_text_from_image_bytes(img_bytes)

        records = []
        lines = text.split("\n")

        fio_pattern = re.compile(
            r"([А-ЯӘҒҚҢӨҰҮҺІ][а-яәғқңөұүһі]+\s+[А-ЯӘҒҚҢӨҰҮҺІ][а-яәғқңөұүһі]+(?:\s+[А-ЯӘҒҚҢӨҰҮҺІ][а-яәғқңөұүһі]+)?)"
        )
        amount_pattern = re.compile(r"(\d{1,3}(?:\s?\d{3})*(?:[.,]\d{2}))")

        for line in lines:
            fio_match = fio_pattern.search(line)
            amount_matches = amount_pattern.findall(line)

            if fio_match and amount_matches:
                fio = fio_match.group(1).strip()
                raw_amount = amount_matches[-1]
                try:
                    net_float = float(
                        raw_amount.replace(" ", "").replace(",", ".").replace("\xa0", "")
                    )
                    if net_float > 0:
                        records.append({"ФИО": fio, "Выплачено (5-15А)": net_float})
                except ValueError:
                    continue

        return pd.DataFrame(records)
    except Exception as e:
        print(f"Ошибка парсинга изображения: {e}")
        return pd.DataFrame()


def reconcile_salary(df_accruals: pd.DataFrame, df_payments: pd.DataFrame):
    """Модуль автоматической сверки ведомости и формы 5-15А."""
    if df_accruals.empty or df_payments.empty:
        return pd.DataFrame(), ["Ошибка: Один из наборов данных пуст."]

    acc_grouped = (
        df_accruals.groupby("ФИО", as_index=False)["К выдаче (Ведомость)"].sum()
    )
    pay_grouped = (
        df_payments.groupby("ФИО", as_index=False)["Выплачено (5-15А)"].sum()
    )

    merged = pd.merge(acc_grouped, pay_grouped, on="ФИО", how="outer").fillna(0)
    merged["Расхождение (₸)"] = (
        merged["К выдаче (Ведомость)"] - merged["Выплачено (5-15А)"]
    )

    def get_status(row):
        diff = abs(row["Расхождение (₸)"])
        if diff < 0.01:
            return "✅ Совпадает"
        elif row["К выдаче (Ведомость)"] == 0:
            return "⚠️ Есть в 5-15А, нет в ведомости"
        elif row["Выплачено (5-15А)"] == 0:
            return "⚠️ Есть в ведомости, нет в 5-15А"
        else:
            return "❌ Расхождение сумм"

    merged["Статус"] = merged.apply(get_status, axis=1)

    risk_comments = []
    discrepancies = merged[merged["Статус"] != "✅ Совпадает"]

    if not discrepancies.empty:
        for _, row in discrepancies.iterrows():
            risk_comments.append(
                f"• **{row['ФИО']}**: Ведомость = {row['К выдаче (Ведомость)']:,.2f} ₸, "
                f"5-15А = {row['Выплачено (5-15А)']:,.2f} ₸. "
                f"Разница: **{row['Расхождение (₸)']:,.2f} ₸** ({row['Статус']})"
            )

    return merged, risk_comments