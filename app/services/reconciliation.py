import io
import re
import fitz  # PyMuPDF
import pandas as pd
import pytesseract
from PIL import Image


def parse_excel_accruals(file_bytes):
    """
    Парсинг расчетной ведомости Excel.
    Сканирует все строки без потери сотрудников при сложных объединениях.
    """
    df_raw = pd.read_excel(file_bytes, header=None)
    if df_raw.empty:
        return df_raw

    # Находим первую строку, где начинается шапка
    header_idx = 0
    for idx, row in df_raw.iterrows():
        row_str = " ".join(row.dropna().astype(str)).lower()
        if any(
            keyword in row_str
            for keyword in [
                "фио",
                "фамилия",
                "иин",
                "начислено",
                "к выплате",
                "всего",
            ]
        ):
            header_idx = idx
            break

    # Считываем данные, используя найденный индекс как заголовок
    df = pd.read_excel(file_bytes, header=header_idx)
    df.columns = [
        str(c).strip().replace("\n", " ")
        if pd.notna(c)
        else f"Unnamed_{i}"
        for i, c in enumerate(df.columns)
    ]
    return df


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


def reconcile_salary(df_accruals: pd.DataFrame, df_payments: pd.DataFrame):
    """
    Универсальный сведениитель данных по ВСЕМ сотрудникам без исключения.
    """
    risk_comments = []
    records = []

    if df_accruals.empty:
        return pd.DataFrame(), ["Файл начислений пуст или не прочитан."]

    # 1. Авто-поиск колонки с ФИО
    fio_col = None
    for col in df_accruals.columns:
        series_str = df_accruals[col].dropna().astype(str)
        # Ищем колонку с текстовыми ФИО (слово с заглавной буквы)
        match_count = series_str.str.contains(
            r"[А-ЯӘҒҚҢӨҰҮҺІA-Z][а-яәғқңөұүһіa-z]+\s+[А-ЯӘҒҚҢӨҰҮҺІA-Z]", regex=True
        ).sum()
        if match_count > 1:
            fio_col = col
            break

    if fio_col is None:
        fio_col = df_accruals.columns[0]

    # Stop-words для фильтрации мусора, итогов и системных строк
    stop_words = [
        "nan",
        "none",
        "фио",
        "ф.и.о.",
        "сотрудник",
        "наименование",
        "всего",
        "итого",
        "подпись",
        "руководитель",
        "бухгалтер",
        "начислено",
        "удержано",
    ]

    months = [
        "Январь",
        "Февраль",
        "Март",
        "Апрель",
        "Май",
        "Июнь",
        "Июль",
        "Август",
        "Сентябрь",
        "Октябрь",
        "Ноябрь",
        "Декабрь",
    ]

    # 2. Обход всех строк
    for idx, row in df_accruals.iterrows():
        raw_fio = str(row[fio_col]).strip()
        fio_lower = raw_fio.lower()

        if (
            len(raw_fio) < 3
            or fio_lower in stop_words
            or any(s in fio_lower for s in ["всего", "итого", "подпись"])
        ):
            continue

        # Форматирование ФИО
        fio_clean = re.sub(r"\s+", " ", raw_fio)

        curr_bal = 0.0

        for m_idx, m_name in enumerate(months):
            kvyd_val = 0.0
            paid_val = 0.0

            # Ищем суммы в колонках, относящихся к текущему месяцу
            for col in df_accruals.columns:
                col_str = str(col).lower()
                if (
                    m_name.lower() in col_str
                    or f".{m_idx+1:02d}." in col_str
                    or f"_{m_idx+1}" in col_str
                ):
                    val = row[col]
                    try:
                        parsed_val = float(
                            str(val).replace(",", ".").replace(" ", "")
                        )
                        if not pd.isna(parsed_val):
                            if (
                                "выдач" in col_str
                                or "начисл" in col_str
                                or "сумма" in col_str
                            ):
                                kvyd_val += parsed_val
                            elif "выплат" in col_str or "перечисл" in col_str:
                                paid_val += parsed_val
                    except ValueError:
                        pass

            end_bal = curr_bal + kvyd_val - paid_val
            status = "Закрыто" if abs(end_bal) < 0.01 else "Расхождение"

            records.append({
                "fio": fio_clean,
                "month": m_name,
                "start_bal": round(curr_bal, 2),
                "kvyd": round(kvyd_val, 2),
                "paid": round(paid_val, 2),
                "end_bal": round(end_bal, 2),
                "status": status,
            })

            curr_bal = end_bal

    df_result = pd.DataFrame(records)

    if not df_result.empty:
        total_people = df_result["fio"].nunique()
        risk_comments.append(f"Успешно обработано сотрудников: {total_people}.")
    else:
        risk_comments.append(
            "Не удалось выделить список сотрудников. Проверьте форматирование файла."
        )

    return df_result, risk_comments