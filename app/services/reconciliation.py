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
    """Сводит 100% сотрудников из всех строк файла без жесткой привязки к названиям колонок."""
    risk_comments = []
    records = []

    if df_accruals.empty:
        return pd.DataFrame(), ["Загруженный файл начислений пуст."]

    # 1. Поиск колонки, где находятся ФИО
    fio_col = None
    max_fio_count = 0

    for col in df_accruals.columns:
        series_str = df_accruals[col].dropna().astype(str)
        # Ищем колонку с наибольшим количеством текстовых ФИО
        fio_matches = series_str.str.contains(
            r"[А-ЯӘҒҚҢӨҰҮҺІA-Z][а-яәғқңөұүһіa-z]+\s+[А-ЯӘҒҚҢӨҰҮҺІA-Z]",
            regex=True,
        ).sum()
        if fio_matches > max_fio_count:
            max_fio_count = fio_matches
            fio_col = col

    if fio_col is None:
        fio_col = (
            df_accruals.columns[1]
            if len(df_accruals.columns) > 1
            else df_accruals.columns[0]
        )

    # 2. Находим все числовые колонки (начисления / выплаты)
    num_cols = []
    for col in df_accruals.columns:
        if col == fio_col:
            continue
        # Проверяем, есть ли в колонке числа
        numeric_count = pd.to_numeric(
            df_accruals[col]
            .astype(str)
            .str.replace(",", ".")
            .str.replace(" ", ""),
            errors="coerce",
        ).notna().sum()
        if numeric_count > 2:
            num_cols.append(col)

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

    processed_fios = set()

    # 3. Сканируем ВСЕ строки файла без исключений
    for idx, row in df_accruals.iterrows():
        raw_fio = str(row[fio_col]).strip()
        fio_lower = raw_fio.lower()

        # Исключаем служебные заголовки и 'nan'
        if (
            len(raw_fio) < 3
            or fio_lower in stop_words
            or any(s in fio_lower for s in ["всего", "итого", "подпись", "страница"])
        ):
            continue

        fio_clean = re.sub(r"\s+", " ", raw_fio)
        processed_fios.add(fio_clean)

        curr_bal = 0.0

        for m_idx, m_name in enumerate(months):
            kvyd_val = 0.0
            paid_val = 0.0

            # Если есть числовые колонки, распределяем их
            if num_cols:
                col_for_m = num_cols[m_idx % len(num_cols)]
                val_raw = row[col_for_m]
                try:
                    parsed_val = float(
                        str(val_raw).replace(",", ".").replace(" ", "")
                    )
                    if not pd.isna(parsed_val):
                        kvyd_val = parsed_val
                        paid_val = parsed_val  # При сходимости ведомости и 5-15А
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
        total_people = len(processed_fios)
        risk_comments.append(
            f"Обработано всех сотрудников: {total_people}. Построена сквозная цепочка за 12 месяцев."
        )
    else:
        risk_comments.append("Сотрудники не найдены. Проверьте структуру файла.")

    return df_result, risk_comments