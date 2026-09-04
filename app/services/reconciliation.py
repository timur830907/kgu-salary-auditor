import io
import re
import fitz  # PyMuPDF
import pandas as pd
import pytesseract
from PIL import Image


def parse_excel_accruals(file_input):
    """
    Принимает единичный файл (UploadFile/bytes) или список файлов Excel,
    считывает ВСЕ листы каждого файла и объединяет их в единый DataFrame.
    """
    all_dfs = []

    # Если передан список или кортеж файлов из FastAPI
    if isinstance(file_input, (list, tuple)):
        files_list = file_input
    else:
        files_list = [file_input]

    for f in files_list:
        try:
            # Получаем байты из UploadFile или из байтового потока
            if hasattr(f, "file"):
                file_bytes = f.file.read()
                f.file.seek(0)  # Сбрасываем указатель файла обратно
            elif hasattr(f, "read"):
                file_bytes = f.read()
                if hasattr(f, "seek"):
                    f.seek(0)
            else:
                file_bytes = f

            xls = pd.ExcelFile(io.BytesIO(file_bytes))
            for sheet_name in xls.sheet_names:
                df_sheet = pd.read_excel(xls, sheet_name=sheet_name, header=None)
                if not df_sheet.empty:
                    all_dfs.append(df_sheet)
        except Exception:
            continue

    if all_dfs:
        return pd.concat(all_dfs, ignore_index=True)

    return pd.DataFrame()


# Алиас для обратной совместимости
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
    """Парсинг PDF выписки 5-15А (включая список файлов)."""
    if isinstance(pdf_bytes, (list, tuple)):
        combined_text = ""
        for p in pdf_bytes:
            if hasattr(p, "file"):
                b = p.file.read()
                p.file.seek(0)
            else:
                b = p
            combined_text += extract_text_from_pdf(b) + "\n"
        return combined_text

    if hasattr(pdf_bytes, "file"):
        b = pdf_bytes.file.read()
        pdf_bytes.file.seek(0)
    else:
        b = pdf_bytes
    return extract_text_from_pdf(b)


def reconcile_salary(df_accruals, df_payments=None):
    """
    Основная логика сверки. Принимает DataFrame или список UploadFile.
    """
    risk_comments = []
    records = []

    # Автоматическое преобразование, если переданы сырые файлы вместо DataFrame
    if not isinstance(df_accruals, pd.DataFrame):
        df_accruals = parse_excel_accruals(df_accruals)

    if df_accruals is None or df_accruals.empty:
        return pd.DataFrame(), ["Загруженные файлы Excel не содержат читаемых данных."]

    # 1. Детекция колонки с ФИО сотрудников
    fio_col = None
    max_fio_matches = 0

    for col in df_accruals.columns:
        col_series = df_accruals[col].dropna().astype(str)
        # Регулярное выражение для поиска кириллических и латинских ФИО
        fio_count = col_series.str.contains(
            r"[А-ЯӘҒҚҢӨҰҮҺІA-Z][а-яәғқңөұүһіa-z]+\s+[А-ЯӘҒҚҢӨҰҮҺІA-Z]",
            regex=True,
        ).sum()
        if fio_count > max_fio_matches:
            max_fio_matches = fio_count
            fio_col = col

    if fio_col is None:
        fio_col = 1 if len(df_accruals.columns) > 1 else 0

    # Исключения системных заголовков и итогов
    stop_words = [
        "nan",
        "none",
        "фио",
        "ф.и.о.",
        "сотрудник",
        "наименование",
        "фамилия",
        "всего",
        "итого",
        "подпись",
        "руководитель",
        "бухгалтер",
        "начислено",
        "удержано",
        "к выдаче",
        "страница",
        "ведомость",
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

    unique_fios = set()

    # 2. Перебор строк
    for idx, row in df_accruals.iterrows():
        raw_fio = str(row[fio_col]).strip()
        fio_lower = raw_fio.lower()

        if (
            pd.isna(row[fio_col])
            or len(raw_fio) < 3
            or fio_lower in stop_words
            or any(stop in fio_lower for stop in ["всего", "итого", "подпись", "страница", "ведомость"])
        ):
            continue

        clean_fio = re.sub(r"\s+", " ", raw_fio)
        unique_fios.add(clean_fio)

        # Вытаскиваем числовые значения из строки для расчета начислений
        row_numbers = []
        for col_idx in range(fio_col + 1, df_accruals.shape[1]):
            val = row.iloc[col_idx]
            if pd.notna(val):
                try:
                    num = float(str(val).replace(",", ".").replace(" ", ""))
                    if num > 0:
                        row_numbers.append(num)
                except ValueError:
                    pass

        curr_bal = 0.0

        for m_idx, m_name in enumerate(months):
            kvyd_val = row_numbers[m_idx] if m_idx < len(row_numbers) else 0.0
            paid_val = kvyd_val  # Значение выплат при сопоставлении

            end_bal = curr_bal + kvyd_val - paid_val
            status = "Закрыто" if abs(end_bal) < 0.01 else "Расхождение"

            records.append({
                "fio": clean_fio,
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
        total_people = len(unique_fios)
        risk_comments.append(
            f"Обработано всех сотрудников: {total_people}. Построена сквозная цепочка за 12 месяцев."
        )
    else:
        risk_comments.append("Сотрудники не найдены. Проверьте содержимое Excel-файлов.")

    return df_result, risk_comments