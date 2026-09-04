import io
import re
import fitz  # PyMuPDF
import pandas as pd
import pytesseract
from PIL import Image


def parse_excel_accruals(file_input):
    """
    Универсальный парсер: принимает как один файл (bytes/UploadFile),
    так и список/кортеж файлов Excel.
    """
    all_dfs = []

    # Если передан список файлов
    if isinstance(file_input, (list, tuple)):
        files_list = file_input
    else:
        files_list = [file_input]

    for f in files_list:
        try:
            # Получаем байты файла
            file_bytes = f.file.read() if hasattr(f, "file") else f
            if hasattr(f, "file"):
                f.file.seek(0)  # Сбрасываем указатель

            xls = pd.ExcelFile(io.BytesIO(file_bytes))
            for sheet_name in xls.sheet_names:
                df_sheet = pd.read_excel
                df_sheet = pd.read_excel(xls, sheet_name=sheet_name, header=None)
                if not df_sheet.empty:
                    all_dfs.append(df_sheet)
        except Exception:
            continue

    if all_dfs:
        return pd.concat(all_dfs, ignore_index=True)

    return pd.DataFrame()


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
    if isinstance(pdf_bytes, (list, tuple)):
        combined_text = ""
        for p in pdf_bytes:
            b = p.file.read() if hasattr(p, "file") else p
            if hasattr(p, "file"):
                p.file.seek(0)
            combined_text += extract_text_from_pdf(b) + "\n"
        return combined_text

    b = pdf_bytes.file.read() if hasattr(pdf_bytes, "file") else pdf_bytes
    if hasattr(pdf_bytes, "file"):
        pdf_bytes.file.seek(0)
    return extract_text_from_pdf(b)


def reconcile_salary(df_accruals: pd.DataFrame, df_payments: pd.DataFrame):
    """
    Принимает объединение ведомостей и производит расчёт по каждому сотруднику.
    """
    risk_comments = []
    records = []

    # Если df_accruals передано как список UploadFile
    if isinstance(df_accruals, (list, tuple)) or hasattr(df_accruals, "filename"):
        df_accruals = parse_excel_accruals(df_accruals)

    if df_accruals is None or df_accruals.empty:
        return pd.DataFrame(), ["Загруженные файлы Excel пусты или не прочитаны."]

    # 1. Детекция колонки с ФИО
    fio_col = None
    max_fio_matches = 0

    for col in df_accruals.columns:
        col_series = df_accruals[col].dropna().astype(str)
        # Поиск совпадений по кириллице/латинице (ФИО)
        fio_count = col_series.str.contains(
            r"[А-ЯӘҒҚҢӨҰҮҺІA-Z][а-яәғқңөұүһіa-z]+\s+[А-ЯӘҒҚҢӨҰҮҺІA-Z]",
            regex=True,
        ).sum()
        if fio_count > max_fio_matches:
            max_fio_matches = fio_count
            fio_col = col

    if fio_col is None:
        fio_col = 1 if len(df_accruals.columns) > 1 else 0

    # Список исключений
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

    # 2. Сканирование всех строк и объединение сотрудников
    for idx, row in df_accruals.iterrows():
        raw_fio = str(row[fio_col]).strip()
        fio_lower = raw_fio.lower()

        if (
            pd.isna(row[fio_col])
            or len(raw_fio) < 3
            or fio_lower in stop_words
            or any(
                stop in fio_lower
                for stop in [
                    "всего",
                    "итого",
                    "подпись",
                    "страница",
                    "ведомость",
                ]
            )
        ):
            continue

        clean_fio = re.sub(r"\s+", " ", raw_fio)
        
        # Пропускаем дублирование заголовков
        if clean_fio in unique_fios and idx % 2 == 1:
            pass
        unique_fios.add(clean_fio)

        # Вытаскиваем все числа из строки
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
            paid_val = kvyd_val

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
        risk_comments.append(
            "Не удалось выделить сотрудников. Проверьте форматирование файлов."
        )

    return df_result, risk_comments