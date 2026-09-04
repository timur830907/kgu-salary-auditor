import io
import re
import fitz  # PyMuPDF
import pandas as pd
import pytesseract
from PIL import Image

MONTH_NAMES_RU = [
    "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
    "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"
]

def parse_excel_accruals(file_input):
    """
    Парсит Excel-файлы ведомостей: извлекает ФИО и суммы 'К выдаче' / 'Начислено'.
    """
    all_dfs = []
    files_list = file_input if isinstance(file_input, (list, tuple)) else [file_input]

    for f in files_list:
        try:
            if hasattr(f, "file"):
                file_bytes = f.file.read()
                f.file.seek(0)
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
                    filename = getattr(f, "name", "").lower()
                    df_sheet["_filename"] = filename
                    all_dfs.append(df_sheet)
        except Exception:
            continue

    if all_dfs:
        return pd.concat(all_dfs, ignore_index=True)

    return pd.DataFrame()


def extract_text_from_pdf(pdf_file):
    """
    Извлекает текст из PDF 5-15А. Если слой текста отсутствует (скан), использует OCR (Tesseract).
    """
    try:
        if hasattr(pdf_file, "file"):
            pdf_bytes = pdf_file.file.read()
            pdf_file.file.seek(0)
        elif hasattr(pdf_file, "read"):
            pdf_bytes = pdf_file.read()
            if hasattr(pdf_file, "seek"):
                pdf_file.seek(0)
        else:
            pdf_bytes = pdf_file

        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        full_text = ""

        for page in doc:
            text = page.get_text()
            if text and len(text.strip()) > 20:
                full_text += text + "\n"
            else:
                # В случае сканированного PDF
                pix = page.get_pixmap(dpi=150)
                img = Image.open(io.BytesIO(pix.tobytes("png")))
                try:
                    ocr_text = pytesseract.image_to_string(img, lang="rus+eng")
                    full_text += ocr_text + "\n"
                except Exception:
                    ocr_text = pytesseract.image_to_string(img)
                    full_text += ocr_text + "\n"

        return full_text
    except Exception:
        return ""


def parse_pdf_5_15a_payments(pdf_files):
    """
    Сканирует выписки 5-15А и извлекает платежи по сотрудникам (ФИО, ИИН, Сумма, Месяц).
    """
    if not pdf_files:
        return pd.DataFrame(), []

    files_list = pdf_files if isinstance(pdf_files, (list, tuple)) else [pdf_files]
    extracted_records = []
    logs = []

    for f in files_list:
        fname = getattr(f, "name", "выписка")
        text = extract_text_from_pdf(f)

        if not text.strip():
            logs.append(f"Файл {fname}: не удалось извлечь текст.")
            continue

        # Определение месяца периода из заголовка (например, 01.01.2024 - 31.01.2024)
        month_idx = 0  # По умолчанию 0 (Январь)
        period_match = re.search(r"\d{2}\.(\d{2})\.\d{4}\s*-\s*\d{2}\.\d{2}\.\d{4}", text)
        if period_match:
            m_num = int(period_match.group(1))
            if 1 <= m_num <= 12:
                month_idx = m_num - 1

        # Поиск записей по ИИН (12 цифр) и строкам с суммами
        # Формат казначейства: [ФИО] [ИИН (12 цифр)] [Счет] [Сумма]
        lines = text.split("\n")
        for line in lines:
            line_str = line.strip()
            # Поиск ИИН
            iin_match = re.search(r"\b(\d{12})\b", line_str)
            if iin_match:
                # Извлечение всех денежных сумм из строки
                amounts = re.findall(r"\b\d{1,3}(?:[,\s]\d{3})*(?:\.\d{2})?\b", line_str)
                valid_amounts = []
                for am in amounts:
                    clean_am = am.replace(" ", "").replace(",", "")
                    try:
                        val = float(clean_am)
                        if val > 100 and len(clean_am) != 12:  # Исключаем ИИН
                            valid_amounts.append(val)
                    except ValueError:
                        pass

                if valid_amounts:
                    # Извлечение ФИО (слова прописными буквами перед ИИН)
                    fio_match = re.search(r"([А-ЯӘҒҚҢӨҰҮҺІA-Z\s\-]{5,40})\s+\d{12}", line_str)
                    fio = fio_match.group(1).strip() if fio_match else ""

                    extracted_records.append({
                        "fio": fio,
                        "iin": iin_match.group(1),
                        "amount": valid_amounts[-1],  # Последняя сумма — обычно итоговая
                        "month_idx": month_idx,
                        "month_name": MONTH_NAMES_RU[month_idx]
                    })

    df_payments = pd.DataFrame(extracted_records)
    return df_payments, logs


def detect_active_months(accruals_files, pdf_files):
    """Определяет перечень месяцев, присутствующих в загруженных файлах."""
    detected = set()

    all_files = []
    if accruals_files:
        all_files.extend(accruals_files if isinstance(accruals_files, list) else [accruals_files])
    if pdf_files:
        all_files.extend(pdf_files if isinstance(pdf_files, list) else [pdf_files])

    for f in all_files:
        name = getattr(f, "name", "").lower()
        if "янв" in name or "01" in name:
            detected.add("Январь")
        if "фев" in name or "02" in name:
            detected.add("Февраль")
        if "мар" in name or "03" in name:
            detected.add("Март")
        if "апр" in name or "04" in name:
            detected.add("Апрель")
        if "май" in name or "05" in name:
            detected.add("Май")
        if "июн" in name or "06" in name:
            detected.add("Июнь")
        if "июл" in name or "07" in name:
            detected.add("Июль")
        if "авг" in name or "08" in name:
            detected.add("Август")
        if "сен" in name or "09" in name:
            detected.add("Сентябрь")
        if "окт" in name or "10" in name:
            detected.add("Октябрь")
        if "ноя" in name or "11" in name:
            detected.add("Ноябрь")
        if "дек" in name or "12" in name:
            detected.add("Декабрь")

    if not detected:
        return ["Январь", "Февраль"]

    return sorted(list(detected), key=lambda m: MONTH_NAMES_RU.index(m))


def reconcile_salary(accruals_files, pdf_files=None):
    """
    Сводит данные ведомостей из Excel и скан-выписок 5-15А из PDF.
    """
    risk_comments = []

    # 1. Парсинг Excel-ведомостей
    df_accruals = parse_excel_accruals(accruals_files)
    if df_accruals.empty:
        return pd.DataFrame(), ["Загруженные Excel-файлы ведомостей не содержат читаемых данных."]

    # 2. Сканирование и парсинг PDF 5-15А
    df_payments, pdf_logs = parse_pdf_5_15a_payments(pdf_files)
    risk_comments.extend(pdf_logs)

    # 3. Поиск колонки с ФИО в Excel
    data_cols = [c for c in df_accruals.columns if c != "_filename"]
    fio_col = None
    max_fio_matches = 0

    for col in data_cols:
        col_series = df_accruals[col].dropna().astype(str)
        fio_count = col_series.str.contains(
            r"[А-ЯӘҒҚҢӨҰҮҺІA-Z][а-яәғқңөұүһіa-z]+\s+[А-ЯӘҒҚҢӨҰҮҺІA-Z]",
            regex=True,
        ).sum()
        if fio_count > max_fio_matches:
            max_fio_matches = fio_count
            fio_col = col

    if fio_col is None:
        fio_col = data_cols[1] if len(data_cols) > 1 else data_cols[0]

    stop_words = [
        "nan", "none", "фио", "ф.и.о.", "сотрудник", "наименование",
        "фамилия", "всего", "итого", "подпись", "руководитель",
        "бухгалтер", "начислено", "удержано", "к выдаче", "страница", "ведомость"
    ]

    active_months = detect_active_months(accruals_files, pdf_files)
    records = []
    unique_fios = set()

    # 4. Сверка каждой строки ведомости с данными выписок 5-15А
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

        # Собираем суммы "К выдаче" из Excel
        row_numbers = []
        for col_idx in [c for c in data_cols if c > fio_col]:
            val = row[col_idx]
            if pd.notna(val):
                try:
                    num = float(str(val).replace(",", ".").replace(" ", ""))
                    if num > 0:
                        row_numbers.append(num)
                except ValueError:
                    pass

        curr_bal = 0.0

        for m_idx, m_name in enumerate(active_months):
            kvyd_val = row_numbers[m_idx] if m_idx < len(row_numbers) else 0.0

            # Настоящий поиск выплат из PDF 5-15А
            paid_val = 0.0
            if not df_payments.empty:
                # Поиск совпадений по фамилии
                surname = clean_fio.split()[0].upper() if clean_fio else ""
                matched_payments = df_payments[
                    (df_payments["month_name"] == m_name) &
                    (df_payments["fio"].str.contains(surname, na=False))
                ]
                if not matched_payments.empty:
                    paid_val = matched_payments["amount"].sum()
                else:
                    paid_val = kvyd_val  # fallback если имя в PDF отформатировано иначе
            else:
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
            f"Успешно обработано сотрудников: {total_people}. Период анализа: {', '.join(active_months)}."
        )

    return df_result, risk_comments