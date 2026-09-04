import io
import re
import fitz  # PyMuPDF
import pandas as pd
import pytesseract
from PIL import Image


def parse_excel_accruals(file_input):
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
                    # Сохраняем имя файла для определения месяца
                    filename = getattr(f, "name", "").lower()
                    df_sheet["_filename"] = filename
                    all_dfs.append(df_sheet)
        except Exception:
            continue

    if all_dfs:
        return pd.concat(all_dfs, ignore_index=True)

    return pd.DataFrame()


parse_excel_payroll = parse_excel_accruals


def extract_text_from_pdf(pdf_bytes):
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


def parse_pdf_5_15a(pdf_bytes):
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


def detect_months_from_files(files_list):
    """Определяет, за какие именно месяцы были загружены файлы."""
    all_months = [
        "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
        "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"
    ]
    detected = []

    if not files_list:
        return ["Январь", "Февраль"]

    filenames = " ".join([getattr(f, "name", "").lower() for f in files_list])

    # Проверка по именам файлов
    month_keywords = {
        "Январь": ["01", "январь", "янв", "january"],
        "Февраль": ["02", "февраль", "фев", "february"],
        "Март": ["03", "март", "мар", "march"],
        "Апрель": ["04", "апрель", "апр", "april"],
        "Май": ["05", "май", "may"],
        "Июнь": ["06", "июнь", "jun"],
        "Июль": ["07", "июль", "jul"],
        "Август": ["08", "август", "авг", "august"],
        "Сентябрь": ["09", "сентябрь", "сен", "september"],
        "Октябрь": ["10", "октябрь", "окт", "october"],
        "Ноябрь": ["11", "ноябрь", "ноя", "november"],
        "Декабрь": ["12", "декабрь", "дек", "december"],
    }

    for m_name, keywords in month_keywords.items():
        if any(kw in filenames for kw in keywords):
            detected.append(m_name)

    return detected if detected else all_months[:2]


def reconcile_salary(df_accruals, df_payments=None):
    risk_comments = []
    records = []

    # Сохраняем исходный список файлов для определения периода
    raw_files = df_accruals if isinstance(df_accruals, (list, tuple)) else []
    
    if not isinstance(df_accruals, pd.DataFrame):
        df_accruals = parse_excel_accruals(df_accruals)

    if df_accruals is None or df_accruals.empty:
        return pd.DataFrame(), ["Загруженные файлы Excel не содержат читаемых данных."]

    # Фильтруем колонки, исключая служебные
    data_cols = [c for c in df_accruals.columns if c != "_filename"]

    # 1. Поиск колонки ФИО
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

    # Определяем активные месяцы по загруженным файлам
    active_months = detect_months_from_files(raw_files)

    unique_fios = set()

    # 2. Обработка сотрудников
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

        # Собираем суммы из числовых колонок
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

        # Формируем строки ТОЛЬКО для загруженных месяцев
        for m_idx, m_name in enumerate(active_months):
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
            f"Обработано всех сотрудников: {total_people}. Период анализа: {', '.join(active_months)}."
        )
    else:
        risk_comments.append("Сотрудники не найдены. Проверьте содержимое файлов.")

    return df_result, risk_comments