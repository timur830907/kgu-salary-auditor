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

def clean_number(val):
    """
    Корректно очищает числа из Excel/PDF от неразрывных пробелов (\xa0),
    обычных пробелов и запятых в разделителях.
    """
    if pd.isna(val):
        return 0.0
    s = str(val).replace("\xa0", "").replace(" ", "").strip()
    s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def parse_excel_accruals(file_input):
    """
    Парсит загруженные Excel-файлы ведомостей.
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
                    df_sheet["_filename"] = getattr(f, "name", "").lower()
                    all_dfs.append(df_sheet)
        except Exception:
            continue

    return pd.concat(all_dfs, ignore_index=True) if all_dfs else pd.DataFrame()


def extract_text_from_pdf(pdf_file):
    """
    Извлекает текст из PDF-файла 5-15А. Использует Tesseract OCR, если текстовый слой отсутствует.
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
                pix = page.get_pixmap(dpi=150)
                img = Image.open(io.BytesIO(pix.tobytes("png")))
                try:
                    full_text += pytesseract.image_to_string(img, lang="rus+eng") + "\n"
                except Exception:
                    full_text += pytesseract.image_to_string(img) + "\n"
        return full_text
    except Exception:
        return ""


def parse_pdf_5_15a_payments(pdf_files):
    """
    Парсит PDF 5-15А и извлекает фактически переведенные суммы по сотрудникам и месяцам.
    """
    if not pdf_files:
        return pd.DataFrame(), []

    files_list = pdf_files if isinstance(pdf_files, (list, tuple)) else [pdf_files]
    extracted_records = []
    logs = []

    for f in files_list:
        fname = getattr(f, "name", "выписка").lower()
        text = extract_text_from_pdf(f)

        # Определение месяца по имени файла или тексту
        month_idx = 0
        for idx, m_name in enumerate(MONTH_NAMES_RU):
            if m_name.lower()[:3] in fname:
                month_idx = idx
                break

        lines = text.split("\n")
        for line in lines:
            line_str = line.strip()
            if not line_str:
                continue

            # Извлечение сумм
            amounts = re.findall(r"\b\d{1,3}(?:[\s,.]?\d{3})*(?:[.,]\d{2})?\b", line_str)
            valid_amounts = []
            for am in amounts:
                parsed = clean_number(am)
                if parsed > 0 and len(str(int(parsed))) < 10:
                    valid_amounts.append(parsed)

            if valid_amounts:
                # Извлечение ФИО (Фамилия И.О. или Полное имя)
                fio_match = re.search(
                    r"([А-ЯӘҒҚҢӨҰҮҺІA-Z][а-яәғқңөұүһіa-z]+\s+[А-ЯӘҒҚҢӨҰҮҺІA-Z]\.?(?:\s*[А-ЯӘҒҚҢӨҰҮҺІA-Z]\.?)?)",
                    line_str
                )
                fio = fio_match.group(1).strip() if fio_match else ""

                if fio:
                    extracted_records.append({
                        "fio": fio,
                        "amount": valid_amounts[-1],  # Итоговая сумма по строке
                        "month_idx": month_idx,
                        "month_name": MONTH_NAMES_RU[month_idx]
                    })

    return pd.DataFrame(extracted_records), logs


def detect_active_months(accruals_files, pdf_files):
    """
    Определяет список уникальных месяцев из загруженных файлов.
    """
    detected = set()
    all_files = []
    if accruals_files:
        all_files.extend(accruals_files if isinstance(accruals_files, list) else [accruals_files])
    if pdf_files:
        all_files.extend(pdf_files if isinstance(pdf_files, list) else [pdf_files])

    for f in all_files:
        name = getattr(f, "name", "").lower()
        for idx, m_name in enumerate(MONTH_NAMES_RU):
            if m_name.lower()[:3] in name or f"{idx+1:02d}" in name:
                detected.add(m_name)

    if not detected:
        return ["Январь", "Февраль", "Март"]

    return sorted(list(detected), key=lambda m: MONTH_NAMES_RU.index(m))


def reconcile_salary(accruals_files, pdf_files=None):
    """
    Сводит данные ведомостей из Excel и скан-выписок 5-15А из PDF.
    """
    risk_comments = []
    df_accruals = parse_excel_accruals(accruals_files)
    if df_accruals.empty:
        return pd.DataFrame(), ["Загруженные Excel-файлы ведомостей не содержат читаемых данных."]

    df_payments, pdf_logs = parse_pdf_5_15a_payments(pdf_files)
    risk_comments.extend(pdf_logs)

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

    stop_words = ["nan", "none", "фио", "ф.и.о.", "сотрудник", "наименование", "фамилия", "всего", "итого", "подпись", "ведомость"]
    active_months = detect_active_months(accruals_files, pdf_files)
    records = []
    unique_fios = set()

    for idx, row in df_accruals.iterrows():
        raw_fio = str(row[fio_col]).strip()
        fio_lower = raw_fio.lower()

        if pd.isna(row[fio_col]) or len(raw_fio) < 3 or any(stop in fio_lower for stop in stop_words):
            continue

        clean_fio = re.sub(r"\s+", " ", raw_fio)
        unique_fios.add(clean_fio)

        row_numbers = []
        for col_idx in [c for c in data_cols if c > fio_col]:
            val = row[col_idx]
            num = clean_number(val)
            if num > 0:
                row_numbers.append(num)

        curr_bal = 0.0
        surname = clean_fio.split()[0].upper()

        for m_idx, m_name in enumerate(active_months):
            kvyd_val = row_numbers[m_idx] if m_idx < len(row_numbers) else 0.0
            paid_val = 0.0

            if not df_payments.empty:
                # Четкий поиск выплат по фамилии и месяцу
                matched_payments = df_payments[
                    (df_payments["month_name"] == m_name) &
                    (df_payments["fio"].str.upper().str.contains(surname, na=False))
                ]
                if not matched_payments.empty:
                    paid_val = matched_payments["amount"].sum()

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
        risk_comments.append(f"Обработано сотрудников: {len(unique_fios)}. Период: {', '.join(active_months)}.")

    return df_result, risk_comments