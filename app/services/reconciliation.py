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
    Корректно преобразует любые значения Excel/PDF в float,
    учитывая разделители тысяч, неразрывные пробелы и запятые.
    """
    if pd.isna(val) or val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    
    s = str(val).replace("\xa0", "").replace(" ", "").strip()
    s = s.replace(",", ".")
    # Извлекаем первое полноценное число из строки
    match = re.search(r"-?\d+(?:\.\d+)?", s)
    if match:
        try:
            return float(match.group(0))
        except ValueError:
            return 0.0
    return 0.0


def normalize_fio(fio_str):
    """
    Приводит ФИО к единому виду для сравнения (например: 'Ешкеева А.Т.' -> 'ЕШКЕЕВААТ')
    """
    if not fio_str:
        return ""
    clean = re.sub(r"[^А-Яа-яӘғқңөұүһіA-Za-z]", "", str(fio_str))
    return clean.upper()


def extract_text_from_pdf(pdf_file):
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
                    df_sheet["_filename"] = getattr(f, "name", "").lower()
                    all_dfs.append(df_sheet)
        except Exception:
            continue

    return pd.concat(all_dfs, ignore_index=True) if all_dfs else pd.DataFrame()


def parse_pdf_5_15a_payments(pdf_files):
    if not pdf_files:
        return pd.DataFrame(), []

    files_list = pdf_files if isinstance(pdf_files, (list, tuple)) else [pdf_files]
    extracted_records = []
    logs = []

    for f in files_list:
        fname = getattr(f, "name", "выписка").lower()
        text = extract_text_from_pdf(f)

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

            # Поиск всех сумм в строке
            amounts = re.findall(r"\b\d{1,3}(?:[\s,.]?\d{3})*(?:[.,]\d{2})?\b", line_str)
            valid_amounts = []
            for am in amounts:
                parsed = clean_number(am)
                if parsed > 0 and len(str(int(parsed))) < 10:
                    valid_amounts.append(parsed)

            if valid_amounts:
                fio_match = re.search(
                    r"([А-ЯӘҒҚҢӨҰҮҺІA-Z][а-яәғқңөұүһіa-z]+\s+[А-ЯӘҒҚҢӨҰҮҺІA-Z]\.?(?:\s*[А-ЯӘҒҚҢӨҰҮҺІA-Z]\.?)?)",
                    line_str
                )
                fio = fio_match.group(1).strip() if fio_match else ""

                if fio:
                    extracted_records.append({
                        "fio": fio,
                        "fio_norm": normalize_fio(fio),
                        "surname": fio.split()[0].upper(),
                        "amount": valid_amounts[-1],
                        "month_idx": month_idx,
                        "month_name": MONTH_NAMES_RU[month_idx]
                    })

    return pd.DataFrame(extracted_records), logs


def detect_active_months(accruals_files, pdf_files):
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
        return ["Январь", "Февраль", "Март", "Апрель"]

    return sorted(list(detected), key=lambda m: MONTH_NAMES_RU.index(m))


def reconcile_salary(accruals_files, pdf_files=None):
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
        fio_norm = normalize_fio(clean_fio)
        surname = clean_fio.split()[0].upper()
        unique_fios.add(clean_fio)

        row_numbers = []
        for col_idx in [c for c in data_cols if c > fio_col]:
            val = row[col_idx]
            num = clean_number(val)
            if num > 0:
                row_numbers.append(num)

        curr_bal = 0.0

        for m_idx, m_name in enumerate(active_months):
            kvyd_val = row_numbers[m_idx] if m_idx < len(row_numbers) else 0.0
            paid_val = 0.0

            if not df_payments.empty:
                # 1. Точное совпадение по нормированному ФИО и месяцу
                matched = df_payments[
                    (df_payments["month_name"] == m_name) &
                    (df_payments["fio_norm"] == fio_norm)
                ]
                # 2. Если не найдено — поиск по Фамилии
                if matched.empty:
                    matched = df_payments[
                        (df_payments["month_name"] == m_name) &
                        (df_payments["surname"] == surname)
                    ]

                if not matched.empty:
                    paid_val = matched["amount"].sum()

            end_bal = curr_bal + kvyd_val - paid_val
            
            # Логика статуса и риска переплаты
            if abs(end_bal) < 0.01:
                status = "Закрыто"
            elif end_bal < 0:
                status = "Переплата (Риск)"
            else:
                status = "Недоплата / Расхождение"

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