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
    Преобразует любые значения из Excel/PDF в float,
    корректно обрабатывая неразрывные пробелы (\xa0),
    разделители тысяч и финансовые форматы.
    """
    if pd.isna(val) or val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    
    s = str(val).replace("\xa0", "").replace(" ", "").strip()
    s = s.replace(",", ".")
    match = re.search(r"-?\d+(?:\.\d+)?", s)
    if match:
        try:
            return float(match.group(0))
        except ValueError:
            return 0.0
    return 0.0


def normalize_fio(fio_str):
    """
    Нормализует ФИО для точного совпадения (удаляет спецсимволы и пробелы).
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

            amounts = re.findall(r"\b\d{1,3}(?:[\s,.]?\d{3})*(?:[.,]\d{2})?\b", line_str)
            valid_amounts = []
            for am in amounts:
                parsed = clean_number(am)
                if parsed > 1000:  # Игнорируем технические мелкие числа/номера страниц
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
    files_list = accruals_files if isinstance(accruals_files, (list, tuple)) else [accruals_files]
    
    df_payments, pdf_logs = parse_pdf_5_15a_payments(pdf_files)
    risk_comments.extend(pdf_logs)

    active_months = detect_active_months(accruals_files, pdf_files)
    records = []
    unique_fios = set()

    # Помесячная обработка каждого файла Excel
    for f in files_list:
        fname = getattr(f, "name", "").lower()
        m_name = active_months[0]
        for m in MONTH_NAMES_RU:
            if m.lower()[:3] in fname:
                m_name = m
                break

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
                df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
                if df.empty:
                    continue

                # Поиск колонки с ФИО
                fio_col = None
                max_fios = 0
                for col in df.columns:
                    cnt = df[col].dropna().astype(str).str.contains(
                        r"[А-ЯӘҒҚҢӨҰҮҺІA-Z][а-яәғқңөұүһіa-z]+\s+[А-ЯӘҒҚҢӨҰҮҺІA-Z]", regex=True
                    ).sum()
                    if cnt > max_fios:
                        max_fios = cnt
                        fio_col = col

                if fio_col is None:
                    continue

                stop_words = ["nan", "none", "фио", "ф.и.о.", "сотрудник", "наименование", "фамилия", "всего", "итого"]

                for idx, row in df.iterrows():
                    raw_fio = str(row[fio_col]).strip()
                    if pd.isna(row[fio_col]) or len(raw_fio) < 3 or any(w in raw_fio.lower() for w in stop_words):
                        continue

                    clean_fio = re.sub(r"\s+", " ", raw_fio)
                    fio_norm = normalize_fio(clean_fio)
                    surname = clean_fio.split()[0].upper()
                    unique_fios.add(clean_fio)

                    # Поиск реальной суммы "К выдаче" (берем максимальное/последнее крупное число правее ФИО)
                    candidate_amounts = []
                    for c in range(fio_col + 1, len(row)):
                        val = clean_number(row[c])
                        if val > 500:  # Отсекаем порядковые номера, дни, ставки (1, 2, 21, 0.5)
                            candidate_amounts.append(val)

                    kvyd_val = candidate_amounts[-1] if candidate_amounts else 0.0
                    paid_val = 0.0

                    if not df_payments.empty:
                        matched = df_payments[
                            (df_payments["month_name"] == m_name) &
                            (df_payments["fio_norm"] == fio_norm)
                        ]
                        if matched.empty:
                            matched = df_payments[
                                (df_payments["month_name"] == m_name) &
                                (df_payments["surname"] == surname)
                            ]
                        if not matched.empty:
                            paid_val = matched["amount"].sum()

                    end_bal = kvyd_val - paid_val
                    if abs(end_bal) < 0.01:
                        status = "Закрыто"
                    elif end_bal < 0:
                        status = "Переплата (Риск)"
                    else:
                        status = "Недоплата / Расхождение"

                    records.append({
                        "fio": clean_fio,
                        "month": m_name,
                        "start_bal": 0.0,
                        "kvyd": round(kvyd_val, 2),
                        "paid": round(paid_val, 2),
                        "end_bal": round(end_bal, 2),
                        "status": status
                    })
        except Exception:
            continue

    df_result = pd.DataFrame(records)
    if not df_result.empty:
        risk_comments.append(f"Обработано сотрудников: {len(unique_fios)}. Период: {', '.join(active_months)}.")

    return df_result, risk_comments