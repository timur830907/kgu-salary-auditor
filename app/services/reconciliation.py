import io
import re
import fitz  # PyMuPDF
import pandas as pd
import pytesseract
from PIL import Image

ALL_MONTHS_RU = [
    "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
    "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"
]

MONTH_ALIASES = {
    "янв": "Январь", "01": "Январь", "январь": "Январь", "января": "Январь",
    "фев": "Февраль", "02": "Февраль", "февраль": "Февраль", "февраля": "Февраль",
    "мар": "Март", "03": "Март", "март": "Март", "марта": "Март",
    "апр": "Апрель", "04": "Апрель", "апрель": "Апрель", "апреля": "Апрель",
    "май": "Май", "05": "Май", "мая": "Май",
    "июн": "Июнь", "06": "Июнь", "июнь": "Июнь", "июня": "Июнь",
    "июл": "Июль", "07": "Июль", "июль": "Июль", "июля": "Июль",
    "авг": "Август", "08": "Август", "август": "Август", "августа": "Август",
    "сен": "Сентябрь", "09": "Сентябрь", "сентябрь": "Сентябрь", "сентября": "Сентябрь",
    "окт": "Октябрь", "10": "Октябрь", "октябрь": "Октябрь", "октября": "Октябрь",
    "ноя": "Ноябрь", "11": "Ноябрь", "ноябрь": "Ноябрь", "ноября": "Ноябрь",
    "дек": "Декабрь", "12": "Декабрь", "декабрь": "Декабрь", "декабря": "Декабрь"
}

def clean_number(val):
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
    if not fio_str:
        return ""
    clean = re.sub(r"[^А-Яа-яӘғқңөұүһіA-Za-z]", "", str(fio_str))
    return clean.upper()


def detect_month_from_text_or_name(text, fname):
    combined = (str(fname) + " " + str(text)).lower()
    for key, month_name in MONTH_ALIASES.items():
        if re.search(r"\b" + re.escape(key) + r"\b", combined):
            return month_name
    return "Январь"


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
        return pd.DataFrame()

    files_list = pdf_files if isinstance(pdf_files, (list, tuple)) else [pdf_files]
    extracted_records = []

    for f in files_list:
        fname = getattr(f, "name", "").lower()
        text = extract_text_from_pdf(f)
        m_name = detect_month_from_text_or_name(text[:500], fname)

        lines = text.split("\n")
        for line in lines:
            line_str = line.strip()
            if not line_str:
                continue

            amounts = re.findall(r"\b\d{1,3}(?:[\s,.]?\d{3})*(?:[.,]\d{2})?\b", line_str)
            valid_amounts = [clean_number(am) for am in amounts if clean_number(am) > 1000]

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
                        "month_name": m_name
                    })

    return pd.DataFrame(extracted_records)


def reconcile_salary(accruals_files, pdf_files=None):
    risk_comments = []
    files_list = accruals_files if isinstance(accruals_files, (list, tuple)) else [accruals_files]
    df_payments = parse_pdf_5_15a_payments(pdf_files)

    employee_data = {}
    active_months_set = set()
    stop_words = ["nan", "none", "фио", "ф.и.о.", "сотрудник", "наименование", "фамилия", "всего", "итого", "подпись"]

    for f in files_list:
        fname = getattr(f, "name", "").lower()

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

                # Определение месяца по содержимому листа
                sample_text = " ".join(df.iloc[:5].fillna("").astype(str).values.flatten())
                m_name = detect_month_from_text_or_name(sample_text, fname)
                active_months_set.add(m_name)

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

                for idx, row in df.iterrows():
                    raw_fio = str(row[fio_col]).strip()
                    if pd.isna(row[fio_col]) or len(raw_fio) < 3 or any(w in raw_fio.lower() for w in stop_words):
                        continue

                    clean_fio = re.sub(r"\s+", " ", raw_fio)

                    candidates = []
                    for c in range(fio_col + 1, len(row)):
                        val = clean_number(row[c])
                        if val > 1000:
                            candidates.append(val)

                    kvyd_val = max(candidates) if candidates else 0.0

                    if clean_fio not in employee_data:
                        employee_data[clean_fio] = {}

                    employee_data[clean_fio][m_name] = employee_data[clean_fio].get(m_name, 0.0) + kvyd_val

        except Exception:
            continue

    active_months = [m for m in ALL_MONTHS_RU if m in active_months_set]
    if not active_months:
        active_months = ALL_MONTHS_RU

    records = []

    for fio, months_dict in employee_data.items():
        fio_norm = normalize_fio(fio)
        surname = fio.split()[0].upper()
        running_balance = 0.0

        for m_name in active_months:
            start_bal = running_balance
            kvyd_val = months_dict.get(m_name, 0.0)
            paid_val = 0.0

            if not df_payments.empty:
                matched = df_payments[
                    (df_payments["month_name"] == m_name) &
                    ((df_payments["fio_norm"] == fio_norm) | (df_payments["surname"] == surname))
                ]
                if not matched.empty:
                    paid_val = matched["amount"].sum()

            end_bal = start_bal + kvyd_val - paid_val
            running_balance = end_bal

            if abs(end_bal) < 0.01:
                status = "Закрыто"
            elif end_bal < 0:
                status = "Переплата"
            else:
                status = "Недоплата"

            records.append({
                "fio": fio,
                "month": m_name,
                "start_bal": round(start_bal, 2),
                "kvyd": round(kvyd_val, 2),
                "paid": round(paid_val, 2),
                "end_bal": round(end_bal, 2),
                "status": status
            })

    df_result = pd.DataFrame(records)
    if not df_result.empty:
        risk_comments.append(f"Обработано сотрудников: {len(employee_data)}. Период: {', '.join(active_months)}.")

    return df_result, risk_comments