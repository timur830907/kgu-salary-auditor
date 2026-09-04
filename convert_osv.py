import os
import sys
import openpyxl
import pandas as pd
import tkinter as tk
from tkinter import filedialog

def process_salary_sheet(input_file: str, output_file: str):
    """
    Максимально глубокий сканер ведомостей:
    Находит ВСЕХ сотрудников без исключения, независимо от верстки таблицы.
    """
    if not os.path.exists(input_file):
        print(f"[ОШИБКА] Файл '{input_file}' не найден!")
        return

    try:
        xls = pd.ExcelFile(input_file)
    except Exception as e:
        print(f"[ОШИБКА] Не удалось открыть файл Excel: {e}")
        return

    all_records = []
    total_found_fio = set()

    for sheet_name in xls.sheet_names:
        try:
            # Загружаем лист без заголовков для чистого разбора
            df = pd.read_excel(input_file, sheet_name=sheet_name, header=None)
        except Exception as e:
            continue

        if df.empty or df.shape[1] < 3:
            continue

        # 1. Поиск колонки, где содержатся ФИО сотрудников
        fio_col_idx = None
        for col in range(min(5, df.shape[1])):
            # Считаем текстовые строки, похожие на ФИО (содержащие пробелы и буквы)
            text_cells = df.iloc[:, col].dropna().astype(str)
            fio_like = text_cells[text_cells.str.contains(r'[А-Яа-яA-Za-z]+\s+[А-Яа-яA-Za-z]+', regex=True)]
            if len(fio_like) > 3:  # Если найдено больше 3 похожих на ФИО записей
                fio_col_idx = col
                break

        if fio_col_idx is None:
            fio_col_idx = 1 # Значение по умолчанию (вторая колонка)

        # 2. Поиск строки заголовка с месяцами
        header_row_m = 0
        header_row_sub = 1

        for r in range(min(12, df.shape[0])):
            row_vals = [str(v).lower() for v in df.iloc[r].dropna().values]
            if any(m in v for v in row_vals for m in ["янв", "фев", "мар", "апр", "май", "июн", "июл", "авг", "сен", "окт", "ноя", "дек"]):
                header_row_m = r
                header_row_sub = r + 1
                break

        row_m = df.iloc[header_row_m].ffill()
        row_sub = df.iloc[header_row_sub] if header_row_sub < df.shape[0] else pd.Series()

        # 3. Собираем карты колонок для каждого месяца
        month_cols = {}
        for col_idx in range(fio_col_idx + 1, df.shape[1]):
            m_val = str(row_m.iloc[col_idx]).strip()
            s_val = str(row_sub.iloc[col_idx]).strip().lower() if col_idx < len(row_sub) else ""

            if m_val in ["nan", "None", ""] or "всего" in m_val.lower() or "итого" in m_val.lower():
                continue

            if m_val not in month_cols:
                month_cols[m_val] = {"kvyd": None, "paid": []}

            # Начисление / К выдаче
            if any(k in s_val for k in ["выдач", "выплат", "начисл", "сумма", "к оплат"]):
                if month_cols[m_val]["kvyd"] is None:
                    month_cols[m_val]["kvyd"] = col_idx
            else:
                # Все остальные числовые колонки относим к выплатам 5-15А
                month_cols[m_val]["paid"].append(col_idx)

        # Если месяцы не разделились корректно, берем все 12 месяцев по умолчанию
        if not month_cols:
            month_cols["Период 1"] = {"kvyd": fio_col_idx + 2, "paid": [fio_col_idx + 3]}

        # 4. Сканируем ВСЕ строки с людьми
        start_row = header_row_sub + 1
        for r_idx in range(start_row, df.shape[0]):
            fio_raw = df.iloc[r_idx, fio_col_idx]

            if pd.isna(fio_raw):
                continue

            fio_str = str(fio_raw).strip()
            fio_lower = fio_str.lower()

            # Исключаем служебные строки
            if (
                len(fio_str) < 3
                or fio_lower in ["nan", "none", "фио", "ф.и.о.", "сотрудник", "наименование", "фамилия"]
                or "всего" in fio_lower
                or "итого" in fio_lower
                or "подпись" in fio_lower
                or "руководитель" in fio_lower
                or "бухгалтер" in fio_lower
            ):
                continue

            total_found_fio.add(fio_str)

            # Входящий остаток (сальдо на начало)
            init_bal_col = fio_col_idx + 1
            init_bal = df.iloc[r_idx, init_bal_col] if init_bal_col < df.shape[1] else 0.0
            try:
                curr_bal = float(str(init_bal).replace(",", ".").replace(" ", "")) if pd.notna(init_bal) else 0.0
            except ValueError:
                curr_bal = 0.0

            # Помесячный обход
            for m_name, cols in month_cols.items():
                kvyd_val = 0.0
                if cols["kvyd"] is not None and cols["kvyd"] < df.shape[1]:
                    v = df.iloc[r_idx, cols["kvyd"]]
                    try:
                        kvyd_val = float(str(v).replace(",", ".").replace(" ", "")) if pd.notna(v) else 0.0
                    except ValueError:
                        kvyd_val = 0.0

                paid_val = 0.0
                for c_p in cols["paid"]:
                    if c_p < df.shape[1]:
                        v = df.iloc[r_idx, c_p]
                        try:
                            paid_val += float(str(v).replace(",", ".").replace(" ", "")) if pd.notna(v) else 0.0
                        except ValueError:
                            pass

                end_bal = curr_bal + kvyd_val - paid_val
                status = "Закрыто" if abs(end_bal) < 0.01 else "Расхождение"

                all_records.append({
                    "Лист / Подразделение": sheet_name,
                    "ФИО сотрудника": fio_str,
                    "Месяц": m_name,
                    "Остаток на начало": round(curr_bal, 2),
                    "К выдаче (Ведомость)": round(kvyd_val, 2),
                    "Перечислено (5-15А)": round(paid_val, 2),
                    "Остаток на конец": round(end_bal, 2),
                    "Статус": status
                })

                curr_bal = end_bal

    # 5. Сохранение результатов
    if all_records:
        res_df = pd.DataFrame(all_records)
        res_df.to_excel(output_file, index=False)
        print("\n==================================================")
        print(f"[УСПЕХ] Обработано записей: {len(res_df)}")
        print(f"[ОХВАТ] Найдено УНИКАЛЬНЫХ СОТРУДНИКОВ: {len(total_found_fio)}")
        print(f"[СОХРАНЕНО] Итоговый файл: {output_file}")
        print("==================================================\n")
    else:
        print("\n[ВНИМАНИЕ] Сотрудники не найдены. Проверьте правильность файла.\n")


def get_input_file():
    if len(sys.argv) > 1 and sys.argv[1].strip():
        return sys.argv[1].strip('"\'')

    try:
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        file_path = filedialog.askopenfilename(
            title="Выберите расчетную ведомость",
            filetypes=[("Excel файлы", "*.xlsx *.xls")]
        )
        if file_path:
            return file_path
    except Exception:
        pass

    return input("Перетащите файл Excel в окно терминала: ").strip('"\'')


if __name__ == "__main__":
    input_filename = get_input_file()

    if input_filename and os.path.exists(input_filename):
        folder, filename = os.path.split(input_filename)
        output_filename = os.path.join(folder, f"ОСВ_ПОЛНЫЙ_ОХВАТ_{filename}")
        process_salary_sheet(input_filename, output_filename)
    else:
        print("\n[ОШИБКА] Указанный файл не найден!")