import io
import os
import re
import sys
from pathlib import Path
import pandas as pd
import streamlit as st

# Настройка путей импорта
BASE_DIR = Path(__file__).resolve().parent
sys.path.append(str(BASE_DIR))
sys.path.append(str(BASE_DIR / "app"))

try:
    from app.services.reconciliation import (
        parse_excel_accruals,
        extract_text_from_pdf,
    )
except ModuleNotFoundError:
    from services.reconciliation import (
        parse_excel_accruals,
        extract_text_from_pdf,
    )

st.set_page_config(
    page_title="Аудит ЗП и формы 5-15А (ПП РК № 1193)",
    page_icon="🏛️",
    layout="wide",
)

st.title("🏛️ Оплата труда гражданских служащих, работников организаций, содержащихся за счет средств государственного бюджета, работников казенных предприятий №1193")
st.caption("Расчет начислений, удержаний (ОПВ, ИПН, ВОСМС, ОПВР), разовая сверка и сквозная пофамильная ОСВ за 12 месяцев")

tab1, tab2, tab3 = st.tabs([
    "📊 Калькулятор начислений и удержаний (ПП РК № 1193)",
    "🔍 Разовая сверка ведомостей и 5-15А",
    "👤 Пофамильная сквозная ОСВ за 12 месяцев"
])

# =============================================================================
# Вкладка 1: Калькулятор
# =============================================================================
with tab1:
    st.header("Расчет оклада, надбавок и удержаний по ПП РК № 1193")
    
    col_cat, col_stazh, col_rate = st.columns(3)
    with col_cat:
        category_group = st.selectbox(
            "Группа / Категория должности",
            [
                "Блок A (Управленческий персонал: A1, A2, A3)",
                "Блок B1 (Врачи, Профессорско-преподавательский состав)",
                "Блок B2 (Учителя, Врачи-специалисты, Методисты)",
                "Блок B3 (Педагоги, Средний медперсонал, Специалисты)",
                "Блок B4 (Ассистенты, Воспитатели, Техспециалисты)",
                "Блок C1 (Административный персонал)",
                "Блок C2 (Административный персонал)",
                "Блок C3 (Административный персонал)",
                "Блок D (Вспомогательный / Рабочий персонал: D1, D2, D3, D4, D5)",
            ]
        )
    with col_stazh:
        stazh = st.number_input("Стаж работы (лет)", min_value=0, max_value=50, value=5)
    with col_rate:
        rate = st.number_input("Ставка (доля ставки)", min_value=0.1, max_value=2.0, value=1.0, step=0.25)

    col_bdo, col_qual = st.columns(2)
    with col_bdo:
        bdo = st.number_input("Базовый должностной оклад (БДО), ₸", value=17697.0)
    with col_qual:
        qual_category = st.selectbox(
            "Квалификационная категория / Статус",
            [
                "Без категории",
                "Педагог-модератор (+30% к ДО)",
                "Педагог-эксперт (+35% к ДО)",
                "Педагог-исследователь (+40% к ДО)",
                "Педагог-мастер (+50% к ДО)",
                "Высшая категория медперсонала (+30% к ДО)",
                "Первая категория медперсонала (+20% к ДО)",
                "Вторая категория медперсонала (+10% к ДО)"
            ]
        )

    st.subheader("Доплаты, надбавки и особые условия (ПП РК № 1193)")
    d_col1, d_col2, d_col3 = st.columns(3)
    
    with d_col1:
        harmful_conditions = st.checkbox("Особые / вредные условия труда (+10-30%)")
        class_guidance = st.checkbox("Классное руководство / заведование (+20-25%)")
        checking_notebooks = st.checkbox("Проверка письменных работ / тетрадей (+20-25%)")
        cabinet_management = st.checkbox("Заведование кабинетом / мастерской (+20%)")

    with d_col2:
        ecological_bonus = st.checkbox("Экологическая зона проживания/выплат")
        night_work = st.checkbox("Работа в ночное время / праздничные / выходные (1.5x)")
        combination_pay = st.checkbox("Совмещение должностей / расширение зоны (до 50%)")
        degree_bonus = st.checkbox("Ученая степень (Кандидат наук / PhD / Доктор)")

    with d_col3:
        mzp_value = st.number_input("Минимальная заработная плата (МЗП), ₸", value=85000.0)

    if st.button("Рассчитать полный расчет (Начисления и Удержания)", type="primary", use_container_width=True):
        coeff_base = 3.2
        if "Блок A" in category_group:
            coeff_base = 4.5
        elif "Блок B1" in category_group:
            coeff_base = 4.1
        elif "Блок B2" in category_group:
            coeff_base = 3.8
        elif "Блок B3" in category_group:
            coeff_base = 3.4
        elif "Блок B4" in category_group:
            coeff_base = 3.0
        elif "Блок C" in category_group:
            coeff_base = 2.8
        elif "Блок D" in category_group:
            coeff_base = 2.2

        coeff = coeff_base + (stazh * 0.04)
        base_salary = bdo * coeff * rate

        extra_pay = 0.0
        if harmful_conditions:
            extra_pay += base_salary * 0.15
        if class_guidance:
            extra_pay += base_salary * 0.25
        if checking_notebooks:
            extra_pay += base_salary * 0.20
        if cabinet_management:
            extra_pay += base_salary * 0.20
        if ecological_bonus:
            extra_pay += base_salary * 0.20
        if night_work:
            extra_pay += base_salary * 0.15
        if combination_pay:
            extra_pay += base_salary * 0.30
        if degree_bonus:
            extra_pay += base_salary * 0.35

        if "модератор" in qual_category:
            extra_pay += base_salary * 0.30
        elif "эксперт" in qual_category:
            extra_pay += base_salary * 0.35
        elif "исследователь" in qual_category:
            extra_pay += base_salary * 0.40
        elif "мастер" in qual_category:
            extra_pay += base_salary * 0.50
        elif "Высшая" in qual_category:
            extra_pay += base_salary * 0.30
        elif "Первая" in qual_category:
            extra_pay += base_salary * 0.20
        elif "Вторая" in qual_category:
            extra_pay += base_salary * 0.10

        gross_salary = base_salary + extra_pay

        opv = gross_salary * 0.10
        vosms = gross_salary * 0.02
        ipn = max(0.0, (gross_salary - opv - vosms - mzp_value) * 0.10)
        
        total_deductions = opv + vosms + ipn
        net_salary = gross_salary - total_deductions

        social_deductions = (gross_salary - opv) * 0.035
        social_tax = max(0.0, gross_salary * 0.095 - social_deductions)
        osms_employer = gross_salary * 0.03
        opvr = gross_salary * 0.035

        st.success("Расчет успешно выполнен!")
        
        res1, res2, res3 = st.columns(3)
        res1.metric("Должностной оклад (ДО)", f"{base_salary:,.2f} ₸")
        res2.metric("Начислено всего (Gross)", f"{gross_salary:,.2f} ₸")
        res3.metric("К выплате на руки (Net)", f"{net_salary:,.2f} ₸", delta=f"-{total_deductions:,.2f} ₸ удержания")

# =============================================================================
# Вкладка 2: Разовая сверка ведомостей и 5-15А
# =============================================================================
with tab2:
    st.header("Автоматическая сверка ведомостей и формы 5-15А")
    st.write("Загрузите ведомости (Excel) и выписку 5-15А (PDF) для разовой проверки.")

    col_up1, col_up2 = st.columns(2)
    with col_up1:
        uploaded_excels = st.file_uploader(
            "Загрузите ведомости (.xlsx, .xls)", 
            type=["xlsx", "xls"], 
            accept_multiple_files=True,
            key="excel_files_uploader"
        )
    with col_up2:
        uploaded_pdf = st.file_uploader(
            "Загрузите выписку 5-15А (.pdf, .png, .jpg, .jpeg)", 
            type=["pdf", "png", "jpg", "jpeg"], 
            key="pdf_file_uploader"
        )

    if st.button("🚀 Начать автоматическую сверку", use_container_width=True):
        if not uploaded_excels or not uploaded_pdf:
            st.error("Пожалуйста, загрузите ведомости Excel и выписку 5-15А.")
        else:
            with st.spinner("Анализ данных..."):
                try:
                    all_dfs = []
                    total_payout = 0.0
                    for excel in uploaded_excels:
                        df = parse_excel_accruals(io.BytesIO(excel.getvalue()))
                        all_dfs.append(df)
                        for col in df.columns:
                            if any(k in str(col).lower() for k in ["к выплате", "выплате", "на руки"]):
                                total_payout += pd.to_numeric(df[col], errors='coerce').sum()

                    pdf_text = extract_text_from_pdf(uploaded_pdf.getvalue())
                    clean_t = pdf_text.replace('\n', ' ').replace('\r', ' ')
                    matches = re.findall(r'\b\d{1,3}(?:[\s\.]?\d{3})*(?:[,\.]\d{2})?\b', clean_t)
                    valid_amounts = [float(re.sub(r'[^\d.]', '', m.replace(',', '.'))) for m in matches if re.sub(r'[^\d.]', '', m.replace(',', '.'))]
                    valid_amounts = [v for v in valid_amounts if 1000.0 <= v <= 500000000.0]
                    sum_5_15a = max(valid_amounts) if valid_amounts else 0.0

                    m1, m2, m3 = st.columns(3)
                    m1.metric("Сумма по ведомостям (к выплате)", f"{total_payout:,.2f} ₸")
                    m2.metric("Сумма по 5-15А", f"{sum_5_15a:,.2f} ₸")
                    diff = abs(total_payout - sum_5_15a)
                    m3.metric("Разница", f"{diff:,.2f} ₸")
                except Exception as e:
                    st.error(f"Ошибка: {str(e)}")

# =============================================================================
# Вкладка 3: Пофамильная сквозная ОСВ за 12 месяцев
# =============================================================================
with tab3:
    st.header("👤 Пофамильная сквозная оборотно-сальдовая ведомость (12 месяцев)")
    st.write(
        "Загрузите ведомости (Excel) и выписки 5-15А (PDF) за разные месяцы. "
        "Система объединит **всех сотрудников без исключения**, привяжет данные к месяцам и рассчитает последовательный перенос остатков."
    )

    MONTH_NAMES = [
        "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
        "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"
    ]

    def get_month_index(filename):
        fn = filename.lower()
        if "01" in fn or "янв" in fn: return 0
        if "02" in fn or "фев" in fn: return 1
        if "03" in fn or "мар" in fn: return 2
        if "04" in fn or "апр" in fn: return 3
        if "05" in fn or "май" in fn or "мая" in fn: return 4
        if "06" in fn or "июн" in fn: return 5
        if "07" in fn or "июл" in fn: return 6
        if "08" in fn or "авг" in fn: return 7
        if "09" in fn or "сен" in fn: return 8
        if "10" in fn or "окт" in fn: return 9
        if "11" in fn or "ноя" in fn: return 10
        if "12" in fn or "дек" in fn: return 11
        return 0

    col_files1, col_files2 = st.columns(2)
    with col_files1:
        yearly_excels = st.file_uploader(
            "1. Загрузите ведомости за все месяцы (.xlsx, .xls)",
            type=["xlsx", "xls"],
            accept_multiple_files=True,
            key="yearly_excels_files_v2"
        )
    with col_files2:
        yearly_pdfs = st.file_uploader(
            "2. Загрузите выписки 5-15А за все месяцы (.pdf)",
            type=["pdf"],
            accept_multiple_files=True,
            key="yearly_pdfs_files_v2"
        )

    if st.button("📊 Построить сквозную пофамильную ОСВ", type="primary", use_container_width=True):
        if not yearly_excels:
            st.error("Загрузите файлы ведомостей!")
        else:
            with st.spinner("Анализ всех сотрудников и расчет цепочки остатков..."):
                try:
                    # Словарь для сумм ведомостей: {ФИО: [м1, м2, ..., м12]}
                    employee_ved = {}
                    # Массив кассовых выплат по 5-15А по месяцам [м1, м2, ..., м12]
                    monthly_515a_totals = [0.0] * 12

                    # 1. Чтение ведомостей (Все сотрудники)
                    for excel_file in yearly_excels:
                        m_idx = get_month_index(excel_file.name)
                        df = parse_excel_accruals(io.BytesIO(excel_file.getvalue()))
                        
                        fio_col = None
                        payout_col = None
                        for c in df.columns:
                            c_str = str(c).lower()
                            if any(k in c_str for k in ["фио", "работник", "сотрудник", "ф.и.о", "наименование"]):
                                fio_col = c
                            if any(k in c_str for k in ["к выплате", "выплате", "на руки", "сумма к выплате"]):
                                payout_col = c

                        if fio_col is None and len(df.columns) > 0:
                            fio_col = df.columns[0]
                        if payout_col is None:
                            num_cols = df.select_dtypes(include=['number']).columns
                            if len(num_cols) > 0:
                                payout_col = num_cols[-1]

                        if fio_col and payout_col:
                            for _, row in df.iterrows():
                                name = str(row[fio_col]).strip()
                                val = pd.to_numeric(row[payout_col], errors='coerce')
                                # Включаем абсолютно всех людей, отсекаем только явные пустые строки и «Итого»
                                if pd.notna(val) and len(name) > 1 and "итого" not in name.lower() and "всего" not in name.lower():
                                    if name not in employee_ved:
                                        employee_ved[name] = [0.0] * 12
                                    employee_ved[name][m_idx] += float(val)

                    # 2. Чтение 5-15А
                    if yearly_pdfs:
                        for pdf_file in yearly_pdfs:
                            m_idx = get_month_index(pdf_file.name)
                            p_text = extract_text_from_pdf(pdf_file.getvalue())
                            clean_t = p_text.replace('\n', ' ').replace('\r', ' ')
                            matches = re.findall(r'\b\d{1,3}(?:[\s\.]?\d{3})*(?:[,\.]\d{2})?\b', clean_t)
                            vals = []
                            for m in matches:
                                try:
                                    v = float(re.sub(r'[^\d.]', '', m.replace(',', '.')))
                                    if 1000.0 <= v <= 500000000.0:
                                        vals.append(v)
                                except ValueError:
                                    pass
                            if vals:
                                monthly_515a_totals[m_idx] += max(vals)

                    # 3. Расчет помесячной цепочки остатков для каждого сотрудника
                    all_rows = []
                    for emp_name, monthly_accruals in employee_ved.items():
                        current_start = 0.0
                        for m_idx in range(12):
                            accrued = monthly_accruals[m_idx]
                            
                            # Распределение кассы 5-15А пропорционально начислениям месяца
                            month_total_ved = sum(emp_vals[m_idx] for emp_vals in employee_ved.values())
                            if month_total_ved > 0 and monthly_515a_totals[m_idx] > 0:
                                ratio = min(1.0, monthly_515a_totals[m_idx] / month_total_ved)
                                transferred = accrued * ratio
                            else:
                                transferred = accrued  # если 5-15А за месяц не загружена

                            end_bal = current_start + accrued - transferred
                            
                            # Показываем записи только для активных месяцев (где есть движения или остатки)
                            if accrued > 0 or transferred > 0 or abs(current_start) > 0.01 or abs(end_bal) > 0.01:
                                all_rows.append({
                                    "ФИО сотрудника": emp_name,
                                    "Месяц": MONTH_NAMES[m_idx],
                                    "Остаток на начало месяца, ₸": round(current_start, 2),
                                    "К выдаче (Ведомость), ₸": round(accrued, 2),
                                    "Перечислено (5-15А), ₸": round(transferred, 2),
                                    "Остаток на конец месяца, ₸": round(end_bal, 2),
                                    "Статус": "✅ Закрыто" if abs(end_bal) < 0.01 else ("🚨 Долг" if end_bal > 0 else "🚨 Переплата")
                                })

                            # Перенос остатка в следующий месяц
                            current_start = end_bal

                    result_df = pd.DataFrame(all_rows)

                    st.success(f"Обработано всех сотрудников: {len(employee_ved)}. Построена сквозная цепочка за 12 месяцев.")
                    st.subheader("📋 Детализированная оборотно-сальдовая ведомость")
                    st.dataframe(result_df, use_container_width=True)

                except Exception as e:
                    st.error(f"Ошибка обработки: {str(e)}")