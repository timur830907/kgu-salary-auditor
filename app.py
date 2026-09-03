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

# -----------------------------------------------------------------------------
# Настройка страницы Streamlit
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Аудит ЗП и формы 5-15А (ПП РК № 1193)",
    page_icon="🏛️",
    layout="wide",
)

st.title("🏛️ Оплата труда гражданских служащих, работников организаций, содержащихся за счет средств государственного бюджета, работников казенных предприятий №1193")
st.caption("Расчет начислений, удержаний (ОПВ, ИПН, ВОСМС, ОПВР) и финансовая сверка ведомостей с формой 5-15А")

tab1, tab2 = st.tabs([
    "📊 Калькулятор начислений и удержаний (ПП РК № 1193)",
    "🔍 Сверка ведомостей и формы 5-15А (Аудит рисков)"
])

# =============================================================================
# Вкладка 1: Полный калькулятор (Категории А, Б, С, Д + Расширенные Доплаты)
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
        # Базовый коэффициент по сетке ПП РК № 1193
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

        # Доплаты от оклада
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

        # Доплата за категорию
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

        # Удержания
        opv = gross_salary * 0.10               # ОПВ (10%)
        vosms = gross_salary * 0.02             # ВОСМС (2%)
        ipn = max(0.0, (gross_salary - opv - vosms - mzp_value) * 0.10) # ИПН (10%)
        
        total_deductions = opv + vosms + ipn
        net_salary = gross_salary - total_deductions

        # Налоги работодателя
        social_deductions = (gross_salary - opv) * 0.035
        social_tax = max(0.0, gross_salary * 0.095 - social_deductions)
        osms_employer = gross_salary * 0.03
        opvr = gross_salary * 0.035             # ОПВР (3.5% с 2026 года)

        st.success("Расчет успешно выполнен!")
        
        res1, res2, res3 = st.columns(3)
        res1.metric("Должностной оклад (ДО)", f"{base_salary:,.2f} ₸")
        res2.metric("Начислено всего (Gross)", f"{gross_salary:,.2f} ₸")
        res3.metric("К выплате на руки (Net)", f"{net_salary:,.2f} ₸", delta=f"-{total_deductions:,.2f} ₸ удержания")

        st.subheader("📌 Начисления и Удержания с работника")
        d_c1, d_c2, d_c3, d_c4 = st.columns(4)
        d_c1.write(f"**ОПВ (10%):** {opv:,.2f} ₸")
        d_c2.write(f"**ВОСМС (2%):** {vosms:,.2f} ₸")
        d_c3.write(f"**ИПН (10%):** {ipn:,.2f} ₸")
        d_c4.write(f"**Всего удержаний:** {total_deductions:,.2f} ₸")

        st.subheader("🏛️ Начисления и отчисления за счет работодателя (КГУ)")
        e_c1, e_c2, e_c3, e_c4 = st.columns(4)
        e_c1.write(f"**Социальные отчисления (3.5%):** {social_deductions:,.2f} ₸")
        e_c2.write(f"**ОСМС работодателя (3%):** {osms_employer:,.2f} ₸")
        e_c3.write(f"**ОПВР работодателя (3.5%):** {opvr:,.2f} ₸")
        e_c4.write(f"**Социальный налог:** {social_tax:,.2f} ₸")

# =============================================================================
# Вкладка 2: Сверка нескольких ведомостей и формы 5-15А с проверкой рисков
# =============================================================================
with tab2:
    st.header("Автоматическая сверка ведомостей и формы 5-15А")
    st.write("Загрузите одна или несколько расчетно-платежных ведомостей (Excel) и выписку 5-15А (PDF) для поиска расхождений и оценки аудиторских рисков.")

    col_up1, col_up2 = st.columns(2)
    
    with col_up1:
        st.subheader("1. Расчетно-платежные ведомости")
        uploaded_excels = st.file_uploader(
            "Загрузите ведомости (.xlsx, .xls) — можно выбрать сразу несколько", 
            type=["xlsx", "xls"], 
            accept_multiple_files=True,
            key="excel_files"
        )

    with col_up2:
        st.subheader("2. Выписка по форме 5-15А")
        uploaded_pdf = st.file_uploader(
            "Загрузите выписку 5-15А (.pdf, .png, .jpg, .jpeg)", 
            type=["pdf", "png", "jpg", "jpeg"], 
            key="pdf_file"
        )

    if st.button("🚀 Начать автоматическую сверку и поиск рисков", use_container_width=True):
        if not uploaded_excels or not uploaded_pdf:
            st.error("Пожалуйста, загрузите ведомости (хотя бы 1 файл Excel) и выписку 5-15А перед запуском сверки.")
        else:
            with st.spinner("Анализ ведомостей, распознавание 5-15А и поиск искажений..."):
                try:
                    all_dfs = []
                    total_payroll_accrued = 0.0
                    total_payroll_payout = 0.0

                    # 1. Обработка всех загруженных ведомостей Excel
                    for uploaded_excel in uploaded_excels:
                        excel_stream = io.BytesIO(uploaded_excel.getvalue())
                        df = parse_excel_accruals(excel_stream)
                        all_dfs.append(df)

                        # Суммирование численных значений
                        for col in df.columns:
                            col_name = str(col).lower()
                            if any(k in col_name for k in ["всего начислено", "начислено", "итого"]):
                                total_payroll_accrued += pd.to_numeric(df[col], errors='coerce').sum()
                            if any(k in col_name for k in ["к выплате", "выплате", "на руки", "выплата"]):
                                total_payroll_payout += pd.to_numeric(df[col], errors='coerce').sum()

                    # 2. Извлечение текста из PDF 5-15А
                    pdf_bytes = uploaded_pdf.getvalue()
                    pdf_text = extract_text_from_pdf(pdf_bytes)

                    # Поиск итоговых сумм в тексте 5-15А с помощью регулярных выражений
                    found_numbers = re.findall(r'\b\d{1,3}(?:\s?\d{3})*(?:\.\d{1,2})?\b', pdf_text)
                    numbers_clean = []
                    for num in found_numbers:
                        clean_n = float(num.replace(" ", ""))
                        if clean_n > 1000:
                            numbers_clean.append(clean_n)

                    sum_5_15a = max(numbers_clean) if numbers_clean else 0.0

                    st.success(f"Обработано файлов ведомостей: {len(uploaded_excels)}. Сверка завершена!")

                    # Отображение сводных показателей
                    st.subheader("⚖️ Результаты финансовой сверки")
                    
                    m1, m2, m3 = st.columns(3)
                    m1.metric("Сумма по ведомостям (к выплате)", f"{total_payroll_payout:,.2f} ₸")
                    m2.metric("Сумма по форме 5-15А", f"{sum_5_15a:,.2f} ₸")
                    
                    diff = abs(total_payroll_payout - sum_5_15a)
                    m3.metric("Расхождение / Искажение", f"{diff:,.2f} ₸", delta=f"{'-' if diff > 0 else ''}{diff:,.2f} ₸", delta_color="inverse")

                    # Оценка риска
                    st.subheader("🚨 Аудиторская оценка рисков")
                    if diff > 1.0 or sum_5_15a == 0.0:
                        st.error(f"⚠️ **ВНИМАНИЕ: ВЫЯВЛЕН РИСК ИСКАЖЕНИЯ И НЕОБОСНОВАННОГО РАСХОЖДЕНИЯ!**")
                        st.warning(
                            f"**Комментарий аудитора:** Итоговая сумма по загруженным ведомостям ({total_payroll_payout:,.2f} ₸) "
                            f"не совпадает с выпиской формы 5-15А ({sum_5_15a:,.2f} ₸). "
                            f"Разница составляет **{diff:,.2f} ₸**. "
                            f"Рекомендуется проверить полноту загруженных ведомостей, корректность применения КБК, а также наличие нераспознанных или удержанных сумм."
                        )
                    else:
                        st.success("✅ **РИСКОВ НЕ ВЫЯВЛЕНО:** Данные расчетных ведомостей и формы 5-15А полностью сходятся!")

                    # Вывод сводной таблицы из Excel
                    st.subheader("📋 Объединенные данные из ведомостей")
                    if all_dfs:
                        combined_df = pd.concat(all_dfs, ignore_index=True)
                        st.dataframe(combined_df.head(30), use_container_width=True)

                    # Вывод распознанного текста 5-15А
                    st.subheader("📄 Извлеченный текст выписки 5-15А (OCR)")
                    if pdf_text and len(pdf_text.strip()) > 0:
                        with st.expander("Показать полный текст из формы 5-15А"):
                            st.text_area("Текст 5-15А", value=pdf_text, height=300)

                except Exception as e:
                    st.error(f"Произошла ошибка при обработке данных: {str(e)}")