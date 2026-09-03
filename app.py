import io
import os
import re
import pandas as pd
import streamlit as st
import sys
from pathlib import Path

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
    page_title="Калькулятор ЗП и Аудит формы 5-15А",
    page_icon="🏛️",
    layout="wide",
)

st.title("🏛️ Калькулятор заработной платы и Аудит формы 5-15А")
st.caption("Расчет начислений и удержаний гражданских служащих (ПП РК № 1193) + финансовая сверка")

tab1, tab2 = st.tabs([
    "📊 Калькулятор начислений и удержаний",
    "🔍 Сверка ведомостей и 5-15А"
])

# =============================================================================
# Вкладка 1: Полный калькулятор (Категории А, Б, С, Д + Налоги)
# =============================================================================
with tab1:
    st.header("Расчет оклада, надбавок и удержаний по ПП РК № 1193")
    
    col_cat, col_stazh, col_rate = st.columns(3)
    with col_cat:
        category_group = st.selectbox(
            "Группа / Категория должности",
            [
                "Блоки А (Управленческий персонал)",
                "Блок B1 (Врачи, Профессорско-преподавательский состав)",
                "Блок B2 (Учителя, Врачи-специалисты, Методисты)",
                "Блок B3 (Педагоги, Средний медперсонал, Специалисты)",
                "Блок B4 (Ассистенты, Воспитатели, Техспециалисты)",
                "Блок C1 (Административный персонал)",
                "Блок C2 (Административный персонал)",
                "Блок C3 (Административный персонал)",
                "Блок D (Вспомогательный / Рабочий персонал - D1-D5)",
            ]
        )
    with col_stazh:
        stazh = st.number_input("Стаж работы (лет)", min_value=0, max_value=50, value=5)
    with col_rate:
        rate = st.number_input("Ставка (доля)", min_value=0.1, max_value=2.0, value=1.0, step=0.25)

    col_bdo, col_extra = st.columns(2)
    with col_bdo:
        bdo = st.number_input("Базовый должностной оклад (БДО), ₸", value=17697.0)
    with col_extra:
        st.subheader("Доплаты и условия")
        harmful_conditions = st.checkbox("Особые / вредные условия труда")
        class_guidance = st.checkbox("Классное руководство / проверка тетрадей")
        ecological_bonus = st.checkbox("Экологическая зона (экологическая надбавка)")

    if st.button("Рассчитать полный расчет (Начисления и Удержания)", type="primary", use_container_width=True):
        # Определение коэффициента
        coeff_base = 3.2
        if "Блок B1" in category_group or "А" in category_group:
            coeff_base = 4.2
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

        # Доплаты
        extra_pay = 0.0
        if harmful_conditions:
            extra_pay += base_salary * 0.15
        if class_guidance:
            extra_pay += base_salary * 0.20
        if ecological_bonus:
            extra_pay += base_salary * 0.10

        gross_salary = base_salary + extra_pay

        # Расчет удержаний (Ставки РК 2026)
        opv = gross_salary * 0.10               # ОПВ (10%)
        vosms = gross_salary * 0.02             # ВОСМС (2%)
        mzp = 85000.0                           # МЗП для вычета ИПН
        ipn = max(0.0, (gross_salary - opv - vosms - mzp) * 0.10) # ИПН (10%)
        
        total_deductions = opv + vosms + ipn
        net_salary = gross_salary - total_deductions

        # Налоги работодателя
        social_tax = max(0.0, gross_salary * 0.095 - (gross_salary - opv) * 0.035) # Соцналог
        social_deductions = (gross_salary - opv) * 0.035                           # Соцотчисления (3.5%)
        osms_employer = gross_salary * 0.03                                       # ОСМС работодателя (3%)
        opvr = gross_salary * 0.035                                                # ОПВР (3.5%)

        st.success("Расчет успешно выполнен!")
        
        res1, res2, res3 = st.columns(3)
        res1.metric("Должностной оклад", f"{base_salary:,.2f} ₸")
        res2.metric("Начисления (Gross)", f"{gross_salary:,.2f} ₸")
        res3.metric("К выплате на руки (Net)", f"{net_salary:,.2f} ₸", delta=f"-{total_deductions:,.2f} ₸ удержания")

        st.subheader("📌 Детализация удержаний с работника")
        d_col1, d_col2, d_col3, d_col4 = st.columns(4)
        d_col1.write(f"**ОПВ (10%):** {opv:,.2f} ₸")
        d_col2.write(f"**ВОСМС (2%):** {vosms:,.2f} ₸")
        d_col3.write(f"**ИПН (10%):** {ipn:,.2f} ₸")
        d_col4.write(f"**Всего удержано:** {total_deductions:,.2f} ₸")

        st.subheader("🏛️ Отчисления работодателя (КГУ)")
        e_col1, e_col2, e_col3, e_col4 = st.columns(4)
        e_col1.write(f"**Соц. отчисления (3.5%):** {social_deductions:,.2f} ₸")
        e_col2.write(f"**ОСМС (3%):** {osms_employer:,.2f} ₸")
        e_col3.write(f"**ОПВР (3.5%):** {opvr:,.2f} ₸")
        e_col4.write(f"**Соц. налог:** {social_tax:,.2f} ₸")

# =============================================================================
# Вкладка 2: Исправленная сверка ведомости и 5-15А
# =============================================================================
with tab2:
    st.header("Автоматическая сверка ведомости и формы 5-15А")
    st.write("Загрузите расчетно-платежную ведомость (Excel) и выписку 5-15А (PDF) для проверки.")

    col_up1, col_up2 = st.columns(2)
    
    with col_up1:
        st.subheader("1. Расчетно-платежная ведомость")
        uploaded_excel = st.file_uploader("Загрузите ведомость (.xlsx, .xls)", type=["xlsx", "xls"], key="excel_file")

    with col_up2:
        st.subheader("2. Выписка по форме 5-15А")
        uploaded_pdf = st.file_uploader("Загрузите выписку 5-15А (.pdf, .png, .jpg, .jpeg)", type=["pdf", "png", "jpg", "jpeg"], key="pdf_file")

    if st.button("🚀 Начать сверку данных", use_container_width=True):
        if not uploaded_excel or not uploaded_pdf:
            st.error("Пожалуйста, загрузите оба файла перед запуском сверки.")
        else:
            with st.spinner("Идет обработка файлов и извлечение данных..."):
                try:
                    # Оборачиваем байты в io.BytesIO(), чтобы pandas/PyMuPDF не выдавали ошибку типа 'bytes'
                    excel_stream = io.BytesIO(uploaded_excel.getvalue())
                    df_payroll = parse_excel_accruals(excel_stream)
                    
                    pdf_bytes = uploaded_pdf.getvalue()
                    pdf_text = extract_text_from_pdf(pdf_bytes)

                    st.success("Файлы успешно обработаны!")

                    st.subheader("📋 Извлеченные данные из Excel")
                    if not df_payroll.empty:
                        st.dataframe(df_payroll.head(20), use_container_width=True)
                    else:
                        st.warning("Не удалось автоматически распознать строки в Excel.")

                    st.subheader("📄 Распознанный текст из формы 5-15А (OCR)")
                    if pdf_text and len(pdf_text.strip()) > 0:
                        with st.expander("Показать извлеченный текст выписки 5-15А"):
                            st.text_area("Текст 5-15А", value=pdf_text, height=300)
                    else:
                        st.error("Не удалось извлечь текст из PDF/скана 5-15А.")

                except Exception as e:
                    st.error(f"Произошла ошибка при обработке файлов: {str(e)}")