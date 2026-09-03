import sys
from pathlib import Path

# Добавляем корневую директорию и папку app в sys.path для корректных импортов на Render
BASE_DIR = Path(__file__).resolve().parent
sys.path.append(str(BASE_DIR))
sys.path.append(str(BASE_DIR / "app"))

import io
import os
import re
import pandas as pd
import streamlit as st

# Импорт из модуля services
try:
    from app.services.reconciliation import parse_excel_accruals, extract_text_from_pdf
except ModuleNotFoundError:
    from services.reconciliation import parse_excel_accruals, extract_text_from_pdf

# -----------------------------------------------------------------------------
# Конфигурация страницы Streamlit
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Калькулятор ЗП и Аудит формы 5-15А",
    page_icon="🏛️",
    layout="wide",
)

st.title("🏛️ Калькулятор заработной платы и Аудит формы 5-15А")
st.caption("Система расчета гражданских служащих (ПП РК № 1193) и финансовая сверка")

# -----------------------------------------------------------------------------
# Вкладки приложения
# -----------------------------------------------------------------------------
tab1, tab2 = st.tabs([
    "📊 Калькулятор начислений (ПП РК № 1193)",
    "🔍 Сверка ведомостей и 5-15А"
])

# =============================================================================
# Вкладка 1: Калькулятор начислений
# =============================================================================
with tab1:
    st.header("Расчет оклада и начислений работникам КГУ")
    
    col1, col2 = st.columns(2)
    with col1:
        category = st.selectbox(
            "Категория должности",
            ["B1", "B2", "B3", "B4", "C1", "C2", "C3", "D1", "D2", "D3"]
        )
        stazh = st.number_input("Стаж работы (лет)", min_value=0, max_value=50, value=5)
        bdo = st.number_input("Базовый должностной оклад (БДО), тенге", value=17697.0)

    with col2:
        rate = st.number_input("Ставка (доля ставки)", min_value=0.1, max_value=2.0, value=1.0, step=0.25)
        harmful_conditions = st.checkbox("Особые / вредные условия труда (+10-30%)")
        class_guidance = st.checkbox("Классное руководство / заведование")

    if st.button("Рассчитать начисления", type="primary"):
        coeff_map = {"B1": 4.5, "B2": 4.1, "B3": 3.8, "B4": 3.5, "C1": 3.2, "C2": 3.0, "C3": 2.8, "D1": 2.5, "D2": 2.3, "D3": 2.1}
        coeff = coeff_map.get(category, 3.0) + (stazh * 0.05)
        
        base_salary = bdo * coeff * rate
        extra_pay = base_salary * 0.10 if harmful_conditions else 0.0
        if class_guidance:
            extra_pay += base_salary * 0.15
            
        total_accrual = base_salary + extra_pay
        
        st.success("Расчет успешно выполнен!")
        res_col1, res_col2, res_col3 = st.columns(3)
        res_col1.metric("Должностной оклад", f"{base_salary:,.2f} ₸")
        res_col2.metric("Надбавки и доплаты", f"{extra_pay:,.2f} ₸")
        res_col3.metric("Итого начислено", f"{total_accrual:,.2f} ₸")

# =============================================================================
# Вкладка 2: Автоматическая сверка ведомости и формы 5-15А
# =============================================================================
with tab2:
    st.header("Автоматическая сверка ведомости и формы 5-15А")
    st.write("Загрузите расчетно-платежную ведомость (Excel) и выписку 5-15А (PDF или сканированное изображение) для проверки расхождений.")

    col_up1, col_up2 = st.columns(2)
    
    with col_up1:
        st.subheader("1. Расчетно-платежная ведомость")
        uploaded_excel = st.file_uploader("Загрузите ведомость (.xlsx, .xls)", type=["xlsx", "xls"], key="excel_file")

    with col_up2:
        st.subheader("2. Выписка по форме 5-15А")
        uploaded_pdf = st.file_uploader("Загрузите выписку 5-15А (.pdf, .png, .jpg, .jpeg)", type=["pdf", "png", "jpg", "jpeg"], key="pdf_file")

    if st.button("🚀 Начать сверку данных", use_container_width=True):
        if not uploaded_excel or not uploaded_pdf:
            st.error("Пожалуйста, загрузите оба файла (Excel-ведомость и PDF/скан формы 5-15А) перед запуском сверки.")
        else:
            with st.spinner("Идет обработка файлов и извлечение данных..."):
                try:
                    excel_bytes = uploaded_excel.getvalue()
                    df_payroll = parse_excel_accruals(excel_bytes)
                    
                    pdf_bytes = uploaded_pdf.getvalue()
                    pdf_text = extract_text_from_pdf(pdf_bytes)

                    st.success("Файлы успешно обработаны!")

                    st.subheader("📋 Извлеченные данные из Excel")
                    if not df_payroll.empty:
                        st.dataframe(df_payroll.head(20), use_container_width=True)
                    else:
                        st.warning("Не удалось автоматически распознать строки в Excel. Проверьте структуру файла.")

                    st.subheader("📄 Распознанный текст из формы 5-15А (OCR)")
                    if pdf_text and len(pdf_text.strip()) > 0:
                        with st.expander("Показать извлеченный текст выписки 5-15А"):
                            st.text_area("Текст 5-15А", value=pdf_text, height=300)
                    else:
                        st.error("Не удалось извлечь текст из PDF/скана 5-15А.")

                except Exception as e:
                    st.error(f"Произошла ошибка при обработке файлов: {str(e)}")