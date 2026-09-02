import io
import sys
from pathlib import Path

# Добавление путей для правильного импорта модулей
ROOT_DIR = Path(__file__).resolve().parent
sys.path.append(str(ROOT_DIR))
sys.path.append(str(ROOT_DIR / "app"))

import pandas as pd
import streamlit as st

# Безопасный импорт функций парсинга и сверки
try:
    from app.services.reconciliation import (
        parse_excel_accruals,
        parse_image_5_15a,
        parse_pdf_5_15a,
        reconcile_salary,
    )
except ImportError:
    from services.reconciliation import (
        parse_excel_accruals,
        parse_image_5_15a,
        parse_pdf_5_15a,
        reconcile_salary,
    )

st.set_page_config(
    page_title="Аудит заработной платы и Калькулятор ПП РК № 1193",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("🏛️ Калькулятор заработной платы и Аудит формы 5-15А")
st.caption("Система расчета гражданских служащих (ПП РК № 1193) и финансовая сверка")

tab1, tab2 = st.tabs(["📊 Калькулятор начислений (ПП РК № 1193)", "🔍 Сверка ведомостей и 5-15А"])

# ==========================================
# ВКЛАДКА 1: КАЛЬКУЛЯТОР
# ==========================================
with tab1:
    st.header("Расчет должностного оклада и начислений")
    
    col_params, col_results = st.columns([1, 1])
    
    with col_params:
        st.subheader("Параметры работника")
        bdo = st.number_input("Базовый должностной оклад (БДО), ₸", value=17697, step=100, key="calc_bdo")
        
        cat_group = st.selectbox(
            "Функциональная блок-группа",
            ["A (Управленческий персонал)", "B (Основной персонал)", "C (Административный)", "D (Вспомогательный)"],
            key="calc_cat_group"
        )
        
        category = st.selectbox(
            "Категория должности",
            ["A1", "A2", "A3", "B1", "B2", "B3", "B4", "C1", "C2", "C3", "D"],
            key="calc_category"
        )
        
        exp_years = st.number_input("Стаж работы (лет)", min_value=0, max_value=50, value=5, key="calc_exp")
        
        st.subheader("Надбавки и доплаты")
        has_eco = st.checkbox("Экологическая надбавка (Арал/Семипалатинск)", value=False, key="calc_eco_check")
        eco_rate = st.number_input("Процент экологической надбавки (%)", value=30 if has_eco else 0, step=5, key="calc_eco_rate")
        
        other_allowances = st.number_input("Прочие надбавки (₸)", value=0, step=1000, key="calc_other_allow")

    def get_coefficient(cat, exp):
        base_coeff = {
            "A1": 5.0, "A2": 4.5, "A3": 4.0,
            "B1": 3.8, "B2": 3.4, "B3": 3.1, "B4": 2.8,
            "C1": 2.5, "C2": 2.3, "C3": 2.1, "D": 1.8
        }.get(cat, 2.0)
        exp_bonus = min(exp * 0.05, 1.0)
        return round(base_coeff + exp_bonus, 2)

    coeff = get_coefficient(category, exp_years)
    base_salary = bdo * coeff
    eco_amount = base_salary * (eco_rate / 100.0)
    total_gross = base_salary + eco_amount + other_allowances

    opv = total_gross * 0.10
    vosms = total_gross * 0.02
    mzp = 85000
    ipn_base = max(0, total_gross - opv - mzp - vosms)
    ipn = ipn_base * 0.10
    
    total_deductions = opv + vosms + ipn
    net_salary = total_gross - total_deductions

    opvr = total_gross * 0.035
    social_tax = max(0, (total_gross - opv - vosms) * 0.095)
    osms = total_gross * 0.03
    so = (total_gross - opv) * 0.035

    with col_results:
        st.subheader("Расчетные показатели")
        st.info(f"**Коэффициент:** {coeff}")
        
        st.markdown(f"""
        | Наименование | Сумма (₸) |
        | :--- | :--- |
        | **Должностной оклад (ДО)** | **{base_salary:,.2f}** |
        | Экологическая надбавка | {eco_amount:,.2f} |
        | Прочие надбавки | {other_allowances:,.2f} |
        | **Всего начислено (Gross)** | **{total_gross:,.2f}** |
        """)
        
        st.markdown("### Удержания с работника:")
        st.markdown(f"""
        - ОПВ (10%): `{opv:,.2f} ₸`
        - ВОСМС (2%): `{vosms:,.2f} ₸`
        - ИПН (10%): `{ipn:,.2f} ₸`
        - **На руки (Net): `{net_salary:,.2f} ₸`**
        """)
        
        with st.expander("Отчисления работодателя (Информация)"):
            st.markdown(f"""
            - **ОПВР (3.5%):** `{opvr:,.2f} ₸`
            - Социальные отчисления (3.5%): `{so:,.2f} ₸`
            - ОСМС (3%): `{osms:,.2f} ₸`
            - Социальный налог: `{social_tax:,.2f} ₸`
            """)

# ==========================================
# ВКЛАДКА 2: СВЕРКА ВЕДОМОСТЕЙ И 5-15А
# ==========================================
with tab2:
    st.header("Автоматическая сверка ведомости и формы 5-15А")
    st.write("Загрузите расчетно-платежную ведомость (Excel) и выписку 5-15А (PDF или сканированное изображение) для проверки расхождений.")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("### 1. Расчетно-платежная ведомость")
        excel_file = st.file_uploader(
            "Загрузите ведомость (.xlsx, .xls)", type=["xlsx", "xls"], key="excel_uploader"
        )

    with col2:
        st.markdown("### 2. Выписка по форме 5-15А")
        payment_file = st.file_uploader(
            "Загрузите выписку 5-15А (.pdf, .png, .jpg, .jpeg)",
            type=["pdf", "png", "jpg", "jpeg"],
            key="payment_uploader",
        )

    if excel_file and payment_file:
        if st.button("🚀 Начать сверку данных", type="primary", use_container_width=True, key="start_recon_btn"):
            with st.spinner("Извлечение данных и выполнение сверки..."):
                df_accruals = parse_excel_accruals(excel_file)

                file_ext = payment_file.name.split(".")[-1].lower()
                if file_ext == "pdf":
                    df_payments = parse_pdf_5_15a(payment_file)
                else:
                    df_payments = parse_image_5_15a(payment_file)

                if df_accruals.empty:
                    st.error("Не удалось извлечь записи из расчетной ведомости Excel.")
                elif df_payments.empty:
                    st.error("Не удалось распознать данные из файла 5-15А.")
                else:
                    res_df, risks = reconcile_salary(df_accruals, df_payments)

                    st.success("Сверка успешно выполнена!")

                    st.subheader("📋 Сводный отчет сверки")
                    st.dataframe(res_df, use_container_width=True)

                    if risks:
                        st.subheader("⚠️ Выявленные расхождения и риски")
                        for risk in risks:
                            st.warning(risk)
                    else:
                        st.balloons()
                        st.success("🎉 Полное совпадение! Расхождений не обнаружено.")