import io
import pandas as pd
import streamlit as st
from app.services.reconciliation import reconcile_salary

st.set_page_config(
    page_title="Аудитор заработной платы КГУ",
    page_icon="📊",
    layout="wide",
)

with st.sidebar:
    st.title("⚙️ Настройки аудита")
    st.info("Система авто-анализа начислений и выписок 5-15А государственного учреждения.")
    st.caption("КГУ Salary Auditor v2.5 | Render Edition")

st.title("📊 Аудитор заработной платы КГУ")

# Зафиксированные 3 вкладки
tab_coeffs, tab_reconciliation, tab_export = st.tabs([
    "📐 Заработная плата с коэффициентами",
    "📊 Ведомости и 5-15А (Сверка и Риски)",
    "📥 Экспорт отчетов"
])

# Вкладка 1: Заработная плата с коэффициентами
with tab_coeffs:
    st.subheader("📐 Модуль расчета и проверки окладов с коэффициентами")
    st.info("Здесь выполняется проверка базовых окладов, стажевых коэффициентов и надбавок сотрудников.")
    
    col_c1, col_c2 = st.columns(2)
    with col_c1:
        st.number_input("Базовый должностной оклад (БДО), ₸", value=17697.0)
        st.selectbox("Категория должности", ["G-1", "G-2", "G-3", "G-4", "G-5", "G-6", "G-7", "G-8"])
    with col_c2:
        st.number_input("Коэффициент за стаж / категорию", value=1.00, step=0.01)
        st.number_input("Процент доплат и надбавок, %", value=0.0, step=1.0)

# Вкладка 2: Ведомости и 5-15А (Сверка и Риски переплат)
with tab_reconciliation:
    st.subheader("📊 Ведомости и 5-15А: Анализ расхождений и рисков переплат")
    
    col1, col2 = st.columns(2)
    with col1:
        accruals_files = st.file_uploader(
            "1. Загрузите файлы Excel (ведомости за нужные месяцы)",
            type=["xlsx", "xls"],
            accept_multiple_files=True,
            key="accruals",
        )
    with col2:
        pdf_files = st.file_uploader(
            "2. Загрузите файлы PDF (выписки 5-15А)",
            type=["pdf"],
            accept_multiple_files=True,
            key="pdfs",
        )

    if st.button("📊 Построить сквозную пофамильную ОСВ", type="primary", use_container_width=True):
        if not accruals_files:
            st.warning("Загрузите файлы Excel ведомостей.")
        else:
            with st.spinner("Анализ данных и проверка выписок 5-15А..."):
                df_result, risk_comments = reconcile_salary(accruals_files, pdf_files)
                st.session_state["df_result"] = df_result
                st.session_state["risk_comments"] = risk_comments

    if "df_result" in st.session_state and not st.session_state["df_result"].empty:
        df_result = st.session_state["df_result"]
        risk_comments = st.session_state.get("risk_comments", [])

        if risk_comments:
            for log in risk_comments:
                st.success(log)

        total_people = df_result["fio"].nunique()
        total_kvyd = df_result["kvyd"].sum()
        mismatches = len(df_result[df_result["status"] != "Закрыто"])
        overpayment_risks = len(df_result[df_result["status"] == "Переплата (Риск)"])

        m_col1, m_col2, m_col3, m_col4 = st.columns(4)
        m_col1.metric("Всего сотрудников", total_people)
        m_col2.metric("Всего к выдаче", f"{total_kvyd:,.2f} ₸")
        m_col3.metric("Выявлено расхождений", mismatches)
        m_col4.metric("Риски переплат", overpayment_risks)

        st.subheader("📋 Детализированная оборотно-сальдовая ведомость")
        
        df_display = df_result.rename(columns={
            "fio": "ФИО сотрудника",
            "month": "Месяц",
            "start_bal": "Остаток на начало месяца, ₸",
            "kvyd": "К выдаче (Ведомость), ₸",
            "paid": "Перечислено (5-15А), ₸",
            "end_bal": "Остаток на конец месяца, ₸",
            "status": "Статус"
        })

        st.dataframe(df_display, use_container_width=True, height=500)

# Вкладка 3: Экспорт
with tab_export:
    st.subheader("📥 Выгрузка отчёта")
    if "df_result" in st.session_state and not st.session_state["df_result"].empty:
        df_result = st.session_state["df_result"]
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_result.to_excel(writer, sheet_name="Сводная ОСВ", index=False)
        output.seek(0)

        st.download_button(
            label="💾 Скачать Excel (.xlsx)",
            data=output,
            file_name="Сводная_ОСВ_5_15А.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    else:
        st.info("Сначала запустите анализ на вкладке сверки, чтобы сформировать отчёт.")