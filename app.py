import io
import pandas as pd
import streamlit as st
from app.services.reconciliation import reconcile_salary

st.set_page_config(
    page_title="Аудит оплаты труда гражданских служащих (ППРК №1193)",
    page_icon="📊",
    layout="wide",
)

with st.sidebar:
    st.title("⚙️ Настройки аудита")
    st.info("Система авто-анализа начислений, удержаний и выписок 5-15А государственного учреждения.")
    st.caption("КГУ Salary Auditor v2.5 | Render Edition")

st.title("📊 Аудит оплаты труда гражданских служащих (ППРК №1193)")

tab_coeffs, tab_reconciliation, tab_export = st.tabs([
    "📐 Расчет окладов, удержаний и налогов (ППРК №1193)",
    "📊 Ведомости и 5-15А (Сверка и Риски)",
    "📥 Экспорт отчетов"
])

# ==========================================
# Вкладка 1: Калькулятор
# ==========================================
with tab_coeffs:
    st.subheader("📐 Расчет должностных окладов, надбавок, удержаний и налогов")
    c1, c2, c3 = st.columns(3)
    with c1:
        bdo = st.number_input("Базовый должностной оклад (БДО), ₸", value=17697.0, step=100.0)
    with c2:
        category = st.selectbox("Категория должности", ["A1-1", "B1-1", "C-1", "G-1"])
    with c3:
        coeff = st.number_input("Расчетный коэффициент", value=3.42, step=0.01)

    base_salary = round(bdo * coeff, 2)
    st.metric("Должностной оклад (ДО)", f"{base_salary:,.2f} ₸")

# ==========================================
# Вкладка 2: Сверка
# ==========================================
with tab_reconciliation:
    st.subheader("📊 Ведомости и 5-15А: Сквозной пофамильный аудит")

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
            with st.spinner("Сквозной расчет ОСВ и сканирование 5-15А..."):
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
        mismatches_count = len(df_result[df_result["status"] != "Закрыто"])
        mismatches_sum = df_result[df_result["status"] != "Закрыто"]["end_bal"].abs().sum()
        overpayments_sum = df_result[df_result["status"] == "Переплата"]["end_bal"].abs().sum()

        m_col1, m_col2, m_col3, m_col4 = st.columns(4)
        m_col1.metric("Всего сотрудников", total_people)
        m_col2.metric("Всего к выдаче", f"{total_kvyd:,.2f} ₸")
        m_col3.metric(f"Расхождений ({mismatches_count} стр.)", f"{mismatches_sum:,.2f} ₸")
        m_col4.metric("Сумма переплат (Риск)", f"{overpayments_sum:,.2f} ₸")

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

# ==========================================
# Вкладка 3: Экспорт
# ==========================================
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