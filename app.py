import io
import pandas as pd
import streamlit as st
from app.services.reconciliation import reconcile_salary

st.set_page_config(
    page_title="Аудитор заработной платы КГУ",
    page_icon="📊",
    layout="wide",
)

# Боковое меню
with st.sidebar:
    st.title("⚙️ Настройки аудита")
    st.info("Система авто-анализа начислений и выписок 5-15А государственного учреждения.")
    st.caption("КГУ Salary Auditor v2.5 | Render Edition")

st.title("📊 Аудитор заработной платы КГУ")

# Фиксированные вкладки верхнего уровня
tab_main, tab_export, tab_about = st.tabs(
    ["🔄 Сквозная ОСВ и Сверка", "📥 Экспорт отчетов", "ℹ️ О системе"]
)

# Вкладка 1: Главный модуль
with tab_main:
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
            with st.spinner("Анализ данных и сканирование 5-15А..."):
                df_result, risk_comments = reconcile_salary(accruals_files, pdf_files)
                st.session_state["df_result"] = df_result
                st.session_state["risk_comments"] = risk_comments

    # Отображение результатов при наличии данных
    if "df_result" in st.session_state and not st.session_state["df_result"].empty:
        df_result = st.session_state["df_result"]
        risk_comments = st.session_state.get("risk_comments", [])

        if risk_comments:
            for log in risk_comments:
                st.success(log)

        total_people = df_result["fio"].nunique()
        total_kvyd = df_result["kvyd"].sum()
        mismatches = len(df_result[df_result["status"] != "Закрыто"])

        m_col1, m_col2, m_col3 = st.columns(3)
        m_col1.metric("Всего сотрудников", total_people)
        m_col2.metric("Всего к выдаче", f"{total_kvyd:,.2f} ₸")
        m_col3.metric("Выявлено расхождений", mismatches)

        st.subheader("📋 Детализированная оборотно-сальдовая ведомость")
        
        # Переименование колонок для отображения
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

# Вкладка 2: Экспорт
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
        st.info("Сначала запустите анализ на первой вкладке, чтобы выгрузить отчёт.")

# Вкладка 3: О системе
with tab_about:
    st.subheader("ℹ️ О системе")
    st.write(
        """
        **Аудитор заработной платы КГУ** — инструмент для автоматизации проверки ведомостей и выписок 5-15А.
        
        **Основные функции:**
        - Сквозное сведение оборотно-сальдовой ведомости (ОСВ) по сотрудникам.
        - Распознавание сканированных PDF-выписок 5-15А с использованием OCR (Tesseract).
        - Автоматический расчет остатков и подсчет расхождений.
        """
    )