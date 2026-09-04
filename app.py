import io
import pandas as pd
import streamlit as st

from app.services.reconciliation import reconcile_salary

st.set_page_config(
    page_title="Аудитор заработной платы КГУ",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ------------------------------------------------------------------------------
# Боковая панель (Sidebar)
# ------------------------------------------------------------------------------
with st.sidebar:
    st.title("⚙️ Настройки аудита")
    st.info("Система авто-анализа начислений и выписок 5-15А государственного учреждения.")
    st.markdown("---")
    st.markdown("### 📋 Инструкция:")
    st.markdown(
        """
        1. Перейдите во вкладку **Сквозная ОСВ**.
        2. Загрузите файлы Excel за нужные месяцы.
        3. Загрузите файлы PDF (5-15А).
        4. Нажмите **Построить сквозную пофамильную ОСВ**.
        """
    )
    st.markdown("---")
    st.caption("КГУ Salary Auditor v2.5 | Render Edition")

st.title("📊 Аудитор заработной платы КГУ")

# ------------------------------------------------------------------------------
# Закреплённые вкладки верхнего уровня (не исчезают)
# ------------------------------------------------------------------------------
tab_main, tab_export, tab_about = st.tabs(
    ["🔄 Сквозная ОСВ и Сверка", "📥 Экспорт отчетов", "ℹ️ О системе"]
)

# ==============================================================================
# ВКЛАДКА 1: Сквозная ОСВ
# ==============================================================================
with tab_main:
    col1, col2 = st.columns(2)

    with col1:
        accruals_files = st.file_uploader(
            "1. Загрузите ведомости (.xlsx, .xls)",
            type=["xlsx", "xls"],
            accept_multiple_files=True,
            key="accruals_uploader",
        )

    with col2:
        pdf_files = st.file_uploader(
            "2. Загрузите выписки 5-15А (.pdf)",
            type=["pdf"],
            accept_multiple_files=True,
            key="pdf_uploader",
        )

    st.divider()

    if st.button("📊 Построить сквозную пофамильную ОСВ", type="primary", use_container_width=True):
        if not accruals_files:
            st.warning("Пожалуйста, загрузите хотя бы один файл ведомости Excel.")
        else:
            with st.spinner("Идёт обработка и формирование периода..."):
                try:
                    df_result, risk_comments = reconcile_salary(accruals_files, pdf_files)

                    if not df_result.empty:
                        st.session_state["df_result"] = df_result
                        st.session_state["risk_comments"] = risk_comments

                    else:
                        st.error("Не удалось выделить сотрудников из загруженных файлов.")

                except Exception as e:
                    st.error(f"Ошибка обработки: {str(e)}")

    # Отображение результатов, если они есть в памяти
    if "df_result" in st.session_state and not st.session_state["df_result"].empty:
        df_result = st.session_state["df_result"]
        risk_comments = st.session_state.get("risk_comments", [])

        total_people = df_result["fio"].nunique()
        active_period = df_result["month"].unique().tolist()

        st.success(
            f"Обработано сотрудников: {total_people}. Период в отчёте: {', '.join(active_period)}."
        )

        m_col1, m_col2, m_col3 = st.columns(3)
        with m_col1:
            st.metric("Всего сотрудников", total_people)
        with m_col2:
            total_kvyd = df_result["kvyd"].sum()
            st.metric("Всего к выдаче", f"{total_kvyd:,.2f} ₸")
        with m_col3:
            mismatches = df_result[df_result["status"] != "Закрыто"]
            st.metric("Выявлено расхождений", len(mismatches))

        st.subheader("📋 Детализированная оборотно-сальдовая ведомость")

        display_df = df_result.rename(
            columns={
                "fio": "ФИО сотрудника",
                "month": "Месяц",
                "start_bal": "Остаток на начало месяца, ₸",
                "kvyd": "К выдаче (Ведомость), ₸",
                "paid": "Перечислено (5-15А), ₸",
                "end_bal": "Остаток на конец месяца, ₸",
                "status": "Статус",
            }
        )

        st.dataframe(display_df, use_container_width=True, height=500)

        if risk_comments:
            with st.expander("📌 Примечания аудита и риски", expanded=True):
                for comment in risk_comments:
                    st.write(f"- {comment}")

# ==============================================================================
# ВКЛАДКА 2: Экспорт отчетов
# ==============================================================================
with tab_export:
    st.subheader("📥 Выгрузка результатов анализа")

    if "df_result" in st.session_state and not st.session_state["df_result"].empty:
        df_to_export = st.session_state["df_result"]

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_to_export.to_excel(writer, sheet_name="Сводная ОСВ", index=False)
        output.seek(0)

        st.download_button(
            label="💾 Скачать сводную ОСВ в Excel (.xlsx)",
            data=output,
            file_name="Сводная_ОСВ_Заработная_плата.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

        st.markdown("---")
        st.dataframe(df_to_export.head(10), use_container_width=True)
    else:
        st.info("Сначала выполните анализ на первой вкладке.")

# ==============================================================================
# ВКЛАДКА 3: О системе
# ==============================================================================
with tab_about:
    st.subheader("ℹ️ О программе")
    st.markdown(
        """
        **Аудитор заработной платы КГУ v2.5**
        
        * Поддерживает динамическое формирование периодов (показывает только загруженные месяцы).
        * Корректно сохраняет интерфейс со вкладками.
        * Автоматически сводит данные ведомостей и PDF 5-15А.
        """
    )