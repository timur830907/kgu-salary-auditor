import io
import re
import pandas as pd
import streamlit as st

# Импорт функции сверки из нашего обновленного сервиса
from app.services.reconciliation import reconcile_salary

# ------------------------------------------------------------------------------
# Настройка страницы
# ------------------------------------------------------------------------------
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
        2. Загрузите файлы Excel (ведомости) по всем подразделениям/месяцам.
        3. Загрузите файлы PDF (выписки 5-15А).
        4. Нажмите **Построить сквозную пофамильную ОСВ**.
        """
    )
    st.markdown("---")
    st.caption("КГУ Salary Auditor v2.4 | Render Edition")

# ------------------------------------------------------------------------------
# Главный заголовок
# ------------------------------------------------------------------------------
st.title("📊 Аудитор заработной платы КГУ")
st.markdown(
    """
    Автоматизированная система сбора, анализа и выявления расхождений между расчетно-платежными 
    ведомостями (Excel) и казначейскими выписками формы 5-15А (PDF).
    """
)

# ------------------------------------------------------------------------------
# Вкладки приложения
# ------------------------------------------------------------------------------
tab_main, tab_export, tab_about = st.tabs(
    ["🔄 Сквозная ОСВ и Сверка", "📥 Экспорт отчетов", "ℹ️ О системе"]
)

# ==============================================================================
# ВКЛАДКА 1: Сквозная ОСВ и Сверка
# ==============================================================================
with tab_main:
    col1, col2 = st.columns(2)

    with col1:
        accruals_files = st.file_uploader(
            "1. Загрузите ведомости за все месяцы (.xlsx, .xls)",
            type=["xlsx", "xls"],
            accept_multiple_files=True,
            key="accruals_uploader",
        )

    with col2:
        pdf_files = st.file_uploader(
            "2. Загрузите выписки 5-15А за все месяцы (.pdf)",
            type=["pdf"],
            accept_multiple_files=True,
            key="pdf_uploader",
        )

    st.divider()

    if st.button("📊 Построить сквозную пофамильную ОСВ", type="primary", use_container_width=True):
        if not accruals_files:
            st.warning("Пожалуйста, загрузите хотя бы один файл ведомости Excel.")
        else:
            with st.spinner("Идёт обработка данных, объединение подразделений и сведение сотрудников..."):
                try:
                    # Передаём все загруженные файлы в сервис сверки
                    df_result, risk_comments = reconcile_salary(
                        accruals_files, pdf_files
                    )

                    if not df_result.empty:
                        # Сохраняем результат в session_state для последующего экспорта
                        st.session_state["df_result"] = df_result
                        st.session_state["risk_comments"] = risk_comments

                        # Считаем точное число уникальных сотрудников
                        total_people = df_result["fio"].nunique()

                        st.success(
                            f"Обработано всех сотрудников: {total_people}. Построена сквозная цепочка за 12 месяцев."
                        )

                        # Метрики верхнего уровня
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

                        # Переименовываем столбцы для красивого отображения в Streamlit
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

                        # Вывод сводной таблицы
                        st.dataframe(display_df, use_container_width=True, height=500)

                        # Вывод дополнительных комментариев/аудит-рисков
                        if risk_comments:
                            with st.expander("📌 Примечания аудита и риски", expanded=True):
                                for comment in risk_comments:
                                    st.write(f"- {comment}")
                    else:
                        st.error(
                            "Не удалось выделить сотрудников. Проверьте содержимое загруженных Excel-файлов."
                        )

                except Exception as e:
                    st.error(f"Ошибка обработки: {str(e)}")

# ==============================================================================
# ВКЛАДКА 2: Экспорт отчетов
# ==============================================================================
with tab_export:
    st.subheader("📥 Выгрузка результатов анализа")
    
    if "df_result" in st.session_state and not st.session_state["df_result"].empty:
        df_to_export = st.session_state["df_result"]
        
        # Генерация Excel в памяти
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
        st.markdown("#### Предпросмотр скачиваемых данных:")
        st.dataframe(df_to_export.head(10), use_container_width=True)
    else:
        st.info("Сначала выполните анализ на первой вкладке, чтобы сформировать отчёт для скачивания.")

# ==============================================================================
# ВКЛАДКА 3: О системе
# ==============================================================================
with tab_about:
    st.subheader("ℹ️ О программе")
    st.markdown(
        """
        **Аудитор заработной платы КГУ** разработан для автоматической проверки соответствия 
        бухгалтерского учета первичным казначейским документам.

        **Возможности:**
        * Парсинг многостраничных Excel-ведомостей любого формата 1С.
        * Распознавание сканированных и цифровых PDF-выписок по форме 5-15А с применением Tesseract OCR.
        * Автоматическое сопоставление сотрудников по ФИО вне зависимости от порядка строк.
        * Расчет непрерывного сальдо с переносом остатков из месяца в месяц.
        * Формирование итогового отчета за весь календарный год.
        """
    )