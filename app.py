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

# Фиксированные вкладки верхнего уровня
tab_coeffs, tab_reconciliation, tab_export = st.tabs([
    "📐 Заработная плата с коэффициентами (ППРК №1193)",
    "📊 Ведомости и 5-15А (Сверка и Риски)",
    "📥 Экспорт отчетов"
])

# ==========================================
# Вкладка 1: Полный калькулятор ППРК №1193
# ==========================================
with tab_coeffs:
    st.subheader("📐 Калькулятор окладов, надбавок и доплат (ППРК №1193)")
    st.caption("Расчет базовых должностных окладов (БДО), стажевых коэффициентов и блоков доплат A, B, C, D.")

    c1, c2, c3 = st.columns(3)
    with c1:
        bdo = st.number_input("Базовый должностной оклад (БДО), ₸", value=17697.0, step=100.0)
        block = st.selectbox("Блок / Группа управленческого персонала", [
            "Блок А (Управленческий персонал)",
            "Блок В (Основной персонал)",
            "Блок С (Административный персонал)",
            "Блок D (Вспомогательный персонал)"
        ])
    with c2:
        category = st.selectbox("Категория должности", [
            "A1-1", "A1-2", "B1-1", "B2-1", "B3-1", "B4-1", "C-1", "C-2", "D-1", "D-2", "G-1", "G-2"
        ])
        experience_years = st.number_input("Стаж работы (лет)", min_value=0, max_value=50, value=5)
    with c3:
        step = st.selectbox("Ступень стажа", [
            "До 1 года (Ступень 1)",
            "От 1 до 3 лет (Ступень 2)",
            "От 3 до 5 лет (Ступень 3)",
            "От 5 до 10 лет (Ступень 4)",
            "От 10 до 15 лет (Ступень 5)",
            "Свыше 15 лет (Ступень 6)"
        ])
        coeff = st.number_input("Расчетный коэффициент (по сетке ППРК №1193)", value=3.42, step=0.01)

    st.markdown("---")
    st.subheader("🧩 Доплаты и надбавки (Блоки А, Б, С, Д)")

    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("**Специфические доплаты:**")
        check_psy = st.checkbox("За особые условия труда (10% от БДО/ДО)")
        check_hazard = st.checkbox("За вредность / тяжелые условия (30-50% от БДО)")
        check_rank = st.checkbox("Доплата за классность / квалификацию")
        other_allowance = st.number_input("Прочие надбавки (в сумме), ₸", value=0.0)

    with col_b:
        st.markdown("**Расчет Итогового Оклада:**")
        base_salary = round(bdo * coeff, 2)
        allowance_total = 0.0
        if check_psy:
            allowance_total += base_salary * 0.10
        if check_hazard:
            allowance_total += bdo * 0.30
        allowance_total += other_allowance

        total_calculated = round(base_salary + allowance_total, 2)

        st.metric("Должностной оклад (ДО = БДО × Коэф)", f"{base_salary:,.2f} ₸")
        st.metric("Сумма доплат и надбавок", f"{allowance_total:,.2f} ₸")
        st.metric("ИТОГО Начислено (до удержаний)", f"{total_calculated:,.2f} ₸")

# ==========================================
# Вкладка 2: Ведомости и 5-15А (Сверка и Риски)
# ==========================================
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
            with st.spinner("Анализ данных и сканирование 5-15А..."):
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

# ==========================================
# Вкладка 3: Экспорт отчетов
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
    else:
        st.info("Сначала запустите анализ на второй вкладке.")