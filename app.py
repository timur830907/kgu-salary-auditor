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

st.title("📊 Аудит оплаты труда гражданских служащих, работников организаций, содержащихся за счет средств государственного бюджета, работников казенных предприятий (ППРК №1193)")

tab_coeffs, tab_reconciliation, tab_export = st.tabs([
    "📐 Расчет окладов, удержаний и налогов (ППРК №1193)",
    "📊 Ведомости и 5-15А (Сверка и Риски)",
    "📥 Экспорт отчетов"
])

# ==========================================
# Вкладка 1: Полный расчет начислений и налогов
# ==========================================
with tab_coeffs:
    st.subheader("📐 Расчет должностных окладов, надбавок, удержаний и налогов")
    st.caption("Соответствует Постановлению Правительства РК №1193 и Налоговому кодексу РК.")

    c1, c2, c3 = st.columns(3)
    with c1:
        bdo = st.number_input("Базовый должностной оклад (БДО), ₸", value=17697.0, step=100.0)
        block = st.selectbox("Блок / Группа", [
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
            "До 1 года (Ступень 1)", "От 1 до 3 лет (Ступень 2)",
            "От 3 до 5 лет (Ступень 3)", "От 5 до 10 лет (Ступень 4)",
            "От 10 до 15 лет (Ступень 5)", "Свыше 15 лет (Ступень 6)"
        ])
        coeff = st.number_input("Расчетный коэффициент (ППРК №1193)", value=3.42, step=0.01)

    st.markdown("---")
    
    col_a, col_b, col_c = st.columns(3)
    with col_a:
        st.markdown("**1. Доплаты и надбавки (Блоки A, B, C, D):**")
        check_psy = st.checkbox("За особые условия труда (10% от ДО)")
        check_hazard = st.checkbox("За вредность / тяжелые условия (30% от БДО)")
        other_allowance = st.number_input("Прочие надбавки, ₸", value=0.0)

        base_salary = round(bdo * coeff, 2)
        allowance_total = 0.0
        if check_psy:
            allowance_total += base_salary * 0.10
        if check_hazard:
            allowance_total += bdo * 0.30
        allowance_total += other_allowance

        gross_salary = round(base_salary + allowance_total, 2)
        st.metric("Всего Начислено (Gross)", f"{gross_salary:,.2f} ₸")

    with col_b:
        st.markdown("**2. Удержания с работника:**")
        opv = round(gross_salary * 0.10, 2)
        vosms = round(gross_salary * 0.02, 2)
        ipn_base = gross_salary - opv - vosms - (14 * 4250)  # С учетом 14 МРП вычета
        ipn = round(max(0.0, ipn_base * 0.10), 2)
        
        total_deductions = round(opv + vosms + ipn, 2)
        net_salary = round(gross_salary - total_deductions, 2)

        st.write(f"• ОПВ (10%): **{opv:,.2f} ₸**")
        st.write(f"• ВОСМС (2%): **{vosms:,.2f} ₸**")
        st.write(f"• ИПН (10%): **{ipn:,.2f} ₸**")
        st.metric("К выдаче на руки (Net)", f"{net_salary:,.2f} ₸")

    with col_c:
        st.markdown("**3. Налоги и отчисления работодателя:**")
        opvr = round(gross_salary * 0.035, 2)  # ОПВР 3.5%
        so = round((gross_salary - opv) * 0.035, 2)  # СО 3.5%
        osms = round(gross_salary * 0.03, 2)  # ООСМС 3%
        total_employer_tax = round(opvr + so + osms, 2)

        st.write(f"• ОПВР (3.5%): **{opvr:,.2f} ₸**")
        st.write(f"• Соц. отчисления (3.5%): **{so:,.2f} ₸**")
        st.write(f"• ОСМС (3%): **{osms:,.2f} ₸**")
        st.metric("Нагрузка на работодателя", f"{total_employer_tax:,.2f} ₸")

# ==========================================
# Вкладка 2: Ведомости и 5-15А (Сверка и Риски)
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