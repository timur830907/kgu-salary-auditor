import io
import pandas as pd
import streamlit as st
from app.services.reconciliation import reconcile_salary

st.set_page_config(
    page_title="Аудит оплаты труда гражданских служащих (ППРК №1193)",
    page_icon="📊",
    layout="wide",
)

MRPK_2026 = 3932.0  # МРП на 2026 год
MZP_2026 = 85000.0  # МЗП на 2026 год

with st.sidebar:
    st.title("⚙️ Настройки аудита")
    st.info("Система авто-анализа начислений, удержаний и выписок 5-15А государственного учреждения.")
    st.caption("КГУ Salary Auditor v2.5 | Render Edition")

st.title("📊 Аудит оплаты труда гражданских служащих (ППРК №1193)")

tab_coeffs, tab_reconciliation, tab_export = st.tabs([
    "📐 Расчет окладов, надбавок, удержаний и налогов (ППРК №1193)",
    "📊 Ведомости и 5-15А (Сверка и Риски)",
    "📥 Экспорт отчетов"
])

# ==========================================
# Вкладка 1: Полный калькулятор ППРК №1193
# ==========================================
with tab_coeffs:
    st.subheader("📐 Расчет должностных окладов, надбавок, удержаний и налогов")
    
    col_left, col_right = st.columns(2)
    
    with col_left:
        st.markdown("### 1. Должностной оклад и специфика сферы")
        bdo = st.number_input("Базовый должностной оклад (БДО), ₸", value=17697.0, step=100.0)
        sphere = st.selectbox("Сфера деятельности", ["Образование / Педагоги", "Здравоохранение / Врачи", "Социальное обеспечение", "Общие служащие / Адм."])
        category = st.selectbox("Категория должности", ["A1-1", "A1-2", "B1-1", "B2-1", "C-1", "G-1"])
        coeff = st.number_input("Расчетный коэффициент (по стажу)", value=3.42, step=0.01)
        
        do_base = round(bdo * coeff, 2)
        st.info(f"**Базовый ДО (БДО × Коэфф):** {do_base:,.2f} ₸")

        st.markdown("### 2. Доплаты и надбавки")
        out_check = st.checkbox("Особые условия труда (ОУТ 10% от ДО)", value=True)
        vredность_check = st.checkbox("Вредность / Психоневрологические КГУ (30% от БДО)", value=False)
        
        add_teachers = 0.0
        if "Образование" in sphere:
            class_run = st.checkbox("Классное руководство (25% от БДО)", value=False)
            check_notebooks = st.checkbox("Проверка тетрадей (20% от БДО)", value=False)
            if class_run: add_teachers += bdo * 0.25
            if check_notebooks: add_teachers += bdo * 0.20

        out_val = (do_base * 0.10) if out_check else 0.0
        vred_val = (bdo * 0.30) if vredность_check else 0.0
        
        total_accrued = round(do_base + out_val + vred_val + add_teachers, 2)
        st.success(f"**Всего начислено (Грязными):** {total_accrued:,.2f} ₸")

    with col_right:
        st.markdown("### 3. Удержания и налоги с работника")
        is_pensioner = st.checkbox("Пенсионер (Освобождение от ОПВ, ВОСМС)", value=False)
        is_invalid = st.checkbox("Инвалидность 1, 2 группы (Вычет 882 МРП)", value=False)

        # ОПВ (10%)
        opv = 0.0 if is_pensioner else round(total_accrued * 0.10, 2)
        
        # ВОСМС (2%)
        vosmc = 0.0 if is_pensioner else round(total_accrued * 0.02, 2)
        
        # ИПН (10%)
        mrp_deduction = (882 * MRPK_2026 / 12) if is_invalid else (14 * MRPK_2026)
        ipn_base = total_accrued - opv - vosmc - mrp_deduction
        ipn = round(max(0.0, ipn_base * 0.10), 2)
        
        total_deductions = round(opv + vosmc + ipn, 2)
        to_hand = round(total_accrued - total_deductions, 2)

        st.metric("К выдаче на руки (На руки)", f"{to_hand:,.2f} ₸")
        st.write(f"* **ОПВ (10%):** {opv:,.2f} ₸")
        st.write(f"* **ВОСМС (2%):** {vosmc:,.2f} ₸")
        st.write(f"* **ИПН (10%):** {ipn:,.2f} ₸")

        st.markdown("### 4. Налоги и отчисления работодателя")
        opvr = 0.0 if is_pensioner else round(total_accrued * 0.035, 2) # ОПВР 3.5%
        so = round((total_accrued - opv) * 0.035, 2) # Соцотчисления 3.5%
        osmc = round(total_accrued * 0.03, 2) # ОСМС 3%
        
        st.write(f"* **ОПВР (3.5%):** {opvr:,.2f} ₸")
        st.write(f"* **Социальные отчисления (3.5%):** {so:,.2f} ₸")
        st.write(f"* **ОСМС работодателя (3%):** {osmc:,.2f} ₸")

# ==========================================
# Вкладка 2: Сверка (ОСВ + 5-15А)
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