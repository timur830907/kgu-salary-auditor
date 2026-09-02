import pandas as pd
from reconciliation import reconcile_salary

# Загрузка подготовленных данных
df_accruals = pd.read_excel("ведомость_к_выдаче_январь.xlsx")
df_payments = pd.read_excel("выплаты_5_15А_январь.xlsx")

# Выполнение сверки
df_report = reconcile_salary(df_accruals, df_payments)

# Сохранение итогового отчета
df_report.to_excel("Результат_Сверки.xlsx", index=False)