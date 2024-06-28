import pandas as pd
import numpy as np
import random
from datetime import datetime, timedelta

# Constantes
cost_center = "CL10201006"
co_object_name = "ADMINISTRATION"
cost_elements = list(range(61002905, 61002905 + 20))
cost_element_names = [
    "AUTO INSURANCE", "COMPANY CAR EXPENSE", "FIX RENT OP LEASE CL", "INCENTIVE CASH AWARD",
    "INTERNAL MATL GENERA", "JANITORIAL SERVICES", "LIABILITY INSURANCE", "LODGING - EMPLOYEE",
    "MEALS -EMPLOYEE", "MISC SERVCONTRACT", "OFFICE SUPPLY & EXP", "OP SUPL STORES",
    "OTHER EMPL BENEFITS", "OTHER EMPL SUPPORT", "PENSION SERVICE COST", "PERSONNEL TRANSPORTA",
    "PROPERTY INSURANCE", "SMALL TOOLS ISSUE", "TECH SERVICE OUTSIDE", "TRANSPORT – EMPLOYEE",
    "WASTE DISPOSAL"
]
posted_unit_of_meas = "AU"
purchase_order_text = "Lorem Ipsum"
period = 6
fiscal_year = 2024
document_header_text = "Comentario que se llenara"
offsetting_account_type = "S"
transaction_currencies = ["CLP", "USD"]

user_names = [
    "ALEJAAA", "ARAYACA4", "ARAYARA", "CONCHKC", "CORTECC", "DANNAAE", "ESPINME", "FARKAMF",
    "GUTIECG", "GUTIECG1", "HERNAJH2", "KAPPEVK", "KAUDETK", "KISSVK2", "OLIVAPO", "PESARCP",
    "ROJASJE", "RTP_BATCH", "SANCHMS3", "SANXHKS", "SAP_WFRT", "SEPULAS1", "SZEKEMS"
]

names_of_offsetting_accounts = [
    "$SERPAN LTDA.", "ACCR INCENTIVE CASH", "ACCR LABOR UNION EXP", "ACCER OTHER",
    "ARRENDADORA DE VEHICULOS SA", "AUTORENTAS DEL PACIFICO LTDA", "BANCO SECURITY",
    "BIOMETRYCLOUD SPA", "COMPASS CATERING SA", "COPEC VOLTEX SPA", "ELIANA BERNARDITA DAHMEN MARI",
    "FADAF SERV DE ING SPA", "FULLOFFICE LTDA", "GRIR PRODUCT RELATED", "IGUANA PROACTIVA LTDA",
    "IMPO Y REPR BOX SOLUTION", "INCENTIVE CASH AWARD", "INMOBILIARIA EL ANCLA LTDA",
    "JPM USD OPER WIRE IN", "LIBERTY CIA DE SEG GRAL", "LOG HUALPEN LTDA", "OP SUPPLY INV MODULE",
    "PBO LT PENSION PLAN", "PREPAID INSURANCE", "SERCOL CERTIF LTDA", "SERV GASTRONOMICOS CASINOS",
    "SOC DE EDUCADORAS DE PARVULOS", "SPARE PARTS INV MOD", "TRANSBUILD SPA",
    "TRANSP E INV TRUJILLO SPA", "TRANSPORTES MABE NELIDA MABEL"
]

# Generación de datos
def generate_dates(start_date, end_date, num_dates):
    date_list = []
    current_date = start_date
    while current_date <= end_date:
        date_list.append(current_date)
        current_date += timedelta(days=1)
    return random.choices(date_list, k=num_dates)

def generate_rows():
    rows = []
    for name in names_of_offsetting_accounts:
        dates = generate_dates(datetime(2024, 1, 1), datetime.now(), len(names_of_offsetting_accounts) * 5)
        for date in set(dates):
            sales_per_day = random.randint(2, 5)
            currencies = ["CLP", "USD"]
            for _ in range(sales_per_day):
                if not currencies:
                    currency = random.choice(transaction_currencies)
                else:
                    currency = currencies.pop()
                if currency == "USD":
                    value_tran_curr = round(random.uniform(1, 1000), 2)
                else:
                    value_tran_curr = random.randint(100000, 10000000)
                rows.append([
                    cost_center,
                    co_object_name,
                    random.choice(cost_elements),
                    random.choice(cost_element_names),
                    round(random.uniform(1, 200000), 2),
                    1,
                    posted_unit_of_meas,
                    purchase_order_text,
                    random.randint(8000319288, 8000500000),
                    period,
                    random.randint(800, 42984298),
                    date.strftime('%d-%m-%Y'),
                    fiscal_year,
                    random.choice(user_names),
                    value_tran_curr,
                    currency,
                    document_header_text,
                    offsetting_account_type,
                    random.randint(10000000, 99999999),
                    name
                ])
    return rows

columns = [
    "Cost Center", "CO Objetct Name", "Cost Element", "Cost element name", "Val.in rep.cur.",
    "Total quantity", "Posted unit of meas.", "Purchase Order Text", "Purchasing Document", "Period",
    "Ref document number", "Posting Date", "Fiscal Year", "User Name", "Value tranCurr", 
    "Transaction Currency", "Document Header Text", "Offsetting Account Type", "offsetting account",
    "Name of offsetting account"
]

# Crear DataFrame y escribir a Excel
data = generate_rows()
df = pd.DataFrame(data, columns=columns)
df.to_excel("testPython.xlsx", index=False, sheet_name="Sheet1")


