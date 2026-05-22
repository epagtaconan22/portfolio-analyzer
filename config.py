"""Application defaults, KPI formula text, and account mapping rules (first match wins)."""
# Default settings
ECO_OCC_TARGET = 0.95

# Months for period filtering
MONTHS = {1:"Jan",2:"Feb",3:"Mar",4:"Apr",5:"May",6:"Jun",
          7:"Jul",8:"Aug",9:"Sep",10:"Oct",11:"Nov",12:"Dec"}

QUARTERS = {"Q1":[1,2,3],"Q2":[4,5,6],"Q3":[7,8,9],"Q4":[10,11,12]}

# KPI formula dictionary - drives hover tooltips (web) and Excel cell comments (workbooks)
KPI_FORMULAS = {
    "Actual Income":
        "Total Income = Rental Income/GPR + Other Income - Vacancy - Concessions - Bad Debt (Effective Gross Income).",
    "Budget Income":
        "Budgeted Total Income for the period.",
    "Income Variance":
        "Income Variance = Actual Income - Budget Income. Positive = favorable (above budget).",
    "Income Variance %":
        "Income Variance % = Income Variance / Budget Income.",
    "Actual Expenses":
        "Sum of all Operating Expense bucket accounts. Excludes depreciation, debt service, reserves.",
    "Budget Expenses":
        "Budgeted Operating Expenses for the period.",
    "Expense Variance":
        "Expense Variance = Actual Expenses - Budget Expenses. Negative = favorable (under budget).",
    "Expense Variance %":
        "Expense Variance % = Expense Variance / Budget Expenses.",
    "Actual NOI":
        "NOI = Total Income - Total Operating Expenses.",
    "Budget NOI":
        "Budget NOI = Budget Income - Budget Expenses.",
    "NOI Variance":
        "NOI Variance = Actual NOI - Budget NOI. Positive = favorable.",
    "NOI Variance %":
        "NOI Variance % = NOI Variance / |Budget NOI|. Absolute denominator handles sign flips when budget NOI is negative.",
    "Economic Occupancy %":
        "Economic Occ % = Net Collectible Rental Revenue / GPR. Net Collectible = GPR - Vacancy - Concessions - Bad Debt.",
    "Budget Economic Occupancy %":
        "Budget Eco Occ % = Budget Net Collectible / Budget GPR.",
    "Economic Occupancy Variance":
        "Eco Occ Variance = Actual Eco Occ % - Budget Eco Occ %.",
    "Portfolio Economic Occupancy %":
        "Weighted: Sum(Net Collectible) / Sum(GPR) across all properties. Never an average of property-level percentages.",
    "Physical Occupancy %":
        "Physical Occ % = Occupied Units / Total Units. Sourced from the uploaded Physical Occupancy Report.",
    "Leakage Gap":
        "Leakage Gap = Physical Occ % - Economic Occ %. Positive value means units are occupied but rent is not being fully collected (due to vacancy loss, concessions, or bad debt).",
    "YoY Physical Occ Variance":
        "YoY Physical Occ Variance = Current Year Physical Occ % - Prior Year Physical Occ %.",
    "YoY Economic Occ Variance":
        "YoY Economic Occ Variance = Current Year Eco Occ % - Prior Year Eco Occ %.",
    "YoY Leakage Gap Change":
        "YoY Leakage Gap Change = Current Year Leakage Gap - Prior Year Leakage Gap. Positive = gap widened (worse).",
    "Income Per Unit":
        "Income Per Unit = Actual Income / Total Units. Total Units sourced from Physical Occupancy Report.",
    "Expense Per Unit":
        "Expense Per Unit = Actual Expenses / Total Units. Total Units sourced from Physical Occupancy Report.",
    "NOI Per Unit":
        "NOI Per Unit = Actual NOI / Total Units. Total Units sourced from Physical Occupancy Report.",
    "Amount Per Unit":
        "Amount Per Unit = Account Amount / Total Units. Enables micro-level benchmarking of individual income and expense lines.",
    "Income/Unit":
        "Income Per Unit = Actual Income / Total Units. Total Units sourced from Physical Occupancy Report.",
    "Expense/Unit":
        "Expense Per Unit = Actual Expenses / Total Units. Total Units sourced from Physical Occupancy Report.",
    "NOI/Unit":
        "NOI Per Unit = Actual NOI / Total Units. Total Units sourced from Physical Occupancy Report.",
}

# Account mapping rules: list of (keyword_pattern, category, treatment, in_noi, in_eco_occ)
# Processed in order; first match wins. Patterns are lowercase substrings.
ACCOUNT_MAPPING_RULES = [
    # Excluded / non-operating (check before expense keywords)
    ("depreciation",      "Excluded", "Excluded", False, False),
    ("amortization",      "Excluded", "Excluded", False, False),
    ("debt service",      "Excluded", "Excluded", False, False),
    ("mortgage",          "Excluded", "Excluded", False, False),
    ("principal",         "Excluded", "Excluded", False, False),
    ("interest expense",  "Excluded", "Excluded", False, False),
    ("replacement reserve","Excluded","Excluded", False, False),
    ("capital",           "Excluded", "Excluded", False, False),
    ("distribution",      "Excluded", "Excluded", False, False),
    ("owner draw",        "Excluded", "Excluded", False, False),
    # Vacancy / contra-income
    ("vacancy",           "Vacancy",  "Contra-Income", True, True),
    ("vacancies",         "Vacancy",  "Contra-Income", True, True),
    # Concessions
    ("concession",        "Concessions", "Contra-Income", True, True),
    ("move-in special",   "Concessions", "Contra-Income", True, True),
    ("move in special",   "Concessions", "Contra-Income", True, True),
    # Bad debt
    ("bad debt",          "Bad Debt", "Contra-Income", True, True),
    ("collection loss",   "Bad Debt", "Contra-Income", True, True),
    ("write-off",         "Bad Debt", "Contra-Income", True, True),
    ("write off",         "Bad Debt", "Contra-Income", True, True),
    ("tenant write",      "Bad Debt", "Contra-Income", True, True),
    # Rental Income / GPR (check after vacancy so "vacancy" in name doesn't hit here)
    ("gross potential",   "Rental Income", "Income", True, True),
    ("apartment rent",    "Rental Income", "Income", True, True),
    ("tenant rent",       "Rental Income", "Income", True, True),
    ("subsidy",           "Rental Income", "Income", True, True),
    ("rental income",     "Rental Income", "Income", True, True),
    ("rent revenue",      "Rental Income", "Income", True, True),
    ("rent income",       "Rental Income", "Income", True, True),
    # Other Income
    ("laundry",           "Other Income", "Income", True, False),
    ("parking",           "Other Income", "Income", True, False),
    ("late fee",          "Other Income", "Income", True, False),
    ("application fee",   "Other Income", "Income", True, False),
    ("pet fee",           "Other Income", "Income", True, False),
    ("misc",              "Other Income", "Income", True, False),
    ("other income",      "Other Income", "Income", True, False),
    ("tenant charge",     "Other Income", "Income", True, False),
    ("fee income",        "Other Income", "Income", True, False),
    # Operating Expenses
    ("administrative",    "Operating Expense", "Expense", True, False),
    ("management fee",    "Operating Expense", "Expense", True, False),
    ("management",        "Operating Expense", "Expense", True, False),
    ("payroll",           "Operating Expense", "Expense", True, False),
    ("salary",            "Operating Expense", "Expense", True, False),
    ("wage",              "Operating Expense", "Expense", True, False),
    ("repair",            "Operating Expense", "Expense", True, False),
    ("maintenance",       "Operating Expense", "Expense", True, False),
    ("utility",           "Operating Expense", "Expense", True, False),
    ("utilities",         "Operating Expense", "Expense", True, False),
    ("insurance",         "Operating Expense", "Expense", True, False),
    ("property tax",      "Operating Expense", "Expense", True, False),
    ("real estate tax",   "Operating Expense", "Expense", True, False),
    ("income tax",        "Excluded",          "Excluded", False, False),
    ("compliance",        "Operating Expense", "Expense", True, False),
    ("marketing",         "Operating Expense", "Expense", True, False),
    ("advertising",       "Operating Expense", "Expense", True, False),
    ("contract service",  "Operating Expense", "Expense", True, False),
    ("security",          "Operating Expense", "Expense", True, False),
    ("turnover",          "Operating Expense", "Expense", True, False),
    ("grounds",           "Operating Expense", "Expense", True, False),
    ("janitorial",        "Operating Expense", "Expense", True, False),
    ("pest",              "Operating Expense", "Expense", True, False),
    ("exterminator",      "Operating Expense", "Expense", True, False),
    ("office",            "Operating Expense", "Expense", True, False),
    ("supplies",          "Operating Expense", "Expense", True, False),
    ("telephone",         "Operating Expense", "Expense", True, False),
    ("professional",      "Operating Expense", "Expense", True, False),
    ("accounting",        "Operating Expense", "Expense", True, False),
    ("legal",             "Operating Expense", "Expense", True, False),
    ("audit",             "Operating Expense", "Expense", True, False),
    ("license",           "Operating Expense", "Expense", True, False),
]
