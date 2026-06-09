"""KPI display definitions — extracted from routes/results.py."""

# Tooltip text for KPI labels
KPI_TOOLTIPS: dict[str, str] = {
    "Actual Income":      "Total Income = GPR + Other Income − Vacancy − Concessions − Bad Debt (Effective Gross Income)",
    "Budget Income":      "Budgeted Total Income for the period",
    "Income Variance":    "Actual Income − Budget Income. Positive = favorable (above budget)",
    "Income Variance %":  "Income Variance / Budget Income",
    "Actual Expenses":    "Sum of all Operating Expense accounts. Excludes depreciation, debt service, reserves",
    "Budget Expenses":    "Budgeted Operating Expenses for the period",
    "Expense Variance":   "Actual Expenses − Budget Expenses. Negative = favorable (under budget)",
    "Expense Variance %": "Expense Variance / Budget Expenses",
    "Actual NOI":         "NOI = Total Income − Total Operating Expenses",
    "Budget NOI":         "Budget NOI = Budget Income − Budget Expenses",
    "NOI Variance":       "NOI Variance = Actual NOI − Budget NOI. Positive = favorable",
    "NOI Variance %":     "NOI Variance / |Budget NOI|. Absolute denominator handles sign flips when budget NOI is negative",
    "GPR":                "Gross Potential Rent — total scheduled rent before any deductions",
    "Vacancy":            "Vacancy loss — rent foregone from unoccupied units",
    "Concessions":        "Move-in specials and rent concessions",
    "Bad Debt":           "Collection losses and write-offs",
    "Net Collectible":    "GPR − Vacancy − Concessions − Bad Debt",
    "Eco Occ %":          "Economic Occupancy % = Net Collectible / GPR",
    "Budget Eco Occ %":   "Budget Economic Occupancy % = Budget Net Collectible / Budget GPR",
    "Eco Occ Variance":   "Actual Eco Occ % − Budget Eco Occ %",
    "Physical Occ %":     "Physical Occ % = Occupied Units / Total Units. Sourced from Physical Occupancy Report",
    "Leakage Gap":        "Physical Occ % − Economic Occ %. Positive = units occupied but rent not being fully collected",
    "Income/Unit":        "Actual Income / Total Units (from Physical Occupancy Report)",
    "Expense/Unit":       "Actual Expenses / Total Units (from Physical Occupancy Report)",
    "NOI/Unit":           "Actual NOI / Total Units (from Physical Occupancy Report)",
}

# (label, key, fmt, favorable_positive, group_id, is_group_header)
# None entries are visual separators (blank rows in the table).
SUMMARY_KPI_DEFINITIONS = [
    ("Actual Income",      "actual_income",        "currency", None,  "group_income",   True),
    ("Budget Income",      "budget_income",        "currency", None,  "group_income",   False),
    ("Income Variance",    "income_variance",      "currency", True,  "group_income",   False),
    None,
    ("Actual Expenses",    "actual_expenses",      "currency", None,  "group_expenses", True),
    ("Budget Expenses",    "budget_expenses",      "currency", None,  "group_expenses", False),
    ("Expense Variance",   "expense_variance",     "currency", False, "group_expenses", False),
    None,
    ("Actual NOI",         "actual_noi",           "currency", None,  "group_noi",      True),
    ("Budget NOI",         "budget_noi",           "currency", None,  "group_noi",      False),
    ("NOI Variance",       "noi_variance",         "currency", True,  "group_noi",      False),
    None,
    ("GPR",                "gpr",                  "currency", None,  "group_gpr",      True),
    ("Vacancy",            "vacancy",              "currency", None,  "group_gpr",      False),
    ("Concessions",        "concessions",          "currency", None,  "group_gpr",      False),
    ("Bad Debt",           "bad_debt",             "currency", None,  "group_gpr",      False),
    ("Net Collectible",    "net_collectible",      "currency", None,  "group_gpr",      False),
    None,
    ("Eco Occ %",          "eco_occ_pct",          "pct",      None,  "group_eco_occ",  True),
    ("Budget Eco Occ %",   "budget_eco_occ_pct",   "pct",      None,  "group_eco_occ",  False),
    ("Eco Occ Variance",   "eco_occ_variance",     "pct",      True,  "group_eco_occ",  False),
    ("Physical Occ %",     "physical_occ_pct",     "pct",      None,  None,             False),
    ("Leakage Gap",        "leakage_gap",          "pct",      False, None,             False),
    None,
    ("Income/Unit",        "income_per_unit",      "currency", None,  None,             False),
    ("Expense/Unit",       "expense_per_unit",     "currency", None,  None,             False),
    ("NOI/Unit",           "noi_per_unit",         "currency", None,  None,             False),
]

# Maps actual KPI key → the budget equivalent for YoY budget comparison
BUDGET_YOY_KEY: dict[str, str] = {
    "actual_income":        "budget_income",
    "budget_income":        "budget_income",
    "income_variance":      "income_variance",
    "income_variance_pct":  "income_variance_pct",
    "actual_expenses":      "budget_expenses",
    "budget_expenses":      "budget_expenses",
    "expense_variance":     "expense_variance",
    "expense_variance_pct": "expense_variance_pct",
    "actual_noi":           "budget_noi",
    "budget_noi":           "budget_noi",
    "noi_variance":         "noi_variance",
    "noi_variance_pct":     "noi_variance_pct",
    "eco_occ_pct":          "budget_eco_occ_pct",
}

YOY_CURRENCY_KEYS = frozenset({
    "actual_income", "budget_income", "income_variance",
    "actual_expenses", "budget_expenses", "expense_variance",
    "actual_noi", "budget_noi", "noi_variance",
})

YOY_FAVORABLE_IF_POSITIVE = frozenset({
    "actual_income", "budget_income",
    "actual_noi",    "budget_noi",    "noi_variance",
    "eco_occ_pct",
})

PCT_VARIANCE_THRESHOLD_KEYS = frozenset({
    "income_variance_pct", "expense_variance_pct",
    "noi_variance_pct",    "eco_occ_variance",
})
