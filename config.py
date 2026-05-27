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
    # ConAm Yardi income netting entries (40000-series; always net to zero via Memo Contra)
    ("market rent",       "Excluded", "Excluded", False, False),
    ("loss/gain to lease","Excluded", "Excluded", False, False),
    ("memo contra",       "Excluded", "Excluded", False, False),
    # Capital expenditure abbreviation (catches "Cap Exp-..." and "Sal Cap Exp-...")
    ("cap exp",           "Excluded", "Excluded", False, False),
    # Abbreviated depreciation / interest (Yardi ConAm format)
    ("depr exp",          "Excluded", "Excluded", False, False),
    ("interest exp",      "Excluded", "Excluded", False, False),
    # Partnership expenses (ConAm prefix "Prtnshp Exp-...")
    ("prtnshp",           "Excluded", "Excluded", False, False),
    # Financing fees excluded from NOI
    ("fee exp-loan",      "Excluded", "Excluded", False, False),
    ("remarketing",       "Excluded", "Excluded", False, False),
    # Vacancy / contra-income
    ("vacancy",           "Vacancy",  "Contra-Income", True, True),
    ("vacancies",         "Vacancy",  "Contra-Income", True, True),
    # Employee unit occupancy (units held by staff at no/reduced rent)
    ("employee unit",     "Vacancy",  "Contra-Income", True, True),
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
    # ConAm "Base Scheduled Rent" detail lines (41000-1000 tenant + 41000-1100 subsidy caught above)
    # The 41000-1798 subtotal row is skipped via account-code detection in _is_skip_row.
    ("base scheduled rent","Rental Income","Income", True, True),
    # Other Income
    ("laundry",           "Other Income", "Income", True, False),
    ("parking",           "Other Income", "Income", True, False),
    ("late fee",          "Other Income", "Income", True, False),
    ("application fee",   "Other Income", "Income", True, False),
    ("pet fee",           "Other Income", "Income", True, False),
    ("miscellaneous income", "Other Income",      "Income",  True, False),
    ("misc income",          "Other Income",      "Income",  True, False),
    ("misc expense",         "Operating Expense", "Expense", True, False),
    ("miscellaneous expense","Operating Expense", "Expense", True, False),
    ("misc",                 "Other Income",      "Income",  True, False),
    ("other income",      "Other Income", "Income", True, False),
    ("tenant charge",     "Other Income", "Income", True, False),
    ("fee income",        "Other Income", "Income", True, False),
    # ConAm-specific Other Income
    ("lease termination", "Other Income", "Income", True, False),
    ("employee rent",     "Other Income", "Income", True, False),
    # Operating Expenses
    ("administrative",    "Operating Expense", "Expense", True, False),
    ("management fee",    "Operating Expense", "Expense", True, False),
    ("management",        "Operating Expense", "Expense", True, False),
    ("payroll",           "Operating Expense", "Expense", True, False),
    ("salary",            "Operating Expense", "Expense", True, False),
    ("wage",              "Operating Expense", "Expense", True, False),
    # ConAm salary abbreviations ("Sal P/M-...", "Sal Maint-...", "Sal Burden-...", "Sal-...")
    ("sal p/m",           "Operating Expense", "Expense", True, False),
    ("sal maint",         "Operating Expense", "Expense", True, False),
    ("sal burden",        "Operating Expense", "Expense", True, False),
    ("sal-",              "Operating Expense", "Expense", True, False),
    ("sal ",              "Operating Expense", "Expense", True, False),
    # Personnel/HR costs (training, travel, employee relations)
    ("personnel",         "Operating Expense", "Expense", True, False),
    ("repair",            "Operating Expense", "Expense", True, False),
    ("maintenance",       "Operating Expense", "Expense", True, False),
    # Repair & Maintenance abbreviated ("R & M-General", "R & M-HVAC", etc.)
    ("r & m",             "Operating Expense", "Expense", True, False),
    ("utility",           "Operating Expense", "Expense", True, False),
    ("utilities",         "Operating Expense", "Expense", True, False),
    ("insurance",         "Operating Expense", "Expense", True, False),
    # Insurance abbreviated ("Ins - Property And Casualty", "Ins - General Liability")
    ("ins - ",            "Operating Expense", "Expense", True, False),
    ("property tax",      "Operating Expense", "Expense", True, False),
    ("real estate tax",   "Operating Expense", "Expense", True, False),
    ("income tax",        "Excluded",          "Excluded", False, False),
    # Property tax variants ("Taxes-Real Estate", "Taxes-Personal Property"; income tax excluded above)
    ("taxes-",            "Operating Expense", "Expense", True, False),
    ("compliance",        "Operating Expense", "Expense", True, False),
    ("marketing",         "Operating Expense", "Expense", True, False),
    ("advertising",       "Operating Expense", "Expense", True, False),
    # Advertising abbreviated ("Adv-Periodicals", "Adv-Apt Guide")
    ("adv-",              "Operating Expense", "Expense", True, False),
    # Marketing/promotion variants
    ("promotion",         "Operating Expense", "Expense", True, False),
    ("website",           "Operating Expense", "Expense", True, False),
    ("lease up",          "Operating Expense", "Expense", True, False),
    ("contract service",  "Operating Expense", "Expense", True, False),
    # Fee expenses — loan-related excluded above; catch all others here
    ("fee exp-",          "Operating Expense", "Expense", True, False),
    # Collection agency fees
    ("collections",       "Operating Expense", "Expense", True, False),
    # Resident activities
    ("activit",           "Operating Expense", "Expense", True, False),
    # Office duplicating/copying
    ("duplicating",       "Operating Expense", "Expense", True, False),
    ("security deposit",  "Other Income",      "Income",  True, False),
    ("security service",  "Operating Expense", "Expense", True, False),
    ("security guard",    "Operating Expense", "Expense", True, False),
    ("security",          "Operating Expense", "Expense", True, False),
    ("turnover",          "Operating Expense", "Expense", True, False),
    ("grounds",           "Operating Expense", "Expense", True, False),
    ("janitorial",        "Operating Expense", "Expense", True, False),
    ("pest",              "Operating Expense", "Expense", True, False),
    ("exterminating",     "Operating Expense", "Expense", True, False),
    ("exterminator",      "Operating Expense", "Expense", True, False),
    ("office",            "Operating Expense", "Expense", True, False),
    ("supplies",          "Operating Expense", "Expense", True, False),
    ("telephone",         "Operating Expense", "Expense", True, False),
    ("professional",      "Operating Expense", "Expense", True, False),
    ("accounting",        "Operating Expense", "Expense", True, False),
    ("legal",             "Operating Expense", "Expense", True, False),
    ("audit",             "Operating Expense", "Expense", True, False),
    ("license",           "Operating Expense", "Expense", True, False),
    # Utilities (individual utility types not covered by "utility"/"utilities")
    ("electricity",       "Operating Expense", "Expense", True, False),
    ("electric",          "Operating Expense", "Expense", True, False),
    ("water",             "Operating Expense", "Expense", True, False),
    ("gas",               "Operating Expense", "Expense", True, False),
    ("sewer",             "Operating Expense", "Expense", True, False),
    ("garbage",           "Operating Expense", "Expense", True, False),
    ("trash",             "Operating Expense", "Expense", True, False),
    ("internet",          "Operating Expense", "Expense", True, False),
    ("cable",             "Operating Expense", "Expense", True, False),
    ("dsl",               "Operating Expense", "Expense", True, False),
    # Additional operating expense categories common in Yardi/LIHTC
    ("bank charge",       "Operating Expense", "Expense", True, False),
    ("bank fee",          "Operating Expense", "Expense", True, False),
    ("processing fee",    "Operating Expense", "Expense", True, False),
    ("payroll processing","Operating Expense", "Expense", True, False),
    ("plumbing",          "Operating Expense", "Expense", True, False),
    ("appliance",         "Operating Expense", "Expense", True, False),
    ("hvac",              "Operating Expense", "Expense", True, False),
    ("elevator",          "Operating Expense", "Expense", True, False),
    ("electrical",        "Operating Expense", "Expense", True, False),
    ("roof",              "Operating Expense", "Expense", True, False),
    ("painting",          "Operating Expense", "Expense", True, False),
    ("flooring",          "Operating Expense", "Expense", True, False),
    ("cleaning",          "Operating Expense", "Expense", True, False),
    ("alarm",             "Operating Expense", "Expense", True, False),
    ("fire",              "Operating Expense", "Expense", True, False),
    ("sprinkler",         "Operating Expense", "Expense", True, False),
    ("landscape",         "Operating Expense", "Expense", True, False),
    ("pool",              "Operating Expense", "Expense", True, False),
    ("vehicle",           "Operating Expense", "Expense", True, False),
    ("auto",              "Operating Expense", "Expense", True, False),
    ("postage",           "Operating Expense", "Expense", True, False),
    ("printing",          "Operating Expense", "Expense", True, False),
    ("dues",              "Operating Expense", "Expense", True, False),
    ("subscriptions",     "Operating Expense", "Expense", True, False),
    ("computer",          "Operating Expense", "Expense", True, False),
    ("software",          "Operating Expense", "Expense", True, False),
    ("training",          "Operating Expense", "Expense", True, False),
    ("education",         "Operating Expense", "Expense", True, False),
    ("resident service",  "Operating Expense", "Expense", True, False),
    ("social service",    "Operating Expense", "Expense", True, False),
    ("key",               "Operating Expense", "Expense", True, False),
    ("lock",              "Operating Expense", "Expense", True, False),
    ("credit check",      "Operating Expense", "Expense", True, False),
    ("background check",  "Operating Expense", "Expense", True, False),
    ("screening",         "Operating Expense", "Expense", True, False),
    ("snow",              "Operating Expense", "Expense", True, False),
    ("signage",           "Operating Expense", "Expense", True, False),
    ("storage",           "Operating Expense", "Expense", True, False),
    ("moving",            "Operating Expense", "Expense", True, False),
    # Concession-style items
    ("free unit",         "Concessions", "Contra-Income", True, True),
    # Interest income (non-operating, exclude from NOI)
    ("interest income",   "Excluded", "Excluded", False, False),
    ("oth inc-interest",  "Excluded", "Excluded", False, False),
    ("interest repl",     "Excluded", "Excluded", False, False),
    ("interest from",     "Excluded", "Excluded", False, False),
    ("interest on bond",  "Excluded", "Excluded", False, False),
    ("unrealized gain",   "Excluded", "Excluded", False, False),
    ("unrealized loss",   "Excluded", "Excluded", False, False),
    # LIHTC financing fees (below-the-line; excluded from operating NOI)
    ("bond fee",          "Excluded", "Excluded", False, False),
    ("bond/loan",         "Excluded", "Excluded", False, False),
    ("issuer fee",        "Excluded", "Excluded", False, False),
    ("trustee fee",       "Excluded", "Excluded", False, False),
    ("ground lease",      "Excluded", "Excluded", False, False),
    ("gp administration", "Excluded", "Excluded", False, False),
    ("admin obligation",  "Excluded", "Excluded", False, False),
    # ConAm-specific non-operating items
    ("permanent loan",    "Excluded", "Excluded", False, False),
    ("swap",              "Excluded", "Excluded", False, False),
    ("grant revenue",     "Excluded", "Excluded", False, False),
    # Additional operating expenses common in Yardi/LIHTC
    ("uniform",           "Operating Expense", "Expense", True, False),
    ("drapery",           "Operating Expense", "Expense", True, False),
    ("carpet",            "Operating Expense", "Expense", True, False),
    ("consulting",        "Operating Expense", "Expense", True, False),
    ("courtesy patrol",   "Operating Expense", "Expense", True, False),
    ("surveillance",      "Operating Expense", "Expense", True, False),
    ("monitoring",        "Operating Expense", "Expense", True, False),
    ("solar",             "Operating Expense", "Expense", True, False),
    ("recruitment",       "Operating Expense", "Expense", True, False),
    ("tenant service",    "Operating Expense", "Expense", True, False),
    ("resident program",  "Operating Expense", "Expense", True, False),
    # Equipment expenses (ConAm: "Equip Exp-Copy Machine", "Equip Exp-Off Equip")
    ("equip exp",         "Operating Expense", "Expense", True, False),
    # Other misc expenses (ConAm: "Oth Exp-Supportive Services", "Oth Exp-Relocation")
    ("oth exp",           "Operating Expense", "Expense", True, False),
    # Insurance: MIP (Mortgage Insurance Premium — FHA/HUD loan operating cost)
    ("ins exp",           "Operating Expense", "Expense", True, False),
    # Events and decorations (resident community events)
    ("events",            "Operating Expense", "Expense", True, False),
    ("decorat",           "Operating Expense", "Expense", True, False),
    # Tenant damage fee income
    ("damage fee",        "Other Income", "Income", True, False),
    # NSF / late charges → Other Income
    ("nsf",               "Other Income", "Income", True, False),
    ("late charge",       "Other Income", "Income", True, False),
    ("returned check",    "Other Income", "Income", True, False),
]

# Canonical property name map — normalizes the raw names the parser infers from
# sheet names / Yardi title rows to the short display names used in the web app
# and Excel workbooks.  Keys are the exact strings the parser may produce;
# values are the canonical names.  Applied during upload so all downstream
# storage, calculations, and exports use the canonical names.
PROPERTY_NAME_MAP: dict[str, str] = {
    # ── Full names (from financial workbook sheet names) ──────────────────────
    "4760 W. Melrose Ave.":            "Emerson",
    "Allanza Apt. Homes":              "Allanza",
    "Alora Family":                    "Alora",
    "Arbor Green Residences":          "Arbor Green",
    "Aurora Apartments":               "Aurora",
    "Connections Housing":             "Connections",
    "Dahlia Apartments":               "Dahlia",
    "Estrella Apartments":             "Estrella",
    "Luna at Pacific Highlands Ranch": "Luna",
    "Monte Vista Apartments":          "Monte Vista I",
    "Monte Vista II Family Housing":   "Monte Vista II",
    "Nova Apartments":                 "Nova",
    "Orange Gardens Apartments":       "Orange Gardens",
    "Sage Pointe Apartments":          "Sage Pointe",
    "Solterra Senior Residences":      "Solterra",
    "Sonoma Court Apartments":         "Sonoma Court",
    "Ten Fifty B Street":               "1050B I",
    "Ten Fifty B Street Hsg Ptrs":     "1050B II",
    "Ventaliso Apartments":            "Ventaliso",
    "Vitalia (Bascom) Apts.":          "Vitalia",
    # ── Yardi 26-char truncated names (stored literally in occupancy cells) ───
    # Yardi truncates property names longer than 23 characters and appends "..."
    # so the cell contains exactly [first 23 chars + "..."] = 26 chars total.
    # These entries ensure the occupancy–financial join succeeds even when the
    # occupancy report stores the truncated form.
    "Orange Gardens Apartmen...":      "Orange Gardens",
    "Ten Fifty B Street Hsg ...":      "1050B II",
    "Monte Vista II Family H...":      "Monte Vista II",
    "Solterra Senior Residen...":      "Solterra",
    "Luna at Pacific Highlan...":      "Luna",
}
