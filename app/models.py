from dataclasses import dataclass, field
from typing import Optional

@dataclass
class RawRow:
    property_name: str
    pm_name: str
    source_workbook: str
    source_sheet: str
    source_type: str        # "Actual" | "Budget" | "Actual+Budget" | "Unknown"
    source_row: int
    account_code: str
    account_name: str
    year: int
    month: int              # 1-12
    amount: float
    original_amount: float
    notes: str = ""

@dataclass
class MappedRow:
    property_name: str
    pm_name: str
    source_workbook: str
    source_sheet: str
    source_type: str
    source_row: int
    account_code: str
    account_name: str
    year: int
    month: int
    amount: float
    original_amount: float
    account_category: str   # "Rental Income" | "Other Income" | "Vacancy" |
                            # "Concessions" | "Bad Debt" | "Operating Expense" |
                            # "Excluded" | "Review Needed"
    kpi_mapping: str
    include_in_noi: bool
    include_in_eco_occ: bool
    treatment: str          # "Income" | "Contra-Income" | "Expense" | "Excluded" | "Review Needed"
    notes: str = ""

@dataclass
class OccupancyRow:
    property_name: str
    year: int
    month: int
    occupied_units: int
    total_units: int

    @property
    def physical_occ_pct(self) -> Optional[float]:
        return self.occupied_units / self.total_units if self.total_units > 0 else None

@dataclass
class PropertyPeriodKPIs:
    property_name: str
    pm_name: str
    year: int
    month: int              # 0 = period-level aggregate
    period: str             # "Full Year" | "Q1" | "Jan" | etc.
    # Income
    actual_income: Optional[float] = None
    budget_income: Optional[float] = None
    income_variance: Optional[float] = None
    income_variance_pct: Optional[float] = None
    # Expenses
    actual_expenses: Optional[float] = None
    budget_expenses: Optional[float] = None
    expense_variance: Optional[float] = None
    expense_variance_pct: Optional[float] = None
    # NOI
    actual_noi: Optional[float] = None
    budget_noi: Optional[float] = None
    noi_variance: Optional[float] = None
    noi_variance_pct: Optional[float] = None
    # Economic occupancy components
    gpr: Optional[float] = None
    vacancy: Optional[float] = None
    concessions: Optional[float] = None
    bad_debt: Optional[float] = None
    net_collectible: Optional[float] = None
    eco_occ_pct: Optional[float] = None
    budget_eco_occ_pct: Optional[float] = None
    eco_occ_variance: Optional[float] = None
    # Physical occupancy
    total_units: Optional[int] = None
    occupied_units: Optional[int] = None
    physical_occ_pct: Optional[float] = None
    leakage_gap: Optional[float] = None
    yoy_physical_occ_variance: Optional[float] = None
    yoy_eco_occ_variance: Optional[float] = None
    yoy_leakage_gap_change: Optional[float] = None
    # Per unit (requires total_units from occupancy report)
    income_per_unit: Optional[float] = None
    expense_per_unit: Optional[float] = None
    noi_per_unit: Optional[float] = None
    # Drivers and commentary
    top_noi_driver_1: str = ""
    top_noi_driver_2: str = ""
    top_eco_occ_driver_1: str = ""
    top_eco_occ_driver_2: str = ""
    commentary: str = ""
    source_key: str = ""
    is_below_eco_occ_target: bool = False
    is_carveout: bool = False
    # Property metadata (sourced from PROPERTY_METADATA in config.py)
    city: str = ""
    tenancy_type: str = ""

@dataclass
class SourceIndexEntry:
    source_workbook: str
    source_sheet: str
    property_name: str
    pm_name: str
    year: int
    source_type: str
    processed: bool
    rows_extracted: int
    reason_if_excluded: str = ""
    notes: str = ""

@dataclass
class MappingEntry:
    account_code: str
    account_name: str
    assigned_category: str
    kpi_mapping: str
    treatment: str
    include_in_noi: bool
    include_in_eco_occ: bool
    notes: str = ""

@dataclass
class QualityCheck:
    check_name: str
    passed: bool
    detail: str = ""

@dataclass
class ARAgingRow:
    property_name: str       # After PROPERTY_NAME_MAP normalization
    pm_name: str             # Extracted from filename prefix before first "_"
    source_file: str         # Basename of source file
    receivable_type: str     # "Tenant Rent" | "Subsidy"
    year: int
    month: int               # 1–12
    charge_amount: float     # Col 1
    current_owed: float      # Col 2
    owed_0_30: float         # Col 3
    owed_31_60: float        # Col 4
    owed_61_90: float        # Col 5
    owed_over_90: float      # Col 6
    prepayments: float       # Col 7 (negative = credits)

    @property
    def total_over_60(self) -> float:
        """61-90 + Over-90 — amounts more than 60 days past due."""
        return self.owed_61_90 + self.owed_over_90

    @property
    def pct_over_60(self) -> Optional[float]:
        """% of charge_amount that is >60 days past due."""
        if self.charge_amount and self.charge_amount > 0:
            return self.total_over_60 / self.charge_amount
        return None

    # Legacy alias kept for any downstream callers that pre-date the >60 change
    @property
    def total_overdue(self) -> float:
        """Alias for total_over_60 (renamed from >30 to >60 days)."""
        return self.total_over_60

    @property
    def pct_overdue(self) -> Optional[float]:
        """Alias for pct_over_60 (renamed from >30 to >60 days)."""
        return self.pct_over_60
