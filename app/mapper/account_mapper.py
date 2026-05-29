"""Maps raw account names to buckets using keyword rules from config."""

from config import ACCOUNT_MAPPING_RULES, JSCO_ACCOUNT_CODE_RULES
from app.models import RawRow, MappedRow, MappingEntry

# (category, treatment, include_in_noi, include_in_eco_occ)
MappingTuple = tuple[str, str, bool, bool]


def map_account_name(
    account_name: str,
    custom_mapping: dict[str, MappingTuple] | None = None,
    account_code: str = "",
) -> MappingTuple:
    """
    Returns (category, treatment, include_in_noi, include_in_eco_occ).

    Lookup priority:
      1. custom_mapping (user-uploaded CSV overrides)
      2. JSCO_ACCOUNT_CODE_RULES (MR-prefix code exact match)
      3. ACCOUNT_MAPPING_RULES (keyword substring match on account name)
      4. "Review Needed" fallback

    custom_mapping keys are lowercase account names.
    """
    name_lower = account_name.lower().strip()

    # 1. User-provided custom mapping takes highest priority.
    if custom_mapping:
        if name_lower in custom_mapping:
            return custom_mapping[name_lower]

    # 2. JSCO code-based exact match (MR-prefix codes are deterministic).
    if account_code:
        code_upper = account_code.upper().strip()
        if code_upper in JSCO_ACCOUNT_CODE_RULES:
            return JSCO_ACCOUNT_CODE_RULES[code_upper]

    # 3. Keyword substring matching on account name.
    for keyword, category, treatment, in_noi, in_eco in ACCOUNT_MAPPING_RULES:
        if keyword in name_lower:
            return category, treatment, in_noi, in_eco

    return "Review Needed", "Review Needed", False, False


def map_rows(
    raw_rows: list[RawRow],
    custom_mapping: dict[str, MappingTuple] | None = None,
) -> tuple[list[MappedRow], list[MappingEntry]]:
    """
    Maps all RawRows to MappedRows. Also returns a deduplicated MappingEntry
    list for the Assumptions_Mapping backup tab.
    """
    mapped: list[MappedRow] = []
    seen: dict[str, MappingEntry] = {}

    for row in raw_rows:
        cat, treatment, in_noi, in_eco = map_account_name(
            row.account_name, custom_mapping, account_code=row.account_code
        )
        kpi_mapping = _kpi_mapping_label(cat)

        mapped.append(MappedRow(
            property_name=row.property_name,
            pm_name=row.pm_name,
            source_workbook=row.source_workbook,
            source_sheet=row.source_sheet,
            source_type=row.source_type,
            source_row=row.source_row,
            account_code=row.account_code,
            account_name=row.account_name,
            year=row.year,
            month=row.month,
            amount=row.amount,
            original_amount=row.original_amount,
            notes=row.notes,
            account_category=cat,
            kpi_mapping=kpi_mapping,
            include_in_noi=in_noi,
            include_in_eco_occ=in_eco,
            treatment=treatment,
        ))

        key = f"{row.account_code}|{row.account_name.lower()}"
        if key not in seen:
            seen[key] = MappingEntry(
                account_code=row.account_code,
                account_name=row.account_name,
                assigned_category=cat,
                kpi_mapping=kpi_mapping,
                treatment=treatment,
                include_in_noi=in_noi,
                include_in_eco_occ=in_eco,
            )

    return mapped, list(seen.values())


def _kpi_mapping_label(category: str) -> str:
    return {
        "Rental Income": "GPR / Rental Income",
        "Other Income":  "Other Income",
        "Vacancy":       "Vacancy",
        "Concessions":   "Concessions",
        "Bad Debt":      "Bad Debt",
        "Operating Expense": "Operating Expense",
        "Excluded":      "Excluded",
        "Review Needed": "Review Needed",
    }.get(category, "Review Needed")
