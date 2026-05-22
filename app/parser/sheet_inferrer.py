"""Infers source type (Actual / Budget / Actual+Budget / Unknown) from sheet metadata."""

_ACTUAL_BUDGET_PATTERNS = ["actual vs budget", "actual/budget", "act vs bud", "act/bud"]
# Include leading-edge variants so "Act 2024" / "Bud Q1" sheet names match
_ACTUAL_PATTERNS = ["actual", "actuals", " act ", "act "]
_BUDGET_PATTERNS = ["budget", "budgeted", " bud ", "bud "]


def infer_sheet_type(sheet_name: str, header_row: list[str], title_rows: list[list]) -> str:
    """
    Args:
        sheet_name: The worksheet tab name.
        header_row: List of column header strings (any case; normalized internally).
        title_rows: First few rows of the sheet as lists of cell values.
    Returns:
        "Actual" | "Budget" | "Actual+Budget" | "Unknown"
    """
    name_lower = sheet_name.lower()

    # Check sheet name first
    for pat in _ACTUAL_BUDGET_PATTERNS:
        if pat in name_lower:
            return "Actual+Budget"
    for pat in _BUDGET_PATTERNS:
        if pat in name_lower:
            return "Budget"
    for pat in _ACTUAL_PATTERNS:
        if pat in name_lower:
            return "Actual"

    # Check title rows (first 5 rows, first 5 cells each — keywords can appear past col C)
    for row in title_rows[:5]:
        for cell in row[:5]:
            if not cell:
                continue
            cell_lower = str(cell).lower()
            for pat in _ACTUAL_BUDGET_PATTERNS:
                if pat in cell_lower:
                    return "Actual+Budget"
            for pat in _BUDGET_PATTERNS:
                if pat in cell_lower:
                    return "Budget"
            for pat in _ACTUAL_PATTERNS:
                if pat in cell_lower:
                    return "Actual"

    # Check column headers for alternating actual/budget pattern
    headers_lower = [str(h).lower() for h in header_row]
    has_actual_cols = any("act" in h or "actual" in h for h in headers_lower)
    has_budget_cols = any("bud" in h or "budget" in h for h in headers_lower)

    if has_actual_cols and has_budget_cols:
        return "Actual+Budget"
    if has_actual_cols:
        return "Actual"
    if has_budget_cols:
        return "Budget"

    return "Unknown"
