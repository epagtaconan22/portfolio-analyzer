"""Calculates NOI, variances, and top-2 account-level drivers per property per month."""

from collections import defaultdict
from typing import Optional
from app.models import MappedRow, PropertyPeriodKPIs
from config import MONTHS

def _safe_variance_pct(variance: Optional[float], denominator: Optional[float]) -> Optional[float]:
    if variance is None or denominator is None:
        return None
    if denominator == 0:
        return None
    return variance / abs(denominator)

def calculate_noi(mapped_rows: list[MappedRow]) -> list[PropertyPeriodKPIs]:
    """
    Returns one PropertyPeriodKPIs per (property, year, month) for each
    property/year/month combination present in mapped_rows.
    Actual and Budget rows are processed separately and merged.
    Sign convention: Income = positive, Expense = positive cost (subtracted), Contra-Income = positive deduction.
    """
    Actual = "Actual"
    Budget = "Budget"

    # accumulators: key = (property_name, pm_name, year, month, source_type)
    income_acc: dict[tuple, float] = defaultdict(float)
    expense_acc: dict[tuple, float] = defaultdict(float)
    # For driver analysis: (property, pm, year, month, source_type, account_name) -> net_contribution
    account_acc: dict[tuple, float] = defaultdict(float)

    for row in mapped_rows:
        if not row.include_in_noi:
            continue
        source = row.source_type if row.source_type in (Actual, Budget) else Actual
        key = (row.property_name, row.pm_name, row.year, row.month, source)

        if row.treatment == "Income":
            income_acc[key] += row.amount
            account_acc[key + (row.account_name,)] += row.amount
        elif row.treatment == "Contra-Income":
            # Normalize: subtract from income (amount should be positive = a cost)
            amount = abs(row.amount)
            income_acc[key] -= amount
            account_acc[key + (row.account_name,)] -= amount
        elif row.treatment == "Expense":
            # Normalize: amount should be positive = a cost
            amount = abs(row.amount)
            expense_acc[key] += amount
            account_acc[key + (row.account_name,)] -= amount  # negative = cost in driver analysis

    # Collect all (property, pm, year, month) combos
    all_keys: set[tuple] = set()
    for (prop, pm, yr, mo, src) in list(income_acc.keys()) + list(expense_acc.keys()):
        all_keys.add((prop, pm, yr, mo))

    result: list[PropertyPeriodKPIs] = []
    for (prop, pm, yr, mo) in sorted(all_keys):
        act_key = (prop, pm, yr, mo, Actual)
        bud_key = (prop, pm, yr, mo, Budget)

        actual_income = income_acc.get(act_key)
        actual_expenses = expense_acc.get(act_key)
        budget_income = income_acc.get(bud_key)
        budget_expenses = expense_acc.get(bud_key)

        actual_noi = None
        if actual_income is not None and actual_expenses is not None:
            actual_noi = actual_income - actual_expenses
        elif actual_income is not None:
            actual_noi = actual_income
        elif actual_expenses is not None:
            actual_noi = -actual_expenses

        budget_noi = None
        if budget_income is not None and budget_expenses is not None:
            budget_noi = budget_income - budget_expenses
        elif budget_income is not None:
            budget_noi = budget_income
        elif budget_expenses is not None:
            budget_noi = -budget_expenses

        inc_var = (actual_income - budget_income) if (actual_income is not None and budget_income is not None) else None
        exp_var = (actual_expenses - budget_expenses) if (actual_expenses is not None and budget_expenses is not None) else None
        noi_var = (actual_noi - budget_noi) if (actual_noi is not None and budget_noi is not None) else None

        # Driver analysis: compare actual vs budget at account level
        driver1, driver2 = _top_drivers(account_acc, prop, pm, yr, mo)

        kpis = PropertyPeriodKPIs(
            property_name=prop,
            pm_name=pm,
            year=yr,
            month=mo,
            period=MONTHS.get(mo, str(mo)),
            actual_income=actual_income,
            budget_income=budget_income,
            income_variance=inc_var,
            income_variance_pct=_safe_variance_pct(inc_var, budget_income),
            actual_expenses=actual_expenses,
            budget_expenses=budget_expenses,
            expense_variance=exp_var,
            expense_variance_pct=_safe_variance_pct(exp_var, budget_expenses),
            actual_noi=actual_noi,
            budget_noi=budget_noi,
            noi_variance=noi_var,
            noi_variance_pct=_safe_variance_pct(noi_var, budget_noi),
            top_noi_driver_1=driver1,
            top_noi_driver_2=driver2,
        )
        result.append(kpis)

    return result


def _top_drivers(
    account_acc: dict,
    prop: str, pm: str, yr: int, mo: int,
) -> tuple[str, str]:
    """Identify top-2 account-level contributors to NOI variance (actual vs budget)."""
    Actual, Budget = "Actual", "Budget"
    act_prefix = (prop, pm, yr, mo, Actual)
    bud_prefix = (prop, pm, yr, mo, Budget)

    # Gather all account names for this property/period
    account_names: set[str] = set()
    for key in account_acc:
        if key[:5] in (act_prefix, bud_prefix):
            account_names.add(key[5])

    variances: list[tuple[float, str]] = []
    for name in account_names:
        act_val = account_acc.get(act_prefix + (name,), 0)
        bud_val = account_acc.get(bud_prefix + (name,), 0)
        var = act_val - bud_val
        variances.append((abs(var), f"{name} (${var:+,.0f})"))

    variances.sort(reverse=True)
    driver1 = variances[0][1] if len(variances) > 0 else ""
    driver2 = variances[1][1] if len(variances) > 1 else ""
    return driver1, driver2
