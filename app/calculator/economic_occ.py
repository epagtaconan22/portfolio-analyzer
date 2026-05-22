"""Enriches PropertyPeriodKPIs with economic occupancy data derived from financial statements."""

from collections import defaultdict
from typing import Optional
from app.models import MappedRow, PropertyPeriodKPIs

_ECO_OCC_CATS = {"Rental Income", "Vacancy", "Concessions", "Bad Debt"}

def enrich_eco_occ(
    mapped_rows: list[MappedRow],
    kpis: list[PropertyPeriodKPIs],
) -> list[PropertyPeriodKPIs]:
    """
    Enriches each KPI record with gpr, vacancy, concessions, bad_debt,
    net_collectible, eco_occ_pct, budget_eco_occ_pct, eco_occ_variance.
    Modifies kpis in-place and returns them.
    """
    # Accumulate: (property, pm, year, month, source_type, category) -> amount
    acc: dict[tuple, float] = defaultdict(float)
    for row in mapped_rows:
        if not row.include_in_eco_occ:
            continue
        src = row.source_type if row.source_type in ("Actual", "Budget") else "Actual"
        key = (row.property_name, row.pm_name, row.year, row.month, src, row.account_category)
        # For Contra-Income, store the raw signed amount so that negative values
        # (sign errors in source data) propagate correctly and produce eco_occ > 1.0,
        # which Quality Checks can then flag. The formula gpr - v - c - b handles
        # the subtraction; no abs() normalisation is applied here.
        amount = row.amount
        acc[key] += amount

    for kpi in kpis:
        prop, pm, yr, mo = kpi.property_name, kpi.pm_name, kpi.year, kpi.month

        for src_type, is_budget in [("Actual", False), ("Budget", True)]:
            def get(cat, _src=src_type):
                return acc.get((prop, pm, yr, mo, _src, cat), None)

            gpr = get("Rental Income")
            vacancy = get("Vacancy")
            concessions = get("Concessions")
            bad_debt = get("Bad Debt")

            if gpr is not None:
                v = vacancy or 0
                c = concessions or 0
                b = bad_debt or 0
                net_collectible = gpr - v - c - b
                eco_occ = net_collectible / gpr if gpr != 0 else None
            else:
                net_collectible = None
                eco_occ = None

            if not is_budget:
                kpi.gpr = gpr
                kpi.vacancy = vacancy
                kpi.concessions = concessions
                kpi.bad_debt = bad_debt
                kpi.net_collectible = net_collectible
                kpi.eco_occ_pct = eco_occ
            else:
                kpi.budget_eco_occ_pct = eco_occ

        if kpi.eco_occ_pct is not None and kpi.budget_eco_occ_pct is not None:
            kpi.eco_occ_variance = kpi.eco_occ_pct - kpi.budget_eco_occ_pct

    return kpis
