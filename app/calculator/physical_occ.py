"""Enriches KPIs with physical occupancy, leakage gap, YoY variances, and per-unit metrics."""

from typing import Optional
from app.models import OccupancyRow, PropertyPeriodKPIs


def _yoy_delta(curr: Optional[float], prev: Optional[float]) -> Optional[float]:
    """Return curr - prev, or None if either value is absent."""
    if curr is None or prev is None:
        return None
    return curr - prev


def enrich_physical_occ(
    occ_rows: list[OccupancyRow],
    kpis: list[PropertyPeriodKPIs],
) -> list[PropertyPeriodKPIs]:
    """
    Enriches each KPI with physical occ data. Modifies kpis in-place and returns them.
    Properties with no matching occupancy row get None for all physical/per-unit fields.
    """
    occ_index: dict[tuple, OccupancyRow] = {}
    for row in occ_rows:
        occ_index[(row.property_name.lower(), row.year, row.month)] = row

    for kpi in kpis:
        occ = occ_index.get((kpi.property_name.lower(), kpi.year, kpi.month))
        if occ is None:
            continue

        kpi.total_units = occ.total_units
        kpi.occupied_units = occ.occupied_units
        kpi.physical_occ_pct = occ.physical_occ_pct

        if kpi.physical_occ_pct is not None and kpi.eco_occ_pct is not None:
            kpi.leakage_gap = kpi.physical_occ_pct - kpi.eco_occ_pct

        units = occ.total_units
        if units:
            if kpi.actual_income is not None:
                kpi.income_per_unit = kpi.actual_income / units
            if kpi.actual_expenses is not None:
                kpi.expense_per_unit = kpi.actual_expenses / units
            if kpi.actual_noi is not None:
                kpi.noi_per_unit = kpi.actual_noi / units

    by_prop_month: dict[tuple, list[PropertyPeriodKPIs]] = {}
    for kpi in kpis:
        key = (kpi.property_name.lower(), kpi.month)
        by_prop_month.setdefault(key, []).append(kpi)

    for group in by_prop_month.values():
        group.sort(key=lambda k: k.year)
        for i in range(1, len(group)):
            curr = group[i]
            prev = group[i - 1]
            if curr.year - prev.year != 1:
                # Years are not consecutive; skip to avoid misrepresenting as YoY.
                continue
            curr.yoy_physical_occ_variance = _yoy_delta(curr.physical_occ_pct, prev.physical_occ_pct)
            curr.yoy_eco_occ_variance = _yoy_delta(curr.eco_occ_pct, prev.eco_occ_pct)
            curr.yoy_leakage_gap_change = _yoy_delta(curr.leakage_gap, prev.leakage_gap)

    return kpis
