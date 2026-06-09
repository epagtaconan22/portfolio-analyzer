"""Formatting helpers — replaces Jinja2 currency and pct template filters."""

from typing import Optional


def fmt_currency(value: Optional[float]) -> str:
    """Format a dollar amount: $1,234,567 or ($50,000) for negatives. '—' for None."""
    if value is None:
        return "—"
    try:
        v = float(value)
    except (TypeError, ValueError):
        return "—"
    if v < 0:
        return f"(${abs(v):,.0f})"
    return f"${v:,.0f}"


def fmt_pct(value: Optional[float]) -> str:
    """Format a decimal ratio as a percentage: 0.954 → '95.4%'. '—' for None."""
    if value is None:
        return "—"
    try:
        return f"{float(value):.1%}"
    except (TypeError, ValueError):
        return "—"
