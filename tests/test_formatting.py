from app.ui.formatting import fmt_currency, fmt_pct


def test_fmt_currency_positive():
    assert fmt_currency(1234567) == "$1,234,567"


def test_fmt_currency_negative():
    assert fmt_currency(-50000) == "($50,000)"


def test_fmt_currency_zero():
    assert fmt_currency(0) == "$0"


def test_fmt_currency_none():
    assert fmt_currency(None) == "—"


def test_fmt_currency_float():
    assert fmt_currency(99999.99) == "$100,000"


def test_fmt_pct_positive():
    assert fmt_pct(0.954) == "95.4%"


def test_fmt_pct_negative():
    assert fmt_pct(-0.032) == "-3.2%"


def test_fmt_pct_zero():
    assert fmt_pct(0) == "0.0%"


def test_fmt_pct_none():
    assert fmt_pct(None) == "—"
