"""Post-generation validation: ZIP integrity, XML parsing, error token scan."""

import os
import zipfile
from app.models import QualityCheck

_ERROR_TOKENS = ["#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A"]
_MAIN_TABS    = {"Dashboard", "Property Quarterly KPIs", "Property Monthly KPIs", "AR Aging"}
_BACKUP_TABS  = {"Raw_Data", "Source_Index", "Assumptions_Mapping",
                 "Budget_vs_Actual", "Account_Detail", "Economic_Occupancy",
                 "Quality_Checks", "AR_Aging_Detail"}


def validate_workbook(path: str, expected_tabs: set[str]) -> list[QualityCheck]:
    checks: list[QualityCheck] = []
    fname = os.path.basename(path)

    # 1. ZIP integrity
    try:
        with zipfile.ZipFile(path) as z:
            bad = z.testzip()
        checks.append(QualityCheck(
            f"{fname}: Valid ZIP package", bad is None,
            detail="" if bad is None else f"Corrupt file: {bad}",
        ))
    except Exception as e:
        checks.append(QualityCheck(f"{fname}: Valid ZIP package", False, detail=str(e)))
        return checks

    # 2. XML parsing + error token scan
    error_tokens_found = []
    try:
        with zipfile.ZipFile(path) as z:
            for name in z.namelist():
                if not name.startswith("xl/worksheets/") or not name.endswith(".xml"):
                    continue
                content = z.read(name).decode("utf-8", errors="replace")
                for token in _ERROR_TOKENS:
                    if token in content:
                        error_tokens_found.append(f"{name}: {token}")
    except Exception as e:
        checks.append(QualityCheck(f"{fname}: XML parsing", False, detail=str(e)))
        return checks

    checks.append(QualityCheck(
        f"{fname}: No formula errors", len(error_tokens_found) == 0,
        detail="; ".join(error_tokens_found) if error_tokens_found else "Clean",
    ))

    # 3. Worksheet tab names
    try:
        import openpyxl
        wb = openpyxl.load_workbook(path, read_only=True)
        actual_tabs = set(wb.sheetnames)
        wb.close()
        missing = expected_tabs - actual_tabs
        checks.append(QualityCheck(
            f"{fname}: Required tabs present", len(missing) == 0,
            detail="" if not missing else f"Missing: {', '.join(sorted(missing))}",
        ))
    except Exception as e:
        checks.append(QualityCheck(f"{fname}: Tab check", False, detail=str(e)))

    return checks


def validate_both_workbooks(main_path: str, backup_path: str) -> list[QualityCheck]:
    checks = validate_workbook(main_path, _MAIN_TABS)
    checks += validate_workbook(backup_path, _BACKUP_TABS)
    return checks
