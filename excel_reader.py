"""The framework's ONLY test-data reader.

One workbook, one sheet per module. Rows are found by scanning to the last
populated row, so adding data never needs a code change.

    from excel_reader import get_test_cases

    for row in get_test_cases("Customer"):
        print(row["TestCaseID"], row["Customer Name"])

Nothing else in the framework opens a workbook: the path lives in EXCEL_FILE
below and only here, the file is opened once per run, and every module reads its
own sheet through get_test_cases(). To point the suite at a different copy of
the workbook, change EXCEL_FILE - that is the single switch.

The framework is READ-ONLY with respect to the workbook. It never creates,
rebuilds or overwrites it. If the file is missing the run stops with a message
asking for it to be placed in the test_data folder.
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any, Dict, List

from openpyxl import load_workbook

log = logging.getLogger("construction_flow")

# --------------------------------------------------------------------------- #
# THE ONE TEST-DATA FILE
#
# Every module of the framework reads from this workbook and no other. It is
# written relative to the project folder, so it is the same file on every
# machine and after a `git clone`:
#
#     ...\GenZOpss_Construction_Flow\test_data\Construction_Flow_Data.xlsx
#
# Sheet per module: Create_Your_Account, Customer, Suppliers, Items, Categories,
# Quotations, Sales_Orders.
# --------------------------------------------------------------------------- #
PROJECT_ROOT = Path(__file__).resolve().parent
EXCEL_FILE = PROJECT_ROOT / "test_data" / "Construction_Flow_Data.xlsx"

# Spellings of the Execute cell that mean "run this row".
YES_VALUES = {"yes", "y", "true", "1"}

_workbook = None                           # opened once per run
_cache: Dict[str, List["Row"]] = {}        # rows per sheet, read once


class Row(dict):
    """One worksheet row. Behaves like a dict, but explains a wrong column name."""

    def __init__(self, values: dict, sheet: str, number: int) -> None:
        super().__init__(values)
        self.sheet = sheet
        self.number = number

    @property
    def test_case_id(self) -> str:
        return str(self.get("TestCaseID", "") or f"{self.sheet}-row{self.number}")

    def __missing__(self, column: str) -> Any:
        raise KeyError(
            f"There is no '{column}' column in sheet '{self.sheet}' of "
            f"{EXCEL_FILE.name}.\nColumns in that sheet: {', '.join(self)}"
        )


def _clean(value: Any) -> str:
    """Excel value -> the string the application expects.

    Excel hands back numbers and dates, so a PIN code typed as a number arrives
    as 562123.0 and a date as datetime(2026, 9, 2, 0, 0).
    """
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if hasattr(value, "strftime"):                       # date / datetime
        return value.strftime("%Y-%m-%d")
    return str(value).strip()


def _open_workbook():
    """Open the one workbook, once, and keep it for the rest of the run."""
    global _workbook
    if _workbook is not None:
        return _workbook

    if not EXCEL_FILE.exists():
        raise FileNotFoundError(
            "\n" + "=" * 70 +
            f"\nConstruction_Flow_Data.xlsx not found. Please place the workbook"
            f"\ninside the test_data folder."
            f"\n\n  Expected : {EXCEL_FILE}"
            f"\n  Folder   : {'exists' if EXCEL_FILE.parent.exists() else 'is missing too'}"
            f"\n\nThe framework only reads this file - it never creates it."
            f"\n(The path is set in one place only - EXCEL_FILE in excel_reader.py.)"
            "\n" + "=" * 70
        )

    log.info("Loading Excel File... (%s)", EXCEL_FILE)
    try:
        _workbook = load_workbook(EXCEL_FILE, data_only=True)
    except Exception as error:                     # corrupt file, open in Excel
        raise RuntimeError(
            f"\n{'=' * 70}"
            f"\nTEST DATA FILE COULD NOT BE OPENED"
            f"\n\n  File   : {EXCEL_FILE}"
            f"\n  Reason : {error}"
            f"\n\nIf the workbook is open in Excel, close it and run again."
            f"\n{'=' * 70}"
        ) from error
    return _workbook


def read_sheet(sheet_name: str) -> List[Row]:
    """Every populated row of a sheet, in order. Blank rows are skipped."""
    if sheet_name in _cache:
        return _cache[sheet_name]

    workbook = _open_workbook()
    if sheet_name not in workbook.sheetnames:
        raise ValueError(
            f"There is no '{sheet_name}' sheet in {EXCEL_FILE}.\n"
            f"Sheets in the workbook: {', '.join(workbook.sheetnames)}"
        )

    sheet = workbook[sheet_name]
    headers = [_clean(cell) for cell in next(sheet.iter_rows(max_row=1,
                                                            values_only=True))]
    headers = [header for header in headers if header]

    rows: List[Row] = []
    # Row 1 is the header, so data starts at row 2. The last row is wherever the
    # data stops - never a hardcoded number.
    for number, values in enumerate(sheet.iter_rows(min_row=2, values_only=True),
                                    start=2):
        cells = [_clean(value) for value in values]
        if not any(cells):
            continue                                     # blank spacer row
        cells += [""] * (len(headers) - len(cells))      # short row -> blanks
        rows.append(Row(dict(zip(headers, cells)), sheet_name, number))

    _cache[sheet_name] = rows
    return rows


def get_test_cases(sheet_name: str) -> List[Row]:
    """The rows of a sheet whose Execute column says YES."""
    log.info("Loading Sheet: %s", sheet_name)
    rows = read_sheet(sheet_name)
    runnable = [row for row in rows
                if str(row.get("Execute", "")).strip().lower() in YES_VALUES]

    skipped = len(rows) - len(runnable)
    log.info("Found %d executable test case%s.%s", len(runnable),
             "" if len(runnable) == 1 else "s",
             f" ({skipped} skipped with Execute=NO)" if skipped else "")
    return runnable
