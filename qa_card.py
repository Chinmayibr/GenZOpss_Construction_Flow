"""THE QA CARD - one module's result, the way a manual tester would write it.

    WHAT WAS ENTERED  ->  WHAT WAS CREATED  ->  WHAT WAS CHECKED  ->  RESULT

One module, one card, and the card is short enough to read in a few seconds:

    ================================================
    QUOTATIONS VALIDATION
    ================================================

    WHAT WAS ENTERED
      Item      : OPC Cement 53 Grade
      Quantity  : 185
      Rate      : 445
      Discount %: 6
      Tax %     : 18

    WHAT WAS CHECKED
      Figure                       Expected      Actual  Difference  Result
      Subtotal (before discount)   82,325       82,325           0   PASS
      Tax Amount                   13,929.39    13,929.39        0   PASS
      Quotation Grand Total        91,314.89    91,314.89        0   PASS

    RESULT: PASS

REPORTING ONLY. This module drives nothing, reads no page and works nothing
out. Every figure on every card is read straight off a Check that
calculation_validation.py had already recorded - its expected, its actual, its
difference and its verdict - so a card cannot say anything the validation did
not say. That is why it is a file of its own: a change here can move where a
number is printed and never what the number is.

It is deliberately NOT the "every calculation checked" table. That table has
fourteen columns because a calculation REVIEW needs fourteen columns - the
inputs, the formula, the reason, the source of every figure. It is still in the
report, underneath this, for the reader who needs it. This card is for the
reader who needs to know whether the module passed.

WHAT IS LEFT OUT, ON PURPOSE:
  * the formula, the inputs and the "read from" notes - on the table below
  * a passing check's reason - "it matched" is what PASS already says
  * anything internal: no variable names, no locators, no tracebacks
A check that FAILED or that could not be verified keeps its reason, because
that is the one line a reader actually has to read.

No import of calculation_validation here, and none is wanted: the reports are
taken duck-typed, which is what keeps this file unable to reach into the
engine and what stops the two importing each other.
"""

from __future__ import annotations

from typing import Dict, List, Optional, Sequence, Tuple

import qa_report
import stock_report

PASS = "PASS"
FAIL = "FAIL"
SKIPPED = "SKIPPED"

#: The picture beside each module's name. Decoration with a job: a reader
#: scrolling a report of eleven test cases finds "the stock one" or "the money
#: one" by its icon before they have read a word.
ICONS: Dict[str, str] = {
    "Create Your Account": "\U0001f511",     # key
    "Customer": "\U0001f465",                # people
    "Suppliers": "\U0001f3e2",               # office
    "Items": "\U0001f3f7",                   # tag
    "Categories": "\U0001f5c2",              # dividers
    "Quotations": "\U0001f4b0",              # money bag
    "Sales_Orders": "\U0001f6d2",            # trolley
    "Purchase_Orders": "\U0001f4e6",         # parcel
    "GRN": "\U0001f69a",                     # lorry
    "Purchase_Bills": "\U0001f9fe",          # receipt
    "Material_Indents": "\U0001f4cb",        # clipboard
}
DEFAULT_ICON = "\U0001f4c4"                  # page

#: Which document register entry names the record a module created. The numbers
#: themselves come from stock_report.DOCUMENTS, which the FLOW filled in as the
#: application allotted them - so a card names the record the run really made,
#: and a module whose record this build prints no number for says so.
DOCUMENT_OF: Dict[str, Tuple[str, str]] = {
    "Quotations": ("Quotation", "Quotation"),
    "Sales_Orders": ("Sales Order", "Sales Order"),
    "Purchase_Orders": ("Purchase Order", "PO Number"),
    "GRN": ("GRN", "GRN Number"),
    "Purchase_Bills": ("Purchase Bill", "Bill Number"),
    "Material_Indents": ("Material Indent", "Indent Number"),
}

#: Which module's card carries the badge the record ended the run showing, and
#: what to call it. A stock movement is only trustworthy once the record behind
#: it is APPROVED, so the two records that put the goods into stock print their
#: status beside their number. Read back from stock_report.STATUSES, which the
#: FLOW filled in from the badge verify_status() had already read - so a card
#: says what the screen said, and a module whose badge this run never read
#: prints no status line at all rather than the status the sheet expected.
STATUS_OF: Dict[str, Tuple[str, str]] = {
    "GRN": ("GRN", "GRN Status"),
    "Purchase_Bills": ("Purchase Bill", "Purchase Bill Status"),
}

#: The module name as a reader says it out loud. "Sales_Orders" is a sheet
#: name; nobody calls it that.
SPOKEN: Dict[str, str] = {
    "Sales_Orders": "SALES ORDER",
    "Purchase_Orders": "PURCHASE ORDER",
    "Purchase_Bills": "PURCHASE BILL",
    "Material_Indents": "MATERIAL INDENT",
    "Quotations": "QUOTATION",
    "Customer": "CUSTOMER",
    "Suppliers": "SUPPLIER",
    "Items": "ITEM",
    "Categories": "CATEGORY",
    "GRN": "GRN",
}


def figure(value: Optional[float]) -> str:
    """A number the way a person writes it: 234, 82,325, 13,929.39.

    Two decimals only when there are two decimals. The rest of the framework
    prints money as "82,325.00" because a column of money is read by lining the
    decimal points up - but this card is a list of answers, not a ledger, and
    "234.00" for two hundred and thirty-four bags invites a reader to look for
    a currency that is not there.
    """
    if value is None:
        return "-"
    try:
        number = float(value)
    except (TypeError, ValueError):
        return str(value)
    if abs(number - round(number)) < 0.0000001:
        return f"{int(round(number)):,}"
    return f"{number:,.2f}"


def mark(status: object) -> str:
    """One verdict with the tick or cross a reader's eye lands on first.

    The word stays beside the symbol - a result column of bare glyphs cannot be
    pasted into a defect, read aloud, or searched for the word FAIL.
    """
    text = str(status or "").strip().upper()
    if text == PASS:
        return "✅ PASS"
    if text == FAIL:
        return "❌ FAIL"
    return f"⚠ {text or 'NOT CHECKED'}"


def title(module: str) -> str:
    """"QUOTATION VALIDATION", with its icon in front of it."""
    icon = ICONS.get(module, DEFAULT_ICON)
    spoken = SPOKEN.get(module, module.replace("_", " ").upper())
    return f"{icon} {spoken} VALIDATION"


# --------------------------------------------------------------------------- #
# WHAT WAS ENTERED, AND WHAT WAS CREATED
# --------------------------------------------------------------------------- #

#: The subject fields, in the order a tester reads them off a form. They are
#: read from the report's context, which subject_of() took off the workbook row
#: and calc.subject() added to - both of which existed long before this card.
SUBJECT_ORDER = ("Item", "Item Code", "Quantity", "Rate", "Discount %", "Tax %")


def entered(report) -> List[List[str]]:
    """WHAT WAS ENTERED - the values this row typed in, and nothing else."""
    context = getattr(report, "context", {}) or {}
    pairs = [[label, str(context.get(label, "")).strip()]
             for label in SUBJECT_ORDER
             if str(context.get(label, "")).strip()]
    return pairs


def created(report) -> List[List[str]]:
    """WHAT WAS CREATED - the record the application allotted a number to."""
    module = getattr(report, "module", "")
    known = DOCUMENT_OF.get(module)
    if known is None:
        return []
    kind, label = known
    ref = stock_report.document(kind)
    pairs = [[label, ref]] if ref else []

    # THE BADGE, under the number it belongs to. Absent when this run never
    # read one - a blank status line, or one filled in from the sheet's
    # expected value, would both read as "the application showed APPROVED".
    badge = STATUS_OF.get(module)
    if badge is not None:
        kind_of_badge, badge_label = badge
        showing = stock_report.status(kind_of_badge)
        if showing:
            pairs.append([badge_label, showing.upper()])

    # THE OPENING STOCK, on the card of the module that read it. It is the
    # figure the whole stock story is measured from, and the Purchase Orders
    # module is where the run takes it - so "what did this item start with?"
    # is answered on the card rather than three attachments away. Read back
    # from the register the flow filled in; absent when it was never read.
    if module == "Purchase_Orders":
        opening = stock_report.quantity("Opening Stock")
        if opening is not None:
            pairs.append(["Stock Before This Purchase", figure(opening)])
    return pairs


# --------------------------------------------------------------------------- #
# WHAT WAS CHECKED - expected, actual, difference, verdict. Four columns.
# --------------------------------------------------------------------------- #

CHECK_HEADINGS = ("Figure", "Expected", "Actual", "Difference", "Result")


def checked_rows(report) -> List[List[str]]:
    """One line per calculation: what it was, both figures, the verdict.

    The inputs, the formula, the tolerance and the source of each figure are
    NOT here. They are on the "Calculation Validation" table underneath, which
    is what a calculation review reads - and putting them here as well is what
    made the report a wall of text a QA tester had to mine for the answer.
    """
    rows: List[List[str]] = []
    for check in getattr(report, "checks", []) or []:
        rows.append([
            str(getattr(check, "name", "")),
            figure(getattr(check, "expected", None)),
            figure(getattr(check, "actual", None)),
            figure(getattr(check, "difference", None)),
            str(getattr(check, "status", "")),
        ])
    return rows


def unresolved(report) -> List[List[str]]:
    """The checks that FAILED or could not be verified, with their reason.

    Only these carry a reason on this card. "It matched" is everything a
    passing check's reason ever says, and printing it once per row is how a
    short card becomes a long one.
    """
    lines: List[List[str]] = []
    for check in getattr(report, "checks", []) or []:
        status = str(getattr(check, "status", ""))
        if status == PASS:
            continue
        reason = str(getattr(check, "note", "") or "").strip()
        lines.append([f"{getattr(check, 'name', '')} ({status})",
                      reason or "(no reason recorded)"])
    return lines


def verdict(report) -> str:
    """The module's own verdict, as the engine already worked it out."""
    return str(getattr(report, "overall", "") or
               getattr(report, "calculation", "") or "")


# --------------------------------------------------------------------------- #
# THE CARD
# --------------------------------------------------------------------------- #

def card_text(report) -> str:
    """The card as the plain text a QA person pastes into a ticket."""
    module = getattr(report, "module", "")
    rows = checked_rows(report)
    reasons = unresolved(report)

    parts = [_banner(title(module))]

    identity = [["Test Case", str(getattr(report, "test_case_id", "") or "")]]
    identity = [pair for pair in identity if pair[1]]
    if identity:
        parts.append(qa_report.labelled(identity))

    for heading, pairs in (("WHAT WAS ENTERED", entered(report)),
                           ("WHAT WAS CREATED", created(report))):
        if pairs:
            parts.append(qa_report.section(heading,
                                           qa_report.labelled(pairs, "  ")))

    if rows:
        parts.append(qa_report.section(
            "WHAT WAS CHECKED",
            _plain_table(CHECK_HEADINGS, rows)))

    if reasons:
        parts.append(qa_report.section(
            "WHAT COULD NOT BE CONFIRMED",
            qa_report.labelled(reasons, "  ")))

    parts.append(f"RESULT: {mark(verdict(report))}")
    return qa_report.joined(parts)


def card_page(report) -> str:
    """The same card as the colour page a stakeholder opens."""
    module = getattr(report, "module", "")
    result = verdict(report)
    rows = checked_rows(report)
    reasons = unresolved(report)

    passed = len([row for row in rows if row[4] == PASS])
    failed = len([row for row in rows if row[4] == FAIL])

    blocks = [qa_report.kpi_block("Result", [
        ["Figures checked", len(rows)],
        ["Matched", passed, PASS if passed else ""],
        ["Did not match", failed, FAIL if failed else ""],
        ["Result", result, result],
    ])]
    for heading, pairs in (("What was entered", entered(report)),
                           ("What was created", created(report))):
        if pairs:
            blocks.append(qa_report.facts_block(heading, pairs))
    if rows:
        blocks.append(qa_report.table_block(
            "What was checked", CHECK_HEADINGS, rows,
            note="Expected is a value worked out independently from the same "
                 "inputs; Actual is what the application displayed."))
    if reasons:
        blocks.append(qa_report.facts_block(
            "What could not be confirmed", reasons,
            note="These were NOT counted as passes. Nothing was assumed "
                 "about them."))
    return qa_report.page(title(module).split(" ", 1)[-1].title(),
                          str(getattr(report, "test_case_id", "") or ""),
                          blocks, verdict=result)


def _banner(heading: str) -> str:
    """The card's heading with a rule under it that reaches the end of it.

    qa_report.banner() lays the rule out on len(heading), which is right for
    every other banner in the framework and one character short here: an emoji
    is one character and two columns wide, so the rule stopped short of the
    title it was underlining.
    """
    width = len(heading) + sum(1 for glyph in heading if ord(glyph) > 0x2100)
    return heading + "\n" + "=" * width


def _plain_table(headings: Sequence[str], rows: Sequence[Sequence[str]]) -> str:
    """A small fixed-width table, laid out on its own widest cell.

    text_table() in data_verification.py does this for the big tables and adds
    a rule, a legend and a wrapped Reason column that this card does not want -
    so the four columns are laid out here, where the widths can stay narrow
    enough for the card to be read at a glance.
    """
    columns = len(headings)
    widths = [len(str(head)) for head in headings]
    for row in rows:
        for index in range(columns):
            cell = str(row[index]) if index < len(row) else ""
            widths[index] = max(widths[index], len(cell))

    def line(cells: Sequence[object]) -> str:
        out = []
        for index in range(columns):
            cell = str(cells[index]) if index < len(cells) else ""
            # The name reads left to right; every figure reads right to left.
            out.append(f"{cell:<{widths[index]}}" if index == 0
                       else f"{cell:>{widths[index]}}")
        return "  ".join(out).rstrip()

    return "\n".join(["  " + line(headings),
                      "  " + "-" * (sum(widths) + 2 * (columns - 1))]
                     + ["  " + line(row) for row in rows])
