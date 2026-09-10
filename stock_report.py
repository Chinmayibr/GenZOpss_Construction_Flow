"""ITEMS & INVENTORY - STOCK VALIDATION, in the shape a QA reader asks for it.

    BEFORE  ->  TRANSACTION  ->  AFTER  ->  EXPECTED vs ACTUAL  ->  RESULT

One stock movement, one card, and the card answers the seven questions a QA
person actually has:

    1. what was the stock BEFORE the transaction?
    2. what transaction happened?
    3. how much was added or removed?
    4. what SHOULD the stock be afterwards?
    5. what does the application actually show?
    6. is the difference 0?
    7. PASS or FAIL?

REPORTING ONLY. Nothing here reads a page, drives a flow or works a verdict
out. Every figure on every card is a reading the run already took or a check
calculation_validation.py already recorded and this module read back - so a
card can never say anything the validation did not say. That is the whole
reason it is a separate file: there is no way for a change here to move a
number, only to change how it is displayed.

The cards are also counted (see counts()), which is what lets the final QA
summary report "Stock Validations: N | Passed: N | Failed: N" from the same
records the cards are drawn from, rather than from a tally kept by hand.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Sequence, Tuple

import qa_report

#: Where the ACTUAL figure on every one of these cards is read from. Named once
#: so the cards, the formulas and the summary cannot drift into calling the
#: same column three different things. Named for what the APPLICATION calls
#: it - the item's own page prints a field headed exactly "Current Stock",
#: and that is the wording and the figure a person compares this report
#: against, so the report has to call it the same thing.
STOCK_COLUMN = "Current Stock"

#: WHERE the Current Stock figure comes from, for the "read from ..."
#: sentences - the item's OWN page, not the Items & Inventory list (see
#: Construction_Flow.stock_now()). Kept apart from STOCK_COLUMN, which is the
#: field's own name: a sentence combining both would read "Current Stock is
#: read from Current Stock," which says nothing.
STOCK_SOURCE = "the item's own page (Current Stock)"

PASS = "PASS"
FAIL = "FAIL"
SKIPPED = "SKIPPED"
NOT_CHECKED = "NOT CHECKED"

#: The unit this run's stock quantities were printed in - "PCS", "KG" - read
#: off the application (Construction_Flow.read_current_stock()) and never
#: guessed: a run whose reading never carried a unit prints none, rather than
#: assuming every item is counted in "PCS".
UNIT: str = ""


def note_unit(unit: str) -> None:
    """Record the unit the application printed alongside a stock quantity.

    One unit for the whole report: every figure in it is the SAME item's
    stock, read at different moments, so there is one unit to record - not a
    register keyed per reading like DOCUMENTS/QUANTITIES/STATUSES. A blank
    reading is ignored, so a screen that failed to show a unit cannot blank
    out one an earlier, working reading already established.
    """
    global UNIT
    text = str(unit or "").strip()
    if text:
        UNIT = text.upper()


def num(value: Optional[float]) -> str:
    """A quantity as the report prints it - "300 PCS", "457.5", "-" for nothing.

    Quantities, not money: a stock figure is a count of things, and printing
    "300.00" for three hundred bags invites a reader to look for a currency
    that is not there. The unit is appended only once this run has actually
    captured one (see note_unit()) - never assumed.
    """
    if value is None:
        return "-"
    try:
        text = f"{float(value):g}"
    except (TypeError, ValueError):
        return str(value)
    return f"{text} {UNIT}" if UNIT else text


def money(value: Optional[float]) -> str:
    """A money figure, printed the way the rest of this report prints one.

    The same format as calculation_validation.money(), written out here rather
    than imported so that this module keeps depending on nothing but the
    renderer. A before-and-after card for an AMOUNT - a supplier's outstanding,
    say - is the same card as one for a quantity; only the formatting differs.
    """
    if value is None:
        return "-"
    try:
        return f"{float(value):,.2f}"
    except (TypeError, ValueError):
        return str(value)


def text_table(headings: Sequence[str],
               rows: Sequence[Sequence[object]]) -> str:
    """A plain-text table whose columns are sized by what is actually in them.

    REPORTING ONLY. The stock tables used to be laid out on hard-coded column
    widths, so an item name longer than the width it was given ran into the
    figure beside it and a reader had to count characters to tell which column
    a number was in. The widths are measured here instead.
    """
    body = [[str(cell) for cell in row] for row in rows]
    width = [max(len(str(headings[index])),
                 *(len(row[index]) if index < len(row) else 0
                   for row in body)) if body else len(str(headings[index]))
             for index in range(len(headings))]
    def line(cells: Sequence[str]) -> str:
        return "  ".join(f"{str(cell):<{width[index]}}"
                         for index, cell in enumerate(cells)).rstrip()
    return "\n".join([line(headings),
                      "-" * (sum(width) + 2 * (len(width) - 1))]
                     + [line(row) for row in body])


def _same(left: Optional[float], right: Optional[float],
          tolerance: float = 0.001) -> bool:
    """Are these two readings the same figure?"""
    if left is None or right is None:
        return False
    return abs(float(left) - float(right)) <= tolerance


@dataclass
class Movement:
    """One stock movement, before and after, with its verdict.

    `quantities` is what the transaction was made of - the GRN's received
    quantity, or the indent's requested / approved / issued / returned - as
    (label, value) pairs in the order a reader reads them. They are kept as
    separate lines and never added together here: REQUESTED, APPROVED, ISSUED
    and RETURNED are four different numbers and the report says so.
    """

    title: str                                   # "GRN STOCK VALIDATION"
    item: str = ""                               # "OPC Cement 53 Grade (CODE)"
    document: str = ""                           # "GRN-0007" / "IND-0012"
    #: The short name for the run-wide table's first column. Written out rather
    #: than derived from the title, because title-casing "GRN STOCK VALIDATION"
    #: produces "Grn" - a report that cannot spell the module it is reporting on
    #: is a report a reader stops trusting the numbers of.
    stage: str = ""
    before_label: str = "Stock Before"
    before: Optional[float] = None
    quantities: List[List[object]] = field(default_factory=list)
    #: The terms that actually MOVE the stock, as (sign, label, value). Kept
    #: separate from `quantities` on purpose, and this is the most important
    #: line in the file: REQUESTED and APPROVED are shown to a reader because
    #: they are part of the story, but they move nothing. A sum that added them
    #: in would read "500 + 43 + 43 - 43 = 462" - arithmetic that does not work
    #: and that quietly claims the request and the approval each changed the
    #: stock. Only what the store really moved belongs in the sum.
    terms: List[List[object]] = field(default_factory=list)
    formula: str = ""                            # "300 + 200"
    expected: Optional[float] = None
    actual: Optional[float] = None
    difference: Optional[float] = None
    result: str = NOT_CHECKED
    note: str = ""
    #: Print this card's figures as MONEY rather than as quantities. A
    #: supplier's outstanding moves before-and-after exactly as stock does, and
    #: it is the same card - but "90000" where a reader expects "90,000.00"
    #: reads as a quantity, so the formatter is chosen per card.
    amount: bool = False
    #: What the "Expected ..." line and the bare actual-value line are about.
    #: "Current Stock" for a stock card - the application's OWN field name,
    #: on the item's own page - "Outstanding" for a money one, because a card
    #: headed SUPPLIER OUTSTANDING VALIDATION that then says "Expected Stock"
    #: is a card a reader has to correct in their head before they can use it.
    subject: str = "Current Stock"
    #: WHERE the Actual figure was read from, printed beside it. Defaulted to
    #: the stock column because that is what most of these cards are about, and
    #: settable because not all of them are: a purchase-order COUNT read off
    #: the Purchase Orders list that told a reader it came from Items &
    #: Inventory would be citing the wrong screen for the right number, which
    #: is the one kind of error a validation report cannot afford.
    source: str = STOCK_SOURCE

    @property
    def show(self):
        """The formatter this card's figures are printed with."""
        return money if self.amount else num

    # -- the one thing this module works out, and it changes no verdict ----- #
    @property
    def moved(self) -> Optional[bool]:
        """Did the stock actually change? None when one side was not read.

        DISPLAY ONLY - it is never compared, never counted and never allowed
        near a status. It exists because "before" and "after" showing the same
        figure is the one thing a stock report must not let a reader skim past:
        a transaction that moved nothing looks exactly like a passing check
        when only the closing figure is printed.
        """
        if self.before is None or self.actual is None:
            return None
        return not _same(self.before, self.actual)

    @property
    def transaction_note(self) -> str:
        """The warning line for a movement where nothing moved."""
        if self.moved is not False:
            return ""
        if not any(_significant(value) for _, value in self.quantities):
            return ""
        return (f"The {self.subject.lower()} is the SAME before and after this "
                f"transaction. Something was moved, so it should not be - read "
                f"the figures below before treating this as a pass.")

    # -- the plain-text card ----------------------------------------------- #
    def text(self) -> str:
        """The card as the plain text a QA person pastes into a ticket."""
        before = [[self.before_label, self.show(self.before)]]
        transaction = [[label, self.show(value) if not isinstance(value, str)
                        else value] for label, value in self.quantities]
        after: List[List[object]] = [
            [f"Expected {self.subject}", self.show(self.expected)
             + (f"   ({self.formula})" if self.formula else "")],
            [self.subject, self.show(self.actual)
             + (f"   ({self.source})" if self.source else "")],
            ["Difference", self.show(self.difference)],
        ]

        parts = [qa_report.banner(self.title, "-")]
        if self.document:
            parts.append(qa_report.labelled([["Document", self.document]]))
        parts += [
            qa_report.section("BEFORE TRANSACTION",
                              qa_report.labelled(before, "  ")),
            qa_report.section("TRANSACTION",
                              qa_report.labelled(transaction, "  ")),
            qa_report.section("AFTER TRANSACTION",
                              qa_report.labelled(after, "  ")),
            f"Result: {self.result}",
        ]
        for extra in (self.transaction_note, self.note):
            if extra:
                parts.append(extra)
        return qa_report.joined(parts)

    # -- the same card as HTML blocks -------------------------------------- #
    def blocks(self) -> List[str]:
        """The card as the blocks of an HTML page."""
        transaction = [[label, self.show(value) if not isinstance(value, str)
                        else value] for label, value in self.quantities]
        return [
            qa_report.kpi_block(self.title, [
                [self.before_label, self.show(self.before)],
                [f"Expected {self.subject}", self.show(self.expected)],
                [self.subject, self.show(self.actual)],
                ["Difference", self.show(self.difference)],
                ["Result", self.result, self.result],
            ]),
            qa_report.facts_block("Transaction", (
                ([["Document", self.document]] if self.document else [])
                + transaction), note=self.transaction_note),
            qa_report.pre_block("How the expected figure was worked out",
                                self.sum_lines()),
            qa_report.note_block("Observation", self.note),
        ]

    def sum_lines(self) -> str:
        """The arithmetic, laid out so a reader can add it up by hand.

        ONLY the terms that move stock - see `terms`. The quantities that do
        not move any (requested, approved) are on the card above, where they
        belong, and not in a sum they would make untrue.
        """
        # Laid out on the longest label this card actually carries, not on a
        # fixed width: "Outstanding Before This Run's Purchase" is longer than
        # any stock label, and a fixed column ran the figure straight into it.
        show = self.show
        labels = ([self.before_label] + [str(label) for _, label, _ in self.terms]
                  + [f"Expected {self.subject}", self.subject,
                     "Difference", "Result"])
        width = max(36, max(len(label) for label in labels) + 2)
        lines = [f"    {self.before_label:<{width}}{show(self.before):>10}"]
        for sign, label, value in self.terms:
            if value is None:
                continue
            lines.append(f"  {sign} {str(label):<{width}}{show(value):>10}")
        lines += [
            f"  {'-' * (width + 12)}",
            f"  = {'Expected ' + self.subject:<{width}}"
            f"{show(self.expected):>10}",
            f"    {self.subject:<{width}}{show(self.actual):>10}",
            f"    {'Difference':<{width}}{show(self.difference):>10}",
            f"    {'Result':<{width}}{self.result:>10}",
        ]
        if self.source:
            lines += ["", f"    {self.subject} is read from {self.source}."]
        return "\n".join(lines)

    def summary_row(self) -> List[str]:
        """One line of the run-wide "every stock validation" table."""
        show = self.show
        return [self.stage or self.title.replace(" STOCK VALIDATION", ""),
                self.item, self.document or "-", show(self.before),
                "; ".join(f"{label} {show(value)}"
                          for label, value in self.quantities
                          if _significant(value)) or "-",
                show(self.expected), show(self.actual), show(self.difference),
                self.result]


def _significant(value: object) -> bool:
    """Is this quantity worth printing as part of the sum?"""
    try:
        return value is not None and float(value) != 0.0
    except (TypeError, ValueError):
        return bool(str(value or "").strip())


#: Every stock movement this run recorded, in the order it happened. Read by
#: the combined page and by the final QA summary.
MOVEMENTS: List[Movement] = []


def record(movement: Movement) -> Movement:
    """Keep one movement for the combined page and the summary counts."""
    MOVEMENTS.append(movement)
    return movement


def reset() -> None:
    """Start a run with no movements - used by the tests of this framework."""
    global UNIT
    MOVEMENTS.clear()
    DOCUMENTS.clear()
    QUANTITIES.clear()
    STATUSES.clear()
    UNIT = ""


# --------------------------------------------------------------------------- #
# THE DOCUMENTS BEHIND THE MOVEMENTS
#
# "The stock went from 0 to 234" is only half an answer; the other half is
# WHICH purchase order, WHICH goods receipt and WHICH bill did it. Every one of
# these numbers is allotted by the application and read off the screen by the
# flow that created the record - so they are recorded here as they are read,
# and never derived, guessed or built out of a pattern.
#
# A document this run did not capture is simply absent. It is never filled in
# with a plausible-looking number, and the report says "(not captured by this
# run)" rather than inventing one.
# --------------------------------------------------------------------------- #

#: (kind, reference, what the reference is) in the order the run created them.
DOCUMENTS: List[List[str]] = []


def note_document(kind: str, reference: str, detail: str = "") -> None:
    """Record one document reference, exactly as the application allotted it.

    Ignores an empty reference: a blank line under "Purchase Bill" tells a
    reader the bill has no number, which is a claim about the application. The
    honest report is that this run did not capture one.
    """
    text = str(reference or "").strip()
    if not text:
        return
    for entry in DOCUMENTS:
        if entry[0] == kind and entry[1] == text:
            return                                 # already recorded
    DOCUMENTS.append([kind, text, detail])


def documents() -> List[List[str]]:
    """The document register, for the report's own table."""
    return [list(entry) for entry in DOCUMENTS]


# --------------------------------------------------------------------------- #
# THE STOCK MOVEMENT SUMMARY - the whole run as a running balance
#
#     Transaction              Quantity     Stock
#     ---------------------------------------------
#     Initial Stock                 0          0
#     GRN Approved               +234        234
#     Material Indent Issued       -7        227
#
# The one table that answers "what happened to this item?" in the order it
# happened. The Stock column is a READING, not a running total worked out here:
# it is filled in only on the line whose reading the run actually took, and
# left blank on a line it did not. A balance carried forward by arithmetic
# would agree with itself all the way down a column that the application never
# said anything about.
# --------------------------------------------------------------------------- #

def ledger_rows(movements: Optional[Sequence[Movement]] = None
                ) -> List[List[str]]:
    """Transaction | Quantity | Stock, in the order the stock moved."""
    # The Item Created card is the opening READING, not a movement: it is the
    # "Initial Stock" line below, and printing it twice would put a
    # transaction in this ledger that never happened.
    cards = [card for card in (MOVEMENTS if movements is None else movements)
             if card.stage != ITEM_STAGE]
    if not cards:
        return []

    # "Initial Stock  0  0" - the quantity column is 0 because nothing moved
    # to get there, not because nothing is known. When the opening reading was
    # never taken there is no stock to report either, and both columns say so.
    opening = opening_stock()
    rows: List[List[str]] = [
        ["Initial Stock", "0" if opening is not None else "-", num(opening)],
    ]
    for card in cards:
        moving = [term for term in card.terms if term[2] is not None]
        for index, (sign, label, value) in enumerate(moving):
            last = index == len(moving) - 1
            rows.append([
                str(label),
                f"{sign}{num(value)}",
                # The reading belongs to the END of the movement. A card with
                # two terms - issued and returned - took ONE reading after
                # both, so the first of them leaves the Stock column empty
                # rather than showing a figure nobody read.
                num(card.actual) if last else "",
            ])
        if not moving:
            rows.append([card.title.replace(" VALIDATION", "").title(),
                         "-", num(card.actual)])
    return rows


def movement_summary_text(item: str = "") -> str:
    """STOCK MOVEMENT SUMMARY, as the plain text a QA person reads first."""
    cards = list(MOVEMENTS)
    if not cards:
        return ("STOCK MOVEMENT SUMMARY\n\n"
                "This run moved no stock, so there is no movement to "
                "summarise - and nothing was assumed.")

    name = item or " | ".join(sorted({card.item for card in cards if card.item}))
    final = last_movement() or cards[-1]

    return qa_report.joined([
        qa_report.banner("STOCK MOVEMENT SUMMARY"),
        qa_report.labelled([["Item", name or "(not named)"]]),
        text_table(("Transaction", "Quantity", "Stock"), ledger_rows()),
        qa_report.labelled([
            ["Current Stock", num(final.actual)],
            ["Expected Current Stock", num(final.expected)],
            ["Difference", num(final.difference)],
            ["Overall Result", result()],
        ]),
        (qa_report.section(
            "DOCUMENTS BEHIND THESE MOVEMENTS",
            qa_report.labelled(
                [[kind, f"{reference}   "
                        f"[{status(kind) or 'no status read'}]"
                        + (f"   ({detail})" if detail else "")]
                 for kind, reference, detail in DOCUMENTS]))
         if DOCUMENTS else ""),
    ])


def movement_summary_blocks(item: str = "") -> List[str]:
    """The same summary as the blocks of an HTML page."""
    cards = list(MOVEMENTS)
    if not cards:
        return [qa_report.note_block(
            "Stock movement summary",
            "This run moved no stock, so there is no movement to summarise.")]

    final = last_movement() or cards[-1]
    return [
        qa_report.kpi_block("Stock movement", [
            ["Initial Stock", num(opening_stock())],
            ["Current Stock", num(final.actual)],
            ["Expected Current Stock", num(final.expected)],
            ["Difference", num(final.difference)],
            ["Overall Result", result(), result()],
        ]),
        qa_report.table_block(
            "The stock, transaction by transaction",
            ("Transaction", "Quantity", "Stock"), ledger_rows(),
            note="The Stock column is what Items & Inventory showed after that "
                 "transaction - a reading, not a running total worked out "
                 "here. A blank means no reading was taken at that point.",
            verdict_column=99, legend=False),
        (qa_report.table_block(
            "The documents behind these movements",
            ("Document", "Reference", "Status", "What it is"),
            documents_with_status(), note=DOCUMENT_STATUS_NOTE,
            verdict_column=99, legend=False) if DOCUMENTS else ""),
    ]


def counts() -> Dict[str, int]:
    """total / passed / failed / not verified, over the movements recorded."""
    results = [movement.result for movement in MOVEMENTS]
    return {
        "total": len(results),
        "passed": results.count(PASS),
        "failed": results.count(FAIL),
        "skipped": len([result for result in results
                        if result not in (PASS, FAIL)]),
    }


def result() -> str:
    """The stock verdict over the whole run."""
    tally = counts()
    if tally["failed"]:
        return FAIL
    if tally["passed"]:
        return PASS
    return SKIPPED if tally["total"] else "NOT APPLICABLE"


# =========================================================================== #
# THE QA-FRIENDLY STOCK VALIDATION SECTIONS
#
#     [STOCK] Stock Validation
#        |-- Stock Increase - GRN
#        |-- Stock Decrease - Material Indent
#        +-- Stock Movement Summary
#
# Everything below this line is PRESENTATION. It reads the Movement cards the
# run already recorded, the document register note_document() already filled
# in and the quantity register note_quantity() already filled in, and lays them
# out as the question-and-answer sheet a manual tester reads without opening a
# line of Python:
#
#     What item?  What was the stock before?  How much moved?  What SHOULD it
#     be now?  What does the application actually show?  Is the difference 0?
#
# Not one figure here is worked out, compared or given a verdict. Expected,
# Actual, Difference and Result are read straight off the Movement, which read
# them straight off the check calculation_validation.py recorded. A change in
# this section can move where a number is printed; it cannot move the number.
# =========================================================================== #

#: The `stage` names the stock cards are recorded under. Named once so the
#: sections, the chain and the summary all find the same cards.
GRN_STAGE = "Stock Increase (GRN)"
INDENT_STAGE = "Stock Decrease (Indent)"

#: The stage the reading taken right after Quotation + Sales Order is recorded
#: under - BEFORE the Material Indent (which is what actually issues stock in
#: this application) has done anything. It sits between GRN_STAGE and
#: INDENT_STAGE in the chain: neither a Quotation nor a Sales Order moves
#: Items & Inventory on its own, so this card's job is to PROVE the stock is
#: still what the purchase flow left it at, with a real reading, rather than
#: to assume it.
SALES_STAGE = "Stock After Sales (Quotation + Sales Order)"

#: The stage the NEW ITEM's opening reading is recorded under. It is a stock
#: card like the other two - "what did Items & Inventory show before anything
#: moved?" - and it is the figure every later stock verdict is measured FROM,
#: which is why it gets a card of its own rather than being a footnote on the
#: GRN one.
ITEM_STAGE = "Item Created (Initial Stock)"

#: What is printed where a figure was never captured. It is deliberately a
#: WORD and not a dash: a blank invites a reader to assume zero, and zero is a
#: claim about the application that nobody made.
NOT_CAPTURED = "NOT AVAILABLE"


def mark(result_text: object) -> str:
    """One verdict with the tick or cross a reader's eye lands on first.

    The word is kept beside the symbol, never replaced by it - a report whose
    result column is a single glyph cannot be pasted into a ticket, read by a
    screen reader, or searched for the word FAIL.
    """
    text = str(result_text or NOT_CHECKED).strip().upper()
    if text == PASS:
        return "✅ PASS"
    if text == FAIL:
        return "❌ FAIL"
    return f"⚠ {text}"


# --------------------------------------------------------------------------- #
# THE QUANTITY REGISTER - the quantities each document was raised for
#
# The document register (DOCUMENTS) answers "WHICH purchase order?"; this one
# answers "for HOW MUCH?". They are kept apart for the same reason REQUESTED
# and ISSUED are: a quotation's quantity and a goods receipt's quantity are two
# different claims about two different records, and the moment they share a
# slot the report can no longer say which screen a figure came off.
#
# Same rule as the documents: a quantity this run did not capture is ABSENT.
# It is never filled in from the quantity of a neighbouring document, however
# obviously equal the two look.
# --------------------------------------------------------------------------- #

#: (label, value, where it was read) in the order the run captured them.
QUANTITIES: List[List[object]] = []


def note_quantity(label: str, value: Optional[float], detail: str = "") -> None:
    """Record one document's quantity, exactly as the run captured it."""
    if value is None or str(value).strip() == "":
        return
    try:
        number = float(value)
    except (TypeError, ValueError):
        return
    for entry in QUANTITIES:
        if entry[0] == label:
            entry[1] = number
            entry[2] = detail or entry[2]
            return
    QUANTITIES.append([label, number, detail])


def quantity(label: str) -> Optional[float]:
    """One captured quantity, or None when this run never captured it."""
    for name, value, _ in QUANTITIES:
        if name == label:
            return float(value)
    return None


# --------------------------------------------------------------------------- #
# THE STATUS REGISTER - the badge each document was showing when it was read
#
# DOCUMENTS answers "WHICH goods receipt?", QUANTITIES answers "for HOW MUCH?",
# and this one answers "and was it APPROVED?". A stock movement is only the
# application's own answer once the record that caused it has been approved, so
# a reader who is shown "+234 received" without "APPROVED" beside it is being
# asked to take the approval on trust.
#
# REPORTING ONLY, and read-only: every value here was already read off the
# screen by verify_status() in the flow, which reads the badge to ASSERT it.
# Nothing here reads a page, and nothing here decides anything - a status this
# run never read is ABSENT, and is never filled in from the sheet's expected
# value however obviously right that looks.
# --------------------------------------------------------------------------- #

#: (kind, status, where it was read) in the order the run captured them.
STATUSES: List[List[str]] = []


def note_status(kind: str, status: str, detail: str = "") -> None:
    """Record one document's status, exactly as the badge on screen spelt it.

    A later reading REPLACES an earlier one, because a record is read twice -
    before its approval and after it - and the status a card should carry is
    the one the record ended the run in. An empty reading is ignored: a blank
    beside "GRN Status" would be a claim that the application showed no badge,
    when all it means is that this run did not capture one.
    """
    text = str(status or "").strip()
    if not text:
        return
    for entry in STATUSES:
        if entry[0].casefold() == str(kind).casefold():
            entry[1] = text
            entry[2] = detail or entry[2]
            return
    STATUSES.append([str(kind), text, detail])


def status(kind: str) -> str:
    """One captured status, or "" when this run captured none for it."""
    for name, text, _ in STATUSES:
        if name.casefold() == str(kind).casefold():
            return text
    return ""


def status_shown(kind: str) -> str:
    """A status as the cards print it - never blank, never invented."""
    return status(kind) or NOT_CAPTURED


def statuses() -> List[List[str]]:
    """The status register, for a report that wants the whole of it."""
    return [list(entry) for entry in STATUSES]


def documents_with_status() -> List[List[str]]:
    """The document register with each record's badge beside its number.

        Document        Reference   Status                      What it is
        Purchase Order  PO-00123    Approved                    the order...
        GRN             GRN/0045    Approved                    the receipt...
        Purchase Bill   PB-00088    Approved                    the bill...
        Quotation       QT-00011    (not captured by this run)  the quote...

    A JOIN of two registers the flow filled in separately, and nothing more:
    the reference off DOCUMENTS, the badge off STATUSES, matched on the kind
    of document. Neither register is touched, and a kind this run never read a
    badge for says so - the Status column is never filled in from a
    neighbouring document's badge, and never from the sheet's Expected Status.
    """
    return [[kind, reference, status_shown(kind), detail]
            for kind, reference, detail in DOCUMENTS]


#: What the Status column means, said once under every table that prints it.
DOCUMENT_STATUS_NOTE = (
    "Status is the badge the record itself was showing when this run last "
    "read it - read off the screen, not taken from the sheet. Only the "
    "records whose status this run reads carry one; the rest say so rather "
    "than borrowing a neighbouring document's badge."
)


def document(kind: str) -> str:
    """One captured document reference, or "" when the run captured none."""
    for name, ref, _ in DOCUMENTS:
        if name.casefold() == kind.casefold():
            return ref
    return ""


def reference(kind: str) -> str:
    """A document reference as the report prints it, never blank."""
    return document(kind) or NOT_CAPTURED


def movement_by_stage(stage: str) -> Optional[Movement]:
    """The recorded card for one stage of the chain, if the run got that far."""
    for card in MOVEMENTS:
        if card.stage == stage:
            return card
    return None


def shown(value: Optional[float]) -> str:
    """A quantity as the sections print it - never blank, never invented."""
    return NOT_CAPTURED if value is None else num(value)


def item_name() -> str:
    """The item this run's stock chain is about, as the cards recorded it."""
    names = [card.item for card in MOVEMENTS if card.item and not card.amount]
    return names[0] if names else "(not named)"


#: "OPC Cement 53 Grade (ITEM-OPC53-034)" -> the name, the code between the
#: LAST parentheses. Movement.item is always built as "name (code)" (see
#: Construction_Flow.py's Movement() calls); this is a display-only split so
#: a card can print "Item Name" and "Item ID" as two lines instead of one.
#: The name half must be non-empty - "(not named)" on its own is the WHOLE
#: placeholder, not a nameless item with code "not named".
_ITEM_IDENTITY = re.compile(r"^(\S.*?)\s*\(([^()]*)\)\s*$")


def item_identity(card_item: str) -> Tuple[str, str]:
    """(name, code) split out of a Movement's combined `item` string."""
    match = _ITEM_IDENTITY.match(str(card_item or ""))
    if not match:
        return str(card_item or "(not named)"), ""
    return match.group(1), match.group(2) or ""


def _asked(card: Optional[Movement]) -> Dict[str, object]:
    """A card's quantity lines as a lookup, so a missing one reads as None."""
    if card is None:
        return {}
    return {str(label): value for label, value in card.quantities}


# --------------------------------------------------------------------------- #
# STOCK INCREASE VALIDATION - GRN
# --------------------------------------------------------------------------- #

INCREASE_RULE = "Current Stock Before Purchase + GRN Approved Quantity"
DECREASE_RULE = "Current Stock Before Consumption - ACTUAL Issued Quantity"
#: Quotation and Sales Order are business documents; neither one moves Items &
#: Inventory in this application (the stock only changes at GRN approval and
#: at a Material Indent's stock issue) - so the rule this checkpoint proves is
#: that the stock is UNCHANGED here, with a real reading, not an assumption.
SALES_RULE = ("Current Stock After Purchase (Quotation and Sales Order do not "
             "move inventory - stock is expected to be UNCHANGED here; it is "
             "reduced later, when the Material Indent issues it)")

_NO_INCREASE = ("This run recorded no GRN stock movement, so there is no stock "
                "increase to show here - and none was assumed.")
_NO_DECREASE = ("This run recorded no Material Indent stock movement, so there "
                "is no stock decrease to show here - and none was assumed.")
_NO_SALES = ("This run recorded no post-Sales-Order stock reading, so there is "
            "no before/after figure to show here - and none was assumed.")


def increase_pairs() -> List[List[str]]:
    """The GRN card as the label : value lines a manual tester reads.

    Every line is a lookup. The document references come from the register the
    flow filled in as the application allotted them, the quantities from the
    register the calculations filled in, and Expected / Actual / Difference /
    Result off the Movement - which took them off the recorded check.
    """
    card = movement_by_stage(GRN_STAGE)
    if card is None:
        return []
    asked = _asked(card)
    received = asked.get("GRN Approved / Received Quantity")
    name, code = item_identity(card.item or item_name())
    return [
        ["Item Name", name],
        ["Item ID", code or NOT_CAPTURED],
        ["PO Number", reference("Purchase Order")],
        ["PO Quantity", shown(quantity("Purchase Order Quantity"))],
        ["GRN Number", card.document or NOT_CAPTURED],
        ["GRN Approved Quantity",
         NOT_CAPTURED if received is None else num(float(received))],
        # The two badges, each printed under the record it belongs to. They are
        # read back from the status register, which verify_status() filled in
        # from the badge it had already read to assert it - so a card can only
        # ever say what the screen said, and says NOT CAPTURED when this run
        # never got as far as reading one.
        ["GRN Status", status_shown("GRN")],
        ["Purchase Bill Number", reference("Purchase Bill")],
        ["Purchase Bill Status", status_shown("Purchase Bill")],
        ["Current Stock Before Purchase", shown(card.before)],
        ["Expected Current Stock", shown(card.expected)],
        ["Current Stock", shown(card.actual)],
        ["Difference", shown(card.difference)],
        ["Result", mark(card.result)],
    ]


def decrease_pairs() -> List[List[str]]:
    """The Material Indent card as the same label : value lines.

    REQUESTED, APPROVED, ISSUED and RETURNED each get a line of their own, and
    the sum underneath uses the ISSUED one. That separation is the whole point
    of this section: this application issues what it can, so a report that let
    the requested quantity stand in for the issued one would call a correct
    application wrong by the difference between them.
    """
    card = movement_by_stage(INDENT_STAGE)
    if card is None:
        return []
    asked = _asked(card)

    def line(label: str) -> str:
        value = asked.get(label)
        return NOT_CAPTURED if value is None else num(float(value))

    name, code = item_identity(card.item or item_name())
    return [
        ["Item Name", name],
        ["Item ID", code or NOT_CAPTURED],
        ["Quotation", reference("Quotation")],
        ["Quotation Quantity", shown(quantity("Quotation Quantity"))],
        ["Sales Order", reference("Sales Order")],
        ["Sales Order Quantity", shown(quantity("Sales Order Quantity"))],
        ["Indent Number", card.document or NOT_CAPTURED],
        ["Current Stock Before Consumption", shown(card.before)],
        ["Material Indent Quantity (Requested)", line("Requested Quantity")],
        ["Approved Quantity", line("Approved Quantity")],
        ["Actual Issued Quantity", line("Actual Issued Quantity")],
        ["Returned Quantity", line("Returned Quantity")],
        ["Expected Current Stock", shown(card.expected)],
        ["Current Stock", shown(card.actual)],
        ["Difference", shown(card.difference)],
        ["Result", mark(card.result)],
    ]


def worked_sum(card: Movement, rule: str) -> str:
    """The arithmetic in the lines a tester writes out by hand.

        Current Stock Before Consumption - ACTUAL Issued Quantity
        234 - 7 = 227

        Expected Current Stock = 227
        Current Stock          = 227
        Difference             = 0
        Result                 = PASS

    `card.formula` already carries THIS row's own figures ("234 - 7"), so the
    sum can only ever be the sum the validation made.
    """
    show = card.show
    if card.expected is None:
        worked = ("This run did not work an expected figure out, because one "
                  "of the two readings the rule needs was never captured.")
    elif card.formula:
        worked = f"{card.formula} = {show(card.expected)}"
    else:
        worked = show(card.expected)
    width = max(10, len(f"Expected {card.subject}"))
    lines = [
        rule,
        worked,
        "",
        f"{('Expected ' + card.subject):<{width}} = {show(card.expected)}",
        f"{card.subject:<{width}} = {show(card.actual)}",
        f"{'Difference':<{width}} = {show(card.difference)}",
        f"{'Result':<{width}} = {mark(card.result)}",
    ]
    if card.source:
        lines += ["", f"{card.subject} is read from {card.source}."]
    return "\n".join(lines)


def _issue_caveat(card: Optional[Movement]) -> str:
    """Why the sum uses ISSUED, spelled out with THIS row's own figures.

    Said on every indent card, not only the ones where the two figures differ:
    a reader who is told the rule only when it bites cannot tell whether the
    other runs followed it.
    """
    if card is None:
        return ""
    asked = _asked(card)
    requested = asked.get("Requested Quantity")
    issued = asked.get("Actual Issued Quantity")
    wanted = NOT_CAPTURED if requested is None else num(float(requested))
    given = NOT_CAPTURED if issued is None else num(float(issued))
    line = (f"Requested {wanted}, actually issued {given}. The sum above uses "
            f"the ISSUED quantity.")
    if requested is not None and issued is not None and \
            not _same(float(requested), float(issued)):
        return (f"{line} The application issued a DIFFERENT quantity from the "
                f"one requested, so the stock was reduced by {given} and not "
                f"by {wanted}. A report built from the request would have "
                f"called a correct application wrong.")
    return (f"{line} This application issues what it can - a row asking for "
            f"more than the store holds is issued what the store holds - so "
            f"the requested quantity is never assumed to be the issued one.")


def _transposed(pairs: List[List[str]]) -> Tuple[List[str], List[str]]:
    """A Field/Value pairs list, pivoted into (headers, one data row).

    Each field becomes its own COLUMN instead of its own ROW - "Item Name",
    "Item ID", "Expected Stock", ... run left to right along the top, with
    this item's values in the single row underneath, the way a QA reader
    reads a spreadsheet rather than a form.
    """
    headers = [str(label) for label, _ in pairs]
    row = [str(value) for _, value in pairs]
    return headers, row


#: A test's DESCRIPTION panel has no stylesheet of its own to reach (see
#: qa_report.embed()'s own docstring on why it never carries a <style> tag),
#: so a table built for it needs its border on every cell INLINE - the class-
#: based rules in qa_report._CSS (used by every ATTACHMENT page, which IS its
#: own self-contained document) never apply here and the table rendered as
#: unbordered, "floating" text. Kept local to this embedding path only -
#: qa_report.table_block()/_CSS are shared by many other cards and are left
#: exactly as they are.
_EMBED_CELL_STYLE = ("border:1px solid #d0d7de;padding:8px 10px;"
                     "text-align:center;vertical-align:middle")
_EMBED_HEAD_STYLE = (_EMBED_CELL_STYLE + ";background:#eef2f7;"
                     "font-weight:650;color:#2c3a4d;font-size:12.5px;"
                     "text-transform:uppercase;letter-spacing:.4px")


def _bordered_table_html(headers: List[str], row: List[str],
                         verdict_index: int = -1) -> str:
    """One header row + one data row as a fully inline-styled <table>,
    wrapped so it scrolls horizontally instead of collapsing into a vertical
    layout when it is wider than the screen."""
    width = len(headers)
    at = verdict_index if verdict_index >= 0 else width - 1
    head = "".join(f"<th style='{_EMBED_HEAD_STYLE}'>{qa_report.escape(h)}</th>"
                   for h in headers)
    cells = []
    for index in range(width):
        value = row[index] if index < len(row) else ""
        content = (qa_report.badge(value) if index == at
                   else qa_report.escape(value))
        cells.append(f"<td style='{_EMBED_CELL_STYLE}'>{content}</td>")
    return ("<div style='width:100%;overflow-x:auto'>"
            "<table style='width:max-content;min-width:100%;"
            "border-collapse:collapse'>"
            f"<thead><tr>{head}</tr></thead>"
            f"<tbody><tr>{''.join(cells)}</tr></tbody>"
            "</table></div>")


def _table_html(pairs: List[List[str]], card: Optional[Movement]) -> str:
    """The pairs as a bare horizontal <table> fragment - no page, no title,
    no kpi tiles - for embedding directly into a test's own description
    panel (qa_report.embed()'s `table_html`), so the SAME figures a reader
    sees there are never ALSO dumped underneath as a <pre> block."""
    if card is None or not pairs:
        return ""
    headers, row = _transposed(pairs)
    if row:
        row[-1] = card.result          # bare word - see _section_page()'s note
    return _bordered_table_html(headers, row)


def item_table_html() -> str:
    """ITEM STOCK - BEFORE TRANSACTION, as a bare <table> for a description."""
    return _table_html(item_pairs(), movement_by_stage(ITEM_STAGE))


def increase_table_html() -> str:
    """ITEM STOCK - AFTER PURCHASE, as a bare <table> for a description."""
    return _table_html(increase_pairs(), movement_by_stage(GRN_STAGE))


def _section_text(heading: str, pairs: List[List[str]],
                  card: Optional[Movement], rule: str, missing: str,
                  table: bool = False) -> str:
    """One section as the plain text a QA person pastes into a ticket.

    `table` renders `pairs` HORIZONTALLY - each field its own column header,
    with the one row of values directly underneath it (text_table() pivoted
    via _transposed()) - instead of the colon-aligned Field : Value list.
    Opt-in, so a caller that does not ask for it keeps the exact layout it
    always had.
    """
    if card is None or not pairs:
        return qa_report.joined([qa_report.banner(heading), missing])
    if table:
        headers, row = _transposed(pairs)
        body = text_table(headers, [row])
    else:
        body = qa_report.labelled(pairs)
    parts = [qa_report.banner(heading),
             body,
             qa_report.section("FORMULA", worked_sum(card, rule))]
    for extra in (card.transaction_note, card.note):
        if extra:
            parts.append(extra)
    return qa_report.joined(parts)


def _section_page(title: str, pairs: List[List[str]],
                  card: Optional[Movement], rule: str, missing: str,
                  subject: str, extra_note: str = "",
                  table: bool = False) -> str:
    """The same section as the colour page a stakeholder opens.

    `table` renders `pairs` as a genuine HORIZONTAL HTML table - a <thead> row
    of column headers (one per field) over a <tbody> row of this item's values
    (qa_report.table_block(), pivoted via _transposed()) - instead of
    qa_report.facts_block()'s one-row-per-field layout. Wrapped in its own
    scrollable container so a wide row (many purchase-detail columns) scrolls
    horizontally rather than forcing anything to wrap onto a new row. Opt-in,
    so a caller that does not ask for it keeps the exact layout it always had.
    """
    if card is None or not pairs:
        return qa_report.page(title, "",
                              [qa_report.note_block("Nothing to show", missing)],
                              verdict=NOT_CHECKED)
    if table:
        headers, row = _transposed(pairs)
        # pairs' own Result value is mark(card.result) - "✅ PASS" - built for
        # the plain-text table, where there is no colour and the symbol is the
        # only visual cue. table_block() colours its verdict column itself
        # (qa_report.badge()/style_of()) but only recognises the BARE word, so
        # the decorated text is swapped for card.result here - otherwise the
        # badge falls back to a colourless grey for a result that did pass.
        if row:
            row[-1] = card.result
        transaction_block = (
            "<div style='overflow-x:auto'>"
            + qa_report.table_block("The transaction, in full", headers, [row],
                                    note=card.transaction_note)
            + "</div>")
    else:
        transaction_block = qa_report.facts_block(
            "The transaction, in full", pairs, note=card.transaction_note)
    blocks = [
        qa_report.kpi_block(f"Did the stock move correctly? ({subject})", [
            [card.before_label, card.show(card.before)],
            [f"Expected {card.subject}", card.show(card.expected)],
            [card.subject, card.show(card.actual)],
            ["Difference", card.show(card.difference)],
            ["Result", card.result, card.result],
        ]),
        transaction_block,
        qa_report.pre_block("How the expected figure was worked out",
                            worked_sum(card, rule)),
        qa_report.note_block("Observation", card.note),
        qa_report.note_block("Requested is not Issued", extra_note),
    ]
    return qa_report.page(title, card.item or item_name(), blocks,
                          verdict=card.result)


def increase_text() -> str:
    """ITEM STOCK - AFTER PURCHASE, as plain text."""
    return _section_text("ITEM STOCK - AFTER PURCHASE", increase_pairs(),
                         movement_by_stage(GRN_STAGE), INCREASE_RULE,
                         _NO_INCREASE, table=True)


def increase_page() -> str:
    """The same section as an HTML page."""
    return _section_page("Item Stock - After Purchase", increase_pairs(),
                         movement_by_stage(GRN_STAGE), INCREASE_RULE,
                         _NO_INCREASE, "GRN", table=True)


def decrease_text() -> str:
    """ITEM STOCK - AFTER CONSUMPTION, as plain text."""
    card = movement_by_stage(INDENT_STAGE)
    body = _section_text("ITEM STOCK - AFTER CONSUMPTION",
                         decrease_pairs(), card, DECREASE_RULE, _NO_DECREASE,
                         table=True)
    caveat = _issue_caveat(card)
    return qa_report.joined([body, caveat]) if caveat else body


def decrease_page() -> str:
    """The same section as an HTML page."""
    card = movement_by_stage(INDENT_STAGE)
    return _section_page("Item Stock - After Consumption",
                         decrease_pairs(), card, DECREASE_RULE, _NO_DECREASE,
                         "Material Indent", extra_note=_issue_caveat(card),
                         table=True)


# --------------------------------------------------------------------------- #
# STOCK AFTER SALES - Quotation + Sales Order, BEFORE the Material Indent
# --------------------------------------------------------------------------- #

def sales_pairs() -> List[List[str]]:
    """The post-sales checkpoint as the label : value lines a reader wants.

    The quantities are shown for CONTEXT - what was quoted and what was
    ordered - never as something this card expects to have moved the stock:
    neither document touches Items & Inventory in this application, which is
    exactly what the Expected / Actual pair below is proving with a real
    reading instead of assuming it.
    """
    card = movement_by_stage(SALES_STAGE)
    if card is None:
        return []
    name, code = item_identity(card.item or item_name())
    return [
        ["Item Name", name],
        ["Item ID", code or NOT_CAPTURED],
        ["Quotation", reference("Quotation")],
        ["Quotation Quantity", shown(quantity("Quotation Quantity"))],
        ["Sales Order", reference("Sales Order")],
        ["Sales Order Quantity", shown(quantity("Sales Order Quantity"))],
        ["Current Stock After Purchase", shown(card.before)],
        ["Expected Current Stock", shown(card.expected)],
        ["Current Stock", shown(card.actual)],
        ["Difference", shown(card.difference)],
        ["Result", mark(card.result)],
    ]


def sales_text() -> str:
    """ITEM STOCK - AFTER SALES (QUOTATION + SALES ORDER), as plain text."""
    return _section_text("ITEM STOCK - AFTER SALES (QUOTATION + SALES ORDER)",
                         sales_pairs(), movement_by_stage(SALES_STAGE),
                         SALES_RULE, _NO_SALES, table=True)


def sales_page() -> str:
    """The same section as an HTML page."""
    return _section_page("Item Stock - After Sales (Quotation + Sales Order)",
                         sales_pairs(), movement_by_stage(SALES_STAGE),
                         SALES_RULE, _NO_SALES, "Quotation + Sales Order",
                         table=True)


# --------------------------------------------------------------------------- #
# STOCK MOVEMENT SUMMARY - before, expected, actual, verdict, one row each
#
#     Transaction    Quantity  Stock Before  Expected  Actual  Result
#     ------------------------------------------------------------------
#     Initial Stock       -          -           -         0   MEASURED
#     GRN Approved     +234          0         234       234   PASS
#     Indent Issue       -7        234         227       227   PASS
#
# The opening line is MEASURED and not PASS on purpose: it is the reading every
# later verdict is measured FROM, and there was nothing yet to check it
# against. A green tick on a row nobody checked is the one thing that would
# make the other two rows worth less.
# --------------------------------------------------------------------------- #

MOVEMENT_HEADINGS = ("Transaction", "Item", "Current Stock Before",
                     "Qty Added/Removed", "Expected Current Stock",
                     "Current Stock", "Difference", "Result")


def movement_rows(movements: Optional[Sequence[Movement]] = None
                  ) -> List[List[str]]:
    """The BEFORE / AFTER table - one row per transaction that touched stock.

        Transaction | Item | Current Stock Before | Qty Added/Removed |
        Expected Current Stock | Current Stock | Difference | Result

    Every cell is read off a recorded Movement, which read it off the check
    calculation_validation.py made. Nothing is added up here, and the Actual
    column is always the application's own figure - see STOCK_COLUMN.

    The synthetic "Initial Stock" line is only drawn when this run recorded no
    Item Created card. When it did, THAT card is the opening line and carries
    a real expected-vs-actual verdict instead of the word MEASURED.
    """
    cards = [card for card in (MOVEMENTS if movements is None else movements)
             if not card.amount]
    if not cards:
        return []

    rows: List[List[str]] = []
    if not any(card.stage == ITEM_STAGE for card in cards):
        opening = opening_stock(cards)
        rows.append([
            "Initial Stock", cards[0].item, "-", "-", "-", shown(opening), "-",
            "MEASURED" if opening is not None else NOT_CHECKED,
        ])
    for card in cards:
        show = card.show
        moving = [term for term in card.terms if term[2] is not None]
        if moving:
            moved = ", ".join(f"{sign}{show(value)}"
                              for sign, _, value in moving)
        else:
            # A card WITH terms whose figures were never read is not a card
            # that moved nothing - it is a card nobody could read. "0" there
            # would be this report inventing a quantity.
            moved = NOT_CAPTURED if card.terms else "0"
        rows.append([
            card.stage or card.title.replace(" VALIDATION", "").title(),
            card.item, shown(card.before), moved, shown(card.expected),
            shown(card.actual), shown(card.difference), card.result,
        ])
    return rows


# --------------------------------------------------------------------------- #
# THE TRANSACTION CHAIN - which record did what, in the order it happened
#
#     Item              OPC Cement 53 Grade
#      |
#      v
#     Purchase Order    PO-00061            x 234
#      |
#      v
#     GRN               GRN/0058            x 234
#     ...
#
# Every reference is one the application allotted and the flow read off the
# screen. A step this run never reached prints "(not captured by this run)" and
# stays in the chain, because a missing link is information: it is the point
# the story stops.
# --------------------------------------------------------------------------- #

def chain_rows() -> List[List[str]]:
    """Step | Reference | Quantity | What it did, top to bottom."""
    grn = movement_by_stage(GRN_STAGE)
    sales = movement_by_stage(SALES_STAGE)
    indent = movement_by_stage(INDENT_STAGE)
    received = _asked(grn).get("GRN Approved / Received Quantity")
    asked = _asked(indent)

    def q(value: object) -> str:
        return "-" if value is None else num(float(value))

    after_grn = shown(grn.actual) if grn is not None else NOT_CAPTURED
    after_sales = shown(sales.actual) if sales is not None else NOT_CAPTURED
    after_indent = shown(indent.actual) if indent is not None else NOT_CAPTURED

    return [
        ["Item", item_name(), "-", "the catalogue item this chain moves"],
        ["Purchase Order", reference("Purchase Order"),
         q(quantity("Purchase Order Quantity")), "the quantity ordered"],
        ["GRN", reference("GRN"), q(received),
         "the receipt raised against that order"],
        ["GRN Approval", reference("GRN"), q(received),
         "the approval that puts the goods into stock"],
        ["Purchase Bill", reference("Purchase Bill"), "-",
         "the bill the approved receipt created"],
        ["Stock Increased", "-", q(received),
         f"stock is now {after_grn}"],
        ["Quotation", reference("Quotation"),
         q(quantity("Quotation Quantity")), "the quantity quoted"],
        ["Sales Order", reference("Sales Order"),
         q(quantity("Sales Order Quantity")),
         "the quantity the customer ordered"],
        ["Stock Checked", "-", "-",
         f"stock after Quotation + Sales Order is {after_sales} (no movement "
         f"expected yet - the indent below is what actually issues it)"],
        ["Material Indent", reference("Material Indent"),
         q(asked.get("Requested Quantity")), "the quantity requested"],
        ["Material Issued", reference("Material Indent"),
         q(asked.get("Actual Issued Quantity")),
         "the quantity the store ACTUALLY gave out"],
        ["Stock Decreased", "-", q(asked.get("Actual Issued Quantity")),
         f"stock is now {after_indent}"],
    ]


def chain_text() -> str:
    """The chain as one table: which record did what, in the order it happened.

    A TABLE and not the arrow diagram it used to be. The diagram spent three
    lines on every step - the step, the reference, and a sentence in brackets
    - which is thirty-three lines to say what eleven rows say, and it pushed
    the stock figures off the first screen of the attachment.
    """
    return qa_report.joined([
        qa_report.banner("TRANSACTION CHAIN"),
        text_table(("Step", "Reference", "Quantity", "What it did"),
                   chain_rows()),
        f"A step showing {NOT_CAPTURED} is one this run did not reach, or one "
        f"this build prints no number for. Nothing was filled in for it.",
    ])


def chain_block() -> str:
    """The same chain as the table of an HTML page."""
    return qa_report.table_block(
        "The transaction chain, and the records behind it",
        ("Step", "Reference", "Quantity", "What it did"), chain_rows(),
        note="Top to bottom is the order the records were created. Every "
             "reference is the one the application allotted; a step this run "
             "did not reach says so rather than showing a number.",
        verdict_column=99, legend=False)


# --------------------------------------------------------------------------- #
# THE WHOLE SECTION - the three parts, one under the other
# --------------------------------------------------------------------------- #

def summary_text(item: str = "") -> str:
    """STOCK MOVEMENT SUMMARY, as the table plus the chain, in plain text."""
    rows = movement_rows()
    if not rows:
        return ("STOCK MOVEMENT SUMMARY\n\n"
                "This run moved no stock, so there is no movement to "
                "summarise - and nothing was assumed.")

    tally = counts()
    return qa_report.joined([
        qa_report.banner("STOCK MOVEMENT SUMMARY"),
        qa_report.labelled([["Item", item or item_name()]]),
        text_table(MOVEMENT_HEADINGS, rows),
        final_stock_text(),
        qa_report.labelled([
            ["Stock Validations Checked", str(tally["total"])],
            ["Passed", str(tally["passed"])],
            ["Failed", str(tally["failed"])],
            ["Overall Stock Result", mark(result())],
        ]),
        chain_text(),
    ])


def summary_page(item: str = "") -> str:
    """The same summary as the colour page a stakeholder opens."""
    rows = movement_rows()
    tally = counts()
    blocks = [qa_report.kpi_block("Stock movement", [
        ["Checked", tally["total"]],
        ["Passed", tally["passed"], PASS if tally["passed"] else ""],
        ["Failed", tally["failed"], FAIL if tally["failed"] else ""],
        ["Overall Result", result(), result()],
    ])]
    if rows:
        blocks.append(final_block())
        blocks.append(qa_report.table_block(
            "Stock before, quantity moved, stock after", MOVEMENT_HEADINGS,
            rows,
            note=f"{STOCK_COLUMN} is what {STOCK_SOURCE} showed after that "
                 f"transaction - a reading, not a running total worked out "
                 f"here. The opening line is MEASURED because there was "
                 f"nothing yet to check it against."))
    else:
        blocks.append(qa_report.note_block(
            "Nothing to show",
            "This run moved no stock, so there is no movement to summarise."))
    blocks.append(chain_block())
    if DOCUMENTS:
        blocks.append(qa_report.table_block(
            "The documents behind these movements",
            ("Document", "Reference", "Status", "What it is"),
            documents_with_status(), note=DOCUMENT_STATUS_NOTE,
            verdict_column=99, legend=False))
    return qa_report.page("Stock Movement Summary", item or item_name(),
                          blocks, verdict=result())


def increase_result() -> str:
    """The GRN stock movement's verdict, for the run's one-line summary."""
    card = movement_by_stage(GRN_STAGE)
    return NOT_CHECKED if card is None else card.result


def decrease_result() -> str:
    """The Material Indent movement's verdict, for the same summary."""
    card = movement_by_stage(INDENT_STAGE)
    return NOT_CHECKED if card is None else card.result


# =========================================================================== #
# 📦 CURRENT STOCK VALIDATION - the page this report is opened for
#
#     CURRENT STOCK BEFORE  ->  WHAT MOVED  ->  EXPECTED CURRENT STOCK
#                           vs  ACTUAL CURRENT STOCK  ->  DIFFERENCE  ->  RESULT
#
# The most important question this suite answers, on one screen, in the order
# a QA person asks it. Everything below reads the Movement cards the run
# already recorded; not one figure is worked out here, and the two sides are
# deliberately kept apart:
#
#   EXPECTED CURRENT STOCK  the reporting/validation side - stock before plus
#                           what was received, or stock before minus what was
#                           ACTUALLY issued.
#   ACTUAL CURRENT STOCK    what the application itself showed in Items &
#                           Inventory, and nothing else.
#
# A figure this run never captured prints NOT AVAILABLE. Nothing is filled in,
# and no number on this page was invented to complete a sum.
# =========================================================================== #

#: What each card's closing reading is called on the FINAL STOCK block. Written
#: out rather than derived from the stage name, because "Stock After Stock
#: Increase (GRN)" is what deriving it produces.
AFTER_LABELS = {
    GRN_STAGE: "Current Stock After GRN",
    SALES_STAGE: "Current Stock After Quotation & Sales Order",
    INDENT_STAGE: "Current Stock After Material Indent",
}

ITEM_RULE = "A catalogue item created by this run starts with 0 stock"

_NO_ITEM = ("This run did not record an opening stock reading for a newly "
            "created item, so there is no initial stock card - and none was "
            "assumed.")

_NO_STOCK = ("This run recorded no stock movement, so there is no current "
             "stock to validate - and nothing was assumed.")


def stock_cards() -> List[Movement]:
    """The cards that MOVED stock, in the order they happened.

    Money cards (a supplier's outstanding) and the opening reading are both
    left out: the first is not stock, and the second is the figure the others
    are measured from rather than a movement of its own.
    """
    return [card for card in MOVEMENTS
            if not card.amount and card.stage != ITEM_STAGE]


def opening_stock(movements: Optional[Sequence[Movement]] = None
                  ) -> Optional[float]:
    """The stock this run started from - what every later verdict measures from.

    The Item Created card's reading when this run took one, and otherwise the
    "before" of the first transaction. Never 0 by assumption: an item the
    catalogue already held does not start at 0, and a reading nobody took is
    not a reading of zero.
    """
    cards = [card for card in (MOVEMENTS if movements is None else movements)
             if not card.amount]
    initial = next((card for card in cards if card.stage == ITEM_STAGE), None)
    if initial is not None and initial.actual is not None:
        return initial.actual
    moved = [card for card in cards if card.stage != ITEM_STAGE]
    if moved:
        return moved[0].before
    return initial.before if initial is not None else None


def last_movement() -> Optional[Movement]:
    """The last card that moved stock - the one carrying the closing figure."""
    cards = stock_cards()
    return cards[-1] if cards else None


# --------------------------------------------------------------------------- #
# 1. THE NEW ITEM - CURRENT STOCK BEFORE ANY TRANSACTION
# --------------------------------------------------------------------------- #

def item_pairs() -> List[List[str]]:
    """The Item Created card as the label : value lines a QA reader reads."""
    card = movement_by_stage(ITEM_STAGE)
    if card is None:
        return []
    name, code = item_identity(card.item or item_name())
    return [
        ["Item Name", name],
        ["Item ID", code or NOT_CAPTURED],
        ["Current Stock Before Transaction", shown(card.before)],
        ["Expected Current Stock", shown(card.expected)],
        ["Current Stock", shown(card.actual)],
        ["Difference", shown(card.difference)],
        ["Result", mark(card.result)],
    ]


def item_text() -> str:
    """ITEM STOCK - BEFORE TRANSACTION, as plain text."""
    return _section_text("ITEM STOCK - BEFORE TRANSACTION", item_pairs(),
                         movement_by_stage(ITEM_STAGE), ITEM_RULE, _NO_ITEM,
                         table=True)


def item_page() -> str:
    """The same card as the HTML page a stakeholder opens."""
    return _section_page("Item Stock - Before Transaction", item_pairs(),
                         movement_by_stage(ITEM_STAGE), ITEM_RULE, _NO_ITEM,
                         "Initial Stock", table=True)


# --------------------------------------------------------------------------- #
# 2. THE FINAL STOCK - initial, everything that moved, what is left
# --------------------------------------------------------------------------- #

def _after_label(card: Movement) -> str:
    """What this card's closing reading is called on the final-stock block."""
    return AFTER_LABELS.get(
        card.stage, f"Current Stock After {card.stage or card.title}")


def final_stock_pairs() -> List[List[str]]:
    """Initial Stock, every term that moved it, and what is left.

        Initial Stock                     : 0
        GRN Approved                      : +234
        Current Stock After GRN           : 234
        Material Indent Issued            : -7
        Current Stock After Material Indent : 227
        Expected Current Stock            : 227
        Current Stock                     : 227
        Difference                        : 0
        Overall Stock Result              : PASS
    """
    cards = stock_cards()
    if not cards:
        return []
    last = cards[-1]
    name, code = item_identity(item_name())
    pairs: List[List[str]] = [
        ["Item Name", name],
        ["Item ID", code or NOT_CAPTURED],
        ["Initial Stock", shown(opening_stock())],
    ]
    for card in cards:
        for sign, label, value in card.terms:
            if value is None:
                continue
            pairs.append([str(label), f"{sign}{num(float(value))}"])
        pairs.append([_after_label(card), shown(card.actual)])
    pairs += [
        ["Expected Current Stock", shown(last.expected)],
        ["Current Stock", shown(last.actual)],
        ["Difference", shown(last.difference)],
        ["Overall Stock Result", mark(result())],
    ]
    return pairs


def final_formula() -> str:
    """"0 + 234 - 7 = 227", written out of the figures this run recorded.

    Empty when any part of the chain was never captured. A sum with a hole in
    it is worse than no sum: it looks like arithmetic and is not.
    """
    initial = opening_stock()
    cards = stock_cards()
    if initial is None or not cards:
        return ""
    total = float(initial)
    parts = [num(initial)]
    for card in cards:
        for sign, _, value in card.terms:
            if value is None:
                return ""
            parts.append(f"{sign} {num(float(value))}")
            total += float(value) if sign == "+" else -float(value)
    if len(parts) == 1:
        return ""
    return " ".join(parts) + f" = {num(total)}"


def _final_caveat() -> str:
    """Said only when the chain and the last check do not land on one figure."""
    last = last_movement()
    formula = final_formula()
    if last is None or not formula or last.expected is None:
        return ""
    chained = formula.rsplit("=", 1)[-1].strip()
    if chained == num(last.expected):
        return ""
    return (f"The chain above adds up to {chained}, and the last "
            f"transaction's own expected figure is {num(last.expected)}. They "
            f"differ because every stage is checked against the stock that was "
            f"actually READ before it, never against the stock the stage "
            f"before it should have left behind - which is what stops one "
            f"wrong figure hiding the next one.")


def final_stock_text() -> str:
    """FINAL CURRENT STOCK VALIDATION, as plain text."""
    pairs = final_stock_pairs()
    if not pairs:
        return ""
    return qa_report.joined([
        qa_report.banner("FINAL CURRENT STOCK VALIDATION", "-"),
        qa_report.labelled(pairs),
        qa_report.section("FORMULA", final_formula()),
        _final_caveat(),
    ])


def final_block() -> str:
    """The same final stock as the blocks of an HTML page."""
    last = last_movement()
    if last is None:
        return ""
    return "\n".join(block for block in [
        qa_report.kpi_block("Final Current Stock Validation", [
            ["Initial Stock", num(opening_stock())],
            ["Expected Current Stock", num(last.expected)],
            ["Current Stock", num(last.actual)],
            ["Difference", num(last.difference)],
            ["Overall Result", result(), result()],
        ]),
        qa_report.facts_block("How the stock got there", final_stock_pairs(),
                              note=final_formula()),
        qa_report.note_block("Note", _final_caveat()),
    ] if block)


# --------------------------------------------------------------------------- #
# 3. THE WHOLE THING ON ONE PAGE
# --------------------------------------------------------------------------- #

def _transaction_pairs(card: Movement) -> List[List[str]]:
    """One transaction in the lines a QA reader needs, and no others.

    Stock before, what actually moved, what the stock SHOULD be, what the
    application shows, the difference, the verdict. The requested and approved
    quantities are not here on purpose - they move no stock, they are on the
    section card that is about them, and on this page they would sit in a
    column a reader is adding up.
    """
    show = card.show
    pairs: List[List[str]] = []
    if card.document:
        pairs.append(["Document", card.document])
    pairs.append([card.before_label, shown(card.before)])
    for sign, label, value in card.terms:
        pairs.append([str(label), NOT_CAPTURED if value is None
                      else f"{sign}{show(value)}"])
    pairs += [
        [f"Expected {card.subject}", shown(card.expected)],
        [card.subject, shown(card.actual)],
        ["Difference", shown(card.difference)],
        ["Result", mark(card.result)],
    ]
    return pairs


def _transaction_text(heading: str, card: Optional[Movement]) -> str:
    """One numbered transaction block of the current-stock page."""
    if card is None:
        return ""
    parts = [qa_report.banner(heading, "-"),
             qa_report.labelled(_transaction_pairs(card))]
    if card.formula:
        parts.append(f"Formula:\n{card.formula} = {card.show(card.expected)}")
    if card.transaction_note:
        parts.append(card.transaction_note)
    return qa_report.joined(parts)


def _numbered() -> List[List[object]]:
    """(heading, card) for every stage this run reached, in order."""
    stages = [(ITEM_STAGE, "ITEM CREATED - CURRENT STOCK BEFORE TRANSACTION"),
              (GRN_STAGE, "AFTER GRN"),
              (SALES_STAGE, "AFTER QUOTATION + SALES ORDER"),
              (INDENT_STAGE, "AFTER MATERIAL INDENT")]
    found = [(heading, movement_by_stage(stage)) for stage, heading in stages]
    numbered: List[List[object]] = []
    for heading, card in found:
        if card is None:
            continue
        numbered.append([f"{len(numbered) + 1}. {heading}", card])
    return numbered


def current_stock_text() -> str:
    """CURRENT STOCK VALIDATION - the whole stock story, in one attachment.

    The order a QA person reads it in: what the item started with, what each
    transaction did to it, the before-and-after table, and the final stock with
    the sum written out.
    """
    numbered = _numbered()
    if not numbered:
        return qa_report.joined([qa_report.banner("CURRENT STOCK VALIDATION"),
                                 _NO_STOCK])
    parts = [
        qa_report.banner("CURRENT STOCK VALIDATION"),
        qa_report.labelled([
            ["Item", item_name()],
            ["Expected Current Stock", "worked out by this report: stock "
                                       "before + received, or stock before "
                                       "- ACTUALLY issued"],
            [STOCK_COLUMN, f"read from {STOCK_SOURCE}"],
        ]),
    ]
    parts += [_transaction_text(heading, card) for heading, card in numbered]
    parts += [
        qa_report.section("CURRENT STOCK - BEFORE AND AFTER EVERY TRANSACTION",
                          text_table(MOVEMENT_HEADINGS, movement_rows())),
        final_stock_text(),
    ]
    return qa_report.joined(parts)


def current_stock_page() -> str:
    """The same page as the colour HTML a stakeholder opens."""
    numbered = _numbered()
    if not numbered:
        return qa_report.page("Current Stock Validation", "",
                              [qa_report.note_block("Nothing to show",
                                                    _NO_STOCK)],
                              verdict=NOT_CHECKED)
    tally = counts()
    blocks = [qa_report.kpi_block("Current stock validation", [
        ["Checked", tally["total"]],
        ["Passed", tally["passed"], PASS if tally["passed"] else ""],
        ["Failed", tally["failed"], FAIL if tally["failed"] else ""],
        ["Overall Result", result(), result()],
    ])]
    for heading, card in numbered:
        blocks.append(qa_report.facts_block(
            str(heading), _transaction_pairs(card),
            note=(f"{card.formula} = {card.show(card.expected)}"
                  if card.formula else "")))
        blocks.append(qa_report.note_block("Observation",
                                           card.transaction_note))
    blocks.append(qa_report.table_block(
        "Current stock - before and after every transaction",
        MOVEMENT_HEADINGS, movement_rows(),
        note=f"Expected Current Stock is worked out by this report; "
             f"{STOCK_COLUMN} is read from {STOCK_SOURCE}. A blank or NOT "
             f"AVAILABLE is a figure this run did not capture - nothing was "
             f"filled in."))
    blocks.append(final_block())
    return qa_report.page("Current Stock Validation", item_name(), blocks,
                          verdict=result())


# =========================================================================== #
# 📦 STOCK MOVEMENT - the six lines a reader wants before any of the sections
# above, not instead of them
#
#     Item Created
#     0 -> 0
#
#     GRN Approved
#     0 + 234 = 234
#     Expected: 234   Actual: 234   Difference: 0   PASS
#
#     Material Indent Issued
#     234 - 7 = 227
#     Expected: 227   Actual: 227   Difference: 0   PASS
#
#     Final Stock: 227
#
# Every figure here is read off a Movement the sections above already
# recorded - card.before / .actual / .expected / .difference / .result / and
# .formula, which already carries THIS run's own numbers ("234 - 7"). Nothing
# is worked out, and nothing is worked out AGAIN: this is a shorter sentence
# about the same cards, not a second opinion on them.
# =========================================================================== #

def _simple_line(card: Movement) -> str:
    """"0 + 234 = 234" / "234 - 7 = 227" - the card's own formula, closed out.

    `card.formula` is the left-hand side only ("234 - 7"); the closing
    "= <expected>" is added here so this line reads as one piece of
    arithmetic, the way a manual tester would write it on paper.
    """
    if not card.formula:
        return "(not enough was captured to work this out)"
    return f"{card.formula} = {card.show(card.expected)}"


def simple_movement_text() -> str:
    """📦 STOCK MOVEMENT - the whole chain in six lines, plain text.

    Deliberately carries NONE of the context the sections above carry - no PO
    number, no GRN status, no requested-vs-issued caveat. That context is
    what makes THOSE sections trustworthy; it is also what makes them the
    wrong thing to open first. This is the one a reader opens first, for the
    plain yes-or-no, before going to the fuller card for how.
    """
    item_card = movement_by_stage(ITEM_STAGE)
    grn_card = movement_by_stage(GRN_STAGE)
    sales_card = movement_by_stage(SALES_STAGE)
    indent_card = movement_by_stage(INDENT_STAGE)
    if (item_card is None and grn_card is None and sales_card is None
            and indent_card is None):
        return qa_report.joined([qa_report.banner("📦 STOCK MOVEMENT"),
                                 _NO_STOCK])

    lines = [qa_report.banner("📦 STOCK MOVEMENT"), "", f"Item: {item_name()}"]

    if item_card is not None:
        lines += ["", "Item Created",
                  f"{shown(item_card.before)} -> {shown(item_card.actual)}",
                  mark(item_card.result)]

    for label, card in (("GRN Approved", grn_card),
                       ("Quotation + Sales Order Checked", sales_card),
                       ("Material Indent Issued", indent_card)):
        if card is None:
            continue
        lines += ["", label, _simple_line(card),
                  f"Expected Current Stock: {shown(card.expected)}   "
                  f"Current Stock: {shown(card.actual)}   "
                  f"Difference: {shown(card.difference)}",
                  mark(card.result)]

    final = last_movement()
    if final is not None:
        lines += ["", f"Current Stock: {shown(final.actual)}"]
    return "\n".join(lines) + "\n"


def simple_movement_page() -> str:
    """The same six lines, as the HTML card a stakeholder opens first."""
    item_card = movement_by_stage(ITEM_STAGE)
    grn_card = movement_by_stage(GRN_STAGE)
    sales_card = movement_by_stage(SALES_STAGE)
    indent_card = movement_by_stage(INDENT_STAGE)
    if (item_card is None and grn_card is None and sales_card is None
            and indent_card is None):
        return qa_report.page("Stock Movement", "",
                              [qa_report.note_block("Nothing to show",
                                                    _NO_STOCK)],
                              verdict=NOT_CHECKED)

    blocks = []
    if item_card is not None:
        blocks.append(qa_report.kpi_block("Item Created", [
            ["Before", shown(item_card.before)],
            ["After", shown(item_card.actual)],
            ["Result", item_card.result, item_card.result],
        ]))
    for label, card in (("GRN Approved", grn_card),
                       ("Quotation + Sales Order Checked", sales_card),
                       ("Material Indent Issued", indent_card)):
        if card is None:
            continue
        blocks.append(qa_report.kpi_block(label, [
            ["Formula", _simple_line(card)],
            ["Expected Current Stock", shown(card.expected)],
            ["Current Stock", shown(card.actual)],
            ["Difference", shown(card.difference)],
            ["Result", card.result, card.result],
        ]))
    final = last_movement()
    if final is not None:
        blocks.append(qa_report.kpi_block(
            "Final Stock", [["Current Stock", shown(final.actual)]]))
    return qa_report.page("Stock Movement", item_name(), blocks,
                          verdict=result())


# =========================================================================== #
# 📦 THE TOP-OF-REPORT HEADLINE - Initial / Increase / Decrease / Final, each
# as Previous Stock -> Quantity -> Expected -> Actual -> Calculation -> Result,
# and nothing else.
#
# This is a FOURTH view of the same three Movement records the sections above
# already carry - it computes nothing, reads nothing from a page, and cannot
# disagree with `increase_text()` / `decrease_text()` / `final_stock_text()`,
# because every figure here is that same card's `.before` / `.actual` /
# `.expected` / `.result` / `.formula` / `.terms`. It exists only to be short
# enough to embed in the test's own DESCRIPTION (qa_report.describe(), which
# Allure draws above the steps and attachments - see
# Construction_Flow.test_validation_summary) so a QA or management reader gets
# the plain answer before opening a single attachment.
# =========================================================================== #

def _headline_quantities(card: Movement) -> List[str]:
    """Each term that moved this card's stock - "<label>: <value>", no sign.

    Whatever the run actually captured - a GRN's one Received quantity, or an
    indent's Issued AND Returned - printed as its own line rather than folded
    into a single made-up "quantity" that would misstate what moved.
    """
    return [f"{label}: {card.show(value)}"
            for _, label, value in card.terms if value is not None]


def _headline_section(emoji: str, heading: str, card: Optional[Movement],
                      missing: str, name: str = "", code: str = "") -> List[str]:
    """One transaction section: Item Name, Item ID, Current Stock Before,
    Quantity, Expected Current Stock, Current Stock, Calculation, Result -
    and nothing else.
    """
    lines = [f"{emoji} {heading}", ""]
    if card is None:
        return lines + [missing, ""]
    if name:
        lines.append(f"Item Name: {name}")
        lines.append(f"Item ID: {code or NOT_CAPTURED}")
    lines.append(f"{card.before_label}: {shown(card.before)}")
    lines += _headline_quantities(card)
    lines.append(f"Expected {card.subject}: {shown(card.expected)}")
    lines.append(f"{card.subject}: {shown(card.actual)}")
    if card.formula:
        lines.append(f"Calculation: {card.formula} = {card.show(card.expected)}")
    lines.append(mark(card.result))
    lines.append("")
    return lines


def top_stock_summary_text() -> str:
    """ITEM STOCK VALIDATION - the four sections a QA/management reader asks
    for first, in order: ITEM STOCK - BEFORE TRANSACTION, ITEM STOCK - AFTER
    PURCHASE, ITEM STOCK - AFTER CONSUMPTION, FINAL CURRENT STOCK VALIDATION -
    meant to be embedded ahead of the QA Summary card in the report's own
    description, never as a replacement for the fuller cards (`item_text()`,
    `increase_text()`, `decrease_text()`, `final_stock_text()`) still attached
    below it.
    """
    item_card = movement_by_stage(ITEM_STAGE)
    grn_card = movement_by_stage(GRN_STAGE)
    sales_card = movement_by_stage(SALES_STAGE)
    indent_card = movement_by_stage(INDENT_STAGE)
    if (item_card is None and grn_card is None and sales_card is None
            and indent_card is None):
        return qa_report.joined(["📦 ITEM STOCK VALIDATION", _NO_STOCK])

    name, code = item_identity(item_name())
    lines = ["📦 ITEM STOCK VALIDATION", "", f"Item: {name}", ""]

    if item_card is not None:
        lines += ["📦 ITEM STOCK - BEFORE TRANSACTION", "",
                  f"Item Name: {name}", f"Item ID: {code or NOT_CAPTURED}",
                  f"Expected Current Stock: {shown(item_card.expected)}",
                  f"Current Stock: {shown(item_card.actual)}",
                  mark(item_card.result), ""]

    lines += _headline_section("📈", "ITEM STOCK - AFTER PURCHASE", grn_card,
                               _NO_INCREASE, name, code)
    lines += _headline_section(
        "📊", "ITEM STOCK - AFTER SALES (QUOTATION + SALES ORDER)",
        sales_card, _NO_SALES, name, code)
    lines += _headline_section("📉", "ITEM STOCK - AFTER CONSUMPTION",
                               indent_card, _NO_DECREASE, name, code)

    final = last_movement()
    if final is not None:
        lines.append("📊 FINAL CURRENT STOCK VALIDATION")
        lines.append("")
        lines.append(f"Item Name: {name}")
        lines.append(f"Item ID: {code or NOT_CAPTURED}")
        lines.append(f"Initial Stock: {shown(opening_stock())}")
        for card in stock_cards():
            for sign, label, value in card.terms:
                if value is None:
                    continue
                lines.append(f"{label}: {sign}{num(float(value))}")
        formula = final_formula()
        lines.append(f"Expected Current Stock: {shown(final.expected)}")
        lines.append(f"Current Stock: {shown(final.actual)}")
        if formula:
            lines.append(f"Calculation: {formula}")
        lines.append(mark(result()))
    return "\n".join(lines).rstrip() + "\n"
