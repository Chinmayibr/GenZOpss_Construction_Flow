"""Calculation validation - "did the application work out the right number?"

`data_verification.py` answers "did the application STORE what was typed?".
This module answers the other half: "did the application CALCULATE the right
result?" - and it reports both halves per module, so a row that saved perfectly
while adding up wrongly is a FAIL, not a PASS.

    Functional Flow  +  Calculation Validation  =  Overall Module Result

THE TWO HALVES ARE ANSWERED SEPARATELY, AND STAY SEPARATE

The functional verdict and the calculation verdict are two answers to two
different questions, and neither is allowed to overwrite the other. A module
whose flow was BLOCKED - no OTP typed, no stock in the warehouse - never
reached a calculation, so it says nothing at all about the application's
arithmetic; a run in which every figure checked was correct reports

    Overall Calculation Result : PASS WITH SKIPS
    Functional Result          : BLOCKED
    Overall Project Result     : BLOCKED

and that is a consistent report, not a contradictory one. The calculation
verdict is worked out in calculation_result() from the CHECK counts alone -
see the note there for what letting a module status reach it used to do.

HOW IT IS USED

    calc.start_test("Quotations", row)          # once per Excel row
    calc.check("Line Amount",
               inputs={"Quantity": 567, "Rate": 345},
               formula="Quantity x Rate",
               expected=195_615,                # worked out HERE, from inputs
               actual=page_value)               # read off the application
    report = calc.finish_test()                 # PASS / FAIL / BLOCKED

Every check becomes a step in the SAME Allure report, carrying the seven things
a calculation review needs: the name, the inputs, the formula, Expected, Actual,
Difference and the Result. At the end of the run `summary()` builds the
whole-project table - module by module, functional vs calculation vs overall -
entirely from what actually executed. Nothing in the summary is hard-coded.

THREE RULES THIS MODULE IS BUILT ON

1.  EXPECTED AND ACTUAL MUST COME FROM DIFFERENT PLACES. `expected` is worked
    out in Python from the inputs the test typed (the workbook) or from a figure
    the application published earlier in the flow; `actual` is read off the
    screen after the application has done its own sum. Passing the same number
    in twice proves nothing, so check() refuses it - see SAME_SOURCE.

2.  NO INVENTED ARITHMETIC. A formula that is not known for certain is not
    guessed at: skip() records the check as SKIPPED with the reason, and the
    summary counts it as unverified. A module the application has no arithmetic
    on is NOT APPLICABLE, not PASS.

3.  A CONVENTION DIFFERENCE IS NOT A DEFECT UNTIL IT IS PROVED TO BE ONE. When
    a check fails, `alternatives` lets the caller list the other orderings the
    same figure could follow (tax before discount, say). If one of them matches
    to the penny, the report says so in the note. The check still FAILS - the
    declared formula is what was asked for - but nobody has to reverse-engineer
    the application from a bare "expected 21000, actual 21600".
"""

from __future__ import annotations

import logging
import os
import re
from contextlib import contextmanager
from dataclasses import dataclass, field
from decimal import Decimal, ROUND_HALF_UP
from typing import Dict, Iterator, List, Optional, Sequence, Tuple

# Reused rather than rewritten: one number parser, one table renderer, one
# Allure attachment helper for the whole framework.
from data_verification import (as_number, attach_html, clean_text, text_table)
import qa_card
# Presentation only: the badge, the tiles and the page every HTML attachment in
# the framework is drawn with. See the module docstring of qa_report.py.
import qa_report

log = logging.getLogger("construction_flow")

#: A money comparison is exact to the paisa. Anything looser hides a real
#: rounding fault; anything tighter fails on float arithmetic.
MONEY_TOLERANCE = 0.01

#: Percentages and quantities are compared a little more loosely, because the
#: screen rounds them for display (33.33%, 12.5 MT).
RATIO_TOLERANCE = 0.05

#: A calculation mismatch fails the row. Set GENZ_FAIL_ON_CALC_MISMATCH=0 to go
#: back to reporting them without failing - the numbers still appear in the
#: report either way.
FAIL_ON_MISMATCH = os.environ.get("GENZ_FAIL_ON_CALC_MISMATCH", "1") != "0"

PASS = "PASS"
FAIL = "FAIL"
SKIPPED = "SKIPPED"
BLOCKED = "BLOCKED"
NOT_APPLICABLE = "NOT APPLICABLE"
NOT_AUTOMATED = "NOT AUTOMATED"

#: Some arithmetic was proved right and some could not be checked at all. It is
#: NOT a FAIL - nothing was found to be wrong - and it is not a bare PASS
#: either, because the report must not claim cover it does not have.
PASS_WITH_SKIPS = "PASS WITH SKIPS"

#: WHY a module did not finish, when it did not. This is the difference between
#: "the application is broken", "the environment could not support the row" and
#: "the suite itself fell over", and the report has to keep the three apart -
#: they go to three different people.
DATA_ENVIRONMENT = "DATA / ENVIRONMENT"
PRECONDITION = "BLOCKED PRECONDITION"
APPLICATION = "APPLICATION"
AUTOMATION = "AUTOMATION"

#: The categories that are NOT a defect of the product or of the suite: the row
#: could not be run, so nothing about the application was proved either way.
BLOCKING_CATEGORIES = (DATA_ENVIRONMENT, PRECONDITION)

#: What check() says when both sides turn out to be the same reading.
SAME_SOURCE = ("Expected and Actual were both read from the SAME screen "
               "reading, so this check would have compared the application "
               "with itself and passed whatever the application had shown. It "
               "proves nothing and is not counted as a pass.")


# --------------------------------------------------------------------------- #
# THE HOST'S HELPERS
#
# Same wiring as data_verification: this module must NOT import
# Construction_Flow (a circle, and under `python Construction_Flow.py` the suite
# is __main__, so importing it by name would build a second copy of everything).
# The helpers are injected with bind() instead, and until they are, the
# fallbacks below keep the module usable on its own - which is what lets it be
# unit-tested without a browser.
# --------------------------------------------------------------------------- #

try:
    import allure
    from allure_commons.types import AttachmentType
    _ALLURE = True
except ImportError:                                        # pragma: no cover
    _ALLURE = False


def _default_attach_text(name: str, body: str) -> None:
    if _ALLURE:
        allure.attach(body, name=name, attachment_type=AttachmentType.TEXT)


def _default_screenshot(page, name: str) -> None:
    return None


@contextmanager
def _default_step(title: str) -> Iterator[None]:
    log.info("%s", title)
    if _ALLURE:
        with allure.step(title):
            yield
    else:
        yield


class _Host:
    step = staticmethod(_default_step)
    attach_text = staticmethod(_default_attach_text)
    screenshot = staticmethod(_default_screenshot)


HOST = _Host()


def bind(**helpers) -> None:
    """Hand this module the framework's own step/attach/screenshot helpers."""
    for name, helper in helpers.items():
        if helper is not None:
            setattr(HOST, name,
                    staticmethod(helper) if callable(helper) else helper)


# --------------------------------------------------------------------------- #
# READING NUMBERS OFF THE PAGE
#
# The totals panels of this application are CSS grids with no table, no ids and
# no labels tying a caption to its figure - the same trap the line-item row set
# (see THE LINE-ITEM ROW in Construction_Flow.py). Anything that walks the DOM
# for them breaks on the next re-style.
#
# So the page is read as TEXT, once, and the figures are found by the words
# printed beside them. That is what a person reading the screen does, it does
# not care how the grid is built, and it costs one round trip instead of one per
# label.
# --------------------------------------------------------------------------- #

#: A number as the application prints it: 1,95,615.00 / ₹1,95,615 / 18% / -250.
#: The sign may stand on either side of the currency symbol - this application
#: writes a deduction as "-₹5,121.20" - see data_verification._NUMBER_LIKE.
_MONEY = r"[-+]?[₹$€£\s]*[-+]?[\d][\d,]*(?:\.\d+)?\s*%?"

#: Every screen reading taken this run gets its own number, so two figures can
#: be told apart even when they come from the same page at different moments.
_READINGS = 0


class Displayed(float):
    """A figure READ OFF a screen, remembering which reading it came from.

    This is what makes "expected and actual must come from different places"
    enforceable rather than merely asked for. A figure the framework works out
    is a plain float; a figure lifted off the page is one of these, tagged with
    the reading it came from - and arithmetic on it produces a plain float
    again, so a value the framework has actually DONE something with is no
    longer marked as a bare screen reading.

    check() then refuses the one case that proves nothing: both sides being the
    same figure, out of the same reading of the same screen. A figure published
    by an EARLIER module carries an earlier reading's tag, so cross-module
    checks - which are exactly the point - still work.
    """
    source: str

    def __new__(cls, value: float, source: str) -> "Displayed":
        reading = super().__new__(cls, value)
        reading.source = source
        return reading


class PageNumbers:
    """Every figure on one screen, addressed by the words printed beside it.

    Snapshot, not a live view: the text is read once and then queried, so a
    panel that re-renders between two lookups cannot give two different answers
    to the same question. Re-read the page to get a fresh one.
    """

    def __init__(self, text: str = "", url: str = "") -> None:
        global _READINGS
        _READINGS += 1
        self.text = text or ""
        self.url = url
        #: What every figure out of THIS reading is tagged with.
        self.source = f"{url}#reading-{_READINGS}"
        self.lines = [clean_text(line) for line in self.text.splitlines()]
        self.lines = [line for line in self.lines if line]

    @classmethod
    def read(cls, page) -> "PageNumbers":
        """Snapshot whatever is on screen now. Never raises - {} on a dead page."""
        try:
            return cls(page.inner_text("body"), page.url)
        except Exception as error:                          # page closed, nav
            log.warning("   could not read the page for calculations: %s",
                        str(error).splitlines()[0])
            return cls("", getattr(page, "url", ""))

    def find(self, *labels: str, near: int = 2) -> Optional[float]:
        """The figure printed for `labels`, the first one that is on screen.

        Two shapes are handled, because the application uses both:

            Subtotal            ₹1,95,615.00      <- caption and figure together
            Subtotal                              <- caption alone, the figure
            ₹1,95,615.00                             on one of the next lines

        `near` is how many following lines may be searched for the figure of a
        caption that stands on its own. Kept small on purpose: a bigger window
        starts stealing the NEXT caption's number.

        A label only ever matches its OWN caption. "Tax" does not answer for
        "Tax Rate", and "Total" does not answer for "Total Tax" - if what
        follows the label is more words, this is a different caption and the
        search moves on. Without that rule the tax RATE (18) was being read as
        the tax AMOUNT on a panel that prints both, which fails a document whose
        arithmetic is perfectly correct.
        """
        for label in labels:
            wording = re.compile(rf"^\W*{label}\b\s*[:\-]?\s*(.*)$", re.IGNORECASE)
            for index, line in enumerate(self.lines):
                match = wording.match(line)
                if not match:
                    continue
                rest = match.group(1).strip()
                if rest[:1].isalpha():
                    continue                   # "Tax" against "Tax Rate"
                value = _last_number(rest)
                if value is not None:
                    return Displayed(value, self.source)
                # The caption is on a line of its own - the figure is below it.
                for following in self.lines[index + 1:index + 1 + near]:
                    if _is_only_a_number(following):
                        return Displayed(as_number(following), self.source)
                    # A line that is another caption ends the search: the figure
                    # belonging to this one is simply not on screen.
                    if re.search(r"[A-Za-z]{3}", following):
                        break
        return None

    def found(self, *labels: str) -> bool:
        return self.find(*labels) is not None


def _last_number(text: str) -> Optional[float]:
    """The last figure in a piece of text - the value, when a caption leads."""
    found = re.findall(_MONEY, text or "")
    for candidate in reversed(found):
        value = as_number(candidate)
        if value is not None:
            return value
    return None


def _is_only_a_number(line: str) -> bool:
    """True when a line holds a figure and nothing else worth reading."""
    line = clean_text(line)
    return bool(line) and as_number(line) is not None


def magnitude(value: Optional[float]) -> Optional[float]:
    """A screen reading without its sign, still tagged with where it came from.

    For the one figure this application prints as a deduction: the totals panel
    writes the discount as "-₹5,121.20" while the discount itself is 5,121.20.
    Comparing the two as they stand would fail a document that is perfectly
    correct, and plain `abs()` would drop the Displayed tag - which is what
    stops an expected value being compared with itself. So the tag is carried
    over deliberately here rather than lost by accident.
    """
    if value is None:
        return None
    if isinstance(value, Displayed):
        return Displayed(abs(float(value)), value.source)
    return abs(float(value))


def money(value: Optional[float]) -> str:
    """A figure as the report should print it, or a dash when there is none."""
    if value is None:
        return "-"
    return f"{value:,.2f}"


# --------------------------------------------------------------------------- #
# ONE CHECK, ONE MODULE, ONE RUN
# --------------------------------------------------------------------------- #

class _CheckFailed(AssertionError):
    """Raised inside a failed check's own step, and caught the moment it leaves.

    Its only job is to colour that step red in the report - an AssertionError
    rather than a plain Exception because Allure reads the first as a FAILED
    step and the second as a BROKEN one, and a wrong number is a failure, not a
    crash. It never escapes _record().
    """


# --------------------------------------------------------------------------- #
# WHAT A CALCULATION IS ABOUT - the subject line of every check
#
# "Quotations | Line Amount | Quantity x Rate | 63,513.45 | 63,513.45" is a
# correct report and an unreadable one: a reviewer cannot tell WHICH item was
# quoted, how many of them, or at what rate, without opening the attachment
# underneath. Those figures were always recorded - they are the check's own
# `inputs`, and the row that produced them - they were simply never lifted into
# the tables and the step titles where a reader looks first.
#
# So every check now carries a SUBJECT: the item, the quantity, the rate, the
# discount and the tax rate it was about. NOTHING here takes part in the
# arithmetic. It is read from two places that already exist, in this order:
#
#   1. the check's OWN inputs      - the values that really went into this sum
#   2. the workbook row            - the test data the module was given
#
# and whatever neither of them holds is left blank rather than guessed at. The
# `inputs` dict is not touched, because qa_report.substituted() builds the
# "worked out as" line from it and adding a label there would change that line.
# --------------------------------------------------------------------------- #

#: What a sheet calls each field. Several spellings because every sheet is named
#: after its own screen: a Quotation row says "Item Name" and a Sales Order row
#: says "Product" for the same thing, and a Purchase Order row carries an "Item
#: Code" as well. First spelling that the row has, wins.
SUBJECT_COLUMNS: Dict[str, Tuple[str, ...]] = {
    "Item": ("Item Name", "Product", "Material", "Item"),
    "Item Code": ("Item Code", "SKU"),
    "Quantity": ("Quantity", "Qty"),
    "Rate": ("Rate", "Unit Price"),
    "Discount %": ("Discount %", "Discount"),
    "Tax %": ("Tax %", "GST %", "Tax"),
}

#: The keyword a caller of subject() uses -> the label the report prints. So the
#: suite can write `calc.subject(item=..., item_code=...)` and every module ends
#: up with the same column headings whether its subject came from the workbook
#: or was handed over by the flow.
SUBJECT_FIELDS: Dict[str, str] = {
    "item": "Item",
    "item_code": "Item Code",
    "quantity": "Quantity",
    "rate": "Rate",
    "discount_pct": "Discount %",
    "tax_pct": "Tax %",
}

#: How the same five things are recognised in a CHECK's own inputs, whose labels
#: are written for a reader rather than for a parser: "Tax % (the rate the
#: application's own form showed)", "Rate on the purchase order", "Purchase
#: Order Quantity". Order matters - the percentages are claimed before the plain
#: Rate and Quantity, so a "Tax %" label cannot be read as a rate.
#:
#: Rate is matched case-SENSITIVELY on purpose: several input labels explain
#: themselves with the lower-case word "rate" in a parenthesis ("(the rate the
#: application charged)"), and those are prose, not the field.
_INPUT_PATTERNS: Tuple[Tuple[str, object], ...] = (
    ("Item", re.compile(r"^(?:Item|Product|Material)\b")),
    ("Tax %", re.compile(r"^(?:Tax|GST)\s*%")),
    ("Discount %", re.compile(r"^Discount\s*%")),
    ("Quantity", re.compile(r"\b(?:Quantity|Qty)\b", re.IGNORECASE)),
    ("Rate", re.compile(r"\bRate\b")),
)

#: WHEN a percentage may be taken off the workbook row instead of the check's
#: own inputs: only when this calculation is about that percentage, which its
#: FORMULA is what says. The item, the quantity and the rate carry no such rule
#: - they identify what the sum was about, and that is always worth printing.
#:
#: Tax is deliberately stricter than discount: it falls back only for a formula
#: about a tax RATE ("Taxable Value x Tax% / 100"), never one about a tax
#: AMOUNT ("IGST = Tax"). The workbook's Tax column says 28 for cement and this
#: application charges the 18% its HSN code carries - a difference the report
#: states in one place, on purpose, and must not repeat as a stray column.
_ROW_FALLBACK_WHEN: Dict[str, object] = {
    "Discount %": re.compile(r"discount", re.IGNORECASE),
    "Tax %": re.compile(r"(?:tax|gst)\s*%", re.IGNORECASE),
}


def subject_of(row) -> Dict[str, str]:
    """The item, quantity, rate, discount and tax a workbook row is about.

    Read straight off the row the module was handed - no lookup, no default and
    no invention: a sheet without a Rate column simply has no rate to show, and
    the report says nothing rather than something plausible.

    Anything dict-like will do (excel_reader.Row is a dict), and a row that is
    not there at all gives {} - this module is never allowed to fall over on
    behalf of the report.
    """
    context: Dict[str, str] = {}
    if row is None:
        return context
    try:
        for field_name, columns in SUBJECT_COLUMNS.items():
            for column in columns:
                value = clean_text(row.get(column, ""))
                if value:
                    context[field_name] = value
                    break
    except AttributeError:                       # not dict-like: nothing to read
        return {}
    return context


@dataclass
class Check:
    """One calculation: what went in, what was expected, what the app showed."""
    module: str
    name: str
    inputs: Dict[str, object] = field(default_factory=dict)
    formula: str = ""
    expected: Optional[float] = None
    actual: Optional[float] = None
    tolerance: float = MONEY_TOLERANCE
    status: str = SKIPPED
    note: str = ""
    #: The workbook row's own subject - item, quantity, rate, discount, tax.
    #: Display only: it is never read by the comparison. See subject_of().
    context: Dict[str, str] = field(default_factory=dict)

    @property
    def difference(self) -> Optional[float]:
        if self.expected is None or self.actual is None:
            return None
        return round(self.actual - self.expected, 6)

    # --- the subject of this calculation, for the tables and the step title --
    def sourced(self, field_name: str) -> Tuple[str, str]:
        """One subject field as (value, where it came from).

        The check's OWN inputs win, because they are the values that really
        went into THIS sum: a GRN's received quantity is the receipt's figure,
        not the quantity the workbook asked for, and printing the second in
        place of the first would be quietly relabelling one number as another.
        The workbook row is the fallback, and it says so.

        A percentage only falls back to the row when this calculation is
        actually ABOUT that percentage - see _ROW_FALLBACK_WHEN. Without that
        rule a Line Amount check would have carried "Tax %: 28" off its row
        while the application had charged 18% from the item's HSN code, and the
        two figures sitting in one table would have read as a contradiction
        where there is none.
        """
        for wanted, pattern in _INPUT_PATTERNS:
            if wanted != field_name:
                continue
            for label, value in self.inputs.items():
                if pattern.search(str(label)) and str(value or "").strip():
                    return _show(value), "used by this calculation"

        needed = _ROW_FALLBACK_WHEN.get(field_name)
        if needed is not None and not needed.search(self.formula or ""):
            return "", ""
        value = self.context.get(field_name, "")
        return (value, "from the test data row") if value else ("", "")

    def about(self, field_name: str) -> str:
        """One subject field's value, or "" when this check has no such field."""
        return self.sourced(field_name)[0]

    @property
    def item(self) -> str:
        """What this calculation is about - the item, with its code when known."""
        name = self.about("Item")
        code = self.context.get("Item Code", "")
        if name and code and code.casefold() not in name.casefold():
            return f"{name} ({code})"
        return name or code

    @property
    def quantity(self) -> str:
        return self.about("Quantity")

    @property
    def rate(self) -> str:
        return self.about("Rate")

    @property
    def discount_pct(self) -> str:
        return self.about("Discount %")

    @property
    def tax_pct(self) -> str:
        return self.about("Tax %")

    @property
    def subject_line(self) -> str:
        """"Item: OPC Cement 53 Grade | Qty: 155 | Rate: 413" - or "" if unknown."""
        return " | ".join(f"{label}: {value}" for label, value in
                          (("Item", self.item), ("Qty", self.quantity),
                           ("Rate", self.rate), ("Disc %", self.discount_pct),
                           ("Tax %", self.tax_pct)) if value)

    @property
    def inputs_text(self) -> str:
        """Every input this check used, on one line: "Quantity=155, Rate=413"."""
        return ", ".join(f"{label}={_show(value)}"
                         for label, value in self.inputs.items()
                         if str(value or "").strip() != "")

    def at_a_glance(self) -> List[List[object]]:
        """The whole check as the label/value list a QA reviewer reads first.

        Everything a reviewer asks of one calculation, in the order they ask
        it, and every line read back off what the run already recorded.
        """
        facts: List[List[object]] = [["Module", self.module],
                                     ["Calculation", self.name]]
        # WHERE each figure came from, beside the figure. A quantity the check
        # itself used and a quantity read off the workbook row are two
        # different claims, and a card that printed them identically would be
        # inviting a reviewer to treat them as one.
        for label, field_name in (("Item / Product", "Item"),
                                  ("Quantity", "Quantity"),
                                  ("Rate", "Rate"),
                                  ("Discount %", "Discount %"),
                                  ("Tax %", "Tax %")):
            value, source = self.sourced(field_name)
            if label == "Item / Product":
                value = self.item
            if value:
                facts.append([label, f"{value}  ({source})" if source else value])
        worked = self.worked_out
        facts += [
            ["Input values used", self.inputs_text or "(none recorded)"],
            ["Formula", self.formula or "(not stated)"],
        ]
        if worked and self.expected is not None:
            facts.append(["Worked out as", f"{worked} = {money(self.expected)}"])
        facts += [
            ["Expected", money(self.expected)],
            ["Actual (application)", money(self.actual)],
            ["Difference", money(self.difference)],
            ["Result", self.status],
            ["Reason", self.reason],
        ]
        return facts

    @property
    def headline(self) -> str:
        """The one line a step title carries: what, about what, how much.

        "Line Amount: PASS | Item: OPC Cement 53 Grade | Qty: 155 |
         Rate: 413.00 | Expected 58,893.80 | Actual 58,893.80"

        This is the single most-read string in the whole report - it is what
        Allure prints in the step tree, where a reviewer scans a module without
        opening anything at all.
        """
        parts = [f"{self.name}: {self.status}"]
        if self.subject_line:
            parts.append(self.subject_line)
        parts.append(f"Expected {money(self.expected)}")
        parts.append(f"Actual {money(self.actual)}")
        return " | ".join(parts)

    @property
    def worked_out(self) -> str:
        """The formula with THIS row's own numbers in it - "567 x 345".

        The line a manual tester writes underneath the rule when they check it
        by hand. Empty when the formula is a sentence rather than an expression,
        or when the labels in it are not the ones the inputs carry: a
        substitution nobody can follow is worse than none - see
        qa_report.substituted().
        """
        return qa_report.substituted(self.formula, self.inputs)

    def report(self) -> str:
        """What a calculation review asks for, in the order it asks for it.

        Module, calculation, the inputs, the formula, the formula with this
        row's numbers in it, Expected, Actual, Difference, Result - and, when
        something did not add up, the note that says which other convention the
        application's figure WOULD have matched.
        """
        lines = [f"MODULE: {self.module}", "", f"Calculation: {self.name}", ""]
        # WHAT THIS SUM IS ABOUT, before the sum itself. The item, the quantity
        # and the rate used to be legible only by reading the input list below
        # and knowing which label meant what.
        subject = [(label, value) for label, value in
                   (("Item / Product", self.item), ("Quantity", self.quantity),
                    ("Rate", self.rate), ("Discount %", self.discount_pct),
                    ("Tax %", self.tax_pct)) if value]
        if subject:
            lines.append("Calculated for:")
            lines += [f"    {label:<16}: {value}" for label, value in subject]
            lines.append("")
        if self.inputs:
            lines.append("Input Values:")
            for label, value in self.inputs.items():
                lines.append(f"    {label}: {_show(value)}")
            lines.append("")
        if self.formula:
            lines += ["Formula:", f"    {self.formula}", ""]
        worked = self.worked_out
        if worked and self.expected is not None:
            lines += ["Worked out as:",
                      f"    {worked} = {money(self.expected)}", ""]
        lines += [
            "Expected:", f"    {money(self.expected)}", "",
            "Actual (what the application displayed):",
            f"    {money(self.actual)}", "",
            "Difference:", f"    {money(self.difference)}", "",
            "Result:", f"    {self.status}",
        ]
        if self.note:
            lines += ["", "Observation:", f"    {self.note}"]
        return "\n".join(lines)

    def page(self) -> str:
        """The same calculation as the HTML card the report shows.

        Three blocks, in the order a reviewer works through one calculation:
        WHAT it was about and how it came out (at a glance), WHICH values went
        into it, and the formula with this row's own numbers in it.
        """
        worked = self.worked_out
        return qa_report.page(
            f"Calculation - {self.name}",
            " | ".join(part for part in
                       (self.module, f"Item: {self.item}" if self.item else "",
                        f"Qty: {self.quantity}" if self.quantity else "",
                        f"Rate: {self.rate}" if self.rate else "") if part),
            [qa_report.facts_block("At a glance", self.at_a_glance()),
             qa_report.facts_block(
                 "Input values used for this calculation",
                 [[label, _show(value)] for label, value in self.inputs.items()]
                 or [["(none recorded)", ""]]),
             qa_report.pre_block("Formula", "\n".join(part for part in (
                 self.formula or "(not stated)",
                 f"{worked} = {money(self.expected)}"
                 if worked and self.expected is not None else "") if part)),
             qa_report.table_block(
                 "Validation", ("Expected", "Actual (application)",
                                "Difference", "Result"),
                 [[money(self.expected), money(self.actual),
                   money(self.difference), self.status]]),
             qa_report.note_block("Observation", self.note)],
            verdict=self.status)

    #: The columns of the calculation tables, in the order a QA review asks for
    #: them. ITEM, QTY, RATE, DISC % and TAX % stand in columns of their own
    #: rather than inside the formula: "Quantity x Rate" tells a reviewer the
    #: rule, and only "155 x 413.00" tells them what was actually checked.
    TABLE_HEADINGS: Tuple[str, ...] = (
        "Module", "Calculation", "Item / Product", "Qty", "Rate", "Disc %",
        "Tax %", "Inputs used", "Formula", "Expected", "Actual", "Difference",
        "Result", "Reason")

    #: Where the verdict sits in TABLE_HEADINGS - stated, because Reason comes
    #: after Result and the renderer's default is the last column.
    TABLE_VERDICT_COLUMN = 12

    def as_table_row(self) -> List[str]:
        """One line of the "every calculation checked" table.

        The REASON is a column of its own and not an afterthought: a row that
        says only "SKIPPED" tells a reader that something was not checked and
        leaves them to guess whether that was the application, the data or the
        automation. Every skip in this framework carries its reason, so the
        table prints it.
        """
        return [self.module, self.name, _fit(self.item, 40), self.quantity,
                self.rate, self.discount_pct, self.tax_pct,
                _fit(self.inputs_text, 70), self.formula or "-",
                money(self.expected), money(self.actual),
                money(self.difference), self.status, self.short_reason]

    @property
    def reason(self) -> str:
        """Why this check came out the way it did - "-" when it simply passed."""
        return self.note or ("-" if self.status == PASS else "(no reason "
                                                            "recorded)")

    @property
    def short_reason(self) -> str:
        """The reason, cut to a width a fixed-width table can still be read at.

        The plain-text tables are laid out on the widest cell in the column, so
        one three-line reason would stretch every row of the report to the same
        width. The whole sentence is always in this check's own attachment
        (report()), which is where a reader goes for the detail.
        """
        return _fit(self.reason, 96)


def _show(value: object) -> str:
    """An input value as the report prints it."""
    if isinstance(value, float):
        return f"{value:,.4f}".rstrip("0").rstrip(".")
    return str(value)


def _fit(text: str, width: int) -> str:
    """A cell cut to a width the fixed-width tables can still be read at.

    The plain-text tables lay every column out on its widest cell, so one long
    input list would stretch the whole report to the same width. The full text
    is always on the check's own card, which is where the detail belongs.
    """
    text = " ".join(str(text or "").split())
    return text if len(text) <= width else text[:width - 3].rstrip() + "..."


@dataclass
class ModuleReport:
    """Everything one Excel row proved, or failed to prove, about one module."""
    module: str
    test_case_id: str = ""
    functional: str = SKIPPED
    checks: List[Check] = field(default_factory=list)
    note: str = ""
    #: WHY the functional half did not finish - one of DATA_ENVIRONMENT,
    #: PRECONDITION, APPLICATION, AUTOMATION. Empty when nothing went wrong.
    category: str = ""
    #: What this row was about - item, quantity, rate, discount, tax - read off
    #: the workbook row by subject_of(). Display only; every check copies it.
    context: Dict[str, str] = field(default_factory=dict)

    @property
    def calculation(self) -> str:
        """The verdict over THIS MODULE'S checks, and nothing else.

        Reads only `self.checks`. Whether the module's FLOW worked is a separate
        question with a separate answer (`functional`), and letting one decide
        the other is what made a run with no wrong number anywhere report its
        arithmetic as failed.
        """
        if any(check.status == FAIL for check in self.checks):
            return FAIL
        passed = any(check.status == PASS for check in self.checks)
        skipped = any(check.status == SKIPPED for check in self.checks)
        if passed and skipped:
            return PASS_WITH_SKIPS
        if passed:
            return PASS
        if skipped:
            return SKIPPED
        return NOT_APPLICABLE

    @property
    def overall(self) -> str:
        """Functional and calculation together - the module's real result.

        A wrong number still fails a module whose flow worked, and a module that
        never ran is BLOCKED rather than failed - a row the environment could
        not support proves nothing about the application, in either direction.
        """
        if self.calculation == FAIL:
            return FAIL
        if self.functional in (BLOCKED, NOT_AUTOMATED):
            return self.functional
        if self.functional == FAIL:
            return FAIL
        if self.functional == SKIPPED:
            return SKIPPED
        return PASS

    def counts(self) -> Dict[str, int]:
        totals = {PASS: 0, FAIL: 0, SKIPPED: 0}
        for check in self.checks:
            totals[check.status] = totals.get(check.status, 0) + 1
        return totals


class CalculationValidator:
    """The one calculation recorder for the run. Construction_Flow holds one."""

    def __init__(self) -> None:
        self.reports: List[ModuleReport] = []
        self.current: Optional[ModuleReport] = None

    # --- one Excel row ----------------------------------------------------
    def start_test(self, module: str, row=None) -> None:
        test_case_id = ""
        if row is not None:
            test_case_id = getattr(row, "test_case_id", "") or ""
        # The row's subject - which item, how many, at what rate - captured
        # once here and copied onto every check the row makes, so the tables
        # and the step titles can name what each sum was about. Read only; it
        # takes no part in any comparison. See subject_of().
        self.current = ModuleReport(module=module, test_case_id=test_case_id,
                                    context=subject_of(row))
        self.reports.append(self.current)

    def subject(self, **fields: object) -> None:
        """Add to what the CURRENT row's calculations are said to be about.

        For the modules whose sheet does not carry the subject: a GRN row names
        a purchase order and a Purchase Bill row a bill number, but the item
        they move is the purchase order's, and only the suite knows it. This is
        how the suite hands it over.

        Display only, exactly like the rest of the context - it cannot reach any
        expected or actual value. A blank is ignored rather than stored, so a
        figure the run could not read stays absent from the report instead of
        appearing as an empty column.
        """
        if self.current is None:
            return
        for label, value in fields.items():
            # _show(), not str(): a quantity handed over as a float would
            # otherwise print as "280.0" beside the workbook's own "280" and
            # read as a different number.
            text = clean_text(_show(value) if isinstance(value, float)
                              else value)
            if text:
                self.current.context[SUBJECT_FIELDS.get(label, label)] = text

    def functional(self, status: str, category: str = "") -> None:
        """Record whether the module's own flow worked, and why it did not."""
        if self.current is not None:
            self.current.functional = status
            if category:
                self.current.category = category

    def note(self, text: str) -> None:
        if self.current is not None:
            self.current.note = text

    # --- one calculation --------------------------------------------------
    def check(self, name: str, inputs: Dict[str, object], formula: str,
              expected: Optional[float], actual: Optional[float],
              tolerance: float = MONEY_TOLERANCE,
              alternatives: Optional[Dict[str, float]] = None,
              page=None) -> bool:
        """Compare one independently calculated value with the application's.

        Returns True when they agree. Everything - the inputs, the formula, both
        values, the difference and the verdict - goes into the Allure report as
        a step of its own, whichever way it goes.
        """
        record = Check(module=self._module(), name=name, inputs=dict(inputs),
                       formula=formula, expected=expected, actual=actual,
                       tolerance=tolerance, context=self._context())

        if expected is None:
            record.status = SKIPPED
            record.note = ("The expected value could not be worked out - an "
                           "input it needs was not available.")
        elif actual is None:
            record.status = SKIPPED
            record.note = ("The application did not display this figure on "
                           "this screen, so there was nothing to compare with. "
                           "Nothing was assumed about it.")
        elif _same_source(expected, actual):
            record.status = SKIPPED
            record.note = SAME_SOURCE
        elif abs(actual - expected) <= tolerance:
            record.status = PASS
        else:
            record.status = FAIL
            record.note = _explain_mismatch(actual, alternatives, tolerance)

        self._record(record, page)
        return record.status == PASS

    def skip(self, name: str, reason: str,
             inputs: Optional[Dict[str, object]] = None,
             formula: str = "") -> None:
        """Record a calculation that could NOT be verified, and why.

        This is what keeps the report honest: an unverified calculation is
        visible as unverified, never quietly missing and never counted as a
        pass.
        """
        self._record(Check(module=self._module(), name=name,
                           inputs=dict(inputs or {}), formula=formula,
                           status=SKIPPED, note=reason,
                           context=self._context()))

    def not_applicable(self, reason: str) -> None:
        """Record that this module has no calculations to check."""
        self.note(reason)
        with HOST.step(f"Calculation Validation: {NOT_APPLICABLE}"):
            log.info("   %s", reason)
            HOST.attach_text("Calculation Validation - Not Applicable", reason)

    # --- the row is over --------------------------------------------------
    def finish_test(self, functional: Optional[str] = None,
                    category: str = "", note: str = "") -> Optional[ModuleReport]:
        """Close the row, attach its module result, and hand it back."""
        report = self.current
        self.current = None
        if report is None:
            return None
        if functional is not None:
            report.functional = functional
        if category:
            report.category = category
        if note and not report.note:
            report.note = note
        if report.checks or report.functional in (FAIL, BLOCKED):
            self._attach_module_result(report)
        return report

    def mismatches(self, report: Optional[ModuleReport]) -> List[Check]:
        if report is None:
            return []
        return [check for check in report.checks if check.status == FAIL]

    # --- the run is over --------------------------------------------------
    def summary_rows(self) -> List[List[str]]:
        """The module table: Module | Functional | Calculation | Overall | why."""
        rows: List[List[str]] = []
        for report in self._merged():
            counts = report.counts()
            checks = (f"{counts[PASS]}/{sum(counts.values())} passed"
                      if report.checks else "-")
            rows.append([report.module, report.functional, checks,
                         report.calculation, report.overall,
                         _reason(report)])
        return rows

    # --- the four verdicts, each answering its OWN question ----------------
    @staticmethod
    def calculation_result(passed: int, failed: int, skipped: int) -> str:
        """The arithmetic verdict, from the CHECK counts and nothing else.

        No module status is looked at here, on purpose. A blocked flow means a
        calculation was never reached, not that a calculation was wrong, and the
        one used to be reported as the other: a run in which every single figure
        the application produced agreed with an independently worked-out value
        announced "Overall Calculation Result: FAIL" because a different module
        had been blocked. Nothing was wrong with the numbers, and the report
        sent people looking for a wrong one.
        """
        if failed > 0:
            return FAIL                      # a real mismatch, and only that
        if passed > 0 and skipped > 0:
            return PASS_WITH_SKIPS
        if passed > 0:
            return PASS
        if skipped > 0:
            return SKIPPED                   # nothing proved, nothing wrong
        return NOT_APPLICABLE

    @staticmethod
    def functional_result(failed: int, blocked: int, passed: int,
                          skipped: int) -> str:
        """The flow verdict, from the MODULE statuses and nothing else."""
        if failed > 0:
            return FAIL
        if blocked > 0:
            return BLOCKED
        if passed > 0 and skipped > 0:
            return PASS_WITH_SKIPS
        if passed > 0:
            return PASS
        if skipped > 0:
            return SKIPPED
        return NOT_AUTOMATED

    def totals(self) -> Dict[str, object]:
        """Every number and every verdict in the summary, counted from the run.

        One place works these out, so the headline, the tables, the Allure
        attachments and the pass/fail of the summary test can never disagree
        with each other. Nothing here is written in by hand.
        """
        merged = self._merged()
        checks = [check for report in merged for check in report.checks]

        checks_passed = sum(1 for c in checks if c.status == PASS)
        checks_failed = sum(1 for c in checks if c.status == FAIL)
        checks_skipped = sum(1 for c in checks if c.status == SKIPPED)

        modules_passed = sum(1 for r in merged if r.overall == PASS)
        modules_failed = sum(1 for r in merged if r.overall == FAIL)
        modules_blocked = sum(1 for r in merged if r.overall == BLOCKED)
        modules_not_automated = sum(1 for r in merged
                                    if r.overall == NOT_AUTOMATED)
        modules_skipped = sum(1 for r in merged if r.overall == SKIPPED)

        calculation = self.calculation_result(checks_passed, checks_failed,
                                              checks_skipped)
        # NOT AUTOMATED is not a blocker: nobody claimed to run it. It is
        # counted on its own line so the coverage the report claims stays
        # honest without dragging the functional verdict down with it.
        functional = self.functional_result(modules_failed, modules_blocked,
                                            modules_passed, modules_skipped)

        # The project verdict is the two above, taken together - which is NOT
        # the same as either of them being allowed to overwrite the other.
        if calculation == FAIL or functional == FAIL:
            project = FAIL
        elif functional == BLOCKED:
            project = BLOCKED
        elif functional == SKIPPED and calculation in (SKIPPED, NOT_APPLICABLE):
            project = SKIPPED
        else:
            project = PASS

        return {
            "merged": merged, "checks": checks,
            "modules_total": len(merged),
            "modules_passed": modules_passed,
            "modules_failed": modules_failed,
            "modules_blocked": modules_blocked,
            "modules_not_automated": modules_not_automated,
            "modules_skipped": modules_skipped,
            "checks_total": len(checks),
            "checks_passed": checks_passed,
            "checks_failed": checks_failed,
            "checks_skipped": checks_skipped,
            "calculation_result": calculation,
            "functional_result": functional,
            "project_result": project,
            # EVERY blocked module, whatever category it carries. The list used
            # to be filtered to BLOCKING_CATEGORIES, so a module blocked for
            # any other reason was counted in the headline and then missing
            # from the section that explains the headline.
            "blockers": [r for r in merged if r.overall == BLOCKED],
            # A blocked module is never a defect of either kind - it was never
            # run, so it cannot have proved anything about the application or
            # about the suite. Stated here as well as implied by the category,
            # because these two lists are the ones a reader acts on.
            "automation_defects": [r for r in merged
                                   if r.category == AUTOMATION
                                   and r.overall != BLOCKED],
            "application_defects": [r for r in merged
                                    if r.category == APPLICATION
                                    and r.overall != BLOCKED],
        }

    def summary(self, environment: Sequence[Sequence[object]] = (),
                narrative: str = "",
                test_cases: Optional[Dict[str, int]] = None) -> Tuple[str, str,
                                                                     str]:
        """(overall PROJECT result, the plain-text summary, the HTML one).

        The project result is what comes back, not the calculation one - use
        totals() when a caller needs the halves apart.

        The three optional arguments are what the SUITE knows and this module
        does not, and they are how the execution summary comes out as ONE report
        instead of two half ones:

            environment   Execution Date / Time / Environment / Application URL
            narrative     the run in two or three sentences of QA English
            test_cases    how many test cases passed, failed and were skipped
                          (module counts and calculation counts are already
                          worked out here; TEST CASE counts are pytest's own -
                          see qa_report.TALLY)

        All three default to empty, so an existing `calc.summary()` call is
        unchanged.
        """
        t = self.totals()
        checks = t["checks"]

        headline = [
            "COMPLETE PROJECT VALIDATION SUMMARY",
            "",
            f"Total Modules Checked      : {t['modules_total']}",
            f"Modules Passed             : {t['modules_passed']}",
            f"Modules Failed             : {t['modules_failed']}",
            f"Modules Blocked            : {t['modules_blocked']}",
            f"Modules Not Automated      : {t['modules_not_automated']}",
            f"Modules Skipped            : {t['modules_skipped']}",
            "",
        ] + ([
            f"Total Test Cases           : {test_cases['total']}",
            f"Test Cases Passed          : {test_cases['passed']}",
            f"Test Cases Failed          : "
            f"{test_cases['failed'] + test_cases.get('broken', 0)}",
            f"Test Cases Blocked         : {test_cases.get('blocked', 0)}",
            f"Test Cases Skipped         : {test_cases['skipped']}",
            "",
        ] if test_cases else []) + [
            f"Total Calculations Checked : {t['checks_total']}",
            f"Calculations Passed        : {t['checks_passed']}",
            f"Calculations Failed        : {t['checks_failed']}",
            f"Calculations Skipped       : {t['checks_skipped']}",
            "",
            f"Overall Calculation Result : {t['calculation_result']}",
            f"Functional Result          : {t['functional_result']}",
            f"Overall Project Result     : {t['project_result']}",
            "",
            _verdict_note(t),
        ]

        module_table = text_table(
            ["Module", "Functional", "Checks", "Calculation", "Overall",
             "Reason"],
            self.summary_rows())
        check_table = text_table(
            list(Check.TABLE_HEADINGS),
            [check.as_table_row() for check in checks]) if checks else \
            "No calculations were executed."

        parts = []
        if environment:
            parts.append(qa_report.section("EXECUTION",
                                           qa_report.labelled(environment)))
        parts += ["\n".join(headline)]
        if narrative:
            parts.append(qa_report.section("IN PLAIN WORDS", narrative))
        parts += [qa_report.section("MODULE BY MODULE", module_table),
                  qa_report.section("EVERY CALCULATION CHECKED", check_table)]
        for title, reports in (
                ("BLOCKERS - MODULES THAT COULD NOT BE RUN", t["blockers"]),
                ("AUTOMATION DEFECTS", t["automation_defects"]),
                ("APPLICATION DEFECTS", t["application_defects"])):
            parts.append(_section(title, reports))
        parts.append(qa_report.section("HOW TO READ THIS REPORT",
                                       qa_report.CLASSIFICATION_NOTE))
        plain = "\n\n".join(part for part in parts if part)

        # ONE page, in the order a QA engineer reads it: where and when it ran,
        # the counts as tiles, the run in plain words, then module by module and
        # calculation by calculation. Every figure comes from totals(), which
        # counts what actually executed - nothing on this page is written in by
        # hand and nothing is recounted.
        tiles: List[List[object]] = [
            ["Modules", t["modules_total"]],
            ["Modules passed", t["modules_passed"],
             PASS if t["modules_passed"] else ""],
            ["Modules failed", t["modules_failed"],
             FAIL if t["modules_failed"] else ""],
            ["Modules blocked", t["modules_blocked"],
             BLOCKED if t["modules_blocked"] else ""],
            ["Not automated", t["modules_not_automated"], NOT_AUTOMATED],
        ]
        case_tiles: List[List[object]] = []
        if test_cases:
            failed_cases = test_cases["failed"] + test_cases.get("broken", 0)
            blocked_cases = test_cases.get("blocked", 0)
            case_tiles = [
                ["Test cases", test_cases["total"]],
                ["Passed", test_cases["passed"],
                 PASS if test_cases["passed"] else ""],
                # Failed and Blocked are two tiles and never one. A test case
                # that could not be RUN is not a failed test case, and adding
                # the two together is what made a blocked precondition read as
                # an application defect.
                ["Failed", failed_cases, FAIL if failed_cases else ""],
                ["Blocked", blocked_cases, BLOCKED if blocked_cases else ""],
                ["Skipped", test_cases["skipped"],
                 SKIPPED if test_cases["skipped"] else ""],
            ]

        blocks = [
            qa_report.facts_block("Execution", environment),
            qa_report.kpi_block("Modules", tiles),
            qa_report.kpi_block("Test cases", case_tiles),
            qa_report.kpi_block("Calculations", [
                ["Checked", t["checks_total"]],
                ["Passed", t["checks_passed"],
                 PASS if t["checks_passed"] else ""],
                ["Failed", t["checks_failed"],
                 FAIL if t["checks_failed"] else ""],
                ["Not verifiable", t["checks_skipped"],
                 SKIPPED if t["checks_skipped"] else ""],
            ]),
            qa_report.kpi_block("The three verdicts", [
                ["Calculation result", t["calculation_result"],
                 t["calculation_result"]],
                ["Functional result", t["functional_result"],
                 t["functional_result"]],
                ["Overall project result", t["project_result"],
                 t["project_result"]],
            ]),
            qa_report.note_block("In plain words",
                                 narrative or _verdict_note(t)),
            qa_report.table_block(
                "Module by module",
                ["Module", "Functional", "Checks", "Calculation", "Reason",
                 "Overall"],
                [[row[0], row[1], row[2], row[3], row[5], row[4]]
                 for row in self.summary_rows()]),
            qa_report.table_block(
                "Every calculation checked",
                list(Check.TABLE_HEADINGS),
                [check.as_table_row() for check in checks],
                note="Item, Qty, Rate, Disc % and Tax % are the values this "
                     "calculation actually used - the check's own inputs where "
                     "it has them, and the workbook row it belongs to "
                     "otherwise. Result is the verdict; Reason says why. A "
                     "SKIPPED line is a figure that could not be verified, and "
                     "the reason states which side of it was missing - never "
                     "that the value was assumed.",
                verdict_column=Check.TABLE_VERDICT_COLUMN),
        ]
        if t["blockers"]:
            blocks.append(qa_report.table_block(
                "Blockers - modules that could not be run",
                ["Module", "Category", "Result", "Why it could not be run"],
                [[r.module, r.category or "-", BLOCKED, r.note or "-"]
                 for r in t["blockers"]],
                note="A blocked module was never RUN, so nothing about the "
                     "application was proved there. It is not an application "
                     "defect, not an automation defect and not a pass. The "
                     "test case each one belongs to carries the full "
                     "precondition, root cause and what to do before the next "
                     "run.",
                verdict_column=2))
        for title, reports in (("Automation defects", t["automation_defects"]),
                               ("Application defects", t["application_defects"])):
            if reports:
                blocks.append(qa_report.table_block(
                    title, ["Module", "Category", "Reason"],
                    [[r.module, r.category or "-", _reason(r) or "-"]
                     for r in reports], verdict_column=99, legend=False))
        blocks.append(qa_report.pre_block("How to read this report",
                                          qa_report.CLASSIFICATION_NOTE))

        html = qa_report.page(
            "GenZOpss Construction Flow - QA Execution Summary",
            f"{t['modules_total']} modules | {t['checks_total']} calculations "
            f"| Calculation: {t['calculation_result']} | Functional: "
            f"{t['functional_result']}",
            blocks, verdict=t["project_result"])
        return t["project_result"], plain, html

    def declare(self, module: str, functional: str, note: str,
                category: str = "") -> None:
        """Record a module that did not run at all - blocked, or not automated.

        Used for the modules the application does not expose to this suite. They
        appear in the summary with their real status so the coverage the report
        claims is the coverage there is; they are never counted as passes.
        """
        self.reports.append(ModuleReport(module=module, functional=functional,
                                         note=note, category=category))

    # --- internals --------------------------------------------------------
    def _module(self) -> str:
        return self.current.module if self.current else "(no module)"

    def _context(self) -> Dict[str, str]:
        """A COPY of the row's subject, for one check to keep.

        A copy, because a later `subject()` call must not rewrite the subject of
        a check that has already been recorded - the report would then say a
        calculation was about something it was not.
        """
        return dict(self.current.context) if self.current else {}

    def _record(self, record: Check, page=None) -> None:
        if self.current is not None:
            self.current.checks.append(record)

        # A failed check has to LOOK failed in the report. Allure takes a step's
        # status from whether something was raised inside it, so a wrong figure
        # raises here and is caught immediately: the step goes red, carrying the
        # numbers, and the run carries on to check the rest of the module. What
        # fails the row is the verdict in run_test_case(), not this.
        try:
            # THE STEP TITLE IS THE REPORT MOST PEOPLE READ. Allure prints it in
            # the step tree, where a reviewer scans a whole module without
            # opening a single attachment - so it carries the subject as well as
            # the verdict: which item, how many, at what rate, expected against
            # actual. "Calculation - Line Amount: PASS" said none of that, and
            # the quantity and the rate were legible only inside the attachment
            # underneath it. Nothing new is computed for this line; it is the
            # figures the check already holds. See Check.headline.
            with HOST.step(f"Calculation - {record.headline}"):
                log.info("   %-28s expected %s | actual %s | %s", record.name,
                         money(record.expected), money(record.actual),
                         record.status)
                subject = record.subject_line
                if subject:
                    log.info("      for: %s", subject)
                if record.note:
                    log.info("      %s", record.note)
                # ONE attachment per check, and only for the checks that need
                # one. A passing sum's inputs, formula, expected, actual,
                # difference and result are ALL on its step title above and on
                # its row of the module's calculation table below - so a text
                # attachment per passing check was the same figures a third
                # time, and a module with fourteen checks arrived with fourteen
                # attachments nobody opened. A check that FAILED or that could
                # not be verified is the opposite case: its reason is the whole
                # point, and it is too long for a table cell.
                if record.status != PASS:
                    HOST.attach_text(f"Calculation - {record.name}",
                                     record.report())
                if record.status == FAIL:
                    # The colour card as well, and ONLY for a wrong number: it
                    # is the one a reader has to understand in a hurry, and an
                    # HTML copy of every passing sum would be a report nobody
                    # can find anything in. Every passing check's item,
                    # quantity, rate and inputs are on its step title, in its
                    # own text attachment above, and in the module's own
                    # calculation table - so nothing is lost by keeping this
                    # for the sums that went wrong.
                    attach_html(f"Calculation FAILED - {record.name}",
                                record.page())
                    if page is not None:
                        HOST.screenshot(
                            page, f"Screenshot - Calculation Failed - "
                                  f"{record.name}")
                    raise _CheckFailed(record.report())
        except _CheckFailed:
            pass

    #: The columns a calculation review asks for, in that order. One table for
    #: the whole module, so a reader sees every sum this row checked side by
    #: side instead of opening one attachment per check.
    #:
    #: The same columns as the run-wide table (Check.TABLE_HEADINGS) minus the
    #: Module one, because every row of THIS table is the same module. Derived
    #: rather than written out twice, so the two tables can never drift apart.
    TABLE_HEADINGS = Check.TABLE_HEADINGS[1:]
    TABLE_VERDICT_COLUMN = Check.TABLE_VERDICT_COLUMN - 1

    def _attach_module_result(self, report: ModuleReport) -> None:
        """The module's calculations: the table, and the detail of what broke.

        The "Module Result - <module>" attachment that used to sit here is
        gone. Every line of it was somewhere else already: the functional
        result and the category are in "Execution Status", the three counts and
        the calculation verdict are in this table's own summary line and in the
        HTML tiles below it, and each check's headline is its step title. It
        was a fourth copy of figures a reader had already been given three
        times, and it was the attachment that made the list look long.
        """
        counts = report.counts()
        if not report.checks:
            return

        # THE QA CARD FIRST. Allure lists a test's attachments in the order
        # they were produced, so this is what "first" means: the short,
        # business-readable answer - what was entered, what record was created,
        # every figure as expected / actual / difference / verdict, and the
        # module's result - before the calculation review underneath it.
        #
        # It reads the SAME Check objects the table below is drawn from, so the
        # two cannot disagree. Nothing here can fail a row: a card that could
        # not be built is skipped and every technical attachment still lands.
        try:
            with HOST.step(f"{qa_card.title(report.module)}: "
                           f"{qa_card.mark(report.overall)}"):
                HOST.attach_text(f"{qa_card.title(report.module)}",
                                 qa_card.card_text(report))
                attach_html(f"{qa_card.title(report.module)}",
                            qa_card.card_page(report))
        except Exception as error:            # reporting is never fatal
            log.debug("Could not build the QA card for %s: %s",
                      report.module, error)

        # The same checks as a table: what was worked out, from what, against
        # what the application showed. Nothing is recalculated here - every
        # figure is read back off the Check the engine already recorded.
        # The check's own row, minus its Module cell - see TABLE_HEADINGS above.
        rows = [check.as_table_row()[1:] for check in report.checks]
        summary = (f"{report.module}"
                   + (f" - {report.test_case_id}" if report.test_case_id else "")
                   + f" | {counts[PASS]} passed, {counts[FAIL]} failed, "
                     f"{counts[SKIPPED]} not verifiable | "
                     f"Calculation Validation: {report.calculation}")
        # The table, and then the long form of ONLY the checks that need one.
        # Appending every check's full report here repeated, for a passing
        # module, the whole table again in paragraph form - four screens of
        # text saying what the table above says in four lines. A check that
        # failed or could not be verified still gets its full write-up, because
        # that is the one a reader has to understand.
        detail = [check.report() for check in report.checks
                  if check.status != PASS]
        HOST.attach_text("Calculation Validation",
                         summary + "\n\n"
                         + text_table(self.TABLE_HEADINGS, rows)
                         + ("\n\n" + "\n\n".join(detail) if detail else ""))
        # The same calculations as the page a reader opens: the counts as tiles,
        # the table, then one card per calculation carrying its inputs, its
        # formula with this row's own numbers in it, and the verdict. Everything
        # is read back off the Check the engine already recorded - nothing here
        # works a figure out for itself, so this page cannot disagree with the
        # validation it is displaying.
        attach_html("Calculation Validation", qa_report.page(
            f"Calculation Validation - {report.module}", summary,
            [qa_report.kpi_block("Calculations", [
                ["Checked", len(report.checks)],
                ["Passed", counts[PASS], PASS if counts[PASS] else ""],
                ["Failed", counts[FAIL], FAIL if counts[FAIL] else ""],
                ["Not verifiable", counts[SKIPPED],
                 SKIPPED if counts[SKIPPED] else ""],
                ["Result", report.calculation, report.calculation],
             ]),
             qa_report.table_block("Every calculation this test case checked",
                                   self.TABLE_HEADINGS, rows,
                                   verdict_column=self.TABLE_VERDICT_COLUMN)]
            # ...and the long form of the ones that did NOT simply pass. The
            # table above already carries every passing sum's inputs, formula,
            # expected, actual, difference and result, so a card repeating each
            # of them turned a two-screen page into a ten-screen one.
            + [qa_report.pre_block(
                f"{check.name} - {check.status}", check.report())
               for check in report.checks if check.status != PASS],
            verdict=report.calculation))

    def _merged(self) -> List[ModuleReport]:
        """One line per module, however many rows it ran.

        The summary is about modules, and a module is only as good as its worst
        row: the counts are added up, the functional and overall verdicts take
        the worst of the rows, so ten green rows cannot bury one red one.
        """
        order: List[str] = []
        merged: Dict[str, ModuleReport] = {}
        for report in self.reports:
            if report.module not in merged:
                order.append(report.module)
                merged[report.module] = ModuleReport(
                    module=report.module, functional=report.functional,
                    note=report.note)
            into = merged[report.module]
            into.checks.extend(report.checks)
            was = into.functional
            into.functional = _worst(into.functional, report.functional)
            # The TestCaseID of the row that made the module's verdict what it
            # is - not simply the first row's. A blockers section that names
            # the module but the wrong test case sends a reader to the
            # attachments of a row that worked.
            if into.functional != was or not into.test_case_id:
                into.test_case_id = report.test_case_id or into.test_case_id
            if report.note and not into.note:
                into.note = report.note
            # The category belongs to the row that went wrong, so a later clean
            # row must not erase it - and a real defect outranks a blocker when
            # a module has both, because a defect is the thing to act on.
            if report.category and (not into.category
                                    or (into.category in BLOCKING_CATEGORIES
                                        and report.category
                                        not in BLOCKING_CATEGORIES)):
                into.category = report.category
        return [merged[name] for name in order]


# --------------------------------------------------------------------------- #
# THE MODULE-LEVEL API
#
# ONE validator for the whole run, and the module's own names bound to its
# methods - the same shape data_verification.py exposes (VERIFICATION plus
# `start_test = VERIFICATION.start_test`, ...). That is not decoration: the
# framework has to keep ONE record across eleven modules and dozens of rows, so
# there is exactly one instance and every caller reaches it the same way:
#
#     import calculation_validation as calc
#     calc.start_test("Quotations", row)
#     calc.check(...)
#     calc.finish_test()
#
# Callers must NOT build their own CalculationValidator. A second instance keeps
# a second set of results, and the end-of-run summary would then report a
# fraction of the run as if it were the whole of it. The class stays public for
# one purpose only - testing this module without a browser, which is what the
# harness does.
# --------------------------------------------------------------------------- #

#: The one instance the framework talks to.
VALIDATION = CalculationValidator()

# Safe to call at any point in a run, including when no row is open: start_test
# is what opens one, and everything else no-ops rather than raising if it is
# called outside one. So the hooks in Construction_Flow.py need no conditions
# around them.
start_test = VALIDATION.start_test
subject = VALIDATION.subject
functional = VALIDATION.functional
note = VALIDATION.note
check = VALIDATION.check
skip = VALIDATION.skip
not_applicable = VALIDATION.not_applicable
finish_test = VALIDATION.finish_test
mismatches = VALIDATION.mismatches
declare = VALIDATION.declare
summary = VALIDATION.summary
summary_rows = VALIDATION.summary_rows
totals = VALIDATION.totals
calculation_result = VALIDATION.calculation_result
functional_result = VALIDATION.functional_result


def reports() -> List[ModuleReport]:
    """Everything recorded so far - one entry per row that ran."""
    return VALIDATION.reports


def reset() -> None:
    """Forget every result. For the self-tests only; a run never calls it."""
    VALIDATION.reports.clear()
    VALIDATION.current = None


def _reason(report: ModuleReport) -> str:
    """The 'why' cell: the category and the note, for a module that did not pass."""
    if report.overall == PASS and not report.category:
        return "-"
    parts = [part for part in (report.category, report.note) if part]
    return " - ".join(parts) if parts else "-"


def _section(title: str, reports: Sequence[ModuleReport]) -> str:
    """One of the report's named lists, or a line saying it is empty.

    Printed even when empty, on purpose: "AUTOMATION DEFECTS: none" is a
    finding, and a heading that simply vanishes reads as one nobody checked.
    """
    if not reports:
        return f"{title}\n{'-' * len(title)}\nNone."
    lines = [title, "-" * len(title)]
    for report in reports:
        lines.append(f"{report.module} [{report.category or 'uncategorised'}]")
        for line in (report.note or "No reason was recorded.").splitlines():
            lines.append(f"    {line}")
    return "\n".join(lines)


def _verdict_note(t: Dict[str, object]) -> str:
    """One sentence saying what the three verdicts do and do not mean.

    The headline used to be read as "the sums are wrong" whatever had gone
    wrong. This says which of the two halves is speaking.
    """
    calculation, functional = t["calculation_result"], t["functional_result"]
    lines = []
    if calculation == FAIL:
        lines.append(f"CALCULATION: {t['checks_failed']} calculation(s) did not "
                     f"match an independently worked-out value.")
    elif calculation == PASS_WITH_SKIPS:
        lines.append(f"CALCULATION: no calculation was found to be wrong. "
                     f"{t['checks_passed']} were proved correct and "
                     f"{t['checks_skipped']} could not be checked (the figure "
                     f"was not on screen, or an input for it was missing), so "
                     f"the arithmetic is PASS WITH SKIPS, not FAIL.")
    elif calculation == SKIPPED:
        lines.append("CALCULATION: nothing could be checked, so nothing is "
                     "claimed - this is not a failure.")
    elif calculation == PASS:
        lines.append("CALCULATION: every calculation checked was correct.")
    if functional == BLOCKED:
        lines.append("FUNCTIONAL: one or more modules were BLOCKED - the flow "
                     "could not be RUN, because a precondition, an environment "
                     "or a manual step was not available. Nothing about the "
                     "application was proved there, in either direction: a "
                     "BLOCKED module is NOT an application defect, NOT an "
                     "automation defect and NOT a pass. It does not make the "
                     "calculation result a FAIL either. Each one is listed "
                     "under BLOCKERS with its cause and what has to be done "
                     "before the next run.")
    elif functional == FAIL:
        lines.append("FUNCTIONAL: one or more modules did not reach the "
                     "outcome their row asked for. That is a separate finding "
                     "from the arithmetic above.")
    return "\n".join(lines)


#: Worst-first, so _worst() can pick by position.
_SEVERITY = (FAIL, BLOCKED, NOT_AUTOMATED, SKIPPED, NOT_APPLICABLE, PASS)


def _worst(left: str, right: str) -> str:
    for status in _SEVERITY:
        if status in (left, right):
            return status
    return left


def _same_source(expected: float, actual: float) -> bool:
    """True when both sides came out of the same reading of the same screen.

    Not "they are equal" - equal is the whole point of a passing check. This
    catches the one thing that proves nothing: taking a figure off the screen
    and then holding the application to it. Provenance, not identity, because
    two equal floats are routinely the same object in CPython and refusing
    those would quietly turn honest passes into unverified checks.
    """
    return (isinstance(expected, Displayed) and isinstance(actual, Displayed)
            and expected.source == actual.source)


def _explain_mismatch(actual: float, alternatives: Optional[Dict[str, float]],
                      tolerance: float) -> str:
    """Say what the application's figure DOES match, when it matches something.

    A quotation that taxes the pre-discount amount is following a different
    convention, not adding up wrongly, and the difference between those two is
    the whole value of this note. The check still fails - the declared formula
    is the one that was asked for - but the report says which rule the
    application actually appears to be using.
    """
    if not alternatives:
        return ("The application's figure does not match the declared formula. "
                "Confirm the rule on screen before treating this as a defect.")
    for description, value in alternatives.items():
        if value is not None and abs(actual - value) <= tolerance:
            return (f"The application's figure matches a DIFFERENT rule: "
                    f"{description} ({money(value)}). That is a convention "
                    f"difference, not necessarily a fault - confirm which rule "
                    f"is intended, then correct the declared formula or raise "
                    f"the defect.")
    tried = "; ".join(f"{name} = {money(value)}"
                      for name, value in alternatives.items())
    return (f"The application's figure matches neither the declared formula nor "
            f"any of the other rules tried ({tried}).")


# --------------------------------------------------------------------------- #
# THE ARITHMETIC
#
# Every formula the framework applies is written here, once, in one place, so
# that "what does the suite believe the application does?" has a single answer
# that can be read, reviewed and corrected against the application itself.
#
# These are the ordinary rules of an Indian GST document, and they are what the
# forms in this application present (a line's Rate and Quantity, a Discount in
# per cent, a Tax rate the application derives from the item's HSN code, and a
# grand total under them). They are NOT guesses at hidden business logic: every
# one of them is checked against a figure the application itself displays, and
# where the application displays no such figure the check is SKIPPED rather than
# assumed.
#
# If a form turns out to follow a different convention, change it HERE - not in
# the module that calls it.
# --------------------------------------------------------------------------- #

def paisa(value: float) -> float:
    """Money, rounded to the paisa the way money is rounded.

    Not round(). Python's round() is banker's rounding on a binary float, so it
    takes 2.675 to 2.67 and 0.125 to 0.12 - and an application that rounds
    half up, which is what invoicing does, then differs from this framework by a
    paisa on exactly those figures. A paisa is more than MONEY_TOLERANCE, so
    that difference is a FAILED check against a correct application.

    Decimal(str(value)) is deliberate: Decimal(float) would carry the binary
    error the string form does not have (0.1 -> 0.1000000000000000055...).

    The result is handed back as a float because that is what the whole
    framework compares in; the point is not the type, it is that the value has
    been rounded by the commercial rule before anything is compared to it.
    """
    return float(Decimal(str(value)).quantize(Decimal("0.01"),
                                              rounding=ROUND_HALF_UP))


def line_amount(quantity: Optional[float], rate: Optional[float]) -> Optional[float]:
    """Quantity x Rate."""
    if quantity is None or rate is None:
        return None
    return paisa(quantity * rate)


def discount_amount(amount: Optional[float],
                    discount_pct: Optional[float]) -> Optional[float]:
    """Amount x Discount% / 100."""
    if amount is None or discount_pct is None:
        return None
    return paisa(amount * discount_pct / 100.0)


def taxable_value(amount: Optional[float],
                  discount: Optional[float]) -> Optional[float]:
    """Amount - Discount."""
    if amount is None:
        return None
    return paisa(amount - (discount or 0.0))


def tax_amount(taxable: Optional[float],
               tax_pct: Optional[float]) -> Optional[float]:
    """Taxable Value x Tax% / 100."""
    if taxable is None or tax_pct is None:
        return None
    return paisa(taxable * tax_pct / 100.0)


def grand_total(taxable: Optional[float], tax: Optional[float],
                round_off: Optional[float] = None) -> Optional[float]:
    """Taxable Value + Tax (+ the application's own Round Off, if it shows one)."""
    if taxable is None:
        return None
    return paisa(taxable + (tax or 0.0) + (round_off or 0.0))


def tax_charged(taxable: Optional[float], tax_pct: Optional[float],
                inter_state: Optional[bool]) -> Optional[float]:
    """The tax the DOCUMENT carries: the sum of the components it prints.

    Not the same figure as tax_amount() for a supply inside the state, and the
    difference is a paisa that is not a rounding error:

        Taxable 38,372.40 at 18%
        tax_amount()  = round(6,907.032)                    = 6,907.03
        CGST + SGST   = round(3,453.516) x 2 = 3,453.52 x 2 = 6,907.04

    A GST invoice prints CGST and SGST as separate lines, each to the paisa, and
    its total tax has to be what those two lines add up to - a document whose
    components do not sum to its own total is the thing that is wrong. So the
    application rounds each half and adds them, and this follows the same rule
    rather than holding the application to an order it does not use.

    Proved, not assumed: the CGST and SGST checks compare each half against a
    figure worked out HERE, and they passed to the paisa on the run that raised
    this. What was wrong was the framework's model of the total, not the
    application's arithmetic.

    An inter-state supply prints one IGST line, so there is nothing to sum and
    the single rounding is right - which is why this only ever showed up on an
    in-state document.
    """
    tax = tax_amount(taxable, tax_pct)
    if tax is None or inter_state is None:
        return tax
    split = gst_split(tax, inter_state)
    if inter_state:
        return split["IGST"]
    return paisa((split["CGST"] or 0.0) + (split["SGST"] or 0.0))


def gst_split(tax: Optional[float], inter_state: bool) -> Dict[str, Optional[float]]:
    """IGST for an inter-state document, CGST + SGST for one inside the state."""
    if tax is None:
        return {"IGST": None, "CGST": None, "SGST": None}
    if inter_state:
        return {"IGST": paisa(tax), "CGST": 0.0, "SGST": 0.0}
    half = paisa(tax / 2.0)
    return {"IGST": 0.0, "CGST": half, "SGST": half}


def closing_stock(opening: Optional[float], received: Optional[float] = None,
                  issued: Optional[float] = None,
                  consumed: Optional[float] = None,
                  returned: Optional[float] = None) -> Optional[float]:
    """Opening + Received - Issued - Consumed + Returned."""
    if opening is None:
        return None
    return round(opening + (received or 0.0) - (issued or 0.0)
                 - (consumed or 0.0) + (returned or 0.0), 2)


def outstanding(requested: Optional[float],
                settled: Optional[float]) -> Optional[float]:
    """Requested - Settled: what is still to come on an indent or an order."""
    if requested is None:
        return None
    return round(requested - (settled or 0.0), 2)


def percentage(part: Optional[float], whole: Optional[float]) -> Optional[float]:
    """Part / Whole x 100, with the division by zero left as unanswerable."""
    if part is None or not whole:
        return None
    return round(part / whole * 100.0, 2)
