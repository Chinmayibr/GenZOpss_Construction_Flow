"""qa_report.py - how the Allure report LOOKS to a manual QA engineer.

This module holds nothing but presentation. It does not touch the browser, the
workbook, a locator, a wait or a business rule, it decides no verdict of its
own, and no test behaves differently because of it. Everything here takes
figures that have already been worked out somewhere else and lays them out the
way a tester would write them on paper:

    Test Data -> Entered Data -> Displayed Data -> Expected -> Actual ->
    Difference -> Calculation -> Formula -> Validation -> Result -> Observation

WHY IT IS A FILE OF ITS OWN

The three modules that report - Construction_Flow.py (the suite),
data_verification.py (entered vs displayed) and calculation_validation.py (the
arithmetic) - all need the same headings, the same colours and the same
PASS/FAIL badge. Written three times they drift apart, and a report whose two
halves are styled differently reads as two reports. So the renderer lives here,
once, and all three import it.

It imports NOTHING from the project. data_verification imports this, and
calculation_validation imports data_verification, so anything imported back the
other way would be a circle - and under `python Construction_Flow.py` the suite
is `__main__`, so importing it by name would build a second copy of the whole
thing. The dependency runs one way only:

    qa_report  <-  data_verification  <-  calculation_validation  <-  the suite

WHAT A CALLER GETS

    page(...)          one self-contained HTML page, from the blocks below
    kpi_block(...)     the tiles at the top: counts, each with its own verdict
    facts_block(...)   a Label : Value list (Test Data, Environment, ...)
    table_block(...)   a table whose verdict column is coloured
    note_block(...)    a paragraph of plain QA English
    pre_block(...)     something that has to keep its own layout (a sum)

    labelled(...)      the same Label : Value list as aligned plain text
    banner(...)        a text heading
    failure_story(...) MODULE / TEST / EXPECTED / ACTUAL / DETAIL / REASON
    blocker_story(...) MODULE / PRECONDITION / ROOT CAUSE / WHAT TO DO NEXT
    substituted(...)   "Quantity x Rate" -> "567 x 345", for the formula line

    CLASSIFICATION_NOTE  what PASS / FAIL / BLOCKED / SKIPPED / NOT AUTOMATED
                         each mean - one wording, used by every report
    TALLY              how many test cases passed, failed, were blocked and
                       were skipped
"""

from __future__ import annotations

import logging
import re
from typing import Dict, List, Sequence, Tuple

log = logging.getLogger("construction_flow")


# --------------------------------------------------------------------------- #
# THE VERDICT VOCABULARY
#
# Every status any layer of the framework can report, and the one colour it is
# drawn in wherever it appears. A status that is not in this table is drawn
# neutral rather than guessed at - a grey "MEASURED" is honest, a green one
# would be a pass nobody proved.
# --------------------------------------------------------------------------- #

#: status -> (text colour, background, what it MEANS to a reader)
STATUS_STYLE: Dict[str, Tuple[str, str, str]] = {
    "PASS": ("#12692c", "#e4f6e8", "matched the expected value"),
    "PASSED": ("#12692c", "#e4f6e8", "matched the expected value"),
    "PASS WITH SKIPS": ("#6a5000", "#fdf3d4",
                        "nothing was found to be wrong, but not everything "
                        "could be checked"),
    "FAIL": ("#b3251e", "#fdeaea", "did NOT match the expected value"),
    "FAILED": ("#b3251e", "#fdeaea", "did NOT match the expected value"),
    "BLOCKED": ("#8a4b00", "#fdeedd",
                "could not be run - nothing was proved either way"),
    "BLOCKED PRECONDITION": ("#8a4b00", "#fdeedd",
                             "a manual precondition was not met"),
    "SKIPPED": ("#6a5000", "#fdf3d4", "not checked - the reason is stated"),
    "NOT CHECKED": ("#5a5a5a", "#eeeeee", "no check was made at this stage"),
    # A figure the application does not put on screen, or one this run could not
    # read from the screen it was on. It is NOT a failure and it is NOT a pass:
    # nothing was proved either way, and the reason is printed beside every one
    # of them. It exists so that a value which could not be captured is visibly
    # absent instead of being quietly invented - see transaction_report.py.
    "NOT AVAILABLE": ("#5a5a5a", "#eeeeee",
                      "the application does not expose this value, or it could "
                      "not be read - the reason is stated beside it"),
    "MEASURED": ("#28527a", "#e6eef7",
                 "a reading, with nothing yet to compare it against"),
    "NOT APPLICABLE": ("#5a5a5a", "#eeeeee", "there is nothing here to check"),
    "NOT AUTOMATED": ("#5a5a5a", "#eeeeee", "this suite does not drive it"),
    "APPLICATION": ("#b3251e", "#fdeaea", "an application defect"),
    "AUTOMATION": ("#7a3d9e", "#f3e9fb", "the automation's own problem"),
    "TEST DATA / ENVIRONMENT": ("#8a4b00", "#fdeedd",
                                "the data or the environment stopped the row"),
}

#: The neutral pair anything unrecognised is drawn in.
_NEUTRAL = ("#333333", "#f2f4f7")


def style_of(status: object) -> Tuple[str, str]:
    """The (colour, background) one verdict is always drawn in."""
    key = str(status or "").strip().upper()
    found = STATUS_STYLE.get(key)
    return (found[0], found[1]) if found else _NEUTRAL


def meaning_of(status: object) -> str:
    """What a verdict MEANS, in QA English - for the legend under a table."""
    found = STATUS_STYLE.get(str(status or "").strip().upper())
    return found[2] if found else ""


def escape(text: object) -> str:
    """Text that is safe to put inside HTML."""
    return (str("" if text is None else text)
            .replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def badge(status: object) -> str:
    """One verdict as a coloured pill - the thing a reader's eye lands on."""
    text = str(status or "").strip()
    if not text:
        return ""
    colour, background = style_of(text)
    return (f'<span class="badge" style="color:{colour};'
            f'background:{background}">{escape(text)}</span>')


# --------------------------------------------------------------------------- #
# THE PAGE
#
# One stylesheet for every HTML attachment the framework produces, so the
# calculation report, the stock report and the entered-vs-displayed report are
# recognisably pages of the SAME report rather than three different ones.
# --------------------------------------------------------------------------- #

_CSS = """
*{box-sizing:border-box}
body{font-family:"Segoe UI",Roboto,Arial,sans-serif;font-size:13.5px;
     color:#1f2733;background:#fff;margin:0;padding:18px 20px 28px}
h1{font-size:19px;margin:0 0 2px 0;letter-spacing:.2px}
h2{font-size:14px;text-transform:uppercase;letter-spacing:.7px;color:#41556e;
   margin:22px 0 8px 0;padding-bottom:5px;border-bottom:2px solid #e3e8ef}
p.sub{color:#5c6b7f;margin:0 0 4px 0;font-size:13px}
p.note{color:#3d4a5c;margin:0 0 10px 0;line-height:1.5}
.head{border-left:5px solid #28527a;padding:2px 0 2px 12px;margin-bottom:6px}
.badge{display:inline-block;padding:2px 9px;border-radius:11px;
       font-weight:700;font-size:11.5px;letter-spacing:.4px;white-space:nowrap}
table{border-collapse:collapse;width:100%;margin:0 0 6px 0}
th,td{border:1px solid #dde3ea;padding:6px 10px;text-align:left;
      vertical-align:top}
th{background:#eef2f7;font-weight:650;color:#2c3a4d;font-size:12.5px;
   text-transform:uppercase;letter-spacing:.4px}
tr:nth-child(even) td{background:#fafbfd}
td.num{text-align:right;font-variant-numeric:tabular-nums;
       font-family:Consolas,"Courier New",monospace}
td.verdict{text-align:center;white-space:nowrap}
.facts{width:100%}
.facts th{width:34%;background:#f6f8fb;text-transform:none;letter-spacing:0;
          font-size:13px;color:#41556e;font-weight:600}
.tiles{display:flex;flex-wrap:wrap;gap:10px;margin:2px 0 6px 0}
.tile{flex:1 1 140px;min-width:130px;border:1px solid #dde3ea;border-radius:7px;
      padding:9px 12px;background:#fbfcfe}
.tile .label{font-size:11px;text-transform:uppercase;letter-spacing:.6px;
             color:#68788c;margin-bottom:3px}
.tile .value{font-size:21px;font-weight:700;line-height:1.1}
pre{background:#f7f9fc;border:1px solid #dde3ea;border-radius:6px;
    padding:10px 12px;overflow-x:auto;font-family:Consolas,"Courier New",
    monospace;font-size:12.5px;line-height:1.55;margin:0 0 6px 0;
    white-space:pre-wrap}
.legend{color:#68788c;font-size:12px;margin:2px 0 4px 0}
.empty{color:#68788c;font-style:italic}
"""


def page(title: str, subtitle: str = "", blocks: Sequence[str] = (),
         verdict: object = "") -> str:
    """One self-contained HTML page for an Allure attachment.

    Self-contained on purpose: Allure serves an HTML attachment in an iframe of
    its own with no stylesheet of the report's, so a page that relies on
    anything outside itself arrives unstyled.
    """
    body = "\n".join(block for block in blocks if block)
    pill = f" {badge(verdict)}" if verdict else ""
    return (
        "<html><head><meta charset='utf-8'><title>" + escape(title) +
        "</title><style>" + _CSS + "</style></head><body>"
        f"<div class='head'><h1>{escape(title)}{pill}</h1>"
        + (f"<p class='sub'>{escape(subtitle)}</p>" if subtitle else "")
        + "</div>" + body + "</body></html>"
    )


def kpi_block(title: str, tiles: Sequence[Sequence[object]]) -> str:
    """The counts at the top of a summary: (label, value, verdict-or-blank)."""
    if not tiles:
        return ""
    cards = []
    for tile in tiles:
        label, value = tile[0], tile[1]
        status = tile[2] if len(tile) > 2 else ""
        colour = style_of(status)[0] if status else "#1f2733"
        cards.append(f"<div class='tile'><div class='label'>{escape(label)}"
                     f"</div><div class='value' style='color:{colour}'>"
                     f"{escape(value)}</div></div>")
    return (_heading(title) + "<div class='tiles'>" + "".join(cards) + "</div>")


def facts_block(title: str, pairs: Sequence[Sequence[object]],
                note: str = "") -> str:
    """A Label : Value list - Test Data, Environment, the record's identity."""
    if not pairs:
        return ""
    rows = []
    for pair in pairs:
        label, value = pair[0], pair[1]
        shown = escape(value) if str(value or "").strip() else \
            "<span class='empty'>(not given)</span>"
        rows.append(f"<tr><th>{escape(label)}</th><td>{shown}</td></tr>")
    return (_heading(title)
            + (f"<p class='note'>{escape(note)}</p>" if note else "")
            + "<table class='facts'>" + "".join(rows) + "</table>")


def table_block(title: str, headings: Sequence[str],
                rows: Sequence[Sequence[object]], note: str = "",
                verdict_column: int = -1, legend: bool = True) -> str:
    """A table whose verdict column carries the PASS / FAIL badge.

    `verdict_column` is where the verdict is; -1 (the default) means the last
    one, which is where every table in this framework puts it. Numeric-looking
    cells are right-aligned in a monospace figure font, because a column of
    money is read by lining the decimal points up.
    """
    if not rows:
        return (_heading(title) + "<p class='note empty'>Nothing was recorded "
                                  "here.</p>") if title else ""
    width = len(headings)
    verdict_at = verdict_column if verdict_column >= 0 else width - 1

    head = "".join(f"<th>{escape(heading)}</th>" for heading in headings)
    body: List[str] = []
    seen: List[str] = []
    for row in rows:
        cells = []
        for index in range(width):
            value = row[index] if index < len(row) else ""
            if index == verdict_at:
                text = str(value or "").strip()
                if text and text not in seen:
                    seen.append(text)
                cells.append(f"<td class='verdict'>{badge(text)}</td>")
            elif _looks_numeric(value):
                cells.append(f"<td class='num'>{escape(value)}</td>")
            else:
                cells.append(f"<td>{escape(value)}</td>")
        body.append("<tr>" + "".join(cells) + "</tr>")

    explained = ""
    if legend:
        meanings = [f"{status} = {meaning_of(status)}" for status in seen
                    if meaning_of(status)]
        if meanings:
            explained = (f"<p class='legend'>{escape('  |  '.join(meanings))}"
                         f"</p>")

    return (_heading(title)
            + (f"<p class='note'>{escape(note)}</p>" if note else "")
            + f"<table><thead><tr>{head}</tr></thead><tbody>"
            + "".join(body) + "</tbody></table>" + explained)


def note_block(title: str, body: object) -> str:
    """A paragraph of plain QA English - an observation, a reason, a caveat."""
    text = str(body or "").strip()
    if not text:
        return ""
    paragraphs = "".join(f"<p class='note'>{escape(part)}</p>"
                         for part in text.split("\n\n") if part.strip())
    return _heading(title) + paragraphs


def pre_block(title: str, body: object) -> str:
    """Something whose own layout IS the information - a worked-out sum."""
    text = str(body or "").rstrip()
    if not text:
        return ""
    return _heading(title) + f"<pre>{escape(text)}</pre>"


def _heading(title: str) -> str:
    return f"<h2>{escape(title)}</h2>" if title else ""


# --------------------------------------------------------------------------- #
# THE DESCRIPTION FRAGMENT
#
# page() above is for ATTACHMENTS, which Allure serves in an iframe of their
# own - so a full document with its own <style> block is safe there and cannot
# affect anything else.
#
# A test's DESCRIPTION is different: Allure writes it straight into the report's
# own page. A <style> block inserted that way is applied by the browser to the
# WHOLE report, so a stylesheet with bare `body`, `h1` and `table` rules in it
# would quietly restyle every screen of Allure - a report that looks broken and
# a cause nobody would think to look for. So a description carries INLINE styles
# only and never a <style> tag, which is what embed() is for.
# --------------------------------------------------------------------------- #

_EMBED_ROOT = ("font-family:'Segoe UI',Roboto,Arial,sans-serif;"
               "font-size:13.5px;line-height:1.5;color:#1f2733")
_EMBED_TITLE = ("font-size:17px;font-weight:700;letter-spacing:.2px;"
                "margin-bottom:2px")
_EMBED_SUB = "color:#5c6b7f;font-size:13px;margin-bottom:8px"
_EMBED_BAR = "border-left:5px solid #28527a;padding:2px 0 2px 12px;margin:0 0 10px 0"
_EMBED_PRE = ("background:#f7f9fc;border:1px solid #dde3ea;border-radius:6px;"
              "padding:10px 12px;overflow-x:auto;white-space:pre-wrap;"
              "font-family:Consolas,'Courier New',monospace;font-size:12.5px;"
              "line-height:1.55;margin:0")
_EMBED_BADGE = ("display:inline-block;padding:2px 9px;border-radius:11px;"
                "font-weight:700;font-size:11.5px;letter-spacing:.4px;"
                "margin-left:8px;vertical-align:middle")
_EMBED_TH = ("text-align:left;padding:4px 10px 4px 0;color:#41556e;"
             "font-weight:600;white-space:nowrap;vertical-align:top")
_EMBED_TD = "padding:4px 0;vertical-align:top"


def embed(title: str, body: object, verdict: object = "",
          subtitle: str = "", pairs: Sequence[Sequence[object]] = (),
          table_html: str = "") -> str:
    """A test description: the verdict, the facts and the QA text, inline-styled.

    Everything a reader should see before opening anything, and nothing that can
    escape the element it is drawn in - see the note above for why that matters.

    `table_html` is already-built HTML (e.g. a real <table> from table_block())
    - inserted as-is INSTEAD of the <pre>-wrapped `body`, for a caller whose
    content is itself a table and must not be shown twice, once as a table and
    once as the same figures dumped into a monospace block underneath it. Every
    existing caller that does not pass it keeps the exact <pre> rendering it
    always had.
    """
    colour, background = style_of(verdict)
    pill = (f"<span style=\"{_EMBED_BADGE};color:{colour};"
            f"background:{background}\">{escape(verdict)}</span>"
            if str(verdict or "").strip() else "")

    facts = ""
    if pairs:
        rows = "".join(
            f'<tr><th style="{_EMBED_TH}">{escape(pair[0])}</th>'
            f'<td style="{_EMBED_TD}">{escape(pair[1])}</td></tr>'
            for pair in pairs)
        facts = (f'<table style="border-collapse:collapse;margin:0 0 10px 0">'
                 f"{rows}</table>")

    text = str(body or "").rstrip()
    return (
        f'<div style="{_EMBED_ROOT}">'
        f'<div style="{_EMBED_BAR}">'
        f'<div style="{_EMBED_TITLE}">{escape(title)}{pill}</div>'
        + (f'<div style="{_EMBED_SUB}">{escape(subtitle)}</div>'
           if subtitle else "")
        + "</div>" + facts
        + (table_html if table_html else
           (f'<pre style="{_EMBED_PRE}">{escape(text)}</pre>' if text else ""))
        + "</div>"
    )


_NUMERIC = re.compile(r"^[₹$€£\s]*[-+]?[\d,]+(\.\d+)?\s*%?$")


def _looks_numeric(value: object) -> bool:
    text = str(value or "").strip()
    return bool(text) and text != "-" and bool(_NUMERIC.match(text))


# --------------------------------------------------------------------------- #
# THE SAME INFORMATION AS PLAIN TEXT
#
# Both, always. The HTML page is what a reader opens; the text one is what
# survives being pasted into a defect ticket, an e-mail or a chat window - and
# it is the only one that can be read at all when the report is opened without
# a browser.
# --------------------------------------------------------------------------- #

def labelled(pairs: Sequence[Sequence[object]], indent: str = "") -> str:
    """A Label : Value list, aligned on the colon."""
    items = [(str(pair[0]), str("" if pair[1] is None else pair[1]))
             for pair in pairs]
    if not items:
        return ""
    width = max(len(label) for label, _ in items)
    return "\n".join(f"{indent}{label:<{width}} : {value}"
                     for label, value in items)


def banner(title: str, rule: str = "=") -> str:
    """A heading a text attachment can be scanned by."""
    return f"{title}\n{rule * len(title)}"


def section(title: str, body: object) -> str:
    """One titled block of a plain-text attachment."""
    text = str(body or "").rstrip()
    return f"{banner(title, '-')}\n{text}" if text else ""


def joined(parts: Sequence[object]) -> str:
    """The parts of a text attachment, with one blank line between them."""
    return "\n\n".join(str(part).rstrip() for part in parts
                       if str(part or "").strip())


# --------------------------------------------------------------------------- #
# THE FORMULA, WITH THIS ROW'S OWN NUMBERS IN IT
# --------------------------------------------------------------------------- #

def substituted(formula: str, inputs: Dict[str, object]) -> str:
    """"Quantity x Rate" + {Quantity: 567, Rate: 345} -> "567 x 345".

    What a manual tester writes underneath the formula when they check it by
    hand, and the line that turns an abstract rule into this row's arithmetic.

    Only the labels that actually appear in the formula are put in, longest
    first so "Taxable Value" is not half-replaced by "Value". A formula whose
    labels do not appear in it (a sentence rather than an expression) gets
    nothing back, and the caller prints the formula alone - inventing a
    substitution nobody can check would be worse than leaving it out.
    """
    text = str(formula or "").strip()
    if not text or not inputs:
        return ""

    usable = {label: value for label, value in inputs.items()
              if _substitutable(value)}
    if not usable:
        return ""

    filled = text
    used = 0
    for label in sorted(usable, key=len, reverse=True):
        pattern = re.compile(re.escape(label), re.IGNORECASE)
        if not pattern.search(filled):
            continue
        filled = pattern.sub(_number(usable[label]).replace("\\", ""), filled)
        used += 1
    return filled if used else ""


def _substitutable(value: object) -> bool:
    """True for a value that can stand in the middle of a sum."""
    if isinstance(value, bool) or value is None:
        return False
    if isinstance(value, (int, float)):
        return True
    text = str(value).strip()
    return bool(text) and bool(_NUMERIC.match(text))


def _number(value: object) -> str:
    """A figure as it should read inside a substituted formula."""
    if isinstance(value, (int, float)):
        text = f"{float(value):,.4f}".rstrip("0").rstrip(".")
        return text or "0"
    return str(value).strip()


# --------------------------------------------------------------------------- #
# A FAILURE, IN QA ENGLISH
#
# The report has to answer "what went wrong?" before it answers "what did
# Python raise?". Both are kept - this is the first half.
# --------------------------------------------------------------------------- #

def failure_story(module: str, test_case: str, test_name: str = "",
                  expected: object = "", actual: object = "",
                  detail: Sequence[Sequence[object]] = (),
                  reason: str = "", last_action: str = "", url: str = "",
                  result: str = "FAIL", category: str = "") -> str:
    """The failure a manual tester would have written in the defect ticket.

    Every field is optional: what is known is printed, and what is not is left
    out rather than filled in with a guess.
    """
    header: List[List[object]] = [["MODULE", module or "-"]]
    if test_case:
        header.append(["TEST CASE", test_case])
    if test_name:
        header.append(["TEST", test_name])
    if category:
        header.append(["CATEGORY", category])

    parts = [banner(f"{result}: {module or 'this test case'}"),
             labelled(header)]

    if str(expected or "").strip() or str(actual or "").strip():
        parts.append(section("EXPECTED VS ACTUAL", labelled([
            ["EXPECTED", expected or "(not stated)"],
            ["ACTUAL", actual or "(not stated)"],
        ])))
    if detail:
        parts.append(section("DETAIL", labelled(detail)))

    closing: List[List[object]] = [["RESULT", result]]
    if last_action:
        closing.append(["LAST ACTION THAT WORKED", last_action])
    if url:
        closing.append(["PAGE AT THE TIME", url])
    parts.append(labelled(closing))

    if reason:
        parts.append(section("REASON", reason))
    return joined(parts)


def blocker_story(module: str, precondition: str = "", root_cause: str = "",
                  expected: str = "", actual: str = "",
                  before_rerun: object = "", test_case: str = "") -> str:
    """A BLOCKED module written out the way a QA lead has to hand it on.

    A blocker is not a defect and it is not a pass, so it needs its own shape:
    a defect report asks "what is wrong with the product?" and the answer here
    is "nothing was established". What a reader needs instead is the six things
    below - above all the last one, because a blocker is the only result that
    comes with an action for the person about to re-run the suite.

        BLOCKED MODULE                which module could not be run
        PRECONDITION                  what had to be true before it could
        ROOT CAUSE                    why that was not true, from the evidence
        WHAT WAS EXPECTED             what the step was waiting for
        WHAT ACTUALLY HAPPENED        what the run saw instead
        WHAT NEEDS TO BE DONE         the action that unblocks the re-run
          BEFORE RE-RUN

    Nothing is inferred: every field is passed in by the step that hit the
    blocker, from what it actually observed, and a field nobody could fill is
    left out rather than guessed at.
    """
    fields: List[List[object]] = [["BLOCKED MODULE", module or "-"]]
    if test_case:
        fields.append(["TEST CASE", test_case])
    for label, value in (("PRECONDITION", precondition),
                         ("ROOT CAUSE", root_cause),
                         ("WHAT WAS EXPECTED", expected),
                         ("WHAT ACTUALLY HAPPENED", actual)):
        if str(value or "").strip():
            fields.append([label, value])

    parts = [banner(f"BLOCKED: {module or 'this module'}"), labelled(fields)]

    steps = before_rerun if isinstance(before_rerun, (list, tuple)) \
        else [before_rerun] if str(before_rerun or "").strip() else []
    if steps:
        parts.append(section(
            "WHAT NEEDS TO BE DONE BEFORE RE-RUN",
            "\n".join(f"  {index}. {step}"
                      for index, step in enumerate(steps, start=1))))

    parts.append(section("HOW TO READ THIS", CLASSIFICATION_NOTE))
    return joined(parts)


#: What each verdict means, in one place, so every report says the same thing.
CLASSIFICATION_NOTE = (
    "PASS         the check ran and matched the expected result.\n"
    "FAIL         the check ran and the application's result was wrong, or a\n"
    "             behaviour the suite verified did not hold. A real finding.\n"
    "BLOCKED      the test could NOT be run: a precondition, an environment or\n"
    "             a manual step was not available. Nothing was proved about the\n"
    "             application either way, so this is neither an application\n"
    "             defect nor an automation defect.\n"
    "SKIPPED      one calculation or check could not be validated because the\n"
    "             value or the data it needed was not available. The reason is\n"
    "             printed beside every one of them.\n"
    "NOT AVAILABLE  the application does not display this value, or this run\n"
    "             could not read it from the screen it was on. Nothing was\n"
    "             assumed and nothing was invented - the reason is printed\n"
    "             beside it. Used by the before/after transaction report for a\n"
    "             count or a figure the product does not expose.\n"
    "NOT AUTOMATED  the module is deliberately outside this suite's scope. It\n"
    "             is listed so the coverage the report claims is the coverage\n"
    "             there really is - never counted as passed, failed or skipped."
)


def failure_embed(module: str, test_case: str, story: str,
                  reason: str = "", result: str = "FAIL") -> str:
    """The same failure as the card the report draws at the top of the test.

    A description fragment rather than a page - see embed() for why a
    description must never carry a stylesheet of its own.
    """
    return embed(f"{module or 'Test case'} - {result}", story, verdict=result,
                 subtitle=reason,
                 pairs=[["Test Case", test_case]] if test_case else ())


# --------------------------------------------------------------------------- #
# HOW MANY TEST CASES PASSED
#
# The module tables count MODULES and the calculation tables count SUMS.
# Neither of them counts TEST CASES, which is the first figure a QA lead asks
# for - so it is counted here, from pytest's own outcome for every test that
# ran, and nowhere else. Nothing is written in by hand.
# --------------------------------------------------------------------------- #

class ExecutionTally:
    """One outcome per test case, kept for the execution summary."""

    #: The summary test is a REPORT of the run, not a test case of it, so it
    #: must not be counted as one - it would pass or fail on the strength of
    #: everything else and inflate whichever column it landed in.
    EXCLUDED = ("test_validation_summary",)

    def __init__(self) -> None:
        self.outcomes: Dict[str, str] = {}
        self.titles: Dict[str, str] = {}

    #: The outcomes a test case can end with, from best to worst. BLOCKED sits
    #: between SKIPPED and BROKEN on purpose, and it is a status of its own
    #: rather than a kind of failure:
    #:
    #:    PASSED   it ran and did what the row asked
    #:    SKIPPED  a check inside it had nothing to work with (a used-up row)
    #:    BLOCKED  it could not be RUN - a precondition, an environment or a
    #:             manual step was not there. Nothing was proved either way, so
    #:             it is neither a pass nor a defect, and counting it as either
    #:             would be a lie in one direction or the other.
    #:    BROKEN   the suite itself fell over (setup / teardown)
    #:    FAILED   it ran and the outcome was wrong - a real finding
    ORDER = ("PASSED", "SKIPPED", "BLOCKED", "BROKEN", "FAILED")

    def record(self, nodeid: str, status: str, title: str = "") -> None:
        """Remember one test case's result. The worst wins - see below."""
        if not nodeid or any(name in nodeid for name in self.EXCLUDED):
            return
        if status not in self.ORDER:
            return
        # setup, call and teardown all report. A test that passed its call and
        # then broke in teardown is not a pass, so the worst outcome of the
        # three is the one kept.
        current = self.outcomes.get(nodeid)
        if current is None or self.ORDER.index(status) > self.ORDER.index(current):
            self.outcomes[nodeid] = status
        if title:
            self.titles[nodeid] = title

    def counts(self) -> Dict[str, int]:
        """total / passed / failed / skipped / blocked / broken."""
        values = list(self.outcomes.values())
        return {
            "total": len(values),
            "passed": values.count("PASSED"),
            "failed": values.count("FAILED"),
            "skipped": values.count("SKIPPED"),
            "blocked": values.count("BLOCKED"),
            "broken": values.count("BROKEN"),
        }

    def result(self) -> str:
        """PASS / FAIL / BLOCKED / SKIPPED over the test cases alone."""
        counts = self.counts()
        if counts["failed"] or counts["broken"]:
            return "FAIL"
        if counts["blocked"]:
            return "BLOCKED"
        if counts["passed"] and counts["skipped"]:
            return "PASS WITH SKIPS"
        if counts["passed"]:
            return "PASS"
        if counts["skipped"]:
            return "SKIPPED"
        return "NOT APPLICABLE"

    def rows(self) -> List[List[str]]:
        """Every test case and how it came out, for the summary's own table."""
        return [[self.titles.get(nodeid, nodeid.split("::")[-1]), status]
                for nodeid, status in self.outcomes.items()]


#: The one tally for the run. Construction_Flow's pytest hook feeds it.
TALLY = ExecutionTally()


def reset_tally() -> None:
    """Forget every outcome. For the self-tests only; a run never calls it."""
    TALLY.outcomes.clear()
    TALLY.titles.clear()
