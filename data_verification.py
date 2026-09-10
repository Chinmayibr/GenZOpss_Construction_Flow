"""Data verification - does the application really store what was typed into it?

Every Create in the suite goes through the same three questions, and this module
answers all three by itself, for every module, without a line of code in any
test case:

    1. WHAT WAS ENTERED   every value the framework typed, selected or picked is
                          remembered as it happens, and attached to the Allure
                          report as a table together with a screenshot of the
                          completed form - BEFORE the Save button is clicked.
    2. WHAT IS DISPLAYED  once the record is confirmed saved, its View page is
                          opened and read, and the values on it are attached as a
                          second table together with a screenshot of that page.
    3. DO THEY AGREE      the two are compared field by field into a PASS / FAIL
                          table, with the expected and the actual value spelled
                          out for anything that does not match.

Nothing here is wired into a particular module. The capture hangs off the
framework's own safe_fill / safe_select / safe_click helpers, so a module that
does not exist yet is covered the day it is written - Purchase Orders, GRN,
Enquiry, and every Construction module included. What a module may add is one
line in VIEW_PAGES, and only when its address is not the obvious one.

How it plugs in (Construction_Flow.py):

    bind(...)                     hand over the framework's own helpers, once
    start_test(module, row)       run_test_case(): a new test case begins
    record(label, value)          safe_fill / safe_select / the pickers
    note_skipped(label, value, why)   a field the form does not offer
    before_save(page, "Create Item")  safe_click(), just before a Save click
    after_save(page, name, "Item")    confirm_saved(): the record really exists
    finish_test(page, passed)     run_test_case(): read the View page, compare

Two rules this module keeps to, because it is reporting and nothing else:

    * it NEVER fails a test on its own. A mismatch is reported loudly - in the
      log, in the comparison table and in the PASS/FAIL summary - and the test
      keeps whatever verdict the framework gave it. Set the environment variable
      GENZ_FAIL_ON_DATA_MISMATCH=1 (or FAIL_ON_MISMATCH below) when you want a
      mismatch to fail the test as well.
    * it NEVER leaves the browser somewhere else. The address the page was on
      before the View page was opened is restored afterwards, so the step that
      follows finds the screen it expects.
"""

from __future__ import annotations

import logging
import os
import re
from contextlib import contextmanager
from datetime import datetime
from typing import Dict, Iterator, List, Optional, Sequence, Tuple

from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import Page
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError

# Presentation only - the headings, the colours and the PASS/FAIL badge every
# HTML attachment in the framework shares. It imports nothing back (see the
# module docstring of qa_report.py), so this is not a circle.
import qa_report

log = logging.getLogger("construction_flow")

# --------------------------------------------------------------------------- #
# SETTINGS
# --------------------------------------------------------------------------- #

#: Should a field that does not match fail the test case?
#: OFF by default: this layer is reporting, and turning it into a gate would
#: change the verdict of a suite that already exists. Turn it on here or with
#: GENZ_FAIL_ON_DATA_MISMATCH=1 the day the data checks are trusted enough to
#: block a build.
FAIL_ON_MISMATCH = bool(os.environ.get("GENZ_FAIL_ON_DATA_MISMATCH"))

#: Skip the whole thing (the capture still runs; the View page is not opened).
SKIP_VIEW_VERIFICATION = bool(os.environ.get("GENZ_NO_DATA_VERIFICATION"))

#: How long the View page is given to open before the check gives up and says so.
VIEW_TIMEOUT = 15_000

#: Fields that are never part of the comparison: a search box is not data, and a
#: password is never displayed anywhere.
IGNORED_FIELDS = re.compile(
    r"(^|\s)(search|password|otp|captcha|terms|filter|confirm password)(\s|$)",
    re.IGNORECASE)

#: Suffixes the framework adds to a field description while it is working.
FIELD_SUFFIX = re.compile(r"\s*\((re-entered|retry|attempt)[^)]*\)\s*$",
                          re.IGNORECASE)


# --------------------------------------------------------------------------- #
# THE VIEW PAGE OF EACH MODULE
#
# Only the exceptions need an entry. Anything not listed here gets the obvious
# address - "Purchase Order" -> /purchase-orders - which is the convention the
# application follows, so a future module usually needs nothing at all.
# --------------------------------------------------------------------------- #

class ViewPage:
    """Where a module's records are listed, and how one of them is opened."""

    def __init__(self, list_path: str, opens_detail: bool = True,
                 search: bool = True) -> None:
        #: The list page's address, e.g. "/sales-orders".
        self.list_path = list_path
        #: False for a module that shows everything on the list itself (the
        #: Categories cards), so the check reads the list instead of a record page.
        self.opens_detail = opens_detail
        #: Whether the list page has a search box worth using.
        self.search = search


VIEW_PAGES: Dict[str, ViewPage] = {
    "customer": ViewPage("/customers"),
    "supplier": ViewPage("/suppliers"),
    "item": ViewPage("/items"),
    "category": ViewPage("/categories", opens_detail=False),
    "sub category": ViewPage("/categories", opens_detail=False),
    "quotation": ViewPage("/quotations"),
    "sales order": ViewPage("/sales-orders"),
    "purchase order": ViewPage("/purchase-orders"),
    "grn": ViewPage("/grn"),
    "enquiry": ViewPage("/enquiries"),
    "lead": ViewPage("/leads"),
    "project": ViewPage("/projects"),
    "site": ViewPage("/sites"),
    "indent": ViewPage("/indents"),
    # The sidebar calls it "Material Indents" and the application routes it to
    # /indents, so the address cannot be invented from the name here either.
    "material indent": ViewPage("/indents"),
    "work order": ViewPage("/work-orders"),
    "invoice": ViewPage("/invoices"),
    "payment": ViewPage("/payments"),
}


def view_page_for(what: str) -> ViewPage:
    """The View page of a module, invented from its name when it is not listed.

    "Purchase Order" -> /purchase-orders, "GRN" -> /grns. That is the address the
    application uses for every module it has, so a module added tomorrow is
    covered without touching this file - and when it is not, one line in
    VIEW_PAGES puts it right.
    """
    key = " ".join((what or "").strip().lower().split())
    if key in VIEW_PAGES:
        return VIEW_PAGES[key]

    slug = re.sub(r"[^a-z0-9]+", "-", key).strip("-")
    if slug and not slug.endswith("s"):
        slug += "es" if slug.endswith(("s", "x", "ch", "sh")) else "s"
    return ViewPage(f"/{slug}" if slug else "/")


# --------------------------------------------------------------------------- #
# WHAT THE SAME FIELD IS CALLED ON THE VIEW PAGE
#
# A form says "Customer Name*" and the record page says "Name"; a form says
# "Phone" and the page says "Mobile". Comparing them by their exact wording would
# report every one of those as missing, so each entered field also carries the
# names it can be displayed under.
# --------------------------------------------------------------------------- #
FIELD_ALIASES: Dict[str, Tuple[str, ...]] = {
    "item name": ("name", "item", "product", "product name", "description"),
    "customer name": ("name", "customer", "party", "client", "billed to"),
    "supplier name": ("name", "supplier", "vendor", "party"),
    "parent category": ("category name", "name", "category"),
    "sub category": ("category name", "name", "sub-category", "subcategory"),
    "company name": ("name", "business name", "organisation", "organization"),
    "contact person": ("contact", "owner", "primary contact", "name"),
    "phone": ("mobile", "mobile number", "phone number", "contact",
              "contact number", "phone no"),
    "mobile number": ("mobile", "phone", "phone number", "contact number"),
    "email": ("email address", "e-mail", "mail"),
    "gstin": ("gst", "gst number", "gst no", "gstin number", "gst in"),
    "address": ("street", "address line", "billing address", "location",
                "site address"),
    "pincode": ("pin code", "postal code", "zip", "zip code", "pin"),
    "city": ("town", "district"),
    "state": ("region", "province"),
    "item code": ("code", "sku", "item no", "item number"),
    "hsn code": ("hsn", "hsn/sac", "hsn sac", "hsn code/sac"),
    "item type": ("type", "kind"),
    "category": ("item category", "category name", "group"),
    "purchase price": ("cost price", "purchase rate", "cost", "buy price"),
    "selling price": ("sale price", "price", "selling rate", "sell price"),
    "mrp": ("maximum retail price", "list price", "m.r.p"),
    "credit limit": ("credit", "credit amount"),
    "credit days": ("credit period", "credit term", "credit terms"),
    "account name": ("account holder", "holder name", "beneficiary"),
    "account no": ("account number", "a/c no", "a/c number", "bank account"),
    "bank name": ("bank",),
    "ifsc code": ("ifsc", "ifsc no"),
    "branch": ("bank branch", "branch name"),
    "subject": ("title", "description", "particulars"),
    "valid until": ("valid till", "valid upto", "validity", "expiry date",
                    "expires on"),
    "delivery date": ("expected delivery", "delivery", "delivery on",
                      "expected delivery date"),
    "payment terms": ("payment term", "terms", "payment"),
    "quantity": ("qty", "quantity ordered", "qty."),
    "rate": ("unit price", "price", "unit rate", "rate per unit"),
    "discount": ("discount %", "disc", "discount percent", "discount pct"),
    "tax": ("gst", "gst %", "tax %", "tax rate"),
    "remarks": ("notes", "note", "comments", "description", "remark"),
    "notes": ("remarks", "note", "comments", "description"),
    "color": ("colour",),
    "unit": ("uom", "unit of measure"),
    "status": ("state", "current status"),
    # The purchase side: Purchase Orders -> GRN -> Purchase Bills.
    "expected date": ("expected delivery date", "expected on", "expected",
                      "delivery date", "due date"),
    "purchase order": ("po", "po number", "po no", "po no.", "order",
                       "order number", "against po"),
    "invoice number": ("invoice", "invoice no", "invoice no.", "supplier invoice",
                       "supplier invoice no", "bill number", "bill no"),
    "vehicle number": ("vehicle", "vehicle no", "vehicle no.", "truck number",
                       "truck no", "lorry no"),
    "dc number": ("dc", "dc no", "dc no.", "delivery challan",
                  "delivery challan no", "challan number", "challan no"),
    "bill number": ("bill no", "bill no.", "bill", "invoice number", "invoice no"),
    "inter state": ("inter-state", "inter-state supply", "interstate",
                    "supply type", "igst"),
    # Material Indents. The detail page abbreviates the department to "Dept"
    # and prints the date as "Required by", both beside the IND/nnnn heading.
    "department": ("dept", "site", "location", "cost centre", "cost center"),
    "required by": ("required on", "required date", "needed by", "required"),
    "priority": ("urgency",),
    "reference type": ("reference", "ref type", "against"),
    "indent number": ("indent no", "indent no.", "indent", "ind no"),
    "approved quantity": ("approved", "approved qty", "qty approved"),
    "issued quantity": ("issued", "issued qty", "qty issued"),
    "return quantity": ("returned", "returned qty", "qty returned", "return qty"),
    "return notes": ("return remarks", "reason for return", "remarks"),
}


# --------------------------------------------------------------------------- #
# THE HOST FRAMEWORK'S HELPERS
#
# Injected by bind() so this module never imports Construction_Flow - importing
# it back would be a circle, and under `python Construction_Flow.py` it would
# quietly build a second copy of the whole suite. Every helper has a working
# fallback, so the module also runs on its own.
# --------------------------------------------------------------------------- #

try:
    import allure
    from allure_commons.types import AttachmentType

    _ALLURE = True
except ImportError:                                    # plain `python` run
    _ALLURE = False
    allure = None                                      # type: ignore[assignment]

    class AttachmentType:                              # noqa: D401 - a stand-in
        TEXT = "text/plain"
        HTML = "text/html"
        PNG = "image/png"


def _default_attach_text(name: str, body: str) -> None:
    if _ALLURE:
        allure.attach(body, name=name, attachment_type=AttachmentType.TEXT)


def _default_screenshot(page: Page, name: str) -> None:
    if not _ALLURE:
        return
    try:
        allure.attach(page.screenshot(full_page=True), name=name,
                      attachment_type=AttachmentType.PNG)
    except PlaywrightError as error:
        log.warning("Could not take the '%s' screenshot: %s", name, error)


def _default_wait_ready(page: Page, what: str = "page") -> None:
    try:
        page.wait_for_load_state("domcontentloaded", timeout=VIEW_TIMEOUT)
    except (PlaywrightTimeoutError, PlaywrightError):
        pass


@contextmanager
def _default_step(title: str) -> Iterator[None]:
    log.info("%s", title)
    yield


class _Host:
    """The framework's helpers, or a working stand-in for each of them."""

    attach_text = staticmethod(_default_attach_text)
    screenshot = staticmethod(_default_screenshot)
    wait_ready = staticmethod(_default_wait_ready)
    step = staticmethod(_default_step)
    settle = staticmethod(lambda page: page.wait_for_timeout(400))
    search_in_list = None          # search_in_list(page, term, what) -> bool
    dismiss_overlay = None         # dismiss_overlay(page) -> None


HOST = _Host()


def bind(**helpers) -> None:
    """Hand this module the framework's own helpers.

    Called once from Construction_Flow.py, after they are all defined. Anything
    not passed keeps the fallback above, so the module is never left half wired.
    """
    for name, helper in helpers.items():
        if helper is not None:
            setattr(HOST, name, staticmethod(helper) if callable(helper) else helper)


def attach_html(name: str, body: str) -> None:
    """Attach an HTML fragment - what makes the comparison table colour-coded."""
    if _ALLURE:
        allure.attach(body, name=name, attachment_type=AttachmentType.HTML)


# --------------------------------------------------------------------------- #
# NORMALISING - the same value written two different ways is still the same value
# --------------------------------------------------------------------------- #

_CURRENCY = "₹$€£"
#: A figure as a screen prints it. The sign is allowed on EITHER side of the
#: currency symbol, because this application writes a deduction as "-₹5,121.20"
#: - sign first - and a pattern that only allowed "₹-5,121.20" read that as not
#: a number at all. The consequence was not a wrong figure but a missing one:
#: the quotation's Discount line was reported as "the application did not
#: display this figure" while it was on screen, and the check that needed it was
#: SKIPPED. Verified against the run of 2026-08-17.
_NUMBER_LIKE = re.compile(rf"^[-+]?[{_CURRENCY}\s]*[-+]?[\d,]+(\.\d+)?\s*[%]?$")
_YES = {"yes", "y", "true", "1", "enabled", "active", "on"}
_NO = {"no", "n", "false", "0", "disabled", "inactive", "off"}

_DATE_FORMATS = (
    "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%m/%d/%Y", "%Y/%m/%d", "%d.%m.%Y",
    "%d %b %Y", "%d %B %Y", "%b %d, %Y", "%B %d, %Y", "%d-%b-%Y", "%d-%B-%Y",
    "%b %d %Y", "%B %d %Y", "%d %b, %Y", "%Y-%m-%dT%H:%M:%S",
)


def clean_text(value: object) -> str:
    """A cell or a piece of screen text as one line of plain text."""
    text = "" if value is None else str(value)
    text = text.replace(" ", " ").replace("’", "'")
    return re.sub(r"\s+", " ", text).strip()


def as_number(text: str) -> Optional[float]:
    """The number a value holds - 340, "340.00", "Rs 1,20,000", "18%" - or None."""
    text = clean_text(text)
    if not text or not _NUMBER_LIKE.match(text):
        return None
    stripped = re.sub(rf"[{_CURRENCY}\s,%]", "", text)
    try:
        return float(stripped)
    except ValueError:
        return None


def date_candidates(text: str) -> set:
    """Every date one piece of text could be, as ISO strings.

    A set, not a single value, because 09/10/2026 is a different day in India
    and in the United States and the screen does not say which it means. Two
    values agree when their sets overlap, which is the honest reading.
    """
    text = clean_text(text)
    if not text or len(text) < 6 or len(text) > 30:
        return set()
    found = set()
    for fmt in _DATE_FORMATS:
        try:
            found.add(datetime.strptime(text, fmt).date().isoformat())
        except ValueError:
            continue
    return found


def normalise(value: object) -> str:
    """One canonical spelling of a value, so equal values compare equal.

    Handles what the application does to a value on its way to the screen:
    340 -> 340.00, Yes -> YES, extra spaces, thousands separators, a rupee sign,
    a trailing full stop, and any of a dozen date formats.
    """
    text = clean_text(value)
    if not text:
        return ""

    number = as_number(text)
    if number is not None:
        # 340, 340.0 and 340.00 all become "340"; 340.50 stays "340.5".
        return f"{number:.10g}"

    dates = date_candidates(text)
    if dates:
        return "|".join(sorted(dates))

    folded = text.casefold()
    if folded in _YES:
        return "yes"
    if folded in _NO:
        return "no"

    # Punctuation a screen adds but a form never asked for.
    return re.sub(r"\s+", " ", re.sub(r"[.,;:]+$", "", folded)).strip()


def digits_of(text: str) -> str:
    """Just the digits - what makes 98765 43210 the same phone as 9876543210."""
    return re.sub(r"\D", "", clean_text(text))


def compare(entered: str, shown: Optional[str]) -> Tuple[str, str]:
    """One field: (status, note). Status is PASS, FAIL or SKIPPED.

    Equivalent values are a PASS, which is the whole point of normalise(): the
    application reformats almost everything it is given, and reporting 340 and
    340.00 as a defect would bury the one field that really is wrong.
    """
    entered_text = clean_text(entered)
    shown_text = clean_text(shown) if shown is not None else None

    if not entered_text:
        return "SKIPPED", "nothing was entered in this field"
    if shown_text is None:
        return "SKIPPED", "not displayed on the View page"
    if not shown_text:
        return "FAIL", "the View page shows this field empty"

    if normalise(entered_text) == normalise(shown_text):
        if entered_text != shown_text:
            return "PASS", "same value, displayed in a different format"
        return "PASS", ""

    entered_number, shown_number = as_number(entered_text), as_number(shown_text)
    if entered_number is not None and shown_number is not None:
        if abs(entered_number - shown_number) < 0.005:
            return "PASS", "same number, displayed in a different format"
        return "FAIL", f"expected {entered_text}, the View page shows {shown_text}"

    if date_candidates(entered_text) & date_candidates(shown_text):
        return "PASS", "same date, displayed in a different format"

    # A phone, a GSTIN or an account number, displayed with spaces, dashes or a
    # country code the form was never given: 9876543210 -> +91 98765 43210.
    entered_digits, shown_digits = digits_of(entered_text), digits_of(shown_text)
    if len(entered_digits) >= 6 and len(shown_digits) >= 6:
        if (entered_digits == shown_digits
                or shown_digits.endswith(entered_digits)
                or entered_digits.endswith(shown_digits)):
            return "PASS", "same digits, displayed differently"

    # The View page often prints a field inside a longer line ("Bengaluru,
    # Karnataka 560001"). The value is there, so it saved correctly.
    if len(entered_text) >= 2:
        haystack = normalise(shown_text)
        needle = normalise(entered_text)
        if needle and re.search(rf"(^|\W){re.escape(needle)}($|\W)", haystack):
            return "PASS", f"shown as part of \"{shown_text}\""

    return "FAIL", f"expected \"{entered_text}\", the View page shows \"{shown_text}\""


# --------------------------------------------------------------------------- #
# READING A VIEW PAGE
#
# The record pages of this application are not one shape: some are definition
# lists, some are cards with the label above the value, some are read-only forms
# and some are tables. So the page is read every one of those ways at once and
# the answers are pooled - a label that turns up twice keeps both values, and the
# comparison uses whichever one matches.
# --------------------------------------------------------------------------- #

_READ_PAGE_JS = r"""
() => {
  const clean = t => (t || '').replace(/\s+/g, ' ').trim();
  const pairs = [];
  const seen = new Set();
  const push = (label, value) => {
    label = clean(label).replace(/[:*]+$/, '').trim();
    value = clean(value);
    if (!label || !value) return;
    if (label.length > 60 || value.length > 300) return;
    if (label.toLowerCase() === value.toLowerCase()) return;
    const key = label.toLowerCase() + ' ' + value.toLowerCase();
    if (seen.has(key)) return;
    seen.add(key);
    pairs.push([label, value]);
  };
  const onScreen = el => {
    const rect = el.getBoundingClientRect();
    if (rect.width <= 0 || rect.height <= 0) return false;
    const style = getComputedStyle(el);
    return style.visibility !== 'hidden' && style.display !== 'none';
  };

  /* 1. definition lists - <dt>Label</dt><dd>Value</dd> */
  document.querySelectorAll('dl').forEach(list => {
    const terms = list.querySelectorAll('dt');
    const values = list.querySelectorAll('dd');
    for (let i = 0; i < Math.min(terms.length, values.length); i++) {
      push(terms[i].innerText, values[i].innerText);
    }
  });

  /* 2. tables - a two-cell row is label/value, a wider one is header/cell */
  document.querySelectorAll('table').forEach(table => {
    const headers = [...table.querySelectorAll('thead th')].map(th => clean(th.innerText));
    table.querySelectorAll('tbody tr').forEach(row => {
      const cells = [...row.children];
      if (headers.length && headers.length === cells.length) {
        headers.forEach((header, i) => push(header, cells[i].innerText));
      } else if (cells.length === 2) {
        push(cells[0].innerText, cells[1].innerText);
      }
    });
  });

  /* 3. a read-only form - the value lives in the control, not in the text */
  document.querySelectorAll('input, select, textarea').forEach(field => {
    if (!onScreen(field)) return;
    let label = '';
    if (field.id) {
      const tag = document.querySelector('label[for="' + CSS.escape(field.id) + '"]');
      if (tag) label = tag.innerText;
    }
    if (!label) {
      label = field.getAttribute('aria-label') || field.getAttribute('name')
           || field.getAttribute('placeholder') || '';
    }
    let value = field.value;
    if (field.tagName === 'SELECT' && field.selectedIndex >= 0) {
      value = field.options[field.selectedIndex].text;
    }
    if (field.type === 'checkbox' || field.type === 'radio') {
      value = field.checked ? 'Yes' : 'No';
    }
    push(label, value);
  });

  /* 4. anything the application labels for itself */
  document.querySelectorAll('[data-label], [data-field], [data-testid]').forEach(el => {
    const label = el.getAttribute('data-label') || el.getAttribute('data-field')
               || el.getAttribute('data-testid');
    push(label, el.innerText);
  });

  /* 5. the card pattern - a small label with its value in the next element.
        Only the FIRST text in a card can be its label: in
        <div><p>Item Code</p><p>ITEM-0001</p></div> that rule is what stops
        ITEM-0001 from "labelling" whatever happens to come after it. */
  const LEAVES = 'p, span, div, h1, h2, h3, h4, h5, h6, label, dt, dd, td, li, strong, small, b';
  const firstWithText = parent => {
    if (!parent) return null;
    for (const child of parent.children) if (clean(child.innerText)) return child;
    return null;
  };
  document.querySelectorAll(LEAVES).forEach(el => {
    if (el.children.length) return;                 /* leaves only */
    const label = clean(el.innerText);
    if (!label || label.length > 40) return;
    if (!onScreen(el)) return;
    if (firstWithText(el.parentElement) !== el) return;
    let next = el.nextElementSibling;
    while (next && !clean(next.innerText)) next = next.nextElementSibling;
    if (next) { push(label, next.innerText); return; }
    const parentNext = el.parentElement && el.parentElement.nextElementSibling;
    if (parentNext) push(label, parentNext.innerText);
  });

  const main = document.querySelector('main') || document.body;
  return { pairs: pairs, text: clean(main.innerText).slice(0, 30000) };
}
"""


class ViewData:
    """Everything one View page is showing, ready to be looked up by label."""

    def __init__(self, pairs: Sequence[Sequence[str]], text: str, url: str) -> None:
        self.url = url
        self.text = text
        self.normalised_text = normalise(text)
        #: label (as it reads on screen) -> every value found under it
        self.values: Dict[str, List[str]] = {}
        self.order: List[str] = []
        for label, value in pairs:
            key = label_key(label)
            if not key or not looks_like_a_label(key):
                continue
            if key not in self.values:
                self.values[key] = []
                self.order.append(label)
            if value not in self.values[key]:
                self.values[key].append(value)
        #: label as first seen, for the "View Page Data" table
        self.labels: Dict[str, str] = {label_key(l): l for l in self.order}

    def __bool__(self) -> bool:
        return bool(self.values)

    def lookup(self, field: str) -> List[str]:
        """Every value the page shows for a field, best match first."""
        key = label_key(field)
        if not key:
            return []

        found: List[str] = []
        wanted = {key} | {label_key(alias) for alias in FIELD_ALIASES.get(key, ())}

        for candidate in wanted:                       # 1. the label itself
            for value in self.values.get(candidate, []):
                if value not in found:
                    found.append(value)

        for shown_key, values in self.values.items():  # 2. the label, reversed
            aliases = {label_key(a) for a in FIELD_ALIASES.get(shown_key, ())}
            if key in aliases:
                for value in values:
                    if value not in found:
                        found.append(value)

        if not found:                                  # 3. "Customer Name" ~ "Name"
            for shown_key, values in self.values.items():
                if len(shown_key) < 3:
                    continue
                if _contains_word(shown_key, key) or _contains_word(key, shown_key):
                    for value in values:
                        if value not in found:
                            found.append(value)
        return found

    def mentions_label(self, field: str) -> bool:
        """True when the page prints this field's name somewhere at all."""
        key = label_key(field)
        if not key:
            return False
        if key in self.values:
            return True
        return _contains_word(normalise(self.text), key)

    def mentions_value(self, value: str) -> bool:
        """True when the value itself is somewhere on the page, label or not."""
        needle = normalise(value)
        if not needle or len(needle) < 2:
            return False
        if re.search(rf"(^|\W){re.escape(needle)}($|\W)", self.normalised_text):
            return True
        return self._mentions_date(value)

    def _mentions_date(self, value: str) -> bool:
        """True when a DATE is on the page written a different way round.

        A workbook holds 2026-08-31 and this application prints 31/08/2026, so
        the plain search above cannot find a date it is in fact showing and the
        field is reported as empty. compare() has always accepted the two as the
        same day - it just never got the chance, because a value nothing could
        locate never reached it.
        """
        wanted = date_candidates(value)
        if not wanted:
            return False
        for candidate in re.findall(r"\b[\d]{1,4}[-/.][\d]{1,2}[-/.][\d]{2,4}\b",
                                    self.text):
            if wanted & date_candidates(candidate):
                return True
        return False


def label_key(label: object) -> str:
    """A field name reduced to what two spellings of it have in common."""
    text = clean_text(label).casefold()
    text = re.sub(r"[*:]+", " ", text)
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return " ".join(text.split())


def looks_like_a_label(key: str) -> bool:
    """True when a piece of text can be a field NAME rather than a field value.

    Reading a card grid means guessing which of two neighbouring texts is the
    label, and the guess is wrong half the time - which is how "2523" ends up
    "labelling" the value next to it. A field name has words in it, is short,
    and is never a number, a date or an amount, so anything that is one of those
    is dropped before it can be looked up or printed in the report.
    """
    if not key or len(key) > 40 or not re.search(r"[a-z]", key):
        return False
    if len(key.split()) > 6:
        return False
    return as_number(key) is None and not date_candidates(key)


def _contains_word(haystack: str, needle: str) -> bool:
    """True when `needle` appears in `haystack` as whole words."""
    if not needle or not haystack:
        return False
    return bool(re.search(rf"(^|\W){re.escape(needle)}($|\W)", haystack))


def read_view_page(page: Page) -> ViewData:
    """Read every label / value pair the current page is showing."""
    try:
        payload = page.evaluate(_READ_PAGE_JS)
    except (PlaywrightTimeoutError, PlaywrightError) as error:
        log.warning("   could not read the View page: %s",
                    str(error).splitlines()[0])
        return ViewData((), "", page.url)
    return ViewData(payload.get("pairs", ()), payload.get("text", ""), page.url)


# --------------------------------------------------------------------------- #
# GETTING TO A VIEW PAGE
# --------------------------------------------------------------------------- #

def _origin(page: Page) -> str:
    match = re.match(r"^(https?://[^/]+)", page.url)
    return match.group(1) if match else ""


def _path_of(url: str) -> str:
    return re.sub(r"^https?://[^/]+", "", url).split("?")[0].split("#")[0]


def _on_detail_page(page: Page, config: ViewPage) -> bool:
    """True when the address looks like one record, not the list of them."""
    path = _path_of(page.url).rstrip("/")
    base = config.list_path.rstrip("/")
    if not base or not path.startswith(base):
        return False
    rest = path[len(base):].strip("/")
    return bool(rest) and rest.lower() not in ("new", "create", "add")


def _open_record(page: Page, record_text: str) -> bool:
    """Click the record open in a list. True when something was clicked."""
    name = clean_text(record_text)
    if not name:
        return False

    # A row's own "View" action first - it is the one control on the row that is
    # certain to open the record and certain not to change it.
    row = page.get_by_role("row").filter(has_text=name).first
    try:
        if row.count() > 0:
            for action in ("View", "View Details", "Open"):
                control = row.get_by_role("button", name=action, exact=False) \
                             .or_(row.get_by_role("link", name=action, exact=False))
                if control.count() > 0:
                    control.first.click(timeout=VIEW_TIMEOUT)
                    return True
    except (PlaywrightTimeoutError, PlaywrightError):
        pass

    # Otherwise the record's own name, which every list makes clickable.
    for candidate in (
        page.get_by_role("link", name=name, exact=False).filter(visible=True),
        page.get_by_role("button", name=name, exact=False).filter(visible=True),
        page.get_by_text(name, exact=False).filter(visible=True),
    ):
        try:
            if candidate.count() == 0:
                continue
            candidate.first.click(timeout=VIEW_TIMEOUT)
            return True
        except (PlaywrightTimeoutError, PlaywrightError):
            continue
    return False


def open_view_page(page: Page, record_text: str, what: str) -> Tuple[bool, str]:
    """Put the record's View page on screen. Returns (opened, how).

    Never raises: a View page that cannot be reached is reported in the Allure
    report as a check that could not run, which is a fact about the application,
    not a reason to fail a test that has already proved the record was created.
    """
    config = view_page_for(what)

    if HOST.dismiss_overlay is not None:
        try:
            HOST.dismiss_overlay(page)                 # a form left open by a save
        except Exception as error:                     # noqa: BLE001 - never fatal
            log.debug("   could not clear the screen first: %s", error)

    # 1. Some forms land straight on the new record's page - nothing to open.
    if config.opens_detail and _on_detail_page(page, config):
        HOST.wait_ready(page, f"{what} View page")
        return True, "the application opened it after saving"

    # 2. The list page. Its address is what the sidebar would have navigated to,
    #    and a navigation cannot be swallowed by an overlay.
    origin = _origin(page)
    if not origin:
        return False, "the browser is not on the application"
    try:
        page.goto(f"{origin}{config.list_path}", wait_until="domcontentloaded",
                  timeout=VIEW_TIMEOUT * 2)
    except (PlaywrightTimeoutError, PlaywrightError) as error:
        return False, f"the {config.list_path} page would not open ({error})"
    HOST.wait_ready(page, f"{what} list")

    # 3. Narrow the list down, so the record is on screen and not on page 4.
    if config.search and HOST.search_in_list is not None:
        try:
            HOST.search_in_list(page, record_text, what.lower())
        except Exception as error:                     # noqa: BLE001 - optional
            log.debug("   the %s list has no usable search box: %s", what, error)

    if not config.opens_detail:
        return True, "read from the list page (this module has no record page)"

    # 4. Open the record itself.
    if _open_record(page, record_text):
        HOST.wait_ready(page, f"{what} View page")
        HOST.settle(page)
        if _on_detail_page(page, config) or page.get_by_role("dialog").count() > 0:
            return True, "opened from the list"
        return True, "opened from the list (the address did not change)"

    return True, ("read from the list page - '%s' could not be opened as a record"
                  % record_text)


# --------------------------------------------------------------------------- #
# THE TABLES THAT GO INTO THE REPORT
# --------------------------------------------------------------------------- #

def text_table(headings: Sequence[str], rows: Sequence[Sequence[str]]) -> str:
    """A fixed-width table - the plain-text attachment a report always keeps."""
    columns = list(headings)
    widths = [len(str(heading)) for heading in columns]
    for row in rows:
        for index, cell in enumerate(row):
            widths[index] = max(widths[index], len(str(cell)))

    rule = "-" * (sum(widths) + 3 * (len(columns) - 1))
    lines = [rule,
             "   ".join(str(h).ljust(widths[i]) for i, h in enumerate(columns)),
             rule]
    for row in rows:
        lines.append("   ".join(str(cell).ljust(widths[index])
                                for index, cell in enumerate(row)).rstrip())
    lines.append(rule)
    return "\n".join(lines)


def html_table(title: str, headings: Sequence[str],
               rows: Sequence[Sequence[str]], summary: str = "") -> str:
    """The same table in colour, which is what the report shows at a glance.

    The whole framework's HTML attachments come through here, so this is the one
    place their look is decided - qa_report.py draws it: the verdict in the last
    column becomes a coloured PASS / FAIL / SKIPPED pill, figures are lined up
    on their decimal points, and every verdict the table used is explained
    underneath it in QA English. The signature has not changed, so every
    existing caller got that without a line of its own.
    """
    verdict = str(rows[-1][-1]) if rows and rows[-1] else ""
    return qa_report.page(
        title, summary,
        [qa_report.table_block("", headings, rows)],
        verdict=verdict if len(rows) == 1 else "")


# --------------------------------------------------------------------------- #
# ONE RECORD, FROM THE FORM TO THE VIEW PAGE
# --------------------------------------------------------------------------- #

class CreatedRecord:
    """Everything captured about one Create, from the first field to the report."""

    def __init__(self, module: str, what: str, row=None) -> None:
        self.module = module
        self.what = what
        self.row = row
        self.entered: "Dict[str, str]" = {}     # insertion ordered = screen order
        self.skipped: "Dict[str, str]" = {}     # field -> why it was not entered
        self.record_text = ""
        self.saved = False

    @property
    def title(self) -> str:
        return self.record_text or self.what or self.module


class DataVerification:
    """The capture, the comparison and the reporting for one test case."""

    def __init__(self) -> None:
        self.module = ""
        self.row = None
        self.entered: "Dict[str, str]" = {}
        self.skipped: "Dict[str, str]" = {}
        self.pending: Optional[CreatedRecord] = None
        self.records: List[CreatedRecord] = []
        self.page: Optional[Page] = None

    # -- capture ---------------------------------------------------------- #

    def start_test(self, module: str, row=None) -> None:
        """A new test case: forget everything the last one captured."""
        self.module = module
        self.row = row
        self.entered = {}
        self.skipped = {}
        self.pending = None
        self.records = []

    def record(self, label: str, value: object) -> None:
        """One value entered on the form. Called by the framework's own helpers."""
        field = self._field_name(label)
        if not field:
            return
        self.entered[field] = clean_text(value)
        self.skipped.pop(field, None)

    def note_skipped(self, label: str, value: object, reason: str) -> None:
        """A field that was NOT entered - the form has no such box, or it is
        read-only (the GST badge), or the cell was empty. It belongs in the
        report as a skipped comparison, never as a failure."""
        field = self._field_name(label)
        if not field or field in self.entered:
            return
        text = clean_text(value)
        self.skipped[field] = f"{reason} (sheet says \"{text}\")" if text else reason

    def _field_name(self, label: str) -> str:
        field = FIELD_SUFFIX.sub("", clean_text(label))
        if not field or IGNORED_FIELDS.search(field):
            return ""
        return field

    # -- the save --------------------------------------------------------- #

    def before_save(self, page: Page, description: str) -> None:
        """Just before a Create / Save click: freeze what was entered.

        Called from safe_click() for every button that submits a form, so no
        module and no test case has to remember to do it.
        """
        self.page = page
        if self.pending is not None:                  # a retried click
            return
        if not self.entered and not self.skipped:
            return

        what = re.sub(r"^(create|save|submit|add|update)\b", "", description,
                      flags=re.IGNORECASE).strip() or self.module
        record = CreatedRecord(self.module, what, self.row)
        record.entered = dict(self.entered)
        record.skipped = dict(self.skipped)
        self.pending = record
        self.entered = {}
        self.skipped = {}

        self._attach_entered(record)
        HOST.screenshot(page, f"Create Page Screenshot - {what}")

    def after_save(self, page: Page, record_text: str, what: str) -> None:
        """The framework has proved the record exists: queue it for checking."""
        self.page = page
        record = self.pending
        if record is None:
            # A module that saves without going through before_save (or a second
            # save in the same step) - capture what there is, so the entered data
            # is still in the report.
            if not self.entered:
                return
            record = CreatedRecord(self.module, what, self.row)
            record.entered = dict(self.entered)
            record.skipped = dict(self.skipped)
            self.entered = {}
            self.skipped = {}
            self._attach_entered(record)

        record.what = what or record.what
        record.record_text = clean_text(record_text)
        record.saved = True
        self.records.append(record)
        self.pending = None

    # -- the report ------------------------------------------------------- #

    def finish_test(self, passed: bool = True) -> None:
        """The end of a test case: read every View page and compare.

        Deliberately the LAST thing a test case does. Checking in the middle of
        one would leave the browser on a record page while the module still had
        work to do (a quotation to send, a sub-category to add), and this layer
        is not allowed to change how a test behaves.
        """
        pending, self.pending = self.pending, None
        if pending is not None and pending.entered:
            # The form was filled in but never confirmed saved - the entered data
            # is already attached; say why there is no comparison.
            HOST.attach_text(
                "Expected vs Actual",
                f"The data entered for '{pending.what}' is attached above. There "
                f"is no comparison for it: this step does not confirm a saved "
                f"record, so there is no View page to read it back from.")

        records, self.records = self.records, []
        page = self.page
        if page is None or not records:
            return

        if not passed:
            HOST.attach_text(
                "Expected vs Actual",
                "The test case failed, so the View page was not opened - the "
                "screen is left exactly as it was when it failed, for the "
                "failure screenshot.")
            return

        if SKIP_VIEW_VERIFICATION:
            HOST.attach_text("Expected vs Actual",
                             "View page verification is switched off "
                             "(GENZ_NO_DATA_VERIFICATION).")
            return

        for record in records:
            try:
                self._verify(page, record)
            except Exception as error:                 # noqa: BLE001 - never fatal
                log.warning("   the data verification of '%s' could not run: %s",
                            record.title, str(error).splitlines()[0])
                HOST.attach_text(
                    "Data Verification Error",
                    f"The View page check for '{record.title}' could not run: "
                    f"{error}")

    def _verify(self, page: Page, record: CreatedRecord) -> None:
        """Open the View page of one record, read it, and compare it."""
        title = f"Verifying the saved data on the {record.what} View page..."
        with HOST.step(title):
            was_at = page.url
            opened, how = open_view_page(page, record.record_text, record.what)

            if not opened:
                log.warning("   the %s View page could not be opened (%s)",
                            record.what, how)
                HOST.attach_text(
                    "Expected vs Actual",
                    f"The View page of {record.what} '{record.record_text}' "
                    f"could not be opened, so the entered data could not be "
                    f"compared with it.\nReason: {how}")
                self._restore(page, was_at)
                return

            view = read_view_page(page)
            log.info("   the %s View page shows %d labelled value(s) (%s)",
                     record.what, len(view.values), how)
            # NO screenshot of a View page that agreed with what was typed, and
            # no separate "View Page Data" attachment: the values it would show
            # are a block of the comparison page built below, under the heading
            # "View page data (what the application displayed)". A picture and a
            # table of a screen that was correct prove nothing a reader needs -
            # the mismatch screenshot below is taken when there IS something to
            # look at.

            rows, counts = self._compare(record, view)
            self._attach_comparison(record, view, rows, counts, how)
            if counts["FAIL"]:
                # A second picture of the same page, under a name that cannot be
                # missed in a report full of green ones.
                HOST.screenshot(page, f"DATA MISMATCH - {record.what} View Page")
            self._restore(page, was_at)

            if counts["FAIL"]:
                message = (f"{counts['FAIL']} field(s) of {record.what} "
                           f"'{record.record_text}' do not match what was "
                           f"entered - see 'Expected vs Actual'.")
                log.error("   DATA MISMATCH | %s", message)
                if FAIL_ON_MISMATCH:
                    raise AssertionError(message)
            else:
                log.info("   data verified: %d field(s) match on the %s View page",
                         counts["PASS"], record.what)

    def _compare(self, record: CreatedRecord,
                 view: ViewData) -> Tuple[List[List[str]], Dict[str, int]]:
        """Every entered field against the View page, in the order it was typed."""
        rows: List[List[str]] = []
        counts = {"PASS": 0, "FAIL": 0, "SKIPPED": 0}

        for field, entered in record.entered.items():
            candidates = view.lookup(field)
            shown: Optional[str]
            note = ""

            if candidates:
                # A label can hold more than one value on a busy page; the one
                # that matches is the one the application meant.
                shown = candidates[0]
                for candidate in candidates:
                    if compare(entered, candidate)[0] == "PASS":
                        shown = candidate
                        break
            elif view.mentions_value(entered):
                shown = entered
                note = "found on the View page, but not under a label"
            elif view.mentions_label(field):
                shown = ""
            else:
                shown = None

            status, why = compare(entered, shown)
            if note and status == "PASS":
                why = note
            counts[status] += 1
            rows.append([field, entered, "" if shown is None else shown,
                         status, why])

        for field, reason in record.skipped.items():
            counts["SKIPPED"] += 1
            rows.append([field, "", "", "SKIPPED", reason])

        return rows, counts

    # -- attachments ------------------------------------------------------ #

    def _attach_entered(self, record: CreatedRecord) -> None:
        """What the form was given, as text.

        Kept because it is attached at SAVE time, which is the only report a
        record gets when the run never reaches a View page to compare it with
        (a step that does not confirm a save, a test case that failed first).
        The HTML twin of the same rows is gone: when the comparison DOES
        happen, these values are a block on the "Expected vs Actual" page
        anyway, and two attachments of one table is what a reader reads as
        clutter.
        """
        rows = [[field, value] for field, value in record.entered.items()]
        rows += [[field, f"(not entered - {reason})"]
                 for field, reason in record.skipped.items()]
        HOST.attach_text("Test Data Entered",
                         self._header(record) + "\n"
                         + text_table(("Field", "Entered Value"), rows))

    def _attach_comparison(self, record: CreatedRecord, view: ViewData,
                           rows: List[List[str]], counts: Dict[str, int],
                           how: str) -> None:
        verdict = "FAIL" if counts["FAIL"] else "PASS"
        headings = ("Field", "Expected / Entered", "Actual / Displayed", "Notes",
                    "Result")
        summary = (
            f"{record.module} - {record.what}: {record.record_text}\n"
            f"View page : {view.url}\n"
            f"Opened    : {how}\n"
            f"Compared  : {len(rows)} field(s) - {counts['PASS']} PASS, "
            f"{counts['FAIL']} FAIL, {counts['SKIPPED']} SKIPPED\n"
            f"Result    : {verdict}"
        )

        # Result LAST for the display, because the table renderer badges the
        # last cell of a row - so PASS is green and FAIL is red where a reader
        # is actually looking. `rows` itself keeps its order: the failure detail
        # below reads the status out of it by position.
        display = [[cell[0], cell[1], cell[2], cell[4], cell[3]] for cell in rows]

        # THE THREE THINGS A TESTER CHECKS, IN THE ORDER THEY CHECK THEM:
        # what was typed in, what the record's own page then displayed, and
        # whether the two agree field by field. They used to be three separate
        # attachments a reader had to hold in their head at once; this is the
        # same three, on one page, with the verdict at the top of it.
        entered_rows = [[field, value]
                        for field, value in record.entered.items()]
        entered_rows += [[field, f"(not entered - {reason})"]
                         for field, reason in record.skipped.items()]
        shown_rows = [[view.labels.get(key, key), " | ".join(values)]
                      for key, values in view.values.items()]
        field_verdicts = [[cell[0], cell[3]] for cell in rows]

        HOST.attach_text("Expected vs Actual",
                         summary + "\n\n" + text_table(headings, display)
                         + "\n\n" + qa_report.section(
                             "VALIDATION", qa_report.labelled(field_verdicts))
                         + f"\n\nOverall: {verdict}")
        heading = f"Expected vs Actual - {record.what}"
        attach_html(heading, qa_report.page(
            f"{record.module} - {record.what}",
            f"{record.record_text or record.what} | Creation data, View page "
            f"data and the field-by-field validation",
            [qa_report.kpi_block("Validation", [
                ["Fields compared", len(rows)],
                ["Passed", counts["PASS"], "PASS" if counts["PASS"] else ""],
                ["Failed", counts["FAIL"], "FAIL" if counts["FAIL"] else ""],
                ["Not compared", counts["SKIPPED"],
                 "SKIPPED" if counts["SKIPPED"] else ""],
                ["Overall", verdict, verdict],
             ]),
             qa_report.facts_block("Record", [
                ["Module", record.module],
                ["Test Case", getattr(record.row, "test_case_id", "")],
                ["Record", record.record_text or record.what],
                ["View page", view.url],
                ["Opened by", how],
             ]),
             qa_report.table_block(
                 "Creation data (what was entered on the form)",
                 ("Field", "Entered Value"), entered_rows,
                 verdict_column=99, legend=False),
             qa_report.table_block(
                 "View page data (what the application displayed)",
                 ("Field", "Displayed Value"), shown_rows,
                 note="" if shown_rows else "The View page showed no labelled "
                                            "values that could be read.",
                 verdict_column=99, legend=False),
             qa_report.table_block("Validation - expected vs actual",
                                   headings, display)],
            verdict=verdict))
        # "PASS/FAIL Summary" was this same `summary` string on its own, and it
        # is already the first thing in the text attachment above and the tile
        # row of the page beside it. It is not attached a third time.

        failures = [row for row in rows if row[3] == "FAIL"]
        if not failures:
            return

        detail = [f"{len(failures)} field(s) of {record.what} "
                  f"'{record.record_text}' do not match what was entered.",
                  f"View page: {view.url}", ""]
        for field, entered, shown, _status, why in failures:
            detail.append(f"FIELD          : {field}")
            detail.append(f"  Expected     : {entered}")
            detail.append(f"  Actual (View): {shown or '(the field is empty)'}")
            detail.append(f"  Detail       : {why}")
            detail.append("")
        HOST.attach_text("Data Mismatch Details", "\n".join(detail))

    def _header(self, record: CreatedRecord) -> str:
        lines = [f"Module     : {record.module}", f"Record     : {record.what}"]
        row = record.row
        if row is not None:
            lines.append(f"TestCaseID : {getattr(row, 'test_case_id', '')}")
            lines.append(f"Sheet/Row  : {getattr(row, 'sheet', '')} "
                         f"(row {getattr(row, 'number', '')})")
        return "\n".join(lines) + "\n"

    def _restore(self, page: Page, url: str) -> None:
        """Put the browser back where the test left it.

        The check is allowed to look at a record page; it is not allowed to
        change where the next step starts from. A sales-order row with a blank
        Customer works on whatever is on screen, so this is what keeps the
        verification invisible to the modules.
        """
        if not url or page.url == url:
            return
        try:
            page.goto(url, wait_until="domcontentloaded", timeout=VIEW_TIMEOUT * 2)
            HOST.wait_ready(page, "the page the test was on")
        except (PlaywrightTimeoutError, PlaywrightError) as error:
            log.warning("   could not go back to %s after reading the View page: "
                        "%s", url, str(error).splitlines()[0])


#: The one instance the framework talks to.
VERIFICATION = DataVerification()

# The module-level API. Every one of these is safe to call when nothing is being
# captured, so the hooks in Construction_Flow.py need no conditions around them.
start_test = VERIFICATION.start_test
record = VERIFICATION.record
note_skipped = VERIFICATION.note_skipped
before_save = VERIFICATION.before_save
after_save = VERIFICATION.after_save
finish_test = VERIFICATION.finish_test
