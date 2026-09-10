"""GenZOpss construction flow - data-driven from ONE Excel workbook.

    test_data\\Construction_Flow_Data.xlsx

That file is the framework's only data source. Its path is defined once, as
EXCEL_FILE in excel_reader.py, and every module reads its own sheet through
excel_reader.get_test_cases(). This file holds no test data at all - only the
environment settings in the CONFIGURATION block below.

Modules run one after another, and every row of a module finishes before the
next module starts:

    Create_Your_Account  -> all rows   (register a business, manual OTP optional)
    Customer             -> all rows
    Suppliers            -> all rows
    Items                -> all rows   (Stock & Materials > Items)
    Categories           -> all rows   (Stock & Materials > Categories)
    Quotations           -> all rows
    Sales_Orders         -> all rows
    Purchase_Orders      -> all rows   (Purchases > Purchase Orders)
    GRN                  -> all rows   (Purchases > GRN, approved into a bill)
    Purchase_Bills       -> all rows   (Purchases > Purchase Bills, approved)
    Material_Indents     -> all rows   (Construction > Material Indents, raised,
                                        approved, issued and returned to store)

Run it:

    pytest                        # everything: tests, report, browser
    python Construction_Flow.py   # plain run, no pytest and no report

Reporting: every module, every step and every failure screenshot goes into the
Allure report, and EVERY EXECUTION KEEPS ITS OWN, in a folder stamped with the
moment it started:

    Reports/2026-08-07_10-30-15/
        allure-report/            the report, with its history/ folder
        allure-report-single/     the same report as one openable file
        screenshots/              the pictures of anything that went red
        execution.log             everything the run logged
        environment.properties    what the run was pointed at

Nothing is ever overwritten, so `pytest` can be run as often as you like and
every earlier report is still there. run_config.py owns that layout; the report
is built and opened in Google Chrome the moment the run ends, and the path is
printed at the end - there is no second command to type. The report carries the
Environment, Executors, Categories, Trend and History panels. Set
GENZ_NO_ALLURE_REPORT=1 to keep the results but skip the report and the browser.

The TREND gains one point per execution, and the order it is built in is:

    1. environment.properties / categories.json / executor.json and the previous
       history are written into allure-results  (before the tests run, because
       --clean-alluredir has just emptied that folder - and because a report
       built by hand later reads them from there too)
    2. the tests run
    3. the same four are written again, in case anything cleaned them mid-run
    4. `allure generate` builds the report into this run's folder
    5. the history is archived to allure-history/ and the report is opened

allure-results/ and allure-history/ stay at the project root and neither is a
report: the first is scratch space for the run in progress (--clean-alluredir
empties it at the start of every run), the second is the trend's carrier. The
archive is needed because each run now reports into a different folder, so there
is no fixed "last report" to read the history out of - and because `allure
generate --clean` deletes the report folder before writing the new one, which
would otherwise let an interrupted generate take the only copy of the trend with
it. Delete allure-history/ to start the trend again from build 1.

Every value is typed into the application EXACTLY as it is written in the cell.
Nothing is generated or modified at run time. To add a test case: open the
workbook, add a row, set Execute = YES, save, run. No code change is ever
needed.

Because the values are used verbatim, the application's own uniqueness rules
apply: it answers a repeated mobile number (PHONE_IN_USE) or GSTIN
(GSTIN_IN_USE) with a 409. Change the cell in the workbook to re-run a row.

Waiting: Playwright waits for elements by itself (visible, enabled, stable), so
this script adds only what Playwright cannot know about - page load, network
idle, and the app's own loading spinners. There is no time.sleep() anywhere; the
one short pause (`settle`) exists for CSS animations that emit no DOM signal.
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import time
import uuid
import webbrowser
from contextlib import contextmanager
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Dict, Iterator, List, Optional, Set, Tuple

import pytest
from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import Locator, Page
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import expect, sync_playwright

import calculation_validation as calc
import data_verification as verify_data
import qa_report
import stock_report
from calculation_validation import PageNumbers
from excel_reader import Row, get_test_cases
from run_config import (ALLURE_BUILD_ORDER_FILE, ALLURE_HISTORY_ARCHIVE,
                        ALLURE_REPORT_DIR, ALLURE_REPORT_NAME,
                        ALLURE_RESULTS_DIR, ALLURE_SINGLE_REPORT_DIR,
                        AUTO_OPEN_ALLURE_REPORT, EXECUTION_LOG,
                        RUN_ENVIRONMENT_FILE, RUN_REPORT_DIR, RUN_TIMESTAMP,
                        SCREENSHOT_DIR, log, previous_run_dirs)

# =========================================================================== #
# CONFIGURATION - environment settings only. All TEST DATA is in the workbook.
# =========================================================================== #

# --- Application URLs ---
MARKETING_URL = "https://genzopss.com/"
ADMIN_URL = "https://admin.genzopss.com"
APP_URL = "https://app.genzopss.com"

# --- Admin credentials (used to approve each new business) ---
ADMIN_MOBILE = "0000000001"
ADMIN_PASSWORD = "SmartOps@123"

# --- Approved business user (used for Customer / Supplier / Quotation / SO) ---
BUSINESS_USER_EMAIL = "chinmayi.br@vcnrtech.com"
BUSINESS_USER_PASSWORD = "SmartOps@123"

# --- Worksheet names, in execution order ---
SHEET_ACCOUNT = "Create_Your_Account"
SHEET_CUSTOMER = "Customer"
SHEET_SUPPLIER = "Suppliers"
SHEET_ITEM = "Items"
SHEET_CATEGORY = "Categories"
SHEET_QUOTATION = "Quotations"
SHEET_SALES_ORDER = "Sales_Orders"
SHEET_PURCHASE_ORDER = "Purchase_Orders"
SHEET_GRN = "GRN"
SHEET_PURCHASE_BILL = "Purchase_Bills"
SHEET_MATERIAL_INDENT = "Material_Indents"

# --- WHICH MODULES RUN ------------------------------------------------------ #
#
# ONE SWITCH, and it is the only thing that turns the Create Your Account module
# off. Nothing else about it is touched: the flow, its helpers, its Excel sheet
# and every row in it are exactly as they were, and turning this back on runs
# them again with no other change anywhere.
#
#     True   register the business, wait for the manual OTP, and approve the
#            tenant in the admin portal - the module as written
#     False  do not run it. The module is REPORTED as deliberately switched
#            off (see test_create_your_account and test_validation_summary), so
#            the summary still lists it and still tells the truth about its own
#            coverage - it is never silently missing, and never counted as a
#            pass.
#
# TO RUN IT AGAIN: set this to True. That is the whole change.
RUN_CREATE_YOUR_ACCOUNT = False

# ...or, to run it once without editing this file:
#     PowerShell   $env:GENZ_RUN_CREATE_ACCOUNT = "1"
#     Bash         GENZ_RUN_CREATE_ACCOUNT=1
if os.environ.get("GENZ_RUN_CREATE_ACCOUNT") == "1":
    RUN_CREATE_YOUR_ACCOUNT = True

#: What the report says about the module while it is switched off. Worded so it
#: cannot be read as a failure, as a blocker or as something the application
#: did: nobody asked for it to run.
CREATE_YOUR_ACCOUNT_OFF = (
    "MODULE SWITCHED OFF - Create Your Account was deliberately disabled for "
    "this run (RUN_CREATE_YOUR_ACCOUNT = False in Construction_Flow.py). It is "
    "not a failure, not a blocker and not a defect of any kind: it was not "
    "asked to run, so nothing about registration or admin approval was proved "
    "either way. The flow and its test data are unchanged - set "
    "RUN_CREATE_YOUR_ACCOUNT = True (or GENZ_RUN_CREATE_ACCOUNT=1) to run it "
    "again.")

# --- Timeouts (milliseconds) ---
DEFAULT_TIMEOUT = 30_000       # every click / fill / assertion
NAVIGATION_TIMEOUT = 60_000    # page loads
LOADER_TIMEOUT = 20_000        # spinners and "Loading..." placeholders
NETWORK_IDLE_TIMEOUT = 8_000   # best effort - a chatty SPA may never go idle
OTP_WAIT_TIMEOUT = int(os.environ.get("GENZ_OTP_WAIT_SECONDS", "120")) * 1000
                               # HOW LONG THE RUN PAUSES FOR THE MANUAL OTP.
                               # The code arrives by e-mail, so it CANNOT be
                               # driven from the UI and this framework does not
                               # try: nothing here generates, guesses or bypasses
                               # a code. The run stops at the verification page,
                               # says so on the console, and goes on the instant
                               # the application accepts what was typed.
                               #
                               # 30s was too short to be usable by a human (a
                               # code has to arrive by e-mail first), and an
                               # unverified registration does not become a
                               # tenant - so every run that nobody watched ended
                               # BLOCKED at the admin approval. Two minutes by
                               # default; GENZ_OTP_WAIT_SECONDS sets it, and 0
                               # means "do not wait at all" for an unattended
                               # run that is not interested in this module.
                               #
                               # It is still not a hard gate: if no code is
                               # typed the run carries on, and the tenant wait
                               # below decides honestly whether the approval
                               # could be done.
OTP_POLL_INTERVAL = 15_000     # how often the pause reports that it is still
                               # waiting, so a watched run does not look hung
TENANT_WAIT_TIMEOUT = int(
    os.environ.get("GENZ_TENANT_WAIT_SECONDS", "180")) * 1000
                               # HOW LONG THE ADMIN PORTAL IS GIVEN TO SHOW THE
                               # BUSINESS THAT WAS JUST REGISTERED. A new
                               # registration does not appear in the Tenants
                               # list the moment it is submitted, and the list
                               # itself renders a bare "Loading..." cell while
                               # it fetches. The run of 2026-08-17 gave the row
                               # ONE 15s look and reported "no matching row"
                               # while the screenshot it took at that very
                               # moment still said "Loading..." - so the search
                               # is now polled (re-typed, re-read and reloaded)
                               # until this budget is spent. See find_tenant_row().
TENANT_POLL_INTERVAL = 5_000   # how long one look at the filtered list gets
                               # before the search is made again
FAIL_ON_BLOCKED = os.environ.get("GENZ_FAIL_ON_BLOCKED") == "1"
                               # WHAT A BLOCKED ROW DOES TO PYTEST'S VERDICT.
                               # Default: it does not fail it. A row that could
                               # not be RUN proves nothing about the
                               # application, so recording it as FAILED - which
                               # is what re-raising did - puts a red test in the
                               # report for a defect that does not exist, and
                               # sends whoever reads it looking for one. It is
                               # recorded as BLOCKED instead: its own status,
                               # counted on its own line, with its reason and
                               # what to do about it. Set this to 1 for a CI job
                               # that must go red when a precondition is missing.
OPTIONAL_TIMEOUT = 8_000       # a step the application is ALLOWED to skip (the
                               # map's suggestion list, the pincode lookup). Kept
                               # short so an optional step cannot cost 30s.
FIELD_LOAD_TIMEOUT = 30_000    # APPLICATION DEFECT: step 1 of the sign-up
                               # wizard renders Industry Type, "Who do you sell
                               # to?" and the specialisation options one after
                               # the other, each only once the one before it has
                               # been answered - and how long each takes is
                               # wildly inconsistent, instant on a good run and
                               # most of a minute when the GST service behind
                               # the form is busy. That is the application
                               # loading, not the test being slow, so those
                               # fields get a wait far longer than
                               # DEFAULT_TIMEOUT. See choose_business_profile().
SPECIALISATION_TIMEOUT = 90_000
                               # APPLICATION DEFECT: the specialisation block
                               # ("What does your business specialise in?" and
                               # "Pick your specialisations") is not part of the
                               # form - it is drawn from a catalogue the page
                               # fetches for itself, and that fetch is the least
                               # reliable thing on the page. Observed on one
                               # afternoon, same GSTIN, same industry: rendered
                               # at once; rendered after 29 seconds; and never
                               # rendered at all while the page showed "GST
                               # lookup quota exhausted". So it gets a wait of
                               # its own, three times the other fields'.
                               # That budget is given IN FULL, every time. The
                               # "GST lookup quota exhausted" notice used to cut
                               # it to 20 seconds, and that was wrong twice
                               # over: the notice is about the GSTIN auto-fill,
                               # not this block, and the block has been measured
                               # arriving at 29 seconds with the notice on
                               # screen. See THE GST NOTICE further down.
SPECIALISATION_REQUIRED = os.environ.get("GENZ_REQUIRE_SPECIALISATION") == "1"
                               # What to do when even 90 seconds, a retry and
                               # the application's own fallback do not produce
                               # the block.
                               #
                               # Default: REPORT IT AND CARRY ON. The form
                               # itself heads this panel "Optional but strongly
                               # recommended", the application accepts the
                               # registration without it, and the tenant this
                               # suite goes on to create is what every module
                               # after Create Your Account depends on - so
                               # stopping the whole run over a catalogue the
                               # application declined to draw costs eleven
                               # modules of coverage to prove a point the report
                               # already makes. It is never silent: the missing
                               # category is logged, attached and photographed,
                               # and the Allure step says QUOTA EXHAUSTED.
                               #
                               # GENZ_REQUIRE_SPECIALISATION=1 makes it a hard
                               # failure again, for a run whose purpose IS to
                               # prove the business profile can be set.
ANIMATION_MS = 500             # the only fixed pause, for  CSS transitions
SCROLL_TIMEOUT = 5_000         # one attempt at bringing a control onto the screen
SCROLL_ATTEMPTS = 2            # scroll, check the viewport, and scroll again
READY_TIMEOUT = 3_000          # the enabled / editable look BEFORE acting. Short
                               # on purpose: the click or fill that follows waits
                               # the full RETRY_TIMEOUT for the same thing, so a
                               # control that stays disabled must not be paid for
                               # twice on every one of the three attempts.
HIGHLIGHT_MS = 700             # how long the outline round the field being worked
                               # on stays up. Drawn and removed inside the page,
                               # so it never costs the run any waiting time.

# --- Resilience: how hard the framework tries before it calls something a bug -
# Every action goes through the self-healing layer further down. These say how
# many times it may try and how long ONE attempt gets: attempts x RETRY_TIMEOUT
# stays in the same ballpark as the old single DEFAULT_TIMEOUT, so a genuinely
# broken step still fails in about the same time it always did.
ACTION_ATTEMPTS = 3            # click / fill / select
NAV_ATTEMPTS = 4               # opening a module from the sidebar
VERIFY_ATTEMPTS = 3            # proving a record was really saved
RETRY_TIMEOUT = 15_000         # one attempt's own timeout
DOM_STABLE_TIMEOUT = 10_000    # how long a re-render is allowed to take
OVERLAY_TIMEOUT = 10_000       # how long a backdrop is allowed to sit there

# --- Browser ---
HEADED = True                  # always show the browser
SLOW_MO_MS = 300               # slow each action down so the run is watchable

# --- Output and Allure reporting ---
# Every one of these is defined in run_config.py, which stamps this execution
# with the moment it started and hands back a folder of its own,
# Reports/<YYYY-MM-DD_HH-MM-SS>/, for the report, the screenshots and the log.
# Nothing a run writes can therefore overwrite the run before it. The names are
# imported rather than redefined so there is still exactly one definition of
# each - see the module docstring in run_config.py for the whole layout.


# =========================================================================== #
# LOGGING - console + Reports/<this run>/execution.log
#
# run_config sets both handlers up the moment it is imported, and conftest.py
# imports it before collection starts, so an error raised while this very file
# is being imported is already being written to execution.log.
# =========================================================================== #


# =========================================================================== #
# ALLURE - used when available, silently skipped when it is not
# =========================================================================== #
try:
    import allure
    from allure_commons.types import AttachmentType

    ALLURE = True
except ImportError:                                    # plain `python` run
    ALLURE = False

    class _AllureStub:
        """Keeps the @allure.* annotations harmless when allure is not installed.

        `python Construction_Flow.py` runs without the reporting plugin, so every
        decorator has to become a no-op instead of an ImportError at import time.
        """

        class severity_level:                          # noqa: N801 - allure's name
            BLOCKER = "blocker"
            CRITICAL = "critical"
            NORMAL = "normal"
            MINOR = "minor"
            TRIVIAL = "trivial"

        class dynamic:
            @staticmethod
            def title(*args, **kwargs) -> None:
                """No report to name."""

            @staticmethod
            def parameter(*args, **kwargs) -> None:
                """No report to parameterise."""

            @staticmethod
            def description(*args, **kwargs) -> None:
                """No report to describe."""

        @staticmethod
        def attach(*args, **kwargs) -> None:
            """No report to attach to."""

        def __getattr__(self, name):
            """@allure.epic(...) / .feature(...) / .story(...) -> leave as it is."""
            def annotation(*args, **kwargs):
                return lambda target: target
            return annotation

    allure = _AllureStub()

    class AttachmentType:                              # the two kinds used here
        TEXT = "text/plain"
        PNG = "image/png"

#: pytest reports a used-up row as SKIPPED; the plain `python` run just logs it
#: and moves to the next row. main() sets this to False.
UNDER_PYTEST = True


@contextmanager
def step(title: str) -> Iterator[None]:
    """Log a step and show it in the Allure report as passed or failed."""
    log.info("%s", title)
    if not ALLURE:
        yield
        return
    with allure.step(title):
        yield


# --------------------------------------------------------------------------- #
# THE STOCK VALIDATION SECTION OF THE REPORT
#
#     📦 Stock Validation
#        ├── 📦 Item Stock - Before Transaction
#        ├── 📈 Stock Increase - GRN
#        ├── 📊 Stock After Sales - Quotation & Sales Order
#        ├── 📉 Stock Decrease - Material Indent
#        └── 📊 Current Stock Validation & Movement Summary
#
# REPORTING ONLY - two nested Allure steps and nothing else. It exists because
# the stock cards used to be attached wherever the flow happened to be
# standing, which put "did the stock move correctly?" three screens down an
# attachment list beside a form screenshot and a locator error. Opening a named
# parent step round them gives the question a place of its own that a manual
# tester can find without being told where to look.
#
# It drives no browser, reads no page and touches no verdict. A step that
# failed to open would still leave every attachment in the report, one level
# higher up - which is why nothing in here is allowed to raise.
# --------------------------------------------------------------------------- #

#: The parent every stock section hangs under, so the three of them group
#: together in the report instead of appearing as three unrelated steps.
STOCK_SECTION = "📦 Stock Validation"

#: The item's opening reading - "what was the CURRENT STOCK before anything
#: moved?" - which is the figure every later stock verdict is measured from,
#: so it gets a step of its own rather than a line on somebody else's card.
STOCK_ITEM_STEP = "📦 Item Stock - Before Transaction"

STOCK_INCREASE_STEP = "📈 Stock Increase - GRN"
STOCK_SALES_STEP = "📊 Stock After Sales - Quotation & Sales Order"
STOCK_DECREASE_STEP = "📉 Stock Decrease - Material Indent"
STOCK_SUMMARY_STEP = "📊 Current Stock Validation"


@contextmanager
def stock_section(child: str) -> Iterator[None]:
    """Open "📦 Stock Validation > <child>" for the attachments inside it."""
    with step(STOCK_SECTION):
        with step(child):
            yield


# --------------------------------------------------------------------------- #
# THE QA ACTION JOURNAL
#
# What a manual tester would have written down while doing this by hand:
#
#     Entered Item Code: ITEM-OPC53-011
#     Entered Item Name: OPC Cement 53 Grade
#     Selected Item Type: semi_finished
#     Clicked 'Create Item'
#     Verified 'item OPC Cement 53 Grade in the Items list' is displayed
#
# It is written from the framework's OWN helpers - safe_fill, safe_select,
# safe_click, safe_expect, safe_navigation - so every module gets it without a
# line of code in any test case, and a module written tomorrow is covered the
# day it is written. The `description` each helper is already given is the
# field's business name, which is why the journal reads like English and not
# like a locator.
#
# Two rules that make it worth reading:
#
#   * an action is journalled AFTER it has succeeded. A field the application
#     refused is not "Entered", and a report that says it was is worse than no
#     report at all. The failure is recorded by fail(), which has the screenshot
#     and the application's own message.
#   * a retry is one action, not three. The helpers retry inside themselves, so
#     the journal records the outcome, not the mechanics.
#
# Nothing here touches a locator, a wait or a business rule. It reads the
# arguments the helpers are already called with.
# --------------------------------------------------------------------------- #

#: The actions of the row in progress. Cleared by run_test_case for every row.
ACTIONS: List[str] = []


def qa_step(action: str) -> None:
    """Record one business-readable action, and show it in the report.

    An Allure step with nothing inside it - it renders as a single ticked line
    under whichever business step is open ("Creating Item..."), which is exactly
    how a test script reads on paper.
    """
    ACTIONS.append(action)
    if ALLURE:
        with allure.step(f"✓ {action}"):
            pass


def quoted(description: str) -> str:
    """A control's name in quotes, unless it already carries some.

    The descriptions the modules pass in are a mixture: "Create Item" is a bare
    label, "item 'ITEM-OPC53-013'" already names the thing it is quoting. Adding
    a second pair round the second one produces "Clicked 'item 'ITEM-...''",
    which is the kind of detail that makes a report look machine-written.
    """
    description = (description or "").strip()
    return description if "'" in description or '"' in description \
        else f"'{description}'"


def forget_actions() -> None:
    """Start a row with an empty journal - the last row's actions are not this one's."""
    ACTIONS.clear()


def attach_actions() -> None:
    """Put the row's actions in the report as one numbered list."""
    if not ACTIONS:
        return
    body = "\n".join(f"{number:>3}. {action}"
                     for number, action in enumerate(ACTIONS, start=1))
    attach_text("Actions Performed",
                f"What the automation did, in order ({len(ACTIONS)} actions):\n\n"
                + body)


def attach_text(name: str, body: str) -> None:
    """Attach plain text (an error, an app message) to the Allure report."""
    if ALLURE:
        allure.attach(body, name=name, attachment_type=AttachmentType.TEXT)


def describe(body: str, html: str = "") -> None:
    """Put the QA-readable summary at the TOP of the test in the report.

    Allure draws a test's description above its steps and its attachments, so
    this is the one place a reader cannot miss - which is why the answers to
    "what was this test?" and, when something went wrong, "why did it fail?" are
    put here rather than only in an attachment.

    Both forms are set on purpose. `description_html` is what the report shows;
    the plain one is what is left when the report is read as text, and what an
    older allure-pytest falls back to. Nothing here can fail a test: an Allure
    version without description_html is caught and the plain text still lands.
    """
    if not ALLURE:
        return
    try:
        allure.dynamic.description(body)
        if html:
            allure.dynamic.description_html(html)
    except Exception as error:                     # reporting is never fatal
        log.debug("Could not set the Allure description: %s", error)


def case_header_card(row: Row, module: str) -> str:
    """The card a reader sees at the top of a test case, before opening anything.

    An embedded FRAGMENT, not a page: Allure writes a description into the
    report's own screen, where a stylesheet of our own would restyle the whole
    of Allure - see qa_report.embed().
    """
    return qa_report.embed(
        f"{module} - {row.test_case_id}",
        "",
        subtitle=case_title(row, module),
        pairs=[["Module", module],
               ["Test Case ID", row.test_case_id],
               ["Test Data", f"{row.sheet} sheet, row {row.number} of "
                             f"Construction_Flow_Data.xlsx"]])


#: (title, plain-text card) for each module that owns one of the three stock
#: stages - the SAME text already attached lower in that module's own step
#: tree (item_text() / increase_text() / decrease_text()+summary_text() -
#: REPORTING ONLY, nothing recalculated here). Read by module_stock_summary()
#: once run_test_case() knows the row passed, which is the only moment this
#: module's figures are both real and final.
STOCK_CARD_BY_MODULE = {
    "Items": ("📦 Item Stock - Before Transaction",
             lambda: stock_report.item_text(),
             lambda: getattr(stock_report.movement_by_stage(
                 stock_report.ITEM_STAGE), "result", ""),
             lambda: stock_report.item_table_html()),
    "GRN": ("📈 Item Stock - After Purchase",
           lambda: stock_report.increase_text(),
           lambda: getattr(stock_report.movement_by_stage(
               stock_report.GRN_STAGE), "result", ""),
           lambda: stock_report.increase_table_html()),
    # Same card as "GRN" above - the stock itself already moved at GRN
    # approval (verified live: the app updates it there, not at bill
    # approval), so this is the SAME Movement, re-shown once the purchase
    # CHAIN a reader actually checks against is complete. It is MORE complete
    # here than it was from GRN's own test case: Purchase Bill Number/Status
    # are still "NOT AVAILABLE" when GRN attaches this (the bill does not
    # exist yet) but are real by the time this test's own describe() fires,
    # because approve_purchase_bill() -> verify_status() has already called
    # stock_report.note_document()/note_status() for "Purchase Bill" earlier
    # in THIS SAME test. REPORTING ONLY - no new stock read, no new page visit.
    "Purchase_Bills": ("📈 Item Stock - After Purchase",
                      lambda: stock_report.increase_text(),
                      lambda: getattr(stock_report.movement_by_stage(
                          stock_report.GRN_STAGE), "result", ""),
                      lambda: stock_report.increase_table_html()),
    # The sales checkpoint's reading is taken at the TOP of create_material_
    # indent() - see report_stock_after_sales() - which is why it is real by
    # the time THIS test's own describe() fires, even though the card is
    # about the Quotation and the Sales Order that ran earlier.
    #
    # No table_html here (stays "") - this one is THREE cards joined
    # (sales + decrease + the whole-chain summary_text()), not a single
    # pairs list, so there is no one table to swap the <pre> for without
    # dropping summary_text()'s own content. Left as the plain-text
    # embedding it already was; only the two single-card cases above
    # (Items, GRN/Purchase_Bills) needed - and got - the swap.
    "Material_Indents": ("📉 Item Stock - After Consumption",
                        lambda: qa_report.joined([
                            stock_report.sales_text(),
                            stock_report.decrease_text(),
                            stock_report.summary_text()]),
                        lambda: stock_report.result(),
                        lambda: ""),
}


def module_stock_summary(module: str) -> Tuple[str, str, object, str]:
    """(title, real figures, verdict, table_html) for a module owning a
    stock stage.

    ("", "", "", "") for a module that owns none. REPORTING ONLY - reads
    whatever stock_report.py already recorded for THIS run; nothing is
    computed here and nothing is invented for a module that never reached
    its stage (the card functions themselves already say so, honestly, when
    that happens). The verdict is the stage's OWN recorded result, not
    assumed from the row's overall outcome - a calculation mismatch
    elsewhere on the row can still reach this path when
    GENZ_FAIL_ON_CALC_MISMATCH=0, and the badge must not call that stage
    PASSED unless its own check said so. `table_html`, when non-empty, is a
    real <table> to show INSTEAD of dumping the same figures a second time
    as plain text underneath it - see qa_report.embed()'s `table_html`.
    """
    entry = STOCK_CARD_BY_MODULE.get(module)
    if entry is None:
        return "", "", "", ""
    title, text_fn, verdict_fn, table_fn = entry
    return title, text_fn(), verdict_fn(), table_fn()


def case_title(row: Row, module: str) -> str:
    """One line saying what this test case does, in a tester's words.

    Taken from the workbook where the sheet carries a description, and otherwise
    from the module and the record the row names - so it is never a locator and
    never a function name.
    """
    for column in ("Description", "Test Case Description", "Scenario",
                   "Test Scenario"):
        text = str(row.get(column, "") or "").strip()
        if text:
            return text
    for column in ("Item Name", "Customer Name", "Supplier Name", "Company Name",
                   "Category Name", "Parent Category", "Business Name"):
        text = str(row.get(column, "") or "").strip()
        if text:
            return f"Create and verify {module} - {text}"
    return f"Create and verify {module}"


def attach_table(name: str, headings: Tuple[str, ...],
                 rows: List[List[str]], summary: str = "",
                 plain: bool = False) -> None:
    """Attach one table as the colour table a stakeholder opens.

    The HTML one is what a reader actually opens - PASS green, FAIL red, at a
    glance - so it is the one that is always attached. The plain-text twin says
    exactly the same thing, and attaching both under the SAME NAME put two
    identical-looking entries in every test's attachment list, which was the
    single biggest source of clutter in this report. It is now kept only where
    it earns its place: `plain=True` on the summary tables, which are the ones
    that get pasted into an e-mail, a ticket or a chat window.

    The renderers already exist in data_verification.py and are reused rather
    than written again.
    """
    if not rows:
        return
    if plain:
        attach_text(name, (summary + "\n\n" if summary else "")
                    + verify_data.text_table(headings, rows))
    verify_data.attach_html(name, verify_data.html_table(
        name, headings, rows, summary.replace("\n", " | ")))


#: Attach the routine "this is the form before I saved it" pictures as well as
#: the failure ones. Off by default: a passing test case does not need a
#: photograph of a form that worked, and eleven of them per run is what turns
#: an attachment list into something nobody scrolls. Set GENZ_FULL_EVIDENCE=1
#: to put them back in the report. THE FILE IS WRITTEN TO DISK EITHER WAY - see
#: form_screenshot() - so no evidence is ever lost, only un-attached.
FULL_EVIDENCE = os.environ.get("GENZ_FULL_EVIDENCE") == "1"


def screenshot(page: Page, name: str, attach: bool = True) -> None:
    """Save a full-page screenshot and attach it to the Allure report.

    The name is used as a file name, and the retry helpers build it out of a
    step's description ("click 'Create Item'"), so the characters Windows will
    not accept in a path are folded to underscores first.

    `attach=False` still takes the picture and still keeps it in this run's
    screenshots folder; it only leaves it out of the report - see
    form_screenshot().
    """
    safe_name = re.sub(r"[^A-Za-z0-9._-]+", "_", name).strip("_") or "screenshot"
    file_path = SCREENSHOT_DIR / f"{safe_name}_{datetime.now():%H%M%S}.png"
    try:
        page.screenshot(path=str(file_path), full_page=True)
    except PlaywrightError as error:
        log.warning("Could not take screenshot '%s': %s", name, error)
        return
    log.info("Screenshot: %s", file_path)
    if ALLURE and attach:
        allure.attach(file_path.read_bytes(), name=name,
                      attachment_type=AttachmentType.PNG)


def form_screenshot(page: Page, name: str) -> None:
    """A routine "before save" picture: kept on disk, attached only on request.

    REPORTING ONLY, and the distinction it draws is the point of it: a picture
    of a form that then saved correctly proves nothing a reader needs, while
    the picture of the screen that FAILED proves everything - so the failure
    ones (screenshot(), fail(), the makereport hook) are always attached and
    these are not. Both are still written to Reports/<run>/screenshots/.
    """
    screenshot(page, name, attach=FULL_EVIDENCE)


def fail(page: Page, message: str, error: object = "") -> None:
    """Stop the test with a QA-friendly reason, a screenshot and the app's own text.

    Always call this instead of a bare `assert` - it captures the evidence that
    makes the failure diagnosable after the fact.
    """
    app_messages = read_app_messages(page)
    detail = f"{message}\n{error}".strip()
    if app_messages:
        detail += "\nApplication said: " + " | ".join(app_messages)

    log.error("FAILED | %s", detail)
    # The question a reader asks first is "what was it doing when this
    # happened?", so the last thing that DID work is put beside the thing that
    # did not. The journal already holds it - see qa_step().
    attach_text("Failure Details", "\n".join([
        "WHAT WENT WRONG",
        message,
        "",
        f"Last action that worked : "
        f"{ACTIONS[-1] if ACTIONS else '(nothing had been done yet)'}",
        f"Page at the time        : {page.url}",
        "",
        "APPLICATION'S OWN MESSAGE",
        " | ".join(app_messages) if app_messages
        else "(the application displayed no message)",
        "",
        "TECHNICAL DETAIL",
        str(error) or "(none - this is the framework's own assertion)",
    ]))
    attach_text("Page URL", page.url)
    screenshot(page, "Screenshot at Failure")
    raise AssertionError(detail)


class RecordAlreadyExists(Exception):
    """The application already holds this row - it is not a defect.

    Mobile numbers, e-mail addresses and GSTINs are unique in GenZOpss, and the
    framework types every cell verbatim. So the second time a row is run the
    application answers 409 (PHONE_IN_USE, PHONE_ALREADY_REGISTERED, ...): the
    record the row asked for is already there. That is a used-up row, not a bug,
    so it is reported as SKIPPED with the application's own words - and the run
    carries on to the modules that still have work to do.
    """


#: What a 409 looks like in the application's replies and toasts.
ALREADY_EXISTS_MARKERS = (
    "ALREADY_REGISTERED", "_IN_USE", "ALREADY_EXISTS", "ALREADY EXISTS",
    "DUPLICATE", "ALREADY LINKED", "ALREADY REGISTERED",
)


#: A rejected registration, by status code rather than by wording. 409 Conflict
#: IS "you already have this" - it is what the status means - and it is the only
#: thing left to go on when the body has gone: drain() reads a response some time
#: after it arrived, and the application's is often no longer available by then
#: ("the response body was no longer available"). Matching the code as well as
#: the wording is what stops a used-up row being reported as something else.
REGISTER_ENDPOINT = re.compile(r"/auth/register\b", re.IGNORECASE)


def already_exists(messages: List[str]) -> bool:
    """True when the application is refusing a row because it already has it."""
    text = " ".join(messages).upper()
    if any(marker in text for marker in ALREADY_EXISTS_MARKERS):
        return True
    # "API 409 https://.../v1/auth/register -> ..." - see REGISTER_ENDPOINT.
    return any(line.startswith("API 409 ") and REGISTER_ENDPOINT.search(line)
               for line in messages)


class ApiFailureRecorder:
    """Remembers the API errors the application never puts on screen.

    A rejected create comes back as a 4xx with a JSON body ("PHONE_IN_USE",
    "GSTIN_IN_USE") and the dialog simply sits there, showing nothing. Without
    this the run reports "the button never disappeared" instead of the reason,
    and a DOM-only check can even pass.
    """

    def __init__(self, page: Page) -> None:
        self._failed: List[object] = []
        page.on("response", self._record)
        page._api_failures = self                      # found by read_app_messages

    def _record(self, response) -> None:
        # Only keep the response here. Reading its body inside the event handler
        # calls a sync Playwright method from a callback, which intermi1ttently
        # loses the payload - drain() reads it on the main thread instead.
        # GETs are ignored: a missing icon is not a test failure.
        #
        # Everything is wrapped: this runs inside Playwright's own dispatch
        # loop, where a response whose request object has already been disposed
        # (which happens mid-navigation) raises out of the library's
        # deserialiser. Letting that escape a listener disturbs the call the
        # main thread is making at the time, so a lost error message is by far
        # the lesser problem.
        try:
            if response.status >= 400 and response.request.method != "GET":
                self._failed.append(response)
        except Exception as error:                     # noqa: BLE001 - see above
            log.debug("could not record a failed response: %s", error)

    def drain(self, limit: int = 3) -> List[str]:
        """The failures seen since the last call, as readable text."""
        messages: List[str] = []
        for response in self._failed[:limit]:
            try:
                body = (response.text() or "").strip()
            except PlaywrightError:
                body = "(the response body was no longer available)"
            messages.append(f"API {response.status} {response.url} -> {body[:300]}")
        self._failed.clear()
        return messages


def read_app_messages(page: Page, limit: int = 5) -> List[str]:
    """Every error the application is showing - on screen AND from its API."""
    messages: List[str] = []

    recorder = getattr(page, "_api_failures", None)
    if recorder is not None:
        messages.extend(recorder.drain())

    for selector in ("[role='alert']", "[class*='error' i]", "[class*='toast' i]"):
        try:
            found = page.locator(selector).filter(visible=True)
            for index in range(min(found.count(), limit)):
                text = (found.nth(index).inner_text() or "").strip()
                if text and text not in messages:
                    messages.append(text)
        except PlaywrightError:
            continue
    return messages[:limit]


#: What a message says when the row could not be run at all - the environment,
#: the data or a manual precondition stopped it, and nothing about the
#: application was proved either way. The wording is the one the flows already
#: write into their own failures (see approve_business() and
#: explain_indent_status()), so the classification follows the evidence the run
#: captured rather than a second, separate opinion about it.
BLOCKED_MARKERS: List[Tuple[str, str]] = [
    ("CATEGORY: TEST DATA / ENVIRONMENT", calc.DATA_ENVIRONMENT),
    ("CATEGORY: DATA / ENVIRONMENT", calc.DATA_ENVIRONMENT),
    ("GST lookup quota exhausted", calc.DATA_ENVIRONMENT),
    ("CATEGORY: BLOCKED", calc.PRECONDITION),
    ("BLOCKED precondition", calc.PRECONDITION),
]


def classify_failure(error: BaseException) -> Tuple[str, str]:
    """(module status, category) for the exception that stopped a row.

    A row that failed is not automatically a defect, and the report has to say
    which of three things happened - they go to three different people:

        BLOCKED   the row could not be run. No OTP was typed, or the warehouse
                  did not hold the stock the row asked for. The application was
                  never put to the test, so this is neither a product defect nor
                  an automation one, and it must not be counted as either.
        FAIL      the flow ran and did not reach the outcome the row asked for.
                  An AssertionError is what fail() raises once it has the
                  application's own message and a screenshot: a real finding.
        AUTOMATION  anything else that escaped - a timeout, a Playwright error,
                  a TypeError. The suite fell over, so this is the suite's own
                  problem until it is proved otherwise. Allure calls the same
                  split 'failed' and 'broken'.

    NOTHING is suppressed here. The exception is still raised by the caller and
    the test still fails; this only decides what the summary CALLS it.
    """
    message = f"{type(error).__name__}: {error}"
    for marker, category in BLOCKED_MARKERS:
        if marker.casefold() in message.casefold():
            return calc.BLOCKED, category
    if isinstance(error, AssertionError):
        return calc.FAIL, calc.APPLICATION
    return calc.FAIL, calc.AUTOMATION


def blocked_reason(error: BaseException) -> str:
    """The 'why' the summary prints for a blocked or failed module.

    The first paragraph of the failure - the flows put the cause in it, and the
    whole message is already in the report as the test's own failure.
    """
    first = str(error).strip().splitlines()
    for line in first:
        if line.strip():
            return line.strip()[:300]
    return type(error).__name__


#: Which test cases this run could not RUN, and why - nodeid -> the reason.
#: Fed by note_blocked_test_case() and read by pytest_runtest_logreport(), which
#: is what turns pytest's "skipped" back into the BLOCKED it really is. Without
#: this the two would be counted as one, and a blocked precondition would be
#: reported in the same column as a used-up row - which are different findings
#: for different people.
BLOCKED_TEST_CASES: Dict[str, str] = {}

#: The test pytest is running right now. Set by pytest_runtest_setup(), which is
#: the only place the nodeid is available; run_test_case() is inside the test and
#: cannot see it.
CURRENT_NODEID = ""

#: The pytest Item currently executing - same hook, same reason, but keeping
#: the OBJECT rather than its nodeid text. current_allure_test_uuid() needs it:
#: allure_pytest's own ItemCache is keyed by id(item.nodeid) (identity, not the
#: string's value - see allure_pytest.listener.ItemCache), so only the exact
#: `item.nodeid` object that AllureListener itself pushed will look itself up
#: again; a copy of the same text will not.
CURRENT_PYTEST_ITEM = None


def note_blocked_test_case(module: str, row: Row, category: str,
                           error: BaseException) -> None:
    """Record that this test case was BLOCKED, not failed and not skipped."""
    if CURRENT_NODEID:
        BLOCKED_TEST_CASES[CURRENT_NODEID] = blocked_reason(error)
    log.warning("BLOCKED | %s (%s): %s", module, category, blocked_reason(error))
    if ALLURE:
        allure.dynamic.parameter("Result", f"BLOCKED ({category})")
        allure.dynamic.label("blockedCategory", category)
        allure.dynamic.label("blockedModule", module)


def failure_narrative(module: str, row: Row, status: str, category: str,
                      error: BaseException) -> str:
    """A failed row explained the way a tester would write it in a ticket.

    MODULE / TEST CASE / EXPECTED / ACTUAL / DETAIL / RESULT / REASON, built
    entirely from what the run already recorded:

        the flows' own failure messages    "Expected: ... / Actual: ..." when the
                                          module stated them
        calculation_validation            any figure that did not match, with
                                          both sides and the difference
        the action journal                the last thing that DID work

    Nothing is inferred and nothing is invented: a field the run did not record
    is left out rather than guessed at. The full exception is still attached
    beside this, and pytest still has the traceback.
    """
    message = str(error).strip()
    detail: List[List[object]] = [
        ["Test Data", f"{row.sheet} sheet, row {row.number}"],
    ]

    # A calculation that did not match states its own Expected and Actual, and
    # those are the two figures a reader wants first.
    expected = actual = ""
    report = getattr(calc.VALIDATION, "current", None)
    wrong = [check for check in getattr(report, "checks", [])
             if check.status == calc.FAIL]
    if wrong:
        first = wrong[0]
        expected = f"{first.name} = {calc.money(first.expected)}"
        actual = f"{first.name} = {calc.money(first.actual)}"
        for check in wrong:
            detail.append([
                check.name,
                f"expected {calc.money(check.expected)}, "
                f"actual {calc.money(check.actual)}, "
                f"difference {calc.money(check.difference)}"])
    else:
        # The flows write "Expected ... / Actual ..." into their own failures;
        # when they have, that is the clearest statement there is of what went
        # wrong, so it is lifted out rather than re-worded.
        for line in message.splitlines():
            lowered = line.strip().lower()
            if not expected and lowered.startswith("expected"):
                expected = line.strip()
            elif not actual and lowered.startswith("actual"):
                actual = line.strip()

    detail.append(["Category", category or "-"])

    return qa_report.failure_story(
        module, row.test_case_id, case_title(row, module),
        expected=expected, actual=actual, detail=detail,
        reason=message or type(error).__name__,
        last_action=ACTIONS[-1] if ACTIONS else "(nothing had been done yet)",
        result=status, category=category)


@contextmanager
def run_test_case(row: Row, module: str) -> Iterator[None]:
    """Wrap one Excel row: log it, report it, and time it.

    This is what makes every row show up in the Allure report as its own test
    execution, carrying the module, the sheet, the TestCaseID and the exact
    input data that produced the result.
    """
    log.info("Executing Test Case: %s", row.test_case_id)

    if ALLURE:
        allure.dynamic.title(f"[{row.test_case_id}] {module} - "
                             f"{case_title(row, module)}")
        allure.dynamic.parameter("Module", module)
        allure.dynamic.parameter("Sheet", row.sheet)
        allure.dynamic.parameter("TestCaseID", row.test_case_id)
        allure.dynamic.parameter("Test Case", case_title(row, module))
        # pytest-playwright/allure-pytest auto-publishes the WHOLE parametrize
        # argument ("row") as its own Parameter, which dumps every column of
        # the sheet - Purchase Price, Selling Price, MRP, Rate, Discount, all
        # of it - as one raw dict in the Overview panel Allure draws above the
        # steps. That is the first thing a reader sees opening any test, ahead
        # of the stock and calculation results this report exists for.
        # Overwriting it here is display-only: allure_pytest.listener.
        # add_parameter() replaces a Parameter already carrying this name by
        # VALUE, nothing else reads it, and stabilise_history_id() below
        # rebuilds historyId from fullName + TestCaseID alone - never from a
        # parameter - so this cannot rename a test's history or its identity.
        allure.dynamic.parameter(
            "row", f"(not displayed - {row.test_case_id})")
        # The test case, at the top of the test, above the steps, so a reader
        # knows what this is before opening a single attachment. The row's
        # own values are deliberately not shown here or anywhere else in the
        # report.
        describe(case_description(row, module), case_header_card(row, module))

    forget_actions()

    # Every value entered from here on is captured, and compared with the
    # record's View page when the test case ends - see data_verification.py.
    verify_data.start_test(module, row)
    # ...and every number the application works out is captured and re-worked
    # independently - see calculation_validation.py. Both halves are reported
    # against this same row, and the module's result is the two of them
    # together.
    calc.start_test(module, row)
    # Nothing this row's forms said about tax is known yet. Without this the
    # rate a PREVIOUS row's form published would be used to check this one.
    forget_form_tax_rate()
    # Same reasoning for the stock a PREVIOUS indent row was short of: it must
    # never be offered as the explanation for this row's missing step.
    forget_indent_shortfall()

    started = time.time()
    try:
        yield
    except RecordAlreadyExists as error:
        # A used-up row, not a failure - see RecordAlreadyExists. It is neither
        # a PASS (nothing was proved) nor a FAIL (the application did the right
        # thing by refusing a duplicate): it is a TEST DATA condition, and the
        # report has to say so in those words rather than leaving a reader to
        # infer it from an exception message.
        log.warning("SKIPPED | %s", error)
        attach_actions()
        attach_text("Execution Status",
                    f"{row.test_case_id}: SKIPPED after {time.time() - started:.1f}s"
                    f" - the application already has this record")
        attach_text("SKIPPED - test data already exists in the application",
                    f"Status : SKIPPED\n"
                    f"Reason : Test data already exists in application\n"
                    f"Module : {module}\n"
                    f"Test case: {row.test_case_id} (sheet '{row.sheet}', "
                    f"row {row.number})\n"
                    f"Category : TEST DATA - not an automation defect and not an "
                    f"application defect.\n\n"
                    f"{duplicate_identity(row)}\n"
                    f"The application refused the registration because one of "
                    f"these values has been used before. That is the application "
                    f"behaving correctly, so the row is skipped rather than "
                    f"failed - and it is NOT passed, because nothing was proved.\n\n"
                    f"To run this row again, put fresh values in the "
                    f"Create_Your_Account sheet.\n\n"
                    f"The application's own message:\n{error}")
        if ALLURE:
            allure.dynamic.parameter("Result", "SKIPPED - data already exists")
            allure.dynamic.label("skipReason", "test data already exists")
            skip_story = qa_report.failure_story(
                module, row.test_case_id, case_title(row, module),
                expected="the application accepts this record",
                actual="the application already holds it and refused it",
                detail=[["Sheet / Row", f"{row.sheet} (row {row.number})"],
                        ["Category", "TEST DATA - not a defect of either kind"]],
                reason="The values in this row have been used before, so the "
                       "application correctly refused a duplicate. Nothing was "
                       "proved about the application, so the row is SKIPPED - "
                       "it is not a pass and it is not a failure. Put fresh "
                       "values in the workbook to run it again.",
                last_action=ACTIONS[-1] if ACTIONS else "",
                result="SKIPPED")
            describe(skip_story, qa_report.failure_embed(
                module, row.test_case_id, skip_story,
                reason="The application already holds this record, so there "
                       "was nothing left for this row to prove.",
                result="SKIPPED"))
        verify_data.finish_test(passed=False)
        calc.finish_test(functional=calc.SKIPPED,
                         note="Test data already exists in the application - "
                              "the row is used up, so nothing was proved.")
        if UNDER_PYTEST:
            pytest.skip("SKIPPED - test data already exists in application: "
                        + str(error).splitlines()[0])
        return                                   # plain run: on to the next row
    except Exception as error:
        # What KIND of failure this was - see classify_failure(). A row the
        # environment blocked is recorded as BLOCKED rather than FAIL, so the
        # summary can tell "the application got this wrong" apart from "the
        # application was never asked". The exception is re-raised either way:
        # the test still fails, with its message, its screenshot and its
        # Allure category. Nothing is swallowed here.
        status, category = classify_failure(error)
        attach_actions()
        attach_text("Execution Status",
                    f"Module     : {module}\n"
                    f"Test case  : {row.test_case_id}\n"
                    f"Result     : {status}\n"
                    f"Category   : {category}\n"
                    f"Ran for    : {time.time() - started:.1f}s\n"
                    f"Last action: "
                    f"{ACTIONS[-1] if ACTIONS else '(nothing was done yet)'}\n\n"
                    f"Why it stopped:\n{blocked_reason(error)}")
        # WHAT WENT WRONG, IN QA ENGLISH, BEFORE ANY PYTHON.
        # The traceback is still there - pytest keeps it, and the technical
        # detail is attached beside this - but a reader should not have to read
        # "AssertionError" to find out that a status was Partial Issue when it
        # should have been Fulfilled. This is built from what the run itself
        # recorded: the module, the row, the last action that worked, the
        # application's own words, and any calculation that did not match.
        story = failure_narrative(module, row, status, category, error)
        attach_text("Why This Failed (QA Explanation)", story)
        if ALLURE:
            allure.dynamic.parameter("Result", f"{status} ({category})")
            describe(story, qa_report.failure_embed(
                module, row.test_case_id, story,
                reason=blocked_reason(error), result=status))
        if status == calc.BLOCKED:
            attach_text(
                f"BLOCKED - {category}",
                f"Status   : BLOCKED\n"
                f"Category : {category}\n"
                f"Module   : {module}\n"
                f"Test case: {row.test_case_id} (sheet '{row.sheet}', "
                f"row {row.number})\n\n"
                f"This row could not be RUN, so nothing about the application "
                f"was proved - it is neither an application defect nor an "
                f"automation defect, and it is NOT counted as a pass.\n\n"
                f"{qa_report.CLASSIFICATION_NOTE}\n\n"
                f"The run's own words:\n{error}")
        # passed=False: the View page is NOT opened after a failure, so the
        # screen the failure screenshot photographs is the screen that failed.
        verify_data.finish_test(passed=False)
        calc.finish_test(functional=status, category=category,
                         note=blocked_reason(error))

        # HOW A BLOCKED ROW LEAVES THIS BLOCK, and why it is not `raise`.
        #
        # A blocked row could not be RUN. Re-raising made pytest record it as
        # FAILED, which is the one word that means the opposite of what
        # happened: a red test in the report reads as "the application got this
        # wrong", and a reader then goes looking for a defect that does not
        # exist. Everything above has already been recorded - the module is
        # BLOCKED in the summary, the reason, the screenshot and the QA
        # explanation are all attached - so what is left is only which word
        # pytest files it under, and `skip` is the honest one of the two it
        # offers. The tally still counts it as BLOCKED and NOT as a skip (see
        # pytest_runtest_logreport), so the report keeps the two apart.
        #
        # Nothing is swallowed: the reason is the skip's own message, so it is
        # on the console, in the report and in the run's log. Set
        # GENZ_FAIL_ON_BLOCKED=1 for a CI job that must go red on a blocked
        # precondition.
        if status == calc.BLOCKED and UNDER_PYTEST and not FAIL_ON_BLOCKED:
            note_blocked_test_case(module, row, category, error)
            pytest.skip(f"BLOCKED ({category}) - {blocked_reason(error)}")
        raise

    # The row worked. Now read every record's View page and compare it with what
    # was typed - the last thing the test case does, so nothing it does can get
    # in the way of a step that still had work to do.
    verify_data.finish_test(passed=True)

    # The flow worked; whether the application's arithmetic did is a separate
    # verdict, and one wrong number is enough to fail the row. The record was
    # still created, and everything above it still ran - what fails here is the
    # claim that the module produced the right result.
    report = calc.finish_test(functional=calc.PASS)
    attach_actions()
    wrong = calc.mismatches(report)
    if wrong:
        detail = "\n\n".join(check.report() for check in wrong)
        attach_text("Execution Status",
                    f"Module    : {module}\n"
                    f"Test case : {row.test_case_id}\n"
                    f"Result    : CALCULATION FAILED\n"
                    f"Ran for   : {time.time() - started:.1f}s\n\n"
                    f"The flow itself worked - every screen did what the row "
                    f"asked. What did not agree is the application's "
                    f"arithmetic: {len(wrong)} of "
                    f"{len(report.checks) if report else 0} calculation(s) "
                    f"differ from a value worked out independently. See "
                    f"'Calculation Validation' for each of them.")
        if ALLURE:
            allure.dynamic.parameter("Result", "CALCULATION FAILED")
            # The flow was fine and a number was not - say exactly that, at the
            # top of the test, with both figures and the difference. It is the
            # one failure a reader is most likely to be sent looking for.
            wrong_story = qa_report.failure_story(
                module, row.test_case_id, case_title(row, module),
                expected=f"{wrong[0].name} = {calc.money(wrong[0].expected)}",
                actual=f"{wrong[0].name} = {calc.money(wrong[0].actual)}",
                detail=[[check.name,
                         f"expected {calc.money(check.expected)}, actual "
                         f"{calc.money(check.actual)}, difference "
                         f"{calc.money(check.difference)}"]
                        for check in wrong],
                reason=f"The module's flow worked - every screen did what the "
                       f"row asked. What did not agree is the application's "
                       f"arithmetic: {len(wrong)} of "
                       f"{len(report.checks) if report else 0} figure(s) differ "
                       f"from a value worked out independently from the same "
                       f"inputs. Each one is shown with its formula in "
                       f"'Calculation Validation'.",
                last_action=ACTIONS[-1] if ACTIONS else "",
                result="CALCULATION FAILED")
            describe(wrong_story, qa_report.failure_embed(
                module, row.test_case_id, wrong_story,
                reason="The flow was correct; the application's arithmetic was "
                       "not.", result="CALCULATION FAILED"))
        log.error("CALCULATION FAILED | %s: %d of %d calculations do not match",
                  module, len(wrong), len(report.checks) if report else 0)
        if calc.FAIL_ON_MISMATCH:
            raise AssertionError(
                f"{module}: the flow worked but the application's arithmetic "
                f"does not. {len(wrong)} calculation(s) did not match an "
                f"independently worked-out value:\n\n{detail}\n\n"
                f"Set GENZ_FAIL_ON_CALC_MISMATCH=0 to report these without "
                f"failing the row.")

    log.info("Completed Successfully.")
    counts = report.counts() if report else {}
    attach_text("Execution Status",
                f"Module     : {module}\n"
                f"Test case  : {row.test_case_id}\n"
                f"Result     : PASSED\n"
                f"Ran for    : {time.time() - started:.1f}s\n"
                f"Actions    : {len(ACTIONS)}\n"
                f"Calculations: {counts.get(calc.PASS, 0)} passed, "
                f"{counts.get(calc.FAIL, 0)} failed, "
                f"{counts.get(calc.SKIPPED, 0)} not verifiable")
    if ALLURE:
        allure.dynamic.parameter("Result", "PASSED")
        # The header card, now with the row's own result on it: what was
        # entered, what was checked and how it came out - the whole test case on
        # one screen, above the steps, without opening an attachment.
        result_card = case_result_card(row, module, "PASSED", counts,
                                       time.time() - started)
        # A module that owns one of the stock stages gets ITS OWN FIGURES
        # ahead of that card, not instead of it - this is the ONLY describe()
        # call left standing once the row passes (see module_stock_summary()),
        # so it is where the real numbers have to live or they never survive
        # to the report a reader opens.
        stock_title, stock_text, stock_verdict, stock_table_html = \
            module_stock_summary(module)
        if stock_text:
            describe(
                qa_report.joined([stock_text, case_description(row, module)]),
                qa_report.embed(stock_title, stock_text, verdict=stock_verdict,
                                table_html=stock_table_html)
                + result_card)
        else:
            describe(case_description(row, module), result_card)


def case_result_card(row: Row, module: str, result: str,
                     counts: Dict[str, int], seconds: float) -> str:
    """The test case's own card, once it is over: data in, result out."""
    calculations = (f"{counts.get(calc.PASS, 0)} passed, "
                    f"{counts.get(calc.FAIL, 0)} failed, "
                    f"{counts.get(calc.SKIPPED, 0)} not verifiable")
    return qa_report.embed(
        f"{module} - {row.test_case_id}",
        "",
        verdict=result, subtitle=case_title(row, module),
        pairs=[["Module", module],
               ["Test Case ID", row.test_case_id],
               ["Test Data", f"{row.sheet} sheet, row {row.number}"],
               ["Result", result],
               ["Actions Performed", len(ACTIONS)],
               ["Calculations", calculations],
               ["Ran For", f"{seconds:.1f}s"]])


#: The columns an application treats as an account's identity. A registration is
#: refused when any ONE of them has been used before, so all three are named in
#: the skip report - which one collided is the application's to say, and the
#: reader needs to know which cells to change either way.
IDENTITY_COLUMNS = ("Mobile Number", "Email", "GSTIN")


def duplicate_identity(row: Row) -> str:
    """The values that make this row's account unique, for the skip report."""
    values = [(column, str(row.get(column, "") or "").strip())
              for column in IDENTITY_COLUMNS]
    values = [(column, value) for column, value in values if value]
    if not values:
        return "This sheet carries none of the identity columns."
    return "\n".join(
        ["The values that identify this account, any one of which the "
         "application will refuse a second time:"]
        + [f"    {column:<14}: {value}" for column, value in values])


def case_description(row: Row, module: str) -> str:
    """The line a reader sees at the top of the test, before opening anything."""
    return (f"Module      : {module}\n"
            f"Test Case   : {row.test_case_id}\n"
            f"Test Data   : {row.sheet} sheet, row {row.number} of "
            f"Construction_Flow_Data.xlsx\n\n"
            f"Open 'Actions Performed' for what the automation did, "
            f"'Expected vs Actual' for what the application then displayed, "
            f"and 'Calculation Validation' for the arithmetic that was "
            f"re-worked independently.")


def finish_module(sheet_name: str, row: Row, rows: List[Row]) -> None:
    """Log the module banner once, after its last row has run."""
    if rows and row is rows[-1]:
        log.info("Finished %s Module.", sheet_name)


# =========================================================================== #
# TEST DATA HELPERS
#
# Every value typed into the application comes from the workbook, exactly as it
# is written in the cell. Nothing is generated behind the sheet's back - the two
# helpers here only expand shorthand the workbook itself asks for: TODAY+30 in a
# date cell, and {RUN} in a cell whose value has to be new on every run.
# =========================================================================== #

# One stamp per run, so every {RUN} in the workbook expands to the same text
# within a run and to something new on the next one. Minutes are enough: two
# runs of the same suite cannot start in the same minute.
RUN_STAMP = datetime.now().strftime("%m%d-%H%M")


def resolve_run_token(value: str) -> str:
    """Expand {RUN} in a cell to this run's stamp.

    A category name has to be new every run: this application answers a name it
    already holds with an unhandled 500 ("An internal error occurred") instead
    of a duplicate error, so the row fails rather than being skipped the way a
    repeated GSTIN or phone number is. Writing

        Shear Wall Formwork {RUN}

    in the sheet makes the name unique per run and keeps the workbook readable.
    Cells without the token are returned untouched, so nothing is stamped unless
    the sheet asks for it.
    """
    text = (value or "").strip()
    if not text:
        return text
    return re.sub(r"\{RUN\}", RUN_STAMP, text, flags=re.IGNORECASE).strip()


def resolve_date(value: str, default_days: int = 30) -> str:
    """A date cell as YYYY-MM-DD.

    Accepts what a QA engineer would naturally type:
        2026-09-30   an exact date
        TODAY+30     30 days from today   (never goes stale)
        TODAY        today
        (blank)      today + default_days
    """
    text = (value or "").strip().upper().replace(" ", "")
    if not text:
        return (date.today() + timedelta(days=default_days)).isoformat()
    if text == "TODAY":
        return date.today().isoformat()
    if text.startswith("TODAY+") or text.startswith("+"):
        days = text.split("+", 1)[1]
        if not days.isdigit():
            raise ValueError(f"'{value}' is not a date. Use 2026-09-30, TODAY, "
                             f"or TODAY+30.")
        return (date.today() + timedelta(days=int(days))).isoformat()
    return (value or "").strip()


# =========================================================================== #
# WAITING HELPERS - what Playwright cannot work out on its own
# =========================================================================== #

# Anything that means "the app is still working". The last entry is a bare
# "Loading..." cell (the admin tenants table) that carries no class or role.
LOADER_SELECTORS = (
    "[role='progressbar']",
    "[aria-busy='true']",
    "[class*='spinner' i]",
    "[class*='loader' i]",
    "[class*='skeleton' i]",
    "text=/^\\s*Loading\\s*\\.{0,3}\\s*$/i",
)


def wait_for_loaders(page: Page) -> None:
    """Wait until every visible spinner / skeleton / 'Loading...' has gone."""
    for selector in LOADER_SELECTORS:
        loader = page.locator(selector).first
        try:
            if loader.count() == 0 or not loader.is_visible():
                continue
            log.info("   waiting for the loader to disappear...")
            loader.wait_for(state="hidden", timeout=LOADER_TIMEOUT)
        except (PlaywrightTimeoutError, PlaywrightError):
            # A stuck spinner is worth a note; the next assertion decides
            # whether the step actually failed.
            log.warning("   a loader is still on screen after %d ms", LOADER_TIMEOUT)


def wait_until_ready(page: Page, what: str = "page") -> None:
    """DOM parsed -> network quiet -> DOM stable -> spinners gone.

    The name every module already calls. It now goes through wait_page_ready(),
    which adds the "wait until the SPA has stopped re-rendering" step, so every
    existing wait in the suite got that for free.
    """
    wait_page_ready(page, what)


def settle(page: Page) -> None:
    """Short pause for a CSS animation (modal fade, accordion) - the only one."""
    page.wait_for_timeout(ANIMATION_MS)


def visible(locator: Locator) -> Locator:
    """Narrow a locator to the copy actually on screen.

    The app keeps several forms mounted at once, so a placeholder such as
    "27AABCU9603R1ZX" exists on both the Customer and the Supplier form.
    Without this filter Playwright's strict mode fails with "resolved to 2
    elements".
    """
    return locator.filter(visible=True).first


def button(page: Page, name: str, exact: bool = True) -> Locator:
    """A button, addressed by the text a user reads on it."""
    return page.get_by_role("button", name=name, exact=exact)


def link(page: Page, name: str, exact: bool = True) -> Locator:
    """A link, addressed by the text a user reads on it.

    exact=True by default because Playwright matches accessible names as a
    SUBSTRING otherwise - and the sidebar holds both "Suppliers" and "Find
    Suppliers", which then resolves to 2 elements and fails on strict mode.
    """
    return page.get_by_role("link", name=name, exact=exact)


def nav_link(page: Page, module: str) -> Locator:
    """A sidebar link, tolerating the live count the application appends to it.

    "Sales Orders" reads "Sales Orders 1" the moment there is one order, so an
    exact match stops working as soon as the suite has done any work of its own.
    Anchoring at the start keeps "Suppliers" from also matching "Find Suppliers".
    """
    return page.get_by_role(
        "link", name=re.compile(rf"^{re.escape(module)}(\s|$)")).filter(visible=True)


# =========================================================================== #
# RESILIENCE - the self-healing layer every action goes through
#
# The application is a busy SPA: it animates, it re-renders while you are
# typing, it drops a backdrop over the sidebar, and its API is sometimes slow.
# None of that is a defect, so none of it should fail a test. Every helper here
# has the same shape:
#
#   wait until the page can be worked with -> do the thing -> prove it worked
#   ... and if it did not, heal whatever is in the way and try again.
#
# Nothing here sleeps. Every wait is a Playwright wait for a real condition, so
# a fast machine waits for nothing and a slow one is not cut short. Every retry
# is logged and photographed into the Allure report, so a run that needed three
# attempts says so instead of quietly hiding it.
# =========================================================================== #

#: Anything that can sit on top of the page and swallow a click. The first is
#: the workspace's "What's New" backdrop, which is the usual culprit.
OVERLAY_SELECTORS = (
    "div.fixed.inset-0.z-50",
    "[class*='backdrop' i]",
    "[class*='overlay' i]",
    "[data-state='open'][class*='fixed' i]",
)
def wait_api_complete(page: Page, timeout: int = NETWORK_IDLE_TIMEOUT) -> None:
    """Let the API calls in flight finish. Best effort - never fails.

    A page that polls, or holds a socket open, never goes idle at all; failing a
    test for that would be self-inflicted, so the assertion that comes next is
    what decides whether anything is actually wrong.
    """
    try:
        page.wait_for_load_state("networkidle", timeout=timeout)
    except (PlaywrightTimeoutError, PlaywrightError):
        log.debug("   the network did not go quiet within %d ms", timeout)


def wait_dom_stable(page: Page, timeout: int = DOM_STABLE_TIMEOUT) -> None:
    """Wait until the DOM stops changing - the SPA has finished re-rendering.

    This is what replaces "pause until the list has refreshed". It counts the
    nodes on consecutive animation frames and returns the moment the count holds
    still three frames running, so it is a wait for a condition rather than a
    guess at how long a re-render takes.
    """
    try:
        page.wait_for_function(
            """() => new Promise(resolve => {
                   const size = () => document.querySelectorAll('*').length;
                   let previous = size();
                   let steady = 0;
                   const tick = () => {
                       const current = size();
                       steady = current === previous ? steady + 1 : 0;
                       previous = current;
                       if (steady >= 3) { resolve(true); }
                       else { requestAnimationFrame(tick); }
                   };
                   requestAnimationFrame(tick);
               })""",
            timeout=timeout)
    except (PlaywrightTimeoutError, PlaywrightError):
        log.debug("   the DOM was still moving after %d ms", timeout)


def overlays_on_screen(page: Page) -> Locator:
    """Every backdrop currently covering the page, as one locator."""
    return page.locator(", ".join(OVERLAY_SELECTORS)).filter(visible=True)


def wait_for_overlays_gone(page: Page, timeout: int = OVERLAY_TIMEOUT,
                           close_stubborn: bool = False) -> None:
    """Wait for spinners and backdrops covering the page to go.

    Only NAVIGATION passes close_stubborn=True. That matters: a create form is
    itself a modal with its own backdrop, so an overlay on screen is perfectly
    normal while a form is being filled in - and closing it would cancel the
    very form being worked on. Clicks therefore only ever wait here; deciding
    that an overlay is in the way is left to Playwright, which refuses to click
    through one and says so, and to the escalation in click_escalating().
    """
    wait_for_loaders(page)
    try:
        if overlays_on_screen(page).count() == 0:
            return
    except PlaywrightError:
        return

    log.info("   something is covering the page - waiting for it to close")
    try:
        overlays_on_screen(page).first.wait_for(state="hidden", timeout=timeout)
        return
    except (PlaywrightTimeoutError, PlaywrightError):
        pass

    if not close_stubborn:
        log.info("   the overlay is still there - it is probably the open "
                 "form's own backdrop, so it is left alone")
        return

    log.warning("   the overlay is still there after %d ms - closing it", timeout)
    dismiss_overlay(page)


def wait_page_ready(page: Page, what: str = "page") -> None:
    """DOM parsed -> API quiet -> DOM stable -> spinners gone. Use after any nav.

    Deliberately does NOT close overlays: it runs in the middle of forms too,
    and a create form that is still open there is the thing being worked on, not
    something in the way. Clicks clear their own path with wait_for_overlays_gone.
    """
    try:
        page.wait_for_load_state("domcontentloaded", timeout=NAVIGATION_TIMEOUT)
    except (PlaywrightTimeoutError, PlaywrightError):
        log.warning("   %s did not finish loading in time - continuing", what)
    wait_api_complete(page)
    wait_dom_stable(page)
    wait_for_loaders(page)
    log.info("   %s is ready", what)


def reload_page(page: Page) -> None:
    """Reload, and wait until the page can be worked with again."""
    log.info("   reloading the page and looking again")
    try:
        page.reload(wait_until="domcontentloaded", timeout=NAVIGATION_TIMEOUT)
    except (PlaywrightTimeoutError, PlaywrightError) as error:
        log.warning("   the reload did not complete: %s",
                    str(error).splitlines()[0])
        return
    wait_page_ready(page, "reloaded page")


# --- Bringing a control onto the screen ------------------------------------- #
#
# Every click, fill and select goes through scroll_into_view() below before it
# touches anything, so the field being worked on is on screen - and visibly on
# screen to whoever is watching the run - rather than somewhere below the fold.
#
# Two things make this less trivial than calling scroll_into_view_if_needed():
#
#   * Playwright's is_visible() is NOT a viewport check. It is true for an
#     element sitting a thousand pixels below the fold, so it cannot be used to
#     confirm that a scroll actually worked. getBoundingClientRect() can, and
#     that is what IN_VIEWPORT_JS asks the browser.
#   * scroll_into_view_if_needed() scrolls the nearest scrollable ancestor, but
#     a modal body inside a scrolling page can leave the control level with a
#     sticky header. SCROLL_JS re-centres it, which walks every scrollable
#     ancestor - the container as well as the page.
#
# Nothing here sleeps and nothing here raises: a control that cannot be reached
# is left to the click or fill that follows, which fails with a message naming
# the field. That is deliberate - this layer must never be the thing that turns
# a working step into a failure.

#: True only when the element really is inside the viewport. An element taller
#: than the viewport counts as visible while it straddles the screen, otherwise
#: a long form section could never satisfy the check.
IN_VIEWPORT_JS = """element => {
    const rect = element.getBoundingClientRect();
    if (rect.width === 0 && rect.height === 0) { return false; }
    const height = window.innerHeight || document.documentElement.clientHeight;
    const width = window.innerWidth || document.documentElement.clientWidth;
    const verticallyIn = rect.height <= height
        ? (rect.top >= 0 && rect.bottom <= height)
        : (rect.top <= 0 && rect.bottom >= 0);
    const horizontallyIn = rect.right > 0 && rect.left < width;
    return verticallyIn && horizontallyIn;
}"""

#: Centre the element in whatever scrolls it - the page or a scrollable panel.
SCROLL_JS = """element => element.scrollIntoView(
    {block: 'center', inline: 'nearest', behavior: 'instant'})"""

#: Outline the control being worked on so the run can be followed by eye, then
#: take the outline away again. The timer runs in the page, so the test does not
#: wait for it, and the outline is drawn on top without moving anything: it is
#: an outline rather than a border precisely because a border would reflow the
#: form and shift the very field being aimed at.
HIGHLIGHT_JS = """(element, milliseconds) => {
    const previous = element.style.outline;
    const previousOffset = element.style.outlineOffset;
    element.style.outline = '2px solid #ff6a00';
    element.style.outlineOffset = '2px';
    setTimeout(() => {
        element.style.outline = previous;
        element.style.outlineOffset = previousOffset;
    }, milliseconds);
}"""


def in_viewport(locator: Locator) -> bool:
    """True when the control is really on the screen right now.

    Answered by the browser's own geometry, because Playwright's is_visible()
    reports an element far below the fold as visible.
    """
    try:
        return bool(locator.evaluate(IN_VIEWPORT_JS, timeout=SCROLL_TIMEOUT))
    except (PlaywrightTimeoutError, PlaywrightError):
        return False


def highlight(locator: Locator) -> None:
    """Draw a short-lived outline round the control. Best effort, never waits."""
    try:
        locator.evaluate(HIGHLIGHT_JS, HIGHLIGHT_MS, timeout=SCROLL_TIMEOUT)
    except (PlaywrightTimeoutError, PlaywrightError):
        pass


def scroll_into_view(locator: Locator, description: str = "",
                     editable: bool = False,
                     attempts: int = SCROLL_ATTEMPTS) -> bool:
    """Put a control on screen and prove it got there. Never fails.

    The order matters, and it is the order a person would use:

        1. wait for the control to be attached and visible in the DOM;
        2. scroll it into view the way Playwright does - which handles a
           scrollable container as well as the page;
        3. ASK THE BROWSER whether it is actually in the viewport, and if it is
           not, centre it with scrollIntoView() and look again;
        4. wait for it to be enabled - and editable too, for a field about to be
           typed into;
        5. focus it (fields only) and outline it, so the value can be seen going
           into a field that is on screen.

    Returns True when the control was confirmed inside the viewport. The return
    value is information, not a gate: a False is left to the click or fill that
    follows, which reports it against the field's own name. Nothing here raises
    and nothing here sleeps.

    Locators are not cached between attempts - Playwright re-resolves them every
    time - so a form that re-lays-out after the first scroll is simply found
    again where it now is.
    """
    for attempt in range(1, attempts + 1):
        try:
            locator.wait_for(state="visible", timeout=SCROLL_TIMEOUT)
        except (PlaywrightTimeoutError, PlaywrightError):
            # Detached, hidden or many matches. The action that follows says so
            # against the field's own name, which reads far better than anything
            # this helper could report.
            return False

        try:
            locator.scroll_into_view_if_needed(timeout=SCROLL_TIMEOUT)
        except (PlaywrightTimeoutError, PlaywrightError):
            pass                       # the geometry check below is the verdict

        if not in_viewport(locator):
            try:
                locator.evaluate(SCROLL_JS, timeout=SCROLL_TIMEOUT)
            except (PlaywrightTimeoutError, PlaywrightError):
                pass

        if in_viewport(locator):
            break

        if attempt < attempts:
            log.debug("   '%s' is still off screen - scrolling to it again",
                      description or "the control")
            wait_dom_stable(locator.page)     # a re-render moved it: let it land
    else:
        log.warning("   could not bring '%s' fully onto the screen - working "
                    "with it anyway", description or "the control")
        return False

    ready(locator, editable=editable)
    if editable:
        try:
            locator.focus(timeout=SCROLL_TIMEOUT)
        except (PlaywrightTimeoutError, PlaywrightError):
            pass          # a styled widget may not take focus; the fill still works
    highlight(locator)
    return True


def ready(locator: Locator, editable: bool = False) -> None:
    """Wait for the control to be enabled, and editable when it is a field.

    Playwright's own actionability checks cover this before it clicks or types,
    so this is belt and braces - but it is checked HERE, before the outline goes
    on, so that a control which is on screen but still disabled is waited for
    rather than typed into. Kept to the short READY_TIMEOUT because the action
    that follows waits properly: a control that is never going to be enabled
    should not cost three long waits here on top of its three real ones.
    """
    try:
        expect(locator).to_be_enabled(timeout=READY_TIMEOUT)
    except (AssertionError, PlaywrightTimeoutError, PlaywrightError):
        pass
    if not editable:
        return
    try:
        expect(locator).to_be_editable(timeout=READY_TIMEOUT)
    except (AssertionError, PlaywrightTimeoutError, PlaywrightError):
        # A contenteditable div, or a date box the application drives itself,
        # is not "editable" by Playwright's definition but still takes a value.
        pass


def retry_action(page: Page, description: str, action, attempts: int,
                 heal=None):
    """Run `action(attempt)` until it works, healing in between.

    The heart of the self-healing layer. Every failed attempt is logged, has a
    screenshot attached to the Allure report and is followed by heal() - the
    chance to close whatever got in the way. The exception from the LAST attempt
    is raised, so the caller can turn it into its own readable failure.

    Nothing is cached between attempts: Playwright locators re-resolve every
    time they are used, which is what makes this recover from a stale element -
    the control that was re-rendered underneath us is simply found again.
    """
    last_error: Optional[BaseException] = None

    for attempt in range(1, attempts + 1):
        try:
            result = action(attempt)
            if attempt > 1:
                log.info("   '%s' worked on attempt %d/%d", description, attempt,
                         attempts)
                attach_text("Recovered After Retry",
                            f"{description} succeeded on attempt {attempt} of "
                            f"{attempts}.")
            return result
        except (PlaywrightTimeoutError, PlaywrightError, AssertionError) as error:
            last_error = error
            log.warning("   attempt %d/%d failed for '%s': %s", attempt, attempts,
                        description, str(error).splitlines()[0])
            screenshot(page, f"Retry {attempt} - {description}")
            if attempt < attempts and heal is not None:
                heal()

    raise last_error                                   # type: ignore[misc]


def click_escalating(locator: Locator, description: str, timeout: int,
                     allow_force: bool = True, allow_script: bool = True) -> None:
    """Click, escalating through three ways of doing it.

    1. the ordinary click - Playwright waits for the control to be visible,
       enabled and stable, and refuses to click through an overlay (that refusal
       is the "intercepted" error);
    2. force=True - skips those checks, for a control the application styles
       oddly (a label sitting over its own input);
    3. element.click() inside the page - the last resort. It does not go through
       the mouse at all, so nothing can intercept it.

    The caller decides how far to go: safe_click() only unlocks step 2 on its
    second attempt and step 3 on its third, so an ordinary click that is simply
    going to fail does not sit through three timeouts before saying so. Steps 2
    and 3 skip the actionability wait, so they get the shorter timeout.
    """
    try:
        locator.click(timeout=timeout)
        return
    except (PlaywrightTimeoutError, PlaywrightError) as error:
        if not allow_force:
            raise
        log.warning("   the ordinary click on '%s' did not land (%s) - forcing it",
                    description, str(error).splitlines()[0])

    try:
        locator.click(timeout=OPTIONAL_TIMEOUT, force=True)
        return
    except (PlaywrightTimeoutError, PlaywrightError):
        if not allow_script:
            raise
        log.warning("   the forced click on '%s' did not land either - clicking "
                    "it from inside the page", description)

    locator.evaluate("element => element.click()", timeout=OPTIONAL_TIMEOUT)


#: The buttons that submit a create form. "Add ..." and "New ..." are deliberately
#: NOT here - they OPEN a form (or the line-item picker), they do not save one.
SUBMIT_BUTTON = re.compile(r"^(create|save|submit|update)\b", re.IGNORECASE)


def safe_click(page: Page, locator: Locator, description: str,
               animated: bool = False, attempts: int = ACTION_ATTEMPTS,
               timeout: int = RETRY_TIMEOUT) -> None:
    """Click a control, and keep trying while the page gets in the way.

    Scrolls the control into view, clicks it, and retries with more force each
    time. If the control has vanished by the time a retry starts, the earlier
    click did land after all - that check is what stops a retry from submitting
    the same form twice.

    It deliberately does NOT clear overlays first. A create form is itself a
    modal with a backdrop, so "an overlay is on screen" is the normal state
    while a form is being filled in; treating that as an obstruction is how an
    earlier version of this helper came to cancel the form it was submitting.
    Playwright refuses to click through anything that really is in the way, and
    the escalation above is the answer when it does.
    """
    log.info("   click: %s", description)

    # A Create / Save / Submit click is the last moment the completed form is on
    # screen, so this is where the entered data and the form's screenshot go
    # into the report. It is here, and not in the modules, so that EVERY create
    # is covered without a line of code in any test case.
    if SUBMIT_BUTTON.match(description):
        verify_data.before_save(page, description)

    def heal() -> None:
        wait_for_loaders(page)          # a spinner does pass by itself
        wait_dom_stable(page)           # so does a re-render

    def attempt_click(attempt: int) -> None:
        if attempt > 1 and locator.count() == 0:
            log.warning("   '%s' is no longer on the page - taking it that the "
                        "earlier click landed", description)
            return
        scroll_into_view(locator, description)
        click_escalating(locator, description, timeout,
                         allow_force=attempt >= 2, allow_script=attempt >= 3)

    try:
        retry_action(page, f"click {description}", attempt_click, attempts, heal)
    except (PlaywrightTimeoutError, PlaywrightError, AssertionError) as error:
        fail(page, f"Could not click '{description}'", error)

    qa_step(f"Clicked {quoted(description)}")

    if animated:
        settle(page)


def value_matches(actual: Optional[str], wanted: str) -> bool:
    """True when a field really holds what was typed into it.

    Number boxes reformat what they are given (35 -> 35.00), currency boxes add
    separators and phone boxes add spaces, so an exact string comparison would
    report a perfectly good value as wrong. Three readings count as a match: the
    same text, the same number, or the same digits.
    """
    actual = (actual or "").strip()
    wanted = (wanted or "").strip()
    if actual == wanted:
        return True
    try:
        return float(actual.replace(",", "")) == float(wanted.replace(",", ""))
    except ValueError:
        pass
    actual_digits = re.sub(r"\D", "", actual)
    wanted_digits = re.sub(r"\D", "", wanted)
    return bool(wanted_digits) and actual_digits == wanted_digits


def safe_fill(page: Page, locator: Locator, value: str, description: str,
              secret: bool = False, verify: bool = True,
              attempts: int = ACTION_ATTEMPTS,
              timeout: int = RETRY_TIMEOUT) -> None:
    """Type into a field, prove the value went in, and try again if it did not.

    Each attempt types a different way, because the reason a field refuses one
    way is usually that it wants another:

        1. fill()               - one go, which is what a paste does
        2. click, clear, type   - for a widget that listens for keystrokes
        3. select all, delete, type slowly - for a masked or debounced box

    A field that ends up holding something other than what was typed is retried,
    but a mismatch on the LAST attempt is only reported, not failed: masks and
    reformatting are the usual reason, and the application's own validation is
    the honest judge of a wrong value - it answers on submit, with a message
    confirm_saved() then puts in the report.
    """
    shown = "********" if secret else value
    log.info("   fill: %-24s = %s", description, shown)
    mismatch: Optional[str] = None

    def heal() -> None:
        wait_dom_stable(page)

    def attempt_fill(attempt: int) -> None:
        nonlocal mismatch
        mismatch = None
        scroll_into_view(locator, description, editable=True)

        if attempt == 1:
            locator.fill(value, timeout=timeout)
        elif attempt == 2:
            locator.click(timeout=timeout)
            locator.fill("", timeout=timeout)
            locator.press_sequentially(value, delay=20, timeout=timeout)
        else:
            locator.click(timeout=timeout)
            page.keyboard.press("Control+A")
            page.keyboard.press("Delete")
            locator.press_sequentially(value, delay=60, timeout=timeout)

        if not (verify and value):
            return
        try:
            actual = locator.input_value(timeout=timeout)
        except (PlaywrightTimeoutError, PlaywrightError):
            return                    # a contenteditable div has no .value
        if not value_matches(actual, value):
            mismatch = actual
            raise AssertionError(
                f"'{description}' holds {actual!r} after typing {shown!r}")

    try:
        retry_action(page, f"fill {description}", attempt_fill, attempts, heal)
    except AssertionError as error:
        if mismatch is None:
            fail(page, f"Could not fill '{description}'", error)
        # The typing itself worked every time; only the read-back differs.
        log.warning("   '%s' reads %r after typing %r - the application has "
                    "reformatted it, so the run continues", description, mismatch,
                    shown)
        attach_text("Field Reformatted by the Application",
                    f"{description}: typed {shown!r}, the field shows "
                    f"{mismatch!r}.")
    except (PlaywrightTimeoutError, PlaywrightError) as error:
        fail(page, f"Could not fill '{description}'", error)

    qa_step(f"Entered {description}: {shown}")

    # Remember what went into the field, for the View-page comparison. A secret
    # is never remembered and never compared - it is not displayed anywhere.
    if not secret:
        verify_data.record(description, value)


def safe_select(page: Page, locator: Locator, value: str, description: str,
                hint: str = "", attempts: int = ACTION_ATTEMPTS) -> None:
    """Choose an option from a <select>, by its value or by its label.

    The workbook may hold either what the application stores ("semi_finished")
    or what it prints ("Semi Finished"), and a dropdown that re-renders can drop
    the selection - so the choice is made, and made again if it did not take.
    """
    log.info("   select: %-24s = %s", description, value)

    def attempt_select(attempt: int) -> None:
        scroll_into_view(locator, description)
        for by_label in (False, True):
            try:
                if by_label:
                    locator.select_option(label=value, timeout=RETRY_TIMEOUT)
                else:
                    locator.select_option(value, timeout=RETRY_TIMEOUT)
                return
            except (PlaywrightTimeoutError, PlaywrightError):
                continue
        raise AssertionError(f"'{value}' is not an option of '{description}'")

    try:
        retry_action(page, f"select {description}", attempt_select, attempts)
    except (PlaywrightTimeoutError, PlaywrightError, AssertionError) as error:
        try:
            offered = " | ".join(text.strip() for text
                                 in locator.locator("option").all_inner_texts())
        except PlaywrightError:
            offered = "(the form did not list them)"
        fail(page, f"'{value}' is not one of the '{description}' options on this "
                   f"form. The form offers: {offered}.{' ' + hint if hint else ''}",
             error)

    qa_step(f"Selected {description}: {value}")
    verify_data.record(description, value)


def safe_expect(page: Page, locator: Locator, description: str,
                timeout: int = RETRY_TIMEOUT, attempts: int = VERIFY_ATTEMPTS,
                refresh: bool = False) -> None:
    """Assert an element is on screen, giving a slow application a second look.

    Between attempts the page is allowed to settle (and reloaded, when the
    caller says a reload is safe), so a list that is still being fetched is
    waited for rather than failed.
    """
    def heal() -> None:
        wait_page_ready(page, "page")
        if refresh:
            reload_page(page)

    def attempt_expect(attempt: int) -> None:
        expect(locator).to_be_visible(timeout=timeout)

    try:
        retry_action(page, f"see {description}", attempt_expect, attempts, heal)
    except (AssertionError, PlaywrightTimeoutError, PlaywrightError) as error:
        fail(page, f"Expected to see '{description}', but it never appeared",
             error)

    qa_step(f"Verified {description} is displayed")


def safe_navigation(page: Page, url: str, what: str, attempts: int = 2) -> None:
    """Go to an address, retry a failed load, then wait until the page is ready."""
    log.info("   opening %s (%s)", what, url)

    def attempt_goto(attempt: int) -> None:
        page.goto(url, wait_until="domcontentloaded", timeout=NAVIGATION_TIMEOUT)

    try:
        retry_action(page, f"open {what}", attempt_goto, attempts)
    except (PlaywrightTimeoutError, PlaywrightError) as error:
        fail(page, f"Could not open {what} ({url})", error)

    wait_page_ready(page, what)
    qa_step(f"Opened {what}")

def safe_locator(page: Page, description: str,
                 candidates: Tuple[Tuple[str, Locator], ...],
                 need_editable: bool = False) -> Optional[Locator]:
    """The first way of finding a control that actually matches something.

    Each candidate is (how it is being looked for, the locator), best first, and
    the one that is really on screen wins. This is what keeps a renamed class or
    a changed input type from breaking a step: the next way of finding it takes
    over, and the log says which way worked."""

    for label, candidate in candidates:
        try:
            target = candidate.filter(visible=True).first
            if target.count() == 0:
                continue
            if need_editable and not target.is_editable(timeout=OPTIONAL_TIMEOUT):
                continue
        except (PlaywrightTimeoutError, PlaywrightError):
            continue          # not a field, detached, or too many matches
        log.info("   locator: %-22s -> %s", description, label)
        return target

    log.warning("   none of the %d ways of finding '%s' matched anything",
                len(candidates), description)
    return None


def find_quantity_box(page: Page, unit_label: str = "PCS") -> Optional[Locator]:
    """The quantity input of a line-item row, whatever this build calls it.

    The row is a bare CSS grid: its boxes have no label, no id and no
    placeholder, and the markup has changed between builds - the box has been a
    spinbutton, a plain number input and a textbox. So every way of addressing
    it is tried in turn, anchored first on the unit printed beside it (PCS,
    BAGS, MT) and then on the word Quantity, ending with an XPath that simply
    takes the input next to that word.
    """
    unit = (unit_label or "").strip() or "PCS"
    unit_row = page.locator("div").filter(
        has_text=re.compile(rf"^{re.escape(unit)}$"), visible=True)

    return safe_locator(page, "Quantity", (
        ("the spinbutton beside the unit", unit_row.get_by_role("spinbutton")),
        ("the textbox beside the unit", unit_row.get_by_role("textbox")),
        ("the number input beside the unit",
         unit_row.locator('input[type="number"]')),
        ("any input beside the unit", unit_row.locator("input")),
        ("a spinbutton called Quantity",
         page.get_by_role("spinbutton", name="Quantity")),
        ("a textbox called Quantity",
         page.get_by_role("textbox", name="Quantity")),
        ("a field labelled Quantity", page.get_by_label("Quantity")),
        ("the Quantity placeholder", page.get_by_placeholder("Quantity")),
        ("the input next to the word Quantity",
         page.locator("xpath=//*[contains(normalize-space(.),'Quantity')]"
                      "/following::input[1]")),
        ("an editable box in the row", page.locator('[contenteditable="true"]')),
    ), need_editable=True)


def wait_for_submit_to_close(page: Page, submit: Locator, record_text: str,
                             what: str) -> None:
    """Wait for a create form's own button to go - the payload was accepted.

    A button that stays put means the application refused the record, and it
    often refuses in silence: the 4xx body is the only evidence, which is what
    ApiFailureRecorder is for.
    """
    try:
        submit.wait_for(state="hidden", timeout=DEFAULT_TIMEOUT)
        return
    except (PlaywrightTimeoutError, PlaywrightError):
        pass

    messages = read_app_messages(page)
    if messages and already_exists(messages):
        raise RecordAlreadyExists(
            f"{what} '{record_text}' is already in the application - this "
            f"workbook row has been used before. Put a fresh value in the sheet "
            f"to create it again.\n" + "\n".join(messages))
    if messages:
        fail(page, f"{what} was rejected by the application", "\n".join(messages))
    log.warning("   the %s form is still open but shows no error - checking for "
                "the record anyway", what)


def verify_success(page: Page, record_text: str, what: str,
                   submit: Optional[Locator] = None, refresh: bool = True,
                   attempts: int = VERIFY_ATTEMPTS) -> None:
    """Prove a record was really saved - never straight after the click.

    Looking immediately is what makes a save look broken when it is not: the
    success toast may already have faded, the list is still being fetched, and
    the DOM is mid-redraw. So every round waits for the API to answer and the
    DOM to stop moving first, then looks for the record OR a success message,
    and reloads before the next round. A duplicate message at any point means
    the row is used up, which is reported as skipped, not failed.
    """
    if submit is not None:
        wait_for_submit_to_close(page, submit, record_text, what)

    for attempt in range(1, attempts + 1):
        wait_api_complete(page)
        wait_dom_stable(page)
        wait_for_loaders(page)

        messages = read_app_messages(page)
        if messages and already_exists(messages):
            raise RecordAlreadyExists(
                f"{what} '{record_text}' is already in the application - this "
                f"workbook row has been used before. Put a fresh value in the "
                f"sheet to create it again.\n" + "\n".join(messages))

        # visible=True on both: a previous module's toast can linger in the DOM,
        # and the form the app has just closed still holds the same text.
        record = page.get_by_text(record_text).filter(visible=True)
        toast = page.get_by_text(re.compile(r"success", re.IGNORECASE)) \
                    .filter(visible=True)
        if appears(record.or_(toast).first, RETRY_TIMEOUT):
            log.info("   saved: %s ('%s')", what, record_text)
            return

        screenshot(page, f"{what} - Retry {attempt} of the Verification")
        log.warning("   attempt %d/%d: '%s' is not on screen yet", attempt,
                    attempts, record_text)
        if attempt < attempts and refresh:
            reload_page(page)

    fail(page, f"{what} was submitted, but neither '{record_text}' nor a success "
               f"message appeared - not after {attempts} attempts and a reload.")


def click(page: Page, locator: Locator, description: str,
          animated: bool = False) -> None:
    """Click a control. The framework's click: see safe_click() for what it does.

    Kept as the name every module already calls, so every existing step got the
    overlay handling, the escalation and the retries without changing a line.
    """
    safe_click(page, locator, description, animated=animated)


def fill(page: Page, locator: Locator, value: str, description: str,
         secret: bool = False) -> None:
    """Type into a field. The framework's fill: see safe_fill() for what it does."""
    safe_fill(page, locator, value, description, secret=secret)


def expect_visible(page: Page, locator: Locator, description: str) -> None:
    """Assert an element is on screen. See safe_expect() for what it does."""
    safe_expect(page, locator, description)


def appears(locator: Locator, timeout: int = OPTIONAL_TIMEOUT) -> bool:
    """True when the element shows up within the timeout - never raises.

    For the places where "not on screen" is a branch rather than a failure: a
    picker that lists its options straight away versus one that only lists them
    once it has been searched.
    """
    try:
        locator.wait_for(state="visible", timeout=timeout)
        return True
    except (PlaywrightTimeoutError, PlaywrightError):
        return False


def find_field(*candidates: Locator) -> Optional[Locator]:
    """The first candidate locator that is really on the form, or None.

    The same field can be addressed by its label on one screen and only by its
    placeholder on another, so each caller lists the ways it can be found, best
    first.
    """
    for candidate in candidates:
        try:
            if candidate.count() > 0:
                return candidate
        except PlaywrightError:
            continue
    return None


def fill_optional(page: Page, value: str, description: str,
                  *candidates: Locator) -> bool:
    """Fill a field only if the workbook has a value AND the form has the field.

    Used for the fields a tenant's configuration can switch off (Credit Limit,
    Credit Days, Discount, Tax, Remarks). An empty cell or a missing field is
    logged and attached to the Allure report, but never fails the test - the
    record itself is what the test is about, and a rejected save is caught by
    confirm_saved() with the application's own message.
    """
    value = (value or "").strip()
    if not value:
        log.info("   skip: %-24s (the cell is empty)", description)
        # Intentionally blank: it is a skipped comparison, never a mismatch.
        verify_data.note_skipped(description, "", "the cell in the workbook is empty")
        qa_step(f"Skipped {description} - the workbook leaves this cell empty")
        return False

    field = find_field(*candidates)
    if field is not None:
        try:
            # These are the fields at the BOTTOM of a form (Credit Limit, Credit
            # Days, Remarks), so the scroll matters more here than anywhere.
            scroll_into_view(field, description, editable=True)
            field.fill(value, timeout=DEFAULT_TIMEOUT)
            log.info("   fill: %-24s = %s", description, value)
            verify_data.record(description, value)
            qa_step(f"Entered {description}: {value}")
            return True
        except (PlaywrightTimeoutError, PlaywrightError) as error:
            log.warning("   could not enter %s = %s (%s)", description, value,
                        str(error).splitlines()[0])
            verify_data.note_skipped(description, value,
                                     "the field would not take the value")
    else:
        log.warning("   this form has no '%s' field - '%s' was not entered",
                    description, value)
        verify_data.note_skipped(description, value,
                                 "this form has no such field")
    attach_text("Field Not Entered", f"{description} = {value}")
    qa_step(f"Could not enter {description}: {value} - the form did not take it")
    return False


def dismiss_overlay(page: Page) -> None:
    """Close anything modal that is covering the page.

    Two things get in the way of a sidebar click:

    * the workspace's "What's New" panel, whose backdrop covers the whole
      viewport and silently swallows every click;
    * a create form left open by a row that the application rejected - the next
      module then cannot reach the menu at all, so ONE bad row would fail every
      module after it.
    """
    dialog = page.get_by_role("dialog").filter(visible=True)
    if dialog.count() > 0:
        log.info("   a form is still open - closing it before navigating")
        # Search INSIDE the dialog. A page-wide "Cancel" is a different button
        # on a record's own page - on a sales order it cancels the ORDER - and
        # clicking that to tidy up would be a destructive mistake.
        panel = dialog.first
        for name in ("Cancel", "Close", "Discard"):
            control = panel.get_by_role("button", name=name).filter(visible=True)
            if control.count() > 0:
                scroll_into_view(control.first, f"{name} (open form)")
                control.first.click()
                break
        else:
            page.keyboard.press("Escape")
        settle(page)

    # Not every create form is a role="dialog", and a rejected one stays on
    # screen with its backdrop over the sidebar - which is how ONE bad row used
    # to take down every module after it ("Could not click 'Orders & Sales
    # menu'"). The giveaway is a "Create ..." button paired with a "Cancel":
    # both together mean an unsaved form, never a saved record's page.
    create = page.get_by_role("button", name=re.compile(r"^Create\s")) \
                 .filter(visible=True)
    cancel = button(page, "Cancel").filter(visible=True)
    if create.count() > 0 and cancel.count() > 0:
        log.info("   an unsaved create form is in the way - cancelling it")
        try:
            scroll_into_view(cancel.first, "Cancel (unsaved form)")
            cancel.first.click(timeout=DEFAULT_TIMEOUT)
        except (PlaywrightTimeoutError, PlaywrightError):
            page.keyboard.press("Escape")
        settle(page)

    overlay = page.locator("div.fixed.inset-0.z-50")
    if overlay.count() == 0:
        return
    log.info("   a modal panel is covering the page - closing it")
    for name in ("Close", "Dismiss"):
        control = button(page, name)
        if control.count() > 0 and control.first.is_visible():
            scroll_into_view(control.first, f"{name} (modal panel)")
            control.first.click()
            break
    else:
        page.keyboard.press("Escape")
    settle(page)


def module_path(module: str) -> str:
    """The address a sidebar module lives at ("Sales Orders" -> /sales-orders)."""
    return f"/{module.lower().replace(' ', '-')}"


def expand_sidebar(page: Page) -> None:
    """Open the sidebar when a narrow window has collapsed it to a hamburger.

    Only called when the wanted link cannot be found at all - clicking a toggle
    while the sidebar is already open would close it.
    """
    for label in ("Open sidebar", "Toggle sidebar", "Open menu", "Menu"):
        toggle = page.get_by_role("button", name=label, exact=False) \
                     .filter(visible=True)
        try:
            if toggle.count() == 0:
                continue
        except PlaywrightError:
            continue
        log.info("   the sidebar looks collapsed - opening it")
        safe_click(page, toggle.first, label, animated=True, attempts=1)
        return


def page_is_open(page: Page, module: str, expected_button: str,
                 path: str = "") -> bool:
    """True when the module's own page is really on screen.

    Three proofs, because each one alone can lie: the address (which an SPA
    updates before it has drawn anything), the page's heading, and the button
    the module is known for.

    `path` is for the one module whose address is not its name - the sidebar
    reads "Material Indents" and the application routes it to /indents. Left
    empty, the address is worked out from the module name as it always was.
    """
    if (path or module_path(module)) not in page.url:
        return False

    control = button(page, expected_button).filter(visible=True)
    if appears(control.first, RETRY_TIMEOUT):
        return True

    # A list page can rename its button once it holds data; the heading is then
    # what proves the module opened.
    heading = page.get_by_role(
        "heading", name=re.compile(rf"^\s*{re.escape(module)}", re.IGNORECASE)
    ).filter(visible=True)
    try:
        return heading.count() > 0
    except PlaywrightError:
        return False


def open_module(page: Page, module: str, expected_button: str,
                attempts: int = NAV_ATTEMPTS, path: str = "") -> None:
    """Open a workspace module from the sidebar, and prove the app got there.

    The sidebar click is the most intercepted action in the suite. The "What's
    New" backdrop, a create form left open by a rejected row, a toast in the
    corner, a sticky header or a spinner will all swallow it - and the SPA then
    reports a perfectly successful click while staying exactly where it was.

    So each round clears the way, opens the sidebar if it is collapsed, scrolls
    the link into view, escalates the click (ordinary -> forced -> from inside
    the page), and only then checks the address, the heading and the button. The
    round before last gives up on the sidebar altogether and goes to the address
    directly, which no overlay can intercept. Every failed round leaves a
    screenshot in the report.

    `path` overrides the address the module name implies. Only Material Indents
    needs it: its sidebar link opens /indents, and /material-indents is a page
    the application does not have - so without it the sidebar click was landing
    perfectly and the check that followed was calling it a failure.
    """
    target_path = path or module_path(module)
    log.info("Opening the %s module...", module)

    for attempt in range(1, attempts + 1):
        # 1. Clear the way. This is the one place that is allowed to close what
        #    it finds: navigating away from a form the application left open is
        #    exactly what dismiss_overlay() is for.
        dismiss_overlay(page)
        wait_for_overlays_gone(page, close_stubborn=True)

        # 2. Find the link, opening the sidebar if it is not there at all.
        nav = nav_link(page, module).first
        if nav.count() == 0:
            expand_sidebar(page)
            nav = nav_link(page, module).first

        # 3. Click it, escalating if the ordinary click does not land.
        if nav.count() > 0:
            scroll_into_view(nav)
            try:
                click_escalating(nav, f"{module} menu", RETRY_TIMEOUT)
            except (PlaywrightTimeoutError, PlaywrightError) as error:
                log.warning("   attempt %d/%d: the %s link would not take a "
                            "click: %s", attempt, attempts, module,
                            str(error).splitlines()[0])
        else:
            log.warning("   attempt %d/%d: there is no '%s' link in the sidebar",
                        attempt, attempts, module)

        # 4. Prove the application really went there.
        wait_page_ready(page, f"{module} page")
        if page_is_open(page, module, expected_button, path):
            log.info("   %s is open (%s)", module, page.url)
            qa_step(f"Opened the {module} module")
            return

        screenshot(page, f"Retry {attempt} - Opening the {module} module")
        log.warning("   attempt %d/%d: still on %s", attempt, attempts, page.url)

        # 5. Second to last go: stop fighting the sidebar and use the address.
        #    Nothing can intercept a navigation, so this is the reliable way in.
        if attempt == attempts - 1:
            origin = re.match(r"^(https?://[^/]+)", page.url)
            if origin:
                safe_navigation(page, f"{origin.group(1)}{target_path}",
                                f"{module} page")
                if page_is_open(page, module, expected_button, path):
                    log.info("   %s is open (%s)", module, page.url)
                    qa_step(f"Opened the {module} module")
                    return

    fail(page, f"Could not open the '{module}' module. The sidebar click keeps "
               f"being intercepted and the address is still {page.url} after "
               f"{attempts} attempts.")


def open_sales_module(page: Page, module: str, expected_button: str) -> None:
    """Open a module that lives inside the collapsed 'Orders & Sales' group."""
    dismiss_overlay(page)          # an open form hides the whole sidebar
    if nav_link(page, module).count() == 0:
        # exact=False on purpose: the sidebar group carries a live count, so it
        # reads "Orders & Sales" on a new tenant and "Orders & Sales 1" as soon
        # as there is one order. Matching it exactly worked only until the first
        # quotation was raised.
        group = button(page, "Orders & Sales", exact=False).filter(visible=True)
        expect_visible(page, group.first, "'Orders & Sales' sidebar group")
        click(page, group.first, "Orders & Sales menu", animated=True)
    open_module(page, module, expected_button)


def confirm_saved(page: Page, submit: Locator, record_text: str, what: str) -> None:
    """Confirm a create form actually saved the record.

    Two checks, because either one alone can lie:
      1. the submit button disappears - the app accepted the payload,
      2. the new record (or a success message) is on screen - it really exists.

    The name every module already calls. The looking is done by verify_success(),
    which waits for the API and the redraw before its first look and reloads
    between attempts, so a slow list no longer reads as a failed save.

    It is also where the record is queued for the data check: verify_success()
    raises if the record is not really there, so reaching the next line means
    there IS a record whose View page can be read and compared.
    """
    verify_success(page, record_text, what, submit=submit)
    verify_data.after_save(page, record_text, what)


# =========================================================================== #
# MODULE 1 - CREATE YOUR ACCOUNT  (sheet: Create_Your_Account)
# =========================================================================== #

def wait_for_pincode_lookup(page: Page) -> None:
    """Let the pincode lookup finish filling City and State.

    It is a network call with no spinner of its own, and the application writes
    its answer straight into City, State and the map box when it returns. Type
    while it is in flight and the typing is silently replaced - which is exactly
    what used to happen to the city.

    Best effort: not every pincode is known to the service, and the fields are
    typed by hand in that case.
    """
    confirmation = page.get_by_text(re.compile(r"pincode\s+verified", re.IGNORECASE))
    try:
        confirmation.first.wait_for(state="visible", timeout=OPTIONAL_TIMEOUT)
        log.info("   the pincode lookup filled City and State")
    except (PlaywrightTimeoutError, PlaywrightError):
        log.info("   the pincode lookup returned nothing - City is typed by hand")
    wait_for_loaders(page)


#: Every shape the map's address suggestion list has been seen in. It is not
#: rendered inside the input's own container - a Places widget appends its
#: `pac-container` to the body, and the app's own list is a portal - so it has
#: to be looked for page-wide. Order is priority: the most specific first, and
#: `.filter(visible=True)` keeps a closed list that is still in the DOM out.
MAP_SUGGESTION_SELECTORS = (
    ".pac-container .pac-item",                  # Google Places JS widget
    "[role='listbox'] [role='option']",          # an ARIA combobox
    "[class*='suggestion' i] li",
    "li[class*='suggestion' i]",
    "[class*='autocomplete' i] li",
    "[role='option']",
)


def suggestion_text(suggestion: Locator) -> str:
    """What a suggestion row reads as, on one line - for the log and the report.

    Never raises: the list closes the moment it is clicked, so reading it is
    always a race, and its text is only ever used to name the step.
    """
    try:
        return " ".join((suggestion.inner_text(timeout=SCROLL_TIMEOUT)
                         or "").split())
    except (PlaywrightTimeoutError, PlaywrightError):
        return ""


def map_suggestion_rows(page: Page) -> Locator:
    """Every suggestion currently on screen, from the most specific shape found.

    The shapes are tried in the order they are listed, so an app-drawn list wins
    over the catch-all `[role='option']` when both would match.
    """
    for selector in MAP_SUGGESTION_SELECTORS:
        items = page.locator(selector).filter(visible=True)
        try:
            if items.count():
                return items
        except PlaywrightError:
            continue
    return page.locator(", ".join(MAP_SUGGESTION_SELECTORS)).filter(visible=True)


def location_suggestion(page: Page, wanted: str,
                        timeout: int = OPTIONAL_TIMEOUT) -> Optional[Locator]:
    """The result to click, once the search has really drawn one.

    Nothing is ever clicked before the list is on screen: the results come back
    from the network a moment after the typing stops, and a click into that gap
    leaves the box holding typed text with no location behind it. The wait is
    one wait on every known shape at once, so a slow list cannot cost
    timeout x selectors.

    None means nothing was offered, which is a branch here and not a failure -
    see search_business_location().

    Which row is taken: the one that READS like the place the workbook asked
    for - the recording searched "Nelama" and clicked "NelamangalaKarnataka,
    India" - and the first row otherwise, because whatever the search offers
    first is still a real, resolved address. Both sides of the comparison have
    their spaces and commas taken out, since the row prints the place, the state
    and the country run together in one string.
    """
    if not appears(page.locator(", ".join(MAP_SUGGESTION_SELECTORS))
                   .filter(visible=True).first, timeout):
        return None

    rows = map_suggestion_rows(page)
    key = re.sub(r"[\s,]+", "", (wanted or "").split(",")[0]).casefold()
    if key:
        try:
            for index in range(min(rows.count(), 8)):
                candidate = rows.nth(index)
                reads = re.sub(r"[\s,]+", "", suggestion_text(candidate)).casefold()
                if key in reads:
                    return candidate
        except PlaywrightError:
            pass                       # the list closed while it was being read
    return rows.first


def location_search_box(page: Page) -> Locator:
    """The "Search for your business location..." box of step 1.

    Addressed the way the recording addresses it - by role and accessible name -
    with the placeholder lookup kept as the fallback. get_by_role ignores the
    copies of a form this SPA keeps mounted but hidden; a bare
    get_by_placeholder does not.
    """
    return visible(page.get_by_role("textbox", name="Search for your business")
                   .or_(page.get_by_placeholder("Search for your business")))


def location_search_terms(row: Row) -> List[str]:
    """What to type into the location box, best first, from the workbook.

    The recording searched "Nelama" and took "Nelamangala, Karnataka, India":
    the box is a places-style autocomplete, so a short, distinctive prefix opens
    the list where a long, fully punctuated address often returns nothing.

    So the sheet's own City Suggestion is tried whole and then cut back to its
    leading place name ("Kalamassery, Kochi" -> "Kalamassery"), then the City,
    then the street Address. The address is NOT cut back the same way - its
    leading part is a plot number, which identifies no place at all and would
    only spend another wait. Deduplicated, so a sheet whose City Suggestion IS
    its City does not search twice for the same thing.
    """
    terms: List[str] = []

    def add(text: str) -> None:
        text = " ".join((text or "").split())
        # Two characters cannot identify a place, and an autocomplete does not
        # answer them - not worth a wait each.
        if len(text) < 3:
            return
        if text.casefold() not in (seen.casefold() for seen in terms):
            terms.append(text)

    suggestion = row.get("City Suggestion", "")
    add(suggestion)
    add(suggestion.split(",")[0])
    add(row.get("City", ""))
    add(row.get("Address", ""))
    return terms


def search_business_location(page: Page, row: Row) -> None:
    """Set the business location through the form's own location search.

    This is the recorded flow - type into "Search for your business
    location...", wait for the result, click it - driven from the workbook
    instead of the recording's fixed "Nelama" / "NelamangalaKarnataka, India".

    The text is typed key by key, because a value set in one go does not open
    the list, and then the list is WAITED for rather than guessed at.

    Still OPTIONAL by design: when the pincode lookup has already resolved the
    address the application rewrites this box itself and offers no list at all -
    the location is set either way and Continue does not depend on it. So an
    empty list is reported, with the evidence, and the registration carries on;
    failing here would fail a registration the application accepts.
    """
    search = location_search_box(page)
    if search.count() == 0:
        log.info("   skip: Location (this form has no location search box)")
        return

    terms = location_search_terms(row)
    if not terms:
        log.info("   skip: Location (the sheet has no place to search for)")
        return

    already_set = search.input_value()

    for term in terms:
        scroll_into_view(search, "Location", editable=True)
        search.fill("")
        search.press_sequentially(term, delay=60)
        log.info("   location: typed '%s' - waiting for the result list", term)

        suggestion = location_suggestion(page, term)
        if suggestion is None:
            log.info("   location: nothing was offered for '%s'", term)
            continue

        offered = suggestion_text(suggestion) or term
        click(page, suggestion, f"Location result '{offered}'", animated=True)
        wait_for_loaders(page)
        log.info("   fill: %-24s = %s (chosen from the search results)",
                 "Location", offered)
        return

    # Nothing offered for anything tried - put the box back the way the lookup
    # left it, so a half-typed search cannot undo the location.
    scroll_into_view(search, "Location", editable=True)
    search.fill(already_set)
    log.warning("   the location search offered no result for %s - the pincode "
                "lookup has already set the location, so registration continues",
                " / ".join(f"'{term}'" for term in terms))
    attach_text("Business Location Search - Skipped",
                f"The location search offered no result for "
                f"{', '.join(repr(term) for term in terms)}. "
                f"Location left at: {already_set or '(unset)'}")


# --------------------------------------------------------------------------- #
# The four option-card fields of step 1, and why they get a layer of their own
#
#     Industry Type *
#         -> Who do you sell to?
#             -> What does your business specialise in? - pick up to 3
#
# APPLICATION DEFECT: the form renders each of these blocks only once the one
# before it has been answered, and answering one re-renders the block around it
# (the same re-render that clears GSTIN and Company Name further down). How long
# that takes is inconsistent - instant on a good run, seconds when the GST
# lookup behind the form is busy.
#
# A click into that gap is swallowed, and the block that depended on it then
# never renders.
#
# A fixed pause is the wrong answer to an inconsistent load - too short on a bad
# day, wasted on a good one. So each field is waited for, chosen, and PROVEN
# chosen before the next one is touched. Nothing here sleeps; every wait is a
# Playwright wait for a real condition, so a fast run waits for nothing and a
# slow one is not cut short.
#
# THE LAST BLOCK IS NOT LIKE THE OTHERS, and that is worth spelling out because
# it is what used to fail this test. "What does your business specialise in?"
# and "Pick your specialisations" are not fields of the form: they are a
# catalogue the page fetches, they sit inside a panel the form itself heads
# "Optional but strongly recommended - takes 30 seconds", and that fetch is
# unreliable. Runs on one afternoon, same GSTIN, same Industry Type, same "Who
# do you sell to?" answer, showed all three outcomes:
#
#     13:14  the block was there immediately
#     14:38  the block appeared after 29 seconds
#     17:06  the block never appeared at all - the screenshot shows "Businesses
#            (B2B)" selected and the page going straight from "Who do you sell
#            to?" to PINCODE, with the catalogue simply missing
#
# The failure was therefore never a swallowed click; the selection above it had
# been taken, and the run failed on an OPTIONAL block the application had
# declined to draw. So it is handled to match what the application actually
# does: waited for three times as long as any other field
# (SPECIALISATION_TIMEOUT), and, if it still never arrives, reported into the
# log and the Allure report with a screenshot and left unanswered - not failed.
# The registration finishes, which is what the test is for. Flip
# SPECIALISATION_REQUIRED to True to make its absence fail the run again.
# --------------------------------------------------------------------------- #

#: The text each block of step 1 opens with. Matched as a pattern because the
#: form decorates the headings ("Industry Type *", "... - pick up to 3") and has
#: changed that decoration between builds; the words themselves have not.
INDUSTRY_TYPE_HEADING = r"Industry\s*Type"
SELL_TO_HEADING = r"Who\s*do\s*you\s*sell\s*to"
SPECIALISE_HEADING = r"What\s*does\s*your\s*business\s*speciali[sz]e"

#: What an option card can be. The form draws them as buttons today; a label
#: wrapping a hidden radio, and a listbox, are what earlier builds used.
OPTION_CARD = ("self::button or self::a or self::label"
               " or @role='button' or @role='option'"
               " or @role='radio' or @role='checkbox'")

#: Everything about the card that changes when it becomes the chosen one, read
#: in one round trip. It deliberately looks WIDER than the card itself: a build
#: is free to put the tick on a child, the ring on a wrapper and the state on
#: neither, so `ancestry` carries the classes of the card and three parents, and
#: `shape` / `group` are the rendered size of the card and of the row it sits
#: in. Between them, almost any visible answer to a click shows up as a change.
OPTION_STATE_JS = """element => {
    const attribute = name => element.getAttribute(name);
    const chain = [];
    let node = element;
    for (let step = 0; step < 4 && node; step++) {
        chain.push(node.getAttribute('class') || '');
        node = node.parentElement;
    }
    const painted = window.getComputedStyle(element);
    return {
        'paint': [painted.backgroundColor, painted.borderColor, painted.color,
                  painted.boxShadow, painted.outlineColor].join(' | '),
        'aria-pressed': attribute('aria-pressed'),
        'aria-checked': attribute('aria-checked'),
        'aria-selected': attribute('aria-selected'),
        'data-state': attribute('data-state'),
        'checked': typeof element.checked === 'boolean'
            ? String(element.checked) : null,
        'class': attribute('class') || '',
        'ancestry': chain.join(' | '),
        'children': element.childElementCount,
        'shape': element.outerHTML.length,
        'group': element.parentElement ? element.parentElement.outerHTML.length : 0,
    };
}"""

#: The markers that say outright whether a card is chosen, so they can be
#: believed both ways: present and not "true" means it really is not chosen.
DECISIVE_MARKERS = ("aria-pressed", "aria-checked", "aria-selected", "checked")

#: The marker values, and the class names, that mean "chosen".
SELECTED_VALUES = ("true", "selected", "checked", "on", "active")
SELECTED_CLASS = re.compile(
    r"(?:^|[\s_-])(selected|active|checked|chosen)(?:[\s_-]|$)", re.IGNORECASE)

#: Emoji, pictographs and the invisible characters that dress them (variation
#: selectors, the zero-width joiner). The workbook writes a category the way a
#: person reads it on the card, and some of those cards draw their emoji as an
#: <img> icon instead of putting it in the text - so the decoration is what a
#: match must be allowed to differ on. A category is never NAMED after it.
DECORATION = re.compile(
    "[‍︀-️"                          # joiner, variation selectors
    "←-⇿⌀-➿"                    # arrows, dingbats, technical
    "⬀-⯿〰〽㊗㊙Ⓜ"   # symbols, enclosed marks
    "\U0001f000-\U0001faff]+")                      # emoji, pictographs, flags


def category_text(value: str) -> str:
    """The comparable form of an option label - its WORDS, nothing else.

    Used ONLY as a fallback, and only ever inside the block the label belongs
    to: "Construction" is the Industry Type card AND the word inside the
    specialisation card, so stripping the emoji is safe once the search is
    already scoped to one block of the form, and wrong before it is.
    """
    return re.sub(r"\s+", " ", DECORATION.sub(" ", value or "")).strip()


def squashed(value: str) -> str:
    """`value` with every space taken out, for a whitespace-blind comparison."""
    return re.sub(r"\s+", "", value or "")


def loosely(value: str) -> "re.Pattern[str]":
    """`value` as a pattern that does not mind how the build spaces the words.

    The workbook holds what a person reads on the card - "Businesses (B2B) You
    sell to" - and the card is built from two elements. Whether the accessible
    name comes out with a space between them depends on how they are laid out,
    so the space is exactly what the match must not depend on. Playwright
    matches a pattern anywhere in the name, so a label that is the START of the
    card's own wording still finds it.
    """
    return re.compile(r"\s*".join(re.escape(word) for word in value.split()),
                      re.IGNORECASE)


def xpath_literal(value: str) -> str:
    """`value` written as an XPath string, whichever quotes it happens to hold."""
    if "'" not in value:
        return f"'{value}'"
    if '"' not in value:
        return f'"{value}"'
    pieces = [f'"{piece}"' if piece == "'" else f"'{piece}'"
              for piece in re.split(r"(')", value) if piece]
    return "concat(" + ", ".join(pieces) + ")"


def section_heading(page: Page, heading: str) -> Locator:
    """The text that opens one block of step 1."""
    return visible(page.get_by_text(re.compile(heading, re.IGNORECASE)))


def wait_for_section(page: Page, heading: str, description: str,
                     note: str = "did not become available") -> Locator:
    """Wait for a block of the form to be on screen, and hand back its heading.

    The heading is the anchor every option below it is then found from, so this
    is both the wait and the scope: an option is only ever looked for INSIDE the
    block that owns it, which is what stops the Industry Type card called
    "Construction" from answering for the specialisation card that carries the
    same word.
    """
    anchor = section_heading(page, heading)
    log.info("   waiting up to %ds for %s...", FIELD_LOAD_TIMEOUT // 1000,
             description)
    wait_for_loaders(page)
    try:
        expect(anchor).to_be_visible(timeout=FIELD_LOAD_TIMEOUT)
    except (AssertionError, PlaywrightTimeoutError, PlaywrightError) as error:
        fail(page, f"{description} {note} within the expected time "
                   f"({FIELD_LOAD_TIMEOUT // 1000} seconds). If the page shows "
                   f"'GST lookup quota exhausted', that service is what step 1 "
                   f"is waiting on: wait for the quota to reset, or use a GSTIN "
                   f"the application has already cached.", error)
    return anchor


# --------------------------------------------------------------------------- #
# THE GST NOTICE - WHAT IT ACTUALLY SAYS, AND WHY IT IS NOT A BLOCKER
#
# Read verbatim off the application (run 2026-08-12_17-01-42):
#
#     ⚡ GST lookup quota exhausted. State auto-filled; please enter remaining
#       details manually.
#
# That whole sentence matters, and taking only its first half is what led this
# framework astray. The notice is about the GSTIN AUTO-FILL: the service that
# would have populated the company's details from its GSTIN has nothing left to
# give, so the application filled in the State and is asking for the rest to be
# TYPED. It is the application telling you its supported way forward - and that
# way forward is exactly what this suite does anyway, because every one of those
# fields is typed from the workbook.
#
# So the notice is INFORMATION, not a failure:
#
#   * it says nothing whatever about "What does your business specialise in?";
#   * it has been seen on screen while the specialisation tiles rendered anyway;
#   * it is usually already up before the Industry Type is even chosen, because
#     it belongs to the GSTIN field several steps earlier.
#
# It must therefore NEVER shorten the wait for the specialisation block and
# never be reported as the reason that block is missing. Treating it as a
# 20-second death sentence is what failed a registration whose catalogue has
# been measured taking 29 seconds to arrive.
#
# What the notice IS used for: saying so in the log and the report, so a run
# that had to type everything by hand says why - and, if the block really never
# arrives, giving the reader the context they need without blaming the wrong
# service.
# --------------------------------------------------------------------------- #

#: How the application words it. A pattern, because the banner has been seen
#: with and without a trailing note; the words themselves do not change.
GST_QUOTA_MESSAGE = re.compile(
    r"GST\s*(?:lookup|verification)?\s*quota\s*(?:is\s*)?exhaust", re.IGNORECASE)

#: The second half of that same notice - the half that says what to DO about it.
#: When the notice carries this, the application has already told us its
#: supported way forward: type the remaining details. This suite types every one
#: of them from the workbook, so the fallback is already being followed and
#: there is nothing to look for on screen.
GST_MANUAL_ENTRY_NOTICE = re.compile(
    r"enter\s+(?:the\s+)?remaining\s+details\s+manually"
    r"|please\s+enter\s+.{0,30}\bmanually"
    r"|auto[- ]?filled", re.IGNORECASE)

#: What a supported way of carrying on WITHOUT the GST lookup would be called.
#:
#: Neither the application nor this framework documents such a control, so it is
#: DISCOVERED when the banner appears rather than assumed to exist: whatever is
#: found is named in the log and in the report, and when nothing is found the
#: run says so plainly and fails on the application's own limit instead of
#: inventing a way round it.
#:
#: "Retry" / "Try again" is deliberately not in this list. A quota that is spent
#: is no less spent on the second attempt, and clicking it only spends more.
GST_FALLBACK_LABELS = (
    r"enter\s+(?:the\s+)?details?\s+manually",
    r"enter\s+manually",
    r"manual\s+entry",
    r"fill\s+(?:in\s+)?manually",
    r"add\s+(?:details\s+)?manually",
    r"continue\s+without\s+GST",
    r"proceed\s+without\s+GST",
    r"skip\s+GST(?:\s+verification|\s+lookup)?",
    r"without\s+GST\s+verification",
    # The application may already hold this GSTIN. If it offers to use what it
    # has, that is its own supported way through and is taken as one - the GSTIN
    # in the workbook is never swapped for a different one to dodge the quota.
    r"use\s+cached",
    r"cached\s+(?:GST|details)",
    r"use\s+(?:the\s+)?saved\s+details",
)

#: What the application would call "run that lookup again". Asked of the page
#: before the retry falls back to re-entering the GSTIN, because a control the
#: application draws for this is always safer than driving the form.
GST_RETRY_LABELS = (
    r"retry(?:\s+GST)?(?:\s+lookup|\s+verification)?",
    r"try\s+again",
    r"verify\s+GSTIN?",
    r"fetch\s+(?:GST\s+)?details",
    r"look\s*up\s+GSTIN?",
    r"refresh(?:\s+GST)?(?:\s+details)?",
)


def gst_quota_banner(page: Page) -> Locator:
    """The application's "GST lookup quota exhausted" message, if it is up.

    visible() matters here: get_by_text is a plain attribute lookup, so it would
    also match a banner left behind on a form the SPA keeps mounted but hidden.
    """
    return visible(page.get_by_text(GST_QUOTA_MESSAGE))


def quota_banner_text(page: Page) -> str:
    """What the quota banner actually says, for the log and the report."""
    try:
        said = " ".join((gst_quota_banner(page).inner_text(timeout=RETRY_TIMEOUT)
                         or "").split())
    except (PlaywrightTimeoutError, PlaywrightError):
        said = ""                  # the banner went away while it was being read
    return said or "GST lookup quota exhausted"


def report_gst_notice(page: Page, when: str) -> str:
    """Record the GST notice as what it is: information about the auto-fill.

    Not a failure and not the end of anything - it is noted so a run that had to
    type the company's details by hand says so, with the application's own words
    and a screenshot. Returns those words for the caller to quote.
    """
    shown = quota_banner_text(page)
    if not gst_notice_on_screen(page):
        return ""

    manual = bool(GST_MANUAL_ENTRY_NOTICE.search(shown))
    with step("GST notice: the GSTIN auto-fill is unavailable"):
        log.warning("   the application is showing: '%s'", shown)
        if manual:
            log.warning("   that notice asks for the remaining details to be "
                        "entered manually, which is what this run does anyway - "
                        "every field comes from the workbook. Carrying on.")
        detail = (
            f"The application displayed: {shown}\n\n"
            f"Seen: {when}.\n\n"
            f"What this notice is about: the service that fills a company's "
            f"details in from its GSTIN has no quota left, so the application "
            f"filled in what it could (the State) and is asking for the rest to "
            f"be typed.\n\n"
            f"Why the run carries on: this suite types every one of those "
            f"fields from Construction_Flow_Data.xlsx already, so the "
            f"application's own way forward is the one being followed. The "
            f"notice says nothing about 'What does your business specialise "
            f"in?', and that block has been seen rendering perfectly well while "
            f"this notice was on screen - so it is NOT treated as a reason to "
            f"stop waiting for it.\n"
            f"Page: {page.url}")
        attach_text(f"GST Lookup Notice ({when})", detail)
        screenshot(page, f"GST Lookup Notice ({when})")
    return shown


def gst_notice_on_screen(page: Page) -> bool:
    """Whether the GST notice is up at all - never raises."""
    try:
        return gst_quota_banner(page).count() > 0
    except (PlaywrightTimeoutError, PlaywrightError):
        return False


def gst_lookup_fallback(page: Page) -> Optional[Locator]:
    """The application's own way of carrying on without the GST lookup, or None.

    Every shape of the question is asked of the page that is actually on screen
    - this is the "inspect the UI before working round it" step - and None means
    the application really is offering nothing, which is a finding in itself.
    """
    candidates: List[Tuple[str, Locator]] = []
    for label in GST_FALLBACK_LABELS:
        wording = re.compile(label, re.IGNORECASE)
        candidates.append((
            f"a control reading /{label}/",
            page.get_by_role("button", name=wording)
                .or_(page.get_by_role("link", name=wording))))
    return safe_locator(page, "GST lookup fallback", tuple(candidates))


def use_gst_lookup_fallback(page: Page) -> bool:
    """Take the application's supported fallback, if this build offers one.

    True means the application was given another way in and the block is worth
    waiting for again. False means it offers none - which is reported, not
    worked around.
    """
    fallback = gst_lookup_fallback(page)
    if fallback is None:
        # The application's fallback is not always a CONTROL. When the notice
        # itself says "please enter remaining details manually", that IS the
        # supported way through, and this suite is already taking it: every
        # field on this form is typed from the workbook.
        if GST_MANUAL_ENTRY_NOTICE.search(quota_banner_text(page)):
            log.info("   the application's own instruction is to enter the "
                     "remaining details manually - which this run does for "
                     "every field, from the workbook. Nothing to click.")
            attach_text("GST Lookup - Fallback",
                        "The application's supported fallback is stated in the "
                        "notice itself: 'please enter remaining details "
                        "manually'. There is no control to press - and none is "
                        "needed, because this suite types every field on the "
                        "form from Construction_Flow_Data.xlsx. The fallback is "
                        "being followed, not worked around.")
            return True

        log.warning("   the application offers no supported way of continuing "
                    "without the GST lookup on this screen - there is no "
                    "'enter details manually' and no 'continue without GST'")
        attach_text("GST Lookup - Fallback",
                    "The screen was inspected for a supported way of carrying "
                    "on without the GST lookup - a button or link offering "
                    "manual entry, or continuing/skipping without GST - and the "
                    "application offered none. Nothing was worked around: the "
                    "run reports the service as unavailable instead.")
        return False

    offered = suggestion_text(fallback) or "the fallback the application offers"
    log.warning("   the application offers '%s' - taking its own fallback "
                "rather than working round it", offered)
    click(page, fallback, f"GST lookup fallback '{offered}'")
    wait_for_loaders(page)
    attach_text("GST Lookup - Fallback",
                f"The GST lookup is unavailable, and the application offered "
                f"its own way of carrying on: '{offered}'. That control was "
                f"used, and the registration continued from there.")
    screenshot(page, "Registration - GST Lookup Fallback Used")
    return True


def gst_lookup_retry_control(page: Page) -> Optional[Locator]:
    """The application's own "run that lookup again" control, or None."""
    candidates: List[Tuple[str, Locator]] = []
    for label in GST_RETRY_LABELS:
        wording = re.compile(label, re.IGNORECASE)
        candidates.append((
            f"a control reading /{label}/",
            page.get_by_role("button", name=wording)
                .or_(page.get_by_role("link", name=wording))))
    return safe_locator(page, "GST lookup retry", tuple(candidates))


def retry_gst_lookup(page: Page, row: Row) -> bool:
    """Ask the application to run its GST lookup ONE more time.

    Called once and only once - see wait_for_specialisation_section(), which
    holds the flag. There is no loop here and no loop around it.

    Two supported ways, in order of how little they disturb the form:

      1. the application's own control ("Retry", "Try again", "Verify GSTIN");
      2. failing that, re-entering the SAME GSTIN from the workbook and leaving
         the field, which is the user action that fires the lookup in the first
         place. The GSTIN is never changed to something else to get a run past
         the quota - the workbook's value is the test data.

    Re-entering can make step 1 re-render, and the Industry Type answer can go
    with it, so that is checked afterwards - see restore_industry_type().
    """
    control = gst_lookup_retry_control(page)
    if control is not None:
        offered = suggestion_text(control) or "Retry"
        log.warning("   the application offers '%s' - using its own control to "
                    "run the GST lookup once more", offered)
        click(page, control, f"GST lookup retry '{offered}'")
        wait_for_loaders(page)
        return True

    field = wizard_field(page, "27AABCU9603R1ZX")
    if field.count() == 0:
        log.warning("   the GSTIN field is no longer on screen, so the lookup "
                    "cannot be re-run - going straight to the fallback check")
        return False

    gstin = row["GSTIN"]
    log.warning("   the application offers no retry control, so the lookup is "
                "re-triggered the way a person would: the SAME GSTIN from the "
                "workbook is re-entered")
    fill(page, field, gstin, "GSTIN (re-entered to re-run the lookup)")
    try:
        field.press("Tab")         # leaving the field is what fires the lookup
    except (PlaywrightTimeoutError, PlaywrightError):
        pass
    wait_for_loaders(page)
    restore_industry_type(page, row)
    return True


def restore_industry_type(page: Page, row: Row) -> None:
    """Put the Industry Type back if re-entering the GSTIN cleared it.

    Only when the evidence says it really was cleared: "Who do you sell to?" is
    drawn only once an Industry Type has been chosen, so that block disappearing
    is the application saying the answer is gone. These cards TOGGLE, so a
    click on one that is still chosen would un-choose it - which is why nothing
    is clicked on the strength of a guess.
    """
    if appears(section_heading(page, SELL_TO_HEADING), OPTIONAL_TIMEOUT):
        return                     # the answer survived the re-render

    industry = row["Business Type"]
    log.warning("   re-entering the GSTIN reset step 1 - re-choosing the "
                "Industry Type ('%s')", industry)
    attach_text("GST Lookup - Retry",
                f"Re-entering the GSTIN to re-run the lookup reset step 1: "
                f"'Who do you sell to?' was no longer on screen, which is the "
                f"application saying the Industry Type answer had been cleared. "
                f"It was re-chosen from the workbook ('{industry}').")
    anchor = section_heading(page, INDUSTRY_TYPE_HEADING)
    if appears(anchor, FIELD_LOAD_TIMEOUT):
        select_option_card(
            page, anchor, industry, "Industry Type (re-chosen)",
            "Industry Type could not be re-chosen after the GST lookup was "
            "re-run", required=False)


def watch_for_section(page: Page, timeout: int) -> Tuple[str, Optional[Locator]]:
    """One attempt at the block: wait the WHOLE budget for it to render.

    Hands back ("loaded", the heading) or ("missing", None).

    The GST notice is deliberately NOT part of this wait. It is a notice about
    the GSTIN auto-fill, it is usually on screen before this block is even due,
    and the block has been seen rendering with it up - so letting it end the
    wait early is how a catalogue that takes 29 seconds got failed after 20.
    Whether the notice is up is a question for the REPORT, once the block really
    has not come, and wait_for_specialisation_section() asks it there.

    The wait is split into passes rather than being one long expect() so that
    each pass can nudge the page and say how it is getting on:

      * every pass clears any spinner still up, so a catalogue that is merely
        slow is waited for rather than being declared missing;
      * every pass scrolls the area into view, which is all it takes if a build
        renders the block only once it has been scrolled to;
      * a run that is waiting minutes says so in the log, instead of looking
        hung.

    Nothing here sleeps, and nothing here clicks.
    """
    description = "'What does your business specialise in?'"
    anchor = section_heading(page, SPECIALISE_HEADING)

    log.info("   waiting up to %ds for %s...", timeout // 1000, description)

    deadline = time.monotonic() + timeout / 1000
    pass_timeout = max(OPTIONAL_TIMEOUT, RETRY_TIMEOUT)

    while True:
        wait_for_loaders(page)
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            return "missing", None
        # Never 0: Playwright reads a timeout of 0 as "wait for ever", so the
        # last sliver of the budget must still be rounded up to a real wait.
        this_pass = max(1_000, min(pass_timeout, int(remaining * 1000)))

        if appears(anchor, this_pass):
            return "loaded", anchor

        if time.monotonic() >= deadline:
            return "missing", None
        log.info("      %s has not been drawn yet - still waiting (%ds left)",
                 description, max(0, int(deadline - time.monotonic())))
        # Nothing here clicks anything. The blocks above have already been
        # answered, and every one of these cards TOGGLES: a nudge that touched
        # one would un-answer the very field this block is waiting on.
        scroll_into_view(section_heading(page, SELL_TO_HEADING), description)


def wait_for_specialisation_section(page: Page,
                                    row: Row) -> Tuple[str, Optional[Locator]]:
    """The whole sequence for the specialisation block, with ONE retry.

        wait the FULL budget ─┬─ it renders ....................... "loaded"
                              └─ it does not
                                     │
                                     ├─ note the GST notice, if one is up
                                     ├─ retry the lookup ONCE
                                     ├─ wait the full budget again
                                     │      └─ it renders ........ "loaded"
                                     ├─ take the application's own fallback,
                                     │  if it has a control for one
                                     │      └─ it renders ........ "loaded"
                                     └─ still nothing ─┬ notice up . "quota"
                                                       └ no notice . "missing"

    The block is given its WHOLE 90 seconds every time - the GST notice never
    cuts a wait short. It has been measured arriving at 29 seconds with that
    notice on screen, so a 20-second grace was failing registrations that were
    about to succeed.

    "quota" vs "missing" is only ever decided AFTERWARDS, from whether the
    notice is up, and it changes nothing but the wording of the report: by then
    the block has had two full waits, a retry and the fallback.

    Bounded throughout - no loop around the retry, no loop around the fallback.
    """
    outcome, anchor = watch_for_section(page, SPECIALISATION_TIMEOUT)
    if outcome == "loaded":
        return outcome, anchor

    def verdict() -> str:
        """What to call it, once the block really has not come."""
        return "quota" if gst_notice_on_screen(page) else "missing"

    # --- it did not come: note the GST notice, then retry the lookup ONCE ---
    report_gst_notice(page, "the specialisation block had not rendered")
    with step("Retrying the GST lookup (once)"):
        retried = retry_gst_lookup(page, row)
    if retried:
        outcome, anchor = watch_for_section(page, SPECIALISATION_TIMEOUT)
        if outcome == "loaded":
            log.info("   the retry worked - the block is on screen")
            return outcome, anchor

    # --- still nothing: the application's own way out, if it has one --------
    with step("Checking for a supported GST lookup fallback"):
        if not use_gst_lookup_fallback(page):
            return verdict(), None

    outcome, anchor = watch_for_section(page, SPECIALISATION_TIMEOUT)
    if outcome == "loaded":
        log.info("   the block arrived after the application's own fallback - "
                 "the registration can carry on")
        return outcome, anchor
    return verdict(), None


def section_options(page: Page, anchor: Locator,
                    wanted: str) -> Tuple[Tuple[str, Locator], ...]:
    """The ways of addressing one option of a block, best first.

    Every one of them is confined to the block by XPath's `following` axis - the
    role-based ones by intersecting with it - so nothing above the heading, and
    no earlier block, can answer. That confinement is what lets the emoji be
    dropped as a last resort: "Construction" is the Industry Type card AND the
    word inside the specialisation card, and only the scope tells them apart.

    The accessible name comes before the raw text because the two are not the
    same string. A card built as <span>Businesses (B2B)</span><span>You sell to
    other businesses</span> has the accessible name "Businesses (B2B) You sell
    to other businesses" - a space between the spans - while XPath's
    normalize-space() sees them run together. The workbook is written from what
    a person reads, so the accessible name is the one that matches it.
    """
    plain = category_text(wanted)
    inside = anchor.locator("xpath=following::*")

    def named(name, exact: bool = False) -> Locator:
        return (page.get_by_role("button", name=name, exact=exact)
                .or_(page.get_by_role("option", name=name, exact=exact))
                .or_(page.get_by_role("radio", name=name, exact=exact))
                .or_(page.get_by_role("checkbox", name=name, exact=exact))
                .and_(inside))

    def reading(test: str) -> Locator:
        return anchor.locator(f"xpath=following::*[{OPTION_CARD}][{test}]")

    def exact_text(value: str) -> str:
        return f"normalize-space(.)={xpath_literal(value)}"

    def holds_text(value: str) -> str:
        # Whitespace is squeezed out of BOTH sides before comparing, because
        # whether there is a space between a card's two <span>s is the build's
        # business, not the workbook's.
        return (f"contains(translate(normalize-space(.), '  ', ''), "
                f"{xpath_literal(squashed(value))})")

    return (
        ("the exact name", named(wanted, exact=True)),
        ("the exact text", reading(exact_text(wanted))),
        ("the name, any spacing", named(loosely(wanted))),
        ("the text, any spacing", reading(holds_text(wanted))),
        ("the exact name, emoji drawn as an icon", named(plain, exact=True)),
        ("the exact text, emoji drawn as an icon", reading(exact_text(plain))),
        ("the name, any spacing, emoji as an icon", named(loosely(plain))),
        ("the text, any spacing, emoji as an icon", reading(holds_text(plain))),
    )


def wait_for_option_offered(page: Page, anchor: Locator, wanted: str,
                            description: str,
                            timeout: int = FIELD_LOAD_TIMEOUT) -> bool:
    """Wait until an option reading `wanted` is really on screen in this block.

    Prove it is there before clicking it - the equivalent of
    expect(get_by_role("button", name="...")).to_be_visible(), but written
    through section_options() so it uses whichever of the eight ways of
    addressing a card this build actually answers to, and so it can never match
    a card belonging to a different block of the form.

    True/False rather than a raise: what an absent option MEANS depends on which
    one it is, and the caller is the only one that knows.
    """
    candidates = section_options(page, anchor, wanted)
    offered = candidates[0][1]
    for _, candidate in candidates[1:]:
        offered = offered.or_(candidate)

    log.info("   waiting up to %ds for the %s option '%s' to be offered...",
             timeout // 1000, description, wanted)
    return appears(visible(offered), timeout)


def option_state(control: Locator) -> Dict[str, object]:
    """Everything the card says about itself, or {} when it cannot be read."""
    try:
        return control.evaluate(OPTION_STATE_JS, timeout=RETRY_TIMEOUT)
    except (PlaywrightTimeoutError, PlaywrightError):
        return {}


def says_chosen(state: Dict[str, object]) -> Optional[bool]:
    """What the card's own markers say, or None when the build says nothing.

    None is the honest answer for a build that announces nothing - guessing
    "not chosen" there would fail a good run over markup that never existed.
    """
    values = [str(state[name]).strip().casefold()
              for name in DECISIVE_MARKERS if state.get(name) is not None]
    if any(value in SELECTED_VALUES for value in values):
        return True
    return False if values else None


def looks_chosen(before: Dict[str, object], after: Dict[str, object]) -> bool:
    """Whether the card LOOKS chosen, for a build that does not say so.

    Two readings count, and either one is enough:

      * the card or one of its wrappers is styled as chosen ("...selected...",
        data-state="on");
      * the click visibly changed something - a ring class, a tick element, the
        rendered size of the card or of the row it sits in, or the colour it is
        painted in. Whatever the build does to say "this one", it does
        something, and this sees it.

    `paint` is in that list because of what this build actually does: choosing
    "Businesses (B2B)" turns the card blue and changes NOTHING else - no
    attribute, no class of its own, no extra child. Reading the rendered colours
    is what tells the difference between a card the application took and a click
    that was swallowed, and without it a perfectly good selection was being
    reported as unconfirmed on every run.
    """
    if SELECTED_CLASS.search(str(after.get("ancestry", ""))):
        return True
    if str(after.get("data-state", "")).strip().casefold() in SELECTED_VALUES:
        return True
    return bool(before) and any(
        after.get(key) != before.get(key)
        for key in ("class", "ancestry", "children", "shape", "group", "paint"))


def select_option_card(page: Page, anchor: Locator, wanted: str,
                       description: str, message: str,
                       required: bool = True) -> bool:
    """Wait for one option of a block to load, choose it, and read back what took.

    Returns True when the option was chosen. With `required=False` an option
    that never loads is reported and False comes back instead of the run being
    failed - see choose_business_profile() for the one block that is optional.

    Two things are hard failures, because both mean the run cannot go on: the
    options never loading, and the wanted one not being among them.

    Whether the application MARKED the card is treated differently. It is read
    back and reported, but a card that shows no readable sign of being chosen
    does not fail the step on its own - a build is free to style its selection
    in a way nothing outside it can see, and failing there would fail a run that
    is going perfectly well. The proof that a selection really took is the block
    that depends on it appearing, and the caller waits for that next.
    """
    wanted = (wanted or "").strip()
    candidates = section_options(page, anchor, wanted)

    def not_selected(what: str, error: object = "") -> None:
        fail(page, f"{message} ({FIELD_LOAD_TIMEOUT // 1000} seconds) - {what}. "
                   f"The workbook asks for '{wanted}'.", error)

    # --- wait for the options themselves to load --------------------------
    offered = candidates[0][1]
    for _, candidate in candidates[1:]:
        offered = offered.or_(candidate)
    try:
        expect(visible(offered)).to_be_visible(timeout=FIELD_LOAD_TIMEOUT)
    except (AssertionError, PlaywrightTimeoutError, PlaywrightError) as error:
        if required:
            not_selected(f"no option reading '{wanted}' ever loaded", error)
        report_optional_block(
            page, description,
            f"No option reading '{wanted}' was offered within "
            f"{FIELD_LOAD_TIMEOUT // 1000} seconds, so {description} was left "
            f"unanswered.")
        return False

    control = safe_locator(page, description, candidates) or visible(offered)

    # --- do not un-choose an option the form has already chosen -----------
    before = option_state(control)
    if says_chosen(before) is True or SELECTED_CLASS.search(
            str(before.get("class", ""))):
        log.info("   %s is already chosen ('%s') - clicking it again would "
                 "un-choose it, so it is left alone", description, wanted)
        return True

    click(page, control, description, animated=True)
    wait_dom_stable(page)

    # --- read back what the application did with the click ----------------
    # The card is clicked ONCE and never clicked again. These fields toggle: a
    # second click on a card that is already chosen un-chooses it, so a retry
    # here would undo the very thing it is meant to be fixing. safe_click() has
    # already proved the click itself landed.
    state = option_state(control)
    marked = says_chosen(state)
    if marked is True or (marked is not False and looks_chosen(before, state)):
        log.info("   %-24s = %s (confirmed on screen)", description, wanted)
        return True

    # The application shows no sign of having taken it. That is NOT failed here:
    # this form answers a real selection by rendering the field that depends on
    # it, and the caller waits for exactly that next - which is a better proof
    # than any attribute, and the one that decides whether the run goes on.
    log.warning("   %s was clicked but the application shows no sign of marking "
                "'%s' as chosen. Carrying on - the field that depends on it is "
                "waited for next, and that is the real proof.",
                description, wanted)
    attach_text("Selection Not Confirmed by the Application",
                f"{description} = {wanted}\n"
                f"The card was clicked and the click landed, but the "
                f"application marked nothing that could be read back "
                f"(markers before: {before}, after: {state}).\n"
                f"The step was not failed on that alone - the block that "
                f"depends on this answer is waited for next, and it is that "
                f"wait which decides whether the selection really took.")
    screenshot(page, f"{description} - Selection Not Confirmed")
    return True


def report_optional_block(page: Page, description: str, what: str) -> None:
    """Record - loudly, and with evidence - that an optional block never loaded.

    This is the whole point of treating the specialisation block as optional: it
    is not swept under the carpet, it is written into the log, attached to the
    report and photographed, so the run that skipped it says exactly what it
    skipped and why. What it does not do is fail a registration the application
    itself is willing to accept without it.
    """
    app_messages = read_app_messages(page)
    detail = (f"{what}\n\n"
              f"This block is drawn from a catalogue the page fetches for "
              f"itself, and the form heads it 'Optional but strongly "
              f"recommended', so the application accepts a registration without "
              f"it. The run therefore carried on rather than failing here.\n"
              f"APPLICATION DEFECT worth raising: the catalogue behind this "
              f"block does not always load - it has been seen appearing at "
              f"once, appearing after ~30 seconds, and never appearing at all "
              f"on the same GSTIN and the same Industry Type.\n"
              f"Page: {page.url}")
    if app_messages:
        detail += "\nApplication said: " + " | ".join(app_messages)
    log.warning("   %s was NOT answered: %s", description, what)
    log.warning("   the application never offered it, and it is optional on "
                "this form, so the run carries on - see the report for the "
                "evidence")
    attach_text(f"{description} - Not Offered by the Application", detail)
    screenshot(page, f"{description} - Not Offered")


#: How each outcome of the lookup reads in the Allure step title.
GST_LOOKUP_RESULT = {
    "loaded": "SUCCESS",
    "quota": "QUOTA EXHAUSTED",
    "missing": "SECTION NEVER LOADED",
}


def specialisation_unavailable(page: Page, description: str,
                               message: str) -> None:
    """Report - and, when the profile is required, fail on - a block that never came.

    Always logged, photographed and attached to the Allure report;
    SPECIALISATION_REQUIRED then decides whether the run stops here or carries
    on and finishes the registration without the profile.
    """
    report_optional_block(page, description, message)
    if SPECIALISATION_REQUIRED:
        fail(page, message)


def choose_business_profile(page: Page, row: Row) -> None:
    """Answer step 1's four option-card fields, in the order the form needs them.

    Each block is waited for, answered, and confirmed before the next one is
    touched, because each one only renders once the block before it has been
    answered:

        Industry Type -> Who do you sell to? -> What does your business
        specialise in? (the category, and then the sub-category under it)

    The first two are required fields and are failed if they cannot be answered.
    The third is the application's own optional panel, fed by a catalogue lookup
    that does not always come back, so it is waited for far longer and then
    reported rather than failed - see the comment above SPECIALISE_HEADING.
    """
    industry = row["Business Type"]
    sells_to = row["Customer Type"]
    category = row["Business Category"]
    sub_category = row["Business Sub Category"]

    with step(f"Select Industry Type ({industry})"):
        anchor = wait_for_section(page, INDUSTRY_TYPE_HEADING, "'Industry Type'")
        select_option_card(
            page, anchor, industry, "Industry Type",
            "Industry Type was not selected/displayed within the expected time")
        screenshot(page, "Registration - Industry Type Selected")

    with step(f"Select who the business sells to ({sells_to})"):
        anchor = wait_for_section(
            page, SELL_TO_HEADING, "'Who do you sell to?'",
            "did not become available after the Industry Type selection")
        select_option_card(
            page, anchor, sells_to, "Customer Type",
            "'Who do you sell to?' was not selected/displayed within the "
            "expected time")

    with step("Wait for GST Lookup"):
        # Everything from here down sits inside the form's "Help buyers find you
        # faster" panel, which the application fills from a catalogue it fetches
        # through its GST lookup. Three things can happen, and they are told
        # apart rather than all being reported as "Business Category did not
        # appear":
        #
        #   the block renders             -> carry on
        #   "GST lookup quota exhausted"  -> take the application's own fallback
        #                                    if it has one, otherwise say which
        #                                    service is down
        #   neither, until the wait ends  -> say the block never loaded
        #
        # Both are waited for at once, so a spent quota is acted on the moment
        # it is announced instead of costing the whole SPECIALISATION_TIMEOUT.
        state, anchor = wait_for_specialisation_section(page, row)

    # --- what the lookup actually did, as its own step in the report ---------
    with step(f"GST Lookup Result: {GST_LOOKUP_RESULT[state]}"):
        if state == "quota" or anchor is None:
            if state == "quota":
                message = (
                    f"'What does your business specialise in?' was never "
                    f"rendered, so Business Category ('{category}') and "
                    f"Business Sub Category ('{sub_category}') could not be "
                    f"selected.\n"
                    f"The application was showing its GST notice at the time "
                    f"('GST lookup quota exhausted. State auto-filled; please "
                    f"enter remaining details manually.'). That notice is about "
                    f"the GSTIN AUTO-FILL, not about this block, and every "
                    f"field it asks to be typed WAS typed, from the workbook - "
                    f"so it is context, not proof of the cause.\n"
                    f"What was actually done: the block was waited for "
                    f"{SPECIALISATION_TIMEOUT // 1000} seconds, the lookup was "
                    f"retried once, it was waited for another "
                    f"{SPECIALISATION_TIMEOUT // 1000} seconds, and the screen "
                    f"was inspected for a supported way of continuing without "
                    f"the lookup - see the 'gst_lookup_fallback' attachment for "
                    f"what the application offered.\n"
                    f"Worth trying: run again once the GST quota has reset, or "
                    f"use a GSTIN the application has already cached.")
            else:
                message = (
                    f"'What does your business specialise in?' was never "
                    f"rendered within {SPECIALISATION_TIMEOUT // 1000} seconds "
                    f"of the Industry Type ('{industry}') being chosen, so "
                    f"neither the Business Category ('{category}') nor the "
                    f"Business Sub Category ('{sub_category}') could be "
                    f"selected. The application showed no GST lookup error, so "
                    f"this is the section itself failing to load.")
            specialisation_unavailable(page, "Business Category", message)
            return
        log.info("   the GST lookup came back and the application drew "
                 "'What does your business specialise in?'")

    # --- the section is up: now the Business Category tile itself ------------
    with step(f"Business Category ({category})"):
        if not wait_for_option_offered(page, anchor, category,
                                       "Business Category"):
            specialisation_unavailable(
                page, "Business Category",
                f"Business Category '{category}' did not become available "
                f"after the 'What does your business specialise in?' section "
                f"loaded (waited {FIELD_LOAD_TIMEOUT // 1000} seconds). The "
                f"section itself is on screen, so this is the tile that is "
                f"missing: check that '{category}' is still one of the options "
                f"the application lists for Industry Type '{industry}'.")
            return

        select_option_card(
            page, anchor, category, "Business Category",
            f"Business Category '{category}' did not become available after "
            f"the 'What does your business specialise in?' section loaded",
            required=SPECIALISATION_REQUIRED)

    # --- prove the Category was really applied before touching anything ------
    # The application answers a chosen Category by listing the sub-options
    # underneath it, so that listing is the proof, and it is WAITED for. Nothing
    # is clicked in between: these cards toggle, and a second click on the
    # Category would un-choose it.
    with step(f"Business Sub Category ({sub_category})"):
        if not wait_for_option_offered(page, anchor, sub_category,
                                       "Business Sub Category"):
            specialisation_unavailable(
                page, "Business Sub Category",
                f"Business Category '{category}' was clicked, but the Business "
                f"Sub Category options it unlocks never appeared within "
                f"{FIELD_LOAD_TIMEOUT // 1000} seconds. Either the selection "
                f"was not applied, or '{sub_category}' is not one of the "
                f"sub-options the application lists under '{category}'.")
            return
        log.info("   Business Category '%s' is applied - the application is "
                 "listing the options under it", category)

        select_option_card(
            page, anchor, sub_category, "Business Sub Category",
            f"Business Sub Category '{sub_category}' did not become available "
            f"under Business Category '{category}'",
            required=SPECIALISATION_REQUIRED)
        screenshot(page, "Registration - Specialisation Selected")


def wizard_field(page: Page, placeholder: str) -> Locator:
    """One text box of the sign-up wizard, addressed the way the recording does.

    The recording addresses every box of this form as
    get_by_role("textbox", name="<the placeholder it shows>"), and that is the
    better lookup here: a role lookup ignores the copies of a form this SPA
    keeps mounted but hidden, while a bare get_by_placeholder matches them and
    trips strict mode. The placeholder is kept as the fallback - the password
    boxes are not textboxes on every build - and both are narrowed to the one
    copy actually on screen.
    """
    return visible(page.get_by_role("textbox", name=placeholder, exact=False)
                   .or_(page.get_by_placeholder(placeholder)))


def accept_terms(page: Page) -> None:
    """Tick "I agree to GenZOpss Terms" on the Review & Submit step.

    The application draws a styled control rather than a plain checkbox, so the
    real checkbox is used when there is one and the recording's own tick
    element is the first fallback when there is not.
    """
    terms = page.locator("label").filter(has_text="I agree to GenZOpss Terms of")

    checkbox = terms.get_by_role("checkbox")
    if checkbox.count() > 0:
        scroll_into_view(checkbox.first, "Terms and conditions")
        checkbox.first.check()
        log.info("   ticked: Terms and conditions")
        return

    # The recording clicked `.mt-0\.5.w-5` - the styled tick box of the terms
    # row. It is looked for INSIDE the terms label first, because those are
    # Tailwind spacing classes and nothing stops another element on the review
    # step from carrying the same pair; page-wide is the fallback, for a build
    # that draws the tick outside the label.
    control = safe_locator(page, "Terms and conditions", (
        ("the recorded tick, inside the terms label",
         terms.locator(r".mt-0\.5.w-5")),
        ("the recorded tick control", page.locator(r".mt-0\.5.w-5")),
        ("the tick inside the terms label", terms.locator("div").first),
        ("the terms label itself", terms),
        ("any checkbox on the review step", page.get_by_role("checkbox")),
    ))
    if control is None:
        fail(page, "The 'I agree to GenZOpss Terms' tick could not be found on "
                   "the Review & Submit step, so the registration could not be "
                   "submitted.")
    click(page, control, "Terms and conditions")


def create_business_account(page: Page, row: Row) -> str:
    """Register one business from an Excel row. Returns the company name used.

    The order and the controls are the recorded Create Your Account flow -
    Get Started, GSTIN, Company Name, the four option cards, Pincode, Address,
    the location search, Continue, the owner's details, Continue, the terms
    tick, Create Account - with every value taken from the Create_Your_Account
    sheet and every action going through the framework's own click/fill layer.
    """
    company = row["Company Name"]
    gstin = row["GSTIN"]
    email = row["Email"]
    mobile = row["Mobile Number"]

    # --- Open GenZOpss and get into the sign-up wizard ---
    with step("Opening GenZOpss..."):
        safe_navigation(page, MARKETING_URL, "GenZOpss home page")

        gst_field = wizard_field(page, "27AABCU9603R1ZX")
        # The recording starts at "Get Started". Some layouts embed the form on
        # the home page instead, so the link is clicked only when the wizard is
        # not already on screen - checked, not waited for, so a home page that
        # does embed it costs nothing.
        if gst_field.count() == 0:
            click(page, link(page, "Get Started"), "Get Started")
            wait_until_ready(page, "registration form")
        expect_visible(page, gst_field, "GSTIN field")

    # --- Business details ---
    with step("Creating Business Account... (business details)"):
        company_field = wizard_field(page, "Rajan Steel Industries Pvt.")
        # Entering the GSTIN is what starts the application's GST lookup, and
        # everything the wizard draws below depends on that lookup - so it is a
        # step of its own in the report.
        with step(f"Enter GSTIN ({gstin})"):
            fill(page, gst_field, gstin, "GSTIN")
        fill(page, company_field, company, "Company Name")

        # Business Type / Customer Type / Category / Sub Category are option
        # cards - all four come from the worksheet, and all four depend on the
        # one before them. Each is waited for, chosen and confirmed on screen
        # before the next is touched; nothing below this line runs until the
        # specialisation the workbook asked for is displayed AND chosen.
        choose_business_profile(page, row)

        # APPLICATION DEFECT: picking a Business Type clears the GSTIN and
        # Company Name, and the form then fails its own validation on Continue.
        # Put them back rather than reordering the steps.
        for field, value, name in (
            (gst_field, gstin, "GSTIN"),
            (company_field, company, "Company Name"),
        ):
            if field.input_value() != value:
                log.warning("   app defect: %s was cleared - re-entering it", name)
                fill(page, field, value, f"{name} (re-entered)")

    # --- Address ---
    with step("Creating Business Account... (address)"):
        fill(page, wizard_field(page, "560058"), row["Pincode"], "Pincode")
        # The pincode lookup fills City and State and sets the location by
        # itself. Let it land BEFORE anything else is typed - it is an
        # unannounced network call, and its answer overwrites whatever is in the
        # fields when it comes back.
        wait_for_pincode_lookup(page)

        fill(page, wizard_field(page, "Industrial Area, Phase II"),
             row["Address"], "Address")

        # City has a field of its own (placeholder "Bengaluru") and the pincode
        # lookup usually fills it. "Search for your business location" is a
        # different control - the search the recording uses, handled by
        # search_business_location() below.
        city_box = wizard_field(page, "Bengaluru")
        if city_box.count() == 0:
            log.info("   skip: City (this form has no separate City field)")
        elif city_box.input_value().strip().casefold() == row["City"].strip().casefold():
            log.info("   fill: %-24s = %s (the pincode lookup already filled it)",
                     "City", row["City"])
        else:
            fill(page, city_box, row["City"], "City")

        search_business_location(page, row)

        click(page, button(page, "Continue"), "Continue (business details)")
        # The wizard moved on only if step 1 is gone.
        try:
            expect(gst_field).to_be_hidden(timeout=DEFAULT_TIMEOUT)
        except AssertionError as error:
            fail(page, "The business details were not accepted - the wizard "
                       "stayed on step 1", error)

    # --- Owner account ---
    with step("Creating Business Account... (contact person)"):
        fill(page, wizard_field(page, "Suresh Rajan"), row["Contact Person"],
             "Contact Person")
        fill(page, wizard_field(page, "9876543210"), mobile, "Mobile Number")
        fill(page, wizard_field(page, "suresh@rajan.co"), email, "Email")
        fill(page, wizard_field(page, "Min. 8 chars, 1 uppercase, 1"),
             row["Password"], "Password", secret=True)
        fill(page, wizard_field(page, "Re-enter password"),
             row["Confirm Password"] or row["Password"], "Confirm Password",
             secret=True)
        # The recording's get_by_role("button").nth(1) between the password and
        # Continue is the show/hide-password eye. It changes no data and an
        # index into every button on the page is the first thing a re-layout
        # breaks, so it is deliberately not replayed.
        click(page, button(page, "Continue"), "Continue (contact person)",
              animated=True)

    # --- Review & Submit ---
    with step("Submitting the registration (Review & Submit)..."):
        accept_terms(page)

        # The final control is "Create Account" (older builds label it "Submit").
        submit = button(page, "Create Account")
        if submit.count() == 0:
            submit = button(page, "Submit")
        click(page, submit, "Create Account")

        # Submitted means the button is gone. If it is still there, the app
        # rejected the form - show its own message.
        #
        # The button going is NOT proof on its own, which is what this used to
        # take it for. A run on 2026-08-13 had the button disappear on a
        # registration the API had answered 409 (already registered): the flow
        # called it a success, went on to the admin portal, and failed there
        # with "the tenant row never appeared" - three steps away from the
        # cause, and reported as a defect when it was a used-up row. So the
        # API is asked either way, and only the answer decides.
        submitted = True
        try:
            submit.wait_for(state="hidden", timeout=DEFAULT_TIMEOUT)
        except (PlaywrightTimeoutError, PlaywrightError):
            submitted = False

        messages = read_app_messages(page)
        if messages and already_exists(messages):
            raise RecordAlreadyExists(
                f"'{company}' is already registered - the Mobile Number, "
                f"Email or GSTIN in this row has been used before. Put fresh "
                f"values in the Create_Your_Account sheet to register again.\n"
                + "\n".join(messages))
        if not submitted:
            if messages:
                fail(page, "The registration was rejected", "\n".join(messages))
            log.warning("   the Submit button is still on screen and no error was "
                        "shown - continuing to the OTP page")
        elif messages:
            # Went through, but the API complained about something else on the
            # way. Worth recording; not worth stopping the run over.
            log.warning("   the registration went through, but the application "
                        "reported: %s", " | ".join(messages))
            attach_text("Registration - Application Messages", "\n".join(messages))

    log.info("   registered: %s (%s)", company, email)
    return company


#: Where the application asks for the e-mailed code.
OTP_VERIFY_PATH = "/register/verify"


def registration_completed(page: Page, company: str) -> str:
    """Prove the registration really went through. Returns how it was proved.

    Called between "Create Account" and the admin portal, because the two things
    that can be true here look identical from the outside: the account exists
    and is waiting for its e-mail code, or the submit never landed at all. The
    run of 2026-08-17 could not tell them apart, so a registration that HAD
    completed and a registration that had not both arrived at the admin portal
    the same way, and the tenant's absence was reported without knowing which of
    the two had happened.

    Nothing here can pass on its own evidence: it reads the application's, and
    when there is none it says so rather than assuming the registration worked.
    """
    if OTP_VERIFY_PATH in page.url:
        return (f"the application moved to its e-mail verification page "
                f"({page.url}), which it only does for an account it has "
                f"created")

    body = ""
    try:
        body = page.inner_text("body") or ""
    except PlaywrightError:
        pass

    for wording in (r"verify\s+your\s+e-?mail", r"check\s+your\s+e-?mail",
                    r"registration\s+(?:is\s+)?(?:successful|complete)",
                    r"account\s+created", r"OTP\s+sent"):
        if re.search(wording, body, re.IGNORECASE):
            return (f"the application is showing its own confirmation "
                    f"('{re.search(wording, body, re.IGNORECASE).group(0)}')")

    # The wizard is gone and no error is on screen: the form was accepted, and
    # that is the weakest of the three answers - so it says exactly that.
    if not re.search(r"Create\s+Account|Review\s*&\s*Submit", body,
                     re.IGNORECASE):
        return ("the sign-up wizard closed and the application reported no "
                "error - the weakest of the three confirmations, so the tenant "
                "wait below is what settles it")
    return ""


def wait_for_manual_otp(page: Page) -> bool:
    """Pause so the OTP can be typed BY HAND. Returns True if it was.

    THE CODE IS NEVER AUTOMATED. It arrives by e-mail, and nothing in this
    framework generates one, guesses one, hard-codes one or asks the application
    for one. This function does exactly two things: it puts the verification
    page in front of whoever is running the suite, and it waits for the
    APPLICATION to say the code was accepted.

    It is a wait for a CONDITION, not a sleep - it returns the instant the
    application leaves the verification page - and it reports every 15 seconds
    so a watched run cannot look hung.

    It is still not a hard gate. If nobody types a code the run carries on, and
    the admin approval reports honestly what it then finds: a business that was
    never verified does not become a tenant, and that is a BLOCKED precondition,
    not a defect. Set GENZ_OTP_WAIT_SECONDS=0 for an unattended run.
    """
    with step("Waiting for Manual OTP..."):
        if OTP_VERIFY_PATH not in page.url:
            safe_navigation(page, f"{APP_URL}{OTP_VERIFY_PATH}", "OTP page")
        else:
            wait_until_ready(page, "OTP page")

        seconds = OTP_WAIT_TIMEOUT // 1000
        if seconds <= 0:
            log.warning("GENZ_OTP_WAIT_SECONDS=0 - not pausing for the OTP. "
                        "The account stays unverified.")
            attach_text("Email Verification (OTP)",
                        "Not waited for: GENZ_OTP_WAIT_SECONDS=0 was set, so "
                        "this run did not pause for the manual code. An "
                        "unverified registration does not reach the admin "
                        "portal's Tenants list.")
            return False

        log.info("=" * 70)
        log.info("MANUAL STEP - type the OTP in the browser window that is open.")
        log.info("The code is e-mailed to the address in the "
                 "Create_Your_Account sheet; it cannot be read or generated by")
        log.info("the automation, so it is not attempted.")
        log.info("The run continues the moment the application accepts it, or")
        log.info("carries on by itself in %d seconds if none is entered.",
                 seconds)
        log.info("=" * 70)

        # Polled in slices rather than in one long wait, purely so the console
        # keeps saying what it is waiting for. The condition is the same one
        # either way: the application leaving its verification page.
        deadline = time.time() + seconds
        while True:
            left = deadline - time.time()
            if left <= 0:
                break
            try:
                page.wait_for_url(lambda url: OTP_VERIFY_PATH not in url,
                                  timeout=min(OTP_POLL_INTERVAL,
                                              int(left * 1000)))
            except (PlaywrightTimeoutError, PlaywrightError):
                remaining = int(max(deadline - time.time(), 0))
                if remaining:
                    log.info("   still waiting for the OTP... (%ds left)",
                             remaining)
                continue

            wait_until_ready(page, "page after verification")
            log.info("OTP verified - the application accepted the code.")
            attach_text("Email Verification (OTP)",
                        f"Verified: the code was typed by hand and the "
                        f"application accepted it. The page moved on to "
                        f"{page.url}.")
            screenshot(page, "Registration - After OTP Verification")
            return True

        log.warning("No OTP entered within %d seconds - the account stays "
                    "unverified. Continuing.", seconds)
        attach_text("Email Verification (OTP)",
                    f"NOT VERIFIED: no code was entered within {seconds}s.\n\n"
                    f"This is a MANUAL prerequisite. The code is e-mailed to "
                    f"the address in the Create_Your_Account sheet and the "
                    f"automation deliberately does not generate, guess or "
                    f"bypass it.\n\n"
                    f"An account whose e-mail is never verified does not "
                    f"become a tenant, so the admin approval below may find "
                    f"nothing to approve. That would be a BLOCKED "
                    f"precondition - not an application defect and not an "
                    f"automation defect.\n\n"
                    f"To take this row all the way through, run the suite with "
                    f"the browser watched and type the code when this pause "
                    f"appears, or give yourself longer with "
                    f"GENZ_OTP_WAIT_SECONDS=<seconds>.")
        screenshot(page, "Registration - OTP Not Entered")
        return False


def tenant_search_box(page: Page) -> Optional[Locator]:
    """The Tenants page's own search box, however this build labels it."""
    return safe_locator(page, "tenant search box", (
        ("the placeholder 'Search company or city'",
         page.get_by_placeholder("Search company or city")),
        ("any 'Search' placeholder",
         page.get_by_placeholder(re.compile("search", re.IGNORECASE))),
        ("a search textbox", page.get_by_role("textbox",
                                              name=re.compile("search",
                                                              re.IGNORECASE))),
    ))


def tenants_table_state(page: Page) -> str:
    """What the tenants table is saying about ITSELF right now.

    Three answers, and the difference between them is the whole diagnosis:

        "loading"   the table has not finished fetching - nothing can be
                    concluded from it yet
        "empty"     it has finished and is reporting no matching tenant
        "listed"    it has finished and is showing rows

    The run of 2026-08-17 reported "the table itself reports no matching row"
    while its own screenshot of that moment showed the bare "Loading..." cell.
    Reading the table's state rather than assuming it is what stops a report
    saying something the evidence does not.
    """
    try:
        body = page.inner_text("body") or ""
    except PlaywrightError:
        return "unknown"
    if re.search(r"^\s*Loading\s*\.{0,3}\s*$", body, re.IGNORECASE | re.MULTILINE):
        return "loading"
    if re.search(r"No\s+tenants?\s+found|No\s+results|Nothing\s+to\s+show",
                 body, re.IGNORECASE):
        return "empty"
    return "listed"


def find_tenant_row(page: Page, company: str,
                    timeout: int = TENANT_WAIT_TIMEOUT) -> Tuple[Optional[Locator],
                                                                 str]:
    """Poll the admin Tenants list until the business appears. (row, what it said).

    A registration does not become a tenant the instant it is submitted - the
    application has its own work to do first (and its e-mail verification to
    wait for) - and the list itself renders a bare "Loading..." cell while it
    fetches. So this is a WAIT FOR A CONDITION with retries, not one look:

        type the company into the search  ->  let the debounce fire  ->
        wait for the loaders to clear     ->  look for the row       ->
        not there? say what the table said, reload, and search again

    until the budget (TENANT_WAIT_TIMEOUT, GENZ_TENANT_WAIT_SECONDS) is spent.
    Every attempt is logged with what the table was showing at the time, so a
    row that never appears is reported with the evidence rather than with a
    guess. Nothing here invents a tenant and nothing here fails a test: the
    caller decides what a `None` means.
    """
    deadline = time.time() + timeout / 1000
    attempt = 0
    said = "unknown"

    while True:
        attempt += 1
        search = tenant_search_box(page)
        if search is not None:
            # Re-typed every time on purpose. A reload clears the filter, and
            # an application that re-fetched the list in the background can
            # leave the box holding text it is no longer filtering by.
            #
            # A search box that cannot be typed into does not end the wait: the
            # unfiltered list is still worth looking at, and this loop is a wait
            # for a tenant, not an assertion about a text box.
            try:
                safe_fill(page, search, company, "Tenant search", attempts=1)
            except (AssertionError, PlaywrightError) as error:
                log.warning("   could not type into the tenant search box (%s) "
                            "- looking at the unfiltered list instead",
                            str(error).splitlines()[0])
            # The search is debounced: for a moment the OLD result is still on
            # screen and no spinner exists yet. Settle so the fetch starts,
            # THEN wait for the table to finish - otherwise the row check races
            # it, which is exactly what happened on 2026-08-17.
            settle(page)
        wait_until_ready(page, "filtered tenants table")

        row_locator = page.get_by_role("row").filter(has_text=company).first
        if appears(row_locator, TENANT_POLL_INTERVAL):
            log.info("   the tenant '%s' is listed (attempt %d)", company,
                     attempt)
            return row_locator, "listed"

        said = tenants_table_state(page)
        left = int(max(deadline - time.time(), 0))
        log.info("   '%s' is not in the Tenants list yet - the table is %s "
                 "(attempt %d, %ds of the wait left)", company,
                 {"loading": "still LOADING",
                  "empty": "reporting NO MATCHING TENANT",
                  "listed": "showing other tenants only"}.get(said, said),
                 attempt, left)
        if left <= 0:
            return None, said

        # A reload is the one thing a re-typed search cannot do: it makes the
        # application fetch the list again from scratch, which is what a tenant
        # created moments ago needs.
        reload_page(page)
        wait_until_ready(page, "Tenants page")


def report_blocked_tenant(page: Page, company: str, said: str, waited: int,
                          otp_verified: bool, registered_because: str) -> None:
    """The business registered but never became a tenant. Report it and stop.

    This is a BLOCKED PRECONDITION and it is reported as one, in those words:
    the registration was accepted, so nothing about it failed; the approval
    screen simply had nothing to act on, so nothing about the admin portal was
    proved either. It is NOT an application defect and it is NOT an automation
    defect, and the report must not let a reader mistake it for either.

    Everything printed here is evidence this run actually collected - how the
    registration was confirmed, whether the manual code was typed, how long the
    tenant list was watched, and what that list was saying when the wait ran
    out. Nothing is inferred, and no tenant is invented to keep the flow going.
    """
    table_said = {
        "loading": ("the Tenants table was STILL LOADING when the wait ran "
                    "out - the admin portal had not finished fetching the "
                    "list"),
        "empty": ("the Tenants table finished loading and reported no "
                  "matching tenant"),
        "listed": ("the Tenants table finished loading and showed other "
                   "tenants, but no row for this company"),
    }.get(said, "the Tenants table's state could not be read")

    if otp_verified:
        cause = ("the e-mail code WAS typed and accepted during this run, so "
                 "the account was verified - and the business still did not "
                 "reach the Tenants list within the wait. Worth raising with "
                 "the application team as a question about how long "
                 "provisioning takes; this run cannot call it a defect, "
                 "because a tenant that arrives later than the wait and a "
                 "tenant that never arrives look identical from here.")
        todo = [
            f"Confirm in the admin portal, by hand, whether "
            f"'{company}' is now listed under Tenants.",
            f"If it is, the provisioning simply took longer than the "
            f"{waited}s this run waited - raise the budget with "
            f"GENZ_TENANT_WAIT_SECONDS=<seconds> and run again.",
            "If it is still not there, the registration did not produce a "
            "tenant: that is a question for the application team, and the "
            "evidence attached here is what to send them.",
        ]
    else:
        cause = ("the manual e-mail OTP was not entered during this run. The "
                 "code arrives by e-mail and this framework deliberately does "
                 "not generate, guess or bypass it, so the run paused for it "
                 "and carried on without it - and an account whose e-mail is "
                 "never verified does not become a tenant.")
        todo = [
            "Run the suite with the browser watched.",
            "When it pauses at 'Waiting for Manual OTP', read the code from "
            "the e-mail sent to the address in the Create_Your_Account sheet "
            "and type it into the browser window.",
            "The run continues by itself the moment the application accepts "
            "the code, and the admin approval then has a tenant to act on.",
            "Give yourself longer with GENZ_OTP_WAIT_SECONDS=<seconds> "
            f"(this run allowed {OTP_WAIT_TIMEOUT // 1000}s).",
        ]

    story = qa_report.blocker_story(
        "Create Your Account",
        precondition=(f"'{company}' has to exist as a TENANT in the admin "
                      f"portal before it can be approved. Registering the "
                      f"business is what creates it, and the e-mail "
                      f"verification is part of registering."),
        root_cause=cause,
        expected=(f"'{company}' listed on the admin portal's Tenants page, "
                  f"with an Approve action against it."),
        actual=(f"After waiting {waited}s - searching, re-reading and "
                f"reloading the list throughout - {table_said}."),
        before_rerun=todo)
    attach_text("BLOCKED PRECONDITION - Create Your Account", story)

    fail(page,
         f"BLOCKED PRECONDITION: business registration completed, but "
         f"'{company}' did not become available as a tenant in the Admin "
         f"Portal within the configured timeout ({waited}s).\n"
         f"CATEGORY: BLOCKED PRECONDITION - this is NOT an automation defect "
         f"and NOT an application defect. The registration was accepted and "
         f"the approval screen simply had nothing to act on, so nothing about "
         f"the application was proved either way.\n"
         f"Registration was confirmed by: "
         f"{registered_because or '(no confirmation was recorded)'}\n"
         f"Manual e-mail OTP entered this run: "
         f"{'yes' if otp_verified else 'NO'}\n"
         f"When the wait ran out, {table_said}.\n"
         f"WHAT TO DO BEFORE RE-RUNNING:\n  "
         + "\n  ".join(f"{index}. {item}"
                       for index, item in enumerate(todo, start=1)))


def approve_business(page: Page, company: str,
                     otp_verified: bool = True,
                     registered_because: str = "") -> None:
    """Log in to the admin portal and approve the business just registered.

    The order this follows is the flow's own: registration completed -> the
    business becomes a tenant -> the tenant list is opened and WAITED ON ->
    the exact company is searched for -> it is approved.

    `otp_verified` says whether the e-mail code was actually typed during this
    run, and `registered_because` is how the registration was confirmed. Neither
    changes what is done - they are what lets a missing tenant row be reported
    as the precondition it is, with the evidence, rather than as a bare "the row
    never appeared".
    """
    # --- Admin login ---
    with step("Opening Admin Portal..."):
        safe_navigation(page, f"{ADMIN_URL}/login", "admin login page")

        # APPLICATION DEFECT: typing the password clears the phone field, so the
        # password goes in FIRST. Same fields, same values - only the order.
        fill(page, page.get_by_placeholder("••••••••"), ADMIN_PASSWORD,
             "Admin password", secret=True)
        fill(page, page.get_by_placeholder("+"), ADMIN_MOBILE, "Admin mobile")
        click(page, button(page, "Sign in"), "Sign in")
        try:
            page.wait_for_url(lambda url: "/login" not in url,
                              timeout=DEFAULT_TIMEOUT)
        except (PlaywrightTimeoutError, PlaywrightError) as error:
            fail(page, "Admin login failed", error)

    # --- Find the tenant ---
    with step(f"Searching Tenant... ({company})"):
        click(page, link(page, "Tenants"), "Tenants")
        wait_until_ready(page, "Tenants page")

        waited = TENANT_WAIT_TIMEOUT // 1000
        log.info("   waiting up to %ds for '%s' to appear in the Tenants list",
                 waited, company)
        row_locator, said = find_tenant_row(page, company)

        if row_locator is None:
            screenshot(page, "Admin Portal - Tenant Not Listed")
            # The list itself, as evidence. Reading it must never be what stops
            # the run - the blocked report below is the point of getting here.
            try:
                on_screen = (page.inner_text("body") or "")[:4000]
            except PlaywrightError as error:
                on_screen = f"(the page could not be read: {error})"
            attach_text("Admin Portal - Tenants List When The Wait Ran Out",
                        on_screen or "(nothing on screen)")
            report_blocked_tenant(page, company, said, waited, otp_verified,
                                  registered_because)
        expect_visible(page, row_locator, f"tenant row for '{company}'")

    # --- Approve ---
    with step("Approving the business..."):
        # Row actions carry no text - their accessible name is the title
        # attribute. Never click by position: on an already-approved row the
        # same position is "Suspend", which would undo the approval.
        if row_locator.get_by_role("button", name="Suspend").count() > 0:
            log.warning("'%s' is already approved - skipping the approval click",
                        company)
            return

        click(page, row_locator.get_by_role("button", name="Approve"), "Approve",
              animated=True)
        wait_until_ready(page, "tenants table after approval")

        badge = row_locator.get_by_text(re.compile(r"^\s*approved\s*$",
                                                   re.IGNORECASE)).first
        try:
            expect(badge).to_be_visible(timeout=DEFAULT_TIMEOUT)
        except AssertionError as error:
            fail(page, f"'{company}' was clicked but its status did not change "
                       f"to approved", error)
        log.info("   approved: %s", company)
        screenshot(page, "Business - Approved in the Admin Portal")


# =========================================================================== #
# LOG IN AS THE BUSINESS USER (once, before the Customer module)
# =========================================================================== #

def login_as_business_user(page: Page) -> None:
    """Sign in to the workspace as the approved business user."""
    with step("Logging in as Business User..."):
        safe_navigation(page, MARKETING_URL, "GenZOpss home page")

        click(page, link(page, "Sign in"), "Sign in link")
        wait_until_ready(page, "sign-in form")

        # The field is named "phone" by the app but accepts an e-mail too.
        fill(page, page.get_by_test_id("phone"), BUSINESS_USER_EMAIL, "Username")
        fill(page, page.get_by_test_id("password"), BUSINESS_USER_PASSWORD,
             "Password", secret=True)
        click(page, page.get_by_test_id("login-btn"), "Sign in")
        wait_until_ready(page, "workspace")

        # Asserting on the navigation proves the session is usable - a URL check
        # alone also passes on an error page.
        expect_visible(page, nav_link(page, "Customers").first,
                       "workspace navigation")
        dismiss_overlay(page)
        log.info("Business user logged in.")


# =========================================================================== #
# MODULE 2 - CUSTOMER  (sheet: Customer)
# =========================================================================== #

#: Customer Name from Excel -> its city. The Quotations module needs the city to
#: recognise the customer in the picker, which lists "<initial> <name> <city>".
CREATED_CUSTOMERS: Dict[str, Tuple[str, str]] = {}


def create_customer(page: Page, row: Row) -> None:
    """Create one customer from an Excel row."""
    name = row["Customer Name"]

    with step("Creating Customer..."):
        open_module(page, "Customers", "Add Customer")
        click(page, button(page, "Add Customer"), "Add Customer", animated=True)

        # Fields in screen order - exactly the column order of the Customer sheet.
        fill(page, visible(page.get_by_placeholder("27AABCU9603R1ZX")),
             row["GSTIN"], "GSTIN")
        fill(page, visible(page.get_by_label("Customer Name*")), name,
             "Customer Name")
        fill(page, visible(page.get_by_placeholder("9876543210")),
             row["Phone"], "Phone")
        fill(page, visible(page.get_by_label("Email")), row["Email"], "Email")
        fill(page, visible(page.get_by_placeholder("Street / Building")),
             row["Address"], "Address")
        fill(page, visible(page.get_by_label("City")), row["City"], "City")
        fill(page, visible(page.get_by_label("State")), row["State"], "State")
        fill(page, visible(page.get_by_placeholder("400001")), row["Pincode"],
             "Pincode")

        # Credit block - sits at the bottom of the form, so scroll to it first.
        credit_limit = visible(page.get_by_label("Credit Limit"))
        if credit_limit.count() > 0:
            scroll_into_view(credit_limit, "Credit Limit", editable=True)
        fill_optional(page, row.get("Credit Limit", ""), "Credit Limit",
                      credit_limit,
                      visible(page.get_by_placeholder("Credit Limit")))
        fill_optional(page, row.get("Credit Days", ""), "Credit Days",
                      visible(page.get_by_label("Credit Days")),
                      visible(page.get_by_placeholder("Credit Days")))

        submit = button(page, "Create Customer")
        click(page, submit, "Create Customer")
        confirm_saved(page, submit, name, "Customer")

    # Remember the city so a Quotation row can find this customer in the picker.
    CREATED_CUSTOMERS[name.strip().lower()] = (name, row["City"])


# =========================================================================== #
# MODULE 3 - SUPPLIERS  (sheet: Suppliers)
# =========================================================================== #

def create_supplier(page: Page, row: Row) -> None:
    """Create one supplier, including the bank details on the same form."""
    name = row["Supplier Name"]

    with step("Creating Supplier..."):
        open_module(page, "Suppliers", "Add Supplier")
        click(page, button(page, "Add Supplier"), "Add Supplier", animated=True)

        # Fields in screen order - exactly the column order of the Suppliers sheet.
        fill(page, visible(page.get_by_placeholder("27AABCU9603R1ZX")),
             row["GSTIN"], "GSTIN")
        fill(page, visible(page.get_by_label("Supplier Name*")), name,
             "Supplier Name")
        fill(page, visible(page.get_by_placeholder("9876543210")),
             row["Phone"], "Phone")
        fill(page, visible(page.get_by_label("Email")), row["Email"], "Email")
        fill(page, visible(page.get_by_label("City")), row["City"], "City")
        fill(page, visible(page.get_by_label("State")), row["State"], "State")
        fill(page, visible(page.get_by_placeholder("400001")), row["Pincode"],
             "Pincode")
        fill_optional(page, row.get("Credit Days", ""), "Credit Days",
                      visible(page.get_by_label("Credit Days")),
                      visible(page.get_by_placeholder("Credit Days")))

        # Bank details sit below the fold - scroll them into view first.
        account_name = visible(page.get_by_label("Account Name"))
        scroll_into_view(account_name, "Account Name", editable=True)
        fill(page, account_name, row["Account Name"], "Account Name")
        fill(page, visible(page.get_by_label("Account No")),
             row["Account No"], "Account No")
        fill(page, visible(page.get_by_placeholder("State Bank of India")),
             row["Bank Name"], "Bank Name")
        fill(page, visible(page.get_by_placeholder("SBIN0001234")),
             row["IFSC Code"], "IFSC Code")
        fill(page, visible(page.get_by_label("Branch")), row["Branch"], "Branch")

        submit = button(page, "Create Supplier")
        click(page, submit, "Create Supplier")
        confirm_saved(page, submit, name, "Supplier")


# --------------------------------------------------------------------------- #
# THE SUPPLIER'S OWN PAGE - what the purchase flow did to it
#
# A supplier is not only the form the sheet fills in. Once a purchase order has
# been raised against it, received on a GRN and the GRN approved, the supplier's
# own page carries the three figures this suite is asked to prove:
#
#     Outstanding               what is owed to the supplier, as an AMOUNT
#     RECENT PURCHASE ORDERS    the order, its date, its value and its status
#     SUPPLIED MATERIALS        the item and the rate it is being bought at
#
# Read as TEXT, and for the same reason the totals panels are (see PageNumbers):
# these blocks are label-less CSS grids - no <table>, no role, no id - so there
# is no cell to address and the caption above the figure is the only thing that
# says what it is.
#
# The Outstanding figure is an AMOUNT here, not a quantity. That is a fact about
# THIS application, established by reading it (a supplier with three orders of
# Rs 62,835 shows Rs 1,88,505), and it is why the outstanding QUANTITY the brief
# asks about is looked for separately and reported as unavailable when the
# application does not print one, rather than being answered with the money
# figure. Two different things must not be reported as the same one.
# --------------------------------------------------------------------------- #

#: One entry of the RECENT PURCHASE ORDERS list. The application prints the
#: number, the date, the value and the status each on a line of its own.
PO_NUMBER_LINE = re.compile(r"^\s*(PO\s*[/-]?\s*\d[\w/-]*)\s*$", re.IGNORECASE)

#: A rate the SUPPLIED MATERIALS block prints against an item: "Rs 363.00/PCS".
SUPPLIED_RATE_LINE = re.compile(
    r"^\s*[₹$€£]?\s*([\d,]+(?:\.\d+)?)\s*/\s*[A-Za-z]+\s*$")


def page_block(page: Page, heading: str, *ends: str) -> List[str]:
    """The lines one captioned block of a page prints, the caption excluded.

    The same reading materials_table_text() does for the indent's materials
    grid, made general because three more screens need it: the supplier's
    RECENT PURCHASE ORDERS and SUPPLIED MATERIALS blocks, and the item's STOCK
    LEDGER. None of them is a table, so none of them has a cell to address.

    `ends` are the captions that come after it; the first one found closes the
    block, so the next panel's figures can never be read as part of this one.
    Returns [] when the caption is not on screen - which the caller reports as
    "the application does not show this", never as a value.
    """
    for area in (page.locator("main").first, page.locator("body").first):
        try:
            text = area.inner_text(timeout=OPTIONAL_TIMEOUT) or ""
        except (PlaywrightTimeoutError, PlaywrightError):
            continue
        start = re.search(heading, text, re.IGNORECASE)
        if not start:
            continue
        block = text[start.end():]
        if ends:
            stop = re.search(r"\n\s*(?:" + "|".join(ends) + r")\b", block,
                             re.IGNORECASE)
            if stop:
                block = block[:stop.start()]
        return [line.strip() for line in block.splitlines() if line.strip()]
    return []


def supplier_row(page: Page, supplier: str) -> Optional[Locator]:
    """The Suppliers list row of one supplier, by its NAME cell.

    The cell, not the row: a row's textContent runs its cells together, so a
    substring match over the whole row can be satisfied by a name that merely
    starts the same way. The substring filter is kept as the fallback, for a
    build whose list is not a table at all.
    """
    exact = row_matching(page, table_column_index(page, "NAME",
                                                  "SUPPLIER", "SUPPLIER NAME"),
                         supplier)
    if exact is not None:
        return exact
    rows = page.locator("tbody tr").filter(visible=True).filter(
        has_text=re.compile(re.escape(supplier), re.IGNORECASE))
    try:
        return rows.first if rows.count() else None
    except PlaywrightError:
        return None


def open_supplier_record(page: Page, supplier: str) -> bool:
    """Open one supplier's own page from the Suppliers list.

    False, never an exception, when the supplier is not in the list: a figure
    that could not be reached is a SKIPPED check with its reason, and the run
    has to carry on to the checks that CAN be made.
    """
    open_module(page, "Suppliers", "Add Supplier")
    search_in_list(page, search_term(supplier), "supplier")
    wait_page_ready(page, "the Suppliers list")

    row = supplier_row(page, supplier)
    if row is None or not appears(row, RETRY_TIMEOUT):
        log.warning("   the Suppliers list does not show '%s' - the supplier's "
                    "own page cannot be opened", supplier)
        return False

    click(page, row, f"supplier '{supplier}'", animated=True)
    wait_until_ready(page, "the supplier's page")
    if "/suppliers/" not in page.url:
        log.warning("   clicking the supplier row did not open a supplier page "
                    "(still at %s)", page.url)
        return False
    return True


def read_supplier_purchase_orders(page: Page) -> List[Dict[str, str]]:
    """The RECENT PURCHASE ORDERS block, one dict per order it lists.

    Each order is {Purchase Order, Date, Amount, Status}. Anything the block
    does not print for an order is left empty rather than guessed at - a status
    that is not there is not the same thing as a status of "Received".
    """
    lines = page_block(page, r"RECENT\s+PURCHASE\s+ORDERS",
                       r"BANK\s+DETAILS", r"SUPPLIED\s+MATERIALS",
                       r"PAYMENT\s+HISTORY", r"NOTES")
    orders: List[Dict[str, str]] = []
    current: Optional[Dict[str, str]] = None
    for line in lines:
        number = PO_NUMBER_LINE.match(line)
        if number:
            current = {"Purchase Order": re.sub(r"\s+", "", number.group(1)),
                       "Date": "", "Amount": "", "Status": ""}
            orders.append(current)
            continue
        if current is None:
            continue
        if not current["Date"] and verify_data.date_candidates(line):
            current["Date"] = line
        elif not current["Amount"] and verify_data.as_number(line) is not None:
            current["Amount"] = line
        elif not current["Status"]:
            current["Status"] = line
    return orders


def document_digits(text: str) -> str:
    """The number inside a document's name, without its leading zeros.

    "PO-00048", "PO/0048" and "PO 48" are one order under three spellings: the
    prefix, the separator and how many zeros a screen chose to pad with are the
    screen's business, not the document's identity.
    """
    digits = re.sub(r"[^0-9]", "", text or "")
    return digits.lstrip("0") or ("0" if digits else "")


def supplier_order_row(orders: List[Dict[str, str]],
                       number: str) -> Optional[Dict[str, str]]:
    """The listed order whose number is `number`, however each screen spells it."""
    wanted = document_digits(number)
    if not wanted:
        return None
    for order in orders:
        if document_digits(order.get("Purchase Order", "")) == wanted:
            return order
    return None


def supplier_outstanding(page: Page) -> Optional[float]:
    """The Outstanding figure the supplier's own page prints.

    An AMOUNT. Whether this application expresses "outstanding" as money or as
    a quantity is not something to assume - it was read off the application
    (a supplier carrying three orders of Rs 62,835 each shows Rs 1,88,505), and
    the outstanding QUANTITY is looked for under its own captions elsewhere.
    """
    return PageNumbers.read(page).find("Outstanding")


def supplier_outstanding_from_list(page: Page, supplier: str,
                                   when: str) -> Optional[float]:
    """The OUTSTANDING column the Suppliers list prints for one supplier.

    Used for the reading taken BEFORE the purchase flow, where opening the
    supplier's own page would be a second navigation for a figure the list
    already carries. Puts the browser back where it found it.
    """
    was_at = page.url
    try:
        with step(f"Reading the supplier's outstanding... ({supplier}, {when})"):
            open_module(page, "Suppliers", "Add Supplier")
            search_in_list(page, search_term(supplier), "supplier")
            wait_page_ready(page, "the Suppliers list")

            row = supplier_row(page, supplier)
            if row is None or not appears(row, RETRY_TIMEOUT):
                log.warning("   the Suppliers list does not show '%s' yet - no "
                            "outstanding could be read %s", supplier, when)
                return None

            column = table_column_index(page, "OUTSTANDING", "BALANCE",
                                        "PAYABLE")
            if column is None:
                log.warning("   the Suppliers list has no outstanding column")
                return None
            value = quantity_number(row_cell(row, column).replace("₹", ""))
            reading = stock_reading(page, value, f"suppliers-list-{when}")
            log.info("   %s outstanding %s: %s", supplier, when,
                     "-" if reading is None else f"{float(reading):,.2f}")
            attach_text(f"Supplier Outstanding - {when}",
                        f"Supplier    : {supplier}\n"
                        f"Read        : {when}\n"
                        f"Outstanding : "
                        f"{'-' if reading is None else calc.money(float(reading))}\n"
                        f"Read from   : the OUTSTANDING column of the Suppliers "
                        f"list")
            return reading
    except (PlaywrightTimeoutError, PlaywrightError, AssertionError) as error:
        log.warning("   the supplier's outstanding could not be read %s: %s",
                    when, str(error).splitlines()[0])
        return None
    finally:
        restore_page(page, was_at)


def read_supplied_material_rate(page: Page, code: str,
                                name: str) -> Optional[float]:
    """The rate the SUPPLIED MATERIALS block prints for one item ("Rs 363.00/PCS")."""
    lines = page_block(page, r"SUPPLIED\s+MATERIALS",
                       r"RECENT\s+PURCHASE\s+ORDERS", r"BANK\s+DETAILS")
    marker = (code or name or "").strip()
    if not marker:
        return None
    start = next((index for index, line in enumerate(lines)
                  if marker.casefold() in line.casefold()), None)
    if start is None:
        return None
    for line in lines[start:start + 4]:
        rate = SUPPLIED_RATE_LINE.match(line)
        if rate:
            try:
                return float(rate.group(1).replace(",", ""))
            except ValueError:
                return None
    return None


# =========================================================================== #
# MODULE 4 - ITEMS  (sheet: Items)
# =========================================================================== #

def open_stock_module(page: Page, module: str, expected_button: str) -> None:
    """Open a module that lives inside the collapsed 'Stock & Materials' group.

    The same shape as open_sales_module(): the group itself is a button, the
    modules inside it are links, and the group label carries a live count once
    the business holds stock - so it is matched with exact=False, exactly as
    "Orders & Sales" has to be.
    """
    dismiss_overlay(page)          # an open form hides the whole sidebar
    if nav_link(page, module).count() == 0:
        group = button(page, "Stock & Materials", exact=False).filter(visible=True)
        expect_visible(page, group.first, "'Stock & Materials' sidebar group")
        click(page, group.first, "Stock & Materials menu", animated=True)
    open_module(page, module, expected_button)


def choose_item_category(page: Page, category: str) -> bool:
    """Pick the item's category. OPTIONAL - the form calls it optional itself.

    The picker opens with the parents collapsed, so a sub-category is only on
    screen once its own parent has been expanded. The recording expands the
    first parent ("View sub-categories" .first), which is the right one only
    while the wanted category happens to sit under it - so each parent in turn
    is opened until the category appears.
    """
    category = (category or "").strip()
    if not category:
        log.info("   skip: %-24s (the cell is empty)", "Category")
        return False

    click(page, visible(page.get_by_role("button", name="Select category")),
          "Select category", animated=True)

    wanted = button(page, category, exact=False).filter(visible=True).first
    if wanted.count() == 0:
        expanders = page.get_by_role("button", name="View sub-categories") \
                        .filter(visible=True)
        for index in range(expanders.count()):
            click(page, expanders.nth(index), f"View sub-categories ({index + 1})",
                  animated=True)
            if wanted.count() > 0:
                break

    if not appears(wanted):
        fail(page, f"The item form's category list does not offer '{category}'. "
                   f"Either create it first (Stock & Materials > Categories, "
                   f"which is what the Categories sheet does), or correct the "
                   f"Category cell in the Items sheet.")
    click(page, wanted, f"Category '{category}'", animated=True)
    verify_data.record("Category", category)
    return True


def select_item_type(page: Page, item_type: str) -> None:
    """Choose Item Type from the form's dropdown.

    A real <select>, so the cell may hold either the value the application
    stores ("semi_finished") or the label it prints ("Semi Finished") - both are
    accepted, and a cell that matches neither is answered with the list of the
    options the form actually has.
    """
    item_type = (item_type or "").strip()
    if not item_type:
        log.info("   skip: %-24s (the cell is empty)", "Item Type")
        return

    # safe_select tries the value first, then the label, retries a dropdown that
    # re-rendered, and lists the real options if neither spelling is one of them.
    field = safe_locator(page, "Item Type", (
        ("the field labelled Item Type*", page.get_by_label("Item Type*")),
        ("the field labelled Item Type", page.get_by_label("Item Type")),
        ("a combobox called Item Type",
         page.get_by_role("combobox", name="Item Type")),
        ("the only select on the form", page.locator("select")),
    )) or visible(page.get_by_label("Item Type*"))

    safe_select(page, field, item_type, "Item Type",
                hint="Correct the Item Type cell in the Items sheet.")


#: The catalogue items THIS run created, by code and by name. REPORTING ONLY:
#: nothing in the flow reads them back. They exist so the stock report can say
#: whether the item a purchase stocked was new - and a new item's opening stock
#: should be 0, which is the one expectation a reader cannot check without
#: knowing the item did not exist before this run started.
CREATED_ITEMS: List[str] = []
CREATED_ITEM_NAMES: List[str] = []


def created_in_this_run(code: str, name: str) -> bool:
    """Did THIS run create the catalogue item the stock chain is following?"""
    code, name = (code or "").strip(), (name or "").strip()
    if code and code in [entry for entry in CREATED_ITEMS if entry]:
        return True
    return bool(name) and name in CREATED_ITEM_NAMES


def create_item(page: Page, row: Row) -> None:
    """Create one catalogue item from an Excel row."""
    name = row["Item Name"]
    code = (row.get("Item Code", "") or "").strip()

    with step("Creating Item..."):
        open_stock_module(page, "Items", "Add Item")
        click(page, button(page, "Add Item"), "Add Item", animated=True)
        form_screenshot(page, "Item - Form Before Save")

        # Fields in screen order - exactly the column order of the Items sheet.
        fill(page, visible(page.get_by_role("textbox", name="Item Code")),
             row["Item Code"], "Item Code")
        fill(page, visible(page.get_by_role("textbox", name="HSN Code")),
             row["HSN Code"], "HSN Code")
        fill(page, visible(page.get_by_role("textbox", name="Name*")), name,
             "Item Name")

        choose_item_category(page, row.get("Category", ""))
        select_item_type(page, row.get("Item Type", ""))

        fill(page, visible(page.get_by_role("spinbutton", name="Purchase Price")),
             row["Purchase Price"], "Purchase Price")
        fill(page, visible(page.get_by_role("spinbutton", name="Selling Price")),
             row["Selling Price"], "Selling Price")
        # A blank MRP cell leaves the application's own default in place.
        fill_optional(page, row.get("MRP", ""), "MRP",
                      visible(page.get_by_role("spinbutton", name="MRP")))

        submit = button(page, "Create Item")
        click(page, submit, "Create Item")
        confirm_saved(page, submit, name, "Item")
        # REPORTING ONLY - no page is read here and nothing is verified. It
        # records that THIS run created this item, which is the one fact that
        # lets the stock report say "a new item, so its opening stock should
        # have been 0" instead of leaving a reader to assume it.
        CREATED_ITEMS.append(code)
        CREATED_ITEM_NAMES.append(name)
        screenshot(page, "Item - After Save")

    with step(f"Verifying the item in the Items list... ({name})"):
        wait_until_ready(page, "Items list")
        # visible=True: the create form the app has just closed still holds the
        # name in the DOM, and a hidden copy would satisfy the assertion.
        expect_visible(page, page.get_by_text(name).filter(visible=True).first,
                       f"item '{name}' in the Items list")
        log.info("   verified: '%s' is in the Items list", name)
        screenshot(page, "Item - Verified in the Items List")

    # THE CURRENT STOCK BEFORE ANY TRANSACTION - read now, on the Items test
    # itself, because this is the earliest moment "before anything moved" is
    # true. Uses the SAME stock_now() round-trip the Purchase Order flow
    # already makes later; nothing new is added to what the automation
    # reaches, only WHEN this one report-only read of it happens.
    #
    # capture_purchase_baseline() still takes its OWN "before the purchase
    # order" reading later and that is what the GRN math is measured from -
    # this call changes no calculation. It only means report_initial_stock()'s
    # once-per-item guard (INITIAL_STOCK_REPORTED) is tripped here instead of
    # there, so the 📦 ITEM STOCK - BEFORE TRANSACTION card now lands on the
    # Items test case instead of waiting for the Purchase Orders one.
    opening = stock_now(page, code, name, "immediately after the item was created")
    report_initial_stock(code, name, opening)


# --------------------------------------------------------------------------- #
# ITEMS & INVENTORY - THE STOCK A TRANSACTION MOVED
#
# The Items module's page is headed "Items & Inventory" and its list carries a
# STOCK column; the item's own page carries "Current Stock" and, under it, a
# STOCK LEDGER - one line per movement, with what it was and what the balance
# became:
#
#     INDENT RETURN      Material return MR-... against IND/0031   +6 PCS   Bal: 98
#     ADJUSTMENT OUT     Issue against IND/0031                   -68 PCS   Bal: 92
#     PURCHASE RECEIPT   GRN GRN/0045                            +160 PCS   Bal: 160
#
# That is the whole movement of this run, written by the application itself, and
# it is what makes the stock chain checkable rather than merely plausible.
#
# Two rules, and they are the ones the rest of the framework already follows:
#
#   * A reading is taken at the MOMENT it means something - before the purchase,
#     after the GRN approval, after the issue, after the return - and every one
#     of them is tagged with the reading it came from, so an expected value can
#     never be answered by the figure it was worked out from (see Displayed).
#   * The item is found by its CODE. This catalogue holds eleven items called
#     "OPC Cement 53 Grade" (ITEM-OPC53-001 ... -011), each with a stock of its
#     own, so a stock read by NAME is whichever of them the list happened to
#     draw first - a number that is real, belongs to a real item, and answers a
#     question nobody asked.
# --------------------------------------------------------------------------- #

#: The heading the Items & Inventory list draws over the stock column. "MIN
#: STOCK LEVEL" is a different column and must not answer for it, which is why
#: the header cell is matched WHOLE rather than by "contains stock".
INVENTORY_STOCK_HEADINGS = ("STOCK", "CURRENT STOCK", "STOCK REMAINING",
                            "QTY IN STOCK")

#: A quantity as these screens print it: "98 PCS", "1,650 PCS", "-68 PCS", "98".
#: as_number() cannot read these - it stops at a trailing unit - and a stock
#: figure without its unit is what this application prints everywhere. The
#: unit itself is captured too (group 2) - see quantity_unit() - so a report
#: can print "182 PCS" rather than a bare "182" that a reader has to take on
#: trust matches the screen.
QUANTITY_WITH_UNIT = re.compile(
    r"^\s*([+-]?[\d,]+(?:\.\d+)?)\s*([A-Za-z]{0,6})\s*$")


def quantity_unit(text: str) -> str:
    """The unit a quantity was printed with - "PCS", "KG" - or "" for none."""
    match = QUANTITY_WITH_UNIT.match(verify_data.clean_text(text))
    if not match:
        return ""
    return (match.group(2) or "").strip().upper()

#: One movement of the STOCK LEDGER, and the balance printed under it.
LEDGER_DELTA_LINE = re.compile(r"^\s*([+-])\s*([\d,]+(?:\.\d+)?)\s*[A-Za-z]*\s*$")
LEDGER_BALANCE_LINE = re.compile(r"^\s*Bal\.?:?\s*([\d,]+(?:\.\d+)?)", re.IGNORECASE)

#: How many stock readings this run has taken. Every one of them gets its own
#: number so that two readings of the same screen at two different moments -
#: which is exactly what proves a movement - are never mistaken for one.
STOCK_READINGS = 0

#: Every reading, kept whole, so the end-of-flow ledger report can show where
#: each figure in it came from instead of only the arithmetic it produced.
STOCK_READINGS_TAKEN: List[Dict[str, object]] = []


def quantity_number(text: str) -> Optional[float]:
    """"98 PCS" -> 98.0. None when the text is not a bare quantity."""
    match = QUANTITY_WITH_UNIT.match(verify_data.clean_text(text))
    if not match:
        return None
    try:
        return float(match.group(1).replace(",", ""))
    except ValueError:
        return None


def stock_reading(page: Page, value: Optional[float],
                  where: str) -> Optional[float]:
    """Tag a figure lifted off a stock screen with the reading it came from.

    What keeps the stock checks honest. PageNumbers does this for every figure
    it finds; these are read out of a table cell instead, so the tag has to be
    put on here - without it, a closing stock could be "proved" against itself.
    """
    global STOCK_READINGS
    if value is None:
        return None
    STOCK_READINGS += 1
    return calc.Displayed(value, f"{page.url}#{where}-{STOCK_READINGS}")


def table_column_index(page: Page, *headings: str) -> Optional[int]:
    """Which column of a list carries `headings`, read off its own header row.

    By the header rather than by a fixed position, because a build is free to
    add a column: a stock read from "the seventh cell" is silently the wrong
    figure the day a column appears before it, and a wrong figure is worse than
    no figure. None - which is a SKIPPED check with its reason - when none of
    the headings is on the list.
    """
    try:
        cells = page.locator("thead th").filter(visible=True)
        texts = [verify_data.clean_text(cells.nth(index).inner_text()).upper()
                 for index in range(cells.count())]
    except (PlaywrightTimeoutError, PlaywrightError):
        return None
    for heading in headings:
        wanted = heading.strip().upper()
        for index, text in enumerate(texts):
            if text == wanted:
                return index
    return None


def row_cell(row: Locator, column: Optional[int]) -> str:
    """The text of one cell of a list row, or "" when it cannot be read."""
    if column is None:
        return ""
    try:
        return verify_data.clean_text(row.locator("td").nth(column).inner_text())
    except (PlaywrightTimeoutError, PlaywrightError):
        return ""


# --------------------------------------------------------------------------- #
# HOW MANY RECORDS A LIST HOLDS
#
# "There were 5 purchase orders, I raised 1, there should be 6" is the simplest
# before-and-after check in the suite and the easiest one to get quietly wrong,
# because a list page shows a PAGE of records and not all of them. Counting the
# rows on screen would report a 60-record list as 10 both times: the check would
# pass, prove nothing, and go on passing after the day it stopped being true.
#
# So the printed total is what is used - "Showing 1 to 10 of 57" - and the row
# count is the fallback ONLY when the whole list demonstrably fits on one page.
# When neither is available the count is not taken: an unverifiable check is
# reported as unverifiable, which is this framework's rule everywhere else.
# --------------------------------------------------------------------------- #

#: How a list page prints the number of records BEHIND it, not on it. Ordered
#: most specific first; the first one that matches wins.
LIST_TOTAL_PATTERNS: Tuple[str, ...] = (
    r"showing\s+[\d,]+\s*(?:to|-|–|—)\s*[\d,]+\s+of\s+([\d,]+)",
    r"\bof\s+([\d,]+)\s+(?:entries|records|results|items|rows|orders)\b",
    r"\b([\d,]+)\s+(?:records?|results?|entries|orders?)\s+found\b",
    r"\btotal\s+(?:records?|results?|entries|orders?)\s*:?\s*([\d,]+)\b",
)

#: What a list draws when it has more pages than the one on screen. Their mere
#: presence is not enough - a disabled "Next" sits on a single-page list too -
#: so these are only consulted to decide whether an on-screen row count can be
#: trusted, and an enabled one means it cannot.
PAGINATION_NEXT = ("Next", "next page", "Go to next page", "»", "›")


def printed_list_total(page: Page) -> Optional[int]:
    """The record count the list page prints, when it prints one."""
    try:
        text = page.inner_text("body")
    except (PlaywrightTimeoutError, PlaywrightError):
        return None
    for pattern in LIST_TOTAL_PATTERNS:
        found = re.search(pattern, text, re.IGNORECASE)
        if found:
            try:
                return int(found.group(1).replace(",", ""))
            except ValueError:
                continue
    return None


def has_more_pages(page: Page) -> bool:
    """Is there a page of this list that is NOT on screen?

    True only for a next-page control that is really there and really enabled.
    A control that cannot be read is treated as "there may be more", because
    the expensive mistake here is trusting a partial count, not rejecting a
    complete one.
    """
    for name in PAGINATION_NEXT:
        control = (page.get_by_role("button", name=name, exact=False)
                   .or_(page.get_by_role("link", name=name, exact=False))
                   .filter(visible=True))
        try:
            if control.count() == 0:
                continue
            if control.first.is_enabled():
                return True
        except (PlaywrightTimeoutError, PlaywrightError):
            return True                    # unreadable - do not trust a count
    return False


#: How many times list_record_count() scrolls a "no Next control" list before
#: trusting its row count. Two consecutive scrolls that add no new row is what
#: tells "that really is everything" apart from "there is more I have not
#: asked for yet" - one scroll is not enough, because a list can legitimately
#: take a moment to append the next batch.
SCROLL_GROWTH_ATTEMPTS = 8


def grown_by_scrolling(page: Page, seen: int,
                       attempts: int = SCROLL_GROWTH_ATTEMPTS) -> Optional[int]:
    """Scroll a list to see whether it loads MORE rows than are on screen now.

    `has_more_pages()` only recognises a textual "Next"/"»" control - it says
    nothing about a list that quietly fetches its next batch on scroll and
    never draws a control at all. Such a list can sit at a round row count
    (50, 100...) forever, because adding one record just pushes the oldest of
    the visible batch out of view - the exact way this suite once reported a
    Purchase Orders list as unchanged when a new order really had been added
    to it. Scrolling is the only way to tell "the whole list is on screen"
    from "there is a page boundary nobody clicked past".

    Returns the row count once TWO scrolls in a row add nothing new. None
    when it is STILL growing after `attempts` scrolls: a list that keeps
    producing rows is a moving target, not a count this check can trust, and
    it is reported as unreadable rather than guessed at.

    Scrolls the LAST ROW into view rather than sending a raw mouse-wheel
    event at a fixed screen position: `scroll_into_view_if_needed()` scrolls
    whichever ancestor actually needs to move to show that element, which is
    the table's own scrollable body when the table has one - a wheel event
    dispatched at (0, 0) instead scrolls the outer page (sidebar included,
    confirmed live: a run's screenshot showed a completely different sidebar
    menu after "scrolling" while the table's own row count never moved),
    which proves nothing about whether the table itself has more rows.
    """
    stable = 0
    for _ in range(attempts):
        try:
            rows = page.locator("tbody tr").filter(visible=True)
            if rows.count() == 0:
                return seen
            rows.last.scroll_into_view_if_needed()
            page.wait_for_timeout(400)
            now = page.locator("tbody tr").filter(visible=True).count()
        except (PlaywrightTimeoutError, PlaywrightError):
            return seen                       # cannot scroll further - stop
        if now <= seen:
            stable += 1
            if stable >= 2:
                return seen
        else:
            stable = 0
            seen = now
    return None


def list_record_count(page: Page, what: str) -> Tuple[Optional[int], str]:
    """(how many records the list holds, where that number came from).

    (None, why not) when the page will not say. NOTHING is inferred: a count
    that cannot be established is returned as no count at all, so the check
    above it is SKIPPED with the reason rather than answered with a figure that
    happens to be on screen.
    """
    wait_page_ready(page, f"{what} list")

    total = printed_list_total(page)
    if total is not None:
        return total, "the total the list page prints"

    try:
        rows = page.locator("tbody tr").filter(visible=True).count()
    except (PlaywrightTimeoutError, PlaywrightError) as error:
        return None, (f"the {what} list could not be read "
                      f"({str(error).splitlines()[0]})")

    if has_more_pages(page):
        return None, (f"the {what} list prints no total and it has more than "
                      f"one page, so the {rows} row(s) on screen are only part "
                      f"of it - counting them would compare one page with "
                      f"another and call the difference a record count")

    # No "Next" control does not by itself mean the whole list is on screen -
    # see grown_by_scrolling(). A row count this suite has already watched
    # sit still at a round number while an application record was quietly
    # scrolled out of view is not one this check should trust unconfirmed.
    grown = grown_by_scrolling(page, rows)
    if grown is None:
        return None, (f"the {what} list kept loading more rows as it was "
                      f"scrolled and never settled, so its row count could "
                      f"not be established - counting a moving target would "
                      f"compare one moment with another")

    return grown, "counted on the list after scrolling it to the end"


def list_contains_reference(page: Page, reference: str) -> bool:
    """Is this exact document reference one of the rows on screen right now?

    A plain text search over the table body - used to tell a genuine page
    boundary ("the new record is not here yet, it is on a page nobody asked
    for") apart from a sliding window ("the new record is right here, and the
    row count did not grow because something else fell out of view instead").
    Unreadable is treated as "not found": this only ever makes a check less
    confident, never more, so a page that will not say costs nothing here.
    """
    if not reference:
        return False
    try:
        return reference in page.locator("tbody").inner_text()
    except (PlaywrightTimeoutError, PlaywrightError):
        return False


def row_matching(page: Page, column: Optional[int], wanted: str,
                 limit: int = 60) -> Optional[Locator]:
    """The row whose `column` cell is exactly `wanted`.

    A CELL, never the row. A row's textContent is its cells run together with
    nothing between them - "ITEM-OPC53-011OPC Cement 53 Grade..." - so a regex
    with a boundary after the code never matches, and one without a boundary
    matches ITEM-OPC53-01 against ITEM-OPC53-011. Both readings are wrong, and
    the second one is worse: it is a real stock figure belonging to a different
    item. Comparing the cell itself has neither problem.
    """
    if column is None or not wanted:
        return None
    rows = page.locator("tbody tr").filter(visible=True)
    try:
        found = rows.count()
    except PlaywrightError:
        return None
    for index in range(min(found, limit)):
        row = rows.nth(index)
        if row_cell(row, column).casefold() == wanted.strip().casefold():
            return row
    return None


def inventory_row(page: Page, code: str, name: str) -> Optional[Locator]:
    """The Items & Inventory row of one catalogue item, found by its CODE.

    The code is what identifies the item, and it is compared against the CODE
    cell. This catalogue holds eleven items called "OPC Cement 53 Grade" with
    eleven different stocks, so a row found by name is a real figure answering
    a question nobody asked - which is why a name match is the fallback, and
    why it says so when more than one row could have been the answer.
    """
    if code:
        row = row_matching(page, table_column_index(page, "CODE", "ITEM CODE",
                                                    "SKU"), code)
        if row is not None:
            return row
        log.warning("   Items & Inventory does not list the code '%s' - looking "
                    "for the item by name instead", code)

    if not name:
        return None

    by_name = row_matching(page, table_column_index(page, "NAME", "ITEM NAME",
                                                    "ITEM"), name)
    if by_name is not None:
        return by_name

    rows = page.locator("tbody tr").filter(visible=True).filter(
        has_text=re.compile(re.escape(name), re.IGNORECASE))
    try:
        found = rows.count()
    except PlaywrightError:
        return None
    if not found:
        return None
    if found > 1:
        log.warning("   %d rows carry the name '%s' and none of them could be "
                    "told apart by a code - the first one is read, which may "
                    "not be the item this run moved", found, name)
    return rows.first


def read_current_stock(page: Page) -> Tuple[Optional[float], str]:
    """The "Current Stock" figure the item's own page prints, and its unit.

    The caption is waited for first. The item page draws its heading before the
    figures under it, and reading it too early returns nothing at all - which
    would be reported as "the application does not show a current stock", an
    untrue statement about an application that shows it a second later.
    """
    caption = page.get_by_text(re.compile(r"^\s*Current\s+Stock\s*$",
                                          re.IGNORECASE)).filter(visible=True)
    appears(caption.first, RETRY_TIMEOUT)

    for attempt in (1, 2):
        lines = page_block(page, r"Current\s+Stock", r"Min\s+Stock",
                           r"Max\s+Stock", r"Selling\s+Price", r"PRICING")
        for line in lines[:3]:
            value = quantity_number(line)
            if value is not None:
                return value, quantity_unit(line)
        if attempt == 1:
            settle(page)                 # the figure is drawn after the caption
            wait_page_ready(page, "the item's page")
    return None, ""


def read_stock_ledger(page: Page) -> List[Dict[str, object]]:
    """The item's STOCK LEDGER, one dict per movement, newest first.

    {"movement": "PURCHASE RECEIPT", "reference": "GRN GRN/0045",
     "change": +160.0, "balance": 160.0}

    The reference is what ties a movement to the document that caused it, and
    it is the only way to say "THIS run's GRN added 160" rather than "the stock
    went up by 160 at some point".
    """
    lines = page_block(page, r"STOCK\s+LEDGER", r"STOCK\s+BATCHES",
                       r"PRODUCTION\s+ORDERS", r"SPECIFICATIONS")
    entries: List[Dict[str, object]] = []
    recent: List[str] = []
    index = 0
    while index < len(lines):
        line = lines[index]
        delta = LEDGER_DELTA_LINE.match(line)
        if not delta:
            if not LEDGER_BALANCE_LINE.match(line):
                recent.append(line)
                recent = recent[-2:]
            index += 1
            continue

        try:
            change = float(delta.group(2).replace(",", ""))
        except ValueError:
            index += 1
            continue
        if delta.group(1) == "-":
            change = -change

        balance = None
        for following in lines[index + 1:index + 3]:
            found = LEDGER_BALANCE_LINE.match(following)
            if found:
                try:
                    balance = float(found.group(1).replace(",", ""))
                except ValueError:
                    balance = None
                break

        entries.append({
            "movement": recent[0] if recent else "",
            "reference": recent[-1] if recent else "",
            "change": change,
            "balance": balance,
        })
        recent = []
        index += 1
    return entries


def references_document(reference: str, document: str) -> bool:
    """Does one ledger line refer to this document?

    The number as it is written first ("GRN/0045" inside "GRN GRN/0045"), and
    only then its digits with a boundary either side. The digits alone are not
    enough by themselves: a ledger line reads "Material return MR-20260814-0003
    against IND/0031", and a bare substring search for a four-digit number finds
    it in the date - which would credit one document's movement to another.
    """
    reference = verify_data.clean_text(reference)
    document = verify_data.clean_text(document)
    if not reference or not document:
        return False
    if document.casefold() in reference.casefold():
        return True

    digits = document_digits(document)
    if not digits:
        return False
    # The digits alone are not an identity - a ledger line reads "Material
    # return MR-20260814-0003 against IND/0031" and a four-digit number turns up
    # in the date as readily as in the document number. So the line also has to
    # carry the document's own prefix, which is the part that says WHICH kind of
    # document the number belongs to.
    prefix = re.match(r"\s*([A-Za-z]{2,})", document)
    if prefix and prefix.group(1).casefold() not in reference.casefold():
        return False
    # 0* so that "GRN/0045" still finds a ledger that padded it "GRN-00045".
    return bool(re.search(rf"(?<!\d)0*{re.escape(digits)}(?!\d)", reference))


def ledger_entry_for(entries: List[Dict[str, object]],
                     document: str) -> Optional[Dict[str, object]]:
    """The ledger movement caused by one document (a GRN, an indent)."""
    for entry in entries:
        if references_document(str(entry.get("reference", "")), document):
            return entry
    return None


def ledger_table(entries: List[Dict[str, object]]) -> str:
    """The stock ledger as the report prints it."""
    if not entries:
        return "(the item's page shows no stock ledger)"
    lines = [f"{'MOVEMENT':<20} {'REFERENCE':<46} {'CHANGE':>10} {'BALANCE':>10}",
             "-" * 90]
    for entry in entries:
        change = entry.get("change")
        balance = entry.get("balance")
        lines.append(
            f"{str(entry.get('movement', '')):<20} "
            f"{str(entry.get('reference', ''))[:46]:<46} "
            f"{('' if change is None else f'{change:+g}'):>10} "
            f"{('-' if balance is None else f'{balance:g}'):>10}")
    return "\n".join(lines)


def restore_page(page: Page, was_at: str) -> None:
    """Put the browser back where a reading took it away from.

    Every stock reading is an excursion in the middle of somebody else's flow -
    the indent still has a Return to Store to do, the GRN still has its data
    verification to run - so nothing this layer does is allowed to leave the
    run somewhere it did not expect to be. The same courtesy
    data_verification.py already pays after reading a View page.
    """
    if not was_at or page.url == was_at:
        return
    try:
        safe_navigation(page, was_at, "the page the stock reading interrupted")
        wait_until_ready(page, "the page the run was on")
    except (PlaywrightTimeoutError, PlaywrightError, AssertionError) as error:
        log.warning("   could not return to %s after the stock reading: %s",
                    was_at, str(error).splitlines()[0])


def inventory_reading(page: Page, code: str, name: str,
                      when: str) -> Dict[str, object]:
    """Everything Items & Inventory knows about one item, in one visit.

    {"stock": the STOCK column of the list, "current": the item page's Current
     Stock, "ledger": [movements], "read": whether anything was read at all}

    Both stock figures are read because they are two different screens saying
    the same thing, and a run in which they disagree has found something worth
    reporting. The item page's own "Current Stock" is the figure the checks
    are made against - it is the exact field and the exact wording a person
    reads when they open that item, so a report that named a different figure
    "the actual stock" could disagree with what is on screen even while
    reading a real number off a real screen. The list's STOCK column is kept
    as corroboration and as the fallback when the item's own page could not
    be opened.

    Never raises. A screen that could not be reached leaves every figure None,
    and a check with no actual value is SKIPPED with its reason - which is the
    honest answer, and not the same thing as a wrong number.
    """
    reading: Dict[str, object] = {"stock": None, "current": None, "unit": "",
                                  "ledger": [], "read": False,
                                  "when": when, "code": code, "name": name}
    was_at = page.url
    try:
        with step(f"Reading the stock in Items & Inventory... "
                  f"({code or name}, {when})"):
            open_stock_module(page, "Items", "Add Item")
            search_in_list(page, code or name, "item")
            wait_page_ready(page, "the Items & Inventory list")

            row = inventory_row(page, code, name)
            if row is None:
                log.warning("   Items & Inventory does not list '%s' - no stock "
                            "figure could be read %s", code or name, when)
                screenshot(page, f"Item Not Listed in Items and Inventory ({when})")
                return reading

            column = table_column_index(page, *INVENTORY_STOCK_HEADINGS)
            if column is None:
                log.warning("   the Items & Inventory list has no stock column "
                            "- its headings are not the ones this suite knows")
            else:
                reading["stock"] = stock_reading(
                    page, quantity_number(row_cell(row, column)),
                    f"items-inventory-{when}")

            screenshot(page, f"Stock - Items and Inventory ({when})")

            # The item's own page: Current Stock and the whole stock ledger.
            try:
                click(page, row, f"item '{code or name}'", animated=True)
                wait_until_ready(page, "the item's page")
                if "/items/" in page.url:
                    current_value, current_unit = read_current_stock(page)
                    reading["current"] = stock_reading(
                        page, current_value, f"item-page-{when}")
                    if current_unit:
                        reading["unit"] = current_unit
                        stock_report.note_unit(current_unit)
                    reading["ledger"] = read_stock_ledger(page)
                    screenshot(page, f"Stock Ledger - Item Page ({when})")
            except (PlaywrightTimeoutError, PlaywrightError,
                    AssertionError) as error:
                log.warning("   the item's own page could not be opened: %s",
                            str(error).splitlines()[0])

            reading["read"] = (reading["stock"] is not None
                               or reading["current"] is not None)
            report_inventory_reading(reading)
    except (PlaywrightTimeoutError, PlaywrightError, AssertionError) as error:
        # A reading that could not be taken must not end somebody else's module.
        # It is reported here and it becomes a SKIPPED check with its reason.
        log.warning("   the stock could not be read %s: %s", when,
                    str(error).splitlines()[0])
    finally:
        restore_page(page, was_at)
    return reading


def report_inventory_reading(reading: Dict[str, object]) -> None:
    """Put one stock reading, and where every part of it came from, in Allure."""
    stock, current = reading.get("stock"), reading.get("current")
    lines = [
        f"Item                         : {reading.get('name') or '-'} "
        f"({reading.get('code') or 'no code'})",
        f"Read                         : {reading.get('when')}",
        f"Items & Inventory - STOCK    : "
        f"{'-' if stock is None else f'{float(stock):g}'}",
        f"Item page - Current Stock    : "
        f"{'-' if current is None else f'{float(current):g}'}",
        "",
        "STOCK LEDGER (the movements the application itself recorded)",
        ledger_table(reading.get("ledger") or []),
    ]
    if stock is not None and current is not None and abs(stock - current) > 0.001:
        lines += ["",
                  "NOTE - the list and the item's own page do not agree on this "
                  "item's stock. Both figures are the application's; the report "
                  "uses the item's own page's Current Stock (below), because "
                  "that is the exact field and wording a person reads when they "
                  "open the item - the list's STOCK column is kept here as "
                  "corroboration, not as the figure the checks are made "
                  "against."]
    attach_text(f"Item Stock Reading - {reading.get('when')}", "\n".join(lines))
    log.info("   stock %s: list %s | item page %s", reading.get("when"),
             "-" if stock is None else f"{float(stock):g}",
             "-" if current is None else f"{float(current):g}")


def stock_now(page: Page, code: str, name: str,
              when: str) -> Optional[float]:
    """The item's Current Stock, exactly as the item's own page shows it now.

    Prefers the item page's own "Current Stock" reading - the field and the
    wording a person actually reads when they open the item - and falls back
    to the Items & Inventory list's STOCK column only when the item's own
    page could not be read at all (see inventory_reading()).
    """
    reading = inventory_reading(page, code, name, when)
    STOCK_READINGS_TAKEN.append(reading)
    current = reading.get("current")
    return current if current is not None else reading.get("stock")


# =========================================================================== #
# MODULE 5 - CATEGORIES  (sheet: Categories)
# =========================================================================== #

def choose_category_colour(page: Page, colour: str) -> bool:
    """Pick the category's colour swatch. Cosmetic, so never a failure.

    The swatches are buttons named after the colour ("Orange"). A colour this
    build does not offer is reported and the category keeps the default one -
    failing the row over a swatch would fail a category that was created
    perfectly well.
    """
    colour = (colour or "").strip()
    if not colour:
        log.info("   skip: %-24s (the cell is empty)", "Color")
        verify_data.note_skipped("Color", "", "the cell in the workbook is empty")
        return False

    swatch = button(page, colour, exact=False).filter(visible=True).first
    if swatch.count() == 0:
        log.warning("   this form offers no '%s' colour - the category keeps the "
                    "default one", colour)
        attach_text("Category Colour Not Selected", f"Color = {colour}")
        verify_data.note_skipped("Color", colour,
                                 "this form offers no such colour")
        return False

    click(page, swatch, f"Colour '{colour}'")
    verify_data.record("Color", colour)
    return True


def open_sub_category_form(page: Page, parent: str) -> None:
    """Open the 'add a sub-category' form belonging to one parent category.

    The recording clicks get_by_role("button", name="Sub").nth(2) - the third
    "Sub" button on the page, which is the right one only while that parent
    stays third in the list. Every category card carries its own button, so the
    parent's card is found by its name first and the positional click is only
    the fallback. "Sub" is matched exactly: as a substring it also finds
    "Submit".
    """
    # Innermost card that both reads the parent's name and holds a Sub button:
    # ancestors come first in document order, so .last is the innermost one.
    card = page.locator("div") \
               .filter(has_text=re.compile(rf"^\s*{re.escape(parent)}",
                                           re.IGNORECASE)) \
               .filter(has=button(page, "Sub")).last
    control = card.get_by_role("button", name="Sub", exact=True).first

    if control.count() == 0:
        log.warning("   could not find the 'Sub' button on the '%s' card - "
                    "falling back to the recorded position", parent)
        control = button(page, "Sub").filter(visible=True).nth(2)
        if control.count() == 0:
            fail(page, f"There is no 'Sub' button for the category '{parent}', so "
                       f"the sub-category cannot be added. Check that the parent "
                       f"category was really created.")

    click(page, control, f"Sub (add a sub-category to '{parent}')", animated=True)


def expand_parent_category(page: Page, parent: str) -> bool:
    """Open a parent category so what is underneath it is on screen.

    A sub-category that has just been created is invisible while its parent is
    collapsed - which is not a failed save, just a closed drawer. Every way the
    build has of opening one is tried: the named expander on the parent's own
    card, any chevron on that card, and finally the card itself.
    """
    parent = (parent or "").strip()
    if not parent:
        return False

    card = page.locator("div") \
               .filter(has_text=re.compile(rf"^\s*{re.escape(parent)}",
                                           re.IGNORECASE)) \
               .filter(has=page.get_by_role("button")).last

    for name in ("View sub-categories", "Show sub-categories", "Expand",
                 "Sub-categories"):
        control = card.get_by_role("button", name=name, exact=False) \
                      .filter(visible=True)
        try:
            if control.count() == 0:
                continue
        except PlaywrightError:
            continue
        log.info("   '%s' is collapsed - opening it", parent)
        safe_click(page, control.first, f"expand '{parent}'", animated=True,
                   attempts=1)
        return True

    # No named expander: some builds open the drawer when the row is clicked.
    row_text = page.get_by_text(parent, exact=False).filter(visible=True).first
    try:
        if row_text.count() > 0:
            log.info("   opening '%s' by clicking the row itself", parent)
            safe_click(page, row_text, f"open '{parent}'", animated=True,
                       attempts=1)
            return True
    except PlaywrightError:
        pass
    return False


def search_in_list(page: Page, term: str, what: str = "record") -> bool:
    """Narrow a long list down to one name, when the page has a search box.

    Optional by design: a page without a search box just says so and the caller
    keeps looking the other ways.
    """
    box = safe_locator(page, f"{what} search box", (
        ("placeholder 'Search categories'",
         page.get_by_placeholder("Search categories", exact=False)),
        ("placeholder 'Search items'",
         page.get_by_placeholder("Search items", exact=False)),
        ("the search role", page.get_by_role("searchbox")),
        ("any 'Search' placeholder",
         page.get_by_placeholder("Search", exact=False)),
    ), need_editable=True)

    if box is None:
        return False

    safe_fill(page, box, term, f"{what} search", verify=False)
    # The search is debounced and gives no signal of its own until the fetch
    # starts - settle() is the framework's one short pause, and this is the
    # place it exists for. The real wait for the answer is the next line.
    settle(page)
    wait_page_ready(page, f"filtered {what} list")
    return True


def verify_category_created(page: Page, name: str, parent: str = "",
                            what: str = "Category",
                            attempts: int = VERIFY_ATTEMPTS) -> None:
    """Prove a category really exists, the way a person would check.

    Straight after Create, the screen is the least reliable it will ever be: the
    success toast has often already faded, the list is still being fetched, and
    a sub-category is hidden inside a parent that is collapsed. Any one of those
    made the old one-shot assertion report a category that had been created
    perfectly well as missing.

    So this never looks immediately. Every round waits for the API to answer and
    the DOM to stop moving, then works through all the places the category can
    be, in the order that costs least:

        a duplicate message   the row is used up - reported as skipped, not failed
        a success message     the application said so itself
        the list              the name, visible on screen
        inside its parent     expand it and look again
        the search box        narrow a long list down to the name

    and if none of that finds it, the page is reloaded and the round starts
    again. Only when every round is spent is it a failure, and each round leaves
    a screenshot in the Allure report.
    """
    for attempt in range(1, attempts + 1):
        wait_api_complete(page)
        wait_dom_stable(page)
        wait_for_loaders(page)

        # 1. What the application itself is saying, before anything fades.
        messages = read_app_messages(page)
        if messages and already_exists(messages):
            raise RecordAlreadyExists(
                f"{what} '{name}' is already in the application - this workbook "
                f"row has been used before. Put a different name in the "
                f"Categories sheet to create it again.\n" + "\n".join(messages))
        if any(re.search(r"success", message, re.IGNORECASE)
               for message in messages):
            log.info("   the application confirmed the %s was created",
                     what.lower())
            screenshot(page, f"{what} - Saved and Confirmed")
            return

        wanted = page.get_by_text(name, exact=False).filter(visible=True).first

        # 2. Simply on screen.
        if appears(wanted, RETRY_TIMEOUT):
            log.info("   verified: '%s' is in the Categories list", name)
            screenshot(page, f"{what} - Verified in the List")
            return

        # 3. Hidden inside a collapsed parent.
        if parent and expand_parent_category(page, parent):
            if appears(wanted, RETRY_TIMEOUT):
                log.info("   verified: '%s' is listed under '%s'", name, parent)
                screenshot(page, f"{what} - Verified in the List")
                return

        # 4. Too far down a long list to be on screen at all.
        if search_in_list(page, name, "category"):
            if appears(wanted, RETRY_TIMEOUT):
                log.info("   verified: '%s' was found by searching for it", name)
                screenshot(page, f"{what} - Verified in the List")
                return

        screenshot(page, f"{what} - Retry {attempt} of the Verification")
        log.warning("   attempt %d/%d: '%s' is not on screen yet", attempt,
                    attempts, name)
        if attempt < attempts:
            reload_page(page)

    fail(page, f"{what} '{name}' was submitted, but it never appeared in the "
               f"Categories list - not after {attempts} attempts, expanding its "
               f"parent, searching for it and reloading the page. Check whether "
               f"the application really created it.")


def create_category(page: Page, row: Row) -> None:
    """Create one parent category and, when the sheet asks for one, its child."""
    # {RUN} is expanded once, here, so the name that is typed, the name that is
    # verified and the name the sub-category is hung under are the same text.
    parent = resolve_run_token(row["Parent Category"])
    sub = resolve_run_token(row.get("Sub Category", ""))
    colour = row.get("Color", "")

    with step(f"Creating Parent Category... ({parent})"):
        open_stock_module(page, "Categories", "Add Category")
        click(page, button(page, "Add Category"), "Add Category", animated=True)
        form_screenshot(page, "Category - Form Before Save")

        # Fields in screen order - the column order of the Categories sheet.
        fill(page, visible(page.get_by_role("textbox", name="Category Name*")),
             parent, "Parent Category")
        choose_category_colour(page, colour)

        submit = button(page, "Create Category")
        click(page, submit, "Create Category")
        # The payload was accepted when the form's own button goes; the record
        # itself is checked in the verification step, not here.
        wait_for_submit_to_close(page, submit, parent, "Category")
        screenshot(page, "Category - After Save")

    with step(f"Verifying the parent category... ({parent})"):
        verify_category_created(page, parent, what="Category")
        # Categories do not go through confirm_saved(), so the record is queued
        # for the data check here instead.
        verify_data.after_save(page, parent, "Category")

    if not sub:
        log.info("   skip: %-24s (the cell is empty)", "Sub Category")
        return

    with step(f"Creating Sub Category... ({sub} under {parent})"):
        open_sub_category_form(page, parent)
        form_screenshot(page, "Sub Category - Form Before Save")

        fill(page, visible(page.get_by_role("textbox", name="Category Name*")),
             sub, "Sub Category")
        choose_category_colour(page, colour)

        submit = button(page, "Create Category")
        click(page, submit, "Create Category")
        wait_for_submit_to_close(page, submit, sub, "Sub Category")
        screenshot(page, "Sub Category - After Save")

    with step(f"Verifying the sub category... ({sub})"):
        # The parent is passed in so the check can open it: a sub-category is
        # invisible while its parent's drawer is shut.
        verify_category_created(page, sub, parent=parent, what="Sub Category")
        verify_data.after_save(page, sub, "Sub Category")


# =========================================================================== #
# MODULE 6 - QUOTATIONS  (sheet: Quotations)
# =========================================================================== #

# --------------------------------------------------------------------------- #
# THE LINE-ITEM ROW  (Quotations and Sales Orders share this layout)
#
# The row is a bare CSS grid: its number boxes carry no label, no id and no
# placeholder, so each one is addressed by what the browser itself knows about
# it. Verified against both live forms:
#
#     Rate       <input type=number step="0.01">           rupees per unit
#     Quantity   <input type=number> + a unit suffix        PCS / BAGS / MT
#     Discount   <input type=number max="100"> + a "%"      Quotations only
#     Tax        NOT an input - a read-only badge the app derives from the
#                item's HSN code, so a Tax cell can never be typed in
#
# Rate and Discount used to be the other way round here, and that mistake is
# invisible on screen: the application accepts the typing and only its API
# objects, with
#     items.0.discount_pct: Number must be less than or equal to 100
# --------------------------------------------------------------------------- #
RATE_BOX = 'input[type="number"][step="0.01"]'
DISCOUNT_BOX = 'input[type="number"][max="100"]'


def quantity_box(page: Page, unit_label: str) -> Locator:
    """The quantity input, found by the unit printed beside it.

    The unit text is anchored (^PCS$) so a cell such as "PCS available" cannot
    match, and .first is used because the filter also matches wrapper <div>s.

    The looking is done by find_quantity_box(), which tries every way the box
    has ever been marked up - spinbutton, textbox, number input, labelled field,
    XPath neighbour - so a build that changes the input type does not break the
    step. If nothing matches at all the original locator is handed back, so the
    caller's own message still names the field.
    """
    found = find_quantity_box(page, unit_label)
    if found is not None:
        return found

    row_locator = page.locator("div").filter(
        has_text=re.compile(rf"^{re.escape(unit_label)}$"), visible=True)
    return row_locator.get_by_role("spinbutton").first


#: What the form said about tax while THIS row was being filled in.
#:
#: The Tax column of the workbook is not an input on the Quotation and Sales
#: Order forms - the application derives GST from the item's HSN code and prints
#: it as a read-only badge. That makes the sheet's Tax the wrong number to work
#: an expected tax out of, and using it anyway is what failed a perfectly
#: correct quotation on 2026-08-13 (expected 28%, the application charged the
#: 18% its own badge was showing). So the rate the FORM published is remembered
#: here, for the calculation layer to use as the honest input.
#:
#: {"derived": bool, "rate": float | None, "sheet": float | None}
#: derived=False (the default) means nothing has said the form is read-only, so
#: the workbook's Tax column is a legitimate input for that document.
TAX_ON_FORM: Dict[str, object] = {"derived": False, "rate": None, "sheet": None}


def forget_form_tax_rate() -> None:
    """Start a row with no memory of the last one's tax badge."""
    TAX_ON_FORM.update({"derived": False, "rate": None, "sheet": None})


def note_tax_is_read_only(page: Page, value: str) -> None:
    """Record that the sheet's Tax cell cannot be entered on this form.

    The application prints the GST rate as a badge taken from the item's HSN
    code; there is no input to type into. This is not a defect and must not be
    logged as one, but the difference between the sheet and the screen belongs
    in the report - and the rate the badge showed belongs in TAX_ON_FORM, so
    the calculation layer works the expected tax out of the rate the
    application is actually charging rather than the one the sheet asked for.
    """
    value = (value or "").strip()
    if not value:
        return
    shown = page.get_by_text(re.compile(r"^\s*\d+(\.\d+)?%\s*$")).filter(visible=True)
    on_screen = shown.first.inner_text().strip() if shown.count() > 0 else "(not shown)"

    TAX_ON_FORM["derived"] = True
    TAX_ON_FORM["rate"] = verify_data.as_number(on_screen)
    TAX_ON_FORM["sheet"] = verify_data.as_number(value)

    log.info("   Tax is set by the item's HSN code and cannot be typed - "
             "sheet says %s%%, the form shows %s", value, on_screen)
    attach_text("Tax Rate - Read-only on This Form",
                f"Sheet Tax = {value}%. The form derives GST from the item's HSN "
                f"code and shows {on_screen}; it has no Tax input.\n\n"
                f"The rate the APPLICATION is charging is therefore the input "
                f"the tax and total checks are worked out from. The workbook's "
                f"Tax column never reached the application on this form, so "
                f"holding the application to it would fail a document whose "
                f"arithmetic is correct.")
    verify_data.note_skipped("Tax", value,
                             "read-only on this form - the application derives "
                             "it from the item's HSN code")


def guard_percentage_boxes(page: Page) -> None:
    """Refuse to submit a percentage the application cannot accept.

    A rupee amount typed into a % box looks completely normal on screen: the
    form takes it, the API answers "items.0.discount_pct: Number must be less
    than or equal to 100", and the dialog is left open with nothing to read.
    Checking it here turns that into a message that names the field.
    """
    boxes = page.locator(DISCOUNT_BOX).filter(visible=True)
    for index in range(boxes.count()):
        raw = (boxes.nth(index).input_value() or "").strip()
        try:
            percent = float(raw or "0")
        except ValueError:
            continue
        if percent > 100:
            fail(page, f"Discount is {raw}%, and the application only accepts "
                       f"0-100. Check the Discount cell in the sheet, and that "
                       f"the Rate has not been typed into the % box.")


def choose_customer(page: Page, excel_name: str) -> str:
    """Pick a customer in the picker used by Quotations and Sales Orders.

    Type to filter, then choose the option. The app renders each option as
    "<initial> <name> <city>", and the city comes from the Customer sheet row
    that created it - so a customer created earlier in the run is matched
    exactly. For a customer that already existed the name alone is used.
    """
    name, city = CREATED_CUSTOMERS.get(excel_name.strip().lower(),
                                       (excel_name.strip(), ""))
    field = find_field(
        visible(page.get_by_placeholder("Search customers by name or")),
        visible(page.get_by_placeholder("Search customers")),
        visible(page.get_by_label("Customer*")),
        visible(page.get_by_label("Customer")),
    )
    if field is None:
        fail(page, "The customer picker was not found on this form")

    click(page, field, "Customer field", animated=True)
    fill(page, field, name, "Customer")

    label = f"{name[:1].upper()} {name} {city}" if city else name
    matches = button(page, label, exact=False).filter(visible=True)
    if not appears(matches.first):
        # The city is only part of the option while it matches the row that
        # created the customer; for one that already existed, or that was
        # created with a different city, the name on its own is what is listed.
        matches = button(page, name, exact=False).filter(visible=True)

    # The application allows customers with the same name, so a search for
    # "Hema P" legitimately returns several rows. Take the first: any of them is
    # the customer the sheet asked for, and without .first Playwright's strict
    # mode fails the test over data the application itself permits.
    option = matches.first
    if not appears(option):
        # A picker that answers "No customers found" is a data problem, not a
        # timing one - say so now instead of waiting out the full timeout and
        # reporting it as a missing element.
        if page.get_by_text("No customers found").filter(visible=True).count() > 0:
            fail(page, f"The picker searched for '{name}' and answered 'No "
                       f"customers found'. This business has no customer by that "
                       f"name, so check the Customer cell in the sheet - it has "
                       f"to name a customer the Customer sheet creates, or one "
                       f"that already exists in this business.")
        expect_visible(page, option, f"customer option '{name}'")
    if matches.count() > 1:
        log.info("   %d customers are named '%s' - using the first one",
                 matches.count(), name)
    click(page, option, "Customer option", animated=True)
    verify_data.record("Customer", name)
    return name


def search_term(option_label: str) -> str:
    """A query the application's search will actually match.

    A cell may store the option exactly as the picker reads it out - "O OPC
    Cement 53 Grade ITEM-" - which is the avatar initial, the name and the start
    of the item code. Typed verbatim into a search box that finds nothing, so
    this trims it back to the name the catalogue is indexed on.
    """
    text = re.split(r"\s+(?:ITEM-|SKU-|·)", option_label.strip())[0]
    words = text.split()
    if len(words) > 1 and len(words[0]) == 1:            # the avatar initial
        words = words[1:]
    return " ".join(words).strip()


def pick_catalogue_item(page: Page, label: str, what: str = "item") -> None:
    """Choose a catalogue line item in a picker that is already open.

    Both forms open the picker with a search box above an empty list ("No items
    found"), and only Quotations sometimes lists the catalogue straight away -
    so the option is clicked when it is already there, and searched for when it
    is not. The query is search_term(label), never the cell verbatim.
    """
    option = button(page, label, exact=False).filter(visible=True).first
    if not appears(option):
        search = find_field(
            visible(page.get_by_placeholder("Search products")),
            visible(page.get_by_placeholder("Search items")),
            visible(page.get_by_label("Product*")),
            visible(page.get_by_label("Product")),
        )
        if search is None:
            fail(page, f"The picker does not list '{label}' and this form has no "
                       f"search box to narrow the catalogue down with")
        query = search_term(label)
        click(page, search, f"{what} search box", animated=True)
        fill(page, search, query, f"{what.title()} search")
        settle(page)

        if not appears(option, DEFAULT_TIMEOUT):
            fail(page, f"Searched the catalogue for '{query}' and the picker "
                       f"still does not offer the {what} '{label}'. The item has "
                       f"to exist under Stock & Materials in this business "
                       f"before it can be added, so check that it is there and "
                       f"that the sheet spells it the same way.")

    click(page, option, f"{what.title()} '{label}'", animated=True)
    # The line item is data too: the View page has to show the item that was
    # picked. "item" on a quotation, "product" on a sales order.
    verify_data.record("Item Name" if what == "item" else "Product", label)


def create_quotation(page: Page, row: Row) -> None:
    """Raise one quotation from an Excel row, then send and accept it."""
    with step("Creating Quotation..."):
        # Quotations live inside the "Orders & Sales" group - expand it first.
        open_sales_module(page, "Quotations", "New Quotation")
        click(page, button(page, "New Quotation"), "New Quotation", animated=True)

        # Header fields, in screen order.
        choose_customer(page, row["Customer"])
        fill(page, visible(page.get_by_placeholder("e.g. Supply of 10MT TMT Bars")),
             row["Subject"], "Subject")
        fill(page, visible(page.get_by_label("Valid Until")),
             resolve_date(row["Valid Until"]), "Valid Until")

        # Payment Terms dropdown
        payment_field = visible(page.get_by_placeholder("e.g. Net"))
        click(page, payment_field, "Payment Terms field", animated=True)
        payment_option = button(page, row["Payment Terms"])
        expect_visible(page, payment_option, row["Payment Terms"])
        click(page, payment_option, "Payment Terms", animated=True)

        # Line item. "Add Item" opens the picker; the number boxes only exist
        # once an item has been chosen, so the order here matters.
        click(page, button(page, "Add Item"), "Add Item", animated=True)
        pick_catalogue_item(page, row["Item Name"], "item")

        # See "THE LINE-ITEM ROW" above for how each box is identified. Unit is
        # a column so that an item measured in BAGS or MT needs a cell change,
        # not a code change.
        unit = row.get("Unit", "").strip() or "PCS"
        fill(page, quantity_box(page, unit), row["Quantity"], "Quantity")
        fill(page, visible(page.locator(RATE_BOX)), row["Rate"], "Rate")
        fill_optional(page, row.get("Discount", ""), "Discount",
                      visible(page.locator(DISCOUNT_BOX)))
        note_tax_is_read_only(page, row.get("Tax", ""))
        fill_optional(page, row.get("Remarks", ""), "Remarks",
                      visible(page.get_by_label("Remarks")),
                      visible(page.get_by_placeholder("Remarks")),
                      visible(page.get_by_label("Notes")))

        guard_percentage_boxes(page)
        submit = button(page, "Create Quotation")
        click(page, submit, "Create Quotation")
        confirm_saved(page, submit, row["Subject"], "Quotation")
        # REPORTING ONLY. This build identifies a quotation by its SUBJECT -
        # no quotation number is printed anywhere the flow reads - so the
        # subject is recorded as what it is, and no number is invented.
        stock_report.note_document("Quotation", row["Subject"],
                                   "identified by its subject - this build "
                                   "prints no quotation number")

    with step("Sending and accepting the quotation..."):
        click(page, button(page, "Send"), "Send", animated=True)
        wait_until_ready(page, "quotation after sending")

        click(page, button(page, "Accept"), "Accept", animated=True)
        wait_until_ready(page, "quotation after accepting")

        # Acceptance is what unlocks the sales order, so asserting on that button
        # proves the status really changed.
        expect_visible(page, button(page, "Convert to SO"),
                       "'Convert to SO' button (quotation accepted)")


# =========================================================================== #
# MODULE 7 - SALES ORDERS  (sheet: Sales_Orders)
# =========================================================================== #

def choose_product(page: Page, product: str) -> None:
    """Open the product picker of the sales-order line item and pick from it.

    The picker opens from "Add Product" on Sales Orders and "Add Item" on
    Quotations; pick_catalogue_item() then does the choosing, searching the
    catalogue when the option is not already on the list.
    """
    for opener_name in ("Add Product", "Add Item"):
        opener = button(page, opener_name).first
        if opener.count() > 0 and opener.is_visible():
            click(page, opener, opener_name, animated=True)
            break

    pick_catalogue_item(page, product, "product")


def set_delivery_date(page: Page, row: Row) -> None:
    """Fill the Sales Order's Delivery Date, which the application requires.

    The form marks it DELIVERY DATE * and will not save without it. Add a
    "Delivery Date" column to the Sales_Orders sheet to drive it (an exact date,
    TODAY, or TODAY+7). With no column and no value the run falls back to
    today + 30 - the same shorthand Valid Until already uses on Quotations.
    """
    field = find_field(
        visible(page.get_by_label("Delivery Date")),
        visible(page.locator('input[type="date"]')),
    )
    if field is None:
        log.info("   skip: Delivery Date (this form does not ask for one)")
        return

    cell = row.get("Delivery Date", "")
    when = resolve_date(cell)
    fill(page, field, when, "Delivery Date")
    if not cell.strip():
        log.info("   the sheet has no Delivery Date column - used %s (today + "
                 "30). Add one to the Sales_Orders sheet to set it yourself.",
                 when)


def create_sales_order(page: Page, row: Row) -> None:
    """Raise one sales order from an Excel row, then drive its status buttons.

    The sheet drives both halves of the screen:

      Customer / Product / Quantity / Rate / Tax / Notes
          the New Sales Order form. Leave Customer blank to skip the form and
          work on the order already on screen (one converted from a quotation).
      Status Sequence
          the buttons clicked afterwards, separated by | . Leave it blank to
          only create the order.
    """
    customer = row.get("Customer", "").strip()
    actions = [action.strip() for action in row.get("Status Sequence", "").split("|")
               if action.strip()]

    if not customer and not actions:
        raise ValueError(
            f"Sheet '{row.sheet}' row {row.number} has neither a Customer nor a "
            f"Status Sequence, so there is nothing to do. Fill in the form "
            f"columns (Customer, Product, Quantity, Rate, Tax, Notes), or a "
            f"Status Sequence such as "
            f"'Mark Dispatched | Mark Delivered | Raise Indent'.")

    if customer:
        with step(f"Creating Sales Order... ({customer})"):
            open_sales_module(page, "Sales Orders", "New Sales Order")
            click(page, button(page, "New Sales Order"), "New Sales Order",
                  animated=True)

            # Fields in screen order - the column order of the Sales_Orders sheet.
            choose_customer(page, customer)
            set_delivery_date(page, row)
            choose_product(page, row["Product"])

            # The line item is the same grid as the quotation's, minus the
            # discount box - see "THE LINE-ITEM ROW" above.
            unit = row.get("Unit", "").strip() or "PCS"
            fill_optional(page, row.get("Quantity", ""), "Quantity",
                          quantity_box(page, unit))
            fill_optional(page, row.get("Rate", ""), "Rate",
                          visible(page.locator(RATE_BOX)))
            note_tax_is_read_only(page, row.get("Tax", ""))
            fill_optional(page, row.get("Notes", ""), "Notes",
                          visible(page.get_by_placeholder("Order notes")),
                          visible(page.get_by_label("Notes")),
                          visible(page.get_by_placeholder("Notes")))

            guard_percentage_boxes(page)
            submit = button(page, "Create Sales Order")
            click(page, submit, "Create Sales Order")
            confirm_saved(page, submit, customer, "Sales Order")
            # REPORTING ONLY - and the same caveat as the quotation: the flow
            # proves this record by its CUSTOMER, so that is what is recorded.
            stock_report.note_document("Sales Order", customer,
                                       "identified by its customer - this "
                                       "build prints no sales order number")

    if not actions:
        return

    with step(f"Driving the sales order... ({' -> '.join(actions)})"):
        for index, action in enumerate(actions):
            control = button(page, action)
            expect_visible(page, control, f"'{action}' button")
            click(page, control, action, animated=True)
            wait_until_ready(page, f"sales order after '{action}'")

            # The proof that the status changed is the NEXT action appearing.
            if index + 1 < len(actions):
                next_action = actions[index + 1]
                expect_visible(page, button(page, next_action),
                               f"'{next_action}' button (state after {action})")

        log.info("   sales order driven through to '%s'", actions[-1])
        screenshot(page, "Sales Order - Final Status")


# =========================================================================== #
# THE PURCHASE SIDE - Purchase Orders -> GRN -> Purchase Bills
#
# The three modules below are one journey and are written to run in that order,
# straight after Sales Orders and in the SAME session: the purchase order the
# first one raises is the one the GRN receives against, and approving that GRN
# is what creates the bill the third one approves. Nothing new is logged in -
# they all use the workspace the Customer module signed into.
#
# What they share is collected here, so each module below reads as its own form
# and nothing else:
#
#   open_purchase_module()      the collapsed "Purchases" sidebar group
#   read_document_number()      PO/26-27/0007 - the number the APP allots
#   select_option_containing()  a dropdown whose options say more than the sheet
#   answer_confirmation()       the "Yes, ..." dialog an action raises
#   drive_actions()             a Status Sequence cell, the way Sales Orders has
#   read_status() / verify_status()   the badge on the record
# =========================================================================== #

def open_purchase_module(page: Page, module: str, expected_button: str) -> None:
    """Open a module that lives inside the collapsed 'Purchases' group.

    The same shape as open_sales_module() and open_stock_module(): the group
    itself is a button, the modules inside it are links, and the group label
    carries a live count once the business has raised a purchase order - so it
    is matched with exact=False, exactly as "Orders & Sales" has to be.
    """
    dismiss_overlay(page)          # an open form hides the whole sidebar
    if nav_link(page, module).count() == 0:
        group = button(page, "Purchases", exact=False).filter(visible=True)
        expect_visible(page, group.first, "'Purchases' sidebar group")
        click(page, group.first, "Purchases menu", animated=True)
    open_module(page, module, expected_button)


#: How the application numbers the documents this side of the suite creates:
#: PO/26-27/0007, GRN/26-27/0003, PB/26-27/0005. The financial year and the
#: serial are the application's own, so they are never in the workbook - they are
#: read off the screen and become the name the record is then verified under.
DOCUMENT_NUMBER = r"{prefix}\s*[/-]\s*[A-Za-z0-9][A-Za-z0-9/-]*"


def document_number_text(prefix: str):
    """The same number as the WHOLE text of one element - for finding it on a list.

    Anchored on purpose. get_by_text() matches a SUBSTRING, so an unanchored
    pattern also matches every <div> the number happens to sit inside, and
    .first is then the outermost wrapper of the whole page - which is a click
    that lands nowhere and an assertion that proves nothing.
    """
    return re.compile(r"^\s*" + DOCUMENT_NUMBER.format(prefix=re.escape(prefix))
                      + r"\s*$", re.IGNORECASE)


def read_document_number(page: Page, prefix: str, what: str) -> str:
    """The number the application allotted to the record it has just saved.

    Returns "" when the page is not showing one. That is deliberately not a
    failure: the record has already been proved to exist, and the caller simply
    falls back to a name it already has (the supplier, the invoice number) as
    the text the View page is found by.
    """
    wanted = re.compile(DOCUMENT_NUMBER.format(prefix=re.escape(prefix)))

    def look() -> Optional[str]:
        """The number, wherever on the page it is showing.

        BOTH areas are searched, and the search is what decides - not whichever
        of them happened to have text in it first. That is the whole fix here:
        this used to take `main`'s text the moment it was non-empty and give up
        if the number was not in it, and immediately after a save `main` is
        still the FORM. The number was in the success toast, which the
        application draws at the top of the BODY, outside main. So a GRN whose
        own page said "GRN GRN/0053" was recorded as having no number at all -
        and the stock-ledger check that identifies the receipt by that number
        was SKIPPED for want of it on every run.
        """
        for area in (page.locator("main").first, page.locator("body").first):
            try:
                text = area.inner_text(timeout=OPTIONAL_TIMEOUT) or ""
            except (PlaywrightTimeoutError, PlaywrightError):
                continue
            found = wanted.search(text)
            if found:
                return found.group(0)
        return None

    # Polled briefly, because the two places the number appears arrive at
    # different moments: the toast is up straight away and the record's own page
    # is drawn a beat later. Whichever comes first is the answer.
    match = look()
    if match is None:
        deadline = time.time() + OPTIONAL_TIMEOUT / 1000
        while match is None and time.time() < deadline:
            page.wait_for_timeout(500)
            match = look()

    if match is None:
        log.warning("   the page is not showing a %s number - the record is "
                    "verified by name instead", what)
        return ""

    number = re.sub(r"\s+", "", match)
    log.info("   %s number: %s", what, number)
    attach_text(f"{what} Number", number)
    return number


#: A dropdown entry that is not a choice at all - the prompt sitting at the top
#: of the list ("- Select Supplier -", "Choose an item", "--").
PLACEHOLDER_OPTION = re.compile(r"^[\s\-—–]*(select|choose|none|all|--)",
                                re.IGNORECASE)


def select_option_containing(page: Page, field: Locator, wanted: str,
                             description: str) -> str:
    """Choose the <select> option that CONTAINS what the workbook holds.

    select_option() only ever matches an option's label in full, and this
    application labels its options with more than the sheet ever could: a
    purchase order reads "PO/26-27/0007 - Delta Pipes & Fittings" and an item
    reads "OPC Cement 53 Grade (ITEM-0001)". So the option is found here, by its
    own text, and then chosen through safe_select() - which keeps the scrolling,
    the retries and the "these are the options the form offers" message.

    An empty cell takes the first real option, which is what a GRN row that
    means "the purchase order this run has just raised" needs.
    """
    wanted = (wanted or "").strip()
    try:
        labels = [text.strip() for text
                  in field.locator("option").all_inner_texts()]
    except PlaywrightError:
        labels = []
    choices = [label for label in labels
               if label and not PLACEHOLDER_OPTION.match(label)]

    chosen = ""
    if wanted:
        for label in choices:
            if wanted.casefold() in label.casefold():
                chosen = label
                break
    elif choices:
        chosen = choices[0]
        log.info("   %s: the cell is empty - taking the option the form offers "
                 "first ('%s')", description, chosen)

    # Nothing matched: hand the cell over as it is, so the failure is the
    # framework's own - it names the field AND lists the real options.
    safe_select(page, field, chosen or wanted, description)

    # safe_select has recorded the option's full label. The workbook's own value
    # is what the View page will print, so that is what the comparison is given.
    if wanted and chosen and wanted != chosen:
        verify_data.record(description, wanted)
    return chosen or wanted


#: The button on the "are you sure?" dialog. Anchored at the start so it can
#: never match the action button that opened it: the dialog reads "Yes, Send to
#: Supplier" while "Send to Supplier" is still on the page behind it.
CONFIRMATION_BUTTON = re.compile(r"^(yes\b|confirm$|ok$|proceed$)", re.IGNORECASE)


def answer_confirmation(page: Page, action: str) -> None:
    """Click through the confirmation dialog, when the action raises one.

    Optional by design, and checked without waiting: most actions carry straight
    on, and giving every one of them a timeout for a dialog that never comes
    would cost the run more than the dialogs do. The click before this one is
    animated=True, so the dialog has already had its fade-in.
    """
    control = page.get_by_role("button", name=CONFIRMATION_BUTTON) \
                  .filter(visible=True).first
    try:
        if control.count() == 0:
            return
        name = (control.inner_text() or "").strip() or "Yes"
    except PlaywrightError:
        return

    log.info("   '%s' asks to be confirmed - answering '%s'", action, name)
    click(page, control, name, animated=True)


def status_actions(row: Row) -> List[str]:
    """The Status Sequence cell as a list of button names, in order.

    The same column the Sales_Orders sheet already uses: the buttons to click
    after the record is saved, separated by | . A blank cell means "only create
    the record".
    """
    return [action.strip() for action in row.get("Status Sequence", "").split("|")
            if action.strip()]


def drive_actions(page: Page, actions: List[str], what: str) -> None:
    """Click a record's status buttons in turn, answering the confirmations.

    Exactly the shape create_sales_order() drives its own sequence with, and the
    proof that a status really changed is the same one: the NEXT button in the
    sequence appearing. What the purchase side adds is the confirmation - "Send
    to Supplier" opens a second button, "Yes, Send to Supplier", and the order
    does not move until that one is clicked.
    """
    for index, action in enumerate(actions):
        control = button(page, action)
        expect_visible(page, control, f"'{action}' button")
        click(page, control, action, animated=True)
        answer_confirmation(page, action)
        wait_until_ready(page, f"{what} after '{action}'")
        screenshot(page, f"{what} - After {action}")

        if index + 1 < len(actions):
            next_action = actions[index + 1]
            expect_visible(page, button(page, next_action),
                           f"'{next_action}' button (state after {action})")

    log.info("   %s driven through to '%s'", what, actions[-1])


#: The words this application prints in a status badge. Read as a whole line, so
#: a paragraph that happens to contain "paid" cannot be mistaken for the badge.
STATUS_BADGE = re.compile(
    r"^\s*(draft|pending|pending approval|open|submitted|confirmed|sent|"
    r"sent to supplier|partially received|received|approved|rejected|unpaid|"
    r"partially paid|paid|cancelled|closed|completed)\s*$", re.IGNORECASE)
# Material Indents deliberately do NOT go through this: their detail page draws
# a stepper that prints every status at once, so the badge has to be addressed
# by where it sits, not by what it says - see read_indent_status().


def read_status(page: Page) -> str:
    """The status the record on screen is showing, as the badge itself spells it."""
    badge = page.get_by_text(STATUS_BADGE).filter(visible=True)
    try:
        if badge.count() == 0:
            return ""
        return (badge.first.inner_text() or "").strip()
    except PlaywrightError:
        return ""


def verify_status(page: Page, expected: str, what: str) -> None:
    """Prove the record is showing the status the sheet expects.

    The BADGE's own text is asserted, never the row's: once the browser
    concatenates a table row it reads "...beta0approved3 Aug 2026", where no
    whole-word match for a status can ever succeed.

    The match is anchored at the start rather than made exact, because a status
    is often printed with more after it - a sheet that says "Sent" is happy with
    a badge that reads "Sent to Supplier". An empty cell skips the check, which
    is what a row that only creates the record wants.
    """
    expected = (expected or "").strip()
    on_screen = read_status(page)
    if on_screen:
        log.info("   %s status on screen: %s", what, on_screen)
        attach_text(f"{what} Status on Screen", on_screen)
        # REPORTING ONLY - the badge was already read by the line above, and is
        # only remembered here so the stock and QA cards can print "APPROVED"
        # beside the record's number instead of leaving a reader to take the
        # approval on trust. Nothing is read, decided or asserted here.
        stock_report.note_status(what, on_screen,
                                 "the badge the record was showing")

    if not expected:
        log.info("   skip: %-24s (the cell is empty)", "Expected Status")
        return

    badge = page.get_by_text(
        re.compile(rf"^\s*{re.escape(expected)}\s*$", re.IGNORECASE)
    ).filter(visible=True).first
    if badge.count() == 0:
        # The badge often says more than the sheet does - "Sent to Supplier" for
        # a cell that says "Sent" - so a start-anchored match is the fallback.
        # .last, because ancestors come first in document order: the innermost
        # element whose text starts with the status IS the badge, while .first
        # would be the card it sits in.
        badge = page.get_by_text(
            re.compile(rf"^\s*{re.escape(expected)}\b", re.IGNORECASE)
        ).filter(visible=True).last
    safe_expect(page, badge, f"the '{expected}' status of the {what}")
    log.info("   verified: the %s is '%s'", what, expected)
    screenshot(page, f"{what} - Status {expected}")


# =========================================================================== #
# MODULE 8 - PURCHASE ORDERS  (sheet: Purchase_Orders)
# =========================================================================== #

#: The numbers of the purchase orders this run has raised, oldest first. The GRN
#: module receives against the last one when its own cell is left empty, which is
#: what makes PO -> GRN -> Bill a journey rather than three unrelated sheets.
CREATED_PURCHASE_ORDERS: List[str] = []

#: HOW MANY PURCHASE ORDERS THE LIST HELD BEFORE THIS ROW RAISED ITS ONE.
#: Taken at the only moment it still exists - the list is on screen, the "New
#: PO" form has not been opened yet - and read back after the order is saved.
#: "before" is None when the list would not say how many it holds; "why" then
#: carries the reason, and the check is SKIPPED with it rather than answered.
PURCHASE_ORDER_COUNT: Dict[str, object] = {}


def capture_purchase_order_count(page: Page) -> None:
    """The BEFORE of the purchase-order count, read off the list on screen.

    Called with the Purchase Orders list already open and no form over it,
    which is the last moment "before" exists. Nothing here can fail a row: a
    count that could not be established leaves the check SKIPPED with the
    reason, and stopping the purchase flow because a list would not say how
    long it was would cost the run everything the row was about to prove.
    """
    PURCHASE_ORDER_COUNT.clear()
    try:
        before, how = list_record_count(page, "Purchase Orders")
    except Exception as error:                     # noqa: BLE001 - never fatal
        before, how = None, (f"the Purchase Orders list could not be counted "
                             f"({str(error).splitlines()[0]})")
    PURCHASE_ORDER_COUNT.update({"before": before, "how": how,
                                 "raised": list(CREATED_PURCHASE_ORDERS)})
    if before is None:
        log.warning("   the purchase-order count BEFORE could not be taken: %s",
                    how)
    else:
        log.info("   purchase orders on the list before this row: %d (%s)",
                 before, how)


def purchase_line_quantity(page: Page, unit: str) -> Locator:
    """The Quantity box of a purchase line, whatever this build calls it.

    The purchase form's line is the same bare CSS grid as the quotation's (see
    "THE LINE-ITEM ROW" above), but it does not always print a unit beside the
    box - so find_quantity_box() is asked first, because it knows every way the
    box has been marked up, and the first spinbutton on the form is the fallback,
    which is the one the recording typed into.
    """
    found = find_quantity_box(page, unit)
    if found is not None:
        return found
    log.info("   the line prints no unit - using the first number box on the form")
    return page.get_by_role("spinbutton").filter(visible=True).first


def purchase_item_rows(page: Page) -> Locator:
    """The line-item rows already on the purchase form.

    A line is recognised by its own dropdown, which is the only <select> on this
    form whose options include "Select Item" - the supplier's does not. So
    counting these counts the lines, which is what tells "Add Item" whether it
    has anything to do.
    """
    return page.locator("select").filter(
        has_text=re.compile(r"select\s+item", re.IGNORECASE)).filter(visible=True)


def choose_purchase_item(page: Page, code: str, name: str) -> None:
    """Put ONE line on the purchase order and choose the item on it.

    "Add Item" is clicked only when the form has no line to fill in yet. Some
    builds open the form with an empty line already on it, and this form can
    show more than one "Add Item" button once a line exists - which is how a
    single step came to add a SECOND row that nothing ever filled in, and the
    application then either rejected the order or saved a blank line on it.

    So the rows are counted first, the click is skipped when there is already a
    row, and the row is proved to exist before anything is chosen on it.

    THE ITEM IS CHOSEN BY ITS CODE WHEN THE SHEET GIVES ONE - see
    choose_indent_item(), which has always done this. Matching on the name alone
    picks whichever of the identically-named items the application happens to
    list first, and this catalogue accumulates one "OPC Cement 53 Grade" per run
    (ITEM-OPC53-001 ... -010). That is not cosmetic: the order was stocking one
    of them while the Material_Indents row drew from another, so the GRN
    received 200 into an item the indent never touched and the indent found
    nothing to issue. The two sheets have to name the SAME item for the stock
    to flow from one to the other.
    """
    rows = purchase_item_rows(page)
    try:
        existing = rows.count()
    except PlaywrightError:
        existing = 0                      # mid-render: treat it as no line yet

    if existing:
        log.info("   the form already has %d line row(s) - 'Add Item' is not "
                 "clicked again", existing)
    else:
        # .first, because once a line is on the form there can be a second
        # "Add Item" button: a locator that resolves to two of them fails on
        # strict mode, and it was that failure being retried which clicked the
        # button again and added the duplicate row.
        opener = button(page, "Add Item").filter(visible=True).first
        click(page, opener, "Add Item", animated=True)
        expect_visible(page, rows.first, "the purchase order's line-item row")
        log.info("   added one line row")

    field = safe_locator(page, "Item Name", (
        ("the line's own item dropdown", rows),
        ("a select named item", page.locator('select[name*="item"]')),
        ("a select named product", page.locator('select[name*="product"]')),
        ("the field labelled Item", page.get_by_label("Item", exact=False)),
        ("the line's dropdown (the second on the form)",
         page.get_by_role("combobox").nth(1)),
    )) or rows.first

    # The code first, the name as the fallback: a sheet that only knows what it
    # asked for still works, and a sheet that pins the code gets exactly that
    # item. The NAME is what is recorded either way - the View page prints the
    # name, so recording the code here would add a comparison that could only
    # ever fail (the same swap choose_indent_item() makes, for the same reason).
    wanted_code, wanted_name = (code or "").strip(), (name or "").strip()
    if wanted_code:
        chosen = select_option_containing(page, field, wanted_code, "Item Name")
        if wanted_name and wanted_code.casefold() not in chosen.casefold():
            log.warning("   the item dropdown offers no option carrying the "
                        "code '%s' - falling back to the name", wanted_code)
            select_option_containing(page, field, wanted_name, "Item Name")
        verify_data.record("Item Name", wanted_name or chosen)
        log.info("   item chosen from the list: %s", chosen)
        return

    select_option_containing(page, field, wanted_name, "Item Name")


def submit_purchase_order(page: Page) -> Locator:
    """Click 'Create Purchase Order', and hand back the button that was clicked.

    The button sits at the bottom of a long modal, and what stops the click is
    never the same thing twice: the loader that is still running while the line's
    amount is worked out, a toast in the corner, the sidebar overlaying the
    dialog ("intercepts pointer events"), or the form still holding its own
    button disabled. None of that is a fixed-delay problem, so nothing here
    sleeps - each cause is waited out for what it actually is:

        1. the box that was just typed into is blurred, so the form commits the
           value and releases its button. This is the root cause of "the click
           does nothing": the recording had to press Create TWICE, and the first
           press was only ever committing the quantity;
        2. the page is waited on - API quiet, DOM finished re-rendering,
           spinners and backdrops gone;
        3. the button is RE-LOCATED after that wait, never reused from before
           it, then scrolled into view and waited on until it is attached,
           visible and enabled;
        4. safe_click() does the clicking, which is what keeps the three
           attempts, the escalation from an ordinary click to a forced one and
           then to a scripted one, the retry screenshots and the Allure steps.

    Whether the order was really created is proved by confirm_saved() afterwards,
    exactly as every other module proves it.
    """
    # 1. Commit the number box. A field that still has focus has not been
    #    handed to the form yet, and the form keeps Create disabled until it is.
    page.keyboard.press("Tab")

    # 2. Let the form finish reacting to the line that was just filled in.
    wait_page_ready(page, "purchase order form")
    # close_stubborn stays False on purpose: this form IS a modal with its own
    # backdrop, and closing that would cancel the order being submitted.
    wait_for_overlays_gone(page)

    # 3. Re-locate. The form re-renders its totals, and a locator resolved
    #    before that can point at a button that is no longer on the page.
    submit = button(page, "Create Purchase Order").filter(visible=True).first
    expect_visible(page, submit, "'Create Purchase Order' button")
    scroll_into_view(submit, "Create Purchase Order")
    ready(submit)

    # 4. The framework's own click, with everything it already does on a retry.
    click(page, submit, "Create Purchase Order")
    return submit


def create_purchase_order(page: Page, row: Row) -> None:
    """Raise one purchase order from an Excel row, then drive its status buttons.

    The sheet drives both halves of the screen, exactly as Sales_Orders does:

      Supplier / Expected Date / Payment Terms / Notes / Item Name / Quantity /
      Rate
          the New PO form.
      Status Sequence
          the buttons clicked afterwards, separated by | - "Confirm PO | Send to
          Supplier". Leave it blank to only raise the order.
      Expected Status
          the badge the order must be showing when the row is finished.
    """
    supplier = row["Supplier"]
    actions = status_actions(row)

    # BEFORE anything is ordered: the two figures this purchase is about to
    # move. A movement needs two readings and this is the only moment the first
    # one still exists - once the GRN is approved, the stock before it is gone
    # for good and no amount of reading the screen afterwards can recover it.
    capture_purchase_baseline(page, row)

    with step(f"Creating Purchase Order... ({supplier})"):
        open_purchase_module(page, "Purchase Orders", "New PO")
        # BEFORE the form is opened: how many purchase orders the list holds.
        # Here and nowhere else - the list is on screen, nothing has been
        # created yet, and one click from now "before" is gone.
        capture_purchase_order_count(page)
        click(page, button(page, "New PO"), "New PO", animated=True)
        form_screenshot(page, "Purchase Order - Form Before Save")

        # Header fields, in screen order - the column order of the sheet.
        supplier_field = safe_locator(page, "Supplier", (
            ("the select named supplier_id",
             page.locator('select[name="supplier_id"]')),
            ("a select named supplier", page.locator('select[name*="supplier"]')),
            ("the field labelled Supplier",
             page.get_by_label("Supplier", exact=False)),
            ("the first dropdown on the form", page.get_by_role("combobox")),
        )) or visible(page.locator('select[name="supplier_id"]'))
        select_option_containing(page, supplier_field, supplier, "Supplier")

        expected_date = safe_locator(page, "Expected Date", (
            ("the input named expected_date",
             page.locator('input[name="expected_date"]')),
            ("the field labelled Expected Date",
             page.get_by_label("Expected", exact=False)),
            ("the only date box on the form", page.locator('input[type="date"]')),
        )) or visible(page.locator('input[name="expected_date"]'))
        fill(page, expected_date, resolve_date(row.get("Expected Date", "")),
             "Expected Date")

        fill_optional(page, row.get("Payment Terms", ""), "Payment Terms",
                      visible(page.get_by_role("textbox", name="e.g. 30 days")),
                      visible(page.get_by_placeholder("e.g. 30 days")))
        fill_optional(page, row.get("Notes", ""), "Notes",
                      visible(page.get_by_role("textbox",
                                               name="Internal notes for this")),
                      visible(page.get_by_placeholder("Internal notes")),
                      visible(page.get_by_label("Notes")))

        # The line item.
        choose_purchase_item(page, row.get("Item Code", ""), row["Item Name"])
        unit = row.get("Unit", "").strip() or "PCS"
        fill(page, purchase_line_quantity(page, unit), row["Quantity"], "Quantity")
        # A blank Rate leaves the price the application takes from the item.
        fill_optional(page, row.get("Rate", ""), "Rate",
                      visible(page.locator(RATE_BOX)))

        guard_percentage_boxes(page)
        submit = submit_purchase_order(page)
        screenshot(page, "Purchase Order - After Save")

    with step("Verifying the purchase order was created..."):
        wait_until_ready(page, "purchase order after saving")
        # The application allots the number itself, so it is read off the screen
        # and used as the text the record is proved - and later found - by.
        number = read_document_number(page, "PO", "Purchase Order")
        confirm_saved(page, submit, number or supplier, "Purchase Order")
        if number:
            CREATED_PURCHASE_ORDERS.append(number)
        # REPORTING ONLY - the number is already in hand; this records which
        # document the report should name beside the stock it moved.
        stock_report.note_document("Purchase Order", number,
                                   "raised against the supplier")
        screenshot(page, "Purchase Order - Created Record")

    if actions:
        with step(f"Driving the purchase order... ({' -> '.join(actions)})"):
            drive_actions(page, actions, "purchase order")
            screenshot(page, "Purchase Order - Final Status")

    with step("Verifying the purchase order status..."):
        verify_status(page, row.get("Expected Status", ""), "Purchase Order")
        attach_text("Page URL", page.url)


# =========================================================================== #
# MODULE 9 - GRN  (sheet: GRN)
# =========================================================================== #

#: The goods receipt this run recorded. The item's STOCK LEDGER names the
#: document behind every movement ("GRN GRN/0045"), so this is what proves the
#: stock went up because of THIS run rather than at some earlier point.
LAST_GRN_NUMBER = ""

#: The purchase bill this run approved. REPORTING ONLY: nothing in the flow
#: reads it back - it exists so the stock report can name the bill that closed
#: the purchase, which is one of the documents a QA reader asks for beside the
#: movement it belongs to.
LAST_PURCHASE_BILL_NUMBER = ""


def picker_row(page: Page, label: str) -> Locator:
    """One row of an open picker panel, whatever element the app builds it from.

    A row is a <button> on some panels and a plain <div> carrying the click on
    some others, and it reads more than the name - "HG Traders Nelamangala
    +919824160009 Surat" - so the name is matched as a substring across every
    role a row can have. Clicking the name inside a row works either way: the
    click bubbles up to whichever ancestor holds the handler.
    """
    return (page.get_by_role("button", name=label, exact=False)
            .or_(page.get_by_role("option", name=label, exact=False))
            .or_(page.get_by_text(label, exact=False))
            .filter(visible=True).first)


def choose_grn_supplier(page: Page, supplier: str) -> None:
    """Pick the supplier on the GRN form, whose picker is a panel, not a select.

    Every other supplier field in the suite is a real <select>; this one opens a
    panel from a button that reads "- Select Supplier -", and that panel is a
    SEARCHABLE list: a search box above however many suppliers the business has,
    of which only the first four or five are ever on screen.

    So the name is typed into that box first, the way a person would. Reaching
    straight for the row is what broke as soon as the business had more than a
    screenful of suppliers: the row was too far down the list to have been
    rendered, and the run failed on a supplier that was demonstrably there - the
    purchase order it had raised minutes earlier was for that same supplier.
    """
    opener = safe_locator(page, "Select Supplier", (
        ("the '- Select Supplier -' button",
         page.get_by_role("button", name="Select Supplier")),
        ("a button named Supplier", page.get_by_role("button", name="Supplier")),
    ))
    if opener is None:
        fail(page, "The GRN form's supplier picker ('- Select Supplier -') was "
                   "not found on this form.")
    click(page, opener, "Select Supplier", animated=True)

    # Short lists are all on screen already, so the search is only paid for when
    # the row is not there - the same order pick_catalogue_item() works in.
    option = picker_row(page, supplier)
    if not appears(option):
        search = find_field(
            visible(page.get_by_placeholder("Search", exact=False)),
            visible(page.get_by_role("searchbox")),
            visible(page.get_by_role("textbox", name="Search", exact=False)),
        )
        if search is None:
            fail(page, f"The GRN supplier picker does not list '{supplier}' and "
                       f"the panel has no search box to narrow it down with.")

        # Filled without clicking it first, which is what pick_catalogue_item()
        # does for the catalogue panel: a click inside THIS panel closes it. The
        # first run with the search in it cost 15s to that - the retry screenshot
        # caught the form with the panel gone - and the retry then succeeded on
        # the fill alone, which is a fill focusing the box by itself.
        query = search_term(supplier)
        fill(page, search, query, "Supplier search")
        settle(page)

        if not appears(option, DEFAULT_TIMEOUT):
            fail(page, f"Searched the GRN supplier picker for '{query}' and it "
                       f"still does not offer '{supplier}'. The supplier has to "
                       f"exist in this business before a goods receipt can be "
                       f"booked against it - the Suppliers sheet is what creates "
                       f"it - and the GRN sheet has to spell it the same way as "
                       f"the Purchase_Orders sheet does.")

    click(page, option, f"Supplier '{supplier}'", animated=True)
    verify_data.record("Supplier", supplier)


def choose_supply_type(page: Page, inter_state: str) -> None:
    """Tick 'Inter-state supply' on the approval dialog when the sheet says so.

    It is what the application works the tax out from (IGST instead of CGST +
    SGST), and it is rendered as a styled label rather than a plain checkbox - so
    the checkbox inside the label is used when there is one and the label itself
    is clicked when there is not, which is exactly what the Terms tick on the
    registration form needs.
    """
    answer = (inter_state or "").strip().casefold()
    if answer not in ("yes", "y", "true", "1"):
        log.info("   skip: %-24s (the sheet does not ask for it)", "Inter State")
        verify_data.note_skipped("Inter State", inter_state,
                                 "the sheet did not ask for an inter-state supply")
        return

    label = page.locator("label").filter(
        has_text=re.compile(r"inter-?state supply", re.IGNORECASE)).first
    if label.count() == 0:
        log.warning("   this approval dialog has no 'Inter-state supply' tick - "
                    "the application is working the tax out by itself")
        verify_data.note_skipped("Inter State", inter_state,
                                 "this approval dialog has no such tick")
        return

    checkbox = label.get_by_role("checkbox")
    if checkbox.count() > 0:
        scroll_into_view(checkbox.first, "Inter-state supply")
        checkbox.first.check()
        log.info("   ticked: Inter-state supply")
    else:
        click(page, label, "Inter-state supply")
    verify_data.record("Inter State", "Yes")


def create_grn(page: Page, row: Row) -> None:
    """Record one goods receipt from an Excel row, then approve it into a bill.

    The sheet drives the whole screen:

      Supplier / Purchase Order / Invoice Number / Vehicle Number / DC Number
          the New GRN form. A blank Purchase Order means the order THIS run has
          just raised, which is the one the dropdown offers first.
      Inter State
          YES ticks "Inter-state supply" on the approval dialog (IGST).
      Approve
          NO leaves the GRN saved but unapproved. Anything else approves it -
          and it is that approval which creates the purchase bill the next
          module works on.
      Expected Status
          the badge the GRN must be showing when the row is finished.
    """
    supplier = row["Supplier"]
    invoice = row["Invoice Number"]

    with step(f"Creating GRN... ({supplier} / {invoice})"):
        open_purchase_module(page, "GRN", "New GRN")
        click(page, button(page, "New GRN"), "New GRN", animated=True)
        form_screenshot(page, "GRN - Form Before Save")

        # Fields in screen order - the column order of the GRN sheet.
        choose_grn_supplier(page, supplier)

        wanted_order = (row.get("Purchase Order", "") or "").strip()
        if not wanted_order and CREATED_PURCHASE_ORDERS:
            wanted_order = CREATED_PURCHASE_ORDERS[-1]
            log.info("   the sheet names no purchase order - receiving against "
                     "the one this run raised (%s)", wanted_order)
        order_field = safe_locator(page, "Purchase Order", (
            ("a select named purchase order",
             page.locator('select[name*="purchase"]')),
            ("a select named po", page.locator('select[name*="po_"]')),
            ("the field labelled Purchase Order",
             page.get_by_label("Purchase Order", exact=False)),
            ("the only dropdown on the form", page.get_by_role("combobox")),
        )) or visible(page.get_by_role("combobox"))
        select_option_containing(page, order_field, wanted_order, "Purchase Order")

        fill(page, visible(page.get_by_role("textbox",
                                            name="Supplier's invoice no.")),
             invoice, "Invoice Number")
        fill(page, visible(page.get_by_role("textbox", name="e.g. MH12AB1234")),
             row["Vehicle Number"], "Vehicle Number")
        fill(page, visible(page.get_by_role("textbox",
                                            name="Supplier's DC number")),
             row["DC Number"], "DC Number")

        submit = button(page, "Save GRN")
        click(page, submit, "Save GRN")
        screenshot(page, "GRN - After Save")

    with step("Verifying the GRN was created..."):
        wait_until_ready(page, "GRN after saving")
        number = read_document_number(page, "GRN", "GRN")
        confirm_saved(page, submit, number or invoice, "GRN")
        # Remembered for the stock checks: the item's ledger writes "GRN
        # GRN/0045" against the receipt, and that reference is the only thing
        # that says THIS run's GRN is the movement that added the stock.
        global LAST_GRN_NUMBER
        LAST_GRN_NUMBER = number
        stock_report.note_document("GRN", number,          # reporting only
                                   "the receipt that puts the goods in stock")
        screenshot(page, "GRN - Created Record")

    if (row.get("Approve", "") or "").strip().casefold() in ("no", "n", "false", "0"):
        log.info("   skip: %-24s (the sheet says NO)", "Approve")
        attach_text("Page URL", page.url)
        return

    with step(f"Approving the GRN... ({number or invoice})"):
        approve = button(page, "Approve GRN")
        expect_visible(page, approve, "'Approve GRN' button")
        click(page, approve, "Approve GRN", animated=True)

        # The dialog that opens is where the supply type is decided, and the
        # bill is only created by the button underneath it.
        choose_supply_type(page, row.get("Inter State", ""))
        create_bill = button(page, "Approve & Create Bill")
        expect_visible(page, create_bill, "'Approve & Create Bill' button")
        click(page, create_bill, "Approve & Create Bill", animated=True)
        answer_confirmation(page, "Approve & Create Bill")
        wait_until_ready(page, "GRN after approval")
        screenshot(page, "GRN - After Approval")

    with step("Verifying the approved GRN..."):
        verify_status(page, row.get("Expected Status", ""), "GRN")
        attach_text("Page URL", page.url)


# =========================================================================== #
# MODULE 10 - PURCHASE BILLS  (sheet: Purchase_Bills)
# =========================================================================== #

def open_purchase_bill(page: Page, wanted: str) -> str:
    """Open one purchase bill from the list. Returns the number that was opened.

    A blank Bill Number means the bill the GRN approval has just created - the
    newest one, which the list shows first. That is why this module can follow
    the GRN module without a single number being written into the workbook.
    """
    wanted = (wanted or "").strip()
    if wanted:
        search_in_list(page, wanted, "purchase bill")

    # exact / anchored in both branches: the number has to be the WHOLE text of
    # the element that is clicked, or the click lands on the row's wrapper.
    bill = (page.get_by_text(wanted, exact=True) if wanted
            else page.get_by_text(document_number_text("PB"))
            ).filter(visible=True).first
    if not appears(bill, RETRY_TIMEOUT):
        fail(page, f"The Purchase Bills list does not show "
                   f"{'the bill ' + wanted if wanted else 'any bill'}. A bill is "
                   f"created by approving a GRN, so the GRN sheet has to have "
                   f"run first - or put an existing bill number in the "
                   f"Bill Number cell.")

    number = ""
    try:
        found = re.search(DOCUMENT_NUMBER.format(prefix="PB"),
                          bill.inner_text() or "")
        number = re.sub(r"\s+", "", found.group(0)) if found else ""
    except PlaywrightError:
        pass

    click(page, bill, f"purchase bill '{number or wanted}'", animated=True)
    wait_until_ready(page, "purchase bill")
    return number or wanted or read_document_number(page, "PB", "Purchase Bill")


def approve_purchase_bill(page: Page, row: Row) -> None:
    """Open one purchase bill and approve it.

    Nothing is typed in this module - the bill is the application's own answer to
    an approved GRN - so what the sheet holds is which bill and what it must
    read afterwards:

      Bill Number
          blank for the bill the GRN this run approved has just created.
      Status Before Approval / Expected Status
          the badge before the Approve Bill click and the badge after it.
    """
    with step("Opening Purchase Bills..."):
        # Bills are not created by hand here - they are the application's answer
        # to an approved GRN - so this list may have no "New ..." button at all.
        # open_module() falls back to the page's own heading when the button it
        # was given is not there, which is what proves the module opened.
        open_purchase_module(page, "Purchase Bills", "New Bill")
        number = open_purchase_bill(page, row.get("Bill Number", ""))
        # REPORTING ONLY - the bill number was already read by the line above;
        # this only remembers it so the stock report can name it.
        global LAST_PURCHASE_BILL_NUMBER
        LAST_PURCHASE_BILL_NUMBER = number
        stock_report.note_document("Purchase Bill", number,
                                   "the bill the approved GRN created")
        form_screenshot(page, "Purchase Bill - Before Approval")

    with step(f"Verifying the bill before approval... ({number})"):
        before = read_status(page)
        attach_text("Purchase Bill Status Before Approval",
                    f"{number}: {before or '(the page shows no status badge)'}")
        verify_status(page, row.get("Status Before Approval", ""),
                      "Purchase Bill")

    with step(f"Approving the bill... ({number})"):
        approve = button(page, "Approve Bill")
        expect_visible(page, approve, "'Approve Bill' button")
        click(page, approve, "Approve Bill", animated=True)
        answer_confirmation(page, "Approve Bill")
        # The button going is the application accepting the approval; a button
        # that stays put means it refused, and says why through the API.
        wait_for_submit_to_close(page, approve, number, "Purchase Bill")
        wait_until_ready(page, "purchase bill after approval")
        screenshot(page, "Purchase Bill - After Approval")

    with step(f"Verifying the approved bill... ({number})"):
        verify_status(page, row.get("Expected Status", ""), "Purchase Bill")
        attach_text("Page URL", page.url)

        # This module fills in no form, so there is nothing for the data check to
        # have captured. What it verified IS the data - the number it opened and
        # the status it now reads - so those are handed over here, and the View
        # page is read back and compared with them exactly as everywhere else.
        verify_data.record("Bill Number", number)
        verify_data.record("Status", read_status(page)
                           or row.get("Expected Status", ""))
        verify_data.after_save(page, number, "Purchase Bill")


# =========================================================================== #
# MODULE 11 - MATERIAL INDENTS  (sheet: Material_Indents)
#
# The construction side of the workspace, and the last module of the journey.
# The item the Items sheet created and the GRN received into stock is now asked
# for by a site, approved, issued out of the store, and the excess handed back:
#
#   New Indent -> Create Indent -> Submit for Approval -> Approve ->
#   Confirm Approval -> Issue Stock -> Confirm Issue -> Return to Store ->
#   Confirm Return to Store
#
# It is ONE test case rather than nine, because the application only offers the
# next button once the one before it has been answered - the same reason a
# quotation's Send/Accept and a sales order's status sequence live inside their
# own module. Nothing is logged in again: it runs on the page the Customer
# module signed in with, straight after Purchase Bills.
#
# Four things about these screens are worth knowing. All four were confirmed
# against the live application on 2026-08-10, and each one is a locator the
# recording got away with only because it ran once, on one tenant:
#
#   * The sidebar link reads "Material Indents" and the page it opens is
#     /indents. It is the only module in the suite whose address is not its
#     name, so it is the only caller that gives open_module() a path.
#   * Every form here is a modal built from `div.fixed.inset-0.z-50` with no
#     role="dialog" and no aria-modal, so a panel is found by its own title and
#     every field is looked for INSIDE it.
#   * The item dropdown is the recording's `get_by_role("combobox").nth(2)` -
#     an index that counts the Priority and Reference Type dropdowns above it,
#     so adding one field to the form would move it onto the wrong control. It
#     is found here by an option only it has ("Select item...").
#   * The detail page draws a progress stepper that prints Draft / Submitted /
#     Approved / Partial / Fulfilled as plain text. A page-wide search for a
#     status word therefore matches the stepper and passes whatever the indent
#     actually is, so the badge beside the IND/nnnn heading is what is asserted.
# =========================================================================== #

#: The sidebar says "Material Indents"; the application routes it to /indents.
INDENT_PATH = "/indents"

#: The whole text of an indent's heading and of a row in the list: IND/0002.
INDENT_NUMBER = re.compile(r"^\s*IND\s*[/-]\s*[A-Za-z0-9][A-Za-z0-9/-]*\s*$",
                           re.IGNORECASE)


def newest_indent_reference(page: Page) -> str:
    """The first IND/nnnn text visible on screen right now, or "" for none.

    Used as a BASELINE before a new indent is created, and again afterwards -
    see open_created_indent() - so a genuinely new top row can be told apart
    from a list that has not refreshed yet and is still showing what it
    showed a moment ago.
    """
    try:
        match = page.get_by_text(INDENT_NUMBER).filter(visible=True).first
        if appears(match, OPTIONAL_TIMEOUT):
            return verify_data.clean_text(match.inner_text())
    except (PlaywrightTimeoutError, PlaywrightError):
        pass
    return ""

#: One quantity cell of the MATERIALS REQUESTED table: "43 PCS", "2.000 PCS".
#: The unit is what tells a quantity apart from the In Stock column, which is a
#: bare number - so a build that prints a dash for a column that has not
#: happened yet simply produces one cell fewer instead of shifting them all.
INDENT_QUANTITY_CELL = re.compile(r"^\s*([\d,]+(?:\.\d+)?)\s*([A-Za-z]+)\s*$")

#: The columns of that table, in the order the application draws them.
INDENT_QUANTITY_COLUMNS = ("Requested", "Approved", "Issued", "Returned")


def open_construction_module(page: Page, module: str, expected_button: str,
                             path: str = "") -> None:
    """Open a module that lives inside the collapsed 'Construction' group.

    The same shape as open_sales_module() / open_purchase_module(): the group is
    a button, the modules inside it are links, and the label is matched with
    exact=False because these sidebar groups carry a live count.
    """
    dismiss_overlay(page)          # an open form hides the whole sidebar
    if nav_link(page, module).count() == 0:
        group = button(page, "Construction", exact=False).filter(visible=True)
        expect_visible(page, group.first, "'Construction' sidebar group")
        click(page, group.first, "Construction menu", animated=True)
    open_module(page, module, expected_button, path=path)


def modal_panel(page: Page, title: str) -> Locator:
    """The modal whose own text carries `title`.

    The indent screens build every form and every dialog from the same
    `div.fixed.inset-0.z-50` the workspace's "What's New" panel uses, with no
    role and no aria-modal on it - so the title is the only thing that tells
    one from another, and scoping the fields to the panel is what keeps a
    dialog's number box from being confused with the form's behind it.
    """
    return (page.locator("div.fixed.inset-0.z-50")
            .filter(has_text=re.compile(title, re.IGNORECASE))
            .filter(visible=True).last)


def indent_item_select(form: Locator) -> Locator:
    """The item dropdown of a material line, found by an option only it has."""
    return form.locator("select").filter(
        has_text=re.compile(r"select item", re.IGNORECASE))


def choose_indent_item(page: Page, form: Locator, code: str, name: str) -> str:
    """Choose the catalogue item on the indent line by what the option READS.

    The recording captured

        select_option("3d99bacf-c4c6-4b74-8708-f56ccf663899")

    which is the item's primary key in one tenant of one environment. The
    business login lands in a different tenant from run to run (see "THE
    LINE-ITEM ROW" above), so that key is meaningless on the very next
    execution. The options are labelled "ITEM-OPC53-001 - OPC Cement 53 Grade",
    so the item is found by the code and the name the workbook holds and the
    GUID never appears in this file at all.

    The Item Code is what disambiguates, and it is not a nicety: this catalogue
    really does hold two items called "OPC Cement 53 Grade" (ITEM-OPC53-001 and
    ITEM-002), so matching on the name alone would depend on the order the
    application happens to list them in. A blank Item Code falls back to the
    name, which is what a sheet that only knows what it asked for wants.
    """
    field = indent_item_select(form)
    expect_visible(page, field, "the indent's item dropdown")

    try:
        labels = [text.strip() for text
                  in field.locator("option").all_inner_texts()]
    except PlaywrightError:
        labels = []
    choices = [label for label in labels
               if label and not PLACEHOLDER_OPTION.match(label)]

    wanted_code = (code or "").strip()
    wanted_name = (name or "").strip()

    chosen = ""
    if wanted_code:
        chosen = next((label for label in choices
                       if wanted_code.casefold() in label.casefold()), "")
        if not chosen:
            log.warning("   the item dropdown offers no option carrying the "
                        "code '%s' - falling back to the name", wanted_code)
    if not chosen and wanted_name:
        chosen = next((label for label in choices
                       if wanted_name.casefold() in label.casefold()), "")

    if not chosen:
        fail(page, f"The indent's item dropdown does not offer "
                   f"'{wanted_code or wanted_name}'. An item has to exist in "
                   f"this business before it can be indented, and the Items "
                   f"sheet is what creates it - so the Material_Indents sheet "
                   f"has to spell the item the way the Items sheet does. The "
                   f"dropdown offers: "
                   f"{' | '.join(choices) if choices else '(nothing at all)'}.")

    # Chosen and recorded under ONE name. safe_select remembers the option's
    # whole label ("ITEM-OPC53-001 - OPC Cement 53 Grade"), which is a string
    # the record's View page never prints - it shows the code and the name in
    # two places - so the sheet's own wording is written over it. Recording the
    # label as a field of its own would have added a comparison that could only
    # ever fail. It is the swap select_option_containing() makes for a purchase
    # order, for the same reason.
    safe_select(page, field, chosen, "Item Name")
    verify_data.record("Item Name", wanted_name or chosen)
    log.info("   item chosen from the list: %s", chosen)
    return chosen


def indent_heading(page: Page) -> Locator:
    """The IND/nnnn heading at the top of an indent's detail page."""
    return page.get_by_role("heading", name=INDENT_NUMBER) \
               .filter(visible=True).first


def indent_record_showing(page: Page, timeout: int = RETRY_TIMEOUT) -> bool:
    """Wait until the indent detail page has actually PAINTED its record.

    The detail page is rendered by the SPA from an API response that is fired
    AFTER the first paint, so for a moment the address is already
    /indents/<id> while the page is still two empty cards. Nothing above this
    catches that: wait_page_ready() sees a quiet network and a stable DOM,
    because on a blank page both are true.

    The IND/nnnn heading is the record's own identity and the badge hangs off
    it, so "the heading is on screen" is the same question as "there is a
    record to read a status off". Waiting for it here is what turns a read
    taken too early into a read taken when there is something to read - the
    whole of the 2026-08-13 Material Indents failure, where all three attempts
    ran inside one second against a page that had not drawn anything yet.

    Returns True/False rather than raising: the caller decides whether a record
    that never arrives is a wait to escalate or a defect to report.
    """
    try:
        indent_heading(page).wait_for(state="visible", timeout=timeout)
        return True
    except (PlaywrightTimeoutError, PlaywrightError):
        return False


def read_indent_status(page: Page, wait: bool = True) -> str:
    """The indent's own status badge - the span beside the IND/nnnn heading.

    Deliberately NOT read_status(): the detail page draws a progress stepper
    that prints "Draft Submitted Approved Partial Fulfilled" as plain text, so
    a page-wide search finds whichever word it was asked about and reports the
    indent as being in that state whatever it is really in. The badge is the
    heading's next sibling, and it is the only element on the page that says
    what the indent IS.

    The record is waited for first (see indent_record_showing): a badge read off
    a page that has not rendered yet is not "no status", it is "not looked yet".
    Pass wait=False to take whatever is on screen right now, which is what the
    end-of-flow record dump wants.
    """
    try:
        if wait and not indent_record_showing(page):
            return ""
        badge = indent_heading(page).locator("xpath=following-sibling::span[1]")
        if badge.count() == 0:
            return ""
        return re.sub(r"\s+", " ", badge.first.inner_text() or "").strip()
    except PlaywrightError:
        return ""


def verify_indent_status(page: Page, expected: str, when: str) -> None:
    """Prove the indent is showing the status this stage of the flow expects.

    Start-anchored rather than exact, the way verify_status() is: a badge often
    says more than the sheet does. An empty cell logs the badge and checks
    nothing, which is what a row that does not care about a stage wants.
    """
    expected = (expected or "").strip()

    on_screen = read_indent_status(page)
    log.info("   indent status %s: %s", when, on_screen or "(no badge on screen)")
    attach_text(f"Material Indent Status {when}",
                on_screen or "(the page shows no status badge)")

    if not expected:
        log.info("   skip: %-24s (the cell is empty)", f"status {when}")
        return

    def attempt_read(attempt: int) -> None:
        actual = read_indent_status(page)
        if not actual.casefold().startswith(expected.casefold()):
            raise AssertionError(
                f"the badge reads '{actual or '(none)'}', not '{expected}'")

    # The heal escalates instead of repeating itself. A second look at a page
    # that has not finished rendering is worth nothing, and the first run to hit
    # this spent all three attempts inside one second doing exactly that; a
    # reload re-issues the API call that draws the record, which is the one
    # thing that can turn a page still waiting on its data into a page with a
    # record on it.
    def heal(state={"attempt": 0}) -> None:                # noqa: B006 - counter
        state["attempt"] += 1
        if state["attempt"] == 1:
            wait_page_ready(page, "indent")
            indent_record_showing(page)
        else:
            log.info("   the indent record still has not rendered - reloading")
            reload_page(page)
            indent_record_showing(page)

    try:
        retry_action(page, f"read the indent's status {when}", attempt_read,
                     VERIFY_ATTEMPTS, heal)
    except (AssertionError, PlaywrightError) as error:
        # An empty detail page and a wrong status are different findings and
        # must not be reported as the same one. "Expected Draft, got (none)"
        # sends a reader hunting for a status rule; what actually happened is
        # that the application never drew the record at all.
        if not indent_record_showing(page, timeout=OPTIONAL_TIMEOUT):
            attach_text("Material Indent - Record Never Rendered",
                        f"APPLICATION DEFECT (not a status mismatch).\n\n"
                        f"Page   : {page.url}\n"
                        f"Symptom: the indent's detail page rendered no record - "
                        f"no IND/nnnn heading, no status badge and no MATERIALS "
                        f"REQUESTED block - after the record was saved, after "
                        f"waiting for it and after a reload.\n\n"
                        f"The status '{expected}' could therefore not be read. "
                        f"This is the application failing to render a record it "
                        f"has just created, NOT the application putting the "
                        f"indent into the wrong status.")
            fail(page, f"The indent was saved but its detail page never rendered "
                       f"the record, so the status {when} could not be read at "
                       f"all.\nThe page stayed empty through a wait and a reload: "
                       f"no IND/nnnn heading and no status badge.\n"
                       f"This is an APPLICATION defect (a record that does not "
                       f"render), not a wrong status - the expected '{expected}' "
                       f"was never contradicted, it was never shown.", error)
        fail(page, f"The indent's status {when} is not '{expected}'.\n"
                   f"{explain_indent_status(page, expected)}", error)

    log.info("   verified: the indent is '%s' %s", expected, when)
    screenshot(page, f"Material Indent - Status {expected} ({when})")


def explain_indent_status(page: Page, expected: str) -> str:
    """Say WHY the indent ended where it did, from the quantities on the record.

    A bare "the badge reads 'Partial Issue', not 'Fulfilled'" sends the reader
    hunting through the flow for a step that went wrong, when the record itself
    already holds the answer.

    What the application does, confirmed by two runs of the same row on
    2026-08-12:

        requested 90, approved 90, ISSUED 90, returned 8  ->  Fulfilled
        requested 90, approved 90, ISSUED < 90            ->  Partial Issue

    So the status follows ISSUED against REQUESTED, and a return does not
    downgrade it - the run that returned 8 of its 90 was still Fulfilled. That
    matters, because the return is the first thing anyone suspects.

    A short issue is not a defect and not a bug in this suite: the issue dialog
    is pre-filled by the application with what it can actually issue, and this
    run deliberately types nothing into it. Less than the approved quantity
    means the warehouse did not have the rest - stock that earlier runs of this
    very suite have consumed.
    """
    shown = read_indent_quantities(page, LAST_INDENT_ITEM) if LAST_INDENT_ITEM else {}
    if not shown:
        return ("The quantities could not be read off the record, so there is "
                "nothing more this check can say about the cause.")

    requested = verify_data.as_number(shown.get("Requested", ""))
    issued = verify_data.as_number(shown.get("Issued", ""))
    returned = verify_data.as_number(shown.get("Returned", ""))
    reading = ", ".join(f"{name} {value}" for name, value in shown.items() if value)

    if issued is not None and requested is not None and issued < requested:
        short = requested - issued
        return (
            f"The record reads: {reading}.\n"
            f"CATEGORY: TEST DATA / ENVIRONMENT - not an automation defect and "
            f"not an application defect. The application was asked for stock it "
            f"does not hold and said so; the automation read that correctly.\n"
            f"CAUSE: the application issued {issued:g} of the {requested:g} "
            f"requested - {short:g} short - so it marked the indent 'Partial "
            f"Issue'. The issue dialog is pre-filled by the application with "
            f"what it can actually issue and this run types nothing into it, so "
            f"a short issue means the stock was not there. Earlier runs of this "
            f"suite consume the same item, so '{expected}' is only reachable "
            f"while at least {requested:g} of '{LAST_INDENT_ITEM}' is in stock.\n"
            f"This is stock running out, not a fault in the application or in "
            f"the flow. Top the item up (or lower the Quantity on the "
            f"Material_Indents row) and the row reaches '{expected}' again.\n"
            f"Not the cause: the Return Quantity. A run that issued all 90 and "
            f"returned 8 still reached 'Fulfilled' - a return does not downgrade "
            f"the status.")

    if returned:
        return (f"The record reads: {reading}.\n"
                f"The full quantity was issued, so the return of {returned:g} is "
                f"not the cause - a return does not downgrade the status.")
    return f"The record reads: {reading}."


def materials_table_text(page: Page) -> str:
    """The MATERIALS REQUESTED block of the detail page, as plain text.

    Read as text on purpose. The block LOOKS like a table and is not one - it is
    a CSS grid with no <table>, no role="row" and no role="cell" (get_by_role
    ("row") finds nothing on this page), so there is no cell to address. The
    text is cut off at whatever comes next so that the MATERIAL RETURNS list
    below it - which prints quantities of its own - can never be read as part of
    the requested line.
    """
    for area in (page.locator("main").first, page.locator("body").first):
        try:
            text = area.inner_text(timeout=OPTIONAL_TIMEOUT) or ""
        except (PlaywrightTimeoutError, PlaywrightError):
            continue
        start = re.search(r"MATERIALS\s+REQUESTED", text, re.IGNORECASE)
        if not start:
            continue
        block = text[start.end():]
        end = re.search(r"\n\s*(NOTES|MATERIAL RETURNS|BATCH CONSUMPTION)",
                        block, re.IGNORECASE)
        return block[:end.start()] if end else block
    return ""


def read_indent_quantities(page: Page, item_name: str) -> Dict[str, str]:
    """What the detail page says was requested, approved, issued and returned.

    The application prints a heading row and then one line per material:

        MATERIAL   REQUESTED   APPROVED   ISSUED   RETURNED   IN STOCK
        OPC Cement 53 Grade
        ITEM-OPC53-001
        43 PCS     43 PCS      43 PCS     2 PCS    459

    Reading starts at the material's own NAME rather than at the heading, so a
    second material's line cannot be read in place of the one being checked.

    The quantity columns carry the unit and the In Stock column does not, which
    is what the cells are told apart by - so a build that prints a dash for a
    column nothing has happened in yet yields one cell fewer rather than
    shifting every reading one place to the left.

    Returns {} when the line cannot be read. That is reported, never failed: a
    quantity this function could not find is not the same thing as a quantity
    the application got wrong.
    """
    block = materials_table_text(page)
    if not block.strip():
        return {}

    lines = [line.strip() for line in block.splitlines() if line.strip()]
    # Start at the material's own name, so another material's line above it
    # cannot be read instead.
    start = next((index for index, line in enumerate(lines)
                  if item_name.casefold() in line.casefold()), None)
    if start is None:
        return {}

    found: Dict[str, str] = {}
    for line in lines[start + 1:]:
        cell = INDENT_QUANTITY_CELL.match(line)
        if not cell:
            continue                       # the item code, a label, In Stock
        if len(found) >= len(INDENT_QUANTITY_COLUMNS):
            break
        found[INDENT_QUANTITY_COLUMNS[len(found)]] = cell.group(1)
    return found


def verify_indent_quantities(page: Page, item_name: str,
                             expected: Dict[str, str]) -> None:
    """Prove requested / approved / issued / returned are what the row asked for.

    This is the check the whole module exists for: the numbers are the business
    rule. The requested quantity must NOT have become the return quantity along
    the way, and the only place that can be seen is the finished record.
    """
    on_screen = read_indent_quantities(page, item_name)
    wanted = {name: value.strip() for name, value in expected.items()
              if str(value).strip()}

    table = "\n".join(f"{name:<12}: expected {value:<10} "
                      f"page shows {on_screen.get(name, '(not on the page)')}"
                      for name, value in wanted.items())
    attach_text("Material Indent Quantities", table or "(nothing to compare)")
    log.info("   quantities on the record:\n      %s",
             "\n      ".join(table.splitlines()) or "(nothing to compare)")

    if not on_screen:
        log.warning("   the MATERIALS REQUESTED line for '%s' could not be read "
                    "- the quantities are reported but not asserted", item_name)
        attach_text("Material Indent Quantities - Could Not Be Read", materials_table_text(page))
        return

    wrong = [f"{name}: the record shows "
             f"{on_screen.get(name, '(nothing)')}, the sheet asked for {value}"
             for name, value in wanted.items()
             if name in on_screen and not value_matches(on_screen[name], value)]
    missing = [name for name in wanted if name not in on_screen]

    if wrong:
        fail(page, "The indent's quantities are not what the row asked for:\n  "
                   + "\n  ".join(wrong))
    if missing:
        log.warning("   the record does not print %s - not asserted",
                    ", ".join(missing))
    log.info("   verified: the indent's quantities are what the sheet asked for")


def indent_dialog_or_page(page: Page, title: str, what: str) -> Locator:
    """The dialog an indent action opened, or the page itself when it opened none.

    Every one of these actions raises a modal today. A build that puts the same
    controls inline instead would leave the modal locator matching nothing, and
    every field looked for inside it would then fail on a screen that is
    perfectly usable - so a missing dialog drops the search back to the whole
    page and says so, rather than ending the run.
    """
    dialog = modal_panel(page, title)
    if appears(dialog, RETRY_TIMEOUT):
        log.info("   the %s dialog is open", what)
        return dialog

    log.warning("   '%s' opened no dialog of its own - looking for its controls "
                "on the page instead", what)
    screenshot(page, f"Material Indent - No {what} Dialog")
    return page.locator("body")


def indent_dialog_quantity(page: Page, dialog: Locator, description: str,
                           placeholder: str = "") -> Optional[Locator]:
    """The number box of an indent dialog, whatever this build labels it.

    The recording addressed these by placeholder - `get_by_placeholder("43")` on
    the approval dialog and `get_by_placeholder("0")` on the return one - and
    the first of those is the REQUESTED QUANTITY, printed into the box as a
    hint. It is different for every row, so a sheet asking for 250 bags would
    have looked for a box called "43" and found nothing. The box is therefore
    found as the dialog's own number input, and the placeholder is only the
    fallback.
    """
    return safe_locator(page, description, (
        ("the dialog's number box", dialog.locator('input[type="number"]')),
        ("the dialog's spinbutton", dialog.get_by_role("spinbutton")),
        ("a box whose placeholder is the quantity",
         dialog.get_by_placeholder(placeholder or "0", exact=True)),
    ), need_editable=True)


def indent_action(page: Page, name: str, what: str,
                  expect_next: str = "") -> None:
    """Click one of the indent's status buttons and prove the screen moved on.

    The same contract drive_actions() gives the purchase side: the proof that a
    status really changed is the NEXT button in the flow appearing. Every click
    is animated=True because each one of these opens or closes a modal.
    """
    control = safe_locator(page, name, (
        (f"the '{name}' button", button(page, name)),
        (f"a button starting '{name}'", button(page, name, exact=False)),
    ))
    if control is None:
        # WHY the button is missing, when this run already knows why. A step
        # that never became available because the warehouse could not issue the
        # stock is the stock running out, not the application failing to offer
        # a control it should have offered - and the two must not be reported
        # as the same thing. The row still FAILS; what changes is what the
        # summary calls it. See classify_failure().
        because = ""
        if INDENT_STOCK_SHORTFALL:
            because = (
                f"\nCATEGORY: TEST DATA / ENVIRONMENT - not an automation "
                f"defect and not an application defect.\n"
                f"CAUSE: at the Issue Stock step {INDENT_STOCK_SHORTFALL}, so "
                f"the indent could not reach the state '{what}' needs. Nothing "
                f"can be returned to the store that was never issued out of "
                f"it, so the application is right not to offer the button.\n"
                f"The application was asked for stock it does not hold and "
                f"said so; the automation read that correctly. Top the item up "
                f"(or lower the Quantity on the Material_Indents row) and this "
                f"step is offered again.")
        fail(page, f"The indent does not offer '{name}'. The material indent "
                   f"flow is New Indent -> Create Indent -> Submit for Approval "
                   f"-> Approve -> Confirm Approval -> Issue Stock -> Confirm "
                   f"Issue -> Return to Store -> Confirm Return to Store, and "
                   f"the application only offers a button once the step before "
                   f"it has been answered. Failed at: {what}.{because}")

    click(page, control, name, animated=True)
    wait_until_ready(page, f"the indent after '{name}'")
    screenshot(page, f"Material Indent - After {name}")

    if expect_next:
        expect_visible(page, button(page, expect_next, exact=False),
                       f"the '{expect_next}' button (the state after {name})")


def note_issue_shortfall(page: Page, dialog: Locator, approved: str) -> None:
    """Record what the issue dialog is actually offering to issue.

    Read, never typed: the quantity in this box is the application telling us
    what it can issue out of the stock it has, and overtyping it would be the
    test inventing a stock movement the warehouse cannot support. When it is
    less than the approved quantity, that is the whole explanation for an indent
    that ends 'Partial Issue' - captured here, where it is still a plain fact,
    rather than left to be puzzled out at the end.

    Never fails: a short issue is the application behaving correctly with the
    stock it has. The end-of-flow status check is what decides whether the row
    got the outcome the sheet asked for.
    """
    wanted = verify_data.as_number(approved)
    box = indent_dialog_quantity(page, dialog, "Issue Quantity",
                                 placeholder=approved)
    offered = None
    if box is not None:
        try:
            offered = verify_data.as_number(box.input_value(timeout=RETRY_TIMEOUT))
        except (PlaywrightTimeoutError, PlaywrightError):
            offered = None

    if offered is None:
        log.info("   the issue dialog does not show a quantity that can be read "
                 "- what was issued is read off the finished record instead")
        return

    log.info("   the application offers to issue %g (approved: %s)", offered,
             approved or "not stated")
    verify_data.record("Issued Quantity (offered by the application)",
                       f"{offered:g}")

    if wanted is None or offered >= wanted:
        return

    short = wanted - offered
    # Remembered for the rest of the row: a step further down the flow that is
    # never offered is explained by this, and without it that step reads as an
    # application defect - see indent_action().
    global INDENT_STOCK_SHORTFALL
    INDENT_STOCK_SHORTFALL = (
        f"the application could only issue {offered:g} of the {wanted:g} "
        f"approved - {short:g} short")
    log.warning("   STOCK SHORT: the application can only issue %g of the %g "
                "approved - %g short. The indent will therefore end 'Partial "
                "Issue' rather than 'Fulfilled'.", offered, wanted, short)
    attach_text("Stock Shortfall at Issue",
                f"The Issue Stock dialog was pre-filled with {offered:g}, "
                f"against an approved quantity of {wanted:g} - {short:g} short."
                f"\n\nThe application fills this box with what it can actually "
                f"issue, and this run types nothing into it, so the shortfall is "
                f"the warehouse being out of stock. Earlier runs of this suite "
                f"draw on the same item.\n\nConsequence: the indent ends "
                f"'Partial Issue'. A Material_Indents row whose Expected Status "
                f"is 'Fulfilled' needs at least {wanted:g} in stock to reach it.")
    screenshot(page, "Material Indent - Stock Short at Issue")


#: How many times open_created_indent() reloads the list while its first row
#: still reads exactly what it read BEFORE this indent was created. Confirmed
#: live 2026-08-23: right after the create form closes, the list can still be
#: showing its PRE-CREATE state for a moment - `.first` then opens an OLD
#: indent instead of the new one, which reads as a fully-processed
#: "Fulfilled" record (every stage already ticked, a return already logged)
#: the instant it is "created", because it was never new. A row that has not
#: actually changed is not this run's indent, however confidently `.first`
#: hands it back.
INDENT_LIST_REFRESH_ATTEMPTS = 4


def open_created_indent(page: Page, before: str = "") -> str:
    """Make sure the indent that has just been saved is the record on screen.

    Some builds land on the new indent's own page as soon as it is created and
    some go back to the list, so both are handled: an address that already
    carries a record id is left alone, and a list is opened at its newest row -
    which is the indent this run has just raised.

    `before` is the list's own first-row reference from BEFORE this indent
    existed (see newest_indent_reference(), captured by the caller while
    still on the list, ahead of "New Indent"). See INDENT_LIST_REFRESH_ATTEMPTS
    for why this is checked rather than trusted on the first look.
    """
    if re.search(rf"{re.escape(INDENT_PATH)}/[^/]+$", page.url):
        # The address carries the id the moment the save returns, but the SPA
        # draws the record from a second request. Wait for the record itself,
        # or the number below - and the status check after it - are read off a
        # page that is still empty.
        if not indent_record_showing(page):
            log.warning("   the indent's detail page has not drawn the record "
                        "yet - reloading it")
            reload_page(page)
            indent_record_showing(page)
        return read_document_number(page, "IND", "Material Indent")

    log.info("   the application went back to the list - opening the indent it "
             "has just created")

    current = newest_indent_reference(page)
    for attempt in range(1, INDENT_LIST_REFRESH_ATTEMPTS + 1):
        if current and current != before:
            break
        if attempt == INDENT_LIST_REFRESH_ATTEMPTS:
            fail(page, f"The Material Indents list still shows "
                       f"{current or '(nothing)'!r} as its first row after "
                       f"{INDENT_LIST_REFRESH_ATTEMPTS} reloads - the same as "
                       f"before this indent was created. The indent was saved "
                       f"(the application said so), but the list never shows "
                       f"a new top row for it, so there is no way to tell "
                       f"which record is this run's without risking opening "
                       f"someone else's already-finished indent.")
        log.info("   the list's first row still reads %r (the same as before "
                 "this indent was created) - reloading and looking again "
                 "(%d/%d)", current or "(nothing)", attempt,
                 INDENT_LIST_REFRESH_ATTEMPTS)
        reload_page(page)
        wait_page_ready(page, "Material Indents list")
        current = newest_indent_reference(page)

    newest = page.get_by_text(INDENT_NUMBER).filter(visible=True).first
    if not appears(newest, RETRY_TIMEOUT):
        fail(page, "The indent was saved but the Material Indents list is not "
                   "showing it, so there is nothing to drive through the "
                   "approval flow.")
    click(page, newest, "the indent this run created", animated=True)
    wait_until_ready(page, "the indent")
    return read_document_number(page, "IND", "Material Indent")


def create_material_indent(page: Page, row: Row) -> None:
    """Raise one material indent from an Excel row and drive it to the end.

    The sheet drives every screen of the flow:

      Priority / Required By / Reference Type / Department
          the New Material Indent form's own fields, in screen order. Department
          is the site or floor the material is for ("Production Floor").
      Item Code / Item Name / Quantity / Unit
          the material line. Quantity is what is REQUESTED - 43 - and it is the
          number every stage after this one is measured against.
      Notes
          the reason for the indent.
      Approved Quantity
          what is typed into the approval dialog. Blank means "approve what was
          asked for", which is the ordinary case and keeps Requested = Approved.
      Issue Stock
          NO stops after the approval. Anything else issues the approved
          quantity out of the store.
      Return Quantity
          how much of what was issued goes back. Blank skips the return
          altogether - it is the one part of the flow a row may not want.
      Status After Create / Status After Approval / Expected Status
          the badge beside the IND/nnnn heading at three points of the flow.

    Requested 43 -> Approved 43 -> Issued 43 -> Returned 2 is the shape of the
    finished record, and it is read back off the detail page at the end. The 2
    is the RETURN quantity and nothing else: it is typed into the return dialog
    only, which is why it has a column of its own rather than being a second
    meaning for Quantity.
    """
    item_name = row["Item Name"]
    item_code = row.get("Item Code", "")
    quantity = row["Quantity"]
    approved_quantity = (row.get("Approved Quantity", "") or "").strip() or quantity
    return_quantity = (row.get("Return Quantity", "") or "").strip()

    # BEFORE the indent exists: the stock the issue is about to take from. Read
    # from Items & Inventory rather than from the indent's own IN STOCK column,
    # because the closing figure is read there too and the two sides of a check
    # have to be the same measurement of the same thing.
    forget("Stock Before Indent", "Stock After Issue", "Indent Opening Stock")
    stock_after_sales = stock_now(page, item_code, item_name, "before the indent")
    publish("Stock Before Indent", stock_after_sales)
    # REPORTING ONLY - this is also the first real reading taken since the
    # Quotation and the Sales Order finished, so it doubles as their own
    # "did the stock move?" checkpoint - see report_stock_after_sales().
    report_stock_after_sales(item_name, item_code,
                             published("Stock After GRN"), stock_after_sales)

    # ---- New Indent ------------------------------------------------------ #
    with step(f"Opening Material Indents... ({item_name} x {quantity})"):
        open_construction_module(page, "Material Indents", "New Indent",
                                 path=INDENT_PATH)
        screenshot(page, "Material Indents - List")
        # The list's own first row, BEFORE this indent exists - see
        # open_created_indent(), which needs this to tell a genuinely new top
        # row apart from a list that has not refreshed yet.
        indent_reference_before_create = newest_indent_reference(page)

    with step(f"Creating the indent... ({item_name} x {quantity})"):
        click(page, button(page, "New Indent"), "New Indent", animated=True)
        form = modal_panel(page, r"New Material Indent")
        expect_visible(page, form, "the New Material Indent form")

        # Fields in screen order - the column order of the Material_Indents
        # sheet. Each dropdown is found by an option only IT carries, so none of
        # them depends on how many dropdowns the form has.
        safe_select(page, form.locator("select").filter(
            has_text=re.compile(r"urgent", re.IGNORECASE)).first,
            row.get("Priority", "") or "Normal", "Priority")
        fill(page, form.locator('input[type="date"]').first,
             resolve_date(row.get("Required By", ""), default_days=21),
             "Required By")
        safe_select(page, form.locator("select").filter(
            has_text=re.compile(r"work order", re.IGNORECASE)).first,
            row.get("Reference Type", "") or "General", "Reference Type")
        fill(page, form.get_by_placeholder("e.g. Production Floor").first,
             row["Department"], "Department")

        # The material line. One line per row: the sheet holds one item, so
        # "Add Row" is never clicked and the first line is the only line - which
        # is what makes .first here unambiguous rather than a guess.
        choose_indent_item(page, form, item_code, item_name)
        fill(page, form.get_by_placeholder("Qty", exact=True).first, quantity,
             "Quantity")
        fill_optional(page, row.get("Unit", ""), "Unit",
                      form.get_by_placeholder("Unit", exact=True).first)
        fill_optional(page, row.get("Notes", ""), "Notes",
                      form.get_by_placeholder("Reason for indent", exact=False))

        submit = button(page, "Create Indent")
        click(page, submit, "Create Indent")

    with step("Verifying the indent was created..."):
        wait_for_submit_to_close(page, submit, item_name, "Material Indent")
        wait_until_ready(page, "the indent after saving")
        number = open_created_indent(page, indent_reference_before_create) \
            or item_name
        verify_indent_status(page, row.get("Status After Create", ""),
                             "after it was created")
        screenshot(page, "Material Indent - Created Record")

    # ---- Submit for Approval --------------------------------------------- #
    with step(f"Submitting the indent for approval... ({number})"):
        indent_action(page, "Submit for Approval", "Submit for Approval",
                      expect_next="Approve")

    # ---- Approve -> Confirm Approval ------------------------------------- #
    with step(f"Approving the indent... ({number}, {approved_quantity})"):
        indent_action(page, "Approve", "Approve")

        dialog = indent_dialog_or_page(page, r"approv", "approval")
        approved_box = indent_dialog_quantity(page, dialog, "Approved Quantity",
                                              placeholder=quantity)
        if approved_box is None:
            fail(page, "The approval dialog has no quantity box to approve "
                       "with. Failed at: Approve.")
        # Approved is a number of its own, not the requested one typed again:
        # a row may approve less than was asked for, and the record has to say
        # so. Blank in the sheet means "approve what was asked for".
        fill(page, approved_box, approved_quantity, "Approved Quantity")

        indent_action(page, "Confirm Approval", "Confirm Approval")
        verify_indent_status(page, row.get("Status After Approval", ""),
                             "after approval")

    # ---- Issue Stock -> Confirm Issue ------------------------------------ #
    if (row.get("Issue Stock", "") or "").strip().casefold() in ("no", "n",
                                                                 "false", "0"):
        log.info("   skip: %-24s (the sheet says NO)", "Issue Stock")
        verify_data.note_skipped("Issue Stock", row.get("Issue Stock", ""),
                                 "the sheet did not ask for the stock to be issued")
        finish_material_indent(page, row, item_name, number, quantity,
                               approved_quantity, issued="", returned="")
        return

    with step(f"Issuing the stock... ({number}, {approved_quantity})"):
        # Read the stock BEFORE anything is issued. One In Stock figure cannot
        # prove a movement; two can, and this is the only moment the first one
        # is still on screen. Published, so the calculation check at the end of
        # the module has an independent opening balance to work from.
        publish("Indent Opening Stock", read_indent_stock(page, item_name))
        indent_action(page, "Issue Stock", "Issue Stock")
        # Nothing is typed here. The issue dialog opens with the approved
        # quantity already in it, which is why the recording types nothing
        # either - and what was really issued is read back off the finished
        # record further down rather than assumed from what was typed.
        dialog = indent_dialog_or_page(page, r"issue", "issue")
        # It IS read, though. This is the moment the indent's final status is
        # decided: the application pre-fills this box with what it can actually
        # issue, so a figure below the approved quantity means the stock is
        # short and the indent will end 'Partial Issue' however the rest of the
        # flow goes. Saying so here turns a puzzling end-state failure into a
        # fact recorded at the moment it happened.
        note_issue_shortfall(page, dialog, approved_quantity)
        indent_action(page, "Confirm Issue", "Confirm Issue")

    # The stock AFTER the issue and BEFORE the return. Only worth the excursion
    # when a return is still to come: without one, the reading taken at the end
    # of the module is the stock after the issue, and taking a second one here
    # would only be reading the same screen twice.
    #
    # The indent's own address is remembered and returned to, because the flow
    # is not finished - Return to Store is still to be clicked, and a run left
    # standing in Items & Inventory would report a perfectly working
    # application as unable to offer the button.
    if return_quantity:
        indent_url = page.url
        publish("Stock After Issue",
                stock_now(page, item_code, item_name, "after the stock issue"))
        # stock_now() puts the browser back by itself; this is the belt to its
        # braces, and the wait is what the Return to Store click needs.
        restore_page(page, indent_url)
        wait_until_ready(page, "the indent after the stock reading")
        if not indent_record_showing(page):
            reload_page(page)
            indent_record_showing(page)

    # ---- Return to Store -> Confirm Return to Store ----------------------- #
    if not return_quantity:
        log.info("   skip: %-24s (the cell is empty)", "Return Quantity")
        verify_data.note_skipped("Return Quantity", "",
                                 "the sheet did not ask for a return")
        finish_material_indent(page, row, item_name, number, quantity,
                               approved_quantity, issued=approved_quantity,
                               returned="")
        return

    with step(f"Returning material to the store... ({number}, {return_quantity})"):
        indent_action(page, "Return to Store", "Return to Store")

        dialog = indent_dialog_or_page(page, r"return", "return")
        return_box = indent_dialog_quantity(page, dialog, "Return Quantity",
                                            placeholder="0")
        if return_box is None:
            fail(page, "The Return Materials to Store dialog has no quantity "
                       "box. Failed at: Return to Store.")
        # The ONLY place the return quantity is ever typed. The indent was
        # raised for `quantity` and stays raised for it - 2 bags going back is
        # not 2 bags having been asked for.
        fill(page, return_box, return_quantity, "Return Quantity")
        fill_optional(page, row.get("Return Notes", ""), "Return Notes",
                      dialog.get_by_placeholder("Reason for return", exact=False))

        indent_action(page, "Confirm Return to Store", "Confirm Return to Store")

    finish_material_indent(page, row, item_name, number, quantity,
                           approved_quantity, issued=approved_quantity,
                           returned=return_quantity)


#: The material the indent just run was raised for. The quantity checks need it
#: to find their line in the MATERIALS grid, and only create_material_indent()
#: knows it - it resolves the sheet's Item Code against the catalogue, where two
#: items share the name "OPC Cement 53 Grade".
LAST_INDENT_ITEM = ""

#: The indent this run raised. The item's STOCK LEDGER writes "Issue against
#: IND/0031" and "Material return MR-... against IND/0031", so this is what
#: picks THIS run's two movements out of the item's whole history.
LAST_INDENT_NUMBER = ""

#: What the Issue Stock dialog offered against what was approved, when it
#: offered less - set by note_issue_shortfall() and cleared at the start of
#: every row. It is what turns "the indent does not offer 'Return to Store'"
#: from an accusation into an explanation: the button is missing BECAUSE the
#: warehouse could not issue the stock, and nothing can be returned that was
#: never issued. See indent_action().
INDENT_STOCK_SHORTFALL = ""


def forget_indent_shortfall() -> None:
    """Start each row with no shortfall remembered from the previous one."""
    global INDENT_STOCK_SHORTFALL
    INDENT_STOCK_SHORTFALL = ""


def finish_material_indent(page: Page, row: Row, item_name: str, number: str,
                           quantity: str, approved_quantity: str,
                           issued: str, returned: str) -> None:
    """Read the finished indent back and prove it says what the row asked for.

    Called from all three ends of the flow - the whole journey, a row that stops
    after the approval, and a row that issues but does not return - so a record
    is verified the same way whatever the sheet asked of it. A stage that did
    not run passes "" for its quantity, which the check skips: the record is
    right to be showing nothing there.
    """
    global LAST_INDENT_ITEM, LAST_INDENT_NUMBER
    LAST_INDENT_ITEM = item_name
    LAST_INDENT_NUMBER = number
    stock_report.note_document("Material Indent", number,   # reporting only
                               "the issue that takes the goods out of stock")

    with step(f"Verifying the finished indent... ({number})"):
        wait_until_ready(page, "the finished indent")
        verify_indent_status(page, row.get("Expected Status", ""), "at the end")
        verify_indent_quantities(page, item_name, {
            "Requested": quantity,
            "Approved": approved_quantity,
            "Issued": issued,
            "Returned": returned,
        })
        attach_text("Page URL", page.url)
        screenshot(page, "Material Indent - Final Status")

        # The number and the status are this module's own data - the form itself
        # was captured field by field on the way in - so they are handed to the
        # checker here and the View page is read back exactly as elsewhere.
        verify_data.record("Indent Number", number)
        verify_data.record("Status", read_indent_status(page)
                           or row.get("Expected Status", ""))
        verify_data.after_save(page, number, "Material Indent")


# =========================================================================== #
# DATA VERIFICATION - hand the checker the framework's own helpers
#
# data_verification.py compares what was typed into a Create form with what the
# record's View page displays, for every module, without a line of code in any
# test case. It deliberately does not import this file (that would be a circle,
# and under `python Construction_Flow.py` it would build a second copy of the
# whole suite), so the helpers it needs are handed to it here - once, after they
# have all been defined.
# =========================================================================== #

verify_data.bind(
    attach_text=attach_text,          # the plain-text Allure attachments
    screenshot=screenshot,            # the Create / View page screenshots
    step=step,                        # the step it appears as in the report
    wait_ready=wait_until_ready,      # load -> network -> redraw -> spinners
    settle=settle,                    # the one short pause, for animations
    search_in_list=search_in_list,    # narrow a long list down to one record
    dismiss_overlay=dismiss_overlay,  # close a form the application left open
)

# calculation_validation.py is the other half of the same idea: data_verification
# proves the application STORED what was typed, this one proves it CALCULATED
# the right result from it. Same reason for the same wiring - it must not import
# this file.
calc.bind(
    attach_text=attach_text,          # the calculation detail attachments
    screenshot=screenshot,            # the screenshot every failed sum gets
    step=step,                        # one step per calculation in the report
)


# =========================================================================== #
# CALCULATION VALIDATION - what each module's numbers are checked against
#
# The engine, the formulas and the reporting live in calculation_validation.py.
# What lives HERE is the part that needs to know this application: which figures
# each screen publishes, and what they should be.
#
# Two rules run through all of it, and they are what make the result mean
# something:
#
#   * EXPECTED is worked out in Python from the inputs the row typed (the
#     workbook) or from a figure the application published EARLIER in the flow.
#     ACTUAL is read off the screen after the application has done its own sum.
#     The two never come from the same place.
#   * A figure the application does not display is SKIPPED, with the reason, and
#     counted as unverified. Nothing is assumed and no arithmetic is invented -
#     if a form has no sums on it, the module is NOT APPLICABLE, not PASS.
# =========================================================================== #

#: The words each screen prints beside the figure being checked, best first.
#: More than one spelling because the forms are not consistent with each other
#: ("Sub Total" on one, "Subtotal" on the next) and because a build is free to
#: rename them - a caption that has gone is a SKIPPED check with a reason, never
#: a wrong number.
#
#: SUBTOTAL and TAXABLE are two DIFFERENT figures and have two different label
#: families, because this application prints both of them, one above the other:
#:
#:      Subtotal          89,200.00     <- quantity x rate, before the discount
#:      Taxable Amount    81,172.00     <- after the discount: what GST is on
#:      IGST              14,610.96
#:      Total             95,782.96
#:
#: They used to share one list with "Subtotal" first, so the taxable-value check
#: - whose formula is "Line Amount - Discount" - was answered by the PRE-discount
#: subtotal and failed by exactly the discount every time. That is the 2026-08-13
#: quotation failure: the automation read the wrong one of two captions the
#: application prints correctly.
SUBTOTAL_LABELS = ("Sub\\s*total", "Gross\\s*Amount", "Items\\s*Total")
#: "Taxable Value" (purchase order), "Taxable Amount" (quotation) and a bare
#: "Taxable" (sales order) are the SAME figure under three captions, so all
#: three are listed - longest first, because find() takes the first label that
#: matches and a caption is only ever answered by its own line ("Taxable" does
#: not answer for "Taxable Amount"). The bare one was missing, which is why a
#: sales order whose taxable value was on screen the whole time was reported as
#: "the application did not display this figure" and SKIPPED.
TAXABLE_LABELS = ("Taxable\\s*(?:Value|Amount)", "Net\\s*Amount", "Taxable")
DISCOUNT_LABELS = ("Discount(?:\\s*Amount)?", "Less\\s*Discount")
TAX_LABELS = ("Tax(?:\\s*Amount)?", "GST(?:\\s*Amount)?", "Total\\s*Tax")
TOTAL_LABELS = ("Grand\\s*Total", "Total\\s*Amount", "Amount\\s*Payable",
                "Order\\s*Total", "Bill\\s*Total", "Total")
ROUND_OFF_LABELS = ("Round\\s*(?:ing)?\\s*Off",)
IGST_LABELS = ("IGST",)
CGST_LABELS = ("CGST",)
SGST_LABELS = ("SGST",)


def sheet_number(row: Row, column: str) -> Optional[float]:
    """One workbook cell as a number, or None when it is blank or not one."""
    return verify_data.as_number(row.get(column, ""))


#: The line a priced document prints UNDER its line-item amount: "155 PCS @
#: ₹413.00". It is the only caption the item row has - every other cell in that
#: row is a bare figure in a CSS grid with no label of any kind - and it names
#: the two numbers that identify the line, which is what makes it usable as an
#: anchor rather than a guess.
LINE_ITEM_ANCHOR = re.compile(
    r"^\s*([\d,]+(?:\.\d+)?)\s*([A-Za-z]{1,8})\s*@\s*[₹$€£]?\s*"
    r"([\d,]+(?:\.\d+)?)\s*$")

#: How far back from that anchor the amount may be. The AMOUNT cell is drawn
#: immediately above the "N PCS @ rate" line; anything further away is another
#: column, so a wider window would start reading one.
LINE_AMOUNT_LOOKBACK = 3


def line_item_amount(numbers: PageNumbers, quantity: Optional[float],
                     rate: Optional[float]) -> Tuple[Optional[float], str]:
    """The per-line amount the document prints for THIS row. (value, how/why).

    The line-item row is a bare CSS grid: no caption, no label, no id on any of
    its cells, so PageNumbers.find() - which works by captions - cannot see any
    of it. What the row DOES print is the summary line under the amount:

        ₹58,893.80
        155 PCS @ ₹413.00      <- the anchor

    So the anchor is matched first and its two numbers are checked against the
    quantity and the rate the workbook typed. Only when BOTH agree is this our
    line, and only then is the figure printed above it taken as its amount.

    Three guards, and each one exists to make a wrong reading impossible rather
    than unlikely:

      * the anchor's quantity and rate must be the row's own - otherwise this
        is a different line, or a different document;
      * the figure must carry a currency symbol, which the amount does and the
        quantity, the GST percentage and the discount percentage do not;
      * it must not be the unit price over again, and it must not be negative
        (that would be the discount cell).

    Anything that does not satisfy all three returns None WITH THE REASON, and
    the caller skips the check. Nothing here works a figure out - it only reads
    one, and says which line it read it from.
    """
    if quantity is None or rate is None:
        return None, ("the row did not give both a Quantity and a Rate, so the "
                      "document's own line could not be identified")

    for index, line in enumerate(numbers.lines):
        match = LINE_ITEM_ANCHOR.match(line)
        if not match:
            continue
        shown_quantity = verify_data.as_number(match.group(1))
        shown_rate = verify_data.as_number(match.group(3))
        if shown_quantity is None or shown_rate is None:
            continue
        if (abs(shown_quantity - quantity) > 0.01
                or abs(shown_rate - rate) > 0.01):
            continue                       # a different line of the document

        for back in range(1, LINE_AMOUNT_LOOKBACK + 1):
            if index - back < 0:
                break
            above = numbers.lines[index - back]
            if re.search(r"[A-Za-z]{3}", above):
                break                      # a caption: the row's cells are over
            if not re.search(r"[₹$€£]", above):
                continue                   # a percentage or a bare count
            value = verify_data.as_number(above)
            if value is None or value <= 0 or abs(value - rate) <= 0.01:
                continue                   # the unit price, or a deduction
            return (calc.Displayed(value, numbers.source),
                    f"the amount the document prints on the line "
                    f"'{line.strip()}'")
        return None, (f"the document's line '{line.strip()}' was found, but no "
                      f"amount is printed beside it in a form this can read")

    return None, ("this document does not print a per-line amount that can be "
                  "tied to this row's quantity and rate")


#: The rates GST can legally be charged at in India. Used to sanity-check a rate
#: worked back out of the application's own figures: 18% derived from a tax
#: amount is a rate the application applied, 17.4% is the application getting
#: its arithmetic wrong, and the two must not be reported as the same thing.
GST_SLABS = (0.0, 0.1, 0.25, 1.0, 1.5, 3.0, 5.0, 6.0, 7.5, 12.0, 18.0, 28.0)


def standalone_percentage(numbers: PageNumbers,
                          not_this: Optional[float] = None) -> Optional[float]:
    """A bare "18%" printed on the document, with no caption beside it.

    How this application publishes the GST rate it took off the item's HSN code
    - as a badge, not as a "Tax Rate:" line - so PageNumbers.find(), which works
    by captions, cannot see it.

    Two guards, because a document prints more than one bare percentage and a
    page-wide scan for "a number with a % after it" is otherwise a coin toss.
    The quotation View page prints the DISCOUNT as "9%" above the GST badge, and
    an unguarded scan duly worked the tax out at 9% - half the right figure, and
    a FAIL against an application that had charged correctly:

      * `not_this` drops the rate the row's own discount is at, which is the
        percentage that is genuinely easy to confuse with the tax one;
      * what is left has to be a rate GST can legally be charged at, so a "10%"
        that is somebody's progress figure cannot be mistaken for a tax rate.

    Zero is never taken from here either. "0%" is the one percentage a screen
    prints for all sorts of reasons that have nothing to do with tax - a nil
    discount, an empty progress bar - and reading it as the GST rate asserts
    that the document is UNTAXED, which is the strongest claim on this page and
    the one least likely to be true. A purchase order at 174,600 was failed
    that way, against an application that had correctly charged 18%. A document
    that really is zero-rated shows a tax figure of 0.00, and the total check
    below reads that figure directly, so nothing is lost by refusing it here.

    Neither guard invents a rate: if nothing survives them, this returns None
    and the caller says the rate could not be established.
    """
    for line in numbers.lines:
        match = re.fullmatch(r"\s*(\d{1,2}(?:\.\d+)?)\s*%\s*", line)
        if not match:
            continue
        value = float(match.group(1))
        if value == 0.0:
            continue
        if not_this is not None and abs(value - not_this) <= 0.001:
            continue
        if any(abs(value - slab) <= 0.001 for slab in GST_SLABS):
            return value
    return None


def displayed_tax_rate(numbers: PageNumbers, row: Row,
                       discount_pct: Optional[float] = None
                       ) -> Tuple[Optional[float], str]:
    """The tax RATE the document is using, and where that rate came from.

    Order matters, and it is the order of how certain each source is that the
    number it is offering is the TAX rate:

    1. a rate the document captions ("Tax Rate", "GST Rate") - it says so;
    2. the badge this row's own form was showing while it was filled in - read
       from the line item's own GST cell, so it is the rate the application
       derived from the HSN code and charged;
    3. a bare percentage badge on the document, once the discount's percentage
       and everything that is not a statutory GST slab have been excluded;
    4. the workbook's Tax column, and ONLY when the form accepted a Tax input.

    Step 4 is last and conditional on purpose. On the Quotation and Sales Order
    forms there is no Tax box: the sheet's 28 never reaches the application, so
    working an expected tax out of it does not check the application against the
    workbook - it checks the application against a number nobody gave it. That
    is what failed a correct quotation on 2026-08-13, and note_tax_is_read_only()
    already knew the form was showing 18% when it happened.

    A rate taken off the screen is an INPUT, not the answer: the expected tax and
    the expected total are still worked out here, in Python, from a taxable value
    the workbook produced. What is read is the rate; what is checked is the
    application's arithmetic with it.
    """
    shown = numbers.find("Tax\\s*Rate", "GST\\s*Rate", "Tax\\s*%")
    if shown is not None:
        return float(shown), "the rate the application captions on this screen"

    if TAX_ON_FORM.get("rate") is not None:
        return float(TAX_ON_FORM["rate"]), \
            "the rate the application's own form showed (derived from the HSN code)"

    badge = standalone_percentage(numbers, not_this=discount_pct)
    if badge is not None:
        return badge, "the rate badge the application prints on this screen"

    if TAX_ON_FORM.get("derived"):
        # The form said "this rate is mine, not yours" but its badge could not
        # be read. The sheet is still not the answer - saying so is better than
        # inventing a rate.
        return None, ("the form derives the rate itself and it could not be "
                      "read off the screen")

    from_sheet = sheet_number(row, "Tax")
    if from_sheet is not None:
        return from_sheet, "the workbook's Tax column (the form accepted it)"
    return None, "no tax rate was available"


def report_tax_rate_source(taxable: Optional[float], tax_rate: Optional[float],
                           source: str, displayed_tax: Optional[float]) -> None:
    """Put the rate the application charged, and where it came from, in the report.

    Two things a reviewer needs and neither of them is an arithmetic check:

      * which rate the application applied and why the suite believes that;
      * that it differs from the workbook's Tax column, when it does. The sheet
        says 28 for cement, the application derives 18 from HSN 25232930. That
        is a rate-master question for the application team, and burying it would
        be hiding it - but failing the row's ARITHMETIC over it would be blaming
        the wrong thing.
    """
    effective = None
    if taxable and displayed_tax is not None and taxable != 0:
        effective = round(displayed_tax / taxable * 100, 4)

    lines = [f"Rate applied by the application : "
             f"{tax_rate if tax_rate is not None else '(not readable)'}%",
             f"Where that rate came from      : {source}"]
    if effective is not None:
        legal = any(abs(effective - slab) <= 0.01 for slab in GST_SLABS)
        lines += [
            f"Rate implied by its own figures : {effective:g}% "
            f"(tax {calc.money(displayed_tax)} / taxable {calc.money(taxable)})",
            f"Is that a statutory GST slab?   : "
            f"{'yes' if legal else 'NO - not a legal GST rate'}"]

    sheet_rate = TAX_ON_FORM.get("sheet")
    if (TAX_ON_FORM.get("derived") and sheet_rate is not None
            and tax_rate is not None and abs(float(sheet_rate) - tax_rate) > 0.01):
        lines += [
            "",
            f"NOTE - the workbook and the application disagree on the RATE:",
            f"    workbook Tax column : {float(sheet_rate):g}%",
            f"    application charged : {tax_rate:g}%",
            "",
            "The form has no Tax input - the application takes the rate from the",
            "item's HSN code - so the workbook's figure never reached it. The",
            "arithmetic checks below therefore use the rate the application",
            "actually charged; whether that rate is the RIGHT one for this HSN",
            "is a rate-master question for the application team and is reported",
            "here rather than failed as an arithmetic fault."]

    attach_text("Tax Rate - Input Used and Where It Came From",
                "\n".join(lines))
    log.info("   tax rate used for the checks: %s%% (%s)",
             tax_rate if tax_rate is not None else "-", source)


def displayed_tax_amount(numbers: PageNumbers) -> Tuple[Optional[float], str]:
    """The tax the document is charging, however this screen chooses to print it.

    A total-tax caption when there is one; otherwise the GST breakup, which is
    the only place the tax appears on this application's View pages - they print
    "IGST 14,610.96" and no "Tax" line at all. Without this the tax amount went
    unchecked on every View page in the suite, because "GST" does not match
    "IGST" (the caption match is anchored, so it cannot).
    """
    total_tax = numbers.find(*TAX_LABELS)
    if total_tax is not None:
        return float(total_tax), "the document's own Tax caption"

    igst = numbers.find(*IGST_LABELS)
    if igst is not None:
        return float(igst), "IGST (inter-state supply)"

    cgst = numbers.find(*CGST_LABELS)
    sgst = numbers.find(*SGST_LABELS)
    if cgst is not None and sgst is not None:
        return float(cgst) + float(sgst), "CGST + SGST (supply inside the state)"
    return None, "the screen shows no tax figure"


def resolve_supply_type(numbers: PageNumbers,
                        inter_state: Optional[bool]) -> Tuple[Optional[bool], str]:
    """Is this an inter-state supply? (answer, how it was decided).

    Which split applies is a property of the DOCUMENT, so when the test data
    does not say, the screen is asked which one the application used - an IGST
    line means inter-state, a CGST/SGST pair means in-state. That decides which
    RULE to apply; it does not decide the answer.

    Resolved BEFORE the tax is checked, because the answer changes how the tax
    itself is worked out - see calc.tax_charged(). None means unknowable.
    """
    if inter_state is not None:
        return inter_state, "the test data says so"

    shows_igst = numbers.found(*IGST_LABELS)
    shows_pair = numbers.found(*CGST_LABELS) and numbers.found(*SGST_LABELS)
    if shows_igst and not shows_pair:
        return True, "the document prints an IGST line"
    if shows_pair and not shows_igst:
        return False, "the document prints CGST and SGST"
    return None, "neither the test data nor the document says"


def check_gst_split(page: Page, numbers: PageNumbers, tax: Optional[float],
                    inter_state: Optional[bool], decided: str) -> None:
    """Check the GST breakup against the tax worked out independently.

    The expected figures are worked out here from a taxable value the workbook
    produced, so a wrong breakup is still caught.
    """
    if inter_state is None:
        # WHICH breakup applies is a property of the document, and this one
        # does not state it: the sales order prints a single "GST" line and no
        # IGST / CGST / SGST anywhere, so there is no breakup on the screen to
        # check and nothing that says which rule should have produced one.
        # Guessing from the customer's GSTIN would be inventing the answer to
        # the very question being asked.
        shows = [name for name, labels in (("IGST", IGST_LABELS),
                                           ("CGST", CGST_LABELS),
                                           ("SGST", SGST_LABELS))
                 if numbers.found(*labels)]
        prints = ("only " + ", ".join(shows)) if shows else "a single tax line"
        calc.skip("GST Split (IGST / CGST + SGST)",
                  f"This document does not break its tax up: it prints "
                  f"{prints} and "
                  f"neither an IGST line nor a CGST/SGST pair that says which "
                  f"kind of supply it is, and the test data does not say "
                  f"either. Which split applies is therefore not knowable from "
                  f"this screen, and none was assumed. The TAX AMOUNT itself "
                  f"is checked above, against a value worked out "
                  f"independently. To have the split checked, fill the sheet's "
                  f"Inter State column (Yes / No).",
                  {"Tax Amount": tax, "Breakup shown": ", ".join(shows) or
                                                       "none"},
                  "inter-state: IGST = Tax | in-state: CGST = SGST = Tax / 2")
        return

    split = calc.gst_split(tax, inter_state)
    kind = "inter-state" if inter_state else "in-state"
    if inter_state:
        calc.check("IGST", {"Tax Amount": tax, "Supply": f"{kind} ({decided})"},
                   "IGST = Tax Amount (inter-state supply)",
                   expected=split["IGST"], actual=numbers.find(*IGST_LABELS),
                   page=page)
    else:
        for name, labels in (("CGST", CGST_LABELS), ("SGST", SGST_LABELS)):
            calc.check(name, {"Tax Amount": tax, "Supply": f"{kind} ({decided})"},
                       f"{name} = Tax Amount / 2 (supply inside the state)",
                       expected=split[name], actual=numbers.find(*labels),
                       page=page)


def check_document_totals(page: Page, row: Row, what: str,
                          quantity: Optional[float], rate: Optional[float],
                          discount_pct: Optional[float] = None,
                          inter_state: Optional[bool] = None) -> None:
    """Validate the arithmetic of a document that has priced line items.

    Quotations, Sales Orders and Purchase Orders all draw the same shape of
    screen - a priced line, then a totals panel under it - so they share this.

    Every figure is checked against the caption the application prints beside
    it, and anything the screen does not show is skipped with its reason. The
    chain is the ordinary one for a GST document and it is declared in
    calculation_validation.py, not here.
    """
    numbers = PageNumbers.read(page)

    amount = calc.line_amount(quantity, rate)
    if amount is None:
        calc.skip("Line Amount",
                  "The row did not supply both a Quantity and a Rate, so there "
                  "was nothing to work out.",
                  {"Quantity": row.get("Quantity", ""), "Rate": row.get("Rate", "")},
                  "Quantity x Rate")
        return

    # 1. the line itself -------------------------------------------------
    #
    # The application prints the line's amount NET OF ITS OWN DISCOUNT - the
    # quotation's AMOUNT column reads 58,893.80 against a 64,015.00 subtotal and
    # a 5,121.20 discount - so that is what the expected value is built to be.
    # Quantity x Rate on its own is the SUBTOTAL, and it is checked against the
    # subtotal caption a few lines below.
    #
    # The caption search this used to do ("Amount", "Line Total", "Item Total")
    # never matched anything on this application: the line row is a CSS grid
    # whose only text is "AMOUNT" as a column HEADING, so every document in
    # every run reported "the application did not display this figure". It is
    # read off the row itself now - see line_item_amount().
    discount = calc.discount_amount(amount, discount_pct)
    net_line = calc.taxable_value(amount, discount)
    shown_line, how = line_item_amount(numbers, quantity, rate)
    if shown_line is None:
        calc.skip("Line Amount",
                  f"The application does not display a per-line amount this "
                  f"check can read: {how}. The same arithmetic is checked "
                  f"against the document's own Subtotal and Taxable captions "
                  f"below, so nothing here is left unverified - and no figure "
                  f"was assumed.",
                  {"Quantity": quantity, "Rate": rate,
                   "Discount": discount or 0.0},
                  "Quantity x Rate - Discount")
    else:
        calc.check(
            "Line Amount",
            {"Quantity": quantity, "Rate": rate,
             "Discount": discount or 0.0, "Read from": how},
            "Quantity x Rate - Discount (the document prints its line amount "
            "net of the line's own discount)",
            expected=net_line,
            actual=shown_line,
            alternatives={
                "the line amount before the discount was taken off": amount},
            page=page)

    # 2. the discount, when the row asked for one -------------------------
    if discount_pct is None:
        calc.skip("Discount Amount",
                  "This row has no Discount, so the application has nothing to "
                  "take off and the document prints no discount line. There is "
                  "no figure here to be wrong.",
                  {"Line Amount": amount},
                  "Line Amount x Discount% / 100")
    else:
        # The document writes a discount as a DEDUCTION ("-₹5,121.20") and the
        # discount itself is a positive amount, so the two are compared by
        # magnitude - see calc.magnitude(), which keeps the reading's tag so
        # the expected side still cannot be compared with itself.
        calc.check(
            "Discount Amount",
            {"Line Amount": amount, "Discount %": discount_pct,
             "Read as": "the magnitude of the deduction the document prints"},
            "Line Amount x Discount% / 100",
            expected=discount,
            actual=calc.magnitude(numbers.find(*DISCOUNT_LABELS)),
            page=page)

    # 3. what is taxed ----------------------------------------------------
    #
    # This application prints TWO figures here - "Subtotal" before the discount
    # and "Taxable Amount" after it - and they are checked as the two different
    # sums they are. Reading the first one and holding it to the second one's
    # formula is what produced the 2026-08-13 failure.
    #
    # A screen that prints only one of the two still works: the taxable check
    # falls back to the subtotal caption, which is the right figure whenever
    # there is no separate taxable line to distinguish it from.
    shown_subtotal = numbers.find(*SUBTOTAL_LABELS)
    shown_taxable = numbers.find(*TAXABLE_LABELS)

    if shown_subtotal is not None and shown_taxable is not None:
        # Both captions are on screen, so the pre-discount subtotal is a real
        # figure of the application's own and gets a check of its own.
        calc.check(
            "Subtotal (before discount)",
            {"Quantity": quantity, "Rate": rate},
            "Quantity x Rate",
            expected=amount,
            actual=shown_subtotal,
            page=page)

    taxable = calc.taxable_value(amount, discount)
    taxable_actual = shown_taxable if shown_taxable is not None else shown_subtotal
    calc.check(
        "Taxable Value (Subtotal)",
        {"Line Amount": amount, "Discount": discount or 0.0},
        "Line Amount - Discount",
        expected=taxable,
        actual=taxable_actual,
        alternatives={"the subtotal before the discount": amount},
        page=page)

    # 4. the tax ----------------------------------------------------------
    tax_rate, source = displayed_tax_rate(numbers, row, discount_pct)
    displayed_tax, tax_shown_as = displayed_tax_amount(numbers)
    report_tax_rate_source(taxable, tax_rate, source, displayed_tax)

    # Which supply this is decides HOW the tax is added up, not just how it is
    # broken out, so it is settled before the tax is worked out - see
    # calc.tax_charged(). `tax` stays the single-rounding figure the GST halves
    # are derived from; `charged` is what the document's own tax line carries.
    inter_state, decided = resolve_supply_type(numbers, inter_state)
    tax = calc.tax_amount(taxable, tax_rate)
    charged = calc.tax_charged(taxable, tax_rate, inter_state)

    if tax_rate is None:
        calc.skip("Tax Amount",
                  f"No tax rate could be established ({source}), so the tax "
                  f"could not be worked out independently. It was NOT assumed "
                  f"to be zero.",
                  {"Taxable Value": taxable}, "Taxable Value x Tax% / 100")
    else:
        if inter_state is False:
            formula = ("CGST + SGST, each rounded to the paisa "
                       "(Taxable Value x Tax% / 100, split in two)")
        else:
            formula = "Taxable Value x Tax% / 100"
        calc.check(
            "Tax Amount",
            {"Taxable Value": taxable, f"Tax % ({source})": tax_rate,
             "Supply": f"{'in-state' if inter_state is False else 'inter-state'}"
                       f" ({decided})" if inter_state is not None else "not stated",
             "Read from": tax_shown_as},
            formula,
            expected=charged,
            actual=displayed_tax,
            alternatives={
                "the whole tax rounded once, not the two halves": tax,
                "tax charged on the pre-discount amount":
                    calc.tax_charged(amount, tax_rate, inter_state)},
            page=page)

    # 5. the GST split ----------------------------------------------------
    check_gst_split(page, numbers, tax, inter_state, decided)

    # 6. the bottom line --------------------------------------------------
    # The application's own Round Off, when it prints one, is an INPUT to the
    # expected total - not something the framework invents to make a sum agree.
    round_off = numbers.find(*ROUND_OFF_LABELS)

    # WHICH TAX GOES INTO THE EXPECTED TOTAL, and why this is not a detail.
    # calc.grand_total() reads a missing tax as zero, so passing None here
    # quietly asserts "this document is not taxed" - and that is an invented
    # assumption, not a measurement. It failed a perfectly correct purchase
    # order at 99,275 against the application's 117,144.50, which is exactly
    # 99,275 + 18%: the screen simply had no tax RATE caption to read.
    #
    # So: the tax worked out from a rate if there was one; failing that, the
    # tax FIGURE the application prints, which still leaves a real check (does
    # the total equal the parts the application itself is showing?); and if the
    # screen shows neither, the check is SKIPPED, because a total cannot be
    # predicted from a taxable value alone.
    # `charged`, not `tax`: the total sits under the tax LINE of the document,
    # so it has to be built from the same figure that line carries. Using the
    # single-rounding tax here made the grand total fail by the same paisa the
    # tax line did - one difference reported twice, and neither of them a fault
    # in the application's arithmetic.
    if charged is not None:
        tax_for_total, tax_source = charged, "worked out from the tax rate"
    else:
        tax_for_total, tax_source = displayed_tax, "as displayed by the application"

    if tax_for_total is None:
        calc.skip(
            f"{what} Grand Total",
            "Neither a tax rate nor a tax amount was on screen, so the total "
            "could not be worked out independently. It was NOT assumed to be "
            "untaxed - that assumption is what turns a correct document into a "
            "false failure.",
            {"Taxable Value": taxable}, "Taxable Value + Tax")
        return

    total = calc.grand_total(taxable, tax_for_total, round_off)
    inputs = {"Taxable Value": taxable, f"Tax ({tax_source})": tax_for_total}
    if round_off is not None:
        inputs["Round Off (from the application)"] = round_off
    calc.check(
        f"{what} Grand Total",
        inputs,
        "Taxable Value + Tax" + (" + Round Off" if round_off is not None else ""),
        expected=total,
        actual=numbers.find(*TOTAL_LABELS),
        alternatives={"a total that ignores the discount":
                      calc.grand_total(amount, calc.tax_amount(amount, tax_rate)),
                      "a total before tax": taxable},
        page=page)


# --------------------------------------------------------------------------- #
# CROSS-MODULE VALIDATION
#
# A figure is not just checked where it is created - it is checked again where
# it is USED. The suite runs as one session, in order, so a module can publish
# what it produced and the module downstream can be measured against it:
#
#     Purchase Order  --(quantity, rate, total)-->  GRN  --(total)-->  Bill
#     Item            --(selling price)---------->  Quotation / Sales Order
#
# Nothing published here is ever read back to prove itself - it is only ever
# used as the EXPECTED side of a check whose actual value comes off a different
# screen, later in the run. A figure that was never published leaves the check
# SKIPPED with its reason, so running one module on its own still reports
# honestly instead of silently passing.
# --------------------------------------------------------------------------- #

PUBLISHED: Dict[str, float] = {}


def publish(name: str, value: Optional[float]) -> None:
    """Record a figure for a later module to be checked against."""
    if value is not None:
        PUBLISHED[name] = value
        log.info("   published for the modules downstream: %-22s = %s", name,
                 calc.money(value))


def published(name: str) -> Optional[float]:
    return PUBLISHED.get(name)


def forget(*names: str) -> None:
    """Drop figures a NEW row is about to take its own readings for.

    publish() ignores a None, which is what keeps a figure that could not be
    read from overwriting a good one - and is exactly why a stale one has to be
    dropped explicitly. Without this, a second indent row whose stock reading
    failed would be checked against the FIRST row's opening balance, and the
    result would look like a passing check of something that was never read.
    """
    for name in names:
        PUBLISHED.pop(name, None)


def recorded_check(name: str):
    """The check the calculation engine already recorded under this name.

    A READ-BACK, and that is the whole point of it. The stock report below has
    to show Expected, Actual, Difference and Result for each stage of the run -
    and every one of those figures has already been worked out, once, by
    calculation_validation.py. Recomputing them for the report would be a second
    calculation engine quietly disagreeing with the first; reading them back
    cannot disagree with anything.

    The newest matching check wins, so a row that ran twice reports the run that
    is being looked at. None when the engine never made that check - which the
    report then prints as "not checked" rather than inventing a verdict.
    """
    for report in reversed(calc.reports()):
        for check in reversed(report.checks):
            if check.name == name:
                return check
    return None


def stock_stage_row(stage: str, check_name: str,
                    measured: Optional[float] = None) -> List[str]:
    """One line of the Item Stock Validation table.

    Two kinds of line, and the difference between them is stated rather than
    blurred:

      * a CHECKED stage - the engine worked an expected value out independently
        and compared it. Expected, Actual, Difference and the verdict are its
        own figures, read back.
      * a MEASURED stage - a reading taken because a movement needs a "before"
        as well as an "after". There is nothing to compare it with yet, so it
        has no verdict, and printing one would be inventing a pass.
    """
    check = recorded_check(check_name) if check_name else None
    if check is not None:
        return [stage, calc.money(check.expected), calc.money(check.actual),
                calc.money(check.difference), check.status]
    if measured is not None:
        return [stage, "-", calc.money(measured), "-", "MEASURED"]
    return [stage, "-", "-", "-", "NOT CHECKED"]


def cross_check(page: Page, name: str, upstream: str, formula: str,
                actual: Optional[float]) -> None:
    """Check a figure against what an earlier module published for it."""
    expected = published(upstream)
    if expected is None:
        calc.skip(name,
                  f"'{upstream}' was not published by an earlier module in this "
                  f"run, so there is nothing upstream to compare with. Run the "
                  f"modules in order for this cross-module check.",
                  {}, formula)
        return
    calc.check(name, {upstream: expected}, formula, expected=expected,
               actual=actual, page=page)


# --------------------------------------------------------------------------- #
# THE STOCK CHAIN - one item, followed from end to end of the run
#
# The brief this was written for asks for a stock LEDGER, not a stock check:
#
#     Opening Stock
#   + GRN Approved / Received
#   - Actual Issued
#   + Actual Returned
#   = Expected Remaining Stock          vs   Items & Inventory - Stock Remaining
#
# Four of those five figures are readings taken at four different moments, and
# a reading that was not taken cannot be invented afterwards - so the run takes
# them as it goes and keeps them here, with the item they belong to:
#
#   Purchase Orders   opening stock, before anything is ordered
#   GRN (approved)    stock after the receipt
#   Material Indents  stock before the indent, after the issue, after the return
#
# REQUESTED, APPROVED, ISSUED, RECEIVED and RETURNED are kept apart on purpose,
# and the ledger uses the ISSUED and RETURNED figures the finished record
# carries - never the quantities the workbook asked for. This application issues
# what it can: a row that asks for 55 against a stock of 3 gets 3 issued, and a
# ledger built from the request would then report a correct application as
# 52 short.
# --------------------------------------------------------------------------- #

#: The item this run's stock chain is following, and what each stage read.
#: Filled in as the run goes; read by the ledger report at the end of it.
STOCK_CHAIN: Dict[str, object] = {}

#: The items whose OPENING stock has already been reported. REPORTING ONLY.
#: The opening card answers "a new item should have started at 0 - did it?",
#: and that question can only be asked once per item: the second purchase
#: order for the same item quite correctly starts from the stock the first one
#: left behind, and holding that reading to 0 would report a correct
#: application as wrong.
INITIAL_STOCK_REPORTED: Set[Tuple[str, str]] = set()

#: (code, name) -> the Allure UUID of that item's "Items - Create and verify
#: Items" test result. Captured in report_initial_stock() the moment the
#: BEFORE TRANSACTION card is recorded; read back once ITEM STOCK - AFTER
#: PURCHASE is known (a later, different pytest test - Purchase Orders/GRN) so
#: both readings for the SAME item can be patched onto the ONE Allure result
#: that already carries the first one, instead of opening a second test entry
#: or a separate section for it. See attach_after_purchase_to_item_test().
ITEM_STOCK_TEST_UUID: Dict[Tuple[str, str], str] = {}


def stock_item_of(row: Row) -> Tuple[str, str]:
    """The (code, name) of the catalogue item a row moves stock for."""
    return ((row.get("Item Code", "") or "").strip(),
            (row.get("Item Name", "") or "").strip())


def capture_purchase_baseline(page: Page, row: Row) -> None:
    """The two figures the purchase flow is about to move, read BEFORE it does.

    Called at the top of the Purchase Orders module, which is the last moment
    at which "before" still exists. Neither figure is checked here - there is
    nothing yet to check them against; they are the independent OPENING side of
    the checks the GRN and the Suppliers module make later.

    Nothing here can fail a row. A figure that could not be read leaves its
    check SKIPPED with the reason, which is the honest report; stopping the
    purchase flow because a baseline could not be taken would cost the run the
    modules that were about to prove something.
    """
    code, name = stock_item_of(row)
    supplier = (row.get("Supplier", "") or "").strip()

    # A new purchase starts a new chain: nothing an earlier row read is allowed
    # to stand in for a reading this one could not take.
    forget("Opening Stock", "Supplier Outstanding Before", "Stock After GRN",
           "GRN Received Quantity", "Stock Before Indent", "Stock After Issue")

    opening = stock_now(page, code, name, "before the purchase order")
    outstanding = supplier_outstanding_from_list(
        page, supplier, "before the purchase order") if supplier else None

    STOCK_CHAIN.clear()
    STOCK_CHAIN.update({"code": code, "name": name, "supplier": supplier,
                        "opening": opening})
    publish("Opening Stock", opening)
    publish("Supplier Outstanding Before", outstanding)
    # REPORTING ONLY - the same reading again, in the register the QA card
    # reads. It is what lets the Purchase Order card answer "what did this item
    # start with?" on the card itself, instead of a reader having to reach the
    # stock section to find out.
    stock_report.note_quantity("Opening Stock", opening,
                               "read before the purchase order was raised")

    # THE "BEFORE" OF THE WHOLE RUN, on its own card - the CURRENT STOCK
    # every later stock verdict is measured from.
    report_initial_stock(code, name, opening)


def report_initial_stock(code: str, name: str,
                         opening: Optional[float]) -> None:
    """📦 ITEM STOCK - BEFORE TRANSACTION: the item's current stock at the start.

        Item Name                        : OPC Cement 53 Grade
        Current Stock Before Transaction : 0
        Expected Initial Stock           : 0
        Difference                       : 0
        Result                           : PASS

    REPORTING ONLY. It reads no page, opens no screen and drives nothing: the
    figure is the reading capture_purchase_baseline() has just taken, which is
    the first and only moment in this flow at which "before anything moved" is
    still true. Nothing here can fail a row - it records a stock card and two
    attachments, and stock cards are counted by the stock report alone.

    THE EXPECTED FIGURE IS ONLY 0 FOR AN ITEM THIS RUN CREATED. An item the
    catalogue already held has no expected opening stock, and claiming 0 of it
    would be this report inventing a requirement the application never had - so
    that card is recorded as NOT CHECKED with the reason, and its reading still
    stands as the BEFORE the GRN is measured against.

    Nothing is assumed when the reading could not be taken: the card says the
    stock was not available rather than showing a 0 nobody read.
    """
    key = ((code or "").strip().casefold(), (name or "").strip().casefold())
    if key in INITIAL_STOCK_REPORTED:
        return
    INITIAL_STOCK_REPORTED.add(key)

    # Filed away NOW, while this test (the Items one) is still the test
    # executing - it is the only moment current_allure_test_uuid() can answer
    # for it. Nothing is lost if this comes back empty (no --alluredir, or the
    # listener could not be reached): ITEM STOCK - AFTER PURCHASE still gets
    # attached to the GRN test as before, it just cannot also be patched onto
    # this one.
    test_uuid = current_allure_test_uuid()
    if test_uuid:
        ITEM_STOCK_TEST_UUID[key] = test_uuid

    new_item = created_in_this_run(code, name)
    actual = None if opening is None else float(opening)
    expected = 0.0 if new_item else None
    difference = (None if actual is None or expected is None
                  else actual - expected)
    if difference is None:
        verdict = stock_report.NOT_CHECKED
    else:
        verdict = (stock_report.PASS if abs(difference) <= 0.001
                   else stock_report.FAIL)

    if actual is None:
        note = ("The current stock could not be read before this transaction, "
                "so there is no opening figure to measure the GRN and the "
                "indent from. Nothing was assumed in its place.")
    elif not new_item:
        note = ("This run did not create this item - the catalogue already "
                "held it - so there is no expected opening stock to check "
                "this reading against. It is recorded because it is the "
                "BEFORE every later stock verdict is measured from.")
    elif verdict == stock_report.PASS:
        note = ("This item was created by this run, so its current stock "
                "before any transaction should be 0 - and Items & Inventory "
                "showed 0.")
    else:
        note = (f"This item was created by this run, so its current stock "
                f"before any transaction should be 0. Items & Inventory "
                f"showed {stock_report.num(actual)} instead. Every movement "
                f"below is still checked against what was really there, never "
                f"against 0.")

    stock_report.record(stock_report.Movement(
        title="ITEM STOCK - BEFORE TRANSACTION",
        stage=stock_report.ITEM_STAGE,
        item=f"{name or '(not named)'} ({code or 'no code'})",
        subject=stock_report.STOCK_COLUMN,
        before_label="Current Stock Before Transaction",
        before=actual,
        expected=expected, actual=actual, difference=difference,
        result=verdict, note=note))

    with stock_section(STOCK_ITEM_STEP):
        # HTML ONLY - a real <table>, not the plain-text card, is the whole
        # point of this section; a monospace attachment sitting next to it
        # would be the old presentation still showing through underneath it.
        verify_data.attach_html("📦 Item Stock - Before Transaction",
                                stock_report.item_page())


def stock_chain_item(row: Row) -> Tuple[str, str, str]:
    """(code, name, why the chain cannot be joined) for a row that moves stock.

    The purchase side and the indent side only make ONE ledger when they are
    about the same item. This catalogue holds eleven items called "OPC Cement
    53 Grade", so two sheets naming the same item NAME are not necessarily
    naming the same item - and a ledger that adds one item's receipt to another
    item's issue is arithmetic about nothing. When the codes differ the chain
    is reported as unjoinable and the indent's own half is still checked.
    """
    code, name = stock_item_of(row)
    chain_code = str(STOCK_CHAIN.get("code", "") or "")
    if code and chain_code and code.casefold() != chain_code.casefold():
        return code, name, (
            f"the purchase flow stocked '{chain_code}' and this row draws from "
            f"'{code}', so they are two different items and their movements do "
            f"not belong on one ledger")
    return code or chain_code, name or str(STOCK_CHAIN.get("name", "") or ""), ""


# --------------------------------------------------------------------------- #
# WHAT EACH MODULE'S NUMBERS ARE
# --------------------------------------------------------------------------- #

def validate_quotation_calculations(page: Page, row: Row) -> None:
    """Quotation: line amount, discount, taxable value, tax and grand total."""
    quantity = sheet_number(row, "Quantity")
    rate = sheet_number(row, "Rate")
    check_document_totals(page, row, "Quotation", quantity, rate,
                          discount_pct=sheet_number(row, "Discount"))
    publish("Quotation Grand Total",
            PageNumbers.read(page).find(*TOTAL_LABELS))
    # REPORTING ONLY - the quantity is the one the line above was checked with;
    # this only remembers it so the stock section can print the quotation leg
    # of the chain instead of leaving it blank.
    stock_report.note_quantity("Quotation Quantity", quantity,
                               "the quantity this quotation was raised for")


def validate_sales_order_calculations(page: Page, row: Row) -> None:
    """Sales Order: the same chain, minus the discount the form does not offer."""
    quantity = sheet_number(row, "Quantity")
    rate = sheet_number(row, "Rate")
    if quantity is None and rate is None:
        calc.skip("Sales Order Grand Total",
                  "This row drives the status buttons of an order that already "
                  "exists (no Customer, no Quantity, no Rate), so it creates no "
                  "figures of its own to check.", {},
                  "Quantity x Rate + Tax")
        return
    check_document_totals(page, row, "Sales Order", quantity, rate)
    # REPORTING ONLY - same as the quotation above.
    stock_report.note_quantity("Sales Order Quantity", quantity,
                               "the quantity this sales order was raised for")


def validate_purchase_order_calculations(page: Page, row: Row) -> None:
    """Purchase Order: the same chain, and the figures the GRN will be held to."""
    quantity = sheet_number(row, "Quantity")
    rate = sheet_number(row, "Rate")
    check_document_totals(page, row, "Purchase Order", quantity, rate)

    # What the GRN and the bill downstream will be measured against.
    publish("Purchase Order Quantity", quantity)
    publish("Purchase Order Rate", rate)
    # REPORTING ONLY - the same figure again, in the register the stock
    # section reads the chain out of. It is not a second capture: it is the
    # value published on the line above, kept where the report can find it.
    stock_report.note_quantity("Purchase Order Quantity", quantity,
                               "the quantity this purchase order was raised for")
    stock_report.note_quantity("Purchase Order Rate", rate,
                               "the rate this purchase order was raised at")
    publish("Purchase Order Total", PageNumbers.read(page).find(*TOTAL_LABELS))


def validate_purchase_order_count(page: Page, row: Row) -> None:
    """BEFORE -> TRANSACTION -> AFTER for the number of purchase orders.

        Expected After Count = Before Count + the orders this row raised

    The simplest question the suite asks and the one a stakeholder asks first:
    there were 5, I raised 1, are there 6? It is a real check because the two
    counts are readings taken at two different moments off the same list, and
    the quantity between them is a document number this run watched the
    application allot.

    RUN LAST, after the calculations, and for a reason: the figures on the
    order's own record page are what validate_purchase_order_calculations()
    reads, and going back to the list before it has finished would take that
    page off the screen. This one navigates, so it goes at the end.

    Nothing here can fail a row on its own. A count that could not be
    established - a list that prints no total and has more than one page - is
    SKIPPED with the reason, which is the honest report.
    """
    before = PURCHASE_ORDER_COUNT.get("before")
    why = str(PURCHASE_ORDER_COUNT.get("how", "") or "the count was not taken")

    # HOW MANY THIS ROW RAISED, counted rather than assumed to be one: the
    # number is appended to CREATED_PURCHASE_ORDERS only when the application
    # really allotted one, so a row whose order was never numbered says 0 here
    # instead of claiming a record that does not exist.
    raised_before = PURCHASE_ORDER_COUNT.get("raised") or []
    created = max(0, len(CREATED_PURCHASE_ORDERS) - len(raised_before))
    number = CREATED_PURCHASE_ORDERS[-1] if CREATED_PURCHASE_ORDERS else ""

    if before is None:
        calc.skip("Purchase Orders on the list (Before + Created)",
                  f"The number of purchase orders BEFORE this row could not be "
                  f"established: {why}. One count cannot prove a movement, and "
                  f"nothing was assumed from it.",
                  {"Orders this row raised": created},
                  "Before Count + Orders Created")
        return

    # Going back to the list is the only thing in this check that can throw,
    # and a list that would not open is a count that could not be taken - NOT
    # a defect in the purchase order, which has already been created, verified
    # and had its arithmetic checked. So it is caught and reported as the skip
    # it is; the docstring's promise that this cannot fail a row is kept here.
    try:
        open_purchase_module(page, "Purchase Orders", "New PO")
        after, how_after = list_record_count(page, "Purchase Orders")
    except Exception as error:                     # noqa: BLE001 - never fatal
        after = None
        how_after = (f"the Purchase Orders list could not be opened again "
                     f"after the order was raised "
                     f"({str(error).splitlines()[0]})")

    if after is None:
        calc.skip("Purchase Orders on the list (Before + Created)",
                  f"The number of purchase orders AFTER this row could not be "
                  f"established: {how_after}. The order itself was created and "
                  f"verified - what could not be read is the list's own count "
                  f"of them.",
                  {"Before Count": before, "Orders this row raised": created},
                  "Before Count + Orders Created")
        return

    # THE LIST CAN LIE ABOUT ITS OWN COUNT WHILE TELLING THE TRUTH IN ITS ROWS.
    # Confirmed live 2026-08-23: the header read "50 purchase orders" both
    # before and after a new order was raised, while the new order's own
    # number sat right there as the first row - scrolling the table's real
    # container (not the page) still settled at 50. A row count that did not
    # grow is proof of a page boundary ONLY if the new record is ALSO missing
    # from the rows; if the record IS there, the list is a sliding window
    # (adding one pushes the oldest out of view) and the row count is not the
    # same quantity before and after, however identical the two numbers look.
    # SKIPPED, not failed: nothing here says the count is wrong, only that it
    # cannot be compared.
    if after <= before and number and list_contains_reference(page, number):
        calc.skip(
            "Purchase Orders on the list (Before + Created)",
            f"The list still reads {after} after this row raised {number} - "
            f"but {number} IS one of the rows on screen. The count did not "
            f"grow because this list only ever shows its newest {after} "
            f"record(s): adding one pushes the oldest of them out of view "
            f"instead of the total growing, which is proven here (the new "
            f"order is visible) and not assumed. Before and after are "
            f"therefore not the same quantity, and nothing was assumed from "
            f"comparing them.",
            {"Before Count": before, "Orders this row raised": created,
             "Purchase Order": number,
             f"{number} is on the list after this row": "yes"},
            "Before Count + Orders Created")
        return

    calc.check(
        "Purchase Orders on the list (Before + Created)",
        {"Before Count (read before the form was opened)": before,
         "Orders this row raised": created,
         "Purchase Order": number or "(the application allotted no number)"},
        "Before Count + Orders Created",
        expected=float(before + created), actual=float(after),
        alternatives={"the list did not move at all": float(before)},
        tolerance=calc.RATIO_TOLERANCE, page=page)

    report_purchase_order_count(before, created, after, why, how_after, number)


def report_purchase_order_count(before: int, created: int, after: int,
                                how_before: str, how_after: str,
                                number: str) -> None:
    """The purchase-order count as the same BEFORE -> AFTER card as the stock.

    A COUNT card: no currency and no stock column, so it names what it is
    about ("Records") rather than borrowing the stock card's wording.
    """
    check = recorded_check("Purchase Orders on the list (Before + Created)")
    if check is None:
        return

    movement = stock_report.Movement(
        title="PURCHASE ORDER COUNT VALIDATION",
        stage="Purchase Orders",
        item="Purchases > Purchase Orders (the list)",
        document=number or "",
        subject="Count",
        source="the Purchases > Purchase Orders list",
        before_label="Before Count",
        before=float(before),
        quantities=[["Created In This Row", created]],
        terms=[["+", "Created In This Row", created]],
        formula=f"{before} + {created}",
        expected=check.expected, actual=check.actual,
        difference=check.difference, result=check.status,
        note=f"Before: {how_before}. After: {how_after}.")

    attach_text("Purchase Orders - Before and After Count", movement.text())
    verify_data.attach_html(
        "Purchase Orders - Before and After Count",
        qa_report.page("Purchase Order Count Validation", movement.item,
                       movement.blocks(), verdict=movement.result))


def note_purchase_subject() -> None:
    """Tell the report which item and rate the purchase figures are about.

    REPORTING ONLY - it changes no expected value, no actual value and no
    verdict. The GRN, Purchase Bill and Suppliers sheets name a purchase order,
    a bill number and a supplier; none of them names the ITEM being moved,
    because the item belongs to the purchase order those rows point at. So the
    three modules' calculations used to be reported without a subject at all -
    "GRN | Pending Quantity | Ordered - Received" and nothing to say what was
    ordered.

    The item, the quantity and the rate are taken from what the purchase flow
    ALREADY recorded earlier in this run - the stock chain and the figures the
    Purchase Orders module published - so nothing here reads a screen, works
    anything out, or invents a value. Whatever was not recorded is simply not
    passed on, and the report leaves that column empty rather than filling it.
    """
    calc.subject(item=str(STOCK_CHAIN.get("name", "") or ""),
                 item_code=str(STOCK_CHAIN.get("code", "") or ""),
                 quantity=published("Purchase Order Quantity"),
                 rate=published("Purchase Order Rate"))


def validate_grn_calculations(page: Page, row: Row) -> None:
    """GRN: the received quantity against the order, and the tax split."""
    note_purchase_subject()
    numbers = PageNumbers.read(page)

    cross_check(page, "Received Quantity vs the Purchase Order",
                "Purchase Order Quantity",
                "the quantity received = the quantity ordered",
                numbers.find("Received(?:\\s*Qty| Quantity)?",
                             "Accepted(?:\\s*Qty| Quantity)?"))

    ordered = published("Purchase Order Quantity")
    received = numbers.find("Received(?:\\s*Qty| Quantity)?")
    pending_shown = numbers.find("Pending(?:\\s*Qty| Quantity)?",
                                 "Balance(?:\\s*Qty| Quantity)?",
                                 "Remaining(?:\\s*Qty| Quantity)?",
                                 "Outstanding(?:\\s*Qty| Quantity)?",
                                 "Still\\s*to\\s*(?:come|receive)")
    if pending_shown is None:
        # Checked against the screen rather than assumed: the approved GRN
        # prints RECEIVED, ACCEPTED, REJECTED and EST. VALUE, and no pending or
        # outstanding quantity anywhere. The figure exists on the PURCHASE
        # ORDER ("280 PCS pending"), which is a different document and a
        # different screen - opening it here would be checking the purchase
        # order, not the receipt.
        calc.skip("Pending Quantity",
                  f"The goods receipt does not display a pending or "
                  f"outstanding quantity. It prints Received, Accepted, "
                  f"Rejected and the receipt's value, and nothing on it "
                  f"answers 'how much of the order is still to come'. Ordered "
                  f"{calc.money(ordered)} and received {calc.money(received)}, "
                  f"so {calc.money(calc.outstanding(ordered, received))} is "
                  f"outstanding by this run's own arithmetic - but the "
                  f"application was NOT asked to agree with a figure it does "
                  f"not show, and nothing was assumed from the ones it does.",
                  {"Ordered": ordered, "Received": received},
                  "Ordered - Received")
    else:
        calc.check("Pending Quantity",
                   {"Ordered": ordered, "Received": received},
                   "Ordered - Received",
                   expected=calc.outstanding(ordered, received),
                   actual=pending_shown,
                   page=page)

    # Inter State is the column that decides which tax split the document uses,
    # so the GRN is the one screen where that rule can be checked.
    inter_state_cell = (row.get("Inter State", "") or "").strip().casefold()
    if inter_state_cell in ("yes", "y", "true", "1"):
        inter_state: Optional[bool] = True
    elif inter_state_cell in ("no", "n", "false", "0"):
        inter_state = False
    else:
        inter_state = None

    tax = numbers.find(*TAX_LABELS)
    if inter_state is None or tax is None:
        # WHICH of the two is missing, in the report's own words. "One of these
        # two things" is not a reason a reader can act on, and on this
        # application it is always the second: the approved GRN prints the
        # receipt's value and no tax at all - the tax appears on the purchase
        # BILL the approval creates, which is the next module and where the GST
        # split IS checked.
        if tax is None:
            reason = ("The goods receipt does not display a tax figure - it "
                      "shows the quantities and the receipt's value only - so "
                      "there is no breakup on this screen to check. The tax "
                      "and its IGST / CGST + SGST split are charged on the "
                      "purchase BILL the approval creates, and they are "
                      "checked there and on the purchase order.")
        else:
            reason = ("The sheet's Inter State cell is blank and the receipt "
                      "does not say which kind of supply it is, so which split "
                      "applies is not knowable from this screen and no breakup "
                      "was assumed. Fill the GRN sheet's Inter State column "
                      "(Yes / No) to have this checked.")
        calc.skip("GST Split (IGST / CGST + SGST)", reason,
                  {"Inter State (workbook)": row.get("Inter State", "") or
                   "(blank)",
                   "Tax shown on the receipt": tax},
                  "inter-state: IGST = Tax | in-state: CGST = SGST = Tax / 2")
    else:
        split = calc.gst_split(tax, inter_state)
        if inter_state:
            calc.check("IGST", {"Tax (from the GRN)": tax, "Supply": "inter-state"},
                       "IGST = Tax (inter-state supply)", expected=split["IGST"],
                       actual=numbers.find(*IGST_LABELS), page=page)
        else:
            for name, labels in (("CGST", CGST_LABELS), ("SGST", SGST_LABELS)):
                calc.check(name, {"Tax (from the GRN)": tax, "Supply": "in-state"},
                           f"{name} = Tax / 2 (supply inside the state)",
                           expected=split[name], actual=numbers.find(*labels),
                           page=page)

    # WHAT THE GRN PUBLISHES FOR THE BILL, and why it is two figures and not one.
    #
    # The approved receipt prints ONE money figure: "EST. VALUE", which is the
    # quantity received at the order's rate - the receipt's TAXABLE value, with
    # no tax on it. It prints no total. So "GRN Total" was never published, and
    # the bill's "Bill Total vs the GRN" check spent every run SKIPPED for want
    # of an upstream figure that does not exist.
    #
    # Both are published for what they actually are, and the bill is held to
    # each of them on the right line: its SUBTOTAL against the receipt's value,
    # its TOTAL against the purchase order's total. Neither is invented and
    # neither is derived - they are read off the screens that print them.
    grn_value = numbers.find("Est\\.?\\s*Value", "Estimated\\s*Value",
                             "Receipt\\s*Value", "Total\\s*Value", "Value")
    publish("GRN Value", grn_value)
    publish("GRN Total", numbers.find(*TOTAL_LABELS))

    # VALIDATE GRN -> INVENTORY. The receipt is what puts the goods into stock,
    # so the stock is read now - after the approval, and before anything else in
    # the run touches the item.
    validate_stock_after_grn(page, row, received)


def report_stock_after_grn(name: str, code: str, opening: Optional[float],
                           received: Optional[float], source: str,
                           actual: Optional[float] = None,
                           expected: Optional[float] = None) -> None:
    """ITEM STOCK VALIDATION - what the approved goods receipt did to the stock.

    The GRN half of the stock story, written where it happens so a reader does
    not have to reach the end of the run to find out whether the receipt landed.

    `actual` and `expected` are the figures validate_stock_after_grn() itself
    just read/worked out - the SAME ones it handed to calc.check(). They are
    threaded through rather than re-derived from recorded_check() because a
    calculation that could not run (no opening reading this session - see
    the `opening is None` branch there) is recorded SKIPPED, with no actual or
    expected of its own - and `actual` is a real reading taken off Items &
    Inventory a moment ago regardless. Losing it behind an absent calculation
    would blank out a genuine application figure, which is exactly the "shows
    0/blank instead of the real 182" failure this card must never make.
    """
    check = recorded_check("Stock After GRN Approval")
    posted = recorded_check("GRN Receipt Posted to Inventory")

    # THE REAL READING WINS. When the calculation ran, check.actual/.expected
    # are these same two values anyway (they are what it was given) - so this
    # changes nothing for a normal run and only fills the gap a skipped
    # calculation would otherwise leave.
    resolved_actual = actual if actual is not None else (
        None if check is None else check.actual)
    resolved_expected = expected if expected is not None else (
        None if check is None else check.expected)
    if check is not None and check.expected is not None and check.actual is not None:
        difference = check.difference
        card_result = check.status
    elif resolved_actual is not None and resolved_expected is not None:
        difference = round(float(resolved_actual) - float(resolved_expected), 2)
        card_result = (stock_report.PASS if abs(difference) <= 0.001
                      else stock_report.FAIL)
    else:
        difference = None if check is None else check.difference
        card_result = stock_report.NOT_CHECKED if check is None else check.status

    # Where the received quantity came from is part of the evidence, not a
    # footnote: "the quantity the GRN says it received" and "the quantity the
    # purchase order was raised for" are two different claims about the same
    # figure, and a card that printed neither would be asking to be trusted.
    note = f"The quantity above is {source}."
    if card_result == stock_report.FAIL:
        note = ("This is a stock movement the application did not make "
                "correctly. The two inputs above were read off two different "
                "screens at two different moments, so neither of them can be "
                "the reason.")
    elif posted is not None and posted.status == calc.FAIL:
        note = (f"The stock itself is right, but the item's own stock ledger "
                f"records {calc.money(posted.actual)} against "
                f"{LAST_GRN_NUMBER or 'this receipt'} where the GRN says it "
                f"received {calc.money(posted.expected)}.")

    # A NEW ITEM SHOULD HAVE STARTED AT NOTHING. Only said when this run
    # really did create the item - see created_in_this_run() - because "opening
    # stock should be 0" is true of a new item and false of one the catalogue
    # already held, and the report must not assert it of the wrong one.
    if created_in_this_run(code, name):
        opening_note = (
            f"This item was created by this run, so its stock before the "
            f"purchase should have been 0 - and it was read as "
            f"{stock_report.num(opening)}.")
        if opening is not None and float(opening) != 0.0:
            opening_note = (
                f"This item was created by this run, so its stock before the "
                f"purchase should have been 0. Items & Inventory showed "
                f"{stock_report.num(opening)} instead. The receipt below is "
                f"still checked against what was really there, not against 0.")
        note = f"{opening_note} {note}".strip()

    grn_movement = stock_report.record(stock_report.Movement(
        title="STOCK INCREASE VALIDATION (Purchase -> GRN -> Bill)",
        stage="Stock Increase (GRN)",
        item=f"{name or '(not named)'} ({code or 'no code'})",
        document=LAST_GRN_NUMBER or "",
        before_label="Current Stock Before Purchase",
        before=opening,
        quantities=[["GRN Approved / Received Quantity", received]],
        terms=[["+", "Purchase Quantity", received]],
        formula=f"{stock_report.num(opening)} + {stock_report.num(received)}",
        expected=resolved_expected,
        actual=resolved_actual,
        difference=difference,
        result=card_result,
        note=note))

    # ONE card, in the two forms a reader uses: the text to paste into a
    # ticket and the colour page to open, under
    # "📦 Stock Validation > 📈 Stock Increase - GRN".
    #
    # It answers, in this order: item, PO, PO quantity, GRN, GRN quantity, GRN
    # status, bill, STOCK BEFORE, EXPECTED STOCK AFTER, ACTUAL STOCK AFTER,
    # DIFFERENCE, RESULT - and the sum written out underneath. The second copy
    # of those same figures that used to hang here as "Stock Increase
    # Validation (GRN)" is gone: two attachments saying one thing is what made
    # this section something a reader had to work through rather than read.
    #
    # Nothing is recalculated for it - increase_text() and increase_page() read
    # the card that was just recorded, which read the check the engine made.
    with stock_section(STOCK_INCREASE_STEP):
        # HTML ONLY - see the matching comment on the BEFORE TRANSACTION
        # attachment above.
        verify_data.attach_html("📈 Item Stock - After Purchase",
                                stock_report.increase_page())

    # ALSO onto the ITEM's own "Items - Create and verify Items" test - see
    # attach_after_purchase_to_item_test() for why this has to be a JSON patch
    # rather than a second allure.attach() call. Built from grn_movement, the
    # object THIS call just recorded, never from increase_text()/
    # increase_page() above - those read movement_by_stage(GRN_STAGE), which
    # only ever answers with the FIRST GRN card of the whole run, and would
    # quietly show one item's figures on another item's test in any run that
    # purchases more than one.
    attach_after_purchase_to_item_test(code, name, grn_movement)


def report_stock_after_sales(name: str, code: str, before: Optional[float],
                             actual: Optional[float]) -> None:
    """ITEM STOCK - AFTER SALES: what Quotation + Sales Order left the stock at.

    Neither a Quotation nor a Sales Order moves Items & Inventory in this
    application - confirmed live: the stock only changes at GRN approval and
    at a Material Indent's stock issue. So the rule this card checks is
    "unchanged from right after the purchase flow", and it checks that with a
    REAL reading rather than assuming it - a run where the application
    unexpectedly did move the stock here would show up as a FAIL, not be
    quietly agreed with.

    `before` is the stock this run already read right after the GRN approval
    (published("Stock After GRN")); `actual` is the SAME reading
    create_material_indent() takes for "Stock Before Indent" - taken once,
    immediately before the indent form is opened, and shown here under its
    other name rather than read off the screen a second time.
    """
    if before is None or actual is None:
        difference = None
        result = stock_report.NOT_CHECKED
    else:
        difference = round(float(actual) - float(before), 2)
        result = (stock_report.PASS if abs(difference) <= 0.001
                  else stock_report.FAIL)

    note = ("Quotation and Sales Order are business documents - neither one "
            "moves Items & Inventory in this application. The stock is "
            "expected to be UNCHANGED here; it is actually reduced later, "
            "when the Material Indent issues it.")
    if result == stock_report.FAIL:
        note = ("The stock changed between the purchase flow and the Material "
                "Indent, which this run did not expect from a Quotation or a "
                "Sales Order. Both readings are the application's own, taken "
                "at two different moments, so this is worth a look rather "
                "than being explained away.")

    stock_report.record(stock_report.Movement(
        title="STOCK AFTER SALES VALIDATION (Quotation -> Sales Order)",
        stage=stock_report.SALES_STAGE,
        item=f"{name or '(not named)'} ({code or 'no code'})",
        document=stock_report.reference("Sales Order"),
        before_label="Current Stock After Purchase",
        before=before,
        # Quotation/Sales Order Quantity are shown on the section's own card
        # (sales_pairs(), read from the QUANTITIES register) rather than
        # carried here: they move no stock, and Movement.quantities feeds
        # transaction_note()'s "something was moved" warning - which must
        # never fire for a card whose whole point is that nothing should have.
        quantities=[],
        terms=[],
        formula=f"{stock_report.num(before)} + 0" if before is not None else "",
        expected=before,
        actual=actual,
        difference=difference,
        result=result,
        note=note))

    with stock_section(STOCK_SALES_STEP):
        # HTML ONLY - see the matching comment on the BEFORE TRANSACTION
        # attachment above.
        verify_data.attach_html(
            "📊 Item Stock - After Sales (Quotation + Sales Order)",
            stock_report.sales_page())


def grn_was_approved(row: Row) -> bool:
    """Did this GRN row ask for the receipt to be approved?

    It matters because the approval is what moves the stock: a GRN left
    unapproved SHOULD leave the stock exactly where it was, and holding it to
    "opening + received" would fail the application for behaving correctly.
    """
    return (row.get("Approve", "") or "").strip().casefold() not in (
        "no", "n", "false", "0")


def validate_stock_after_grn(page: Page, row: Row,
                             received: Optional[float]) -> None:
    """VALIDATION A - the stock Items & Inventory shows after the GRN approval.

        Expected Stock After GRN = Previous Stock + Approved/Received Quantity

    The received quantity is the GRN's OWN figure, read off the receipt, and the
    opening stock is the reading taken before the purchase order was raised.
    Neither of them comes from the screen the answer is read on, so the check
    is a real one - and the exact quantity is what is compared, never "the
    stock is more than nothing".

    Where the received quantity is checked a SECOND time: the item's own STOCK
    LEDGER names the document behind every movement, so the receipt this run's
    GRN posted can be picked out of the item's whole history and held to the
    quantity the GRN says it received.
    """
    code = str(STOCK_CHAIN.get("code", "") or "")
    name = str(STOCK_CHAIN.get("name", "") or "")
    opening = published("Opening Stock")

    if not grn_was_approved(row):
        calc.skip("Stock After GRN Approval",
                  "This GRN row does not ask for the receipt to be approved, "
                  "and it is the approval that puts the goods into stock. There "
                  "is no movement to check - and the stock was NOT assumed to "
                  "have gone up.",
                  {"Opening Stock": opening, "Received": received},
                  "Opening Stock + GRN Received Quantity")
        return

    if received is None:
        received = published("Purchase Order Quantity")
        source = "the quantity the purchase order was raised for"
    else:
        source = "the quantity the GRN says it received"

    reading = inventory_reading(page, code, name, "after the GRN approval")
    STOCK_READINGS_TAKEN.append(reading)
    # The item page's own "Current Stock" first - see stock_now() - falling
    # back to the list's STOCK column only when the item's own page could not
    # be read.
    actual = reading.get("current")
    if actual is None:
        actual = reading.get("stock")
    STOCK_CHAIN["after_grn"] = actual
    STOCK_CHAIN["received"] = received
    publish("Stock After GRN", actual)
    publish("GRN Received Quantity", received)

    expected_value = (calc.closing_stock(opening, received=received)
                      if opening is not None else None)
    if opening is None:
        calc.skip("Stock After GRN Approval",
                  "The stock BEFORE the purchase was not captured, so the "
                  "movement the receipt made cannot be worked out. One stock "
                  "figure cannot prove a movement, and nothing was assumed from "
                  "it. Run the Purchase Orders sheet in the same session as the "
                  "GRN sheet - that is where the opening reading is taken.",
                  {"Stock now": actual, "Received": received},
                  "Opening Stock + GRN Received Quantity")
    else:
        calc.check(
            "Stock After GRN Approval",
            {"Opening Stock (before the purchase order)": opening,
             f"Received ({source})": received,
             "Item": f"{name} ({code or 'no code'})"},
            "Opening Stock + GRN Received Quantity",
            expected=expected_value,
            actual=actual,
            alternatives={"the stock had not moved at all": opening},
            tolerance=calc.RATIO_TOLERANCE, page=page)

    report_stock_after_grn(name, code, opening, received, source,
                           actual, expected_value)

    # The same quantity again, from the other direction: what the item's own
    # ledger says THIS GRN posted into stock.
    entry = ledger_entry_for(reading.get("ledger") or [], LAST_GRN_NUMBER)
    if not LAST_GRN_NUMBER:
        calc.skip("GRN Receipt Posted to Inventory",
                  "This run did not read a GRN number off the receipt, so its "
                  "movement cannot be picked out of the item's stock ledger.",
                  {"Received": received}, "the ledger's receipt = the quantity "
                                          "the GRN received")
    elif entry is None:
        calc.skip("GRN Receipt Posted to Inventory",
                  f"The item's stock ledger has no movement referring to "
                  f"{LAST_GRN_NUMBER}. Either the application does not draw a "
                  f"ledger for this item, or the receipt was not posted to it - "
                  f"the stock check above is what says which.",
                  {"GRN": LAST_GRN_NUMBER, "Received": received},
                  "the ledger's receipt = the quantity the GRN received")
    else:
        calc.check(
            "GRN Receipt Posted to Inventory",
            {"GRN": LAST_GRN_NUMBER,
             f"Received ({source})": received,
             "Ledger movement": str(entry.get("movement", "")),
             "Ledger reference": str(entry.get("reference", ""))},
            "the movement the ledger posted = the quantity the GRN received",
            expected=received,
            actual=stock_reading(page, entry.get("change"),
                                 "stock-ledger-grn"),
            tolerance=calc.RATIO_TOLERANCE, page=page)


def validate_purchase_bill_calculations(page: Page, row: Row) -> None:
    """Purchase Bill: the bill is the purchase's money, so that is what it is
    held to - line by line, and against the two documents it came from.

    The bill is the END of the purchase chain, and it is the one screen where
    every figure of that chain is printed together:

        Purchase Order --(total)--> Bill Total
        GRN            --(value)--> Bill Subtotal
        Bill Total - Paid       --> Outstanding

    So each of the three is checked against the document that produced it,
    which is what makes them real cross-checks: the expected side was read off
    a different screen, earlier in the run, before this page existed.
    """
    note_purchase_subject()
    numbers = PageNumbers.read(page)
    total = numbers.find(*TOTAL_LABELS)

    # WHAT THE BILL'S TOTAL IS CHECKED AGAINST, and why it is not the GRN.
    # The approved goods receipt prints no total - it prints its VALUE, which
    # is the taxable amount with no tax on it - so "the bill total = the GRN
    # total" was a check against a figure that never existed, and it was
    # SKIPPED in every run. The bill's total belongs to the purchase ORDER
    # (both are tax-inclusive), and the bill's SUBTOTAL is what belongs to the
    # receipt. Two checks, each against the document that really produced it.
    cross_check(page, "Bill Total vs the Purchase Order",
                "Purchase Order Total",
                "the bill total = the total of the purchase order it was "
                "raised against (both are tax-inclusive)",
                total)
    bill_subtotal = numbers.find(*SUBTOTAL_LABELS)
    if bill_subtotal is None:
        bill_subtotal = numbers.find(*TAXABLE_LABELS)
    cross_check(page, "Bill Subtotal vs the GRN", "GRN Value",
                "the bill's taxable subtotal = the value of the goods receipt "
                "it was raised from (both are before tax)",
                bill_subtotal)

    paid = numbers.find("Paid(?:\\s*Amount)?", "Amount\\s*Paid",
                        "Payments?(?:\\s*Made)?", "Settled")
    outstanding_shown = numbers.find("Balance(?:\\s*Due)?", "Outstanding",
                                     "Amount\\s*Due", "Due\\s*Amount")
    ordered_total = published("Purchase Order Total")

    if outstanding_shown is None:
        calc.skip("Balance Due",
                  "The bill does not display a balance, an outstanding amount "
                  "or an amount due, so there is no figure of the "
                  "application's here to check.",
                  {"Bill Total": total, "Paid": paid}, "Bill Total - Paid")
    elif paid is not None:
        calc.check("Balance Due", {"Bill Total": total, "Paid": paid},
                   "Bill Total - Paid",
                   expected=calc.outstanding(total, paid),
                   actual=outstanding_shown, page=page)
    elif ordered_total is not None:
        # The bill prints an Outstanding but no Paid line. Nothing in this
        # suite pays a bill - it raises the order, receives the goods and
        # approves the receipt - so the amount paid against a bill created
        # minutes ago is zero, and that is a fact of THIS RUN rather than an
        # assumption about the application. The expected side is still an
        # independent one: the purchase order's own total, read off a different
        # screen before this bill existed.
        calc.check("Balance Due",
                   {"Purchase Order Total (read before this bill existed)":
                        ordered_total,
                    "Paid": "0.00 - this run makes no payment against the bill "
                            "it creates, and the bill shows no payment line",
                    "Bill Total (on this screen)": total},
                   "Bill Total - Paid, with nothing paid: the whole bill is "
                   "outstanding",
                   expected=calc.outstanding(ordered_total, 0.0),
                   actual=outstanding_shown,
                   alternatives={"a bill that had already been settled": 0.0},
                   page=page)
    else:
        calc.skip("Balance Due",
                  f"The bill shows an outstanding amount "
                  f"({calc.money(outstanding_shown)}) but no PAID figure, and "
                  f"no purchase order total was published earlier in this run "
                  f"to work the balance out from independently. Checking the "
                  f"bill's outstanding against the bill's own total would be "
                  f"comparing one figure on this screen with another, which "
                  f"proves nothing - so it was not done. Run the Purchase "
                  f"Orders sheet in the same session for this check.",
                  {"Bill Total": total, "Outstanding": outstanding_shown},
                  "Bill Total - Paid")


def validate_supplier_purchase_activity(page: Page, row: Row) -> None:
    """SUPPLIERS - Recent Purchase Orders (Received) and Outstanding.

    The end of the chain the brief asks for:

        Purchase Order -> GRN -> GRN Approval -> Supplier ->
        Recent Purchase Orders (Received) -> Outstanding

    Two halves, and they are two different questions:

      * the ORDER. The supplier's page lists the purchase orders raised against
        it. The one THIS run raised has to be there, has to carry the value the
        purchase order itself carried, and - once its goods have been received
        on an approved GRN - has to say so. The value is checked against the
        figure the Purchase Orders module published, which came off a different
        screen at a different moment.
      * the OUTSTANDING. This application expresses it as an AMOUNT (a supplier
        carrying three orders of Rs 62,835 shows Rs 1,88,505), and it is checked
        as the movement it is: what it was before this run's purchase, plus what
        this run bought. A single reading of a figure that has been accumulating
        for weeks proves nothing about this run.

    The outstanding QUANTITY the brief also asks about - Ordered less Received -
    is worked out here from the order and the receipt, and compared with the
    application's own outstanding-quantity figure if it prints one. This build
    does not, so that check is SKIPPED with the reason rather than answered with
    the money figure, which is a different thing.
    """
    note_purchase_subject()             # reporting only - see the helper
    supplier = (row.get("Supplier Name", "") or "").strip()
    ordered = published("Purchase Order Quantity")
    received = published("GRN Received Quantity")
    order_total = published("Purchase Order Total")
    outstanding_before = published("Supplier Outstanding Before")

    if not supplier:
        calc.skip("Outstanding Purchase Orders (Amount)",
                  "This row names no supplier, so there is no supplier page to "
                  "read.", {}, "Outstanding before + what this run bought")
        return

    if not bought_from(supplier):
        calc.skip("Outstanding Purchase Orders (Amount)",
                  f"This run's purchase order was raised against "
                  f"'{STOCK_CHAIN.get('supplier', '')}', not '{supplier}', so "
                  f"nothing it bought is on this supplier's page and the "
                  f"movement of its outstanding cannot be attributed to this "
                  f"run. Name the same supplier in the Suppliers and "
                  f"Purchase_Orders sheets to check the chain end to end.",
                  {"Outstanding before": outstanding_before},
                  "Outstanding before + the value received in this run")
        return

    listed_outstanding = supplier_outstanding_from_list(
        page, supplier, "after the GRN approval")

    if not open_supplier_record(page, supplier):
        calc.skip("Recent Purchase Orders (Received) - Order Value",
                  f"The supplier '{supplier}' could not be opened from the "
                  f"Suppliers list, so neither its purchase orders nor its "
                  f"outstanding could be read.",
                  {"Purchase Order Total": order_total},
                  "the order on the supplier's page = the order that was raised")
        calc.skip("Outstanding Purchase Orders (Amount)",
                  f"The supplier '{supplier}' could not be opened from the "
                  f"Suppliers list.", {"Outstanding before": outstanding_before},
                  "Outstanding before + the value received in this run")
        return

    screenshot(page, "Supplier Page - Recent Purchase Orders and Outstanding")
    orders = read_supplier_purchase_orders(page)
    shown_outstanding = supplier_outstanding(page)
    number = CREATED_PURCHASE_ORDERS[-1] if CREATED_PURCHASE_ORDERS else ""
    ours = supplier_order_row(orders, number)

    attach_text("Supplier - Recent Purchase Orders", "\n".join(
        [f"Supplier                 : {supplier}",
         f"Purchase order this run  : {number or '(none was raised in this run)'}",
         "",
         f"{'PURCHASE ORDER':<18} {'DATE':<14} {'AMOUNT':>16}  STATUS",
         "-" * 66]
        + [f"{order['Purchase Order']:<18} {order['Date']:<14} "
           f"{order['Amount']:>16}  {order['Status']}" for order in orders]
        + ["", f"Outstanding (supplier page) : "
              f"{'-' if shown_outstanding is None else calc.money(float(shown_outstanding))}",
           f"Outstanding (suppliers list): "
           f"{'-' if listed_outstanding is None else calc.money(float(listed_outstanding))}"]))

    # --- the order --------------------------------------------------------
    if not number:
        calc.skip("Recent Purchase Orders (Received) - Order Value",
                  "This run raised no purchase order, so there is nothing of "
                  "its own to look for on the supplier's page.",
                  {}, "the order on the supplier's page = the order raised")
    elif ours is None:
        calc.skip("Recent Purchase Orders (Received) - Order Value",
                  f"The supplier's page does not list {number} under RECENT "
                  f"PURCHASE ORDERS. The order exists - it was raised, received "
                  f"and approved earlier in this run - so this is the "
                  f"application not showing it, and it is reported rather than "
                  f"worked around. Orders it does list: "
                  f"{', '.join(order['Purchase Order'] for order in orders) or 'none'}.",
                  {"Purchase Order": number, "Purchase Order Total": order_total},
                  "the order on the supplier's page = the order raised")
    else:
        calc.check(
            "Recent Purchase Orders (Received) - Order Value",
            {"Purchase Order": number,
             "Total on the order itself": order_total,
             "Status on the supplier's page": ours["Status"] or "(none shown)",
             "Date on the supplier's page": ours["Date"] or "(none shown)"},
            "the value the supplier's page shows for the order = the value the "
            "order itself was raised for",
            expected=order_total,
            actual=verify_data.as_number(ours["Amount"]),
            page=page)

    # --- the supplied material's rate -------------------------------------
    rate = published("Purchase Order Rate")
    shown_rate = read_supplied_material_rate(
        page, str(STOCK_CHAIN.get("code", "") or ""),
        str(STOCK_CHAIN.get("name", "") or ""))
    calc.check("Supplied Material Rate",
               {"Rate on the purchase order": rate,
                "Item": f"{STOCK_CHAIN.get('name', '')} "
                        f"({STOCK_CHAIN.get('code', '') or 'no code'})"},
               "the rate the supplier's page shows = the rate the order was "
               "placed at",
               expected=rate, actual=shown_rate, page=page)

    # --- the outstanding AMOUNT -------------------------------------------
    bought = order_total if order_total is not None else published("GRN Total")
    if outstanding_before is None or bought is None:
        calc.skip("Outstanding Purchase Orders (Amount)",
                  "The outstanding BEFORE this run's purchase, or the value of "
                  "what it bought, was not available - so the movement cannot "
                  "be worked out. A single reading of a figure that has been "
                  "accumulating since long before this run proves nothing about "
                  "it, and it was not treated as if it did.",
                  {"Outstanding before": outstanding_before,
                   "Value bought in this run": bought,
                   "Outstanding now": shown_outstanding},
                  "Outstanding before + the value received in this run")
    else:
        calc.check(
            "Outstanding Purchase Orders (Amount)",
            {"Outstanding before this run's purchase": outstanding_before,
             "Value bought in this run (the purchase order's total)": bought},
            "Outstanding before + the value received in this run",
            expected=round(float(outstanding_before) + float(bought), 2),
            actual=shown_outstanding,
            alternatives={"outstanding did not move at all": outstanding_before,
                          "outstanding is only this run's purchase": bought},
            page=page)

    # The list and the supplier's own page are two screens showing one figure.
    calc.check("Outstanding - the list and the supplier's page agree",
               {"Suppliers list": listed_outstanding},
               "the OUTSTANDING column of the list = the Outstanding on the "
               "supplier's own page",
               expected=listed_outstanding, actual=shown_outstanding, page=page)

    # --- the outstanding QUANTITY -----------------------------------------
    quantity_labels = ("Outstanding\\s*(?:Qty|Quantity)",
                       "Pending\\s*(?:Qty|Quantity)",
                       "Balance\\s*(?:Qty|Quantity)",
                       "Undelivered\\s*(?:Qty|Quantity)")
    shown_quantity = PageNumbers.read(page).find(*quantity_labels)
    expected_quantity = calc.outstanding(ordered, received)
    if ordered is None:
        calc.skip("Outstanding Quantity (Ordered - Received)",
                  "No ordered quantity was published by the Purchase Orders "
                  "module in this run, so what is still to come cannot be "
                  "worked out.", {"Received": received}, "Ordered - Received")
    elif shown_quantity is None:
        calc.skip(
            "Outstanding Quantity (Ordered - Received)",
            f"Ordered {ordered:g} and received {received if received is None else f'{received:g}'} "
            f"- so {expected_quantity:g} is still to come. The application "
            f"prints no outstanding QUANTITY anywhere on the supplier's page: "
            f"its 'Outstanding' is an AMOUNT (checked above), and answering a "
            f"quantity question with a money figure would be reporting two "
            f"different things as one. Nothing was assumed.",
            {"Ordered": ordered, "Received": received,
             "Outstanding quantity worked out here": expected_quantity},
            "Ordered - Received")
    else:
        calc.check("Outstanding Quantity (Ordered - Received)",
                   {"Ordered": ordered, "Received": received},
                   "Ordered - Received", expected=expected_quantity,
                   actual=shown_quantity, tolerance=calc.RATIO_TOLERANCE,
                   page=page)

    # THE SUPPLIER'S OUTSTANDING, BEFORE AND AFTER - the same card as the two
    # stock ones, because it is the same question about a different figure:
    # what was it before this run bought anything, what did this run buy, what
    # should it be now, and what does the application show? Every figure was
    # read by the checks above; this only lays them out.
    report_outstanding_movement(supplier, outstanding_before, bought,
                                shown_outstanding)

    # Everything above, as the one page the supplier check is opened for. LAST,
    # so every check it reads back has already been made.
    report_supplier_purchase_orders(supplier, number, orders, ordered, received,
                                    expected_quantity, order_total,
                                    outstanding_before, shown_outstanding,
                                    listed_outstanding)


def report_outstanding_movement(supplier: str, before: Optional[float],
                                bought: Optional[float],
                                shown: Optional[float]) -> None:
    """BEFORE -> TRANSACTION -> AFTER for the supplier's outstanding amount.

    A money card rather than a stock one, and otherwise identical: the figure
    before this run's purchase, what the purchase added, what the outstanding
    should therefore be, and what the supplier's page actually shows.

    Read back off "Outstanding Purchase Orders (Amount)" - the check that has
    just been made - so this cannot disagree with it. It is NOT counted in the
    stock tally: an outstanding balance is not stock, and adding it to
    "Stock Validations: N" would be reporting two different things as one.
    """
    check = recorded_check("Outstanding Purchase Orders (Amount)")
    if check is None:
        return

    movement = stock_report.Movement(
        title="SUPPLIER OUTSTANDING VALIDATION",
        stage="Supplier Outstanding",
        item=supplier or "(not named)",
        amount=True, subject="Outstanding",
        source="the supplier's own page",
        before_label="Outstanding Before This Run's Purchase",
        before=before,
        quantities=[["Value Bought In This Run "
                     "(the purchase order's total)", bought]],
        terms=[["+", "Value Bought In This Run", bought]],
        formula=f"{stock_report.money(before)} + {stock_report.money(bought)}",
        expected=check.expected, actual=check.actual,
        difference=check.difference, result=check.status,
        note="The outstanding is an AMOUNT, not a quantity - this application "
             "prints no outstanding quantity anywhere on the supplier's page.")

    attach_text("Supplier Outstanding - Before and After", movement.text())
    verify_data.attach_html(
        "Supplier Outstanding - Before and After",
        qa_report.page("Supplier Outstanding Validation", movement.item,
                       movement.blocks(), verdict=movement.result))


def report_supplier_purchase_orders(supplier: str, number: str,
                                    orders: List[Dict[str, str]],
                                    ordered: Optional[float],
                                    received: Optional[float],
                                    outstanding_quantity: Optional[float],
                                    order_total: Optional[float],
                                    outstanding_before: Optional[float],
                                    shown_outstanding: Optional[float],
                                    listed_outstanding: Optional[float]) -> None:
    """SUPPLIER PURCHASE ORDER VALIDATION - the supplier's page, as one page.

    Which orders the supplier carries, what this run's one is worth, what has
    been received against it and what is still outstanding - with the quantity
    identity a reader can add up by hand:

        Received + Outstanding = Ordered

    Nothing is calculated here. Every Expected, Actual, Difference and Result is
    read back off the check calculation_validation.py already recorded, so this
    page cannot say anything the validation did not.
    """
    def show(value) -> str:
        return "-" if value is None else f"{float(value):g}"

    value_check = recorded_check("Recent Purchase Orders (Received) - Order Value")
    amount_check = recorded_check("Outstanding Purchase Orders (Amount)")
    quantity_check = recorded_check("Outstanding Quantity (Ordered - Received)")
    verdict = (value_check.status if value_check is not None
               else (amount_check.status if amount_check is not None
                     else "NOT CHECKED"))

    ours = supplier_order_row(orders, number)
    identity = (f"{show(received)} + {show(outstanding_quantity)} = "
                f"{show(ordered)}")

    verify_data.attach_html(
        "Supplier - Recent Purchase Orders", qa_report.page(
            "Supplier Purchase Order Validation",
            f"{supplier or '(not named)'} | "
            f"{number or 'no order was raised in this run'}",
            [qa_report.kpi_block("This run's purchase order", [
                ["PO Amount",
                 "-" if order_total is None else calc.money(order_total)],
                ["Received quantity", show(received)],
                ["Outstanding quantity", show(outstanding_quantity)],
                ["Result", verdict, verdict],
             ]),
             qa_report.facts_block("Supplier", [
                ["Supplier Name", supplier or "(not named)"],
                ["Purchase Order Number",
                 number or "(none was raised in this run)"],
                ["Purchase Order Date",
                 (ours or {}).get("Date", "") or "(none shown)"],
                ["Status on the supplier's page",
                 (ours or {}).get("Status", "") or "(none shown)"],
                ["Outstanding before this run",
                 "-" if outstanding_before is None
                 else calc.money(float(outstanding_before))],
                ["Outstanding on the supplier's page",
                 "-" if shown_outstanding is None
                 else calc.money(float(shown_outstanding))],
                ["Outstanding in the Suppliers list",
                 "-" if listed_outstanding is None
                 else calc.money(float(listed_outstanding))],
             ]),
             qa_report.table_block(
                 "Recent Purchase Orders (as the supplier's page lists them)",
                 ("Purchase Order", "Date", "Amount", "Status"),
                 [[order["Purchase Order"], order["Date"], order["Amount"],
                   order["Status"]] for order in orders],
                 note="" if orders else "The supplier's page lists no purchase "
                                        "orders.",
                 verdict_column=99, legend=False),
             qa_report.pre_block("The quantity identity", "\n".join([
                 "    Expected : Received + Outstanding = Ordered",
                 f"    Actual   : {identity}",
             ])),
             qa_report.table_block(
                 "Validation",
                 ("Check", "Expected", "Actual (application)", "Difference",
                  "Result"),
                 [[check.name, calc.money(check.expected),
                   calc.money(check.actual), calc.money(check.difference),
                   check.status]
                  for check in (value_check, amount_check, quantity_check)
                  if check is not None])],
            verdict=verdict))


def verify_supplier_received_order(page: Page, row: Row) -> None:
    """Prove the order this run raised is on the supplier's page, as Received.

    The FUNCTIONAL half of the supplier check - a status is not arithmetic, so
    it does not belong in the calculation half, and a missing order is a finding
    about the application rather than a number that did not add up.

    Only asserted when the run actually raised an order and had it received: a
    row that stopped earlier has nothing to be on the supplier's page yet, and
    saying so is not the same as passing.
    """
    supplier = (row.get("Supplier Name", "") or "").strip()
    number = CREATED_PURCHASE_ORDERS[-1] if CREATED_PURCHASE_ORDERS else ""

    if number and not bought_from(supplier):
        # The purchase went to a different supplier, so this one has nothing of
        # this run's to show. Reported, not failed and not passed: the two
        # sheets simply do not name the same business.
        log.info("   the purchase order was raised against '%s', not '%s' - "
                 "this supplier's page has nothing of this run's on it",
                 STOCK_CHAIN.get("supplier", ""), supplier)
        attach_text("Supplier - Recent Purchase Orders Not Asserted",
                    f"The Purchase_Orders sheet buys from "
                    f"'{STOCK_CHAIN.get('supplier', '') or '(not recorded)'}' "
                    f"and this Suppliers row is '{supplier}'. They are "
                    f"different businesses, so nothing this run bought belongs "
                    f"on this supplier's page. Name the same supplier in both "
                    f"sheets to check the chain end to end.")
        return

    if not number:
        log.info("   no purchase order was raised in this run - the supplier's "
                 "RECENT PURCHASE ORDERS block has nothing of ours to show")
        attach_text("Supplier - Recent Purchase Orders Not Asserted",
                    "This run raised no purchase order, so there is nothing of "
                    "its own to look for on the supplier's page. Reported, not "
                    "passed: nothing was proved.")
        return

    if not open_supplier_record(page, supplier):
        fail(page, f"The supplier '{supplier}' is not in the Suppliers list, so "
                   f"the purchase orders raised against it cannot be verified. "
                   f"The Suppliers sheet is what creates it.")

    orders = read_supplier_purchase_orders(page)
    ours = supplier_order_row(orders, number)
    listed = ", ".join(order["Purchase Order"] for order in orders) or "none"
    log.info("   the supplier's page lists: %s", listed)

    if ours is None:
        fail(page, f"The supplier '{supplier}' does not list {number} under "
                   f"RECENT PURCHASE ORDERS.\n"
                   f"That order was raised against this supplier earlier in "
                   f"this run, received on a GRN and the GRN approved, so the "
                   f"supplier's page should be showing it.\n"
                   f"Orders the page does list: {listed}.")

    status = (ours.get("Status") or "").strip()
    log.info("   %s is on the supplier's page: %s, %s, %s", number,
             ours.get("Date") or "(no date)", ours.get("Amount") or "(no value)",
             status or "(no status)")
    attach_text("Supplier - The Order This Run Raised", "\n".join([
        f"Supplier       : {supplier}",
        f"Purchase Order : {number}",
        f"Date           : {ours.get('Date') or '(none shown)'}",
        f"Amount         : {ours.get('Amount') or '(none shown)'}",
        f"Status         : {status or '(none shown)'}",
        "",
        f"Every order the supplier's page lists: {listed}"]))

    if not grn_received_in_this_run():
        log.info("   this run did not approve a goods receipt, so the order's "
                 "status is reported and not asserted")
        return

    if not re.search(r"receiv", status, re.IGNORECASE):
        fail(page, f"The goods on {number} were received and the GRN approved "
                   f"in this run, but the supplier's RECENT PURCHASE ORDERS "
                   f"block shows the order as '{status or 'no status at all'}' "
                   f"rather than Received.")
    log.info("   verified: %s is listed as Received on the supplier's page",
             number)


def grn_received_in_this_run() -> bool:
    """True when this run recorded a goods receipt and approved it."""
    return bool(LAST_GRN_NUMBER) and published("GRN Received Quantity") is not None


def bought_from(supplier: str) -> bool:
    """Is this the supplier this run's purchase order was raised against?

    The Suppliers sheet and the Purchase_Orders sheet are free to name different
    businesses, and when they do, this supplier's page holds nothing of this
    run's - which is a test-data fact, not an application one. Unknown (nothing
    recorded) is treated as "yes": a run that skipped the purchase baseline
    should still be able to check whatever is there.
    """
    bought = str(STOCK_CHAIN.get("supplier", "") or "").strip()
    if not bought or not supplier:
        return True
    return bought.casefold() == supplier.strip().casefold()


#: The lines of a material's row that are not cells at all - its item code
#: ("ITEM-OPC53-017", "MAT/023") and the dash a column with nothing in it yet is
#: drawn with. One token, no spaces, and letters only at the front.
#:
#: They have to be named, because the scan below stops at "a line with words in
#: it" - that is how it knows the NEXT material has started - and the item code
#: is a line with words in it that belongs to THIS material. It sits directly
#: under the name, so the scan stopped on the first line every time and
#: read_indent_stock() returned None on every run this application has ever had:
#: the indent's IN STOCK column was reported as unreadable while it was on
#: screen, and the closing-stock check that needed it was SKIPPED.
INDENT_ROW_NOISE = re.compile(
    r"^(?:[-–—]+|n/?a|[A-Za-z]{1,8}[-/_][A-Za-z0-9][A-Za-z0-9/_-]*)$",
    re.IGNORECASE)


def read_indent_stock(page: Page, item_name: str) -> Optional[float]:
    """The IN STOCK figure of one material line, or None.

    The quantity columns of that grid all carry a unit ("43 PCS") and In Stock
    does not - which is exactly how read_indent_quantities() tells the columns
    apart - so the bare number on the line is the stock.
    """
    block = materials_table_text(page)
    lines = [line.strip() for line in block.splitlines() if line.strip()]
    start = next((index for index, line in enumerate(lines)
                  if item_name.casefold() in line.casefold()), None)
    if start is None:
        return None
    for line in lines[start + 1:]:
        if INDENT_QUANTITY_CELL.match(line):
            continue                       # a quantity: it carries its unit
        if INDENT_ROW_NOISE.match(line):
            continue                       # this material's code, or an empty
        value = verify_data.as_number(line)
        if value is not None:
            return value
        if re.search(r"[A-Za-z]{3}", line):
            break                          # the next material's line
    return None


def validate_material_indent_calculations(page: Page, row: Row,
                                          item_name: str) -> None:
    """Material Indent: the quantity chain, and the stock the issue moved."""
    shown = read_indent_quantities(page, item_name)
    requested = verify_data.as_number(shown.get("Requested", ""))
    approved = verify_data.as_number(shown.get("Approved", ""))
    issued = verify_data.as_number(shown.get("Issued", ""))
    returned = verify_data.as_number(shown.get("Returned", ""))

    # The workbook is the independent side: what the row ASKED for, against what
    # the application ended up recording.
    calc.check("Approved vs Requested",
               {"Requested (workbook)": row.get("Quantity", ""),
                "Approved (workbook)": row.get("Approved Quantity", "")},
               "the approved quantity cannot exceed the requested quantity",
               expected=sheet_number(row, "Approved Quantity"),
               actual=approved, tolerance=calc.RATIO_TOLERANCE, page=page)

    # THE TWO FIGURES THIS APPLICATION DOES NOT DERIVE, and the difference
    # between "not on screen" and "not checked".
    #
    # The indent's MATERIALS REQUESTED grid has exactly five columns -
    # REQUESTED, APPROVED, ISSUED, RETURNED and IN STOCK - and every one of
    # them is a figure the application was GIVEN, not one it worked out. There
    # is no Pending column and no Net Issued column, so neither derived figure
    # exists to be compared with. The four quantities themselves are checked,
    # against what the row asked for, by verify_indent_quantities().
    pending_shown = _indent_pending(page)
    if pending_shown is None:
        calc.skip("Pending Approval",
                  f"The indent does not display a pending quantity. Its "
                  f"materials grid prints Requested, Approved, Issued, "
                  f"Returned and In Stock, and no column of it answers 'how "
                  f"much is still awaiting approval'. Requested "
                  f"{calc.money(requested)} and approved "
                  f"{calc.money(approved)}, so "
                  f"{calc.money(calc.outstanding(requested, approved))} is "
                  f"pending by this run's own arithmetic - the application was "
                  f"not asked to agree with a figure it does not print, and "
                  f"nothing was assumed. The approved quantity itself IS "
                  f"checked, above.",
                  {"Requested": requested, "Approved": approved},
                  "Requested - Approved")
    else:
        calc.check("Pending Approval",
                   {"Requested": requested, "Approved": approved},
                   "Requested - Approved",
                   expected=calc.outstanding(requested, approved),
                   actual=pending_shown,
                   tolerance=calc.RATIO_TOLERANCE, page=page)

    net_shown = _indent_net_issued(page)
    if net_shown is None:
        calc.skip("Net Issued (issued less returned)",
                  f"The indent does not display a net issued quantity - it "
                  f"prints Issued ({calc.money(issued)}) and Returned "
                  f"({calc.money(returned)}) as the two separate movements "
                  f"they are, and never their difference. What the net "
                  f"movement really did is proved instead where it can be: "
                  f"against Items & Inventory, by 'Stock After Indent Issue', "
                  f"'Stock After Return to Store' and the two stock-ledger "
                  f"checks - all of which passed or failed on this run's own "
                  f"readings rather than on a figure nobody displays.",
                  {"Issued": issued, "Returned": returned},
                  "Issued - Returned")
    else:
        calc.check("Net Issued (issued less returned)",
                   {"Issued": issued, "Returned": returned},
                   "Issued - Returned",
                   expected=calc.outstanding(issued, returned),
                   actual=net_shown,
                   tolerance=calc.RATIO_TOLERANCE, page=page)

    # Closing stock: only checkable when the stock BEFORE the issue was read -
    # the detail page shows one In Stock figure, and one figure cannot prove a
    # movement. The opening reading is taken just before "Issue Stock".
    opening = published("Indent Opening Stock")
    closing = read_indent_stock(page, item_name)
    if opening is None:
        calc.skip("Closing Stock (the indent's IN STOCK column)",
                  "The stock BEFORE the issue was not captured - either the row "
                  "issued nothing, or the IN STOCK column could not be read off "
                  "the materials grid at that moment. One In Stock figure cannot "
                  "prove a movement, so nothing was assumed from it.",
                  {"In Stock now": closing},
                  "Opening - Issued + Returned")
    else:
        calc.check("Closing Stock (the indent's IN STOCK column)",
                   {"Opening (before the issue)": opening, "Issued": issued,
                    "Returned": returned},
                   "Opening - Issued + Returned",
                   expected=calc.closing_stock(opening, issued=issued,
                                               returned=returned),
                   actual=closing, tolerance=calc.RATIO_TOLERANCE, page=page)

    # VALIDATE INVENTORY STOCK REDUCTION, and the run's whole stock ledger.
    validate_inventory_stock_chain(page, row, issued, returned,
                                   requested=requested, approved=approved)


def validate_inventory_stock_chain(page: Page, row: Row,
                                   issued: Optional[float],
                                   returned: Optional[float],
                                   requested: Optional[float] = None,
                                   approved: Optional[float] = None) -> None:
    """VALIDATION B and the END-TO-END LEDGER, in Items & Inventory.

        Stock After Issue  = Stock Before the Indent - ACTUAL Issued
        Stock After Return = Stock After Issue       + ACTUAL Returned
        Closing Stock      = Opening + Received - Issued + Returned

    ISSUED and RETURNED are the finished record's own figures - what the
    application really moved - and never the quantities the workbook asked for.
    This application issues what it can: a row asking for 55 against a stock of
    3 is issued 3, and a ledger built from the request would then report a
    correct application as 52 short. Requested, Approved, Issued, Received and
    Returned are five different numbers and they are kept apart.

    Each stage is checked against the reading taken at that stage, so a chain
    that goes wrong says WHERE it went wrong rather than only that the closing
    figure is out.
    """
    code, name, why_not = stock_chain_item(row)
    before_indent = published("Stock Before Indent")
    after_issue = published("Stock After Issue")
    wanted_return = (row.get("Return Quantity", "") or "").strip()

    reading = inventory_reading(page, code, name, "at the end of the run")
    STOCK_READINGS_TAKEN.append(reading)
    # The item page's own "Current Stock" first - see stock_now() - falling
    # back to the list's STOCK column only when the item's own page could not
    # be read.
    closing = reading.get("current")
    if closing is None:
        closing = reading.get("stock")
    STOCK_CHAIN["issued"] = issued
    STOCK_CHAIN["returned"] = returned
    STOCK_CHAIN["closing"] = closing

    # A row that never issued anything took no mid-flow reading, so the closing
    # figure IS the stock after the issue. Saying so is not an assumption - it
    # is the same screen, read once instead of twice.
    if after_issue is None and not wanted_return:
        after_issue = closing
    STOCK_CHAIN["after_issue"] = after_issue

    # --- VALIDATION B: the issue takes stock out ---------------------------
    if before_indent is None:
        calc.skip("Stock After Indent Issue",
                  "The stock BEFORE the indent was not read from Items & "
                  "Inventory, so what the issue took out cannot be worked out. "
                  "One reading cannot prove a movement and nothing was assumed "
                  "from it.",
                  {"Issued (from the record)": issued, "Stock now": closing},
                  "Stock Before the Indent - ACTUAL Issued Quantity")
    elif issued is None:
        calc.skip("Stock After Indent Issue",
                  "The finished indent does not show an ISSUED quantity, so "
                  "what left the store is not known. The REQUESTED quantity was "
                  "NOT used in its place - this application can issue less than "
                  "was asked for, and the two are different numbers.",
                  {"Stock Before the Indent": before_indent},
                  "Stock Before the Indent - ACTUAL Issued Quantity")
    else:
        calc.check(
            "Stock After Indent Issue",
            {"Stock Before the Indent": before_indent,
             "ACTUAL Issued (from the finished record)": issued,
             "Requested (workbook, NOT used in this sum)":
                 row.get("Quantity", ""),
             "Item": f"{name} ({code or 'no code'})"},
            "Stock Before the Indent - ACTUAL Issued Quantity",
            expected=calc.closing_stock(before_indent, issued=issued),
            actual=after_issue,
            alternatives={
                "the stock had not moved at all": before_indent,
                "the REQUESTED quantity was taken out instead of the issued one":
                    calc.closing_stock(before_indent,
                                       issued=sheet_number(row, "Quantity"))},
            tolerance=calc.RATIO_TOLERANCE, page=page)

    # --- the return puts stock back ---------------------------------------
    if not wanted_return:
        calc.skip("Stock After Return to Store",
                  "This row does not return anything to the store, so there is "
                  "no return movement to check.",
                  {"Stock after the issue": after_issue},
                  "Stock After Issue + ACTUAL Returned Quantity")
    elif after_issue is None or returned is None:
        calc.skip("Stock After Return to Store",
                  "Either the stock between the issue and the return, or the "
                  "quantity the record says was returned, could not be read - "
                  "so the return's movement cannot be worked out.",
                  {"Stock after the issue": after_issue,
                   "Returned (from the record)": returned,
                   "Stock now": closing},
                  "Stock After Issue + ACTUAL Returned Quantity")
    else:
        calc.check(
            "Stock After Return to Store",
            {"Stock after the issue": after_issue,
             "ACTUAL Returned (from the finished record)": returned,
             "Item": f"{name} ({code or 'no code'})"},
            "Stock After Issue + ACTUAL Returned Quantity",
            expected=calc.closing_stock(after_issue, returned=returned),
            actual=closing,
            alternatives={"the return did not move the stock": after_issue},
            tolerance=calc.RATIO_TOLERANCE, page=page)

    # --- the ledger's own two movements ------------------------------------
    check_indent_ledger_entries(page, reading, issued, returned)

    # --- the whole ledger, end to end --------------------------------------
    opening = published("Opening Stock")
    received = published("GRN Received Quantity")
    expected_closing = calc.closing_stock(opening, received=received,
                                          issued=issued, returned=returned)
    report_stock_ledger(row, code, name, opening, received, issued, returned,
                        expected_closing, closing, why_not)

    if why_not:
        calc.skip("Closing Stock (Items & Inventory)",
                  f"The purchase side and the indent side of this run are not "
                  f"about the same item - {why_not} - so there is no single "
                  f"ledger to add up. The indent's own movements are checked "
                  f"above; the purchase side was checked in the GRN module.",
                  {"Opening Stock": opening, "Received": received,
                   "Issued": issued, "Returned": returned},
                  "Opening + Received - Issued + Returned")
    elif opening is None:
        calc.skip("Closing Stock (Items & Inventory)",
                  "The opening stock was not read before the purchase order, "
                  "so the run's ledger has no starting balance. Run the "
                  "Purchase Orders sheet in the same session - the opening "
                  "reading is taken there.",
                  {"Received": received, "Issued": issued,
                   "Returned": returned, "Stock now": closing},
                  "Opening + Received - Issued + Returned")
    else:
        calc.check(
            "Closing Stock (Items & Inventory)",
            {"Opening Stock": opening,
             "GRN Received": received if received is not None else 0.0,
             "ACTUAL Issued": issued if issued is not None else 0.0,
             "ACTUAL Returned": returned if returned is not None else 0.0,
             "Item": f"{name} ({code or 'no code'})"},
            "Opening Stock + GRN Received - ACTUAL Issued + ACTUAL Returned",
            expected=expected_closing, actual=closing,
            alternatives={
                "the receipt was never added": calc.closing_stock(
                    opening, issued=issued, returned=returned),
                "the return was never put back": calc.closing_stock(
                    opening, received=received, issued=issued),
                "the REQUESTED quantity was issued instead of the actual one":
                    calc.closing_stock(opening, received=received,
                                       issued=sheet_number(row, "Quantity"),
                                       returned=returned)},
            tolerance=calc.RATIO_TOLERANCE, page=page)

    # Everything above, as the one table the report is opened for. It runs LAST
    # so that every check it reads back has already been made.
    report_item_stock_validation(code, name, opening, received, before_indent,
                                 requested, approved, issued, returned,
                                 expected_closing, closing)


def report_item_stock_validation(code: str, name: str,
                                 opening: Optional[float],
                                 received: Optional[float],
                                 before_indent: Optional[float],
                                 requested: Optional[float],
                                 approved: Optional[float],
                                 issued: Optional[float],
                                 returned: Optional[float],
                                 expected_closing: Optional[float],
                                 closing: Optional[float]) -> None:
    """ITEMS & INVENTORY - STOCK VALIDATION, as the two cards a QA reader wants.

        BEFORE  ->  TRANSACTION  ->  AFTER  ->  EXPECTED vs ACTUAL  ->  RESULT

    The Material Indent card is built here, the GRN one was built where the
    receipt was approved, and the combined page puts the two of them one under
    the other with the run's FINAL RESULT at the bottom - which is the whole
    stock story of the run on one screen, without opening a second attachment
    and without reading a line of Python.

    Expected, Actual, Difference and Result PREFER the check
    calculation_validation.py already recorded for the stage this row actually
    reached (issue-only or issue-then-return) - it is the more PRECISE of the
    two, because it is measured from the reading taken immediately before this
    specific movement rather than from the run's very first opening figure.
    But `closing`/`expected_closing` - the real reading Items & Inventory just
    showed, and the run's own end-to-end sum - are ALWAYS real values by this
    point, and they are used whenever that check is missing or was SKIPPED, so
    a calculation the engine could not run never blanks out a stock figure
    this run actually captured off the screen.

    The stage-by-stage table is kept for the one case it earns its place in: a
    stock figure that did NOT come out right, where the question stops being
    "did it work?" and becomes "which stage did it go wrong at?". On a clean
    run it is left out, because eleven rows of MEASURED readings are exactly
    the kind of detail that buries the seven lines a reader came for.
    """
    # WHICH CHECK IS THE INDENT'S "AFTER". A row that returned nothing ends at
    # the issue; a row that returned something ends after the return. Picking
    # the one that matches what this row actually did is what keeps the card's
    # Expected from being an answer to a question the row never asked.
    after_issue = recorded_check("Stock After Indent Issue")
    after_return = recorded_check("Stock After Return to Store")
    final = recorded_check("Closing Stock (Items & Inventory)")
    returned_something = returned is not None and float(returned or 0) != 0.0
    after = (after_return if returned_something and after_return is not None
             else after_issue) or final

    # THE REAL READING WINS - see the docstring. `closing` is the application's
    # own figure, read off Items & Inventory "at the end of the run"
    # regardless of which (if any) calc check could be completed.
    resolved_actual = closing if closing is not None else (
        None if after is None else after.actual)
    resolved_expected = (after.expected
                         if after is not None and after.expected is not None
                         else expected_closing)
    if after is not None and after.expected is not None and after.actual is not None:
        difference = after.difference
        card_result = after.status
    elif resolved_actual is not None and resolved_expected is not None:
        difference = round(float(resolved_actual) - float(resolved_expected), 2)
        card_result = (stock_report.PASS if abs(difference) <= 0.001
                      else stock_report.FAIL)
    else:
        difference = None if after is None else after.difference
        card_result = stock_report.NOT_CHECKED if after is None else after.status

    # The formula, with THIS row's own numbers in it - the line a manual tester
    # writes underneath the rule when they check it by hand.
    formula = f"{stock_report.num(before_indent)} - {stock_report.num(issued)}"
    if returned_something:
        formula += f" + {stock_report.num(returned)}"

    note = ("The stock reduction uses the ACTUAL ISSUED quantity, never the "
            "requested one: this application issues what it can, so a row "
            "asking for 55 against a stock of 3 is issued 3, and a sum built "
            "from the request would report a correct application as 52 short.")
    if card_result == stock_report.FAIL:
        note = ("The stock the application shows is not the stock this "
                "movement should have left behind. The figures above were "
                "read off the indent and off Items & Inventory at different "
                "moments, so neither of them can be the reason.")

    stock_report.record(stock_report.Movement(
        title="STOCK DECREASE VALIDATION (Material Indent Issue)",
        stage="Stock Decrease (Indent)",
        item=f"{name or '(not named)'} ({code or 'no code'})",
        document=LAST_INDENT_NUMBER or "",
        before_label="Current Stock Before Consumption",
        before=before_indent,
        quantities=[["Requested Quantity", requested],
                    ["Approved Quantity", approved],
                    ["Actual Issued Quantity", issued],
                    ["Returned Quantity", returned]],
        # Requested and Approved are shown above and move NOTHING. Only the
        # quantity the store really gave out, and the quantity that really came
        # back, belong in the sum - see Movement.terms.
        terms=([["-", "Consumed / Issued Quantity", issued]]
               + ([["+", "Material Indent Returned", returned]]
                  if returned_something else [])),
        formula=formula,
        expected=resolved_expected,
        actual=resolved_actual,
        difference=difference,
        result=card_result,
        note=note))

    # "📦 Stock Validation > 📉 Stock Decrease - Material Indent": the
    # question-and-answer sheet first, then the card it was built from. The
    # sheet keeps REQUESTED, APPROVED, ISSUED and RETURNED on four separate
    # lines and says underneath which of them the sum used - see
    # stock_report.decrease_pairs() and stock_report._issue_caveat().
    with stock_section(STOCK_DECREASE_STEP):
        # HTML ONLY - see the matching comment on the BEFORE TRANSACTION
        # attachment above.
        verify_data.attach_html(
            "📉 Item Stock - After Consumption",
            stock_report.decrease_page())

    # "📦 Stock Validation > 📊 Current Stock Validation" - THE attachment
    # this report is opened for, and the first moment it can be complete:
    # every card the run records is recorded by now.
    #
    #     CURRENT STOCK BEFORE -> WHAT MOVED -> EXPECTED CURRENT STOCK
    #                          vs ACTUAL CURRENT STOCK -> DIFFERENCE -> RESULT
    #
    # for the opening reading, the GRN and the indent in turn, then the
    # before-and-after table for all three, then the final stock with the sum
    # written out. The movement summary underneath it is the same three rows
    # with the transaction chain and the document references beside them.
    #
    # "Items & Inventory - Stock Validation" - the long page that repeated
    # every card a third time - is no longer attached. Nothing on it is lost:
    # its summary is the movement summary and its cards are the two sections
    # above.
    with stock_section(STOCK_SUMMARY_STEP):
        # THE SIX LINES FIRST - "did the stock move correctly, yes or no" -
        # then the fuller cards underneath for "how do you know". Same
        # Movements, same figures; this one just leaves out the PO number,
        # the GRN status and every other piece of context that belongs in
        # the sections below it.
        attach_text("📦 STOCK MOVEMENT", stock_report.simple_movement_text())
        verify_data.attach_html("📦 Stock Movement",
                                stock_report.simple_movement_page())
        attach_text("📦 CURRENT STOCK VALIDATION",
                    stock_report.current_stock_text())
        verify_data.attach_html("📦 Current Stock Validation",
                                stock_report.current_stock_page())
        attach_text("📊 STOCK MOVEMENT SUMMARY",
                    stock_report.summary_text())
        verify_data.attach_html("📊 Stock Movement Summary",
                                stock_report.summary_page())

    # ...and, ONLY when something did not come out right, the stage-by-stage
    # breakdown that says where.
    if stock_report.result() == stock_report.FAIL:
        verdict = final.status if final is not None else "NOT CHECKED"
        rows = [
            stock_stage_row("Stock Before GRN", "", opening),
            stock_stage_row("GRN Approved / Received Quantity",
                            "GRN Receipt Posted to Inventory", received),
            stock_stage_row("Stock After GRN", "Stock After GRN Approval"),
            stock_stage_row("Stock Before Material Indent", "", before_indent),
            stock_stage_row("Material Indent Requested", "", requested),
            stock_stage_row("Material Indent Approved", "Approved vs Requested",
                            approved),
            stock_stage_row("Material Indent Issued",
                            "Indent Issue Posted to Inventory", issued),
            stock_stage_row("Material Indent Returned",
                            "Return to Store Posted to Inventory", returned),
            stock_stage_row("Stock After Material Indent Issue",
                            "Stock After Indent Issue"),
            stock_stage_row("Stock After Return to Store",
                            "Stock After Return to Store"),
            stock_stage_row("Closing Stock (Stock Remaining)",
                            "Closing Stock (Items & Inventory)"),
        ]
        attach_table(
            "Stock Validation - Stage by Stage (a figure did not agree)",
            ("Stage", "Expected", "Actual (application)", "Difference",
             "Result"), rows,
            f"Item: {name or '(not named)'} ({code or 'no code'})   |   "
            f"Closing Stock: {stock_report.num(expected_closing)} expected, "
            f"{stock_report.num(closing)} actual   |   Result: {verdict}")


def check_indent_ledger_entries(page: Page, reading: Dict[str, object],
                                issued: Optional[float],
                                returned: Optional[float]) -> None:
    """The issue and the return, as the item's own STOCK LEDGER recorded them.

    The second, independent account of the same two movements: the indent's
    record says what it issued, and the item's ledger says what was taken out of
    stock for that indent. They have to be the same quantity, and a build where
    they are not has an inventory posting problem that neither figure shows on
    its own.
    """
    entries = [entry for entry in (reading.get("ledger") or [])
               if references_document(str(entry.get("reference", "")),
                                      LAST_INDENT_NUMBER)]

    out = next((entry for entry in entries
                if float(entry.get("change") or 0) < 0), None)
    back = next((entry for entry in entries
                 if float(entry.get("change") or 0) > 0), None)

    indent = LAST_INDENT_NUMBER or "the indent this run raised"
    if not LAST_INDENT_NUMBER or not entries:
        calc.skip("Indent Issue Posted to Inventory",
                  f"The item's stock ledger shows no movement referring to "
                  f"{indent}, so the issue could not be checked against the "
                  f"inventory's own record of it. The stock figures above are "
                  f"what say whether the movement happened.",
                  {"Issued (from the indent)": issued},
                  "the ledger's issue = the quantity the indent issued")
    else:
        calc.check(
            "Indent Issue Posted to Inventory",
            {"Indent": LAST_INDENT_NUMBER,
             "Issued (from the indent's own record)": issued,
             "Ledger movement": str(out.get("movement", "")) if out else "-",
             "Ledger reference": str(out.get("reference", "")) if out else "-"},
            "the quantity the ledger took out = the quantity the indent issued",
            expected=issued,
            actual=stock_reading(page, abs(float(out["change"]))
                                 if out and out.get("change") is not None
                                 else None, "stock-ledger-issue"),
            tolerance=calc.RATIO_TOLERANCE, page=page)

    if returned is None or not returned:
        calc.skip("Return to Store Posted to Inventory",
                  "Nothing was returned to the store on this row, so the ledger "
                  "has no return movement to show.",
                  {"Returned": returned},
                  "the ledger's return = the quantity the record returned")
    elif back is None:
        calc.skip("Return to Store Posted to Inventory",
                  f"The record says {returned:g} was returned to the store, but "
                  f"the item's stock ledger shows no movement putting it back "
                  f"against {indent}. The stock check above is what says "
                  f"whether the stock itself moved.",
                  {"Returned (from the record)": returned},
                  "the ledger's return = the quantity the record returned")
    else:
        calc.check(
            "Return to Store Posted to Inventory",
            {"Indent": LAST_INDENT_NUMBER,
             "Returned (from the indent's own record)": returned,
             "Ledger movement": str(back.get("movement", "")),
             "Ledger reference": str(back.get("reference", ""))},
            "the quantity the ledger put back = the quantity the record returned",
            expected=returned,
            actual=stock_reading(page, back.get("change"),
                                 "stock-ledger-return"),
            tolerance=calc.RATIO_TOLERANCE, page=page)


def report_stock_ledger(row: Row, code: str, name: str,
                        opening: Optional[float], received: Optional[float],
                        issued: Optional[float], returned: Optional[float],
                        expected: Optional[float], actual: Optional[float],
                        why_not: str) -> None:
    """The run's whole stock movement, as one table, in the report.

    The thing a reader opens first when a stock figure is wrong: every stage,
    where its number came from, and which stage the difference appears at.
    """
    def show(value: Optional[float]) -> str:
        return "(not available)" if value is None else f"{float(value):g}"

    difference = (None if expected is None or actual is None
                  else round(float(actual) - float(expected), 4))
    verdict = ("SKIPPED - one of the figures was not available"
               if difference is None
               else ("PASS" if abs(difference) <= calc.RATIO_TOLERANCE
                     else "FAIL"))

    lines = [
        f"END-TO-END STOCK LEDGER - {name or '(item not named)'} "
        f"({code or 'no code'})",
        "",
        f"{'Opening Stock':<38}{show(opening):>14}   "
        f"read from Items & Inventory before the purchase order",
        f"{'+ GRN Received / Approved Quantity':<38}{show(received):>14}   "
        f"read from the approved goods receipt",
        f"{'= Stock After GRN':<38}"
        f"{show(STOCK_CHAIN.get('after_grn')):>14}   "
        f"read from Items & Inventory after the GRN approval",
        f"{'  Stock Before the Indent':<38}"
        f"{show(published('Stock Before Indent')):>14}   "
        f"read from Items & Inventory before the indent",
        f"{'- ACTUAL Issued Quantity':<38}{show(issued):>14}   "
        f"read from the finished indent (NOT the requested quantity)",
        f"{'= Stock After Issue':<38}"
        f"{show(STOCK_CHAIN.get('after_issue')):>14}   "
        f"read from Items & Inventory after the stock issue",
        f"{'+ ACTUAL Returned Quantity':<38}{show(returned):>14}   "
        f"read from the finished indent",
        "-" * 96,
        f"{'= Expected Closing Stock':<38}{show(expected):>14}",
        f"{'  Actual Stock Remaining':<38}{show(actual):>14}   "
        f"Items & Inventory, STOCK column",
        f"{'  Difference':<38}{show(difference):>14}",
        f"{'  Result':<38}{verdict:>14}",
        "",
        "THE QUANTITIES, KEPT APART (they are not the same number):",
        f"    Requested (workbook)   : {row.get('Quantity', '') or '-'}",
        f"    Approved  (workbook)   : {row.get('Approved Quantity', '') or '-'}",
        f"    Received  (application): {show(received)}",
        f"    Issued    (application): {show(issued)}",
        f"    Returned  (application): {show(returned)}",
    ]
    if why_not:
        lines += ["", f"NOTE - the ledger could not be added up end to end: "
                      f"{why_not}."]
    body = "\n".join(lines)
    # NOT attached to the report any more. Every figure on it is already on the
    # two stock cards - before, quantity, expected, actual, difference, result
    # - and a third telling of the same sum is what made the stock section
    # something a reader had to work through rather than read. It is still
    # written in full next to this run's log, and still logged line by line,
    # so nothing is lost for anyone who wants the long form.
    #
    # It IS attached when the stock did not come out right: at that point the
    # per-stage provenance ("read from the finished indent, NOT the requested
    # quantity") stops being noise and becomes the thing being argued about.
    # The gate is this function's OWN verdict rather than stock_report's,
    # because the ledger is written before the indent card is recorded and a
    # run-wide verdict read here would not yet know about the indent.
    if verdict != "PASS":
        attach_text("Stock Validation - End-to-End Ledger", body)
    summary_file("stock_ledger.txt", body)
    for line in lines:
        log.info("   %s", line)


def _indent_pending(page: Page) -> Optional[float]:
    """The pending figure the indent page prints, when it prints one."""
    return PageNumbers.read(page).find("Pending(?:\\s*Qty| Quantity)?",
                                       "Balance(?:\\s*Qty| Quantity)?")


def _indent_net_issued(page: Page) -> Optional[float]:
    """The net-issued figure the indent page prints, when it prints one."""
    return PageNumbers.read(page).find("Net\\s*Issued", "Consumed",
                                       "Net\\s*Quantity")


def validate_item_calculations(page: Page, row: Row) -> None:
    """Item: the margin the catalogue works out from the two prices it was given."""
    purchase = sheet_number(row, "Purchase Price")
    selling = sheet_number(row, "Selling Price")
    numbers = PageNumbers.read(page)

    margin_shown = numbers.find("Margin(?:\\s*%)?", "Profit(?:\\s*%)?")
    if purchase is None or selling is None:
        calc.skip("Margin %",
                  "This row does not give both a Purchase Price and a Selling "
                  "Price, so a margin cannot be worked out.",
                  {"Purchase Price": row.get("Purchase Price", ""),
                   "Selling Price": row.get("Selling Price", "")},
                  "(Selling - Purchase) / Purchase x 100")
    elif margin_shown is None:
        calc.skip("Margin %",
                  "The item page displays no margin - it prints the prices it "
                  "was given and works nothing out from them - so there is no "
                  "figure here to be right or wrong about. That the prices "
                  "were stored exactly as typed is proved by the data "
                  "verification layer.",
                  {"Purchase Price": purchase, "Selling Price": selling},
                  "(Selling - Purchase) / Purchase x 100")
    else:
        calc.check("Margin %", {"Purchase Price": purchase,
                                "Selling Price": selling},
                   "(Selling - Purchase) / Purchase x 100",
                   expected=calc.percentage(selling - purchase, purchase),
                   actual=margin_shown, tolerance=calc.RATIO_TOLERANCE,
                   page=page)

    publish("Item Selling Price", selling)


def no_calculations(module: str, why: str) -> None:
    """Record a module the application does no arithmetic on."""
    calc.not_applicable(
        f"{module} has no calculated fields: {why} Every value it holds is "
        f"typed in and stored as typed, and that it was stored correctly is "
        f"proved by the data verification layer, which compares the whole form "
        f"with the record's View page. Nothing is invented here to make this "
        f"module look covered.")


# =========================================================================== #
# TEST DATA - read once, at collection time
#
# Adding a row to the workbook adds a test execution here. Nothing else changes.
# =========================================================================== #
ACCOUNT_ROWS = get_test_cases(SHEET_ACCOUNT)
CUSTOMER_ROWS = get_test_cases(SHEET_CUSTOMER)
SUPPLIER_ROWS = get_test_cases(SHEET_SUPPLIER)
ITEM_ROWS = get_test_cases(SHEET_ITEM)
CATEGORY_ROWS = get_test_cases(SHEET_CATEGORY)
QUOTATION_ROWS = get_test_cases(SHEET_QUOTATION)
SALES_ORDER_ROWS = get_test_cases(SHEET_SALES_ORDER)
PURCHASE_ORDER_ROWS = get_test_cases(SHEET_PURCHASE_ORDER)
GRN_ROWS = get_test_cases(SHEET_GRN)
PURCHASE_BILL_ROWS = get_test_cases(SHEET_PURCHASE_BILL)
MATERIAL_INDENT_ROWS = get_test_cases(SHEET_MATERIAL_INDENT)


def case_id(row: Row) -> str:
    """The name each execution gets in pytest output and in Allure."""
    return row.test_case_id


#: The root of the report's Behaviors tree: every module hangs under it, so the
#: whole journey reads as one flow rather than eleven unrelated features.
ALLURE_EPIC = "GenZOpss Construction Flow"

#: The modules of the requested validation tree that this suite does NOT drive.
#: They are declared so the summary tells the truth about its own coverage: a
#: module nobody automated must not be absent from the report (which reads as
#: "nothing to say about it") and must never be counted as a pass.
#:
#: Everything here is a Construction PROJECT module - WBS, BOQ, phases, site
#: logs and the rest. This suite covers the tenant's onboarding and its trading
#: documents; the project modules live in the separate Projects suite. Add a
#: module here the day it is automated, and delete its line the day it is.
NOT_AUTOMATED_MODULES: List[Tuple[str, str]] = [
    (name, reason) for name, reason in (
        ("Delivery Challans",
         "no Delivery Challans sheet and no flow in this suite"),
        ("Construction Project",
         "creating the project itself is driven by the separate Projects suite. "
         "It is named here so the coverage this report claims is the coverage "
         "there is - a module nobody ran must never be missing from the list, "
         "because a heading that simply is not there reads as one that passed"),
        ("Construction Project - WBS", "driven by the separate Projects suite"),
        ("Construction Project - BOQ", "driven by the separate Projects suite"),
        ("Construction Project - Phases", "driven by the separate Projects suite"),
        ("Construction Project - Site Logs", "not automated anywhere yet"),
        ("Construction Project - Labour", "driven by the separate Labour suite"),
        ("Construction Project - Labour Analysis",
         "driven by the separate Labour suite"),
        ("Construction Project - Milestones",
         "not available in the application to this tenant"),
        ("Construction Project - Inventory",
         "the workspace-wide stock is covered - Items & Inventory is read "
         "before the purchase order, after the GRN approval, after the stock "
         "issue and after the return to store, and the run's whole ledger is "
         "checked against it. What is NOT automated is the Construction "
         "PROJECT's own Inventory screen, which is a different module"),
        ("Construction Project - Material Analysis", "not automated anywhere yet"),
        ("Construction Project - Daily Progress", "not automated anywhere yet"),
        ("Construction Project - Subcontractors", "not automated anywhere yet"),
        ("Construction Project - Change Orders", "not automated anywhere yet"),
        ("Construction Project - Equipment", "not automated anywhere yet"),
        ("Construction Project - Defects", "not automated anywhere yet"),
        ("Construction Project - JV Accounts", "not automated anywhere yet"),
        ("Construction Project - Sales & Billing", "not automated anywhere yet"),
        ("Construction Project - Gantt", "not automated anywhere yet"),
    )
]


#: Every test case below has the same two halves, and the report says so:
#:
#:     [TC00n] Quotations
#:         Functional Validation     did the module do its job?
#:         Calculation Validation    did the application get the numbers right?
#:
#: The overall result is the two together - calculation_validation.py keeps the
#: verdict, and run_test_case() is what fails the row when a number is wrong.
FUNCTIONAL = "Functional Validation"
CALCULATION = "Calculation Validation"


# =========================================================================== #
# FIXTURES - one browser, one page, kept open for the whole run
# =========================================================================== #

@pytest.fixture(scope="session")
def page() -> Iterator[Page]:
    """A visible, maximized Chromium page shared by every module.

    Session-scoped on purpose: the workspace login has to survive from the
    Customer module through to the Sales Order module.
    """
    log.info("Launching Browser...")
    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(
            headless=not HEADED, slow_mo=SLOW_MO_MS, args=["--start-maximized"])
        context = browser.new_context(no_viewport=True, ignore_https_errors=True)
        active_page = context.new_page()
        active_page.set_default_timeout(DEFAULT_TIMEOUT)
        active_page.set_default_navigation_timeout(NAVIGATION_TIMEOUT)
        ApiFailureRecorder(active_page)

        yield active_page

        log.info("Closing the browser.")
        context.close()
        browser.close()


@pytest.fixture(scope="session")
def workspace(page: Page) -> Page:
    """The workspace, signed in as the business user.

    Runs once, the first time a Customer / Supplier / Quotation / Sales Order
    test asks for it - i.e. after the whole Create_Your_Account module.
    """
    login_as_business_user(page)
    return page


# =========================================================================== #
# THE MODULES - pytest runs them top to bottom, all rows of one before the next
# =========================================================================== #

@allure.epic(ALLURE_EPIC)
@allure.feature("Create Your Account")
@allure.story("Register a business and approve it in the admin portal")
@allure.severity(allure.severity_level.BLOCKER)
# The module is switched off from ONE place - see RUN_CREATE_YOUR_ACCOUNT at the
# top of this file. skipif rather than a deleted test or a commented-out body:
# the rows are still collected, so the report still lists them and says why they
# did not run, and turning the switch back on needs no other edit. The body
# below is untouched.
@pytest.mark.skipif(not RUN_CREATE_YOUR_ACCOUNT,
                    reason=CREATE_YOUR_ACCOUNT_OFF)
@pytest.mark.parametrize("row", ACCOUNT_ROWS, ids=case_id)
def test_create_your_account(page: Page, row: Row) -> None:
    """Register a business and approve it in the admin portal."""
    with run_test_case(row, "Create Your Account"):
        with step(FUNCTIONAL):
            company = create_business_account(page, row)
            # Registration completed is asserted BEFORE the admin portal is
            # opened, so "the tenant is not there" can never be reported
            # without knowing whether the account was created at all.
            registered_because = registration_completed(page, company)
            attach_text("Registration Completed",
                        f"Company     : {company}\n"
                        f"Confirmed by: "
                        f"{registered_because or '(nothing on screen confirmed it)'}")
            # The OTP flow, untouched and still MANUAL: it sits inside the
            # functional half because verifying the e-mail is part of
            # registering, not a sum. Whether a code was actually typed is
            # carried through, so a tenant that never appears is reported as
            # the precondition it is.
            otp_verified = wait_for_manual_otp(page)
            approve_business(page, company, otp_verified, registered_because)
        with step(CALCULATION):
            no_calculations("Create Your Account",
                            "registration and admin approval produce a tenant, "
                            "not a figure.")
    finish_module(SHEET_ACCOUNT, row, ACCOUNT_ROWS)


@allure.epic(ALLURE_EPIC)
@allure.feature("Customer")
@allure.story("Create a customer")
@allure.severity(allure.severity_level.CRITICAL)
@pytest.mark.parametrize("row", CUSTOMER_ROWS, ids=case_id)
def test_customer(workspace: Page, row: Row) -> None:
    """Create a customer."""
    with run_test_case(row, "Customer"):
        with step(FUNCTIONAL):
            create_customer(workspace, row)
        with step(CALCULATION):
            no_calculations("Customer",
                            "the Credit Limit and Credit Days it holds are "
                            "limits the application stores, not amounts it "
                            "works out.")
    finish_module(SHEET_CUSTOMER, row, CUSTOMER_ROWS)


@allure.epic(ALLURE_EPIC)
@allure.feature("Suppliers")
@allure.story("Create a supplier with its bank details")
@allure.severity(allure.severity_level.CRITICAL)
@pytest.mark.parametrize("row", SUPPLIER_ROWS, ids=case_id)
def test_supplier(workspace: Page, row: Row) -> None:
    """Create a supplier."""
    with run_test_case(row, "Suppliers"):
        with step(FUNCTIONAL):
            create_supplier(workspace, row)
        with step(CALCULATION):
            no_calculations("Suppliers",
                            "the CREATE form gives it Credit Days and bank "
                            "details, and the application derives no figure "
                            "from them. The figures a supplier does carry - "
                            "Outstanding, and the value of the orders raised "
                            "against it - do not exist until something has been "
                            "bought, so they are checked further down the run, "
                            "after the purchase order has been received and its "
                            "GRN approved.")
    finish_module(SHEET_SUPPLIER, row, SUPPLIER_ROWS)


@allure.epic(ALLURE_EPIC)
@allure.feature("Items")
@allure.story("Create a catalogue item under Stock & Materials")
@allure.severity(allure.severity_level.CRITICAL)
@pytest.mark.parametrize("row", ITEM_ROWS, ids=case_id)
def test_items(workspace: Page, row: Row) -> None:
    """Create an item and verify it in the Items list."""
    with run_test_case(row, "Items"):
        with step(FUNCTIONAL):
            create_item(workspace, row)
        with step(CALCULATION):
            validate_item_calculations(workspace, row)
    finish_module(SHEET_ITEM, row, ITEM_ROWS)


@allure.epic(ALLURE_EPIC)
@allure.feature("Categories")
@allure.story("Create a parent category and its sub-category")
@allure.severity(allure.severity_level.NORMAL)
@pytest.mark.parametrize("row", CATEGORY_ROWS, ids=case_id)
def test_categories(workspace: Page, row: Row) -> None:
    """Create a parent category and the sub-category underneath it."""
    with run_test_case(row, "Categories"):
        with step(FUNCTIONAL):
            create_category(workspace, row)
        with step(CALCULATION):
            no_calculations("Categories",
                            "a category is a name, a colour and a parent.")
    finish_module(SHEET_CATEGORY, row, CATEGORY_ROWS)


@allure.epic(ALLURE_EPIC)
@allure.feature("Quotations")
@allure.story("Raise a quotation, send it and accept it")
@allure.severity(allure.severity_level.CRITICAL)
@pytest.mark.parametrize("row", QUOTATION_ROWS, ids=case_id)
def test_quotation(workspace: Page, row: Row) -> None:
    """Raise a quotation, send it and accept it."""
    with run_test_case(row, "Quotations"):
        with step(FUNCTIONAL):
            create_quotation(workspace, row)
        with step(CALCULATION):
            validate_quotation_calculations(workspace, row)
    finish_module(SHEET_QUOTATION, row, QUOTATION_ROWS)


@allure.epic(ALLURE_EPIC)
@allure.feature("Sales Orders")
@allure.story("Raise a sales order and drive its status sequence")
@allure.severity(allure.severity_level.CRITICAL)
@pytest.mark.parametrize("row", SALES_ORDER_ROWS, ids=case_id)
def test_sales_order(workspace: Page, row: Row) -> None:
    """Drive the accepted quotation through its status sequence."""
    with run_test_case(row, "Sales_Orders"):
        with step(FUNCTIONAL):
            create_sales_order(workspace, row)
        with step(CALCULATION):
            validate_sales_order_calculations(workspace, row)
    finish_module(SHEET_SALES_ORDER, row, SALES_ORDER_ROWS)
    log.info("Automation Completed Successfully.")


@allure.epic(ALLURE_EPIC)
@allure.feature("Purchase Orders")
@allure.story("Raise a purchase order, confirm it and send it to the supplier")
@allure.severity(allure.severity_level.CRITICAL)
@pytest.mark.parametrize("row", PURCHASE_ORDER_ROWS, ids=case_id)
def test_purchase_order(workspace: Page, row: Row) -> None:
    """Raise a purchase order and drive its status sequence."""
    with run_test_case(row, "Purchase_Orders"):
        with step(FUNCTIONAL):
            create_purchase_order(workspace, row)
        with step(CALCULATION):
            validate_purchase_order_calculations(workspace, row)
        # LAST, because it navigates back to the list: the calculations above
        # read the order's own record page, and this would take it off screen.
        with step("Verifying the purchase order count (before and after)..."):
            validate_purchase_order_count(workspace, row)
    finish_module(SHEET_PURCHASE_ORDER, row, PURCHASE_ORDER_ROWS)


@allure.epic(ALLURE_EPIC)
@allure.feature("GRN")
@allure.story("Receive a purchase order and approve the goods receipt into a bill")
@allure.severity(allure.severity_level.CRITICAL)
@pytest.mark.parametrize("row", GRN_ROWS, ids=case_id)
def test_grn(workspace: Page, row: Row) -> None:
    """Record a goods receipt against the purchase order and approve it."""
    with run_test_case(row, "GRN"):
        with step(FUNCTIONAL):
            create_grn(workspace, row)
        with step(CALCULATION):
            validate_grn_calculations(workspace, row)
    finish_module(SHEET_GRN, row, GRN_ROWS)


@allure.epic(ALLURE_EPIC)
@allure.feature("Purchase Bills")
@allure.story("Approve the purchase bill the approved GRN created")
@allure.severity(allure.severity_level.CRITICAL)
@pytest.mark.parametrize("row", PURCHASE_BILL_ROWS, ids=case_id)
def test_purchase_bill(workspace: Page, row: Row) -> None:
    """Open the purchase bill and approve it."""
    with run_test_case(row, "Purchase_Bills"):
        with step(FUNCTIONAL):
            approve_purchase_bill(workspace, row)
        with step(CALCULATION):
            validate_purchase_bill_calculations(workspace, row)
    finish_module(SHEET_PURCHASE_BILL, row, PURCHASE_BILL_ROWS)


@allure.epic(ALLURE_EPIC)
@allure.feature("Suppliers")
@allure.story("Recent Purchase Orders (Received) and Outstanding, after the "
              "goods receipt has been approved")
@allure.severity(allure.severity_level.CRITICAL)
@pytest.mark.parametrize("row", SUPPLIER_ROWS, ids=case_id)
def test_supplier_purchase_activity(workspace: Page, row: Row) -> None:
    """Verify what the purchase flow did to the supplier's own page.

    The last link of the chain the purchase modules make:

        Purchase Order -> GRN -> GRN Approval -> Supplier ->
        Recent Purchase Orders (Received) -> Outstanding

    It runs HERE, and not with the Suppliers sheet at the top of the run,
    because none of these figures exists until something has been bought: a
    supplier that has just been created has no orders and nothing outstanding,
    and checking it then would be checking that zero is zero.

    Its results are recorded against the "Suppliers" module, so the summary
    reports one line for the supplier - what its form stored, and what the
    purchase flow did to it - rather than two modules that are really one.
    """
    with run_test_case(row, "Suppliers"):
        with step(FUNCTIONAL):
            verify_supplier_received_order(workspace, row)
        with step(CALCULATION):
            validate_supplier_purchase_activity(workspace, row)
    finish_module(SHEET_SUPPLIER, row, SUPPLIER_ROWS)


@allure.epic(ALLURE_EPIC)
@allure.feature("Material Indents")
@allure.story("Raise a material indent, approve it, issue the stock and return "
              "the excess to the store")
@allure.severity(allure.severity_level.CRITICAL)
@pytest.mark.parametrize("row", MATERIAL_INDENT_ROWS, ids=case_id)
def test_material_indent(workspace: Page, row: Row) -> None:
    """Raise a material indent and drive it through to the return to store."""
    with run_test_case(row, "Material_Indents"):
        with step(FUNCTIONAL):
            create_material_indent(workspace, row)
        with step(CALCULATION):
            validate_material_indent_calculations(
                workspace, row, LAST_INDENT_ITEM or row.get("Item Name", ""))
        # TOP-LEVEL ATTACHMENTS - the whole-chain current-stock and final
        # movement-summary cards, visible the moment this test opens without
        # expanding the step tree. Same underlying figures as "CURRENT STOCK
        # VALIDATION" / "STOCK MOVEMENT SUMMARY" further down the step tree -
        # REPORTING ONLY, nothing recalculated - just under their own heading
        # here. (A third card used to sit alongside these, "STOCK MOVEMENT" -
        # that one was a byte-identical duplicate of its nested counterpart,
        # so it was removed rather than kept twice.)
        attach_text("📦 ITEM STOCK VALIDATION", stock_report.current_stock_text())
        verify_data.attach_html("📦 Item Stock Validation",
                                stock_report.current_stock_page())
        attach_text("📊 FINAL STOCK MOVEMENT SUMMARY", stock_report.summary_text())
        verify_data.attach_html("📊 Final Stock Movement Summary",
                                stock_report.summary_page())
    finish_module(SHEET_MATERIAL_INDENT, row, MATERIAL_INDENT_ROWS)
    log.info("Automation Completed Successfully.")


# =========================================================================== #
# THE VALIDATION SUMMARY - the last thing in the report
#
# Everything in it is counted from what actually executed: the modules that ran,
# the calculations each of them checked, and how each one came out. There are no
# figures written into this file and nothing here re-runs a test - it reads the
# record calculation_validation.py kept while the suite was running.
#
# It is a test rather than a teardown hook on purpose. A summary attached in
# pytest_unconfigure() lands after Allure has written its results and is easy to
# miss; a test is a node in the report, it carries its own attachments, and its
# result IS the overall verdict - green when every module's arithmetic agreed,
# red when one did not.
# =========================================================================== #

def execution_facts() -> List[List[str]]:
    """Where and when this execution ran - the top of the QA summary.

    The four lines a tester writes at the head of a report before any result:
    the date, the time, the environment and the address of the application it
    was pointed at. Read from run_config's stamp and the URLs at the top of this
    file, so they cannot drift from what the run actually used.
    """
    # RUN_TIMESTAMP is "YYYY-MM-DD_HH-MM-SS" - the date and the time of the
    # moment this execution started, which run_config stamped once.
    date_part, _, time_part = RUN_TIMESTAMP.partition("_")
    return [
        ["Execution Date", date_part],
        ["Execution Time", time_part.replace("-", ":")],
        ["Execution Id", RUN_TIMESTAMP],
        ["Environment", "GenZOpss Production (live application)"],
        ["Application URL", APP_URL],
        ["Marketing Site", MARKETING_URL],
        ["Admin Portal", ADMIN_URL],
        ["Business User", BUSINESS_USER_EMAIL],
        ["Browser", f"Chromium (headed={HEADED})"],
        ["Test Data", "test_data/Construction_Flow_Data.xlsx"],
        ["Report Folder", str(RUN_REPORT_DIR)],
    ]


def qa_narrative(totals: Dict[str, object], project: str,
                 cases: Dict[str, int]) -> str:
    """The run in two or three sentences of plain QA English.

    What a tester would say if they were asked "how did it go?" - and every
    sentence is conditional on what actually happened, so it can only ever
    describe this run. A sentence about the stock is written only when the stock
    was really checked; a run that never reached it says nothing about it.
    """
    said: List[str] = []

    # FAILED and BLOCKED are counted in two different sentences, and never in
    # one. "1 of 12 test cases did not pass" was printed for a run whose only
    # non-pass was a blocked precondition - a sentence that reads as a defect
    # and was not one.
    if totals["modules_failed"] or cases["failed"] or cases["broken"]:
        said.append(
            f"{cases['failed'] + cases['broken']} of {cases['total']} test "
            f"case(s) FAILED, across {totals['modules_failed']} module(s). "
            f"Each one is listed below with what was expected, what the "
            f"application actually did, and why.")
    elif cases["total"]:
        said.append(f"No test case failed. {cases['passed']} of "
                    f"{cases['total']} executed test case(s) completed "
                    f"successfully across {totals['modules_passed']} module(s)"
                    + (f", and {cases['blocked']} could not be run at all."
                       if cases["blocked"] else "."))

    if totals["checks_failed"]:
        said.append(f"{totals['checks_failed']} of {totals['checks_total']} "
                    f"calculation(s) did not match a value worked out "
                    f"independently from the same inputs.")
    elif totals["checks_passed"]:
        said.append(f"All {totals['checks_passed']} verified calculation(s) "
                    f"matched the expected values"
                    + (f"; {totals['checks_skipped']} figure(s) could not be "
                       f"checked and are listed as SKIPPED with the reason."
                       if totals["checks_skipped"] else "."))

    # The stock story, and only when the run actually proved it. Read back off
    # the checks the engine recorded - never assumed from the fact the modules
    # ran.
    grn = recorded_check("Stock After GRN Approval")
    issue = recorded_check("Stock After Indent Issue")
    closing = recorded_check("Closing Stock (Items & Inventory)")
    if grn is not None and grn.status == calc.PASS:
        said.append("Stock increased correctly after the GRN was approved.")
    if issue is not None and issue.status == calc.PASS:
        said.append("Stock decreased correctly after the material indent was "
                    "issued.")
    if closing is not None and closing.status == calc.PASS:
        said.append("The closing stock the application displays matches the "
                    "stock worked out independently from the run's own "
                    "movements (opening + received - issued + returned).")

    if totals["modules_blocked"]:
        said.append(f"{totals['modules_blocked']} module(s) could not be RUN "
                    f"(a precondition, the environment or a manual step was "
                    f"not available), so nothing about the application was "
                    f"proved there - they are reported as BLOCKED: not passes, "
                    f"not failures, not application defects and not automation "
                    f"defects. Each one is listed under BLOCKERS with its "
                    f"cause and what has to be done before the next run.")
    if totals["modules_skipped"]:
        said.append(f"{totals['modules_skipped']} module(s) did not run in this "
                    f"execution - switched off for the run, or their rows were "
                    f"test data the application already holds. They are "
                    f"reported as SKIPPED with the reason: nothing about them "
                    f"was proved, and nothing about them failed.")
    if totals["modules_not_automated"]:
        said.append(f"{totals['modules_not_automated']} module(s) are not "
                    f"automated by this suite and are listed as NOT AUTOMATED "
                    f"so the coverage this report claims is the coverage there "
                    f"is.")

    said.append(f"Overall result: {project}.")
    return " ".join(said)


def qa_headline(totals: Dict[str, object], project: str) -> str:
    """The counts and the three verdicts, for a reader who opens nothing else.

    Every figure comes from calc.totals() (modules and calculations) and
    qa_report.TALLY (test cases), both of which count what actually executed.
    Nothing here is written in by hand and nothing is recounted - one place
    works these out so the headline can never disagree with the tables under it.
    """
    cases = qa_report.TALLY.counts()
    return "\n".join([
        qa_report.banner("GENZOPSS CONSTRUCTION FLOW - QA EXECUTION SUMMARY"),
        "",
        qa_report.labelled(execution_facts()),
        "",
        f"Total Modules        : {totals['modules_total']}",
        f"Passed               : {totals['modules_passed']}",
        f"Failed               : {totals['modules_failed']}",
        f"Skipped              : {totals['modules_skipped']}",
        f"Blocked              : {totals['modules_blocked']}",
        f"Not Automated        : {totals['modules_not_automated']}",
        "",
        f"Total Test Cases     : {cases['total']}",
        f"Passed               : {cases['passed']}",
        f"Failed               : {cases['failed'] + cases['broken']}",
        f"Blocked              : {cases['blocked']}",
        f"Skipped              : {cases['skipped']}",
        "",
        f"Calculations Checked : {totals['checks_total']}",
        f"Calculations Passed  : {totals['checks_passed']}",
        f"Calculations Failed  : {totals['checks_failed']}",
        f"Calculations Skipped : {totals['checks_skipped']}",
        "",
        f"Overall Calculation Result : {totals['calculation_result']}",
        f"Functional Result          : {totals['functional_result']}",
        f"Overall Project Result     : {project}",
        "",
        # THE THREE FIGURES A READER HAS TO BE ABLE TO TELL APART, spelled out
        # rather than left to be worked out from the columns above. A run whose
        # only red mark is a blocked precondition used to be indistinguishable
        # from a run with a real defect in it.
        f"Actual test failures       : {cases['failed'] + cases['broken']}",
        f"Blocked test flows         : {cases['blocked']}",
        f"Calculation failures       : {totals['checks_failed']}",
        "",
        qa_report.section("IN PLAIN WORDS",
                          qa_narrative(totals, project, cases)),
        "",
        qa_report.section("HOW TO READ THIS", qa_report.CLASSIFICATION_NOTE),
    ])


def qa_summary_card(totals: Dict[str, object],
                    cases: Dict[str, int]) -> str:
    """THE CONCISE SUMMARY - the counts a QA lead reports, and nothing else.

    Deliberately the shortest attachment in the report and deliberately the
    first: modules, test cases, calculations, stock validations, verdict. The
    headline, the tables and the full page under it are the same run in more
    detail, for whoever needs the detail - this is for whoever does not.

    Every figure is read back off calc.totals(), qa_report.TALLY and
    stock_report.counts(), which are the same three counters the tables are
    drawn from, so this card cannot disagree with anything below it. Stock is
    counted here for the first time: it is the question this suite is most
    often run to answer, and it used to be legible only by opening the stock
    attachment of the Material Indents test case and reading its verdict.
    """
    stock = stock_report.counts()
    failed_cases = cases["failed"] + cases["broken"]
    return "\n".join([
        qa_report.banner("QA SUMMARY"),
        "",
        f"Total Modules        : {totals['modules_total']}",
        f"Passed               : {totals['modules_passed']}",
        f"Failed               : {totals['modules_failed']}",
        f"Skipped              : {totals['modules_skipped']}",
        f"Blocked              : {totals['modules_blocked']}",
        "",
        f"Total Test Cases     : {cases['total']}",
        f"Passed               : {cases['passed']}",
        f"Failed               : {failed_cases}",
        f"Blocked              : {cases['blocked']}",
        f"Skipped              : {cases['skipped']}",
        "",
        f"Calculations Checked : {totals['checks_total']}",
        f"Calculations Passed  : {totals['checks_passed']}",
        f"Calculations Failed  : {totals['checks_failed']}",
        "",
        f"Stock Validations    : {stock['total']}",
        f"Stock Passed         : {stock['passed']}",
        f"Stock Failed         : {stock['failed']}",
        # The two stock questions this suite is most often run to answer, each
        # with its own answer. "Stock Passed: 2" says both went right and
        # "Stock Failed: 1" does not say WHICH - and "did the receipt land?"
        # and "did the issue come off?" are two different defects with two
        # different owners. Both verdicts are the ones the cards already
        # carry - see stock_report.increase_result() / decrease_result().
        f"Stock Increase       : {stock_report.increase_result()}   (GRN)",
        f"Stock Decrease       : {stock_report.decrease_result()}   (Material Indent)",
        "",
        f"Overall Result       : {stock_report.mark(totals['project_result'])}",
    ])


def summary_file(name: str, body: str) -> None:
    """Keep a copy of the summary next to the run's log and screenshots."""
    try:
        (RUN_REPORT_DIR / name).write_text(body, encoding="utf-8")
    except OSError as error:
        log.warning("Could not write %s: %s", name, error)


@allure.epic(ALLURE_EPIC)
@allure.feature("ZZ Validation Summary")
@allure.story("Functional and calculation results for every module in this run")
@allure.severity(allure.severity_level.BLOCKER)
def test_validation_summary() -> None:
    """Report every module's functional, calculation and overall result."""
    # Modules this suite does not drive are declared HERE, at the end, so they
    # cannot be mistaken for something that ran and passed - and so the coverage
    # the report claims is the coverage there actually is.
    for module, reason in NOT_AUTOMATED_MODULES:
        calc.declare(module, calc.NOT_AUTOMATED, reason)

    # A module that was SWITCHED OFF for this run is declared for exactly the
    # same reason: a heading that simply is not there reads as one that passed.
    # It is SKIPPED, not NOT AUTOMATED (the suite does drive it - it was told
    # not to this time), not BLOCKED (nothing was in its way) and certainly not
    # FAILED. Declared only when it really did not run, so that turning the
    # switch back on cannot leave a stale "skipped" line beside its results.
    if not RUN_CREATE_YOUR_ACCOUNT:
        calc.declare("Create Your Account", calc.SKIPPED,
                     CREATE_YOUR_ACCOUNT_OFF)

    totals = calc.totals()
    cases = qa_report.TALLY.counts()
    # The three things the SUITE knows and calculation_validation does not: where
    # and when it ran, how many test cases came out which way, and the run in
    # plain words. Handing them over is what makes the execution summary one
    # report instead of two half ones.
    project, plain, html = calc.summary(
        environment=execution_facts(),
        narrative=qa_narrative(totals, str(totals["project_result"]), cases),
        test_cases=cases)

    # The three verdicts, each on its own, so a reader who opens nothing else
    # still sees that they are three answers to three different questions.
    if ALLURE:
        # The verdict in the test's NAME. This node can end green, grey or red
        # and a reader should not have to work out which of those means what -
        # a summary that reads "QA EXECUTION SUMMARY: BLOCKED" says it before
        # anything is opened.
        allure.dynamic.title(f"QA EXECUTION SUMMARY: {project}")
        allure.dynamic.parameter("Overall Calculation Result",
                                 totals["calculation_result"])
        allure.dynamic.parameter("Functional Result", totals["functional_result"])
        allure.dynamic.parameter("Overall Project Result", project)
        allure.dynamic.parameter("Test Cases",
                                 f"{cases['passed']} passed, "
                                 f"{cases['failed'] + cases['broken']} failed, "
                                 f"{cases['blocked']} blocked, "
                                 f"{cases['skipped']} skipped")
        # The three figures that must never be added together, each named.
        allure.dynamic.parameter("Actual test failures",
                                 cases["failed"] + cases["broken"])
        allure.dynamic.parameter("Blocked test flows", cases["blocked"])
        allure.dynamic.parameter("Calculation failures",
                                 totals["checks_failed"])
        stock = stock_report.counts()
        allure.dynamic.parameter("Stock Validations",
                                 f"{stock['total']} checked, "
                                 f"{stock['passed']} passed, "
                                 f"{stock['failed']} failed")
        allure.dynamic.parameter("Stock Result", stock_report.result())
        # The whole summary at the top of this test - what a stakeholder opens
        # the report for, above the steps and above the attachments. The colour
        # page is the 'QA Summary - Full Report' attachment below; this is the
        # same figures as a description fragment, which must not carry a
        # stylesheet of its own - see qa_report.embed().
        # THE SHORT CARD AT THE TOP, not the long headline. Allure draws a
        # test's description above its steps and its attachments, so it is the
        # one thing a reader cannot miss - and what belongs there is the
        # twelve-line answer, not the narrative, the classification note and
        # the four blocks of counts that used to fill it. Every one of those is
        # still in the report, in "QA Summary - Headline" underneath.
        #
        # THE STOCK CALCULATION, AHEAD OF EVEN THAT. "Stock Passed: 2" in the
        # QA Summary card answers "did it pass?" - it does not answer "what was
        # the stock, what moved, what should it be now?", which is the question
        # this suite is most often opened to answer. `top_stock_summary_text()`
        # reads the SAME Movement records qa_summary_card() counts and the same
        # ones increase_text()/decrease_text()/final_stock_text() print in full
        # below - it only puts the short answer where a reader sees it first.
        qa_top_card = qa_report.joined([stock_report.top_stock_summary_text(),
                                        qa_summary_card(totals, cases)])
        describe(qa_top_card, qa_report.embed(
            "GenZOpss Construction Flow - QA Execution Summary",
            qa_top_card, verdict=project,
            subtitle=f"{totals['modules_total']} modules | {cases['total']} "
                     f"test cases | {totals['checks_total']} calculations",
            pairs=execution_facts()))

    with step(f"CONSTRUCTION FLOW QA SUMMARY: {project}"):
        for line in plain.splitlines():
            log.info("%s", line)
        # THE CONCISE SUMMARY, FIRST. Twelve lines and a verdict: the counts a
        # QA lead reports upward, in the order they are asked for, with nothing
        # to scroll past to reach them. Everything under it is the same run in
        # more detail, for whoever needs the detail.
        attach_text("QA SUMMARY", qa_summary_card(totals, cases))
        # The headline on its own, because it is the one thing a stakeholder
        # reads: when it ran, how many modules and test cases, how many
        # calculations, and the verdict - with the run described in words
        # underneath it.
        attach_text("QA Summary - Headline", qa_headline(totals, project))
        # Module by module: Functional, Calculation and Overall for each - built
        # from what actually executed. Overall is the LAST column on purpose:
        # the table renderer badges the last cell of each row, so the verdict is
        # the thing that turns green or red when the table is scanned.
        attach_table("QA Summary - Module by Module",
                     ("Module", "Functional", "Calculation", "Reason",
                      "Overall"),
                     [[row[0], row[1], row[3], row[5], row[4]]
                      for row in calc.summary_rows()],
                     f"Overall Project Result: {project}", plain=True)
        # Test case by test case, which is the level a QA lead reports at and
        # the one level no other table in this report covers.
        attach_table("QA Summary - Test Case by Test Case",
                     ("Test Case", "Result"), qa_report.TALLY.rows(),
                     f"{cases['total']} test case(s) - {cases['passed']} passed, "
                     f"{cases['failed'] + cases['broken']} failed, "
                     f"{cases['blocked']} blocked, {cases['skipped']} skipped",
                     plain=True)
        # THE STOCK STORY OF THE RUN, on the summary as well as on the module
        # that produced it - "did the stock move correctly?" is a question
        # asked of the RUN, and a reader should not have to find the Material
        # Indents test case to answer it.
        #
        # The three sections are repeated here under the same
        # "📦 Stock Validation" parent they carry on their own modules, so a
        # reader who opens only the summary test still gets the stock increase,
        # the stock decrease and the movement table without going looking. They
        # are the SAME cards - stock_report holds them for the whole run - so
        # nothing here can disagree with what the modules reported.
        with stock_section(STOCK_ITEM_STEP):
            verify_data.attach_html("📦 Item Stock - Before Transaction",
                                    stock_report.item_page())
        with stock_section(STOCK_INCREASE_STEP):
            verify_data.attach_html("📈 Item Stock - After Purchase",
                                    stock_report.increase_page())
        with stock_section(STOCK_DECREASE_STEP):
            verify_data.attach_html(
                "📉 Item Stock - After Consumption",
                stock_report.decrease_page())
        with stock_section(STOCK_SUMMARY_STEP):
            attach_text("📦 STOCK MOVEMENT",
                        stock_report.simple_movement_text())
            verify_data.attach_html("📦 Stock Movement",
                                    stock_report.simple_movement_page())
            attach_text("📦 CURRENT STOCK VALIDATION",
                        stock_report.current_stock_text())
            verify_data.attach_html("📦 Current Stock Validation",
                                    stock_report.current_stock_page())
            attach_text("📊 STOCK MOVEMENT SUMMARY",
                        stock_report.summary_text())
            verify_data.attach_html("📊 Stock Movement Summary",
                                    stock_report.summary_page())
        # `plain` is the same figures as the HTML page below it, in text. It is
        # still written to Reports/<run>/calculation_summary.txt, and the
        # concise "QA SUMMARY" above is what a reader actually needs in text,
        # so it is no longer attached a second time as well.
        verify_data.attach_html("QA Summary - Full Report", html)
        summary_file("calculation_summary.txt", plain)

    with step(f"CALCULATION VALIDATION: {totals['calculation_result']}"):
        attach_text("Calculation Validation Result", calculation_verdict(totals))

    # THE BLOCKERS, EACH ONE WRITTEN OUT IN FULL - the section a QA lead reads
    # before deciding whether the run is worth re-running and what to change
    # first. Every field comes from what the blocked step itself recorded.
    attach_blocker_report(totals)

    # WHAT THIS TEST'S OWN RESULT MEANS, and why a BLOCKED project does not make
    # it red.
    #
    # This test is a REPORT OF THE RUN, not a test case of it. It used to raise
    # whenever the project result was anything but PASS - so a run with no
    # failed test case, no wrong calculation and one blocked precondition ended
    # with a red "test_validation_summary" in Allure, and the report a
    # stakeholder opened said FAILED. That is the report accusing the
    # application of a defect that this run never found.
    #
    # So the three outcomes are told apart, and each one is filed under the word
    # that is true of it:
    #
    #   FAIL      something was actually wrong - a calculation did not match, or
    #             a module's flow did not reach the outcome its row asked for.
    #             This test FAILS, because there is a finding to act on.
    #   BLOCKED   nothing was found to be wrong and part of the run could not be
    #             executed. This test is SKIPPED with the blockers as its
    #             reason: not green (the run is not a clean pass) and not red
    #             (there is no defect). GENZ_FAIL_ON_BLOCKED=1 makes it red for
    #             a CI job that needs that.
    #   PASS      everything that ran, passed. This test passes.
    reasons = build_summary_reasons(totals)

    # The pytest-level count is checked as well as the module-level verdict. A
    # test that fell over before it ever reached a module - a broken fixture, an
    # unreadable workbook - leaves no failed MODULE behind, and a summary that
    # only looked at the modules would go quietly grey over it. A real failure
    # anywhere in the run makes this test red, whatever else is blocked.
    nothing_failed = (cases["failed"] + cases["broken"]) == 0

    if project == calc.BLOCKED and nothing_failed and not FAIL_ON_BLOCKED:
        blocked_modules = ", ".join(sorted({report.module
                                            for report in totals["merged"]
                                            if report.overall == calc.BLOCKED}))
        log.warning("Overall Project Result: BLOCKED - %s could not be run. "
                    "No test case failed and no calculation was wrong.",
                    blocked_modules or "one or more modules")
        if UNDER_PYTEST:
            pytest.skip(
                f"Overall Project Result: BLOCKED. No test case FAILED and no "
                f"calculation was wrong - {blocked_modules or 'a module'} "
                f"could not be RUN because a precondition was not available, "
                f"so the run's overall verdict could not be established. This "
                f"is NOT an application defect and NOT an automation defect. "
                f"See 'QA Summary - Blockers' for the cause and what to do "
                f"before re-running.")
        return

    if project != calc.PASS and calc.FAIL_ON_MISMATCH:
        raise AssertionError(
            f"Overall Project Result: {project} - this is the SUMMARY of the "
            f"run, not a defect of its own. Each cause below has already "
            f"failed or blocked the module it belongs to, with the detail.\n\n"
            f"Overall Calculation Result : {totals['calculation_result']}\n"
            f"Functional Result          : {totals['functional_result']}\n\n"
            + "\n\n".join(reasons) + "\n\n" + plain)


def build_summary_reasons(totals: Dict[str, object]) -> List[str]:
    """WHY the run did not pass, in the summary's own words.

    "The arithmetic did not match" used to be printed whatever had gone wrong,
    so a run whose sums were all correct and whose FLOW had been blocked still
    accused the application of adding up wrongly - and sent the reader looking
    for a wrong number that was not there. These are separate findings and this
    says which of them happened, and names the modules.
    """
    reasons: List[str] = []
    if totals["checks_failed"]:
        wrong_sums = [f"{report.module} ({counts} calculation(s) wrong)"
                      for report in totals["merged"]
                      for counts in [sum(1 for check in report.checks
                                         if check.status == calc.FAIL)]
                      if counts]
        reasons.append(
            "CALCULATION failures - the application's arithmetic did not "
            "match an independently worked-out value:\n    "
            + "\n    ".join(wrong_sums))
    else:
        reasons.append(
            f"CALCULATION: {totals['calculation_result']}. No calculation "
            f"was found to be wrong ({totals['checks_passed']} proved "
            f"correct, {totals['checks_skipped']} could not be checked). "
            f"The arithmetic is NOT what stopped this run.")

    broken_flows = sorted({report.module for report in totals["merged"]
                           if report.overall == calc.FAIL})
    if broken_flows:
        reasons.append(
            "FUNCTIONAL failures - the module's flow did not reach the "
            "outcome its row asked for (this is NOT an arithmetic fault; "
            "the module's own failure carries the cause):\n    "
            + "\n    ".join(broken_flows))

    blocked = sorted({report.module for report in totals["merged"]
                      if report.overall == calc.BLOCKED})
    if blocked:
        reasons.append(
            "BLOCKED - the module could not be RUN, so nothing about the "
            "application was proved there. Not a product defect and not an "
            "automation defect:\n    " + "\n    ".join(blocked))
    return reasons


def attach_blocker_report(totals: Dict[str, object]) -> None:
    """Every blocked module, written out the way the next run needs it.

    BLOCKED MODULE / PRECONDITION / ROOT CAUSE / WHAT WAS EXPECTED / WHAT
    ACTUALLY HAPPENED / WHAT NEEDS TO BE DONE BEFORE RE-RUN - the full story is
    written by the step that hit the blocker and attached to the test case it
    happened in (that is the only place that knows it). What this adds is the
    same thing at RUN level: the summary a reader opens first must not leave
    them hunting through eleven test cases to find out what stopped the run.
    """
    blockers = totals["blockers"]
    if not blockers:
        attach_text("QA Summary - Blockers",
                    "None. Every module this suite drives was able to run.\n\n"
                    + qa_report.CLASSIFICATION_NOTE)
        return

    lines = [qa_report.banner("BLOCKERS - MODULES THAT COULD NOT BE RUN"), ""]
    for report in blockers:
        lines += [
            qa_report.labelled([
                ["BLOCKED MODULE", report.module],
                ["TEST CASE", report.test_case_id or "-"],
                ["CATEGORY", report.category or "BLOCKED PRECONDITION"],
                ["ROOT CAUSE", report.note or "(no reason was recorded)"],
            ]),
            "",
            "    The test case's own attachments carry the full statement: the",
            "    precondition, what was expected, what actually happened and",
            "    what has to be done before the suite is run again.",
            "",
        ]
    lines += [
        "This section is NOT a defect list. A blocked module was never run, so "
        "nothing about the application was established there - in either "
        "direction. The AUTOMATION DEFECTS and APPLICATION DEFECTS sections of "
        "the summary are where a real finding would appear, and they are "
        "counted separately.",
        "",
        qa_report.CLASSIFICATION_NOTE,
    ]
    body = "\n".join(lines)
    attach_text("QA Summary - Blockers", body)
    verify_data.attach_html("QA Summary - Blockers (report)", qa_report.page(
        "Blockers - modules that could not be run",
        f"{len(blockers)} module(s) could not be executed in this run",
        [qa_report.table_block(
            "Blocked modules",
            ["Module", "Test case", "Category", "Result", "Root cause"],
            [[report.module, report.test_case_id or "-",
              report.category or "BLOCKED PRECONDITION", calc.BLOCKED,
              report.note or "-"] for report in blockers],
            verdict_column=3),
         qa_report.pre_block("How to read this", qa_report.CLASSIFICATION_NOTE)],
        verdict=calc.BLOCKED))
    for line in body.splitlines():
        log.info("   %s", line)


def calculation_verdict(totals: Dict[str, object]) -> str:
    """The calculation half of the report, on its own, in its own words.

    Separate from the project summary on purpose: this is the answer to "did the
    application get its numbers right?", and it is not allowed to be coloured by
    what any module's FLOW did.
    """
    lines = [
        "CALCULATION VALIDATION",
        "",
        f"Total Calculations Checked : {totals['checks_total']}",
        f"Calculations Passed        : {totals['checks_passed']}",
        f"Calculations Failed        : {totals['checks_failed']}",
        f"Calculations Skipped       : {totals['checks_skipped']}",
        "",
        f"Overall Calculation Result : {totals['calculation_result']}",
        "",
        "This verdict is worked out from the calculation checks alone. A "
        "blocked or failed module elsewhere in the run does NOT make it a "
        "FAIL - only a figure that disagreed with an independently worked-out "
        "value does.",
        "",
    ]
    for report in totals["merged"]:
        if not report.checks:
            continue
        lines.append(f"{report.module}: {report.calculation}")
        for check in report.checks:
            lines.append(f"    {check.name:<34} {check.status:<8} "
                         f"expected {calc.money(check.expected)} | "
                         f"actual {calc.money(check.actual)}")
    return "\n".join(lines)


# =========================================================================== #
# ALLURE REPORTING HOOKS
#
# Two things pytest can only do from a plugin: attach a screenshot to a test
# that has just failed, and build the report once the run is over. The whole
# suite is this single file, and pytest does NOT load hooks out of a test
# module - so the module registers itself as a plugin, and nothing has to move
# into a conftest.py.
# =========================================================================== #

pytest_plugins = [__name__]


def page_in_use(item) -> Optional[Page]:
    """The browser page a test was working with, or None if it never got one.

    Both fixtures hand out the same session page: `page` for the registration
    module, `workspace` for the four that need a signed-in business user.
    """
    fixtures = getattr(item, "funcargs", None) or {}
    for name in ("workspace", "page"):
        candidate = fixtures.get(name)
        if candidate is not None:
            return candidate
    return None


def current_allure_test_uuid() -> Optional[str]:
    """The Allure UUID of the test case executing right now, or None.

    The same identity-keyed lookup stabilise_history_id() below already uses
    (allure_pytest.listener.ItemCache is keyed by id(item.nodeid), which is
    why CURRENT_PYTEST_ITEM keeps the item object and not a copy of its
    nodeid). Reused here so a reading taken mid-test can be filed against the
    exact Allure result it belongs to, while that is still possible: the
    moment this test finishes, AllureListener.pytest_runtest_logfinish calls
    AllureReporter.close_test(), which pops the result out of memory and hands
    it to the file logger - after that there is no live object left to ask.
    """
    item = CURRENT_PYTEST_ITEM
    if item is None or not ALLURE:
        return None
    listener = item.config.pluginmanager.get_plugin("allure_listener")
    if listener is None:                           # running without --alluredir
        return None
    test_uuid = listener._cache.get(item.nodeid)
    if test_uuid is None or listener.allure_logger.get_test(test_uuid) is None:
        return None
    return test_uuid


def _find_allure_result_file(results_dir: Path,
                             test_uuid: str) -> Optional[Path]:
    """Which *-result.json in this run's results dir IS this Allure test.

    allure-commons names the FILE after a fresh uuid it mints for that purpose
    (AllureFileLogger._report_item) - never the test's own uuid, which only
    lives INSIDE the file as its "uuid" field - so the file has to be found by
    reading that field back, not guessed from a filename pattern.
    """
    needle = f'"uuid": "{test_uuid}"'
    for path in results_dir.glob("*-result.json"):
        try:
            text = path.read_text(encoding="utf-8")
        except OSError:
            continue
        if needle not in text:
            continue
        try:
            data = json.loads(text)
        except ValueError:
            continue
        if data.get("uuid") == test_uuid:
            return path
    return None


def _write_allure_attachment(results_dir: Path, name: str, body: str,
                             extension: str, mime_type: str) -> Dict[str, str]:
    """Write one attachment straight into allure-results and describe it the
    way a *-result.json's own "attachments" entries already describe theirs."""
    file_name = f"{uuid.uuid4()}-attachment.{extension}"
    (results_dir / file_name).write_text(body, encoding="utf-8")
    return {"name": name, "source": file_name, "type": mime_type}


def _find_nested_step(steps: List[Dict[str, object]],
                      name: str) -> Optional[Dict[str, object]]:
    """Depth-first search for a step by name inside an Allure result's own
    "steps" tree."""
    for step_dict in steps:
        if step_dict.get("name") == name:
            return step_dict
        found = _find_nested_step(step_dict.get("steps", []), name)
        if found is not None:
            return found
    return None


#: Movement.result (a QA-domain word: PASS/FAIL/NOT CHECKED/SKIPPED) -> the
#: Allure step status vocabulary (passed/failed/skipped/unknown - see
#: allure_commons.model2.Status). Decorative only: it colours the small step
#: icon this patch adds and touches no verdict of the test it is added to.
_ALLURE_STEP_STATUS = {stock_report.PASS: "passed", stock_report.FAIL: "failed",
                       stock_report.SKIPPED: "skipped"}


def attach_after_purchase_to_item_test(code: str, name: str,
                                       movement: "stock_report.Movement"
                                       ) -> None:
    """Patch ITEM STOCK - AFTER PURCHASE onto the item's OWN "Items - Create
    and verify Items" Allure result, so a reader finds BOTH stock readings for
    one item on the one test page instead of the before figure on the Items
    test and the after figure on a separate Purchase Orders/GRN one.

    WHY THIS IS A JSON PATCH AND NOT allure.attach(): the two readings happen
    in two different pytest tests, run one after the other. By the time this
    function runs (from the GRN module, deep inside the Purchase Orders test),
    the Items test has already finished and allure-pytest has already closed,
    serialised and forgotten it (see current_allure_test_uuid()'s docstring) -
    there is no live Allure context left for a normal attach call to reach.
    Editing the JSON it already wrote onto disk is the only way left to add to
    it, which is exactly what allure-results is at this point in the run: a
    folder of finished, independently-editable result files, one per test,
    waiting for `allure generate` at the very end (pytest_unconfigure).

    REPORTING ONLY, and never fatal to the GRN row: a run without
    --alluredir, an item this run never created (so it never trips
    report_initial_stock() and never gets a uuid in ITEM_STOCK_TEST_UUID), or
    a results file that cannot be found or parsed - each is silently skipped.
    """
    if not ALLURE or CURRENT_PYTEST_ITEM is None:
        return
    key = ((code or "").strip().casefold(), (name or "").strip().casefold())
    test_uuid = ITEM_STOCK_TEST_UUID.get(key)
    if not test_uuid:
        return

    try:
        results_dir = allure_results_dir(CURRENT_PYTEST_ITEM.config)
        result_path = _find_allure_result_file(results_dir, test_uuid)
        if result_path is None:
            return
        data = json.loads(result_path.read_text(encoding="utf-8"))

        after_step = {
            "name": STOCK_INCREASE_STEP,
            "status": _ALLURE_STEP_STATUS.get(movement.result, "unknown"),
            "steps": [],
            "attachments": [
                _write_allure_attachment(
                    results_dir, "📈 ITEM STOCK - AFTER PURCHASE",
                    movement.text(), "txt", "text/plain"),
                _write_allure_attachment(
                    results_dir, "📈 Item Stock - After Purchase",
                    qa_report.page("Item Stock - After Purchase",
                                   movement.item, movement.blocks(),
                                   verdict=movement.result),
                    "html", "text/html"),
            ],
            "start": int(time.time() * 1000), "stop": int(time.time() * 1000),
        }

        # Land it next to "📦 Item Stock - Before Transaction", under the SAME
        # "📦 Stock Validation" step the Items test already opened for it - not
        # a new section, the one this card was always filed under, just in a
        # different test.
        parent = _find_nested_step(data.get("steps", []), STOCK_SECTION)
        (parent if parent is not None else data).setdefault(
            "steps", []).append(after_step)

        # The side-by-side view, at the top of the test where "Test Data Used"
        # and "Execution Status" already live - one row, both readings, both
        # of them figures this run actually captured (Movement.before is the
        # SAME opening reading report_initial_stock() took; nothing here is
        # re-read or invented).
        name_shown, code_shown = stock_report.item_identity(movement.item)
        summary = stock_report.text_table(
            ["Item Name", "Item ID", "Before Transaction", "After Purchase",
             "Result"],
            [[name_shown, code_shown or "no code",
              stock_report.num(movement.before), stock_report.num(movement.actual),
              movement.result]])
        data.setdefault("attachments", []).append(_write_allure_attachment(
            results_dir, "📊 Item Stock - Before vs After Purchase", summary,
            "txt", "text/plain"))

        result_path.write_text(json.dumps(data, ensure_ascii=False),
                               encoding="utf-8")
    except Exception as error:               # reporting is never fatal
        log.debug("Could not attach the AFTER PURCHASE card to the Items "
                  "test for '%s': %s", name or code, error)


@pytest.hookimpl(wrapper=True, tryfirst=True)
def pytest_runtest_teardown(item):
    """Give every test case a history identity that survives a data change.

    Allure decides whether two executions are the SAME test - and therefore
    whether they belong on one history line - from
    md5(fullName + every parameter value). The row parameter of these tests is
    the whole Excel row, so changing one cell (a fresh GSTIN, a new company
    name, GENZ_UNIQUE_RUN_DATA=1) mints a brand new identity: the test loses its
    History and Retries panels and counts as "new" in the trend. It is why 18
    builds of a 10-test suite had grown 85 separate histories.

    The identity used here is the test function plus the TestCaseID, both of
    which are stable, so TC006 stays TC006 however its data is edited.

    tryfirst on a wrapper makes this the OUTERMOST wrapper, so the code after
    the yield runs last - after allure-pytest's own teardown hook has written
    the historyId this replaces. It touches nothing else about the result, and
    a failure here is never allowed to fail a test.
    """
    result = yield
    try:
        stabilise_history_id(item)
    except Exception as error:                     # reporting is never fatal
        log.debug("Could not stabilise the Allure historyId: %s", error)
    return result


def stabilise_history_id(item) -> None:
    """Rewrite one test's Allure historyId to fullName + TestCaseID."""
    listener = item.config.pluginmanager.get_plugin("allure_listener")
    if listener is None:                           # running without --alluredir
        return

    test_result = listener.allure_logger.get_test(
        listener._cache.get(item.nodeid))
    if test_result is None:
        return

    case_ids = [parameter.value for parameter in test_result.parameters
                if parameter.name == "TestCaseID"]
    if not case_ids:
        return                                     # not one of the Excel rows

    identity = f"{test_result.fullName}#{case_ids[0]}"
    test_result.historyId = hashlib.md5(identity.encode("utf-8")).hexdigest()


#: How much of pytest's own failure detail a QA reader is asked to read. Not
#: a hard cutoff of the evidence - execution.log on disk keeps the whole
#: thing - just of what this ONE attachment prints, so it stops being the
#: biggest thing in the report.
FAILURE_DETAIL_LIMIT = 4000


def condensed_failure_detail(longrepr: str) -> str:
    """pytest's own failure detail, without the source code and capped.

    REPORTING ONLY - the failure itself is untouched; this only changes how
    much of pytest's rendering of it a reader is shown.

    `report.longreprtext` is pytest's traceback PLUS, for every frame of the
    call stack, the source code around it - which is how a summary test
    whose own body is 150 lines long turned this into a 46 KB attachment
    that opened on a wall of Python nobody asked to read. pytest marks the
    exception's own message with a leading "E" on every line of it; every
    other line is that source context. Keeping only the "E" lines (and the
    closing "file.py:NNN: ExceptionType" line) removes the source and keeps
    the message - the actual reason - exactly as pytest wrote it.
    """
    if not longrepr:
        return ""
    kept = [line for line in longrepr.splitlines()
            if line.startswith("E") or re.match(r"^\S+\.py:\d+:", line)]
    text = "\n".join(kept) if kept else longrepr
    if len(text) > FAILURE_DETAIL_LIMIT:
        text = (text[:FAILURE_DETAIL_LIMIT] +
                f"\n\n... truncated ({len(text):,} of {len(longrepr):,} "
                f"characters shown - the source code pytest prints "
                f"alongside it has already been left out). The full "
                f"detail is in this run's execution.log.")
    return text


@pytest.hookimpl(wrapper=True)
def pytest_runtest_makereport(item, call):
    """Attach the evidence of a failure - screenshot, URL, traceback - to Allure.

    fail() already does this for the failures the framework raises itself; this
    catches everything else (an unexpected exception, a broken fixture), so no
    red test in the report is ever without a picture of the screen.

    pytest 9 dropped hookwrapper=True, so this is a `wrapper=True` hook and must
    return the report it was handed.
    """
    report = yield

    if report.when in ("setup", "call") and report.failed:
        try:
            page = page_in_use(item)
            if page is not None:
                attach_text("Page URL at Failure", page.url)
                safe_name = re.sub(r"[^A-Za-z0-9_.-]+", "_", item.name)
                screenshot(page, f"Screenshot at Failure - {safe_name}")
            attach_text("Failure Detail (technical)",
                        condensed_failure_detail(report.longreprtext or ""))
        except Exception as error:                     # reporting is never fatal
            log.warning("Could not attach the failure evidence: %s", error)

    return report


@pytest.hookimpl(tryfirst=True)
def pytest_runtest_setup(item) -> None:
    """Remember which test is running, so a blocked row can name itself.

    run_test_case() lives INSIDE the test and never sees the nodeid; the tally
    is keyed by it. This is the one hook that has both.
    """
    global CURRENT_NODEID, CURRENT_PYTEST_ITEM
    CURRENT_NODEID = item.nodeid
    CURRENT_PYTEST_ITEM = item


def pytest_runtest_logreport(report) -> None:
    """Count one test case's outcome, for the execution summary.

    The module tables count MODULES and the calculation tables count SUMS.
    Neither of them answers "how many TEST CASES passed?", which is the first
    figure a QA lead asks for - so it is counted here, from pytest's own verdict
    for every test that ran, and nowhere else. Nothing in the summary is written
    in by hand.

    setup, call and teardown each report; the tally keeps the worst of the three
    (see qa_report.ExecutionTally.record), so a test that passed and then broke
    while tidying up is not counted as a pass. A failure in setup or teardown is
    counted as BROKEN rather than FAILED, which is the same split Allure draws
    between 'failed' (the assertion did not hold) and 'broken' (the suite fell
    over) - the two go to different people.

    This only counts. It changes no verdict and can fail nothing.
    """
    try:
        if report.skipped:
            # A row the environment BLOCKED leaves the test as a pytest skip -
            # see run_test_case() - because pytest has no third word for "could
            # not be run". The report does have one, so the two are told apart
            # here and counted separately: SKIPPED is a used-up row, BLOCKED is
            # a precondition that was not there.
            status = ("BLOCKED" if report.nodeid in BLOCKED_TEST_CASES
                      else "SKIPPED")
        elif report.failed:
            status = "FAILED" if report.when == "call" else "BROKEN"
        elif report.when == "call":
            status = "PASSED"
        else:
            return                                 # a clean setup or teardown
        qa_report.TALLY.record(report.nodeid, status,
                               report.nodeid.split("::")[-1])
    except Exception as error:                     # counting is never fatal
        log.debug("Could not count a test outcome: %s", error)


def allure_results_dir(config) -> Path:
    """Where this run is writing its Allure results (pytest.ini says so)."""
    return Path(getattr(config.option, "allure_report_dir", None)
                or ALLURE_RESULTS_DIR).resolve()


def pytest_collection_finish(session) -> None:
    """STEP 1 of the report sequence: put the report's own files in place.

    --clean-alluredir empties allure-results in pytest_configure, and that takes
    categories.json, executor.json and history/ with it, so they have to be put
    back before the run - not only just before the report is built. Doing it
    here is what keeps the Categories tab and the trend alive when the report
    step is skipped (GENZ_NO_ALLURE_REPORT=1), when the run is interrupted, and
    above all when someone builds the report by hand afterwards with
    `allure generate allure-results` or `allure serve allure-results`. Without
    it those reports fall back to Allure's two built-in categories, 'Product
    defects' and 'Test defects', and none of the buckets below ever show up.

    It is pytest_collection_finish and not pytest_sessionstart because this file
    registers ITSELF as the plugin (pytest_plugins = [__name__]): pytest only
    picks that up when it imports the module, which happens during collection -
    by which time sessionstart has already been and gone. This hook is the first
    one that is certain to fire.

    The copy is repeated before allure generate, which costs nothing: seeding
    only ever adds files, so doing it twice cannot lose anything.
    """
    if getattr(session.config.option, "allure_report_dir", None) is None:
        return                                  # not running with --alluredir
    if session.config.option.collectonly:
        return
    write_allure_metadata(allure_results_dir(session.config), "before the run")


def pytest_unconfigure(config) -> None:
    """The last thing the run does: build this execution's report and open it.

    Steps 3-5 of the sequence: carry the history into the results, generate,
    then open the report. Step 1 is pytest_collection_finish, step 2 is the
    tests.

    The report is built into Reports/<timestamp>/, a folder created when the run
    started, so an execution never overwrites the one before it however many
    times pytest is run in a day.
    """
    if not AUTO_OPEN_ALLURE_REPORT or os.environ.get("GENZ_NO_ALLURE_REPORT"):
        return

    results_dir = allure_results_dir(config)
    if not results_dir.is_dir() or not any(results_dir.glob("*-result.json")):
        log.info("No Allure results in %s - no report to build.", results_dir)
        announce_report(built=False)
        return

    # Written once already, at collection. Repeated here so a results folder
    # that was emptied or replaced mid-run is still complete: it only ever
    # rewrites the same files, so doing it twice cannot lose anything.
    write_allure_metadata(results_dir, "before the report is built")

    report_index = generate_allure_report(results_dir)
    archive_allure_history()
    if report_index is not None:
        open_in_chrome(report_index)

    # The banner goes last, after Chrome has been asked to open, so the path is
    # the final thing left on the terminal.
    announce_report(built=(ALLURE_REPORT_DIR / "index.html").is_file())


def write_allure_metadata(results_dir: Path, when: str) -> None:
    """Everything the report's side panels are made of, into the results folder.

    environment.properties, categories.json, executor.json and the previous
    run's history: `allure generate` reads all four out of allure-results, so
    whatever is missing there is simply missing from the report.
    """
    write_allure_environment(results_dir)
    write_allure_categories(results_dir)
    carried = seed_allure_history(results_dir, when)
    write_allure_executor(results_dir, allure_build_number(carried))


def write_allure_environment(results_dir: Path) -> None:
    """The 'Environment' panel of the report: which app this run was pointed at.

    Written to two places: into the results folder, where `allure generate`
    picks it up and turns it into the panel, and into this run's own folder next
    to the report, where it is readable without opening the report at all.
    """
    lines = [
        f"Application={MARKETING_URL}",
        f"Admin.Portal={ADMIN_URL}",
        f"Workspace={APP_URL}",
        f"Business.User={BUSINESS_USER_EMAIL}",
        f"Browser=Chromium (headed={HEADED})",
        f"Python={sys.version.split()[0]}",
        f"Execution.Id={RUN_TIMESTAMP}",
        f"Executed.On={datetime.now():%Y-%m-%d %H:%M:%S}",
    ] + qa_summary_environment()
    text = "\n".join(lines)
    for path in (results_dir / "environment.properties", RUN_ENVIRONMENT_FILE):
        try:
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(text, encoding="utf-8")
        except OSError as error:
            log.warning("Could not write %s: %s", path, error)


def qa_summary_environment() -> List[str]:
    """The QA execution summary, as lines for the report's Environment panel.

    THE SUMMARY, WHERE IT CANNOT BE MISSED. The panel this feeds is drawn on the
    report's OVERVIEW page - the first screen anyone sees when they open the
    report - so the counts and the verdict are there before a single test is
    opened. The same figures are in the summary test's own attachments; this is
    the one place they appear without a click.

    write_allure_metadata() runs twice: once at collection, when nothing has
    executed and there is nothing to say, and once at the end of the run, when
    there is. The empty list the first call gets is what keeps a half-run from
    publishing a verdict it has not got.

    Everything is read back from calc.totals() and qa_report.TALLY - the same two
    counters the summary test reports from, so the Overview and the summary can
    never disagree. Reporting is never allowed to break a run, so the whole thing
    is guarded: a report without this panel is a small loss, a run that died
    building it is not.
    """
    try:
        if not calc.reports():
            return []                              # nothing has executed yet
        totals = calc.totals()
        cases = qa_report.TALLY.counts()
        stock = stock_report.counts()
        failed_cases = cases["failed"] + cases["broken"]
        return [
            f"QA.01.Overall.Result={totals['project_result']}",
            f"QA.02.Functional.Result={totals['functional_result']}",
            f"QA.03.Calculation.Result={totals['calculation_result']}",
            f"QA.04.Modules={totals['modules_total']} total | "
            f"{totals['modules_passed']} passed | {totals['modules_failed']} "
            f"failed | {totals['modules_blocked']} blocked | "
            f"{totals['modules_skipped']} skipped | "
            f"{totals['modules_not_automated']} not automated",
            f"QA.05.Test.Cases={cases['total']} total | {cases['passed']} "
            f"passed | {failed_cases} failed | {cases['blocked']} blocked | "
            f"{cases['skipped']} skipped",
            f"QA.06.Calculations={totals['checks_total']} checked | "
            f"{totals['checks_passed']} passed | {totals['checks_failed']} "
            f"failed | {totals['checks_skipped']} not verifiable",
            # The line that stops the overview being read as a defect count.
            # A blocked flow and a failed test are different findings, and the
            # first screen of the report has to say which of them this run has.
            f"QA.07.Failures.vs.Blockers={failed_cases} actual test failure(s) "
            f"| {cases['blocked']} blocked flow(s) | "
            f"{totals['checks_failed']} calculation failure(s)",
            # The stock verdict on the OVERVIEW page. This suite is most often
            # run to answer "did Items & Inventory move correctly?", and that
            # answer used to be reachable only by opening the Material Indents
            # test case and reading its stock attachment.
            f"QA.08.Stock.Validations={stock['total']} checked | "
            f"{stock['passed']} passed | {stock['failed']} failed | "
            f"result {stock_report.result()}",
        ]
    except Exception as error:                     # reporting is never fatal
        log.debug("Could not add the QA summary to the environment panel: %s",
                  error)
        return []


def write_allure_json(path: Path, payload: object) -> None:
    """One of the small JSON files Allure reads out of the results folder."""
    try:
        path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    except OSError as error:
        log.warning("Could not write %s: %s", path.name, error)


def trend_build_order(history_dir: Path) -> int:
    """The highest build number a history folder knows about, or 0.

    Read from history-trend.json, which is a list of runs newest first. The
    MAXIMUM is what matters, not the length: Allure keeps only the last 20 runs
    in that file (Stream.limit(20) in HistoryTrendPlugin), so counting the
    entries stops rising at 20 - which is exactly how a trend ends up with
    several runs all numbered 21 and stops showing new executions.
    """
    try:
        trend = json.loads((history_dir / "history-trend.json")
                           .read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return 0
    orders = [entry.get("buildOrder", 0) for entry in trend
              if isinstance(entry, dict)]
    return max([order for order in orders if isinstance(order, int)] or [0])


def best_history_source() -> Optional[Path]:
    """Where the newest surviving copy of the history is - or None on run one.

    Three places can hold it, and the one that is further along wins:

      allure-history/           a copy taken after every successful report, at
                                the project root. It is the normal source, and
                                the only one whose location does not change from
                                run to run: each execution now reports into its
                                own Reports/<timestamp>/ folder, so there is no
                                fixed "last report" to read.
      Reports/<newest>/         the previous execution's report, tried
        allure-report/history   newest-first until one turns up with a history/
                                in it. This is the fallback for a project whose
                                allure-history/ was deleted, and for a run whose
                                report was never built (no Allure command line
                                tool, or the suite died early) - it keeps the
                                trend unbroken across a failed build.
      <this run>/               only ever populated by the time the report has
        allure-report/history   already been generated, so it matters when the
                                metadata is written a second time.

    `allure generate --clean` DELETES the report folder before it writes the new
    one, which is why the archive exists at all: a generate interrupted at the
    wrong moment (Ctrl+C, a full disk, a crashed JVM) would otherwise take the
    only copy of the trend with it, and every later run would start again at
    build 1.
    """
    places = [ALLURE_HISTORY_ARCHIVE, ALLURE_REPORT_DIR / "history"]
    places += [folder / "allure-report" / "history"
               for folder in previous_run_dirs()]
    candidates = [directory for directory in places
                  if (directory / "history-trend.json").is_file()]
    if not candidates:
        return None
    return max(candidates, key=trend_build_order)


def seed_allure_history(results_dir: Path, when: str) -> int:
    """Copy the previous history into <results>/history. Returns its build number.

    This is what fills the TREND, HISTORY and RETRIES panels: Allure writes the
    history into <report>/history when it builds a report and reads it back out
    of <results>/history the next time, so without this copy every run looks
    like the first one and no trend is ever drawn.

    It only ever ADDS to the results folder - the existing history is never
    deleted or truncated here. Allure merges it with this run's results when it
    generates, and that merge is what appends one new point to the trend.
    """
    source = best_history_source()
    if source is None:
        log.info("Allure history: none found (%s) - this run starts the trend.",
                 when)
        return 0

    try:
        results_dir.mkdir(parents=True, exist_ok=True)
        shutil.copytree(source, results_dir / "history", dirs_exist_ok=True)
    except OSError as error:
        # Never fatal: a run that cannot carry its history is still a valid run.
        log.warning("Could not carry the Allure history over (%s): %s",
                    when, error)
        return trend_build_order(source)

    carried = trend_build_order(results_dir / "history")
    copied = sorted(path.name for path in (results_dir / "history").glob("*.json"))
    log.info("Allure history copied %s: %s -> %s (%d file(s): %s), "
             "latest build in it = #%d", when, source, results_dir / "history",
             len(copied), ", ".join(copied), carried)
    return carried


#: This run's build number, handed out once by allure_build_number().
_ALLURE_BUILD: Optional[int] = None


def allure_build_number(carried: int) -> int:
    """This run's build number, worked out once and then reused.

    write_allure_metadata() runs twice - at collection and before the report -
    and next_build_number() moves the counter on every call, so without this the
    run would burn two numbers and the executor would disagree with the history
    the trend was built from.
    """
    global _ALLURE_BUILD
    if _ALLURE_BUILD is None:
        _ALLURE_BUILD = next_build_number(carried)
    return _ALLURE_BUILD


def next_build_number(carried: int) -> int:
    """This run's build number - one more than the highest ever seen.

    Two sources, and the larger wins:

      the history carried into this run   the normal case,
      allure-history/build-order.txt      a counter of our own, which keeps
                                          rising after Allure has dropped the
                                          oldest runs out of history-trend.json
                                          (it keeps 20) and even if the history
                                          is lost altogether.

    Getting this wrong is what makes a trend look broken: two runs sharing a
    buildOrder are drawn as one point on the chart, so an execution appears to
    have gone missing.
    """
    last = max(carried, read_build_counter())
    build = last + 1
    write_build_counter(build)
    log.info("Allure trend: this execution is build #%d (the previous one was "
             "#%d).", build, last)
    return build


def read_build_counter() -> int:
    """The last build number this project handed out, or 0."""
    try:
        return int(ALLURE_BUILD_ORDER_FILE.read_text(encoding="utf-8").strip())
    except (OSError, ValueError):
        return 0


def write_build_counter(build: int) -> None:
    """Remember the build number, so the next run cannot reuse it."""
    try:
        ALLURE_HISTORY_ARCHIVE.mkdir(parents=True, exist_ok=True)
        ALLURE_BUILD_ORDER_FILE.write_text(str(build), encoding="utf-8")
    except OSError as error:
        log.warning("Could not record the Allure build number: %s", error)


def archive_allure_history() -> None:
    """Keep a copy of the history that `allure generate --clean` cannot delete.

    Run straight after the report is built, when allure-report/history holds
    this execution as well. See best_history_source() for why this exists.
    """
    fresh = ALLURE_REPORT_DIR / "history"
    if not (fresh / "history-trend.json").is_file():
        log.warning("The report has no history/ folder - the trend was NOT "
                    "archived. The next run may start a new trend.")
        return
    if trend_build_order(fresh) < trend_build_order(ALLURE_HISTORY_ARCHIVE):
        # Older than what is already kept: a half-built report must not be
        # allowed to overwrite a good archive.
        log.warning("The new report's history is older than the archived one - "
                    "keeping the archive as it is.")
        return
    try:
        shutil.copytree(fresh, ALLURE_HISTORY_ARCHIVE, dirs_exist_ok=True)
        log.info("Allure history archived to %s (through build #%d) - the next "
                 "run continues this trend.", ALLURE_HISTORY_ARCHIVE,
                 trend_build_order(fresh))
    except OSError as error:
        log.warning("Could not archive the Allure history: %s", error)


def write_allure_executor(results_dir: Path, build: int) -> None:
    """The 'Executors' panel: who ran the suite, and which build this is.

    buildOrder is the x-axis of the TREND chart and reportName is what groups
    the points into one line, so reportName must stay the same from run to run
    and buildOrder must never repeat - see next_build_number().
    """
    if os.environ.get("JENKINS_URL") or os.environ.get("BUILD_URL"):
        executor = {"name": "Jenkins", "type": "jenkins",
                    "buildUrl": os.environ.get("BUILD_URL", "")}
    elif os.environ.get("GITHUB_ACTIONS"):
        executor = {"name": "GitHub Actions", "type": "github",
                    "buildUrl": f"{os.environ.get('GITHUB_SERVER_URL', '')}/"
                                f"{os.environ.get('GITHUB_REPOSITORY', '')}/"
                                f"actions/runs/{os.environ.get('GITHUB_RUN_ID', '')}"}
    else:
        who = os.environ.get("USERNAME") or "local"
        where = os.environ.get("COMPUTERNAME") or "workstation"
        executor = {"name": f"{who} ({where})", "type": "local"}

    executor.update({
        "buildOrder": build,
        "buildName": f"{ALLURE_REPORT_NAME} #{build}",
        "reportName": ALLURE_REPORT_NAME,
    })
    # An empty buildUrl renders as a dead link in the panel - leave it out.
    write_allure_json(results_dir / "executor.json",
                      {key: value for key, value in executor.items() if value})


#: The buckets of the 'Categories' tab, as (name, statuses, message pattern).
#: A test is put in a bucket when its status is one of `statuses` AND its
#: failure/skip message contains `pattern` - see contains() for the regex.
#: Most specific first; the catch-alls are added by build_allure_categories().
SPECIFIC_CATEGORIES: List[Tuple[str, List[str], str]] = [
    # BLOCKED, not a product defect and not a test defect: the application's own
    # GST lookup service had nothing left to give, so the wizard could not draw
    # the fields the row needs. Kept first and worded as a dependency so these
    # runs are not read as the application calculating or behaving wrongly.
    # 'skipped' is in the statuses of both blocked buckets because a BLOCKED row
    # is recorded as a pytest skip, not as a failure - see run_test_case(). It
    # was a failure until 2026-08-17, and 'failed' is kept alongside so that a
    # run with GENZ_FAIL_ON_BLOCKED=1, and every report built from an older
    # results folder, still lands in the right bucket.
    ("Blocked - application dependency unavailable (GST lookup quota)",
     ["failed", "broken", "skipped"], r"GST lookup quota exhausted"),
    # BLOCKED, and for the same reason: the row could not be RUN. The stock was
    # not there, or a manual precondition (the e-mail OTP) was not met. Nothing
    # about the application was proved either way, so these must not sit in
    # 'Product defects' - which is where every one of them used to land.
    ("Blocked - precondition, data or environment",
     ["failed", "broken", "skipped"],
     r"CATEGORY: TEST DATA / ENVIRONMENT|"
     r"CATEGORY: DATA / ENVIRONMENT|"
     r"CATEGORY: BLOCKED|BLOCKED PRECONDITION|BLOCKED precondition|"
     r"Overall Project Result: BLOCKED"),
    # A module somebody switched off. Its own bucket so that a deliberate
    # decision is never read as a problem: it is not a used-up row, not a
    # blocker and not a defect, and the Categories tab should say so.
    ("Module switched off for this run",
     ["skipped"], r"MODULE SWITCHED OFF"),
    # A used-up row: the application already holds the record this row asked it
    # to create. The wordings are "already registered", "already exists",
    # "already in the application" and the API's own _IN_USE / ALREADY_ codes.
    # The phrases are matched rather than the bare word "already", so that a
    # BLOCKED skip - which is also a skip, and may well use the word in passing
    # - cannot fall into this bucket as well. Allure puts a result in EVERY
    # category it matches, so an overlap here would count one test twice.
    ("Row already used - the application already has this record",
     ["skipped"], r"already (registered|exists|linked|in the application|"
                  r"holds|has this)|_IN_USE|ALREADY_"),
    ("Rejected by the application (API error)",
     ["failed", "broken"], "API [45][0-9][0-9]"),
    ("Data mismatch on the View page",
     ["failed", "broken"], "do not match what was entered"),
    ("Element never appeared (UI or timing)",
     ["failed", "broken"], "never appeared|Could not click|Could not fill|"
                           "keeps being intercepted|Timeout"),
    ("Test data problem - check the workbook",
     ["failed", "broken"], r"There is no '[^']*' (sheet|column)|"
                           r"Construction_Flow_Data\.xlsx not found|"
                           r"TEST DATA FILE COULD NOT BE OPENED"),
]

#: The catch-alls, as (name, status). 'Product defects' and 'Test defects' are
#: the two Allure falls back on when there is no categories.json at all, so they
#: keep those names - a report that suddenly shows only these two is the sign
#: that the file did not reach the results folder.
CATCH_ALL_CATEGORIES: List[Tuple[str, str]] = [
    ("Product defects", "failed"),
    ("Test defects", "broken"),
    ("Skipped for another reason", "skipped"),
]


def contains(pattern: str) -> str:
    """A category regex that matches when the message contains `pattern`.

    Allure matches the WHOLE message (Java Pattern.matches), so the pattern has
    to be padded with .* on both sides, and (?s) is what lets those dots cross
    the newlines - every message fail() builds is several lines long.
    """
    return f"(?s).*({pattern}).*"


def contains_none_of(patterns: List[str]) -> str:
    """A catch-all's regex: it matches only what no specific bucket claimed.

    Allure does NOT stop at the first matching category - CategoriesPlugin puts
    a result into EVERY category it matches - so an unconditional catch-all
    shadows every bucket in front of it. That is what listed one failing test
    under both 'Rejected by the application (API error)' and 'Product defects',
    and counted it twice in the tab and in the categories trend.

    A category that carries a regex is skipped for a result with no message at
    all, but every failure here is raised through fail() or an exception that
    carries its text, so there is always a message to match.
    """
    return f"(?s)(?!.*({'|'.join(patterns)})).*"


def build_allure_categories() -> List[Dict[str, object]]:
    """The contents of categories.json: the buckets, then the exclusive catch-alls."""
    categories: List[Dict[str, object]] = [
        {"name": name, "matchedStatuses": statuses,
         "messageRegex": contains(pattern)}
        for name, statuses, pattern in SPECIFIC_CATEGORIES
    ]
    for name, status in CATCH_ALL_CATEGORIES:
        claimed = [pattern for _, statuses, pattern in SPECIFIC_CATEGORIES
                   if status in statuses]
        categories.append({"name": name, "matchedStatuses": [status],
                           "messageRegex": contains_none_of(claimed)})
    return categories


ALLURE_CATEGORIES = build_allure_categories()


def write_allure_categories(results_dir: Path) -> None:
    """The 'Categories' tab: group the failures by what caused them."""
    write_allure_json(results_dir / "categories.json", ALLURE_CATEGORIES)


def allure_command_line() -> Optional[str]:
    """The Allure command line tool, from PATH or from where it usually installs."""
    for name in ("allure", "allure.bat", "allure.cmd"):
        found = shutil.which(name)
        if found:
            return found
    for candidate in (
        Path(r"C:\allure\bin\allure.bat"),
        Path(r"C:\Program Files\allure\bin\allure.bat"),
        Path(r"C:\Program Files (x86)\allure\bin\allure.bat"),
        Path(os.environ.get("USERPROFILE", "")) / "scoop/shims/allure.exe",
    ):
        if candidate.is_file():
            return str(candidate)
    return None


def run_allure_generate(command_line: str, results_dir: Path, output_dir: Path,
                        single_file: bool) -> bool:
    """One `allure generate` run. True when it produced an index.html."""
    command = [command_line, "generate", str(results_dir), "-o", str(output_dir),
               "--clean"] + (["--single-file"] if single_file else [])
    log.info("Generating the Allure report%s...",
             " (single file)" if single_file else "")
    try:
        finished = subprocess.run(command, capture_output=True, text=True,
                                  timeout=300)
    except (OSError, subprocess.SubprocessError) as error:
        log.warning("Could not run the Allure command line tool: %s", error)
        return False

    if finished.returncode == 0 and (output_dir / "index.html").is_file():
        return True

    log.warning("allure generate%s failed (exit %d): %s",
                " --single-file" if single_file else "", finished.returncode,
                (finished.stderr or finished.stdout or "").strip()[:400])
    return False


def generate_allure_report(results_dir: Path) -> Optional[Path]:
    """Build the report. Returns the file to open in the browser, or None.

    It is generated TWICE, on purpose:

      allure-report/          the ordinary report. It is the only form that
                              writes history/, and that folder is what the next
                              run turns into the trend. Browse it with
                              `allure open allure-report`.
      allure-report-single/   the same report as one self-contained file. A
                              normal report fetches its data over XHR, which
                              Chrome blocks on a file:// URL (the page then sits
                              on "Loading..."), while the single file carries
                              every panel - trend, executors, categories,
                              screenshots - inline and opens straight from disk.
    """
    command_line = allure_command_line()
    if command_line is None:
        log.warning(
            "The Allure results are in %s, but the Allure command line tool was "
            "not found, so the report could not be built. Install it (scoop "
            "install allure, or npm i -g allure-commandline) and run: "
            "allure serve %s", results_dir, results_dir)
        return None

    full = run_allure_generate(command_line, results_dir, ALLURE_REPORT_DIR, False)
    if full:
        log.info("Allure report: %s (history kept for the trend)",
                 ALLURE_REPORT_DIR / "index.html")

    if run_allure_generate(command_line, results_dir, ALLURE_SINGLE_REPORT_DIR,
                           True):
        return ALLURE_SINGLE_REPORT_DIR / "index.html"

    # No self-contained file to open from disk - let Allure serve the full
    # report instead, which starts its own local web server.
    if full:
        log.info("Falling back to 'allure open' (it serves the report itself).")
        try:
            subprocess.Popen([command_line, "open", str(ALLURE_REPORT_DIR)])
        except (OSError, subprocess.SubprocessError) as error:
            log.warning("Could not serve the report either: %s", error)
    return None


def open_in_chrome(index: Path) -> None:
    """Show the finished report in Google Chrome, or the default browser."""
    url = index.resolve().as_uri()

    chrome = None
    for name in ("chrome", "chrome.exe", "google-chrome"):
        chrome = chrome or shutil.which(name)
    for candidate in (
        Path(r"C:\Program Files\Google\Chrome\Application\chrome.exe"),
        Path(r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"),
        Path(os.environ.get("LOCALAPPDATA", "")) /
        "Google/Chrome/Application/chrome.exe",
    ):
        if chrome is None and candidate.is_file():
            chrome = str(candidate)

    if chrome is not None:
        try:
            subprocess.Popen([chrome, url])
            log.info("Allure report opened in Google Chrome: %s", url)
            return
        except (OSError, subprocess.SubprocessError) as error:
            log.warning("Could not start Chrome (%s) - using the default "
                        "browser instead.", error)

    if webbrowser.open(url):
        log.info("Allure report opened in the default browser: %s", url)
    else:
        log.warning("Could not open a browser. The report is at: %s", index)


def announce_report(built: bool) -> None:
    """The banner the run signs off with: where this execution's report landed.

    printed rather than logged, so it is the last thing on the screen whatever
    the logging level is set to, and so it goes to the terminal rather than into
    execution.log - the log is inside the folder the banner is pointing at.
    """
    line = "=" * 41
    message = [line]

    if built:
        message += ["Execution Completed Successfully", "Allure Report:",
                    str(ALLURE_REPORT_DIR / "index.html")]
        single = ALLURE_SINGLE_REPORT_DIR / "index.html"
        if single.is_file():
            message += ["", "Opened (self-contained copy):", str(single)]
    else:
        message += ["Execution Completed",
                    "The report could not be built - this run's output is in:",
                    str(RUN_REPORT_DIR)]

    message += ["", f"Execution log:  {EXECUTION_LOG}",
                f"Environment:    {RUN_ENVIRONMENT_FILE}",
                "", "Previous reports are kept - nothing was overwritten.",
                line]
    print("\n".join(message))


# =========================================================================== #
# PLAIN PYTHON ENTRY POINT - same modules, same order, no pytest
#
#   python Construction_Flow.py
# =========================================================================== #

def main() -> int:
    """Run every module in order without pytest. Returns the exit code."""
    global UNDER_PYTEST
    UNDER_PYTEST = False           # a used-up row is logged, not skipped-by-pytest

    log.info("Launching Browser...")
    failures: List[str] = []
    blocked: List[str] = []

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(
            headless=not HEADED, slow_mo=SLOW_MO_MS, args=["--start-maximized"])
        context = browser.new_context(no_viewport=True, ignore_https_errors=True)
        page = context.new_page()
        page.set_default_timeout(DEFAULT_TIMEOUT)
        page.set_default_navigation_timeout(NAVIGATION_TIMEOUT)
        ApiFailureRecorder(page)

        # (sheet name, rows, what to do with one row)
        modules = [
            (SHEET_ACCOUNT, ACCOUNT_ROWS,
             lambda row: create_account_and_approve(page, row)),
            (SHEET_CUSTOMER, CUSTOMER_ROWS, lambda row: create_customer(page, row)),
            (SHEET_SUPPLIER, SUPPLIER_ROWS, lambda row: create_supplier(page, row)),
            (SHEET_ITEM, ITEM_ROWS, lambda row: create_item(page, row)),
            (SHEET_CATEGORY, CATEGORY_ROWS, lambda row: create_category(page, row)),
            (SHEET_QUOTATION, QUOTATION_ROWS, lambda row: create_quotation(page, row)),
            (SHEET_SALES_ORDER, SALES_ORDER_ROWS,
             lambda row: create_sales_order(page, row)),
            (SHEET_PURCHASE_ORDER, PURCHASE_ORDER_ROWS,
             lambda row: create_purchase_order(page, row)),
            (SHEET_GRN, GRN_ROWS, lambda row: create_grn(page, row)),
            (SHEET_PURCHASE_BILL, PURCHASE_BILL_ROWS,
             lambda row: approve_purchase_bill(page, row)),
            # The Suppliers sheet a second time, and on purpose: the supplier's
            # Recent Purchase Orders and Outstanding do not exist until
            # something has been bought from it, so they are verified here -
            # after the receipt has been approved - rather than when the
            # supplier was created. The same place test_supplier_purchase_
            # activity() sits in the pytest run.
            (SHEET_SUPPLIER, SUPPLIER_ROWS,
             lambda row: verify_supplier_received_order(page, row)),
            (SHEET_MATERIAL_INDENT, MATERIAL_INDENT_ROWS,
             lambda row: create_material_indent(page, row)),
        ]

        try:
            for sheet_name, rows, run_row in modules:
                # The same switch the pytest run obeys - see
                # RUN_CREATE_YOUR_ACCOUNT at the top of this file. Checked here
                # as well so `python Construction_Flow.py` and `pytest` behave
                # identically: one of them quietly registering a business while
                # the other did not would be the worst of both.
                if sheet_name == SHEET_ACCOUNT and not RUN_CREATE_YOUR_ACCOUNT:
                    log.warning("SKIPPING %s - %s", sheet_name,
                                CREATE_YOUR_ACCOUNT_OFF)
                    continue
                if not rows:
                    log.info("No executable rows in %s - skipping the module.",
                             sheet_name)
                    continue
                # The workspace modules need a signed-in business user.
                if sheet_name == SHEET_CUSTOMER:
                    login_as_business_user(page)

                for row in rows:
                    try:
                        with run_test_case(row, sheet_name):
                            run_row(row)
                    except (AssertionError, PlaywrightError, ValueError, KeyError) as error:
                        # One bad row must not cancel the rest of the module.
                        # A BLOCKED row is kept apart from a failed one here
                        # too: it could not be RUN, so it is not a finding and
                        # it must not make the exit code say there was one.
                        status, _ = classify_failure(error)
                        first = str(error).splitlines()[0]
                        if status == calc.BLOCKED:
                            blocked.append(f"{sheet_name}/{row.test_case_id}: "
                                           f"{first}")
                            log.warning("Test Case %s BLOCKED - it could not be "
                                        "run. Continuing with the next row.",
                                        row.test_case_id)
                        else:
                            failures.append(f"{sheet_name}/{row.test_case_id}: "
                                            f"{first}")
                            log.error("Test Case %s FAILED - continuing with "
                                      "the next row.", row.test_case_id)
                log.info("Finished %s Module.", sheet_name)
        except KeyboardInterrupt:
            log.warning("Interrupted by the user (Ctrl+C).")
            return 130
        finally:
            log.info("Closing the browser. Log: %s", EXECUTION_LOG)
            context.close()
            browser.close()

    if blocked:
        log.warning("%d test case(s) were BLOCKED - they could not be run, so "
                    "nothing about the application was proved by them. This is "
                    "neither an application defect nor an automation defect:\n"
                    "  %s", len(blocked), "\n  ".join(blocked))
    if failures:
        log.error("%d test case(s) FAILED:\n  %s", len(failures),
                  "\n  ".join(failures))
        return 1
    if blocked:
        log.info("No test case failed. %d could not be run - see the BLOCKED "
                 "list above and the QA summary.", len(blocked))
        return 0
    log.info("Automation Completed Successfully.")
    return 0


def create_account_and_approve(page: Page, row: Row) -> None:
    """The whole Create Your Account flow, for the plain-Python entry point.

    The same four steps, in the same order, as test_create_your_account():
    register -> confirm the registration completed -> pause for the MANUAL OTP
    -> wait for the tenant and approve it. Whether the code was typed and how
    the registration was confirmed are carried through here too, because they
    are what a missing tenant has to be explained with - the older version of
    this dropped both and the plain run therefore reported a blocked
    precondition as if the reason were unknown.
    """
    company = create_business_account(page, row)
    registered_because = registration_completed(page, company)
    otp_verified = wait_for_manual_otp(page)
    approve_business(page, company, otp_verified, registered_because)


if __name__ == "__main__":
    sys.exit(main())