"""Where THIS execution writes, and the logging that goes with it.

Every run gets a folder of its own, stamped with the moment it started:

    Reports/<YYYY-MM-DD_HH-MM-SS>/
        allure-report/            the HTML report, with its history/ folder
        allure-report-single/     the same report as one openable file
        screenshots/              the pictures of anything that went red
        execution.log             everything the run logged
        environment.properties    what the run was pointed at

Nothing is ever overwritten: a run started at 10:30:15 cannot tread on one
started at 11:02:51, so `pytest` can be run as often as you like and every
report is still there afterwards.

allure-results/ and allure-history/ stay at the project root, and neither is a
report:

    allure-results/   scratch space for the run in progress. pytest.ini writes
                      the raw results there and --clean-alluredir empties it at
                      the start of every run, so there is nothing in it to keep.
    allure-history/   the trend's carrier. Allure draws the TREND, HISTORY and
                      RETRIES panels from a history/ folder that has to reach
                      the NEXT run, and each run's report is now in a different
                      timestamped folder - so the history is archived here, in
                      one fixed place, the moment a report is built.

This module is imported by conftest.py, which is the very first thing pytest
loads. That is deliberate: the folders exist and execution.log is open before
collection starts, so even an error while importing the suite is written down.
"""

from __future__ import annotations

import logging
import sys
from datetime import datetime
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent

# --------------------------------------------------------------------------- #
# THIS RUN'S STAMP
#
# Read once, when this module is first imported, which happens exactly once per
# process. Every folder the execution writes to hangs off it.
# --------------------------------------------------------------------------- #
RUN_TIMESTAMP = f"{datetime.now():%Y-%m-%d_%H-%M-%S}"

# --------------------------------------------------------------------------- #
# Output folders
# --------------------------------------------------------------------------- #
#: The parent of every execution's folder. Created if it is not there.
REPORTS_DIR = PROJECT_ROOT / "Reports"
#: THIS execution's folder. Nothing else writes here, ever.
RUN_REPORT_DIR = REPORTS_DIR / RUN_TIMESTAMP

#: Kept under the old name so nothing that used to write to artifacts/ has to
#: change - it now writes into this run's folder instead.
ARTIFACTS_DIR = RUN_REPORT_DIR
SCREENSHOT_DIR = RUN_REPORT_DIR / "screenshots"
EXECUTION_LOG = RUN_REPORT_DIR / "execution.log"
RUN_ENVIRONMENT_FILE = RUN_REPORT_DIR / "environment.properties"

# --------------------------------------------------------------------------- #
# Allure
# --------------------------------------------------------------------------- #
#: Where pytest writes the raw results (pytest.ini already passes --alluredir).
ALLURE_RESULTS_DIR = PROJECT_ROOT / "allure-results"
#: This run's HTML report. It also holds history/, which is what the next run
#: turns into the trend.
ALLURE_REPORT_DIR = RUN_REPORT_DIR / "allure-report"
#: The same report as one self-contained file, which is the copy Chrome opens.
ALLURE_SINGLE_REPORT_DIR = RUN_REPORT_DIR / "allure-report-single"
#: The trend's safety copy, at the project root rather than inside a run's
#: folder: it has to be findable by the NEXT run without knowing which folder
#: the last one used, and `allure generate --clean` deletes the report folder
#: before it writes the new one. Delete this to start the trend again at 1.
ALLURE_HISTORY_ARCHIVE = PROJECT_ROOT / "allure-history"
#: The build number the last execution used. Allure keeps only the last 20 runs
#: in history-trend.json, so the trend file alone cannot be counted on to say
#: which build this is - this file can.
ALLURE_BUILD_ORDER_FILE = ALLURE_HISTORY_ARCHIVE / "build-order.txt"
#: The name the report and its builds carry in the Executors panel. It is what
#: groups the points of the TREND chart into one line - never change it
#: mid-project.
ALLURE_REPORT_NAME = "GenZOpss Construction Flow"
#: Build the report and open it in Chrome the moment the run ends. Set the
#: environment variable GENZ_NO_ALLURE_REPORT=1 to turn it off (CI, debugging).
AUTO_OPEN_ALLURE_REPORT = True

#: The format the run folders are named in. It is written down here because the
#: whole scheme depends on it sorting the same alphabetically as it does
#: chronologically - which is what lets "the newest folder" be found by name.
TIMESTAMP_FORMAT = "%Y-%m-%d_%H-%M-%S"


def ensure_run_dirs() -> None:
    """Create this execution's output folders. Safe to call as often as you like.

    Called from the bottom of this module, so simply importing it is enough to
    have somewhere to write - which matters because the logging FileHandler
    below opens execution.log immediately, and a FileHandler will not create the
    folder it is pointed at.
    """
    for folder in (REPORTS_DIR, RUN_REPORT_DIR, SCREENSHOT_DIR,
                   ALLURE_RESULTS_DIR):
        folder.mkdir(parents=True, exist_ok=True)


def previous_run_dirs() -> list[Path]:
    """Every earlier execution's folder, newest first. This run's is left out.

    The timestamp format sorts the same alphabetically as it does
    chronologically, so ordering the names backwards is ordering the runs from
    most recent to oldest.
    """
    if not REPORTS_DIR.is_dir():
        return []
    try:
        return sorted((path for path in REPORTS_DIR.iterdir()
                       if path.is_dir() and path.name != RUN_TIMESTAMP),
                      key=lambda path: path.name, reverse=True)
    except OSError:
        return []


def pytest_is_printing() -> bool:
    """True while pytest's own live-logging handler is on the root logger.

    pytest.ini switches log_cli on, and that puts a handler of pytest's on the
    root logger for the phases it covers - collection, and each test's setup,
    call and teardown. It is added when the phase starts and taken off again
    when the phase ends (catching_logs() in _pytest/logging.py), so asking
    whether it is there RIGHT NOW is the same question as "is pytest already
    printing this record".

    Matched by class rather than imported, because it is private to pytest
    (_pytest.logging._LiveLoggingStreamHandler). Its stand-in when log_cli is
    switched off is a NullHandler that prints nothing, and that one must not
    count. If a future pytest renames the class, the worst that happens is the
    duplicate lines come back - no output is ever lost.
    """
    for handler in logging.getLogger().handlers:
        if isinstance(handler, logging.NullHandler):
            continue
        kind = type(handler)
        if (kind.__module__.startswith("_pytest.logging")
                and "LiveLogging" in kind.__name__):
            return True
    return False


class QuietWhilePytestPrints(logging.Filter):
    """Keeps the console handler quiet while pytest is printing the same record.

    Without it every line appears on the terminal twice during collection and
    during the tests - once from pytest's live-logging handler and once from
    ours - while the lines outside those phases (pytest_configure, the report
    being built in pytest_unconfigure) appear only once, because pytest's
    handler is not on the root logger then.

    Filtering rather than removing our handler is what keeps that asymmetry from
    mattering: the console handler stays in place for the whole run and simply
    stands aside for the stretches pytest has covered. A plain
    `python Construction_Flow.py` run has no pytest handler at any point, so
    nothing is ever filtered out of it.

    Only the console handler carries this filter. execution.log is written by a
    handler of its own, which pytest never duplicates, so the log file holds
    every line however the suite was started.
    """

    def filter(self, record: logging.LogRecord) -> bool:
        return not pytest_is_printing()


def configure_logging() -> logging.Logger:
    """Console + Reports/<run>/execution.log, set up once per process.

    basicConfig does nothing when the root logger already has handlers, so a
    second call is harmless - but the guard is explicit here because this module
    is imported both by conftest.py and by the suite itself.
    """
    # Windows consoles are cp1252 and cannot print the app's emoji / rupee signs.
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except (AttributeError, OSError):
        pass

    root = logging.getLogger()
    if not any(getattr(handler, "_genz_run_log", False)
               for handler in root.handlers):
        file_handler = logging.FileHandler(EXECUTION_LOG, encoding="utf-8")
        file_handler._genz_run_log = True                # noqa: SLF001 - our mark

        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.addFilter(QuietWhilePytestPrints())

        logging.basicConfig(
            level=logging.INFO,
            format="%(asctime)s | %(levelname)-7s | %(message)s",
            datefmt="%H:%M:%S",
            handlers=[console_handler, file_handler],
        )
    return logging.getLogger("construction_flow")


ensure_run_dirs()
log = configure_logging()
