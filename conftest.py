"""The first thing pytest loads - it puts this execution's folder in place.

The suite itself is one file, Construction_Flow.py, and it registers itself as a
plugin (`pytest_plugins = [__name__]`), so the fixtures and the reporting hooks
live there rather than here. This conftest exists for the one thing that has to
happen EARLIER than that: pytest imports a rootdir conftest.py before it
collects anything, so importing run_config from here means

    Reports/<YYYY-MM-DD_HH-MM-SS>/     exists, and
    Reports/<YYYY-MM-DD_HH-MM-SS>/execution.log      is open and being written

before collection starts. A suite that cannot even be imported - a missing
workbook, a syntax error, an uninstalled package - therefore still leaves a
folder behind with the reason written in it, instead of failing with nothing on
disk to look at.

Nothing else belongs in this file. Adding a hook here that Construction_Flow.py
also defines would run it twice.
"""

from __future__ import annotations

from run_config import EXECUTION_LOG, RUN_REPORT_DIR, RUN_TIMESTAMP, log


def pytest_configure(config) -> None:
    """Say, before the first test starts, where this execution will report to."""
    log.info("Execution %s", RUN_TIMESTAMP)
    log.info("Output folder: %s", RUN_REPORT_DIR)
    log.info("Execution log: %s", EXECUTION_LOG)
