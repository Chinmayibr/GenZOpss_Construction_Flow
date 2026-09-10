# GenZOpss - Construction Flow

A data-driven Playwright suite. All test data lives in one Excel workbook.

```
Construction_Flow.py                     the automation (config, steps, tests, runner)
test_data/Construction_Flow_Data.xlsx    THE test data - one sheet per module
excel_reader.py                          the ONE reader (path, rows, Execute filter)
data_verification.py                     entered data vs the View page, per Create
run_config.py                            this run's folder, its paths and its log
conftest.py                              loaded first, so the folder exists early
pytest.ini                               pytest + Allure settings
Reports/<YYYY-MM-DD_HH-MM-SS>/           ONE folder per execution, kept for ever
allure-results/                          raw Allure results (scratch, emptied per run)
allure-history/                          the trend's carrier (see Reporting)
```

## One data source

The framework reads **one** Excel file and nothing else:

```
C:\Users\Lenovo\PycharmProjects\GenZOpss_Construction_Flow\test_data\Construction_Flow_Data.xlsx
```

* The path is defined **once**, as `EXCEL_FILE` in `excel_reader.py` (built from
  the project folder, so it survives a move or a `git clone`). Change it there
  and the whole suite follows — nothing else stores a path.
* `excel_reader.py` is the only reader. It opens the workbook **once per run**,
  caches each sheet, and every module gets its rows through `get_test_cases()`.
  There is no second reader and no second workbook object.
* Each module reads its own worksheet: `Create_Your_Account` → the
  `Create_Your_Account` sheet, `Customer` → `Customer`, and so on.
* The framework is **read-only** with respect to the workbook. It never creates,
  rebuilds, overwrites or seeds it — the file you place in `test_data` is the
  only source of data.
* If the file is missing, or is open in Excel, the run stops with a plain-English
  message — not a stack trace:

  ```
  Construction_Flow_Data.xlsx not found. Please place the workbook inside the
  test_data folder.
  ```

## Execution order

Modules run one after another. **Every row of a module finishes before the next
module starts** — never one row from each sheet in turn.

```
Create_Your_Account  -> all Execute=YES rows   (register + admin approval)
Customer             -> all Execute=YES rows
Suppliers            -> all Execute=YES rows   (create the supplier)
Items                -> all Execute=YES rows
Categories           -> all Execute=YES rows
Quotations           -> all Execute=YES rows
Sales_Orders         -> all Execute=YES rows
Purchase_Orders      -> all Execute=YES rows   (Purchases > Purchase Orders)
GRN                  -> all Execute=YES rows   (received + approved into a bill)
                                               + VALIDATE the stock the receipt added
Purchase_Bills       -> all Execute=YES rows   (the bill that approval created)
Suppliers            -> all Execute=YES rows   (Recent Purchase Orders + Outstanding)
Material_Indents     -> all Execute=YES rows   (Construction > Material Indents)
                                               + VALIDATE the stock the issue and
                                                 the return moved
```

The **Suppliers sheet runs twice** and that is deliberate: the second pass reads
what the purchase flow did to the supplier's own page. A supplier that has just
been created has no orders and nothing outstanding, so checking it at that point
would only be checking that zero is zero.

The browser opens once, visible and maximized, and stays open for the whole run.
The business-user login happens once, just before the Customer module.

## The workbook

**Column order = screen order.** After the two fixed columns (`Execute`,
`TestCaseID`), every sheet lists its fields in the order they are entered on the
form — top to bottom, left to right.

| Sheet | Columns |
| --- | --- |
| `Create_Your_Account` | Execute, TestCaseID, GSTIN, Company Name, Business Type, Customer Type, Business Category, Business Sub Category, Pincode, Address, City, City Suggestion, Contact Person, Mobile Number, Email, Password, Confirm Password |
| `Customer` | Execute, TestCaseID, GSTIN, Customer Name, Phone, Email, Address, City, State, Pincode, Credit Limit, Credit Days |
| `Suppliers` | Execute, TestCaseID, GSTIN, Supplier Name, Phone, Email, City, State, Pincode, Credit Days, Account Name, Account No, Bank Name, IFSC Code, Branch |
| `Quotations` | Execute, TestCaseID, Customer, Subject, Valid Until, Payment Terms, Item Name, Quantity, Unit, Rate, Discount, Tax, Remarks |
| `Sales_Orders` | Execute, TestCaseID, Customer, Product, Quantity, Rate, Tax, Notes, Status Sequence |
| `Purchase_Orders` | Execute, TestCaseID, Supplier, Expected Date, Payment Terms, Notes, Item Name, Unit, Quantity, Rate, Status Sequence, Expected Status |
| `GRN` | Execute, TestCaseID, Supplier, Purchase Order, Invoice Number, Vehicle Number, DC Number, Inter State, Approve, Expected Status |
| `Purchase_Bills` | Execute, TestCaseID, Bill Number, Status Before Approval, Expected Status |
| `Material_Indents` | Execute, TestCaseID, Priority, Required By, Reference Type, Department, Item Code, Item Name, Quantity, Unit, Notes, Approved Quantity, Issue Stock, Return Quantity, Return Notes, Status After Create, Status After Approval, Expected Status |

Every column is a field the automation actually types or selects — nothing spare.
`Create_Your_Account` follows the wizard: business details, then the address
block, then the contact person.

* **Execute** — `YES` runs the row, `NO` skips it. There is a YES/NO picker on
  the cell.
* **TestCaseID** — appears in the console log, in the pytest test name and in the
  Allure report.
* **Valid Until** — an exact date (`2026-09-30`) or an offset (`TODAY+30`), so it
  never goes stale.
* **Quotations → Unit** — the unit shown on the line item (`PCS`, `BAGS`, `MT`).
  The quantity box has no label, so it is found through this unit text; change
  the item's unit in the app and only the cell changes. Blank means `PCS`.
* **Status Sequence** — the sales-order buttons clicked after the order is
  created, separated by `|`. Leave it blank to only create the order; leave
  `Customer` blank to skip the form and only drive the buttons.
* **Quotations / Sales_Orders → Customer** — must match a `Customer` sheet row.
  The automation looks that row up to get the city, because the customer picker
  lists options as `<initial> <name> <city>`.
* **Purchase_Orders / GRN / Purchase_Bills → the numbers** — `PO/26-27/0007`,
  `GRN/26-27/0003` and `PB/26-27/0005` are allotted by the application, never by
  the workbook. They are read off the screen and used as the name each record is
  verified and looked up under. So `GRN → Purchase Order` and
  `Purchase_Bills → Bill Number` are normally left **blank**: blank means "the
  one this run has just created", which is what makes the three sheets one
  journey.
* **GRN → Inter State** — `YES` ticks *Inter-state supply* on the approval
  dialog, which is what the application works IGST out from.
* **GRN → Approve** — anything but `NO` approves the receipt, and it is that
  approval which creates the purchase bill the next sheet works on.
* **Material_Indents → Quantity vs Return Quantity** — `Quantity` is what the
  site *asks* for and is never changed afterwards; `Approved Quantity` is what
  the approval dialog is given (blank = approve what was asked for); `Return
  Quantity` is only what goes back to the store. The finished record is read
  back and must read Requested 43 → Approved 43 → Issued 43 → Returned 2.
* **Material_Indents → Item Code** — the catalogue holds more than one item
  called *OPC Cement 53 Grade*, so the code is what picks the right one. The
  dropdown is matched on the text it displays, never on the item's id.
* **Material_Indents → Issue Stock / Return Quantity** — `NO` stops the row
  after the approval; a blank `Return Quantity` issues the stock but returns
  none of it. Either way the record is still verified at whatever point it
  stopped.
* **Expected Status / Status Before Approval** — the badge the record must be
  showing, matched from the start of the badge's own text (`Sent` is happy with
  `Sent to Supplier`). A blank cell reads the badge and reports it without
  asserting it.
* **Credit Limit, Credit Days, Discount, Tax, Remarks, Notes** — entered when the
  cell has a value and the form shows the field. A blank cell, or a field a
  tenant's configuration hides, is logged and attached to the Allure report
  instead of failing the row.

## Adding a test case

1. Open `test_data\Construction_Flow_Data.xlsx`.
2. Add a row to the sheet you want.
3. Set **Execute = YES**.
4. Save and close.
5. Run the automation.

No Python change, ever. Rows are found by scanning to the last populated row, so
there is no row limit and no hardcoded row number.

## Install

```bash
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe -m playwright install chromium
```

## Run

This is a **Python** Playwright project. Every command runs inside the `.venv`,
and there is no build step — no Maven, no `pom.xml`, no `mvn`.

```bash
pytest                           # tests, report, and Chrome - nothing else to type
python Construction_Flow.py      # plain run, no pytest and no report
```

Each row appears in Allure as its own test execution, showing the module, sheet,
TestCaseID, execution status, input data used, execution time, and — on failure —
the error details and a screenshot.

### First-time setup

```bash
python -m venv .venv
.venv\Scripts\activate
python -m pip install -r requirements.txt
python -m playwright install chromium
```

### Recording a flow (codegen)

Playwright's recorder is a **Python CLI command**. With the `.venv` active:

```bash
playwright codegen https://genzopss.com/
python -m playwright codegen https://genzopss.com/     # same thing, explicit
playwright codegen https://genzopss.com/ -o recording_new.py
```

The workspace modules live behind a login, so save the signed-in state once and
reuse it instead of signing in on every recording session:

```bash
playwright codegen https://app.genzopss.com/login --save-storage=auth.json
playwright codegen https://app.genzopss.com --load-storage=auth.json
```

`auth.json` is in `.gitignore` and is never committed.

> **Not this.** The Java form of the same command —
> `mvn exec:java -Dexec.mainClass=com.microsoft.playwright.CLI ...` — belongs to
> the Java framework in `VCNR_InfraTraq_Labour`, which has a `pom.xml`. Run it
> here and Maven stops with `MissingProjectException: there is no POM in this
> directory`, because this project has none and needs none.

What codegen emits is a **starting point, not a drop-in**: raw `page.click(...)`
calls with brittle selectors and the values hardcoded. This framework routes
every action through the self-healing retry layer (`ACTION_ATTEMPTS`,
`RETRY_TIMEOUT`) and takes every value from the workbook, so the usual move is to
lift the selectors out of the recording and hand-write the step — which is what
`recording.py` and `recording_purchase_order.py` are.

## Where a run's output goes

**Every execution keeps its own report.** Nothing is ever overwritten, so
`pytest` can be run as often as you like and each earlier report is still there:

```
Reports/
    2026-08-07_10-30-15/
        allure-report/          the report, with its history/ folder
        allure-report-single/   the same report as one openable file
        screenshots/            the pictures of anything that went red
        execution.log           everything the run logged
        environment.properties  what the run was pointed at
    2026-08-07_11-02-51/
    2026-08-08_09-14-45/
```

The folder is named for the moment the run started, in `YYYY-MM-DD_HH-MM-SS`.
That format sorts the same alphabetically as it does chronologically, which is
what lets the next run find the newest one by name. `run_config.py` stamps the
timestamp once per process and owns every path built from it; `conftest.py`
imports it before collection starts, so the folder exists and `execution.log` is
already being written by the time the suite is imported — a run that dies on a
missing workbook still leaves the reason on disk.

`allure-results/` and `allure-history/` stay at the project root, and neither is
a report: the first is scratch space for the run in progress (`--clean-alluredir`
empties it at the start of every run), the second is the trend's carrier.

The run signs off with the path:

```
=========================================
Execution Completed Successfully
Allure Report:
...\Reports\2026-08-07_10-30-15\allure-report\index.html

Opened (self-contained copy):
...\Reports\2026-08-07_10-30-15\allure-report-single\index.html

Execution log:  ...\Reports\2026-08-07_10-30-15\execution.log
Environment:    ...\Reports\2026-08-07_10-30-15\environment.properties

Previous reports are kept - nothing was overwritten.
=========================================
```

The report opens in **Google Chrome**, or in the default browser if Chrome is
not installed. Set `GENZ_NO_ALLURE_REPORT=1` to keep the results but skip the
report and the browser.

## Reporting - the Trend across executions

`pytest` builds and opens the report by itself. One point is added to the
**Trend** per execution, in this order:

| # | Step | Where |
| --- | --- | --- |
| 1 | write `environment.properties`, `categories.json`, `executor.json` and the previous history into `allure-results` | `pytest_collection_finish` |
| 2 | run the tests | — |
| 3 | write the same four again (nothing is ever deleted, only added) | `pytest_unconfigure` |
| 4 | `allure generate` into `Reports/<this run>/allure-report/` | `pytest_unconfigure` |
| 5 | archive the history to `allure-history/`, open the report, print the path | `pytest_unconfigure` |

Step 1 exists because `--clean-alluredir` empties `allure-results` at the start
of every run - history and `categories.json` included. It also runs *before* the
tests so that a report built by hand afterwards (`allure serve allure-results`)
still finds them; when `categories.json` is missing, Allure silently falls back
to its two built-in buckets, **Product defects** and **Test defects**, and none
of the project's own categories appear. Every copy is logged, e.g.:

```
Allure history copied before the run: ...\allure-report\history -> ...\allure-results\history
    (5 file(s): categories-trend.json, duration-trend.json, history-trend.json,
    history.json, retry-trend.json), latest build in it = #18
Allure trend: this execution is build #19 (the previous one was #18).
Allure history archived to ...\allure-history (through build #19)
```

Two things keep the trend honest over the long run:

* **The build number is `max(buildOrder) + 1`, never `len(trend) + 1`.** Allure
  keeps only the **last 20** runs in `history-trend.json`, so counting the
  entries freezes at 20 and every run from the 21st on would reuse build #21 —
  two executions drawn as one point.  A counter in `allure-history/build-order.txt`
  backs this up, so the number keeps rising even if the history is lost.
* **`allure-history/` holds a copy at the project root**, for two reasons. Each
  run now reports into its own `Reports/<timestamp>/` folder, so there is no
  fixed "last report" to read the history out of; and `allure generate --clean`
  deletes the report folder before it writes the new one, so a generate
  interrupted at that moment used to take the only copy of the trend with it and
  the next run started again at build 1. If the archive is ever lost, the trend
  is picked up from the newest folder under `Reports/` that has a `history/` in
  it, newest first — so a failed build does not break the line either.

To start the trend over, delete `allure-history/` (and `Reports/`, if you want
the old reports gone with it).

A test also keeps its own history (the History and Retries tabs, and the trend's
new/fixed/flaky figures) when the workbook is edited: Allure identifies a test by
`fullName` + every parameter, and the `row` parameter is the whole Excel row, so
one changed cell used to make a test look brand new. The suite pins the identity
to the test function + `TestCaseID` instead.

## Data verification - does the application really store what was typed?

Every **Create** in the suite is checked against the record's own **View page**,
automatically. No test case asks for it and no module contains a line of it: the
capture hangs off the framework's own `safe_fill` / `safe_select` / `safe_click`,
so a module written tomorrow (Purchase Order, GRN, Enquiry, any Construction
module) is covered the day it is written.

What each Create adds to the Allure report:

| Attachment | What it is |
| --- | --- |
| **Test Data Entered** | every value typed, selected or picked, as a table |
| **Create Page Screenshot** | the completed form, taken *before* Save is clicked |
| **View Page Data** | every label / value the record's View page displays |
| **View Page Screenshot** | the View page the values were read from |
| **Expected vs Actual** | entered vs displayed, PASS / FAIL per field |
| **PASS/FAIL Summary** | the counts and the verdict |
| **Data Mismatch Details** | only when something differs: field, expected, actual |

```
------------------------------------------------------------------------------
Field            Entered Value   View Value   Status    Notes
------------------------------------------------------------------------------
Item Code        ITEM-0001       ITEM-0001    PASS
Item Name        OPC Cement      OPC Cement   PASS
Category         Concrete Works  Concrete...  PASS
Purchase Price   340             ₹340.00      PASS      same value, different format
Selling Price    450             ₹450.00      PASS      same value, different format
MRP              500             ₹500.00      PASS      same value, different format
Tax                                           SKIPPED   read-only on this form
------------------------------------------------------------------------------
```

* **Formatting is never a defect.** `340` and `340.00`, `Yes` and `YES`,
  `9876543210` and `+91 98765 43210`, `2026-09-30` and `30 Sep 2026`, a rupee
  sign, thousands separators and extra spaces are all normalised before the
  comparison, so only a real difference reads as `FAIL`.
* **Blank and hidden fields are skipped, not failed.** An empty cell, a field a
  tenant's configuration hides, and a read-only value (the GST badge, derived
  from the item's HSN code) are reported as `SKIPPED` with the reason.
* **The View page is found by itself.** `/customers`, `/items`, `/sales-orders` -
  the address a module's records live at is derived from its name, the record is
  searched for and opened, and the page is read whatever shape it is (cards,
  definition lists, tables, read-only forms). A module whose address is not the
  obvious one gets one line in `VIEW_PAGES` in `data_verification.py`.
* **It never changes how a test behaves.** The check runs as the last thing a
  test case does, and puts the browser back on the address the test left it on.
  A mismatch is reported everywhere but does not fail the test - set
  `GENZ_FAIL_ON_DATA_MISMATCH=1` (or `FAIL_ON_MISMATCH` in
  `data_verification.py`) to make it fail as well. `GENZ_NO_DATA_VERIFICATION=1`
  keeps the entered-data tables and skips reading the View pages.
* **A failed test is left alone.** The View page is not opened, so the failure
  screenshot is a picture of the screen that actually failed.

## The Allure report - what a QA reader sees

The report is written for someone who will not open a line of Python. Every
execution carries the same shape:

```
Epic     GenZOpss Construction Flow
Feature  <module>                     e.g. Items, GRN, Material Indents
Story    <the business flow>          e.g. "Receive a purchase order and approve it"
Test     [TC012] Items                title, with Module / Sheet / TestCaseID / Result
```

and under it, **steps in business language, not Playwright**. The framework's own
`safe_fill` / `safe_select` / `safe_click` / `safe_expect` / `safe_navigation`
write them from the field descriptions they are already called with, so every
module gets them without a line of code in a test case:

```
Functional Validation
   ✓ Opened the Items module
   ✓ Clicked 'Add Item'
   ✓ Entered Item Code: ITEM-OPC53-011
   ✓ Entered Item Name: OPC Cement 53 Grade
   ✓ Selected Item Type: semi_finished
   ✓ Entered Purchase Price: 363
   ✓ Clicked 'Create Item'
   ✓ Verified item 'OPC Cement 53 Grade' in the Items list is displayed
Calculation Validation
   Calculation - Margin %: PASS
```

An action is journalled **after** it has succeeded, and a retry is one action
rather than three - so the list is what happened, not what was attempted.

| Attachment | Answers |
| --- | --- |
| **Test Data Used** | what data did this execution use? (the workbook row, as a table) |
| **Actions Performed** | what did the automation do, in order? |
| **Expected vs Actual** | what was entered, what the application displayed, PASS/FAIL per field |
| **Calculation Validation** | every sum this module checked: formula, expected, actual, difference, result |
| **Calculation Validation - `<name>`** | one sum in full, including the inputs it was worked out from |
| **Item Stock Validation** | the stock at every stage, expected vs actual, with the verdict |
| **Item Stock Validation - Calculation** | the stock arithmetic written out, and the five quantities kept apart |
| **Execution Status** | PASSED / FAILED / BLOCKED / SKIPPED, how long, and why |
| **Failure Details** | what went wrong, the last action that worked, and the application's own message |
| **QA Summary - Headline** | the run's counts and the three verdicts |
| **QA Summary - Module by Module** | Module / Functional / Calculation / Overall, colour-coded |

Screenshots are attached where they are evidence - the completed form, the saved
record, the View page, Items & Inventory at each stock reading, the final status,
and any failed calculation - not once per Playwright action.

## Stock and supplier validation - what the purchase flow actually moved

`calculation_validation.py` proves the application's arithmetic; this is the part
of it that follows ONE item and ONE supplier across the whole run, and it is the
only kind of check a single screen cannot answer. A stock figure on its own says
nothing - two readings of it, either side of a transaction, say everything.

Five readings are taken as the run goes, each at the only moment it exists:

| When | What is read | Where from |
| --- | --- | --- |
| before the purchase order | opening stock, supplier outstanding | Items & Inventory, Suppliers list |
| after the GRN approval | stock after the receipt | Items & Inventory |
| before the indent | stock before the issue | Items & Inventory |
| after Confirm Issue | stock after the issue | Items & Inventory |
| at the end of the run | stock remaining | Items & Inventory |

and they produce this, in the Allure report and in
`Reports/<run>/stock_ledger.txt`:

```
END-TO-END STOCK LEDGER - OPC Cement 53 Grade (ITEM-OPC53-011)

Opening Stock                                     98   read before the purchase order
+ GRN Received / Approved Quantity               160   read from the approved goods receipt
= Stock After GRN                                258   read after the GRN approval
  Stock Before the Indent                        258   read before the indent
- ACTUAL Issued Quantity                          68   read from the finished indent
= Stock After Issue                              190   read after the stock issue
+ ACTUAL Returned Quantity                         6   read from the finished indent
------------------------------------------------------------------------------
= Expected Closing Stock                         196
  Actual Stock Remaining                         196   Items & Inventory, STOCK column
  Difference                                       0
  Result                                        PASS
```

* **REQUESTED, APPROVED, ISSUED, RECEIVED and RETURNED are five different
  numbers** and the ledger keeps them apart. The stock reduction uses the
  quantity the application really ISSUED, read off the finished indent - never
  the quantity the workbook asked for. This application issues what it can: a
  row asking for 55 against a stock of 3 is issued 3, and a ledger built from
  the request would report a correct application as 52 short.
* **Every movement is confirmed twice.** The item's own STOCK LEDGER names the
  document behind each one ("GRN GRN/0045", "Issue against IND/0031"), so this
  run's receipt, issue and return are picked out of the item's whole history and
  held to the quantities the GRN and the indent say they moved.
* **The item is found by its CODE**, compared against the list's CODE cell. The
  catalogue holds eleven items called "OPC Cement 53 Grade" with eleven
  different stocks, so a stock read by name is a real figure answering a
  question nobody asked.
* **The supplier is checked after the receipt, not when it is created.** Its
  RECENT PURCHASE ORDERS block has to list the order this run raised, at the
  value the order itself carries and marked Received; its **Outstanding** is
  checked as the movement it is - what it was before this run's purchase, plus
  what this run bought.
* **Outstanding is an AMOUNT in this application, not a quantity** - established
  by reading it, not assumed. The outstanding QUANTITY (Ordered - Received) is
  still worked out, and reported as SKIPPED with the reason, because answering a
  quantity question with a money figure would be reporting two different things
  as one.
* **A reading that could not be taken is never invented.** It becomes a SKIPPED
  check that says which figure was missing and why, and it never fails a module.
  Every excursion puts the browser back where it found it, so nothing here
  changes how a test behaves.

## Configuration

Only environment settings live in the CONFIGURATION block at the top of
`Construction_Flow.py`: URLs, admin credentials, the business-user credentials,
timeouts, and the browser options. All test data is in the workbook.

Every value is typed **exactly as written in the cell** — the automation never
generates or modifies test data. The application's uniqueness rules therefore
apply as they would to a manual tester: a repeated mobile number
(`PHONE_IN_USE`) or GSTIN (`GSTIN_IN_USE`) is rejected with a 409 that the UI
never displays. To re-run a row, change those cells in the workbook.

## Waiting

No `time.sleep()`. Playwright waits for elements to be visible, enabled and
stable by itself; the script adds only what it cannot know about — page load,
network idle, and the app's own spinners / `Loading...` placeholders
(`wait_until_ready`). The single fixed pause (`settle`, 500 ms) is for CSS
animations that emit no DOM signal.

## Turning the Create Your Account module off (and on again)

One switch, at the top of `Construction_Flow.py`:

```python
RUN_CREATE_YOUR_ACCOUNT = False   # <- True to run it again
```

It is currently **False**, so registration and admin approval do not run. Every
other module runs exactly as before. Nothing else about the module is touched —
the flow, its helpers and its `Create_Your_Account` sheet are unchanged — and
the summary still lists it, as `SKIPPED` with the reason `MODULE SWITCHED OFF`,
so the coverage the report claims stays honest. To run it once without editing
the file: `$env:GENZ_RUN_CREATE_ACCOUNT = "1"`.

## OTP — a manual prerequisite, never automated

The Create_Your_Account module pauses at the application's verification page for
`OTP_WAIT_TIMEOUT` (120 s by default, `GENZ_OTP_WAIT_SECONDS=<seconds>`) so the
code can be typed **by hand**. The code arrives by e-mail: nothing in this
framework generates, guesses, hard-codes or bypasses one. The pause ends the
instant the application accepts the code, and reports every 15 s while it waits
so a watched run cannot look hung. `GENZ_OTP_WAIT_SECONDS=0` skips the pause for
an unattended run.

It is not a hard gate — if no code is typed the run carries on. What decides the
module is what the admin portal then shows.

## Tenant availability, and BLOCKED vs FAILED

Registering a business does not put it in the admin portal's Tenants list at
once, and the list draws a bare `Loading...` cell while it fetches. So the
approval step **polls**: it searches, waits for the loaders, looks for the exact
company, then reloads and searches again until `TENANT_WAIT_TIMEOUT` (180 s,
`GENZ_TENANT_WAIT_SECONDS=<seconds>`) is spent — see `find_tenant_row()`.

If the tenant still is not there, the module is reported as

    BLOCKED PRECONDITION - business registration completed, but the business did
    not become available as a tenant in the Admin Portal within the configured
    timeout.

with the precondition, the root cause, what was expected, what actually
happened, and what to do before the next run. A BLOCKED module is **not** a
pytest failure: it is recorded as a skip carrying that reason, and counted in
the report under its own heading. It is neither an application defect nor an
automation defect, and it is never counted as a pass. `GENZ_FAIL_ON_BLOCKED=1`
makes a blocked precondition fail the run, for a CI job that needs that.

The five verdicts the report distinguishes:

| Verdict | Meaning |
|---|---|
| PASS | the check ran and matched the expected result |
| FAIL | the check ran and the result was wrong — a real finding |
| BLOCKED | it could not be RUN; nothing was proved either way |
| SKIPPED | one calculation could not be validated; the reason is printed |
| NOT AUTOMATED | deliberately outside this suite's scope |
