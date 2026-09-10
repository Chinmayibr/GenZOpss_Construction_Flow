import re
from playwright.sync_api import Playwright, sync_playwright


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page = context.new_page()
    page.goto("https://app.genzopss.com/login")
    page.get_by_role("button", name="Purchases").click()
    page.get_by_role("link", name="Purchase Orders").click()
    page.get_by_role("button", name="New PO").click()
    page.locator("select[name=\"supplier_id\"]").select_option("50212784-3811-4762-8bbc-54365542bf0c")
    page.locator("input[name=\"expected_date\"]").fill("2026-08-31")
    page.get_by_role("textbox", name="e.g. 30 days").click()
    page.get_by_role("textbox", name="e.g. 30 days").fill("30")
    page.get_by_role("textbox", name="Internal notes for this").click()
    page.get_by_role("textbox", name="Internal notes for this").fill("N/A")
    page.get_by_role("button", name="Add Item").click()
    page.get_by_role("combobox").nth(1).select_option("4c68ccc1-036a-41e3-815a-1a322ebe231f")
    page.get_by_role("spinbutton").first.click()
    page.get_by_role("spinbutton").first.fill("21")
    page.get_by_role("button", name="Create Purchase Order").click()
    page.get_by_role("button", name="Create Purchase Order").click()
    page.get_by_role("button", name="Confirm PO").click()
    page.get_by_role("button", name="Send to Supplier").click()
    page.get_by_role("button", name="Yes, Send to Supplier").click()
    page.get_by_role("link", name="GRN").click()
    page.get_by_role("button", name="New GRN").click()
    page.get_by_role("button", name="— Select Supplier —").click()
    page.get_by_role("button", name="Skyline Electricals +").click()
    page.get_by_role("combobox").select_option("ebd340c1-11b7-4f42-ac88-36b7bf2cf506")
    page.get_by_role("textbox", name="Supplier's invoice no.").click()
    page.get_by_role("textbox", name="Supplier's invoice no.").fill("563434")
    page.get_by_role("textbox", name="e.g. MH12AB1234").click()
    page.get_by_role("textbox", name="e.g. MH12AB1234").fill("MH12AB1234")
    page.get_by_role("textbox", name="Supplier's DC number").click()
    page.get_by_role("textbox", name="Supplier's DC number").fill("3455")
    page.get_by_role("button", name="Save GRN").click()
    page.get_by_role("button", name="Approve GRN").click()
    page.locator("label").filter(has_text="Inter-state supplyIGST will").click()
    page.get_by_role("button", name="Approve & Create Bill").click()
    page.get_by_role("link", name="Purchase Bills").click()
    page.locator("div").filter(has_text=re.compile(r"^PB/26-27\/0005$")).click()
    page.get_by_role("button", name="Approve Bill").click()

    # ---------------------
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)
