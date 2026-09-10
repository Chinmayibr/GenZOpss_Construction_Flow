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
    page.locator("input[name=\"expected_date\"]").fill("2026-08-24")
    page.get_by_role("button", name="Add Item").click()
    page.get_by_role("button", name="Add Item").click()
    page.get_by_role("combobox").nth(1).select_option("4c68ccc1-036a-41e3-815a-1a322ebe231f")
    page.get_by_role("button", name="Remove row 2").click()
    page.get_by_role("button", name="Create Purchase Order").click()
    page.close()

    # ---------------------
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)
