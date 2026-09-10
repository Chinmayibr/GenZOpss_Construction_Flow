from playwright.sync_api import Playwright, sync_playwright


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page = context.new_page()
    page.goto("https://genzopss.com/")
    page.get_by_role("link", name="Get Started").click()
    page.get_by_role("textbox", name="27AABCU9603R1ZX").click()
    page.get_by_role("textbox", name="27AABCU9603R1ZX").fill("37AAZCV4662Z1ZG")
    page.get_by_role("textbox", name="Rajan Steel Industries Pvt.").click()
    page.get_by_role("textbox", name="Rajan Steel Industries Pvt.").fill("Vizag Coastal Constructions")
    page.get_by_role("button", name="Construction").click()
    page.get_by_role("button", name="Businesses (B2B) You sell to").click()
    page.get_by_role("button", name="🏗️ Construction").click()
    page.get_by_role("button", name="Civil & Structural").click()
    page.get_by_role("textbox", name="560058").click()
    page.get_by_role("textbox", name="560058").fill("562123")
    page.get_by_role("textbox", name="Industrial Area, Phase II").click()
    page.get_by_role("textbox", name="Industrial Area, Phase II").fill("NElamangala")
    page.get_by_role("textbox", name="Search for your business").click()
    page.get_by_role("textbox", name="Search for your business").fill("Nelama")
    page.get_by_text("NelamangalaKarnataka, India").click()
    page.get_by_role("button", name="Continue").click()
    page.get_by_role("textbox", name="Suresh Rajan").click()
    page.get_by_role("textbox", name="Suresh Rajan").fill("Madhu")
    page.get_by_role("textbox", name="9876543210").click()
    page.get_by_role("textbox", name="9876543210").fill("3455673323")
    page.get_by_role("textbox", name="suresh@rajan.co").click()
    page.get_by_role("textbox", name="suresh@rajan.co").fill("madhu@gmail.com")
    page.get_by_role("textbox", name="Min. 8 chars, 1 uppercase, 1").click()
    page.get_by_role("textbox", name="Min. 8 chars, 1 uppercase, 1").click()
    page.get_by_role("textbox", name="Min. 8 chars, 1 uppercase, 1").fill("SmartOps@123")
    page.get_by_role("textbox", name="Re-enter password").click()
    page.get_by_role("textbox", name="Re-enter password").fill("SmartOps@123")
    page.get_by_role("button").nth(1).click()
    page.get_by_role("button", name="Continue").click()
    page.locator(".mt-0\\.5.w-5").click()
    page.get_by_role("button", name="Create Account").click()
    page.goto("https://app.genzopss.com/register/verify?email=madhu%40gmail.com")

    # ---------------------
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)
