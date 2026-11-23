from playwright.sync_api import Playwright, sync_playwright


def run(playwright: Playwright) -> None:
    browser = playwright.chromium.launch(headless=False)
    context = browser.new_context()
    page = context.new_page()
    page.goto("https://www.iolformula.com/agreement/")
    page.get_by_text("I Agree", exact=True).click()
    page.get_by_role("textbox", name="Surgeon").click()
    page.get_by_role("textbox", name="Surgeon").fill("attack01")
    page.get_by_role("textbox", name="Patient").click()
    page.get_by_role("textbox", name="Patient").fill("KALADA")
    page.get_by_role("textbox", name="ID").click()
    page.get_by_role("textbox", name="ID").fill("105737")
    page.get_by_text("Sex M F Biological sex is").click()
    page.locator("#A-Constant1").click()
    page.locator("#A-Constant1").fill("119.0")
    page.locator("#right-target").click()
    page.locator("#right-target").fill("-0.03")
    page.locator("#right-target").click()
    page.locator("#right-target").fill("-0.03")
    page.locator("#right-target").press("Enter")
    page.locator("#al-right").click()
    page.locator("#al-right").fill("23.33")
    page.locator("input[name=\"k1_right\"]").click()
    page.locator("input[name=\"k1_right\"]").fill("44.25")
    page.locator("input[name=\"k2_right\"]").click()
    page.locator("input[name=\"k2_right\"]").fill("44.75")
    page.locator("#acd-right").click()
    page.locator("#acd-right").fill("2.18")
    page.get_by_role("button", name="Calculate").click()
    page.get_by_role("cell", name="-0.57").click()

    # ---------------------
    context.close()
    browser.close()


with sync_playwright() as playwright:
    run(playwright)
