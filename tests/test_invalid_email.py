import os

from dotenv import load_dotenv
from playwright.sync_api import Page, expect


load_dotenv()


def test_login_with_invalid_email_format_is_blocked(page: Page) -> None:
    base_url = os.getenv("BASE_URL")
    intentionally_invalid_email = "invalidemail"
    assert base_url, "BASE_URL must be set in .env"

    page.goto(base_url, wait_until="domcontentloaded")
    email_field = page.get_by_label("Email")
    email_field.fill(intentionally_invalid_email)
    page.get_by_label("Password").fill("password-for-email-format-test")

    page.get_by_role("button", name="Sign in").click()

    validation = email_field.evaluate(
        "element => ({valid: element.validity.valid, typeMismatch: element.validity.typeMismatch})"
    )
    assert not validation["valid"]
    assert validation["typeMismatch"]
    expect(page.get_by_text("Please enter a valid email.", exact=True)).to_be_visible()
    assert not page.url.endswith("#/dashboard")
    assert not page.get_by_label("Enter security pin code").is_visible()