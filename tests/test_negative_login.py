import os

from dotenv import load_dotenv
from playwright.sync_api import Page, expect


load_dotenv()


def test_login_with_incorrect_password_fails(page: Page) -> None:
    base_url = os.getenv("BASE_URL")
    test_email = os.getenv("TEST_EMAIL")
    intentionally_invalid_password = "intentionally-invalid-password-for-negative-test"
    assert base_url, "BASE_URL must be set in .env"
    assert test_email, "TEST_EMAIL must be set in .env"

    page.goto(base_url, wait_until="domcontentloaded")
    page.get_by_label("Email").fill(test_email)
    page.get_by_label("Password").fill(intentionally_invalid_password)
    page.get_by_role("button", name="Sign in").click()

    error_message = page.get_by_text("Incorrect email or password", exact=True)
    expect(error_message).to_be_visible()
    assert not page.url.endswith("#/dashboard")
    assert not page.get_by_label("Enter security pin code").is_visible()