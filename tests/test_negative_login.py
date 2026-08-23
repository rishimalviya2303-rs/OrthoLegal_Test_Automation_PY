import os

from dotenv import load_dotenv
from playwright.sync_api import Page, expect

from conftest import test_case as report_case


load_dotenv()


@report_case(
    test_case_id="LOGIN_TC_002",
    module="Login Tests",
    title="Verify that if the user enters a valid email and an incorrect password, then an incorrect password error is displayed.",
    test_type="Negative",
    preconditions="The OrthoLegal portal is accessible and a valid test email is configured in .env.",
    steps="1. Open the portal.\n2. Enter the valid email.\n3. Enter an intentionally incorrect password.\n4. Submit the login form.",
    test_data="Valid email from .env; intentionally invalid password.",
    expected_result="The incorrect password error is displayed, and the user does not reach verification or the dashboard.",
    actual_result="The incorrect password error was displayed and authentication did not continue.",
)
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