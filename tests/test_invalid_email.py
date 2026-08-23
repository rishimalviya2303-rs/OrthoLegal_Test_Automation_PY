import os

from dotenv import load_dotenv
from playwright.sync_api import Page, expect

from conftest import test_case as report_case


load_dotenv()


@report_case(
    test_case_id="LOGIN_TC_003",
    module="Login Tests",
    title="Verify that if the user enters an invalid email format, then the user is not allowed to proceed with login.",
    test_type="Negative",
    preconditions="The OrthoLegal portal is accessible.",
    steps="1. Open the portal.\n2. Enter an invalid email format.\n3. Enter a password.\n4. Attempt to submit the login form.",
    test_data="Invalid email format: invalidemail; password used only to test email validation.",
    expected_result="The email validation error is displayed, and the user does not reach verification or the dashboard.",
    actual_result="The email validation error was displayed and authentication did not continue.",
)
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