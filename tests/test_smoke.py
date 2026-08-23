import os
import re

from dotenv import load_dotenv
from playwright.sync_api import Page


load_dotenv()


def test_ortholegal_login(page: Page) -> None:
    base_url = os.getenv("BASE_URL")
    test_email = os.getenv("TEST_EMAIL")
    test_password = os.getenv("TEST_PASSWORD")
    assert base_url, "BASE_URL must be set in .env"
    assert test_email, "TEST_EMAIL must be set in .env"
    assert test_password, "TEST_PASSWORD must be set in .env"

    page.goto(base_url, wait_until="domcontentloaded")
    page.get_by_label("Email").fill(test_email)
    page.get_by_label("Password").fill(test_password)
    page.get_by_role("button", name="Sign in").click()

    pin_field = page.get_by_label("Enter security pin code")
    pin_field.wait_for()
    pin_message = page.get_by_text("Login PIN sent to your email.", exact=False)
    pin_match = re.search(r"\b\d{6}\b", pin_message.inner_text())
    assert pin_match, "A 6-digit verification PIN was not displayed"

    pin_field.fill(pin_match.group())
    page.get_by_role("button", name="Submit").click()

    page.wait_for_url(re.compile(r".*/#/dashboard$"))
    assert page.url.endswith("#/dashboard"), "Login did not reach the dashboard"