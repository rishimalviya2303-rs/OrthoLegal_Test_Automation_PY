import os
import re

from dotenv import load_dotenv
from playwright.sync_api import Page, expect


load_dotenv()


def authenticate_user(page: Page, email_env_var: str, password_env_var: str) -> None:
    base_url = os.getenv("BASE_URL")
    email = os.getenv(email_env_var)
    password = os.getenv(password_env_var)
    assert base_url, "BASE_URL must be set in .env"
    assert email, f"{email_env_var} must be set in .env"
    assert password, f"{password_env_var} must be set in .env"

    page.goto(base_url, wait_until="domcontentloaded")
    page.get_by_label("Email").fill(email)
    page.get_by_label("Password").fill(password)
    page.get_by_role("button", name="Sign in").click()

    pin_field = page.get_by_label("Enter security pin code")
    expect(pin_field).to_be_visible()

    pin_message = page.get_by_text("Login PIN sent to your email.", exact=False)
    expect(pin_message).to_be_visible()
    pin_match = re.search(r"\b\d{6}\b", pin_message.inner_text())
    assert pin_match, "A 6-digit verification PIN was not displayed on the verification-code screen"

    pin_field.fill(pin_match.group())
    page.get_by_role("button", name="Submit").click()

    page.wait_for_url(re.compile(r".*/#/dashboard$"))
    assert page.url.endswith("#/dashboard"), "Home page was not displayed after login"


def authenticate_super_admin(page: Page) -> None:
    authenticate_user(page, "TEST_EMAIL", "TEST_PASSWORD")


def authenticate_case_manager(page: Page) -> None:
    authenticate_user(page, "CASE_MANAGER_EMAIL", "CASE_MANAGER_PASSWORD")
