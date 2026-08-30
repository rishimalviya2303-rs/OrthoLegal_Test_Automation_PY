from playwright.sync_api import Page

from conftest import test_case as report_case
from utils.auth import authenticate_super_admin


@report_case(
    test_case_id="SUPERADMIN_LOGIN_TC_001",
    module="Super Admin - Login",
    title="Verify that if the Super Admin enters valid login credentials and the displayed verification PIN, then the Super Admin is successfully logged in.",
    test_type="Positive",
    preconditions="The OrthoLegal web portal is accessible. A valid Super Admin test account exists. BASE_URL, Super Admin email, and Super Admin password are available through the existing .env configuration.",
    steps="1. Read the application URL from .env.\n2. Navigate to the OrthoLegal login page.\n3. Enter the valid Super Admin email from .env.\n4. Enter the valid Super Admin password from .env.\n5. Submit the login form.\n6. Wait for the verification-code screen.\n7. Read the displayed six-digit verification PIN from the page.\n8. Enter the displayed PIN into the verification-code field.\n9. Submit the verification code.\n10. Verify that the application Home page is displayed.",
    test_data="Valid Super Admin credentials from environment configuration",
    expected_result="The Super Admin should be authenticated successfully after entering valid credentials and the displayed verification PIN, and the Home page should be displayed.",
    actual_result="The Super Admin credentials were accepted, verification was completed successfully, and the Home page was displayed.",
)
def test_super_admin_login(page: Page) -> None:
    authenticate_super_admin(page)
