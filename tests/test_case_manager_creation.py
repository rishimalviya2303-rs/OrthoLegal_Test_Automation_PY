import os
from typing import Any

from dotenv import load_dotenv
from playwright.sync_api import Locator, Page, expect

from conftest import test_case as report_case
from utils.auth import authenticate_case_manager, authenticate_super_admin
from utils.excel_reader import PROJECT_ROOT, read_worksheet


load_dotenv()

USER_TEST_DATA_PATH = PROJECT_ROOT / "test_data" / "Users_Test_Data.xlsx"
SIGNATURE_PATH = PROJECT_ROOT / "test_data" / "Signatures" / "CaseManagetSignature.png"
REQUIRED_COLUMNS = {
    "Select Role",
    "First Name",
    "Middle Name",
    "Last Name",
    "Ortho Legal Email",
    "Personal Email",
    "Password",
    "Confirm Password",
}


def _case_manager_data() -> dict[str, Any]:
    rows = read_worksheet("Users", workbook_path=USER_TEST_DATA_PATH, required_columns=REQUIRED_COLUMNS)
    matches = [
        row for row in rows
        if str(row.get("Select Role", "")).strip().lower() == "case manager"
    ]
    assert matches, (
        "No Case Manager creation row found in test_data/Users_Test_Data.xlsx. "
        "Add a Users row with Select Role=Case Manager, First Name, Middle Name, Last Name, "
        "Ortho Legal Email, Personal Email, Password, and Confirm Password."
    )
    return matches[0]


def _required_text(row: dict[str, Any], column: str) -> str:
    value = row.get(column)
    assert value is not None and str(value).strip(), f"Users test data column '{column}' must contain a value."
    return str(value).strip()


def _case_manager_name(row: dict[str, Any]) -> str:
    return " ".join(
        part for part in [
            _required_text(row, "First Name"),
            _required_text(row, "Middle Name"),
            _required_text(row, "Last Name"),
        ]
        if part
    )


def _visible(locator: Locator) -> Locator:
    return locator.first


def _fill_field(form: Locator, label: str, value: str) -> None:
    field = form.get_by_label(label)
    expect(_visible(field)).to_be_visible()
    _visible(field).fill(value)


def _case_manager_email_from_env() -> str:
    email = os.getenv("CASE_MANAGER_EMAIL")
    assert email and email.strip(), "CASE_MANAGER_EMAIL must be set in .env"
    return email.strip()


def _assert_case_manager_password_configured() -> None:
    password = os.getenv("CASE_MANAGER_PASSWORD")
    assert password and password.strip(), "CASE_MANAGER_PASSWORD must be set in .env"


def _logout(page: Page) -> None:
    account_menu = page.get_by_role("button", name="Account").or_(
        page.get_by_role("button", name="Profile")
    ).or_(
        page.get_by_role("button", name="User")
    )
    if account_menu.count() > 0:
        _visible(account_menu).click()

    logout_button = page.get_by_role("button", name="Logout").or_(
        page.get_by_role("menuitem", name="Logout")
    ).or_(
        page.get_by_text("Logout", exact=True)
    )
    expect(_visible(logout_button)).to_be_visible()
    _visible(logout_button).click()
    expect(page.get_by_label("Email")).to_be_visible()


@report_case(
    test_case_id="SUPERADMIN_USER_TC_001",
    module="Super Admin - Roles and Users",
    title="Verify that if the Super Admin enters valid Case Manager details and saves the user, then the Case Manager is created successfully, displayed in the Users list with the correct name and email, and can log in successfully.",
    test_type="Positive",
    preconditions="The OrthoLegal web portal is accessible. A valid Super Admin test account exists. Case Manager creation data is available in the Users worksheet of the existing Excel input workbook. CASE_MANAGER_EMAIL and CASE_MANAGER_PASSWORD are available through .env. The Case Manager signature file exists in the repository.",
    steps="1. Log in as Super Admin using environment credentials and the displayed verification PIN.\n2. Navigate to Database.\n3. Open Roles and Users.\n4. Verify the Users screen is displayed.\n5. Open the Add User form.\n6. Select the role whose visible name matches the Select Role value from Excel.\n7. Enter Case Manager details from Excel test data.\n8. Verify the Case Manager Ortho Legal email matches CASE_MANAGER_EMAIL from .env.\n9. Upload the Case Manager signature PNG from the repository test_data folder.\n10. Save the user.\n11. Search by the created Case Manager name.\n12. Verify the Users search results show the expected Case Manager name and Ortho Legal email.\n13. Log out the Super Admin.\n14. Log in as the created Case Manager using .env credentials and the displayed verification PIN.\n15. Verify that the authenticated Case Manager home page is displayed.",
    test_data="User creation data from the Users worksheet in test_data/Users_Test_Data.xlsx; Case Manager login email and password from .env; signature from test_data/Signatures/CaseManagetSignature.png. Password and verification PIN are not reported.",
    expected_result="The Case Manager should be created successfully, should be displayed in the Users search results with the same name and Ortho Legal email entered during creation, and should be able to log in successfully to the authenticated Case Manager home page.",
    actual_result="The Case Manager was created successfully, was displayed in the Users search results with the expected name and Ortho Legal email, and logged in successfully to the authenticated Case Manager home page.",
)
def test_creation_of_case_manager(page: Page) -> None:
    case_manager = _case_manager_data()
    selected_role = _required_text(case_manager, "Select Role")
    case_manager_name = _case_manager_name(case_manager)
    ortho_legal_email = _required_text(case_manager, "Ortho Legal Email")
    personal_email = _required_text(case_manager, "Personal Email")
    assert ortho_legal_email == _case_manager_email_from_env(), (
        "Ortho Legal Email in test_data/Users_Test_Data.xlsx must match CASE_MANAGER_EMAIL in .env"
    )
    _assert_case_manager_password_configured()

    assert SIGNATURE_PATH.exists(), f"Signature file was not found at expected path: {SIGNATURE_PATH}"

    authenticate_super_admin(page)

    database_link = page.get_by_role("link", name="Database").or_(page.get_by_text("Database", exact=True))
    expect(_visible(database_link)).to_be_visible()
    _visible(database_link).click()

    roles_and_users_link = page.get_by_role("link", name="Roles and Users").or_(
        page.get_by_text("Roles and Users", exact=True)
    )
    expect(_visible(roles_and_users_link)).to_be_visible()
    _visible(roles_and_users_link).click()

    users_tab = page.get_by_role("tab", name="Users").or_(page.get_by_text("Users", exact=True))
    expect(_visible(users_tab)).to_be_visible()

    add_user_button = page.get_by_role("button", name="Add User")
    expect(add_user_button).to_be_visible()
    add_user_button.click()

    add_user_form = page.get_by_role("dialog").filter(has_text="Add User")
    if add_user_form.count() == 0:
        add_user_form = page.locator("form").filter(has_text="Case Manager")
    expect(_visible(add_user_form)).to_be_visible()
    form = _visible(add_user_form)

    role_checkbox = form.get_by_label(selected_role)
    expect(_visible(role_checkbox)).to_be_visible()
    _visible(role_checkbox).check()
    expect(_visible(role_checkbox)).to_be_checked()

    _fill_field(form, "First Name", _required_text(case_manager, "First Name"))
    _fill_field(form, "Middle Name", _required_text(case_manager, "Middle Name"))
    _fill_field(form, "Last Name", _required_text(case_manager, "Last Name"))
    _fill_field(form, "Ortho Legal Email", ortho_legal_email)
    _fill_field(form, "Personal Email", personal_email)
    _fill_field(form, "Password", _required_text(case_manager, "Password"))
    _fill_field(form, "Confirm Password", _required_text(case_manager, "Confirm Password"))

    signature_input = form.locator("input[type='file']")
    expect(_visible(signature_input)).to_be_attached()
    _visible(signature_input).set_input_files(str(SIGNATURE_PATH))

    uploaded_signature_indicator = form.get_by_text(SIGNATURE_PATH.name, exact=False).or_(
        form.locator("img")
    )
    expect(_visible(uploaded_signature_indicator)).to_be_visible()

    save_button = form.get_by_role("button", name="Save")
    expect(save_button).to_be_enabled()
    save_button.click()

    expect(form).not_to_be_visible()

    search_field = page.get_by_placeholder("Search by Name").or_(page.get_by_label("Search by Name"))
    expect(_visible(search_field)).to_be_visible()
    _visible(search_field).fill(case_manager_name)

    matching_row = page.get_by_role("row").filter(has_text=case_manager_name).filter(has_text=ortho_legal_email)
    expect(_visible(matching_row)).to_be_visible()
    expect(_visible(matching_row).get_by_text(case_manager_name, exact=True)).to_be_visible()
    expect(_visible(matching_row).get_by_text(ortho_legal_email, exact=True)).to_be_visible()

    _logout(page)
    authenticate_case_manager(page)
