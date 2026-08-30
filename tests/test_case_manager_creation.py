import re
from typing import Any

from playwright.sync_api import Locator, Page, expect

from conftest import test_case as report_case
from utils.auth import authenticate_super_admin, authenticate_with_credentials
from utils.excel_reader import PROJECT_ROOT, read_worksheet


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
    middle_name = row.get("Middle Name")
    return " ".join(
        part for part in [
            _required_text(row, "First Name"),
            str(middle_name).strip() if middle_name is not None else "",
            _required_text(row, "Last Name"),
        ]
        if part
    )


def _visible(locator: Locator) -> Locator:
    return locator.first


def _fill_field(form: Locator, label: str | re.Pattern[str], value: str) -> None:
    field = form.get_by_role("textbox", name=label).or_(form.get_by_label(label))
    expect(_visible(field)).to_be_visible()
    _visible(field).fill(value)


def _select_role_from_excel(page: Page, form: Locator, selected_role: str) -> None:
    role_selector = form.get_by_label("Roles")
    expect(_visible(role_selector)).to_be_visible()
    _visible(role_selector).click()

    role_option = page.locator("mat-option").filter(
        has_text=re.compile(rf"^\s*{re.escape(selected_role)}\s*$")
    )
    expect(_visible(role_option)).to_be_visible()
    _visible(role_option).click(force=True)
    expect(form.get_by_text(selected_role, exact=True)).to_be_visible()
    page.keyboard.press("Escape")


def _logout(page: Page) -> None:
    account_menu_candidates = [
        page.get_by_role("button", name="Account"),
        page.get_by_role("button", name="Profile"),
        page.get_by_role("button", name="User"),
        page.get_by_alt_text("Profile Image"),
        page.get_by_text(re.compile(r"^Hi .*!")),
        page.get_by_text("Ortholegal Admin", exact=True),
    ]
    for account_menu in account_menu_candidates:
        if account_menu.count() > 0 and _visible(account_menu).is_visible():
            _visible(account_menu).click()
            break

    logout_button = page.get_by_role("button", name="Logout").or_(
        page.get_by_role("menuitem", name="Logout")
    ).or_(
        page.get_by_role("button", name="Sign out")
    ).or_(
        page.get_by_role("menuitem", name="Sign out")
    ).or_(
        page.get_by_text("Logout", exact=True)
    ).or_(
        page.get_by_text("Sign out", exact=True)
    )
    expect(_visible(logout_button)).to_be_visible()
    _visible(logout_button).click()
    expect(page.get_by_label("Email")).to_be_visible()


@report_case(
    test_case_id="SUPERADMIN_USER_TC_001",
    module="Super Admin - Roles and Users",
    test_case_name="Creation of Case Manager and Login Validation",
    title="Verify that if the Super Admin creates a Case Manager using valid test data, then the Case Manager is created successfully and can log in using the same credentials.",
    test_type="Positive",
    preconditions="The OrthoLegal web portal is accessible. A valid Super Admin test account exists. Case Manager creation data is available in the Users worksheet of the existing Excel input workbook. The Case Manager signature file exists in the repository.",
    steps="1. Log in as Super Admin using environment credentials and the displayed verification PIN.\n2. Navigate to Database.\n3. Open Roles and Users.\n4. Verify the Users screen is displayed.\n5. Open the Add User form.\n6. Select the role whose visible name matches the Select Role value from Excel.\n7. Enter Case Manager details from Excel test data.\n8. Upload the Case Manager signature PNG from the repository test_data folder.\n9. Save the user.\n10. Search by the created Case Manager name.\n11. Verify the Users search results show the expected Case Manager name and Ortho Legal email.\n12. Log out the Super Admin.\n13. Log in as the created Case Manager using the email and password from the same Excel row and the displayed verification PIN.\n14. Verify that the authenticated Case Manager home page is displayed.",
    test_data="User creation and Case Manager login data from the Users worksheet in test_data/Users_Test_Data.xlsx; signature from test_data/Signatures/CaseManagetSignature.png. Password and verification PIN are not reported.",
    expected_result="The Case Manager should be created successfully with the data provided in the Users Excel sheet, should appear in the Users list with the correct name and Ortho Legal email, and should successfully log in using the same credentials after completing verification.",
    actual_result="The Case Manager was created successfully using the Excel test data, the created user was displayed with the expected name and Ortho Legal email, and the Case Manager successfully logged in after verification.",
)
def test_creation_of_case_manager_and_login_validation(page: Page) -> None:
    case_manager = _case_manager_data()
    selected_role = _required_text(case_manager, "Select Role")
    case_manager_name = _case_manager_name(case_manager)
    ortho_legal_email = _required_text(case_manager, "Ortho Legal Email")
    personal_email = _required_text(case_manager, "Personal Email")
    case_manager_password = _required_text(case_manager, "Password")
    confirm_password = _required_text(case_manager, "Confirm Password")
    assert case_manager_password == confirm_password, "Password and Confirm Password in Users_Test_Data.xlsx must match."

    assert SIGNATURE_PATH.exists(), f"Signature file was not found at expected path: {SIGNATURE_PATH}"

    authenticate_super_admin(page)

    database_link = page.get_by_role("link", name="Database").or_(page.get_by_text("Database", exact=True))
    expect(_visible(database_link)).to_be_visible()
    _visible(database_link).click()

    roles_and_users_link = page.get_by_role("link", name="Roles & Users").or_(
        page.get_by_text("Roles & Users", exact=True)
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

    _select_role_from_excel(page, form, selected_role)

    _fill_field(form, re.compile(r"^First name$", re.I), _required_text(case_manager, "First Name"))
    _fill_field(form, re.compile(r"^Middle name$", re.I), str(case_manager.get("Middle Name") or "").strip())
    _fill_field(form, re.compile(r"^Last name$", re.I), _required_text(case_manager, "Last Name"))
    _fill_field(form, re.compile(r"^Ortho\s*Legal email$", re.I), ortho_legal_email)
    _fill_field(form, re.compile(r"^Personal email$", re.I), personal_email)
    _fill_field(form, re.compile(r"^Password$", re.I), case_manager_password)
    _fill_field(form, re.compile(r"^Confirm password$", re.I), confirm_password)

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

    expect(form).not_to_be_visible(timeout=20000)

    search_field = page.get_by_placeholder("Search by Name").or_(page.get_by_label("Search by Name"))
    expect(_visible(search_field)).to_be_visible()
    _visible(search_field).fill(case_manager_name)

    matching_row = page.get_by_role("row").filter(has_text=case_manager_name).filter(has_text=ortho_legal_email)
    expect(_visible(matching_row)).to_be_visible()
    expect(_visible(matching_row).get_by_text(case_manager_name, exact=True)).to_be_visible()
    expect(_visible(matching_row).get_by_text(ortho_legal_email, exact=True)).to_be_visible()

    _logout(page)
    authenticate_with_credentials(page, ortho_legal_email, case_manager_password)
