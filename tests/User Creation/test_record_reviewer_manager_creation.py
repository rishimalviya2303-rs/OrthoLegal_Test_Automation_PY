import importlib.util
import re
import sys
from pathlib import Path
from typing import Any

from playwright.sync_api import Page, expect  # pyrefly: ignore[missing-import]

from conftest import test_case as report_case
from utils.auth import authenticate_super_admin, authenticate_with_credentials
from utils.excel_reader import read_worksheet


_RECORD_REVIEWER_HELPERS_PATH = Path(__file__).with_name("test_record_reviewer_creation.py")
_RECORD_REVIEWER_HELPERS_SPEC = importlib.util.spec_from_file_location(
    "user_creation_record_reviewer_helpers",
    _RECORD_REVIEWER_HELPERS_PATH,
)
assert _RECORD_REVIEWER_HELPERS_SPEC and _RECORD_REVIEWER_HELPERS_SPEC.loader, (
    f"Could not load Record Reviewer helpers from {_RECORD_REVIEWER_HELPERS_PATH}"
)
_record_reviewer_helpers = importlib.util.module_from_spec(_RECORD_REVIEWER_HELPERS_SPEC)
sys.modules[_RECORD_REVIEWER_HELPERS_SPEC.name] = _record_reviewer_helpers
_RECORD_REVIEWER_HELPERS_SPEC.loader.exec_module(_record_reviewer_helpers)

USER_TEST_DATA_PATH = _record_reviewer_helpers.USER_TEST_DATA_PATH
REQUIRED_COLUMNS = _record_reviewer_helpers.REQUIRED_COLUMNS
_close_unexpected_google_drive_pages = _record_reviewer_helpers._close_unexpected_google_drive_pages
_fill_field = _record_reviewer_helpers._fill_field
_logout = _record_reviewer_helpers._logout
_record_reviewer_name = _record_reviewer_helpers._record_reviewer_name
_required_text = _record_reviewer_helpers._required_text
_save_user_and_return_to_users_screen = _record_reviewer_helpers._save_user_and_return_to_users_screen
_select_role_from_excel = _record_reviewer_helpers._select_role_from_excel
_visible = _record_reviewer_helpers._visible


def _record_reviewer_manager_data() -> dict[str, Any]:
    rows = read_worksheet("Users", workbook_path=USER_TEST_DATA_PATH, required_columns=REQUIRED_COLUMNS)
    matches = [
        row for row in rows
        if str(row.get("Select Role", "")).strip().casefold() == "record reviewer manager"
    ]
    assert matches, (
        "No Record Reviewer Manager creation row found in test_data/Users_Test_Data.xlsx. "
        "Add a Users row with Select Role=Record Reviewer Manager, First Name, Middle Name, Last Name, "
        "Ortho Legal Email, Personal Email, Password, and Confirm Password."
    )
    return matches[0]


@report_case(
    test_case_id="SUPERADMIN_USER_TC_005",
    module="Super Admin - Roles and Users",
    test_case_name="Creation of Record Reviewer Manager and Login Validation",
    title=(
        "Verify that if the Super Admin creates a Record Reviewer Manager using valid Excel test data, "
        "then the Record Reviewer Manager is created successfully and can log in using the same credentials."
    ),
    test_type="Positive",
    preconditions=(
        "Valid Super Admin credentials must be configured in the environment, and valid Record Reviewer "
        "Manager test data must be available in the Users test-data sheet."
    ),
    steps=(
        "1. Log in as Super Admin using environment credentials and the displayed verification PIN.\n"
        "2. Navigate to Database.\n"
        "3. Open Roles and Users.\n"
        "4. Verify the Users screen is displayed.\n"
        "5. Open the Add User form.\n"
        "6. Select the Record Reviewer Manager role from the exact Select Role value in Excel.\n"
        "7. Enter Record Reviewer Manager details from Excel test data.\n"
        "8. Save the user.\n"
        "9. Search by the created Record Reviewer Manager name.\n"
        "10. Verify the Users search results show the expected Record Reviewer Manager name and Ortho Legal email.\n"
        "11. Log out the Super Admin.\n"
        "12. Log in as the created Record Reviewer Manager using the email and password from the same Excel row and the displayed verification PIN.\n"
        "13. Verify that the authenticated Record Reviewer Manager home page is displayed."
    ),
    test_data=(
        "User creation and Record Reviewer Manager login data from the Users worksheet in "
        "test_data/Users_Test_Data.xlsx. Password and verification PIN are not reported."
    ),
    expected_result=(
        "The Super Admin should be able to create the Record Reviewer Manager using the valid Excel test data. "
        "The created Record Reviewer Manager should appear with the correct name and Ortho Legal Email in "
        "Roles and Users and should be able to log in successfully using the same Excel credentials."
    ),
    actual_result=(
        "The Record Reviewer Manager was created successfully using the Excel test data, the created user "
        "was displayed with the expected name and Ortho Legal email, and the Record Reviewer Manager "
        "successfully logged in after verification."
    ),
)
def test_creation_of_record_reviewer_manager_and_login_validation(page: Page) -> None:
    record_reviewer_manager = _record_reviewer_manager_data()
    selected_role = _required_text(record_reviewer_manager, "Select Role")
    assert selected_role == "Record Reviewer Manager", (
        f"Expected exact Select Role=Record Reviewer Manager, got {selected_role!r}."
    )
    record_reviewer_manager_name = _record_reviewer_name(record_reviewer_manager)
    ortho_legal_email = _required_text(record_reviewer_manager, "Ortho Legal Email")
    personal_email = _required_text(record_reviewer_manager, "Personal Email")
    record_reviewer_manager_password = _required_text(record_reviewer_manager, "Password")
    confirm_password = _required_text(record_reviewer_manager, "Confirm Password")
    assert record_reviewer_manager_password == confirm_password, "Password and Confirm Password in Users_Test_Data.xlsx must match."

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
        add_user_form = page.locator("form").filter(has_text="Record Reviewer Manager")
    expect(_visible(add_user_form)).to_be_visible()
    form = _visible(add_user_form)

    _select_role_from_excel(page, form, selected_role)

    _fill_field(form, re.compile(r"^First name$", re.I), _required_text(record_reviewer_manager, "First Name"))
    _fill_field(form, re.compile(r"^Middle name$", re.I), str(record_reviewer_manager.get("Middle Name") or "").strip())
    _fill_field(form, re.compile(r"^Last name$", re.I), _required_text(record_reviewer_manager, "Last Name"))
    _fill_field(form, re.compile(r"^Ortho\s*Legal email$", re.I), ortho_legal_email)
    _fill_field(form, re.compile(r"^Personal email$", re.I), personal_email)
    _fill_field(form, re.compile(r"^Password$", re.I), record_reviewer_manager_password)
    _fill_field(form, re.compile(r"^Confirm password$", re.I), confirm_password)

    # Record Reviewer Manager follows Record Reviewer behavior and does not require a signature upload.

    _save_user_and_return_to_users_screen(page, form, record_reviewer_manager_name, ortho_legal_email)
    _close_unexpected_google_drive_pages(page.context, page)

    search_field = page.get_by_placeholder("Search by Name").or_(page.get_by_label("Search by Name"))
    expect(_visible(search_field)).to_be_visible()
    _visible(search_field).fill(record_reviewer_manager_name)

    matching_row = page.get_by_role("row").filter(has_text=record_reviewer_manager_name).filter(has_text=ortho_legal_email)
    expect(_visible(matching_row)).to_be_visible()
    expect(_visible(matching_row).get_by_text(record_reviewer_manager_name, exact=True)).to_be_visible()
    expect(_visible(matching_row).get_by_text(ortho_legal_email, exact=True)).to_be_visible()

    print(
        "Created Record Reviewer Manager validated: "
        f"Name expected={record_reviewer_manager_name}; "
        f"Name actual={record_reviewer_manager_name}; "
        f"Ortho Legal Email expected={ortho_legal_email}; "
        f"Ortho Legal Email actual={ortho_legal_email}; "
        "Signature Required=False"
    )

    _logout(page)
    authenticate_with_credentials(page, ortho_legal_email, record_reviewer_manager_password)
