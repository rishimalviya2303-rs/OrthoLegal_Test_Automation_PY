import re
from typing import Any


from playwright.sync_api import BrowserContext, Locator, Page, expect  # pyrefly: ignore[missing-import]

from conftest import test_case as report_case
from utils.auth import authenticate_super_admin, authenticate_with_credentials
from utils.excel_reader import PROJECT_ROOT, read_worksheet


USER_TEST_DATA_PATH = PROJECT_ROOT / "test_data" / "Users_Test_Data.xlsx"
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


def _editor_manager_data() -> dict[str, Any]:
    rows = read_worksheet("Users", workbook_path=USER_TEST_DATA_PATH, required_columns=REQUIRED_COLUMNS)
    matches = [
        row for row in rows
        if str(row.get("Select Role", "")).strip().lower() == "editor manager"
    ]
    assert matches, (
        "No Editor Manager creation row found in test_data/Users_Test_Data.xlsx. "
        "Add a Users row with Select Role=Editor Manager, First Name, Middle Name, Last Name, "
        "Ortho Legal Email, Personal Email, Password, and Confirm Password."
    )
    return matches[0]


def _required_text(row: dict[str, Any], column: str) -> str:
    value = row.get(column)
    assert value is not None and str(value).strip(), f"Users test data column '{column}' must contain a value."
    return str(value).strip()


def _editor_manager_name(row: dict[str, Any]) -> str:
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
    sign_out_button = page.get_by_role("button", name="Sign Out").or_(
        page.get_by_role("menuitem", name="Sign Out")
    ).or_(
        page.get_by_text("Sign Out", exact=True)
    ).or_(
        page.get_by_text("Sign out", exact=True)
    )
    profile_click_targets = [
        page.get_by_alt_text("Profile Image"),
        page.get_by_text(re.compile(r"^Hi\s+.*!$")),
        page.get_by_text("Super Admin", exact=True),
    ]

    for profile_target in profile_click_targets:
        if profile_target.count() == 0:
            continue

        _visible(profile_target).click(force=True)
        try:
            expect(_visible(sign_out_button)).to_be_visible(timeout=3000)
            break
        except AssertionError:
            continue
    else:
        profile_text = page.get_by_text(re.compile(r"^Hi\s+.*!$"))
        expect(_visible(profile_text)).to_be_visible()
        _visible(profile_text).evaluate(
            "element => (element.closest('button') || element.closest('[role=button]') || element.parentElement)?.click()"
        )

    expect(_visible(sign_out_button)).to_be_visible()
    _visible(sign_out_button).click()
    expect(page.get_by_role("textbox", name="Email")).to_be_visible()


def _close_unexpected_google_drive_pages(context: BrowserContext, app_page: Page) -> list[str]:
    closed_urls = []
    for open_page in context.pages:
        if open_page == app_page or open_page.is_closed():
            continue

        page_url = open_page.url
        if "drive.google.com" in page_url or "accounts.google.com" in page_url:
            closed_urls.append(page_url)
            open_page.close()

    app_page.bring_to_front()
    return closed_urls


def _save_user_and_return_to_users_screen(
    page: Page,
    form: Locator,
    user_name: str,
    ortho_legal_email: str,
) -> None:
    save_button = form.get_by_role("button", name="Save")
    expect(save_button).to_be_enabled()
    save_button.click()

    try:
        expect(form).not_to_be_visible(timeout=30000)
    except AssertionError:
        form_errors = form.locator("mat-error, .mat-error, [role='alert']").filter(
            has_text=re.compile(r"\S")
        )
        assert form_errors.count() == 0, (
            f"User could not be saved. Validation/error message displayed: {_visible(form_errors).inner_text()}"
        )

        close_button = form.get_by_role("button").first
        close_button.click()
        expect(form).not_to_be_visible(timeout=10000)

    _close_unexpected_google_drive_pages(page.context, page)
    expect(page).to_have_url(re.compile(r".*/#/database/user&roles$"))


@report_case(
    test_case_id="SUPERADMIN_USER_TC_009",
    module="Super Admin - Roles and Users",
    test_case_name="Creation of Editor Manager and Login Validation",
    title="Verify that if the Super Admin creates an Editor Manager using valid Excel test data, then the Editor Manager is created successfully and can log in using the same credentials.",
    test_type="Positive",
    preconditions="The OrthoLegal web portal is accessible. A valid Super Admin test account exists. Editor Manager creation data is available in the Users worksheet of the existing Excel input workbook.",
    steps="1. Log in as Super Admin using environment credentials and the displayed verification PIN.\n2. Navigate to Database.\n3. Open Roles and Users.\n4. Verify the Users screen is displayed.\n5. Open the Add User form.\n6. Select the role whose visible name matches the Select Role value from Excel.\n7. Enter Editor Manager details from Excel test data.\n8. Save the user.\n9. Search by the created Editor Manager name.\n10. Verify the Users search results show the expected Editor Manager name and Ortho Legal email.\n11. Log out the Super Admin.\n12. Log in as the created Editor Manager using the email and password from the same Excel row and the displayed verification PIN.\n13. Verify that the authenticated Editor Manager home page is displayed.",
    test_data="User creation and Editor Manager login data from the Users worksheet in test_data/Users_Test_Data.xlsx. Password and verification PIN are not reported.",
    expected_result="The Editor Manager should be created successfully using the data provided in the Users Excel sheet, should appear in the Users search results with the correct name and Ortho Legal email, and should successfully log in using the same credentials after completing six-digit verification.",
    actual_result="The Editor Manager was created successfully using the Excel test data, the created user was displayed with the expected name and Ortho Legal email, and the Editor Manager successfully logged in after verification.",
)
def test_creation_of_editor_manager_and_login_validation(page: Page) -> None:
    editor_manager = _editor_manager_data()
    selected_role = _required_text(editor_manager, "Select Role")
    editor_manager_name = _editor_manager_name(editor_manager)
    ortho_legal_email = _required_text(editor_manager, "Ortho Legal Email")
    personal_email = _required_text(editor_manager, "Personal Email")
    editor_manager_password = _required_text(editor_manager, "Password")
    confirm_password = _required_text(editor_manager, "Confirm Password")
    assert editor_manager_password == confirm_password, "Password and Confirm Password in Users_Test_Data.xlsx must match."

    # ── Step 1: Super Admin Login ──
    authenticate_super_admin(page)

    # ── Step 2–4: Navigate to Database → Roles and Users → Users tab ──
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

    # ── Step 5: Open Add User form ──
    add_user_button = page.get_by_role("button", name="Add User")
    expect(add_user_button).to_be_visible()
    add_user_button.click()

    add_user_form = page.get_by_role("dialog").filter(has_text="Add User")
    if add_user_form.count() == 0:
        add_user_form = page.locator("form").filter(has_text="Editor Manager")
    expect(_visible(add_user_form)).to_be_visible()
    form = _visible(add_user_form)

    # ── Step 6: Select Editor Manager role ──
    _select_role_from_excel(page, form, selected_role)

    # ── Step 7: Populate Excel-driven fields ──
    _fill_field(form, re.compile(r"^First name$", re.I), _required_text(editor_manager, "First Name"))
    _fill_field(form, re.compile(r"^Middle name$", re.I), str(editor_manager.get("Middle Name") or "").strip())
    _fill_field(form, re.compile(r"^Last name$", re.I), _required_text(editor_manager, "Last Name"))
    _fill_field(form, re.compile(r"^Ortho\s*Legal email$", re.I), ortho_legal_email)
    _fill_field(form, re.compile(r"^Personal email$", re.I), personal_email)
    _fill_field(form, re.compile(r"^Password$", re.I), editor_manager_password)
    _fill_field(form, re.compile(r"^Confirm password$", re.I), confirm_password)

    # NOTE: Editor Manager does NOT require a signature upload — intentionally skipped.

    # ── Step 8: Save and return to Users screen ──
    _save_user_and_return_to_users_screen(page, form, editor_manager_name, ortho_legal_email)

    # ── Step 9–10: Search and validate created Editor Manager ──
    search_field = page.get_by_placeholder("Search by Name").or_(page.get_by_label("Search by Name"))
    expect(_visible(search_field)).to_be_visible()
    _visible(search_field).fill(editor_manager_name)

    matching_row = page.get_by_role("row").filter(has_text=editor_manager_name).filter(has_text=ortho_legal_email)
    expect(_visible(matching_row)).to_be_visible()
    expect(_visible(matching_row).get_by_text(editor_manager_name, exact=True)).to_be_visible()
    expect(_visible(matching_row).get_by_text(ortho_legal_email, exact=True)).to_be_visible()

    # ── Step 11: Logout Super Admin ──
    _logout(page)

    # ── Step 12–13: Login as Editor Manager and complete six-digit verification ──
    authenticate_with_credentials(page, ortho_legal_email, editor_manager_password)
