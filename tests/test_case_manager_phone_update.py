import re
from typing import Any

from playwright.sync_api import Page, expect  # pyrefly: ignore[missing-import]

from conftest import test_case as report_case
from utils.auth import authenticate_with_credentials
from utils.excel_reader import PROJECT_ROOT, read_worksheet
from utils.profile import dismiss_profile_picture_popup_if_present


USER_TEST_DATA_PATH = PROJECT_ROOT / "test_data" / "Users_Test_Data.xlsx"
REQUIRED_COLUMNS = {
    "Select Role",
    "Ortho Legal Email",
    "Password",
    "Updated Phone Number",
}


def _case_manager_data() -> dict[str, Any]:
    """Return the Case Manager row from the Users worksheet."""
    rows = read_worksheet("Users", workbook_path=USER_TEST_DATA_PATH, required_columns=REQUIRED_COLUMNS)
    matches = [
        row for row in rows
        if str(row.get("Select Role", "")).strip().lower() == "case manager"
    ]
    assert matches, (
        "No Case Manager row found in test_data/Users_Test_Data.xlsx. "
        "Add a Users row with Select Role=Case Manager, Ortho Legal Email, Password, "
        "and Updated Phone Number."
    )
    return matches[0]


def _required_text(row: dict[str, Any], column: str) -> str:
    """Return a non-blank string value for the given column or fail clearly."""
    value = row.get(column)
    assert value is not None and str(value).strip(), f"Users test data column '{column}' must contain a value."
    return str(value).strip()


def _normalize_phone(raw_value: Any) -> str:
    """Convert an Excel phone value to a plain 10-digit string."""
    if isinstance(raw_value, float) and raw_value.is_integer():
        text = str(int(raw_value))
    elif isinstance(raw_value, int):
        text = str(raw_value)
    else:
        text = str(raw_value).strip()

    assert re.fullmatch(r"[0-9]{10}", text), (
        "Updated Phone Number for Case Manager must contain exactly 10 digits."
    )
    return text


def _strip_digits(formatted_phone: str) -> str:
    """Remove all non-digit characters from a formatted phone string."""
    return re.sub(r"\D", "", formatted_phone)


@report_case(
    test_case_id="CASEMANAGER_PROFILE_TC_001",
    module="Case Manager - Profile Settings",
    test_case_name="Case Manager Phone Number Update and Validation",
    title=(
        "Verify that if the Case Manager updates the phone number from Profile Settings "
        "using a valid 10-digit phone number, then the updated phone number is saved and "
        "displayed correctly."
    ),
    test_type="Positive",
    preconditions=(
        "A valid Case Manager user must already exist and its login credentials and "
        "Updated Phone Number must be available in the Users test-data sheet."
    ),
    steps=(
        "1. Read Case Manager credentials and Updated Phone Number from Users Excel data.\n"
        "2. Login as Case Manager.\n"
        "3. Complete six-digit login verification.\n"
        "4. Dismiss the optional profile-picture popup using its Cancel button if visible.\n"
        "5. Open the top-right Case Manager profile menu.\n"
        "6. Click Settings.\n"
        "7. Open Profile Settings.\n"
        "8. Click Edit.\n"
        "9. Update only the field named Phone with the 10-digit Updated Phone Number from Excel.\n"
        "10. Save the profile.\n"
        "11. Validate that the displayed Phone value matches the Excel phone number "
        "after normalization."
    ),
    test_data=(
        "Case Manager login credentials and Updated Phone Number from the Users worksheet "
        "in test_data/Users_Test_Data.xlsx. Password and verification PIN are not reported."
    ),
    expected_result=(
        "The Case Manager should be able to log in successfully, dismiss the optional "
        "profile-picture popup if it appears, update the phone number from Profile Settings "
        "using the 10-digit value provided in the Users Excel sheet, save the changes "
        "successfully, and see the same phone number displayed on the profile."
    ),
    actual_result=(
        "The Case Manager updated the Phone field from Profile Settings, and after saving, "
        "the displayed Phone value represents the same 10-digit number provided in the "
        "Users Excel test data."
    ),
)
def test_case_manager_phone_number_update_and_validation(page: Page) -> None:
    stage = "reading Case Manager data from Users Excel"

    try:
        case_manager = _case_manager_data()
        ortho_legal_email = _required_text(case_manager, "Ortho Legal Email")
        case_manager_password = _required_text(case_manager, "Password")

        raw_phone = case_manager.get("Updated Phone Number")
        assert raw_phone is not None, (
            "Users test data column 'Updated Phone Number' must contain a value."
        )
        phone_digits = _normalize_phone(raw_phone)

        stage = "logging in as Case Manager"
        authenticate_with_credentials(page, ortho_legal_email, case_manager_password)

        stage = "dismissing the optional profile-picture popup"
        dismiss_profile_picture_popup_if_present(page)

        stage = "opening Profile Settings"
        page.get_by_text(re.compile(r"^Hi\s+.*!$")).click()
        page.get_by_role("button", name="Settings", exact=True).click()
        expect(page).to_have_url(re.compile(r".*/#/admin$"))

        stage = "enabling profile edit mode"
        edit_control = page.get_by_text("Edit", exact=True)
        expect(edit_control).to_be_visible()
        edit_control.click()
        save_button = page.get_by_role("button", name="Save", exact=True)
        expect(save_button).to_be_visible()

        stage = "updating only the Phone field"
        phone_input = page.get_by_placeholder("Enter Phone Number", exact=True)
        expect(phone_input).to_be_editable()
        expect(phone_input).to_have_accessible_name("Phone")
        phone_input.fill(phone_digits)
        phone_input.blur()

        stage = "saving the profile"
        expect(save_button).to_be_enabled()
        save_button.click()
        page.wait_for_load_state("networkidle")

        stage = "validating the saved Phone field after reload"
        page.reload(wait_until="domcontentloaded")
        displayed_phone_locator = page.get_by_placeholder("Enter Phone Number", exact=True)
        expect(displayed_phone_locator).to_have_accessible_name("Phone")
        displayed_phone_text = displayed_phone_locator.input_value()
        normalized_ui_phone = _strip_digits(displayed_phone_text)
        assert normalized_ui_phone == phone_digits, (
            "Saved profile Phone value does not match the Users Excel phone."
        )
        print(f"Saved profile Phone: {displayed_phone_text}")
    except Exception as exc:
        raise AssertionError(
            f"Case Manager phone update failed while {stage} ({type(exc).__name__})."
        ) from None
