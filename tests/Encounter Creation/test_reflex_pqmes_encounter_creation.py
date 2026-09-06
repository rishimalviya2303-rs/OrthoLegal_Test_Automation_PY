import importlib.util
import re
import sys
from pathlib import Path
from typing import Any

from playwright.sync_api import Page, expect  # pyrefly: ignore[missing-import]

from conftest import test_case as report_case
from utils.auth import authenticate_super_admin
from utils.encounter_test_data import (
    REFLEX_PQMES_ENCOUNTER_DETAILS_FIELDS,
    REFLEX_PQMES_ENCOUNTER_SHEET_NAME,
    REFLEX_PQMES_ENCOUNTER_WORKBOOK_PATH,
    read_encounter_test_data,
)


_PQME_HELPERS_PATH = Path(__file__).with_name("test_pqme_encounter_creation.py")
_PQME_HELPERS_SPEC = importlib.util.spec_from_file_location("encounter_creation_pqme_helpers", _PQME_HELPERS_PATH)
assert _PQME_HELPERS_SPEC and _PQME_HELPERS_SPEC.loader, f"Could not load PQME helpers from {_PQME_HELPERS_PATH}"
_pqme_helpers = importlib.util.module_from_spec(_PQME_HELPERS_SPEC)
sys.modules[_PQME_HELPERS_SPEC.name] = _pqme_helpers
_PQME_HELPERS_SPEC.loader.exec_module(_pqme_helpers)

ENCOUNTER_ROW_NUMBER = _pqme_helpers.ENCOUNTER_ROW_NUMBER
EncounterDataError = _pqme_helpers.EncounterDataError
_applicant_name = _pqme_helpers._applicant_name
_applicant_table_name = _pqme_helpers._applicant_table_name
_contains_display_text = _pqme_helpers._contains_display_text
_field = _pqme_helpers._field
_fill_applicant_details = _pqme_helpers._fill_applicant_details
_fill_insurance_details = _pqme_helpers._fill_insurance_details
_fill_injury_details = _pqme_helpers._fill_injury_details
_fill_party_details = _pqme_helpers._fill_party_details
_fill_text = _pqme_helpers._fill_text
_find_upload_file = _pqme_helpers._find_upload_file
_format_ui_date = _pqme_helpers._format_ui_date
_handle_generated_document_tab = _pqme_helpers._handle_generated_document_tab
_open_new_encounter = _pqme_helpers._open_new_encounter
_open_view_encounters = _pqme_helpers._open_view_encounters
_physician_table_name = _pqme_helpers._physician_table_name
_required_value = _pqme_helpers._required_value
_section = _pqme_helpers._section
_select_date = _pqme_helpers._select_date
_select_dropdown = _pqme_helpers._select_dropdown
_select_searchable_dropdown = _pqme_helpers._select_searchable_dropdown
_upload_documents = _pqme_helpers._upload_documents
_visible = _pqme_helpers._visible


REFLEX_PQMES_REQUIRED_EXCEL_FIELDS = {
    "Applicant Details": (
        "First Name",
        "Last Name",
        "Date of Birth",
        "Gender",
        "SSN",
        "Address Line 1",
        "City",
        "State",
        "Zip",
        "Home Phone",
        "Cell Phone",
        "Work Phone",
    ),
    "Injury Details": (
        "Date of Injury",
        "Injured Body Parts",
        "WCAB Number",
        "Claim Number",
    ),
    "Encounter Details": REFLEX_PQMES_ENCOUNTER_DETAILS_FIELDS,
    "Applicant Attorney Details": (
        "Firm",
        "First Name",
        "Last Name",
        "Address Line 1",
        "City",
        "State",
        "Zip",
        "Email",
    ),
    "Employer Details": (
        "Employer",
    ),
    "Defense Attorney Details": (
        "Firm",
        "First Name",
        "Last Name",
        "Address Line 1",
        "City",
        "State",
        "Zip",
        "Email",
    ),
    "Insurance Company Details": (
        "Insurance Company",
        "Insurance Company Phone",
        "Adjuster First Name",
        "Adjuster Last Name",
        "Address Line 1",
        "City",
        "State",
        "Zip",
        "Adjuster Email",
    ),
}


def _validate_required_reflex_pqmes_excel_data(data: dict[str, dict[str, Any]]) -> None:
    missing = []
    for section, fields in REFLEX_PQMES_REQUIRED_EXCEL_FIELDS.items():
        for field in fields:
            value = data[section][field]
            if value is None or not str(value).strip():
                missing.append(f"{section} -> {field}")

    if missing:
        raise EncounterDataError(
            "Reflex PQMES Encounter Excel row 3 is not ready for end-to-end creation. "
            "Populate these required values: "
            + "; ".join(missing)
        )

    encounter_details = data["Encounter Details"]
    actual_fields = tuple(encounter_details.keys())
    if actual_fields != REFLEX_PQMES_ENCOUNTER_DETAILS_FIELDS:
        raise EncounterDataError(
            "Reflex PQMES Encounter Details headers do not match the expected field order. "
            f"Actual fields: {actual_fields}"
        )

    forbidden_fields = (
        "Date of Cover Letter",
        "Date of Declaration",
        "Date of Appointment Call",
        "Appointment Date and Time",
        "Examination Address",
        "Address Line 1",
        "Address Line 2",
        "City",
        "State",
        "Zip",
    )
    unexpected = [field for field in forbidden_fields if field in encounter_details]
    if unexpected:
        raise EncounterDataError(
            "Reflex PQMES Encounter Details contains fields that should not be used: "
            + ", ".join(unexpected)
        )

    encounter_type = _required_value(data, "Encounter Details", "Encounter Type")
    normalized_encounter_type = encounter_type.casefold()
    if "reflex" not in normalized_encounter_type or "pqmes" not in normalized_encounter_type:
        raise EncounterDataError(
            "Reflex PQMES Encounter Excel row 3 must create a Reflex PQMES encounter: "
            f"Encounter Details -> Encounter Type is '{encounter_type}', expected a Reflex PQMES encounter type."
        )


def _fill_reflex_pqmes_encounter_details(page: Page, data: dict[str, dict[str, Any]]) -> None:
    section = _section(page, "Encounter details")
    _select_searchable_dropdown(page, section, "Physician", _required_value(data, "Encounter Details", "Physician"))
    _select_dropdown(page, section, "Medical Evaluation Type", _required_value(data, "Encounter Details", "Medical Evaluation Type"))
    _select_dropdown(page, section, "Encounter type", _required_value(data, "Encounter Details", "Encounter Type"))
    _fill_text(section, "Panel Number", _required_value(data, "Encounter Details", "Panel Number"))
    _select_reflex_date_of_receipt(page, section, _required_value(data, "Encounter Details", "Date of Receipt"))
    _select_dropdown(page, section, "Requesting party", _required_value(data, "Encounter Details", "Requesting Party"))

    for non_reflex_label in (
        "Date of Cover Letter",
        "Date of Declaration",
        "Date of Appointment Call",
        "Appointment date & time",
        "Examination address",
    ):
        assert _field(section, non_reflex_label).count() == 0, (
            f"Reflex PQMES flow should not use this Encounter Details field: {non_reflex_label}"
        )


def _select_reflex_date_of_receipt(page: Page, section: Any, mm_dd_yyyy: str) -> None:
    date_label = _reflex_date_of_receipt_label(section)
    try:
        _select_date(page, section, date_label, mm_dd_yyyy)
        return
    except AssertionError:
        pass

    parsed = _pqme_helpers.datetime.strptime(mm_dd_yyyy.split()[0], "%m/%d/%Y")
    control = _field(section, date_label)
    expect(control).to_be_visible()
    control.click()

    target_label = f"{parsed.strftime('%B')} {parsed.day}, {parsed.year}"
    date_button = page.get_by_role("button", name=re.compile(rf"^\s*{re.escape(target_label)}\s*$", re.I))
    for _ in range(36):
        if date_button.count() and date_button.first.is_visible():
            date_button.first.click()
            break
        next_month = page.get_by_role("button", name=re.compile(r"next month", re.I))
        if next_month.count() == 0 or not next_month.first.is_enabled():
            break
        next_month.first.click()
    else:
        raise AssertionError(f"Could not find Date of Receipt calendar day '{target_label}'.")

    if date_button.count() == 0:
        control.fill(mm_dd_yyyy)
        control.blur()

    page.keyboard.press("Escape")
    expect(control).to_have_value(re.compile(r"\d{1,2}[/-]\d{1,2}[/-]\d{4}"))


def _reflex_date_of_receipt_label(section: Any) -> str:
    for label in ("Date of Receipt", "Date of reciept"):
        if _field(section, label).count():
            return label
    return "Date of Receipt"


def _confirm_reflex_pqmes_encounter_information(page: Page) -> str:
    dialog = page.get_by_role("dialog").or_(page.locator("mat-dialog-container")).filter(
        has_text=re.compile(r"confirm|correct|editable|encounter information", re.I)
    )
    expect(_visible(dialog)).to_be_visible(timeout=30000)
    modal = _visible(dialog)
    dialog_text = re.sub(r"\s+", " ", modal.inner_text()).strip()
    confirm_button = modal.get_by_role("button", name=re.compile(r"^Confirm$", re.I))
    expect(confirm_button).to_be_enabled()
    confirm_button.click()
    expect(modal).not_to_be_visible(timeout=60000)
    return dialog_text


def _complete_reflex_pqmes_confirmation(page: Page) -> str:
    submit_button = page.get_by_role("button", name=re.compile(r"^Submit$", re.I)).last
    expect(submit_button).to_be_enabled()
    submit_button.click()
    return _confirm_reflex_pqmes_encounter_information(page)


def _validate_reflex_pqmes_encounter_row(page: Page, data: dict[str, dict[str, Any]]) -> dict[str, str]:
    _open_view_encounters(page)
    applicant_name = _applicant_name(data)
    applicant_table_name = _applicant_table_name(data)
    physician = _required_value(data, "Encounter Details", "Physician")
    physician_table_name = _physician_table_name(physician)
    encounter_type = _required_value(data, "Encounter Details", "Encounter Type")
    medical_evaluation_type = _required_value(data, "Encounter Details", "Medical Evaluation Type")

    main = page.locator("main")
    search = main.get_by_placeholder(re.compile("Applicant name", re.I)).or_(
        main.get_by_role("textbox", name=re.compile("Applicant name", re.I))
    ).or_(
        main.get_by_role("searchbox", name=re.compile(r"Search", re.I))
    )
    row = page.get_by_role("row").filter(has_text=applicant_table_name).filter(has_text=encounter_type)
    search_terms = [
        applicant_table_name,
        applicant_table_name.title(),
        _required_value(data, "Applicant Details", "First Name"),
        _required_value(data, "Applicant Details", "First Name").title(),
        "",
    ]
    if search.count():
        for term in dict.fromkeys(search_terms):
            search.first.fill(term)
            page.keyboard.press("Enter")
            try:
                expect(_visible(row)).to_be_visible(timeout=8000)
                break
            except AssertionError:
                continue

    try:
        expect(_visible(row)).to_be_visible(timeout=30000)
    except AssertionError as exc:
        visible_rows = page.get_by_role("row").all_inner_texts()
        raise AssertionError(
            "Created Reflex PQMES row was not found by Applicant name and Encounter Type. "
            f"Expected Applicant='{applicant_table_name}', Encounter Type='{encounter_type}'. "
            f"Visible rows: {visible_rows}"
        ) from exc

    row_text = _visible(row).inner_text()
    assert _contains_display_text(row_text, applicant_table_name), (
        "Created Reflex PQMES row does not show the expected Applicant name."
    )
    assert _contains_display_text(row_text, physician_table_name), (
        "Created Reflex PQMES row does not show the expected Physician."
    )
    assert _contains_display_text(row_text, encounter_type), (
        "Created Reflex PQMES row does not show the expected Encounter Type."
    )

    displayed_reflex_value = (
        medical_evaluation_type
        if _contains_display_text(row_text, medical_evaluation_type)
        else "Not displayed in View Encounters row"
    )
    return {
        "Applicant Name": applicant_name,
        "Applicant Name Actual": applicant_table_name,
        "Physician": physician,
        "Physician Actual": physician_table_name,
        "Encounter Type": encounter_type,
        "Medical Evaluation Type": medical_evaluation_type,
        "Displayed Reflex PQMES Value": displayed_reflex_value,
    }


@report_case(
    test_case_id="SUPERADMIN_REFLEX_PQMES_ENCOUNTER_TC_001",
    module="Super Admin - Reflex PQMES Encounter Creation",
    test_case_name="Create Reflex PQMES Encounter and Validate Encounter",
    title=(
        "Verify that if the Super Admin creates a Reflex PQMES encounter using valid Excel test data, "
        "then the Reflex PQMES encounter is created successfully and the correct encounter details "
        "are displayed in View Encounters."
    ),
    test_type="Positive",
    preconditions=(
        "Valid Super Admin credentials must be configured in the environment, valid Reflex PQMES "
        "Encounter test data must be available in Reflex_PQMES_Encounter_Test_Data.xlsx, and all "
        "required master data and upload files must exist."
    ),
    steps=(
        "1. Read Reflex PQMES Encounter data from the two-level Section -> Field Excel workbook.\n"
        "2. Validate required Reflex PQMES Encounter Excel data and upload files.\n"
        "3. Login as Super Admin using existing environment credentials and verification helper.\n"
        "4. Navigate to View Encounters.\n"
        "5. Open New Encounter.\n"
        "6. Fill Applicant Details.\n"
        "7. Fill Injury Details and select requested body parts.\n"
        "8. Fill Reflex PQMES Encounter Details using Date of Receipt.\n"
        "9. Upload Strike Panel and Attorney Demo files.\n"
        "10. Fill Applicant Attorney, Employer, Defense Attorney, and Insurance Company details.\n"
        "11. Save the Encounter and confirm the Reflex PQMES encounter-information popup.\n"
        "12. Handle the generated ANF/PDF tab if one appears and return to OrthoLegal.\n"
        "13. Validate the created Reflex PQMES Encounter row in View Encounters."
    ),
    test_data=(
        "Reflex PQMES Encounter data row 3 from test_data/Reflex_PQMES_Encounter_Test_Data.xlsx, "
        f"worksheet {REFLEX_PQMES_ENCOUNTER_SHEET_NAME}, and upload files from test_data/encounterfiles. "
        "Super Admin password and verification PIN are not reported."
    ),
    expected_result=(
        "The Super Admin should be able to create a Reflex PQMES Encounter using the data provided "
        "in the Reflex PQMES Encounter Excel test-data sheet, and after successful creation the "
        "correct Reflex PQMES Encounter details should be displayed in View Encounters."
    ),
    actual_result=(
        "The Super Admin created the Reflex PQMES Encounter from Excel data, confirmed the encounter "
        "information popup, returned to View Encounters, and validated the created row against the "
        "expected Applicant, Physician, and Encounter Type."
    ),
)
def test_create_reflex_pqmes_encounter_and_validate_encounter(page: Page) -> None:
    stage = "reading Reflex PQMES Encounter Excel data"
    confirmation_text = ""
    try:
        encounter_data = read_encounter_test_data(
            row_number=ENCOUNTER_ROW_NUMBER,
            workbook_path=REFLEX_PQMES_ENCOUNTER_WORKBOOK_PATH,
            sheet_name=REFLEX_PQMES_ENCOUNTER_SHEET_NAME,
        )
        _validate_required_reflex_pqmes_excel_data(encounter_data)
        strike_panel = _find_upload_file("StrikePanel")
        attorney_demo = _find_upload_file("AttorneyDemo")

        stage = "logging in as Super Admin"
        authenticate_super_admin(page)

        stage = "opening View Encounters"
        _open_view_encounters(page)

        stage = "opening New Encounter"
        _open_new_encounter(page)

        stage = "filling Applicant Details"
        _fill_applicant_details(page, encounter_data)

        stage = "filling Injury Details"
        _fill_injury_details(page, encounter_data)

        stage = "filling Reflex PQMES Encounter Details"
        _fill_reflex_pqmes_encounter_details(page, encounter_data)

        stage = "uploading Encounter documents"
        _upload_documents(page, _section(page, "Upload documents"))

        stage = "filling Applicant Attorney Details"
        _fill_party_details(page, encounter_data, "Applicant attorney details")

        stage = "filling Employer Details"
        _fill_party_details(page, encounter_data, "Employer details")

        stage = "filling Defense Attorney Details"
        _fill_party_details(page, encounter_data, "Defense attorney details")

        stage = "filling Insurance Company Details"
        _fill_insurance_details(page, encounter_data)

        stage = "saving and confirming the Reflex PQMES Encounter"
        app_page = page
        before_pages = set(page.context.pages)
        confirmation_text = _complete_reflex_pqmes_confirmation(page)
        document_tab_appeared = any(candidate not in before_pages for candidate in page.context.pages)
        _handle_generated_document_tab(page.context, app_page)

        stage = "validating the created Reflex PQMES Encounter row"
        actual_values = _validate_reflex_pqmes_encounter_row(app_page, encounter_data)
        print(
            "Created Reflex PQMES Encounter validated: "
            f"Applicant expected={actual_values['Applicant Name']}; "
            f"Applicant actual={actual_values['Applicant Name Actual']}; "
            f"Physician expected={actual_values['Physician']}; "
            f"Physician actual={actual_values['Physician Actual']}; "
            f"Encounter Type expected={actual_values['Encounter Type']}; "
            f"Encounter Type actual={actual_values['Encounter Type']}; "
            f"Medical Evaluation Type={actual_values['Medical Evaluation Type']}; "
            f"Displayed Reflex PQMES Value={actual_values['Displayed Reflex PQMES Value']}; "
            f"Date of Receipt={_format_ui_date(_required_value(encounter_data, 'Encounter Details', 'Date of Receipt'))}; "
            f"Confirmation Popup={confirmation_text}; "
            f"Document Tab Appeared={document_tab_appeared}; "
            f"Strike Panel={strike_panel.name}; Attorney Demo={attorney_demo.name}"
        )
    except EncounterDataError:
        raise
    except AssertionError:
        raise
    except Exception as exc:
        raise AssertionError(
            f"Create Reflex PQMES Encounter and Validate Encounter failed while {stage} ({type(exc).__name__}). "
            f"Last confirmation popup text: {confirmation_text}"
        ) from None
