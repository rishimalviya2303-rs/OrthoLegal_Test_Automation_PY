import importlib.util
import re
import sys
from pathlib import Path
from typing import Any

from playwright.sync_api import Page, expect  # pyrefly: ignore[missing-import]

from conftest import test_case as report_case
from utils.auth import authenticate_super_admin
from utils.encounter_test_data import (
    PQMER_ENCOUNTER_SHEET_NAME,
    PQMER_ENCOUNTER_WORKBOOK_PATH,
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
_complete_recipients_popup = _pqme_helpers._complete_recipients_popup
_contains_display_text = _pqme_helpers._contains_display_text
_fill_applicant_details = _pqme_helpers._fill_applicant_details
_fill_encounter_details = _pqme_helpers._fill_encounter_details
_fill_insurance_details = _pqme_helpers._fill_insurance_details
_fill_injury_details = _pqme_helpers._fill_injury_details
_fill_party_details = _pqme_helpers._fill_party_details
_find_upload_file = _pqme_helpers._find_upload_file
_format_ui_date = _pqme_helpers._format_ui_date
_handle_generated_document_tab = _pqme_helpers._handle_generated_document_tab
_open_new_encounter = _pqme_helpers._open_new_encounter
_open_view_encounters = _pqme_helpers._open_view_encounters
_physician_table_name = _pqme_helpers._physician_table_name
_required_value = _pqme_helpers._required_value
_save_encounter = _pqme_helpers._save_encounter
_section = _pqme_helpers._section
_upload_documents = _pqme_helpers._upload_documents
_visible = _pqme_helpers._visible


PQMER_REQUIRED_EXCEL_FIELDS = {
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
    "Encounter Details": (
        "Physician",
        "Medical Evaluation Type",
        "Encounter Type",
        "Panel Number",
        "Date of Appointment Call",
        "Appointment Date and Time",
        "Requesting Party",
        "Examination Address",
        "City",
        "Zip",
    ),
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


def _validate_required_pqmer_excel_data(data: dict[str, dict[str, Any]]) -> None:
    missing = []
    for section, fields in PQMER_REQUIRED_EXCEL_FIELDS.items():
        for field in fields:
            value = data[section][field]
            if value is None or not str(value).strip():
                missing.append(f"{section} -> {field}")

    if missing:
        raise EncounterDataError(
            "PQMER Encounter Excel row 3 is not ready for end-to-end creation. "
            "Populate these required values: "
            + "; ".join(missing)
        )

    appointment_value = _required_value(data, "Encounter Details", "Appointment Date and Time")
    if not re.search(r"\b\d{1,2}:\d{2}\s*(?:AM|PM)\b", appointment_value, re.I):
        raise EncounterDataError(
            "PQMER Encounter Excel row 3 is missing appointment time: "
            "Encounter Details -> Appointment Date and Time must include both date and time."
        )

    encounter_type = _required_value(data, "Encounter Details", "Encounter Type")
    if encounter_type.casefold() != "pqmer":
        raise EncounterDataError(
            "PQMER Encounter Excel row 3 must create a PQMER encounter: "
            f"Encounter Details -> Encounter Type is '{encounter_type}', expected 'PQMER'."
        )


def _validate_pqmer_encounter_row(page: Page, data: dict[str, dict[str, Any]]) -> dict[str, str]:
    _open_view_encounters(page)
    applicant_name = _applicant_name(data)
    applicant_table_name = _applicant_table_name(data)
    physician = _required_value(data, "Encounter Details", "Physician")
    physician_table_name = _physician_table_name(physician)
    encounter_type = _required_value(data, "Encounter Details", "Encounter Type")
    evaluation_type = str(data["Encounter Details"].get("Evaluation Type") or "").strip()
    appointment_date = _format_ui_date(_required_value(data, "Encounter Details", "Appointment Date and Time"))

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
            "Created PQMER row was not found by Applicant name and Encounter Type. "
            f"Expected Applicant='{applicant_table_name}', Encounter Type='{encounter_type}'. "
            f"Visible rows: {visible_rows}"
        ) from exc

    row_text = _visible(row).inner_text()
    assert _contains_display_text(row_text, applicant_table_name), "Created PQMER row does not show the expected Applicant name."
    assert _contains_display_text(row_text, physician_table_name), "Created PQMER row does not show the expected Physician."
    assert _contains_display_text(row_text, encounter_type), "Created PQMER row does not show the expected Encounter Type."
    if evaluation_type:
        assert _contains_display_text(row_text, evaluation_type), "Created PQMER row does not show the expected Evaluation Type."
    assert appointment_date in row_text, "Created PQMER row does not show the expected Appointment Date."
    return {
        "Applicant Name": applicant_name,
        "Applicant Name Actual": applicant_table_name,
        "Physician": physician,
        "Physician Actual": physician_table_name,
        "Encounter Type": encounter_type,
        "Evaluation Type": evaluation_type,
        "Appointment Date": appointment_date,
    }


@report_case(
    test_case_id="SUPERADMIN_PQMER_ENCOUNTER_TC_001",
    module="Super Admin - PQMER Encounter Creation",
    test_case_name="Create PQMER Encounter and Validate Encounter",
    title=(
        "Verify that if the Super Admin creates a PQMER encounter using valid Excel test data, "
        "then the PQMER encounter is created successfully and the correct encounter details are "
        "displayed in View Encounters."
    ),
    test_type="Positive",
    preconditions=(
        "Valid Super Admin credentials must be configured in the environment, valid PQMER Encounter "
        "test data must be available in PQMER_Encounter_Test_Data.xlsx, the referenced Physician "
        "and required master values must exist in the application, and all required Encounter upload "
        "files must be available in the configured test-data folder."
    ),
    steps=(
        "1. Read PQMER Encounter data from the two-level Section -> Field Excel workbook.\n"
        "2. Validate required PQMER Encounter Excel data and upload files.\n"
        "3. Login as Super Admin using existing environment credentials and verification helper.\n"
        "4. Navigate to View Encounters.\n"
        "5. Open New Encounter.\n"
        "6. Fill Applicant Details.\n"
        "7. Fill Injury Details and select requested body parts.\n"
        "8. Fill Encounter Details using the PQMER workbook values.\n"
        "9. Upload Strike Panel and Attorney Demo files.\n"
        "10. Fill Applicant Attorney, Employer, Defense Attorney, and Insurance Company details.\n"
        "11. Save the Encounter and complete the required PQMER save workflow.\n"
        "12. Handle the generated ANF/PDF tab if one appears and return to OrthoLegal.\n"
        "13. Validate the created PQMER Encounter row in View Encounters."
    ),
    test_data=(
        "PQMER Encounter data row 3 from test_data/PQMER_Encounter_Test_Data.xlsx, "
        "worksheet Encounter_Test_Data, and upload files from test_data/encounterfiles. "
        "Super Admin password and verification PIN are not reported."
    ),
    expected_result=(
        "The Super Admin should be able to create a PQMER Encounter using the data provided in the "
        "PQMER Encounter Excel test-data sheet, and after successful creation the correct PQMER "
        "Encounter details should be displayed in View Encounters."
    ),
    actual_result=(
        "The Super Admin created the PQMER Encounter from Excel data, completed the required save "
        "workflow, returned to View Encounters, and validated the created row against the expected "
        "Applicant, Physician, Encounter Type, Evaluation Type, and Appointment Date."
    ),
)
def test_create_pqmer_encounter_and_validate_encounter(page: Page) -> None:
    stage = "reading PQMER Encounter Excel data"
    try:
        encounter_data = read_encounter_test_data(
            row_number=ENCOUNTER_ROW_NUMBER,
            workbook_path=PQMER_ENCOUNTER_WORKBOOK_PATH,
            sheet_name=PQMER_ENCOUNTER_SHEET_NAME,
        )
        _validate_required_pqmer_excel_data(encounter_data)
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

        stage = "filling PQMER Encounter Details"
        _fill_encounter_details(page, encounter_data)

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

        stage = "saving and completing the PQMER Encounter workflow"
        app_page = page
        before_pages = set(page.context.pages)
        _save_encounter(page)
        _complete_recipients_popup(page)
        document_tab_appeared = any(candidate not in before_pages for candidate in page.context.pages)
        _handle_generated_document_tab(page.context, app_page)

        stage = "validating the created PQMER Encounter row"
        actual_values = _validate_pqmer_encounter_row(app_page, encounter_data)
        print(
            "Created PQMER Encounter validated: "
            f"Applicant expected={actual_values['Applicant Name']}; "
            f"Applicant actual={actual_values['Applicant Name Actual']}; "
            f"Physician expected={actual_values['Physician']}; "
            f"Physician actual={actual_values['Physician Actual']}; "
            f"Encounter Type expected=PQMER; "
            f"Encounter Type actual={actual_values['Encounter Type']}; "
            f"Evaluation Type={actual_values['Evaluation Type']}; "
            f"Appointment Date={actual_values['Appointment Date']}; "
            f"Document Tab Appeared={document_tab_appeared}; "
            f"Strike Panel={strike_panel.name}; Attorney Demo={attorney_demo.name}"
        )
    except EncounterDataError:
        raise
    except AssertionError:
        raise
    except Exception as exc:
        raise AssertionError(
            f"Create PQMER Encounter and Validate Encounter failed while {stage} ({type(exc).__name__})."
        ) from None
