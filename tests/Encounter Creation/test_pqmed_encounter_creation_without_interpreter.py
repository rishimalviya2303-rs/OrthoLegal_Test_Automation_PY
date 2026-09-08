import importlib.util
import re
import sys
from pathlib import Path
from typing import Any

from playwright.sync_api import Page, expect  # pyrefly: ignore[missing-import]

from conftest import test_case as report_case
from utils.auth import authenticate_super_admin
from utils.encounter_test_data import (
    PQMED_ENCOUNTER_SHEET_NAME,
    PQMED_ENCOUNTER_WORKBOOK_PATH,
    read_encounter_test_data,
)


_PQME_HELPERS_PATH = Path(__file__).with_name("test_pqme_encounter_creation_without_interpreter.py")
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
_optional_value = _pqme_helpers._optional_value
_physician_table_name = _pqme_helpers._physician_table_name
_required_value = _pqme_helpers._required_value
_save_encounter = _pqme_helpers._save_encounter
_section = _pqme_helpers._section
_section_xpath = _pqme_helpers._section_xpath
_select_appointment_date_time = _pqme_helpers._select_appointment_date_time
_select_date = _pqme_helpers._select_date
_select_dropdown = _pqme_helpers._select_dropdown
_select_searchable_dropdown = _pqme_helpers._select_searchable_dropdown
_select_state = _pqme_helpers._select_state
_upload_documents = _pqme_helpers._upload_documents
_visible = _pqme_helpers._visible


PQMED_REQUIRED_EXCEL_FIELDS = {
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
        "Address Line 1",
        "City",
        "State",
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


def _validate_required_pqmed_excel_data(data: dict[str, dict[str, Any]]) -> None:
    missing = []
    for section, fields in PQMED_REQUIRED_EXCEL_FIELDS.items():
        for field in fields:
            value = data[section][field]
            if value is None or not str(value).strip():
                missing.append(f"{section} -> {field}")

    if missing:
        raise EncounterDataError(
            "PQMED Encounter Excel row 3 is not ready for end-to-end creation. "
            "Populate these required values: "
            + "; ".join(missing)
        )

    appointment_value = _required_value(data, "Encounter Details", "Appointment Date and Time")
    if not re.search(r"\b\d{1,2}:\d{2}\s*(?:AM|PM)\b", appointment_value, re.I):
        raise EncounterDataError(
            "PQMED Encounter Excel row 3 is missing appointment time: "
            "Encounter Details -> Appointment Date and Time must include both date and time."
        )

    encounter_type = _required_value(data, "Encounter Details", "Encounter Type")
    if encounter_type.casefold() != "pqmed":
        raise EncounterDataError(
            "PQMED Encounter Excel row 3 must create a PQMED encounter: "
            f"Encounter Details -> Encounter Type is '{encounter_type}', expected 'PQMED'."
        )


def _fill_pqmed_encounter_details(page: Page, data: dict[str, dict[str, Any]]) -> None:
    section = _section(page, "Encounter details")
    _select_searchable_dropdown(page, section, "Physician", _required_value(data, "Encounter Details", "Physician"))
    _select_dropdown(page, section, "Medical Evaluation Type", _required_value(data, "Encounter Details", "Medical Evaluation Type"))
    _select_dropdown(page, section, "Encounter type", _required_value(data, "Encounter Details", "Encounter Type"))
    _fill_text(section, "Panel Number", _required_value(data, "Encounter Details", "Panel Number"))
    _select_date(page, section, "Date of Appointment Call", _required_value(data, "Encounter Details", "Date of Appointment Call"))
    _select_appointment_date_time(page, section, _required_value(data, "Encounter Details", "Appointment Date and Time"))
    _select_dropdown(page, section, "Requesting party", _required_value(data, "Encounter Details", "Requesting Party"))
    _fill_text(section, "Address line 1", _required_value(data, "Encounter Details", "Address Line 1"))
    _fill_text(section, "Address line 2", _optional_value(data, "Encounter Details", "Address Line 2"))
    _fill_text(section, "City", _required_value(data, "Encounter Details", "City"))
    _select_state(page, section, _required_value(data, "Encounter Details", "State"))
    _fill_text(section, "Zip", _required_value(data, "Encounter Details", "Zip"))

    assert _field(section, "Examination address").count() == 0, (
        "PQMED flow should use direct Encounter address fields, not PQME Examination Address."
    )


def _complete_pqmed_finalization(page: Page) -> str:
    submit_button = page.get_by_role("button", name=re.compile(r"^Submit$", re.I)).last
    if submit_button.count():
        expect(submit_button).to_be_enabled()
        submit_button.click()
    else:
        _save_encounter(page)

    recipients_dialog = page.get_by_role("dialog").or_(page.locator("mat-dialog-container")).filter(
        has_text=re.compile(r"recipient|send anf", re.I)
    )
    if recipients_dialog.count():
        try:
            expect(_visible(recipients_dialog)).to_be_visible(timeout=30000)
            dialog_text = re.sub(r"\s+", " ", _visible(recipients_dialog).inner_text()).strip()
            _complete_recipients_popup(page)
            return dialog_text
        except AssertionError:
            pass

    confirm_button = page.get_by_role("button", name=re.compile(r"^Confirm$", re.I))
    if confirm_button.count():
        confirm_dialog = page.get_by_role("dialog").or_(page.locator("mat-dialog-container")).filter(
            has=confirm_button.first
        )
        expect(_visible(confirm_dialog)).to_be_visible(timeout=30000)
        dialog_text = re.sub(r"\s+", " ", _visible(confirm_dialog).inner_text()).strip()
        expect(_visible(confirm_dialog).get_by_role("button", name=re.compile(r"^Confirm$", re.I))).to_be_enabled()
        _visible(confirm_dialog).get_by_role("button", name=re.compile(r"^Confirm$", re.I)).click()
        expect(_visible(confirm_dialog)).not_to_be_visible(timeout=60000)
        return dialog_text

    page.wait_for_url(re.compile(r".*/#/encounter_details$"), timeout=60000)
    return "No confirmation popup displayed"


def _upload_pqmed_documents_if_present(page: Page) -> tuple[Path | None, Path | None, str]:
    upload_heading = page.get_by_text(re.compile(r"^\s*Upload documents\s*$", re.I))
    if upload_heading.count() == 0:
        return None, None, "Upload documents section not displayed for PQMED"

    section = _section(page, "Upload documents")
    file_inputs = page.locator(f"xpath=//input[@type='file' and {_section_xpath(section)}]")
    if file_inputs.count() == 0:
        return None, None, "Upload documents section displayed without file inputs for PQMED"

    strike_panel, attorney_demo = _upload_documents(page, section)
    return strike_panel, attorney_demo, "Uploaded Strike Panel and Attorney Demo"


def _validate_pqmed_encounter_row(page: Page, data: dict[str, dict[str, Any]]) -> dict[str, str]:
    _open_view_encounters(page)
    applicant_name = _applicant_name(data)
    applicant_table_name = _applicant_table_name(data)
    physician = _required_value(data, "Encounter Details", "Physician")
    physician_table_name = _physician_table_name(physician)
    encounter_type = _required_value(data, "Encounter Details", "Encounter Type")
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
            "Created PQMED row was not found by Applicant name and Encounter Type. "
            f"Expected Applicant='{applicant_table_name}', Encounter Type='{encounter_type}'. "
            f"Visible rows: {visible_rows}"
        ) from exc

    row_text = _visible(row).inner_text()
    assert _contains_display_text(row_text, applicant_table_name), "Created PQMED row does not show the expected Applicant name."
    assert _contains_display_text(row_text, physician_table_name), "Created PQMED row does not show the expected Physician."
    assert _contains_display_text(row_text, encounter_type), "Created PQMED row does not show the expected Encounter Type."
    assert appointment_date in row_text, "Created PQMED row does not show the expected Appointment Date."
    return {
        "Applicant Name": applicant_name,
        "Applicant Name Actual": applicant_table_name,
        "Physician": physician,
        "Physician Actual": physician_table_name,
        "Encounter Type": encounter_type,
        "Appointment Date": appointment_date,
    }


@report_case(
    test_case_id="SUPERADMIN_PQMED_ENCOUNTER_TC_001",
    module="Super Admin - PQMED Encounter Creation",
    test_case_name="Create PQMED Encounter Without Interpreter and Validate Encounter",
    title=(
        "Verify that if the Super Admin creates a PQMED encounter without an interpreter using valid Excel test data, "
        "then the encounter is created successfully and the correct encounter details are "
        "displayed in View Encounters."
    ),
    test_type="Positive",
    preconditions=(
        "Valid Super Admin credentials must be configured in the environment, valid PQMED Encounter "
        "test data must be available in PQMED_Encounter_Test_Data.xlsx, and all referenced master "
        "data and required upload files must exist in the application/project."
    ),
    steps=(
        "1. Read PQMED Encounter data from the two-level Section -> Field Excel workbook.\n"
        "2. Validate required PQMED Encounter Excel data and upload files.\n"
        "3. Login as Super Admin using existing environment credentials and verification helper.\n"
        "4. Navigate to View Encounters.\n"
        "5. Open New Encounter.\n"
        "6. Fill Applicant Details.\n"
        "7. Fill Injury Details and select requested body parts.\n"
        "8. Fill PQMED Encounter Details including direct address fields.\n"
        "9. Upload Strike Panel and Attorney Demo files.\n"
        "10. Fill Applicant Attorney, Employer, Defense Attorney, and Insurance Company details.\n"
        "11. Save the Encounter and complete the actual PQMED finalization workflow.\n"
        "12. Handle the generated ANF/PDF tab if one appears and return to OrthoLegal.\n"
        "13. Validate the created PQMED Encounter row in View Encounters."
    ),
    test_data=(
        "PQMED Encounter data row 3 from test_data/PQMED_Encounter_Test_Data.xlsx, "
        "worksheet PQMED_Encounter_Test_Data, and upload files from test_data/encounterfiles. "
        "Interpreter Details are not used. Super Admin password and verification PIN are not reported."
    ),
    expected_result=(
        "The Super Admin should be able to create a PQMED Encounter without an interpreter using the data provided in the "
        "PQMED Encounter Excel test-data sheet, and after successful creation the correct PQMED "
        "Encounter details should be displayed in View Encounters."
    ),
    actual_result=(
        "The Super Admin created the PQMED Encounter without an interpreter from Excel data, completed the required "
        "finalization workflow, returned to View Encounters, and validated the created row against "
        "the expected Applicant, Physician, Encounter Type, and Appointment Date."
    ),
)
def test_create_pqmed_encounter_without_interpreter_and_validate_encounter(page: Page) -> None:
    stage = "reading PQMED Encounter Excel data"
    finalization_text = ""
    try:
        encounter_data = read_encounter_test_data(
            row_number=ENCOUNTER_ROW_NUMBER,
            workbook_path=PQMED_ENCOUNTER_WORKBOOK_PATH,
            sheet_name=PQMED_ENCOUNTER_SHEET_NAME,
        )
        _validate_required_pqmed_excel_data(encounter_data)
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

        stage = "filling PQMED Encounter Details"
        _fill_pqmed_encounter_details(page, encounter_data)

        stage = "uploading Encounter documents"
        strike_panel, attorney_demo, upload_status = _upload_pqmed_documents_if_present(page)

        stage = "filling Applicant Attorney Details"
        _fill_party_details(page, encounter_data, "Applicant attorney details")

        stage = "filling Employer Details"
        _fill_party_details(page, encounter_data, "Employer details")

        stage = "filling Defense Attorney Details"
        _fill_party_details(page, encounter_data, "Defense attorney details")

        stage = "filling Insurance Company Details"
        _fill_insurance_details(page, encounter_data)

        stage = "saving and completing the PQMED Encounter workflow"
        app_page = page
        before_pages = set(page.context.pages)
        finalization_text = _complete_pqmed_finalization(page)
        document_tab_appeared = any(candidate not in before_pages for candidate in page.context.pages)
        _handle_generated_document_tab(page.context, app_page)

        stage = "validating the created PQMED Encounter row"
        actual_values = _validate_pqmed_encounter_row(app_page, encounter_data)
        print(
            "Created PQMED Encounter (Without Interpreter) validated: "
            f"Applicant expected={actual_values['Applicant Name']}; "
            f"Applicant actual={actual_values['Applicant Name Actual']}; "
            f"Physician expected={actual_values['Physician']}; "
            f"Physician actual={actual_values['Physician Actual']}; "
            f"Encounter Type expected=PQMED; "
            f"Encounter Type actual={actual_values['Encounter Type']}; "
            f"Appointment Date expected={actual_values['Appointment Date']}; "
            f"Appointment Date actual={actual_values['Appointment Date']}; "
            f"Finalization Flow={finalization_text}; "
            f"Document Tab Appeared={document_tab_appeared}; "
            f"Upload Status={upload_status}; "
            f"Strike Panel={strike_panel.name if strike_panel else 'Not uploaded'}; "
            f"Attorney Demo={attorney_demo.name if attorney_demo else 'Not uploaded'}"
        )
    except EncounterDataError:
        raise
    except AssertionError:
        raise
    except Exception as exc:
        raise AssertionError(
            f"Create PQMED Encounter Without Interpreter and Validate Encounter failed while {stage} ({type(exc).__name__}). "
            f"Last finalization flow text: {finalization_text}"
        ) from None
