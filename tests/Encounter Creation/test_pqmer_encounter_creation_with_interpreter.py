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

_PQME_WITH_INTERPRETER_PATH = Path(__file__).with_name("test_pqme_encounter_creation_with_interpreter.py")
_PQME_WITH_INTERPRETER_SPEC = importlib.util.spec_from_file_location("pqme_with_interpreter", _PQME_WITH_INTERPRETER_PATH)
assert _PQME_WITH_INTERPRETER_SPEC and _PQME_WITH_INTERPRETER_SPEC.loader, f"Could not load PQME interpreter helpers"
_pqme_with_interpreter = importlib.util.module_from_spec(_PQME_WITH_INTERPRETER_SPEC)
sys.modules[_PQME_WITH_INTERPRETER_SPEC.name] = _pqme_with_interpreter
_PQME_WITH_INTERPRETER_SPEC.loader.exec_module(_pqme_with_interpreter)

INTERPRETER_REQUIRED_EXCEL_FIELDS = _pqme_with_interpreter.INTERPRETER_REQUIRED_EXCEL_FIELDS
_fill_interpreter_details = _pqme_with_interpreter._fill_interpreter_details

_PQMER_HELPERS_PATH = Path(__file__).with_name("test_pqmer_encounter_creation_without_interpreter.py")
_PQMER_HELPERS_SPEC = importlib.util.spec_from_file_location("pqmer_helpers", _PQMER_HELPERS_PATH)
assert _PQMER_HELPERS_SPEC and _PQMER_HELPERS_SPEC.loader, f"Could not load PQMER helpers"
_pqmer_helpers = importlib.util.module_from_spec(_PQMER_HELPERS_SPEC)
sys.modules[_PQMER_HELPERS_SPEC.name] = _pqmer_helpers
_PQMER_HELPERS_SPEC.loader.exec_module(_pqmer_helpers)

_validate_required_pqmer_excel_data = _pqmer_helpers._validate_required_pqmer_excel_data
_validate_pqmer_encounter_row = _pqmer_helpers._validate_pqmer_encounter_row


@report_case(
    test_case_id="SUPERADMIN_PQMER_ENCOUNTER_TC_002",
    module="Super Admin - PQMER Encounter Creation",
    test_case_name="Create PQMER Encounter With Interpreter and Validate Encounter",
    title=(
        "Verify that if the Super Admin creates a PQMER encounter with an interpreter using valid Excel test data, "
        "then the encounter is created successfully with the interpreter details and the correct encounter information is displayed."
    ),
    test_type="Positive",
    preconditions=(
        "Valid Super Admin credentials must be configured, valid PQMER Encounter "
        "test data including Interpreter Details must be available in the PQMER Encounter Excel workbook, "
        "and all required master data and upload files must exist."
    ),
    steps=(
        "1. Read PQMER Encounter data from the two-level Section -> Field Excel workbook.\n"
        "2. Validate required PQMER Encounter Excel data, Interpreter Details, and upload files.\n"
        "3. Login as Super Admin using existing environment credentials and verification helper.\n"
        "4. Navigate to View Encounters.\n"
        "5. Open New Encounter.\n"
        "6. Fill Applicant Details.\n"
        "7. Fill Injury Details and select requested body parts.\n"
        "8. Fill Encounter Details using the PQMER workbook values.\n"
        "9. Upload Strike Panel and Attorney Demo files.\n"
        "10. Check Interpreter Required and fill Language, Vendor Name, and Vendor Phone.\n"
        "11. Fill Applicant Attorney, Employer, Defense Attorney, and Insurance Company details.\n"
        "12. Save the Encounter and complete the required PQMER save workflow.\n"
        "13. Handle the generated ANF/PDF tab if one appears and return to OrthoLegal.\n"
        "14. Validate the created PQMER Encounter row in View Encounters."
    ),
    test_data=(
        "PQMER Encounter data row 3 from test_data/PQMER_Encounter_Test_Data.xlsx, "
        "worksheet Encounter_Test_Data, including Interpreter Details "
        "(Language, Vendor Name, Vendor Phone), and upload files from test_data/encounterfiles. "
        "Super Admin password and verification PIN are not reported."
    ),
    expected_result=(
        "The Super Admin should be able to create a PQMER Encounter with the Interpreter Required option selected and valid Language, Vendor Name, and Vendor Phone values from the PQMER Encounter Excel test-data sheet, and the Encounter should be created successfully."
    ),
    actual_result=(
        "The Super Admin created the PQMER Encounter with Interpreter Details from Excel data, completed the required save "
        "workflow, returned to View Encounters, and validated the created row against the expected "
        "Applicant, Physician, Encounter Type, Evaluation Type, and Appointment Date."
    ),
)
def test_create_pqmer_encounter_with_interpreter_and_validate_encounter(page: Page) -> None:
    stage = "reading PQMER Encounter Excel data"
    try:
        encounter_data = read_encounter_test_data(
            row_number=ENCOUNTER_ROW_NUMBER,
            workbook_path=PQMER_ENCOUNTER_WORKBOOK_PATH,
            sheet_name=PQMER_ENCOUNTER_SHEET_NAME,
        )
        _validate_required_pqmer_excel_data(encounter_data)
        for interp_section, interp_fields in INTERPRETER_REQUIRED_EXCEL_FIELDS.items():
            for interp_field in interp_fields:
                value = encounter_data.get(interp_section, {}).get(interp_field)
                if value is None or not str(value).strip():
                    raise EncounterDataError(
                        f"Required Excel test data is missing: {interp_section} -> {interp_field}"
                    )
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
        
        stage = "filling Interpreter Details"
        _fill_interpreter_details(page, encounter_data)

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
        interpreter = encounter_data.get("Interpreter Details", {})
        print(
            "Created PQMER Encounter (With Interpreter) validated: "
            f"Applicant expected={actual_values['Applicant Name']}; "
            f"Applicant actual={actual_values['Applicant Name Actual']}; "
            f"Physician expected={actual_values['Physician']}; "
            f"Physician actual={actual_values['Physician Actual']}; "
            f"Encounter Type expected=PQMER; "
            f"Encounter Type actual={actual_values['Encounter Type']}; "
            f"Evaluation Type={actual_values['Evaluation Type']}; "
            f"Appointment Date={actual_values['Appointment Date']}; "
            f"Document Tab Appeared={document_tab_appeared}; "
            f"Strike Panel={strike_panel.name}; Attorney Demo={attorney_demo.name}; "
            f"Language={interpreter.get('Language')}; "
            f"Vendor Name={interpreter.get('Vendor Name')}; "
            f"Vendor Phone={interpreter.get('Vendor Phone')}"
        )
    except EncounterDataError:
        raise
    except AssertionError:
        raise
    except Exception as exc:
        raise AssertionError(
            f"Create PQMER Encounter With Interpreter and Validate Encounter failed while {stage} ({type(exc).__name__})."
        ) from None
