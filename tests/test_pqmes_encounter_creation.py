import re
from typing import Any

from playwright.sync_api import Page, expect  # pyrefly: ignore[missing-import]

from conftest import test_case as report_case
from tests.test_pqmeencounter_creation import (
    ENCOUNTER_ROW_NUMBER,
    EncounterDataError,
    _applicant_name,
    _applicant_table_name,
    _contains_display_text,
    _field,
    _fill_applicant_details,
    _fill_insurance_details,
    _fill_injury_details,
    _fill_party_details,
    _fill_text,
    _find_upload_file,
    _format_ui_date,
    _handle_generated_document_tab,
    _open_new_encounter,
    _open_view_encounters,
    _physician_table_name,
    _required_value,
    _section,
    _select_date,
    _select_dropdown,
    _select_searchable_dropdown,
    _upload_documents,
    _visible,
)
from utils.auth import authenticate_super_admin
from utils.encounter_test_data import (
    PQMES_ENCOUNTER_SHEET_NAME,
    PQMES_ENCOUNTER_WORKBOOK_PATH,
    read_encounter_test_data,
)


PQMES_REQUIRED_EXCEL_FIELDS = {
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
        "Date of Cover Letter",
        "Date of Declaration",
        "Requesting Party",
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


def _validate_required_pqmes_excel_data(data: dict[str, dict[str, Any]]) -> None:
    missing = []
    for section, fields in PQMES_REQUIRED_EXCEL_FIELDS.items():
        for field in fields:
            value = data[section][field]
            if value is None or not str(value).strip():
                missing.append(f"{section} -> {field}")

    if missing:
        raise EncounterDataError(
            "PQMES Encounter Excel row 3 is not ready for end-to-end creation. "
            "Populate these required values: "
            + "; ".join(missing)
        )

    encounter_type = _required_value(data, "Encounter Details", "Encounter Type")
    if encounter_type.casefold() != "pqmes":
        raise EncounterDataError(
            "PQMES Encounter Excel row 3 must create a PQMES encounter: "
            f"Encounter Details -> Encounter Type is '{encounter_type}', expected 'PQMES'."
        )


def _fill_pqmes_encounter_details(page: Page, data: dict[str, dict[str, Any]]) -> None:
    section = _section(page, "Encounter details")
    _select_searchable_dropdown(page, section, "Physician", _required_value(data, "Encounter Details", "Physician"))
    _select_dropdown(page, section, "Medical Evaluation Type", _required_value(data, "Encounter Details", "Medical Evaluation Type"))
    _select_dropdown(page, section, "Encounter type", _required_value(data, "Encounter Details", "Encounter Type"))
    _fill_text(section, "Panel Number", _required_value(data, "Encounter Details", "Panel Number"))
    _select_date(page, section, "Date of Cover Letter", _required_value(data, "Encounter Details", "Date of Cover Letter"))
    _select_date(page, section, "Date of Declaration", _required_value(data, "Encounter Details", "Date of Declaration"))
    _select_dropdown(page, section, "Requesting party", _required_value(data, "Encounter Details", "Requesting Party"))

    for pqme_only_label in (
        "Date of Appointment Call",
        "Appointment date & time",
        "Examination address",
    ):
        assert _field(section, pqme_only_label).count() == 0, (
            f"PQMES flow should not use PQME-only Encounter Details field: {pqme_only_label}"
        )


def _confirm_pqmes_encounter_information(page: Page) -> str:
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


def _complete_pqmes_confirmation(page: Page) -> str:
    submit_button = page.get_by_role("button", name=re.compile(r"^Submit$", re.I)).last
    expect(submit_button).to_be_enabled()
    submit_button.click()
    return _confirm_pqmes_encounter_information(page)


def _validate_pqmes_encounter_row(page: Page, data: dict[str, dict[str, Any]]) -> dict[str, str]:
    _open_view_encounters(page)
    applicant_name = _applicant_name(data)
    applicant_table_name = _applicant_table_name(data)
    physician = _required_value(data, "Encounter Details", "Physician")
    physician_table_name = _physician_table_name(physician)
    encounter_type = _required_value(data, "Encounter Details", "Encounter Type")
    medical_evaluation_type = _required_value(data, "Encounter Details", "Medical Evaluation Type")
    cover_letter_date = _format_ui_date(_required_value(data, "Encounter Details", "Date of Cover Letter"))
    declaration_date = _format_ui_date(_required_value(data, "Encounter Details", "Date of Declaration"))

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
            "Created PQMES row was not found by Applicant name and Encounter Type. "
            f"Expected Applicant='{applicant_table_name}', Encounter Type='{encounter_type}'. "
            f"Visible rows: {visible_rows}"
        ) from exc

    row_text = _visible(row).inner_text()
    assert _contains_display_text(row_text, applicant_table_name), "Created PQMES row does not show the expected Applicant name."
    assert _contains_display_text(row_text, physician_table_name), "Created PQMES row does not show the expected Physician."
    assert _contains_display_text(row_text, encounter_type), "Created PQMES row does not show the expected Encounter Type."
    if _contains_display_text(row_text, medical_evaluation_type):
        displayed_pqmes_value = medical_evaluation_type
    elif cover_letter_date in row_text:
        displayed_pqmes_value = cover_letter_date
    elif declaration_date in row_text:
        displayed_pqmes_value = declaration_date
    else:
        displayed_pqmes_value = "Not displayed in View Encounters row"
    return {
        "Applicant Name": applicant_name,
        "Applicant Name Actual": applicant_table_name,
        "Physician": physician,
        "Physician Actual": physician_table_name,
        "Encounter Type": encounter_type,
        "Medical Evaluation Type": medical_evaluation_type,
        "Displayed PQMES Value": displayed_pqmes_value,
        "Date of Cover Letter": cover_letter_date,
        "Date of Declaration": declaration_date,
    }


@report_case(
    test_case_id="SUPERADMIN_PQMES_ENCOUNTER_TC_001",
    module="Super Admin - PQMES Encounter Creation",
    test_case_name="Create PQMES Encounter and Validate Encounter",
    title=(
        "Verify that if the Super Admin creates a PQMES encounter using valid Excel test data "
        "and confirms the encounter information, then the PQMES encounter is created successfully "
        "and the correct encounter details are displayed in View Encounters."
    ),
    test_type="Positive",
    preconditions=(
        "Valid Super Admin credentials must be configured in the environment, valid PQMES Encounter "
        "test data must be available in PQMES_Encounter_Test_Data.xlsx, the referenced Physician "
        "and required master values must exist in the application, and all required Encounter upload "
        "files must be available in the configured test-data folder."
    ),
    steps=(
        "1. Read PQMES Encounter data from the two-level Section -> Field Excel workbook.\n"
        "2. Validate required PQMES Encounter Excel data and upload files.\n"
        "3. Login as Super Admin using existing environment credentials and verification helper.\n"
        "4. Navigate to View Encounters.\n"
        "5. Open New Encounter.\n"
        "6. Fill Applicant Details.\n"
        "7. Fill Injury Details and select requested body parts.\n"
        "8. Fill PQMES Encounter Details.\n"
        "9. Upload Strike Panel and Attorney Demo files.\n"
        "10. Fill Applicant Attorney, Employer, Defense Attorney, and Insurance Company details.\n"
        "11. Save the Encounter and confirm the PQMES encounter-information popup.\n"
        "12. Handle the generated ANF/PDF tab if one appears and return to OrthoLegal.\n"
        "13. Validate the created PQMES Encounter row in View Encounters."
    ),
    test_data=(
        "PQMES Encounter data row 3 from test_data/PQMES_Encounter_Test_Data.xlsx, "
        "worksheet PQMES_Encounter_Test_Data, and upload files from test_data/encounterfiles. "
        "Super Admin password and verification PIN are not reported."
    ),
    expected_result=(
        "The Super Admin should be able to create a PQMES Encounter using the data provided in the "
        "PQMES Encounter Excel test-data sheet. After completing the form and confirming that the "
        "Encounter information is correct, the PQMES Encounter should be created successfully and "
        "should appear in View Encounters with the correct Applicant, Physician, Encounter Type, "
        "and other applicable PQMES Encounter details."
    ),
    actual_result=(
        "The Super Admin created the PQMES Encounter from Excel data, confirmed the encounter "
        "information popup, returned to View Encounters, and validated the created row against "
        "the expected Applicant, Physician, Encounter Type, and applicable PQMES values."
    ),
)
def test_create_pqmes_encounter_and_validate_encounter(page: Page) -> None:
    stage = "reading PQMES Encounter Excel data"
    confirmation_text = ""
    try:
        encounter_data = read_encounter_test_data(
            row_number=ENCOUNTER_ROW_NUMBER,
            workbook_path=PQMES_ENCOUNTER_WORKBOOK_PATH,
            sheet_name=PQMES_ENCOUNTER_SHEET_NAME,
        )
        _validate_required_pqmes_excel_data(encounter_data)
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

        stage = "filling PQMES Encounter Details"
        _fill_pqmes_encounter_details(page, encounter_data)

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

        stage = "saving and confirming the PQMES Encounter"
        app_page = page
        before_pages = set(page.context.pages)
        confirmation_text = _complete_pqmes_confirmation(page)
        document_tab_appeared = any(candidate not in before_pages for candidate in page.context.pages)
        _handle_generated_document_tab(page.context, app_page)

        stage = "validating the created PQMES Encounter row"
        actual_values = _validate_pqmes_encounter_row(app_page, encounter_data)
        print(
            "Created PQMES Encounter validated: "
            f"Applicant expected={actual_values['Applicant Name']}; "
            f"Applicant actual={actual_values['Applicant Name Actual']}; "
            f"Physician expected={actual_values['Physician']}; "
            f"Physician actual={actual_values['Physician Actual']}; "
            f"Encounter Type expected=PQMES; "
            f"Encounter Type actual={actual_values['Encounter Type']}; "
            f"Medical Evaluation Type={actual_values['Medical Evaluation Type']}; "
            f"Displayed PQMES Value={actual_values['Displayed PQMES Value']}; "
            f"Date of Cover Letter={actual_values['Date of Cover Letter']}; "
            f"Date of Declaration={actual_values['Date of Declaration']}; "
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
            f"Create PQMES Encounter and Validate Encounter failed while {stage} ({type(exc).__name__}). "
            f"Last confirmation popup text: {confirmation_text}"
        ) from None
