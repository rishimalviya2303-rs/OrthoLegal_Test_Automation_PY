import re
from datetime import date, datetime, timedelta
from pathlib import Path
from dataclasses import dataclass
from typing import Any

from playwright.sync_api import BrowserContext, Locator, Page, expect  # pyrefly: ignore[missing-import]

from conftest import test_case as report_case
from utils.auth import authenticate_super_admin
from utils.encounter_test_data import (
    PQME_ENCOUNTER_SHEET_NAME,
    PQME_ENCOUNTER_WORKBOOK_PATH,
    read_encounter_test_data,
)
from utils.business_days import calculate_report_due_date
from utils.excel_reader import PROJECT_ROOT


ENCOUNTER_ROW_NUMBER = 3
ENCOUNTER_FILES_PATH = PROJECT_ROOT / "test_data" / "encounterfiles"

REQUIRED_EXCEL_FIELDS = {
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

INTERPRETER_REQUIRED_EXCEL_FIELDS = {
    "Interpreter Details": (
        "Language",
        "Vendor Name",
        "Vendor Phone",
    ),
}


class EncounterDataError(AssertionError):
    """Raised when Encounter Excel data is not sufficient to run the E2E flow."""


@dataclass(frozen=True)
class EncounterSection:
    page: Page
    title: str


def _visible(locator: Locator) -> Locator:
    return locator.first


def _section(page: Page, title: str) -> EncounterSection:
    heading = page.get_by_text(re.compile(rf"^\s*{re.escape(title)}\s*$", re.I)).first
    expect(heading).to_be_visible()
    return EncounterSection(page=page, title=title)


def _section_xpath(section: EncounterSection) -> str:
    return (
        "preceding::h2[1]["
        f"translate(normalize-space(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz') = "
        f"'{section.title.lower()}'"
        "]"
    )


def _field(section: EncounterSection, label: str) -> Locator:
    container = _field_container(section, label)
    return container.locator("input, textarea, mat-select, [role='textbox'], [role='combobox']").first


def _field_container(section: EncounterSection, label: str) -> Locator:
    label_node = section.page.locator(
        f"xpath=//*[self::mat-label or self::label][{_section_xpath(section)}]"
    ).filter(has_text=re.compile(rf"^\s*{re.escape(label)}\s*(?:\*)?\s*$", re.I)).first
    return label_node.locator(
        "xpath=ancestor::*[contains(concat(' ', normalize-space(@class), ' '), ' mat-form-field ')][1]"
    )


def _required_value(data: dict[str, dict[str, Any]], section: str, field: str) -> str:
    value = data[section][field]
    if value is None or not str(value).strip():
        raise EncounterDataError(f"Encounter Excel row {ENCOUNTER_ROW_NUMBER} is missing: {section} -> {field}")
    return str(value).strip()


def _optional_value(data: dict[str, dict[str, Any]], section: str, field: str) -> str:
    value = data[section][field]
    return "" if value is None else str(value).strip()


def _validate_required_excel_data(data: dict[str, dict[str, Any]]) -> None:
    missing = []
    for section, fields in REQUIRED_EXCEL_FIELDS.items():
        for field in fields:
            value = data[section][field]
            if value is None or not str(value).strip():
                missing.append(f"{section} -> {field}")

    if missing:
        raise EncounterDataError(
            "Encounter Excel row 3 is not ready for end-to-end creation. "
            "Populate these required values: "
            + "; ".join(missing)
        )

    appointment_value = _required_value(data, "Encounter Details", "Appointment Date and Time")
    if not re.search(r"\b\d{1,2}:\d{2}\s*(?:AM|PM)\b", appointment_value, re.I):
        raise EncounterDataError(
            "Encounter Excel row 3 is missing appointment time: "
            "Encounter Details -> Appointment Date and Time must include both date and time."
        )


def _applicant_name(data: dict[str, dict[str, Any]]) -> str:
    applicant = data["Applicant Details"]
    return " ".join(
        part
        for part in (
            str(applicant.get("First Name") or "").strip(),
            str(applicant.get("Middle Name") or "").strip(),
            str(applicant.get("Last Name") or "").strip(),
        )
        if part
    )


def _applicant_table_name(data: dict[str, dict[str, Any]]) -> str:
    applicant = data["Applicant Details"]
    first_name = str(applicant.get("First Name") or "").strip()
    middle_name = str(applicant.get("Middle Name") or "").strip()
    last_name = str(applicant.get("Last Name") or "").strip()
    given_names = " ".join(part for part in (first_name, middle_name) if part)
    return f"{last_name}, {given_names}" if last_name and given_names else _applicant_name(data)


def _physician_table_name(value: str) -> str:
    parts = [part.strip() for part in value.split(",") if part.strip()]
    if len(parts) >= 2:
        first_part = re.sub(r"\b(?:MD|DO)\b\.?", "", parts[0], flags=re.I).strip()
        second_part = re.sub(r"\b(?:MD|DO)\b\.?", "", parts[1], flags=re.I).strip()
        return " ".join(part for part in (second_part, first_part) if part)
    return re.sub(r"\s+", " ", re.sub(r"\b(?:MD|DO)\b\.?", "", value, flags=re.I)).strip()


def _fill_text(section: Locator, label: str, value: str) -> None:
    if not value:
        return
    try:
        control = _field(section, label)
        expect(control).to_be_visible()
        control.fill(value)
    except Exception as exc:
        raise AssertionError(f"Could not fill text field '{label}' ({type(exc).__name__}).") from None


def _select_dropdown(page: Page, section: Locator, label: str, value: str) -> None:
    if not value:
        return
    try:
        _clear_overlays(page)
        control = _field(section, label)
        expect(control).to_be_visible()
        control.click(force=True)
        if page.locator("mat-option").count() == 0:
            control.focus()
            page.keyboard.press("Enter")
        if page.locator("mat-option").count() == 0:
            control.focus()
            page.keyboard.press("Alt+ArrowDown")
        option = page.locator("mat-option").filter(
            has_text=re.compile(rf"^\s*{re.escape(value)}\s*$", re.I)
        )
        expect(_visible(option)).to_be_visible()
        _visible(option).click()
        expect(_field_container(section, label)).to_contain_text(value)
    except Exception as exc:
        raise AssertionError(f"Could not select dropdown field '{label}' ({type(exc).__name__}).") from None


def _select_state(page: Page, section: Locator, value: str) -> None:
    _select_dropdown(page, section, "State", value)


def _digits_only(value: str) -> str:
    return re.sub(r"\D", "", value)


def _clear_overlays(page: Page) -> None:
    page.keyboard.press("Escape")
    page.keyboard.press("Escape")
    page.mouse.click(5, 5)


def _fill_optional_phone(page: Page, section: Locator, label: str, value: str) -> None:
    if not value:
        return

    phone_digits = _digits_only(value)
    assert phone_digits, f"Encounter phone field '{label}' must contain at least one digit."

    try:
        container = _field_container(section, label)
        dropdown = container.get_by_role("combobox")
        if dropdown.count() and dropdown.first.is_visible():
            dropdown.first.click()
            enter_phone = page.get_by_role("option", name="Enter Phone", exact=True)
            expect(enter_phone).to_be_visible()
            enter_phone.click()

        phone_input = container.get_by_role("textbox").or_(container.get_by_placeholder(re.compile("phone", re.I))).first
        expect(phone_input).to_be_editable()
        phone_input.fill(phone_digits)
        phone_input.blur()
    except Exception as exc:
        raise AssertionError(f"Could not fill phone field '{label}' ({type(exc).__name__}).") from None


def _fill_optional_custom_text(page: Page, section: Locator, label: str, value: str) -> None:
    if not value:
        return
    try:
        container = _field_container(section, label)
        dropdown = container.get_by_role("combobox")
        if dropdown.count() and dropdown.first.is_visible():
            dropdown.first.click()
            custom_text = page.get_by_role("option", name="Enter Custom Text", exact=True)
            expect(custom_text).to_be_visible()
            custom_text.click()

        text_input = container.locator("input, textarea, [role='textbox']").first
        expect(text_input).to_be_editable()
        text_input.fill(value)
        text_input.blur()
    except Exception as exc:
        raise AssertionError(f"Could not fill custom text field '{label}' ({type(exc).__name__}).") from None


def _select_date(page: Page, section: Locator, label: str, mm_dd_yyyy: str) -> None:
    if not mm_dd_yyyy:
        return
    try:
        parsed = datetime.strptime(mm_dd_yyyy.split()[0], "%m/%d/%Y")
        container = _field_container(section, label)
        control = _field(section, label)
        expect(control).to_be_visible()
        control.click()

        month_year = parsed.strftime("%B %Y")
        day = str(parsed.day)
        date_button = page.get_by_role("button", name=re.compile(rf"{re.escape(month_year)}.*\b{day}\b", re.I))
        if date_button.count() == 0:
            control.fill(mm_dd_yyyy)
            control.blur()
        else:
            _visible(date_button).click()

        page.keyboard.press("Escape")
        expect(control).to_have_value(re.compile(r"\d{1,2}[/-]\d{1,2}[/-]\d{4}"))
    except Exception as exc:
        raise AssertionError(f"Could not select date field '{label}' ({type(exc).__name__}).") from None


def _select_body_parts(page: Page, section: Locator, value: str) -> None:
    parts = [part.strip() for part in re.split(r"[;,]", value) if part.strip()]
    if not parts:
        return

    _field(section, "Injured body parts").click()
    dialog = page.get_by_role("dialog").filter(has=page.get_by_role("heading", name=re.compile(r"Body Parts", re.I)))
    expect(_visible(dialog)).to_be_visible()
    modal = _visible(dialog)

    for body_part in parts:
        option = modal.get_by_role("option", name=re.compile(rf"^\s*{re.escape(body_part)}\s*$", re.I))
        expect(_visible(option)).to_be_visible()
        _visible(option).click()

    save_button = modal.get_by_role("button", name=re.compile(r"^Save$", re.I))
    expect(save_button).to_be_enabled()
    save_button.click()
    expect(dialog).not_to_be_visible(timeout=10000)
    page.keyboard.press("Escape")


def _find_upload_file(stem: str) -> Path:
    matches = sorted(path for path in ENCOUNTER_FILES_PATH.glob(f"{stem}.*") if path.is_file())
    assert matches, f"Encounter upload file not found: test_data/encounterfiles/{stem}.<actual-extension>"
    assert len(matches) == 1, f"Multiple Encounter upload files found for {stem}; keep only one supported file."
    return matches[0]


def _upload_documents(page: Page, section: Locator) -> tuple[Path, Path]:
    strike_panel = _find_upload_file("StrikePanel")
    attorney_demo = _find_upload_file("AttorneyDemo")
    file_inputs = page.locator(
        f"xpath=//input[@type='file' and {_section_xpath(section)}]"
    )
    expect(file_inputs).to_have_count(2)
    file_inputs.nth(0).set_input_files(str(strike_panel))
    file_inputs.nth(1).set_input_files(str(attorney_demo))
    return strike_panel, attorney_demo


def _fill_interpreter_details(page: Page, data: dict[str, dict[str, Any]]) -> None:
    """Check Interpreter Required and fill Language, Vendor Name, Vendor Phone."""
    for interp_section, interp_fields in INTERPRETER_REQUIRED_EXCEL_FIELDS.items():
        for interp_field in interp_fields:
            value = data.get(interp_section, {}).get(interp_field)
            if value is None or not str(value).strip():
                raise EncounterDataError(
                    f"Required Excel test data is missing: {interp_section} -> {interp_field}"
                )

    checkbox = page.locator("mat-checkbox").filter(
        has_text=re.compile(r"Interpreter\s+Required", re.I)
    ).first
    if not checkbox.is_visible():
        checkbox = page.get_by_label(re.compile(r"Interpreter\s+Required", re.I)).first
    expect(checkbox).to_be_visible()
    checkbox.click()

    section = _section(page, "Interpreter details")

    language = _required_value(data, "Interpreter Details", "Language")
    _select_searchable_dropdown(page, section, "Language", language)

    vendor_name = _required_value(data, "Interpreter Details", "Vendor Name")
    _fill_text(section, "Vendor Name", vendor_name)

    vendor_phone = _required_value(data, "Interpreter Details", "Vendor Phone")
    _fill_optional_phone(page, section, "Vendor Phone", vendor_phone)


def _select_searchable_dropdown(page: Page, section: Locator, label: str, value: str) -> None:
    if not value:
        return
    try:
        _clear_overlays(page)
        control = _field(section, label)
        expect(control).to_be_visible()
        control.click(force=True)
        if control.evaluate("element => ['INPUT', 'TEXTAREA'].includes(element.tagName)"):
            control.fill(value)
        option = page.locator("mat-option").filter(has_text=re.compile(re.escape(value), re.I))
        expect(_visible(option)).to_be_visible(timeout=15000)
        _visible(option).click()
        expect(_field_container(section, label)).to_contain_text(re.compile(re.escape(value), re.I))
    except Exception as exc:
        raise AssertionError(f"Could not select searchable field '{label}' ({type(exc).__name__}).") from None


def _select_appointment_date_time(page: Page, section: Locator, value: str) -> None:
    if not value:
        return
    _fill_text(section, "Appointment date & time", value)


def _format_ui_date(value: str) -> str:
    digits = re.findall(r"\d+", value)
    if len(digits) >= 3:
        month, day, year = digits[0], digits[1], digits[2]
        return f"{int(month):02d}/{int(day):02d}/{int(year):04d}"
    return value.strip()


def _contains_display_text(actual_text: str, expected_text: str) -> bool:
    actual = re.sub(r"\s+", " ", actual_text).strip().casefold()
    expected = re.sub(r"\s+", " ", expected_text).strip().casefold()
    return expected in actual


def _handle_generated_document_tab(context: BrowserContext, app_page: Page) -> None:
    for candidate in context.pages:
        if candidate == app_page or candidate.is_closed():
            continue
        url = candidate.url.lower()
        if "pdf" in url or "anf" in url or "blob:" in url or "drive.google.com" in url:
            candidate.close()
    app_page.bring_to_front()


def _fill_applicant_details(page: Page, data: dict[str, dict[str, Any]]) -> None:
    section = _section(page, "Applicant details")
    applicant = data["Applicant Details"]
    _fill_text(section, "First name", _required_value(data, "Applicant Details", "First Name"))
    _fill_text(section, "Middle name", _optional_value(data, "Applicant Details", "Middle Name"))
    _fill_text(section, "Last name", _required_value(data, "Applicant Details", "Last Name"))
    _select_date(page, section, "Date of Birth", _required_value(data, "Applicant Details", "Date of Birth"))
    _select_dropdown(page, section, "Gender", _required_value(data, "Applicant Details", "Gender"))
    _fill_text(section, "SSN", _required_value(data, "Applicant Details", "SSN"))
    _fill_text(section, "Address line 1", _required_value(data, "Applicant Details", "Address Line 1"))
    _fill_text(section, "Address line 2", _optional_value(data, "Applicant Details", "Address Line 2"))
    _fill_text(section, "City", _required_value(data, "Applicant Details", "City"))
    _select_state(page, section, _required_value(data, "Applicant Details", "State"))
    _fill_text(section, "Zip", _required_value(data, "Applicant Details", "Zip"))
    _fill_optional_phone(page, section, "Home Phone", str(applicant.get("Home Phone") or ""))
    _fill_optional_phone(page, section, "Cell phone", str(applicant.get("Cell Phone") or ""))
    _fill_optional_phone(page, section, "Work phone", str(applicant.get("Work Phone") or ""))
    _fill_text(section, "Email", _optional_value(data, "Applicant Details", "Email"))
    _fill_text(section, "Job title", _optional_value(data, "Applicant Details", "Job Title"))


def _fill_injury_details(page: Page, data: dict[str, dict[str, Any]]) -> None:
    section = _section(page, "Injury details")
    _select_date(page, section, "Date of injury", _required_value(data, "Injury Details", "Date of Injury"))
    _select_body_parts(page, section, _required_value(data, "Injury Details", "Injured Body Parts"))
    _fill_optional_custom_text(page, section, "WCAB Case Number", _required_value(data, "Injury Details", "WCAB Number"))
    _fill_optional_custom_text(page, section, "Claim Number", _required_value(data, "Injury Details", "Claim Number"))


def _fill_encounter_details(page: Page, data: dict[str, dict[str, Any]]) -> None:
    section = _section(page, "Encounter details")
    _select_searchable_dropdown(page, section, "Physician", _required_value(data, "Encounter Details", "Physician"))
    _select_dropdown(page, section, "Medical Evaluation Type", _required_value(data, "Encounter Details", "Medical Evaluation Type"))
    _select_dropdown(page, section, "Encounter type", _required_value(data, "Encounter Details", "Encounter Type"))
    _fill_text(section, "Panel Number", _optional_value(data, "Encounter Details", "Panel Number"))
    _select_date(page, section, "Date of Appointment Call", _required_value(data, "Encounter Details", "Date of Appointment Call"))
    _select_appointment_date_time(page, section, _required_value(data, "Encounter Details", "Appointment Date and Time"))
    _select_dropdown(page, section, "Requesting party", _required_value(data, "Encounter Details", "Requesting Party"))
    _select_searchable_dropdown(page, section, "Examination address", _required_value(data, "Encounter Details", "Examination Address"))

    try:
        city = _field(section, "City")
        zip_code = _field(section, "Zip")
        expect(city).to_have_value(re.compile(r"\S+"))
        expect(zip_code).to_have_value(re.compile(r"\S+"))
    except Exception as exc:
        raise AssertionError(f"Dependent Examination City/Zip did not populate ({type(exc).__name__}).") from None

    evaluation_type = _optional_value(data, "Encounter Details", "Evaluation Type")
    if evaluation_type and _field(section, "Evaluation type").is_enabled():
        _select_dropdown(page, section, "Evaluation type", evaluation_type)


def _fill_party_details(page: Page, data: dict[str, dict[str, Any]], section_name: str) -> None:
    section = _section(page, section_name)
    excel_section = " ".join(word.capitalize() for word in section_name.split())
    if excel_section == "Applicant Attorney Details":
        _fill_text(section, "Firm", _required_value(data, excel_section, "Firm"))
    elif excel_section == "Defense Attorney Details":
        _fill_text(section, "Firm", _required_value(data, excel_section, "Firm"))
    else:
        _fill_text(section, "Employer", _required_value(data, excel_section, "Employer"))

    for excel_field, ui_label in (
        ("First Name", "First name"),
        ("Middle Name", "Middle name"),
        ("Last Name", "Last name"),
        ("Address Line 1", "Address line 1"),
        ("Address Line 2", "Address line 2"),
        ("City", "City"),
        ("Zip", "Zip"),
        ("Fax", "Fax"),
        ("Email", "Email"),
        ("Assistant Email", "Assistant email"),
    ):
        if excel_field in data[excel_section]:
            _fill_text(section, ui_label, _optional_value(data, excel_section, excel_field))
    if "State" in data[excel_section]:
        _select_state(page, section, _optional_value(data, excel_section, "State"))
    if "Phone" in data[excel_section]:
        _fill_optional_phone(page, section, "Phone", _optional_value(data, excel_section, "Phone"))


def _fill_insurance_details(page: Page, data: dict[str, dict[str, Any]]) -> None:
    section = _section(page, "Insurance company details")
    excel_section = "Insurance Company Details"
    for excel_field, ui_label in (
        ("Insurance Company", "Insurance company"),
        ("Payer ID", "Payer ID"),
        ("Adjuster First Name", "Adjuster first name"),
        ("Adjuster Middle Name", "Adjuster middle name"),
        ("Adjuster Last Name", "Adjuster last name"),
        ("Address Line 1", "Address line 1"),
        ("Address Line 2", "Address line 2"),
        ("City", "City"),
        ("Zip", "Zip"),
        ("Adjuster Fax", "Adjuster fax"),
        ("Adjuster Email", "Adjuster email"),
        ("Adjuster Assistant Email", "Adjuster assistant email"),
    ):
        _fill_text(section, ui_label, _optional_value(data, excel_section, excel_field))
    _select_state(page, section, _required_value(data, excel_section, "State"))
    _fill_optional_phone(page, section, "Insurance company phone", _required_value(data, excel_section, "Insurance Company Phone"))
    _fill_optional_phone(page, section, "Adjuster phone", _optional_value(data, excel_section, "Adjuster Phone"))


def _save_encounter(page: Page) -> None:
    save_button = page.get_by_role("button", name=re.compile(r"Save encounter", re.I)).last
    expect(save_button).to_be_enabled()
    save_button.click()


def _complete_recipients_popup(page: Page) -> None:
    dialog = page.get_by_role("dialog").or_(page.locator("mat-dialog-container")).filter(
        has_text=re.compile(r"recipient|send anf", re.I)
    )
    expect(_visible(dialog)).to_be_visible(timeout=30000)
    save_button = _visible(dialog).get_by_role(
        "button",
        name=re.compile(r"Save encounter|Download PDFs and save encounter", re.I),
    )
    expect(save_button).to_be_enabled()
    save_button.click()
    _confirm_post_save_popup(page)
    expect(_visible(dialog)).not_to_be_visible(timeout=60000)


def _confirm_post_save_popup(page: Page) -> None:
    confirm_button = page.get_by_role("button", name=re.compile(r"^Confirm$", re.I))
    try:
        expect(confirm_button.first).to_be_visible(timeout=30000)
    except Exception:
        return

    confirm_container = page.get_by_role("dialog").or_(page.locator("mat-dialog-container")).filter(
        has=confirm_button.first
    )
    if confirm_container.count() and confirm_container.first.is_visible():
        confirm_button = confirm_container.first.get_by_role("button", name=re.compile(r"^Confirm$", re.I))

    expect(confirm_button.first).to_be_enabled()
    confirm_button.first.click()
    if confirm_container.count():
        expect(confirm_container.first).not_to_be_visible(timeout=30000)


def _open_view_encounters(page: Page) -> None:
    view_encounters = page.get_by_role("link", name=re.compile(r"View Encounters", re.I)).or_(
        page.get_by_text("View Encounters", exact=True)
    )
    expect(_visible(view_encounters)).to_be_visible()
    _visible(view_encounters).click()
    page.wait_for_url(re.compile(r".*/#/encounter_details$"))
    expect(page.get_by_role("button", name=re.compile(r"New encounter", re.I))).to_be_visible()


def _open_new_encounter(page: Page) -> None:
    new_encounter = page.get_by_role("button", name=re.compile(r"New encounter", re.I))
    expect(new_encounter).to_be_visible()
    new_encounter.click(force=True)
    page.wait_for_url(re.compile(r".*/#/calendar/create_encounter$"))
    expect(page.get_by_text(re.compile(r"^\s*Applicant details\s*$", re.I))).to_be_visible()


def _parse_ui_date(value: str) -> date:
    """Parse a UI date string (MM/DD/YYYY) into a date object."""
    normalized = _format_ui_date(value)
    return datetime.strptime(normalized, "%m/%d/%Y").date()


def _extract_row_cell_dates(row_locator: Locator, expected_eval_date_str: str) -> tuple[str, str]:
    """Extract Date of Evaluation and Report Due Date text from the matched row."""
    cells = row_locator.locator("td")
    date_pattern = re.compile(r"\b\d{1,2}/\d{1,2}/\d{4}\b")
    candidate_dates: list[str] = []

    for index in range(cells.count()):
        candidate_dates.extend(date_pattern.findall(cells.nth(index).inner_text().strip()))

    expected_eval_normalized = _format_ui_date(expected_eval_date_str)
    date_of_evaluation = ""
    report_due_date = ""

    for candidate in candidate_dates:
        if _format_ui_date(candidate) == expected_eval_normalized:
            date_of_evaluation = candidate
            break

    for candidate in candidate_dates:
        candidate_normalized = _format_ui_date(candidate)
        if candidate_normalized == expected_eval_normalized:
            continue
        try:
            candidate_date = datetime.strptime(candidate_normalized, "%m/%d/%Y").date()
            eval_date = datetime.strptime(expected_eval_normalized, "%m/%d/%Y").date()
            if 25 <= (candidate_date - eval_date).days <= 35:
                report_due_date = candidate
                break
        except ValueError:
            continue

    return date_of_evaluation, report_due_date


def _validate_encounter_row(page: Page, data: dict[str, dict[str, Any]]) -> dict[str, str]:
    _open_view_encounters(page)
    applicant_name = _applicant_name(data)
    applicant_table_name = _applicant_table_name(data)
    physician = _required_value(data, "Encounter Details", "Physician")
    physician_table_name = _physician_table_name(physician)
    encounter_type = _required_value(data, "Encounter Details", "Encounter Type")
    evaluation_type = _optional_value(data, "Encounter Details", "Evaluation Type")
    appointment_value = _required_value(data, "Encounter Details", "Appointment Date and Time")
    appointment_date = _format_ui_date(appointment_value)
    evaluation_date_obj = _parse_ui_date(appointment_value)
    expected_report_due_date_obj = calculate_report_due_date(evaluation_date_obj)
    base_plus_30 = evaluation_date_obj + timedelta(days=30)
    expected_eval_date_str = evaluation_date_obj.strftime("%m/%d/%Y")
    expected_due_date_str = expected_report_due_date_obj.strftime("%m/%d/%Y")

    search = page.get_by_placeholder(re.compile("Applicant name", re.I)).or_(
        page.get_by_role("textbox", name=re.compile("Applicant name", re.I))
    )
    if search.count():
        search.first.fill(applicant_table_name)

    row = page.get_by_role("row").filter(has_text=applicant_table_name).filter(has_text=encounter_type).filter(
        has_text=appointment_date
    )
    expect(_visible(row)).to_be_visible(timeout=30000)
    matched_row = _visible(row)

    row_text = matched_row.inner_text()
    assert _contains_display_text(row_text, applicant_table_name), (
        f"Applicant Name mismatch. Expected '{applicant_table_name}', "
        f"but the row text does not contain it. Row text: {row_text!r}"
    )
    assert _contains_display_text(row_text, physician_table_name), (
        f"Physician mismatch. Expected '{physician_table_name}', "
        f"but the row text does not contain it. Row text: {row_text!r}"
    )
    assert _contains_display_text(row_text, encounter_type), (
        f"Encounter Type mismatch. Expected '{encounter_type}', "
        f"but the row text does not contain it."
    )
    if evaluation_type:
        assert _contains_display_text(row_text, evaluation_type), (
            f"Evaluation Type mismatch. Expected '{evaluation_type}', "
            f"but the row text does not contain it."
        )

    actual_eval_date_str, actual_due_date_str = _extract_row_cell_dates(matched_row, appointment_value)
    if actual_eval_date_str:
        actual_eval_normalized = _format_ui_date(actual_eval_date_str)
        assert actual_eval_normalized == expected_eval_date_str, (
            f"Date of Evaluation mismatch. "
            f"Expected {expected_eval_date_str}, found {actual_eval_normalized}."
        )
    else:
        assert appointment_date in row_text, (
            f"Date of Evaluation not found in row. "
            f"Expected {expected_eval_date_str} to appear in the row text."
        )

    if actual_due_date_str:
        actual_due_normalized = _format_ui_date(actual_due_date_str)
        actual_due_date_obj = datetime.strptime(actual_due_normalized, "%m/%d/%Y").date()
        assert actual_due_date_obj == expected_report_due_date_obj, (
            f"Report Due Date mismatch. "
            f"Evaluation Date: {expected_eval_date_str}, "
            f"Base +30 Date: {base_plus_30.strftime('%m/%d/%Y')} ({base_plus_30.strftime('%A')}), "
            f"Expected adjusted Due Date: {expected_due_date_str}, "
            f"Actual: {actual_due_normalized}."
        )
    else:
        assert expected_due_date_str in row_text, (
            f"Report Due Date not found in row. "
            f"Expected {expected_due_date_str} "
            f"(Evaluation Date {expected_eval_date_str} + 30 days, adjusted). "
            f"Row text: {row_text!r}"
        )

    return {
        "Applicant Name": applicant_name,
        "Applicant Name Actual": applicant_table_name,
        "Physician": physician,
        "Physician Actual": physician_table_name,
        "Encounter Type": encounter_type,
        "Evaluation Type": evaluation_type,
        "Appointment Date": appointment_date,
        "Date of Evaluation": actual_eval_date_str or appointment_date,
        "Expected Report Due Date": expected_due_date_str,
        "Actual Report Due Date": actual_due_date_str or "(matched in row text)",
    }


@report_case(
    test_case_id="SUPERADMIN_PQME_ENCOUNTER_TC_002",
    module="Super Admin - PQME Encounter Creation",
    test_case_name="Create PQME Encounter With Interpreter and Validate Encounter",
    title=(
        "Verify that if the Super Admin creates a PQME encounter with an interpreter "
        "using valid Excel test data, then the encounter is created successfully with "
        "the interpreter details and the correct encounter information is displayed."
    ),
    test_type="Positive",
    preconditions=(
        "Valid Super Admin credentials must be configured, valid PQME Encounter test data "
        "including Interpreter Details must be available in the PQME Encounter Excel workbook, "
        "and all required master data and upload files must exist."
    ),
    steps=(
        "1. Read Encounter data from the two-level Section -> Field Excel workbook.\n"
        "2. Validate required Encounter Excel data, Interpreter Details, and upload files.\n"
        "3. Login as Super Admin using existing environment credentials and verification helper.\n"
        "4. Navigate to View Encounters.\n"
        "5. Open New Encounter.\n"
        "6. Fill Applicant Details.\n"
        "7. Fill Injury Details and select requested body parts.\n"
        "8. Fill Encounter Details and dependent Examination Address.\n"
        "9. Upload Strike Panel and Attorney Demo files.\n"
        "10. Check Interpreter Required and fill Language, Vendor Name, and Vendor Phone.\n"
        "11. Fill Applicant Attorney, Employer, Defense Attorney, and Insurance Company details.\n"
        "12. Save the Encounter and complete the Recipients popup.\n"
        "13. Handle the generated ANF/PDF tab and return to OrthoLegal.\n"
        "14. Search for the created Encounter in View Encounters by Applicant name.\n"
        "15. Validate the Encounter row: Applicant Name, Physician, Encounter Type, "
        "Date of Evaluation, and Report Due Date (Evaluation Date + 30 calendar days, "
        "adjusted backward for weekends and U.S. federal holidays)."
    ),
    test_data=(
        "PQME Encounter data row 3 from test_data/PQME_Encounter_Test_Data.xlsx, "
        "worksheet PQME_Encounter_Test_Data, including Interpreter Details "
        "(Language, Vendor Name, Vendor Phone), and upload files from "
        "test_data/encounterfiles. Super Admin password and verification PIN are not reported."
    ),
    expected_result=(
        "The Super Admin should be able to create a PQME Encounter with the Interpreter "
        "Required option selected and valid Language, Vendor Name, and Vendor Phone values "
        "from the PQME Encounter Excel test-data sheet. After completing the form, uploading "
        "the required documents, and saving the PQME Encounter through the Recipients step, "
        "the Encounter should be created successfully and should appear in View Encounters "
        "with the correct Applicant name, Physician, Encounter Type, Evaluation Type, Date "
        "of Evaluation (matching the Appointment Date), and Report Due Date (Evaluation "
        "Date + 30 calendar days, adjusted backward for weekends and U.S. federal holidays)."
    ),
    actual_result=(
        "The Super Admin created the PQME Encounter with Interpreter Details from Excel data, "
        "completed the Recipients step, returned to View Encounters, and validated the "
        "created row against the expected Applicant name, Physician, Encounter Type, "
        "Evaluation Type, Date of Evaluation, and Report Due Date."
    ),
)
def test_create_pqme_encounter_with_interpreter_and_validate_encounter(page: Page) -> None:
    stage = "reading Encounter Excel data"
    try:
        encounter_data = read_encounter_test_data(
            row_number=ENCOUNTER_ROW_NUMBER,
            workbook_path=PQME_ENCOUNTER_WORKBOOK_PATH,
            sheet_name=PQME_ENCOUNTER_SHEET_NAME,
        )
        _validate_required_excel_data(encounter_data)
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

        stage = "filling Encounter Details"
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

        stage = "saving the Encounter"
        app_page = page
        _save_encounter(page)
        _complete_recipients_popup(page)
        _handle_generated_document_tab(page.context, app_page)

        stage = "validating the created Encounter row"
        actual_values = _validate_encounter_row(app_page, encounter_data)
        interpreter = encounter_data.get("Interpreter Details", {})
        print(
            "Created PQME Encounter (With Interpreter) validated: "
            f"Applicant={actual_values['Applicant Name']}; "
            f"Physician={actual_values['Physician']}; "
            f"Encounter Type={actual_values['Encounter Type']}; "
            f"Evaluation Type={actual_values['Evaluation Type']}; "
            f"Appointment Date={actual_values['Appointment Date']}; "
            f"Date of Evaluation={actual_values['Date of Evaluation']}; "
            f"Expected Report Due Date={actual_values['Expected Report Due Date']}; "
            f"Actual Report Due Date={actual_values['Actual Report Due Date']}; "
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
            f"Create PQME Encounter With Interpreter and Validate Encounter "
            f"failed while {stage} ({type(exc).__name__})."
        ) from None
