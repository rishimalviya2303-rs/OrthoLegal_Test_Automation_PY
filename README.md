# OrthoLegal Automation Framework

## 1. Project Overview

This repository contains a standalone Playwright + Python end-to-end automation framework for the OrthoLegal web application. It is maintained separately from the application source code and drives the browser like a real user.

The current implementation covers Super Admin login, User Creation workflows, Encounter Creation workflows, Excel-driven test data, dynamic login verification, and timestamped Excel execution reporting.

## 2. Tech Stack

| Technology | Purpose |
| --- | --- |
| Python | Automation language |
| Playwright | Browser automation |
| pytest | Test execution |
| pytest-playwright | pytest integration for Playwright fixtures |
| python-dotenv | `.env` configuration loading |
| openpyxl | Excel test-data reading and execution-report generation |

## 3. Project Structure

```text
OrthoLegal_Automation/
|-- conftest.py
|-- README.md
|-- requirements.txt
|-- reporting/
|   |-- __init__.py
|   `-- excel_report.py
|-- reports/
|   `-- OrthoLegal_Automation_Test_Report_<timestamp>_IST.xlsx
|-- test_data/
|   |-- PQME_Encounter_Test_Data.xlsx
|   |-- PQMER_Encounter_Test_Data.xlsx
|   |-- PQMES_Encounter_Test_Data.xlsx
|   |-- PQMED_Encounter_Test_Data.xlsx
|   |-- Reflex_PQMES_Encounter_Test_Data.xlsx
|   |-- Users_Test_Data.xlsx
|   |-- encounterfiles/
|   |   |-- AttorneyDemo.docx
|   |   `-- StrikePanel.docx
|   `-- signatures/
|       |-- CaseManagerSignature.png
|       |-- CaseManagetSignature.png
|       `-- ReportFinalizerSignature.png
|-- tests/
|   |-- test_case_manager_phone_update.py
|   |-- test_encounter_test_data.py
|   |-- test_excel_reader.py
|   |-- test_super_admin_login.py
|   |-- Encounter Creation/
|   |   |-- test_pqme_encounter_creation.py
|   |   |-- test_pqmer_encounter_creation.py
|   |   |-- test_pqmes_encounter_creation.py
|   |   |-- test_pqmed_encounter_creation.py
|   |   `-- test_reflex_pqmes_encounter_creation.py
|   `-- User Creation/
|       |-- test_case_manager_creation.py
|       |-- test_editor_creation.py
|       |-- test_editor_manager_creation.py
|       |-- test_historian_creation.py
|       |-- test_historian_manager_creation.py
|       |-- test_medical_assistant_creation.py
|       |-- test_medical_assistant_manager_creation.py
|       |-- test_record_reviewer_creation.py
|       |-- test_record_reviewer_manager_creation.py
|       `-- test_report_finalizer_creation.py
`-- utils/
    |-- __init__.py
    |-- auth.py
    |-- encounter_test_data.py
    |-- excel_reader.py
    `-- profile.py
```

Generated cache folders such as `.pytest_cache/`, `__pycache__/`, and `.venv/` may exist locally but are not part of the automation design.

## 4. Environment Setup

Create a local `.env` file in the project root. The current authentication helper reads these variable names:

```env
BASE_URL=
TEST_EMAIL=
TEST_PASSWORD=
```

`BASE_URL` is the OrthoLegal application URL. `TEST_EMAIL` and `TEST_PASSWORD` are the Super Admin login credentials used by `utils/auth.py`.

Never place real credentials in `README.md`, execution reports, committed code, or shared test data. The local `.env` file should remain private.

## 5. Installation

From PowerShell:

```powershell
cd "C:\Rishi Files\OrthoLegal_Automation"
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python -m playwright install
```

Use the project virtual environment for test runs:

```powershell
.\.venv\Scripts\python.exe -m pytest --version
```

## 6. Authentication and Dynamic Verification

Login automation is centralized in `utils/auth.py`.

The current flow is:

1. Open `BASE_URL`.
2. Enter the email and password from `.env`.
3. Click `Sign in`.
4. Wait for the verification-code screen.
5. Read the application message containing `Login PIN sent to your email.`.
6. Extract the displayed six-digit verification PIN from the page.
7. Enter the extracted PIN into `Enter security pin code`.
8. Submit the verification code.
9. Validate the authenticated dashboard URL ending in `#/dashboard`.

The six-digit verification PIN is dynamic and must not be hardcoded. Credentials, PINs, tokens, cookies, and other secrets must not be written to reports.

## 7. Test Data Architecture

### User Test Data

User Creation tests read:

```text
test_data/Users_Test_Data.xlsx
```

Worksheet:

```text
Users
```

Current columns include:

```text
Select Role
First Name
Middle Name
Last Name
Ortho Legal Email
Personal Email
Password
Confirm Password
Updated Phone Number
Updated Office Phone Number
```

The User Creation tests select the correct row by matching the `Select Role` value.

Current `Select Role` values in the workbook:

| Role |
| --- |
| Case Manager |
| Historian |
| Historian Manager |
| Record Reviewer |
| Record Reviewer Manager |
| Medical Assistant |
| Medical Assistant Manager |
| Editor |
| Editor Manager |
| Report Finalizer |

Credentials and user input values should come from Excel and `.env`, not from hardcoded test-script values.

### Signature Files

The current signature folder is:

```text
test_data/signatures/
```

Current signature files:

| File | Current usage |
| --- | --- |
| `CaseManagetSignature.png` | Used by `tests\User Creation\test_case_manager_creation.py` |
| `ReportFinalizerSignature.png` | Used by `tests\User Creation\test_report_finalizer_creation.py` |
| `CaseManagerSignature.png` | Present in the folder, but the current Case Manager test uses `CaseManagetSignature.png` |

Signature upload currently applies to Case Manager and Report Finalizer creation tests only.

### Encounter Upload Files

The current encounter upload folder is:

```text
test_data/encounterfiles/
```

Current upload files:

| File |
| --- |
| `AttorneyDemo.docx` |
| `StrikePanel.docx` |

The PQME Encounter Creation test uploads these files through Playwright `set_input_files()`, avoiding the Windows native file dialog.

## 8. Encounter Test Data Workbooks

Encounter data is read with the generic helper in `utils/encounter_test_data.py`. The helper reads a two-level Excel header design:

```text
Section -> Field
```

Examples:

```text
Encounter Details -> Physician
Applicant Details -> First Name
Upload Documents -> Upload Strike Panel
```

Current workbooks:

| Encounter Type | Workbook | Worksheet |
| --- | --- | --- |
| PQME | `PQME_Encounter_Test_Data.xlsx` | `Encounter_Test_Data` |
| PQMER | `PQMER_Encounter_Test_Data.xlsx` | `Encounter_Test_Data` |
| PQMES | `PQMES_Encounter_Test_Data.xlsx` | `PQMES_Encounter_Test_Data` |
| PQMED | `PQMED_Encounter_Test_Data.xlsx` | `PQMED_Encounter_Test_Data` |
| Reflex PQMES | `Reflex_PQMES_Encounter_Test_Data.xlsx` | `Reflex_PQMES_Encounter_Test_Dat` |

Common workbook sections:

| Section | Fields |
| --- | --- |
| Applicant Details | First Name, Middle Name, Last Name, Date of Birth, Gender, SSN, Address Line 1, Address Line 2, City, State, Zip, Home Phone, Cell Phone, Work Phone, Email, Job Title |
| Injury Details | Date of Injury, Injured Body Parts, WCAB Number, Claim Number |
| Upload Documents | Upload Strike Panel, Upload Attorney Demos |
| Applicant Attorney Details | Firm, First Name, Middle Name, Last Name, Address Line 1, Address Line 2, City, State, Zip, Phone, Fax, Email, Assistant Email |
| Employer Details | Employer, Address Line 1, Address Line 2, City, State, Zip, Phone, Fax, Email |
| Defense Attorney Details | Firm, First Name, Middle Name, Last Name, Address Line 1, Address Line 2, City, State, Zip, Phone, Fax, Email, Assistant Email |
| Insurance Company Details | Insurance Company, Payer ID, Insurance Company Phone, Adjuster First Name, Adjuster Middle Name, Adjuster Last Name, Address Line 1, Address Line 2, City, State, Zip, Adjuster Phone, Adjuster Fax, Adjuster Email, Adjuster Assistant Email |

The `Encounter Details` section changes by encounter type.

### PQME Encounter Details

| Order | Field |
| --- | --- |
| 1 | Physician |
| 2 | Medical Evaluation Type |
| 3 | Encounter Type |
| 4 | Panel Number |
| 5 | Date of Appointment Call |
| 6 | Appointment Date and Time |
| 7 | Requesting Party |
| 8 | Examination Address |
| 9 | City |
| 10 | State |
| 11 | Zip |
| 12 | Evaluation Type |

### PQMER Encounter Details

| Order | Field |
| --- | --- |
| 1 | Physician |
| 2 | Medical Evaluation Type |
| 3 | Encounter Type |
| 4 | Panel Number |
| 5 | Date of Appointment Call |
| 6 | Appointment Date and Time |
| 7 | Requesting Party |
| 8 | Examination Address |
| 9 | City |
| 10 | State |
| 11 | Zip |
| 12 | Evaluation Type |

### PQMES Encounter Details

| Order | Field |
| --- | --- |
| 1 | Physician |
| 2 | Medical Evaluation Type |
| 3 | Encounter Type |
| 4 | Panel Number |
| 5 | Date of Cover Letter |
| 6 | Date of Declaration |
| 7 | Requesting Party |

### PQMED Encounter Details

| Order | Field |
| --- | --- |
| 1 | Physician |
| 2 | Medical Evaluation Type |
| 3 | Encounter Type |
| 4 | Panel Number |
| 5 | Date of Appointment Call |
| 6 | Appointment Date and Time |
| 7 | Requesting Party |
| 8 | Address Line 1 |
| 9 | Address Line 2 |
| 10 | City |
| 11 | State |
| 12 | Zip |

PQMED uses direct Encounter address fields instead of the PQME/PQMER `Examination Address` behavior.

### Reflex PQMES Encounter Details

| Order | Field |
| --- | --- |
| 1 | Physician |
| 2 | Medical Evaluation Type |
| 3 | Encounter Type |
| 4 | Panel Number |
| 5 | Date of Receipt |
| 6 | Requesting Party |

Normal PQMES uses `Date of Cover Letter` and `Date of Declaration`. Reflex PQMES uses `Date of Receipt`.

## 9. User Creation Tests

Current User Creation tests are located in:

```text
tests\User Creation\
```

| Role | Test File | Data Source | Notes |
| --- | --- | --- | --- |
| Case Manager | `test_case_manager_creation.py` | `Users_Test_Data.xlsx` | Uploads `CaseManagetSignature.png` |
| Editor | `test_editor_creation.py` | `Users_Test_Data.xlsx` | Creates user, validates search result, then logs in as created user |
| Editor Manager | `test_editor_manager_creation.py` | `Users_Test_Data.xlsx` | Creates user, validates search result, then logs in as created user |
| Historian | `test_historian_creation.py` | `Users_Test_Data.xlsx` | Creates user, validates search result, then logs in as created user |
| Historian Manager | `test_historian_manager_creation.py` | `Users_Test_Data.xlsx` | Creates user, validates search result, then logs in as created user |
| Medical Assistant | `test_medical_assistant_creation.py` | `Users_Test_Data.xlsx` | Leaves the Medical Assistant-specific 4-digit PIN/Create PIN field untouched |
| Medical Assistant Manager | `test_medical_assistant_manager_creation.py` | `Users_Test_Data.xlsx` | Leaves the Medical Assistant-specific 4-digit PIN/Create PIN field untouched |
| Record Reviewer | `test_record_reviewer_creation.py` | `Users_Test_Data.xlsx` | Creates user, validates search result, then logs in as created user |
| Record Reviewer Manager | `test_record_reviewer_manager_creation.py` | `Users_Test_Data.xlsx` | Reuses Record Reviewer helpers for the manager role workflow |
| Report Finalizer | `test_report_finalizer_creation.py` | `Users_Test_Data.xlsx` | Uploads `ReportFinalizerSignature.png` |

The Medical Assistant role-specific 4-digit workflow PIN is separate from the dynamic six-digit login verification PIN and is intentionally left untouched by the current Medical Assistant tests.

## 10. Encounter Creation Tests

Current Encounter Creation tests are located in:

```text
tests\Encounter Creation\
```

| Encounter Type | Test File | Data Source | Notes |
| --- | --- | --- | --- |
| PQME | `test_pqme_encounter_creation.py` | `PQME_Encounter_Test_Data.xlsx` | Uses appointment call/date-time fields, Examination Address, Evaluation Type, and document uploads |
| PQMER | `test_pqmer_encounter_creation.py` | `PQMER_Encounter_Test_Data.xlsx` | Uses the same Encounter Details field structure as PQME |
| PQMES | `test_pqmes_encounter_creation.py` | `PQMES_Encounter_Test_Data.xlsx` | Uses Date of Cover Letter and Date of Declaration |
| PQMED | `test_pqmed_encounter_creation.py` | `PQMED_Encounter_Test_Data.xlsx` | Uses appointment date/time and direct Encounter address fields |
| Reflex PQMES | `test_reflex_pqmes_encounter_creation.py` | `Reflex_PQMES_Encounter_Test_Data.xlsx` | Uses Date of Receipt instead of PQMES cover-letter/declaration dates |

## 11. Standalone Tests

| Test File | Purpose |
| --- | --- |
| `tests\test_super_admin_login.py` | Validates Super Admin login using `.env` credentials and the displayed six-digit verification PIN |
| `tests\test_case_manager_phone_update.py` | Updates and validates Case Manager phone numbers from `Users_Test_Data.xlsx` |
| `tests\test_excel_reader.py` | Unit coverage for Excel reader filtering behavior |
| `tests\test_encounter_test_data.py` | Unit coverage for Encounter workbook generation, two-level headers, grouping, and date/date-time formatting |

## 12. Running Tests

Run commands should be executed from the project root:

```powershell
cd "C:\Rishi Files\OrthoLegal_Automation"
```

### Run All Tests

```powershell
.\.venv\Scripts\python.exe -m pytest --headed -v
```

### Run All User Creation Tests

```powershell
.\.venv\Scripts\python.exe -m pytest "tests\User Creation" --headed -v
```

### Run Individual User Creation Tests

| Role | Command |
| --- | --- |
| Case Manager | `.\.venv\Scripts\python.exe -m pytest "tests\User Creation\test_case_manager_creation.py" --headed -v` |
| Editor | `.\.venv\Scripts\python.exe -m pytest "tests\User Creation\test_editor_creation.py" --headed -v` |
| Editor Manager | `.\.venv\Scripts\python.exe -m pytest "tests\User Creation\test_editor_manager_creation.py" --headed -v` |
| Historian | `.\.venv\Scripts\python.exe -m pytest "tests\User Creation\test_historian_creation.py" --headed -v` |
| Historian Manager | `.\.venv\Scripts\python.exe -m pytest "tests\User Creation\test_historian_manager_creation.py" --headed -v` |
| Medical Assistant | `.\.venv\Scripts\python.exe -m pytest "tests\User Creation\test_medical_assistant_creation.py" --headed -v` |
| Medical Assistant Manager | `.\.venv\Scripts\python.exe -m pytest "tests\User Creation\test_medical_assistant_manager_creation.py" --headed -v` |
| Record Reviewer | `.\.venv\Scripts\python.exe -m pytest "tests\User Creation\test_record_reviewer_creation.py" --headed -v` |
| Record Reviewer Manager | `.\.venv\Scripts\python.exe -m pytest "tests\User Creation\test_record_reviewer_manager_creation.py" --headed -v` |
| Report Finalizer | `.\.venv\Scripts\python.exe -m pytest "tests\User Creation\test_report_finalizer_creation.py" --headed -v` |

### Run All Encounter Creation Tests

```powershell
.\.venv\Scripts\python.exe -m pytest "tests\Encounter Creation" --headed -v
```

### Run Individual Encounter Creation Tests

| Encounter Type | Command |
| --- | --- |
| PQME | `.\.venv\Scripts\python.exe -m pytest "tests\Encounter Creation\test_pqme_encounter_creation.py" --headed -v` |
| PQMER | `.\.venv\Scripts\python.exe -m pytest "tests\Encounter Creation\test_pqmer_encounter_creation.py" --headed -v` |
| PQMES | `.\.venv\Scripts\python.exe -m pytest "tests\Encounter Creation\test_pqmes_encounter_creation.py" --headed -v` |
| PQMED | `.\.venv\Scripts\python.exe -m pytest "tests\Encounter Creation\test_pqmed_encounter_creation.py" --headed -v` |
| Reflex PQMES | `.\.venv\Scripts\python.exe -m pytest "tests\Encounter Creation\test_reflex_pqmes_encounter_creation.py" --headed -v` |

### Run Standalone Tests

| Test | Command |
| --- | --- |
| Super Admin Login | `.\.venv\Scripts\python.exe -m pytest tests\test_super_admin_login.py --headed -v` |
| Case Manager Phone Update | `.\.venv\Scripts\python.exe -m pytest tests\test_case_manager_phone_update.py --headed -v` |
| Excel Reader Unit Tests | `.\.venv\Scripts\python.exe -m pytest tests\test_excel_reader.py -v` |
| Encounter Test Data Unit Tests | `.\.venv\Scripts\python.exe -m pytest tests\test_encounter_test_data.py -v` |

## 13. Execution Reporting

Execution reporting is implemented in `reporting/excel_report.py` and connected through pytest hooks in `conftest.py`.

Current behavior:

| Behavior | Implementation |
| --- | --- |
| Report folder | `reports/` |
| Report format | `.xlsx` |
| Report filename | `OrthoLegal_Automation_Test_Report_<DD_Mon_YYYY_HH-MM-SS_AM/PM_IST>.xlsx` |
| Timestamp timezone | `Asia/Kolkata` |
| Per-run behavior | One pytest execution creates one timestamped workbook |
| Multiple tests in same run | All report-enabled tests from that run are written into the same workbook |
| Previous reports | Preserved; filenames are timestamped and collision-safe |

Report worksheets are organized by the test metadata `module` value. Current report columns are:

| Column |
| --- |
| Test Case ID |
| Module |
| Test Case Name |
| Test Case Title |
| Test Type |
| Preconditions |
| Test Steps |
| Test Data |
| Expected Result |
| Actual Result |
| Status |
| Execution Date/Time |

Only tests decorated with `@test_case(...)` metadata are written to the Excel execution report.

Execution reports must not contain passwords, verification PINs, tokens, cookies, or secrets.

## 14. Current Automation Coverage

| Category | Test | File | Data Source | Status |
| --- | --- | --- | --- | --- |
| Login | Super Admin Login | `tests\test_super_admin_login.py` | `.env` | Automated |
| User Creation | Case Manager | `tests\User Creation\test_case_manager_creation.py` | `Users_Test_Data.xlsx` | Automated |
| User Creation | Editor | `tests\User Creation\test_editor_creation.py` | `Users_Test_Data.xlsx` | Automated |
| User Creation | Editor Manager | `tests\User Creation\test_editor_manager_creation.py` | `Users_Test_Data.xlsx` | Automated |
| User Creation | Historian | `tests\User Creation\test_historian_creation.py` | `Users_Test_Data.xlsx` | Automated |
| User Creation | Historian Manager | `tests\User Creation\test_historian_manager_creation.py` | `Users_Test_Data.xlsx` | Automated |
| User Creation | Medical Assistant | `tests\User Creation\test_medical_assistant_creation.py` | `Users_Test_Data.xlsx` | Automated |
| User Creation | Medical Assistant Manager | `tests\User Creation\test_medical_assistant_manager_creation.py` | `Users_Test_Data.xlsx` | Automated |
| User Creation | Record Reviewer | `tests\User Creation\test_record_reviewer_creation.py` | `Users_Test_Data.xlsx` | Automated |
| User Creation | Record Reviewer Manager | `tests\User Creation\test_record_reviewer_manager_creation.py` | `Users_Test_Data.xlsx` | Automated |
| User Creation | Report Finalizer | `tests\User Creation\test_report_finalizer_creation.py` | `Users_Test_Data.xlsx` | Automated |
| User Profile | Case Manager Phone Update | `tests\test_case_manager_phone_update.py` | `Users_Test_Data.xlsx` | Automated |
| Encounter Creation | PQME | `tests\Encounter Creation\test_pqme_encounter_creation.py` | `PQME_Encounter_Test_Data.xlsx` | Automated |
| Encounter Creation | PQMER | `tests\Encounter Creation\test_pqmer_encounter_creation.py` | `PQMER_Encounter_Test_Data.xlsx` | Automated |
| Encounter Creation | PQMES | `tests\Encounter Creation\test_pqmes_encounter_creation.py` | `PQMES_Encounter_Test_Data.xlsx` | Automated |
| Encounter Creation | PQMED | `tests\Encounter Creation\test_pqmed_encounter_creation.py` | `PQMED_Encounter_Test_Data.xlsx` | Automated |
| Encounter Creation | Reflex PQMES | `tests\Encounter Creation\test_reflex_pqmes_encounter_creation.py` | `Reflex_PQMES_Encounter_Test_Data.xlsx` | Automated |
| Utility | Excel Reader | `tests\test_excel_reader.py` | temporary workbook fixture | Automated |
| Utility | Encounter Test Data | `tests\test_encounter_test_data.py` | temporary workbook fixture | Automated |

## 15. Framework Guidelines

Follow these practices when extending the suite:

1. Use Playwright auto-waiting and assertions instead of arbitrary `time.sleep()`.
2. Prefer stable semantic locators such as roles, labels, text, placeholders, and test IDs.
3. Keep credentials in `.env`.
4. Keep workflow input data in Excel workbooks under `test_data/`.
5. Reuse common helpers from `utils/`, `reporting/`, and `conftest.py`.
6. Keep tests independently runnable.
7. Do not hardcode dynamic verification values.
8. Use project-relative paths for files and uploads.
9. Use `set_input_files()` for uploads instead of depending on native file dialogs.
10. Do not depend on browser page indexes for external or PDF tabs.
11. Keep input workbooks separate from execution reports.
12. Do not write passwords, PINs, tokens, cookies, or other secrets to reports.

## 16. Current Notes and Known Conventions

The active repository does not contain `pytest.ini`.

`requirements.txt` currently lists `openpyxl` and `tzdata`. The test code also imports Playwright, pytest, pytest-playwright behavior, and python-dotenv, so the environment must have those packages installed for the full browser suite to run.

`utils/excel_reader.py` supports simple one-row-header worksheets and provides `read_worksheet()` and `get_test_data()`. The user creation tests currently call `read_worksheet()` and perform role matching with the `Select Role` column.

`utils/encounter_test_data.py` supports Encounter workbooks with two header rows, grouping data by `Section -> Field`, and formatting date/date-time inputs for Encounter workflows.

The temporary workbook lock file `test_data/~$Encounter_Test_Data.xlsx` may appear when Excel has a workbook open and should not be treated as a test-data workbook.
