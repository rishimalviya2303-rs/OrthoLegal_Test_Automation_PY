# OrthoLegal Automation Framework

## 1. Overview

This repository contains a standalone end-to-end test automation framework for the OrthoLegal web portal.

The automation framework is maintained separately from the OrthoLegal application source code and interacts with the application through the browser like an actual user.

The framework currently uses:

* Python
* Playwright
* pytest
* pytest-playwright
* python-dotenv
* openpyxl

<<<<<<< HEAD
The report contains test documentation and results only. Do not put passwords, PINs, tokens, or cookies in report metadata.

## Excel Test Data

Place the input workbook at:

`test_data/OrthoLegal_Test_Data.xlsx`

The reusable reader in `utils/excel_reader.py` treats the first row as headers and returns later rows as dictionaries. It skips rows unless `Execute` is `Yes`, case-insensitively. Future tests can filter actions without adding another reader:

```python
from utils.excel_reader import get_test_data

create_users = get_test_data("Users", action="Create")
```

The `Action` value is preserved in each dictionary; UI actions such as Create, Edit, Deactivate, and Activate will be implemented by future tests. Blank cells are returned as `None`. The input workbook is read only and is never used for execution reports.
=======
The framework is being developed incrementally to support functional testing, negative testing, data-driven regression testing, and automated Excel execution reporting.

---

## 2. Main Objectives

The framework is designed to:

* Automate OrthoLegal web portal workflows.
* Maintain automation independently from the application source code.
* Support positive and negative test scenarios.
* Read configuration and credentials securely from environment variables.
* Support dynamic verification PIN handling during login.
* Support Excel-driven regression test data.
* Generate a new Excel execution report for every pytest run.
* Support future automation for Create, Edit, Activate, and Deactivate workflows.
* Keep the framework simple, reusable, and maintainable.

---

## 3. Technology Stack

| Technology        | Purpose                                  |
| ----------------- | ---------------------------------------- |
| Python            | Automation programming language          |
| Playwright        | Browser automation                       |
| pytest            | Test execution framework                 |
| pytest-playwright | Playwright integration with pytest       |
| python-dotenv     | Environment configuration                |
| openpyxl          | Excel input data and execution reporting |

---

## 4. Project Structure

The project is organized approximately as follows:

```text
OrthoLegal_Automation/
│
├── .venv/
│
├── tests/
│   └── test_smoke.py
│
├── utils/
│   ├── excel_reader.py
│   └── <reporting helper>
│
├── test_data/
│   └── OrthoLegal_Test_Data.xlsx
│
├── reports/
│   └── OrthoLegal_Automation_Test_Report_<IST_TIMESTAMP>.xlsx
│
├── .env
├── .env.example
├── .gitignore
├── conftest.py
├── pytest.ini
├── requirements.txt
└── README.md
```

> The exact file names may evolve as the framework grows. Existing working automation should be preserved while the framework is expanded incrementally.

---

# 5. Initial Setup

## 5.1 Create Virtual Environment

From the project root:

```powershell
python -m venv .venv
```

Activate it in PowerShell:

```powershell
.\.venv\Scripts\Activate.ps1
```

After activation, the terminal should display something similar to:

```text
(.venv) PS C:\Rishi Files\OrthoLegal_Automation>
```

---

## 5.2 Install Dependencies

Install the required Python packages:

```powershell
python -m pip install playwright pytest pytest-playwright python-dotenv openpyxl
```

Install Playwright browser binaries:

```powershell
python -m playwright install
```

Verify Playwright:

```powershell
python -m playwright --version
```

Verify pytest:

```powershell
python -m pytest --version
```

---

# 6. Environment Configuration

Application configuration and credentials must not be hardcoded inside test scripts.

Create a local `.env` file in the project root.

Example:

```env
BASE_URL=https://your-ortholegal-environment.com
TEST_EMAIL=your-test-user@example.com
TEST_PASSWORD=your-password
```

A `.env.example` file should contain only variable names/placeholders and must not contain real credentials.

Example:

```env
BASE_URL=
TEST_EMAIL=
TEST_PASSWORD=
```

The `.env` file should be excluded through `.gitignore`.

Never commit:

* Passwords
* Verification PINs
* Authentication tokens
* Cookies
* Other sensitive credentials

---

# 7. Login Automation

The OrthoLegal login workflow contains an additional verification step.

The automated flow is:

```text
Open OrthoLegal Portal
        ↓
Enter Email
        ↓
Enter Password
        ↓
Click Sign In
        ↓
Credentials Validated
        ↓
Verification Screen
        ↓
Read 6-Digit PIN Displayed on Page
        ↓
Enter Same PIN in Verification Field
        ↓
Submit Verification
        ↓
Verify Dashboard
```

The verification PIN must always be read dynamically from the application.

The PIN must never be hardcoded.

Playwright reads the displayed PIN, extracts the required six-digit value when necessary, enters it into the verification field, and completes authentication.

---

# 8. Current Login Test Coverage

The current automation suite includes the following login scenarios.

## LOGIN_TC_001

**Test Case**

Verify that if the user enters a valid email and valid password and completes PIN verification, then the user is successfully logged in.

**Expected Result**

* Valid credentials are accepted.
* Verification screen is displayed.
* The six-digit PIN is dynamically read from the application.
* The PIN is entered into the verification field.
* Verification succeeds.
* The dashboard is displayed.

---

## LOGIN_TC_002

**Test Case**

Verify that if the user enters a valid email and an incorrect password, then an incorrect password error is displayed.

**Expected Result**

* Login fails.
* The appropriate incorrect-password error is displayed.
* Verification PIN screen is not reached.
* Dashboard is not reached.

---

## LOGIN_TC_003

**Test Case**

Verify that if the user enters an invalid email format and a password, then the user is not allowed to proceed with login.

**Expected Result**

* Authentication does not proceed.
* Email-format validation is triggered.
* User remains on the login screen.
* Verification PIN screen is not reached.
* Dashboard is not reached.

The test should verify the application's actual validation behavior, including native HTML email validation where applicable.

---

# 9. Running Tests

Run all available tests:

```powershell
python -m pytest
```

Run tests with the browser visible:

```powershell
python -m pytest --headed
```

Run a specific test file:

```powershell
python -m pytest tests/test_smoke.py --headed
```

Additional pytest selectors can be used to execute individual scenarios as the suite grows.

---

# 10. Excel Execution Reporting

The framework generates an Excel execution report using `openpyxl`.

A new XLSX report is created for every pytest execution.

Previous execution reports must not be overwritten.

## Report Location

```text
reports/
```

## Filename Format

Reports use the pytest execution start date/time in Indian Standard Time.

Example:

```text
OrthoLegal_Automation_Test_Report_23_Aug_2026_03-24-38_PM_IST.xlsx
```

The timezone is explicitly handled using:

```text
Asia/Kolkata
```

One pytest execution creates one Excel report.

For example:

```text
pytest run
    ↓
LOGIN_TC_001
LOGIN_TC_002
LOGIN_TC_003
    ↓
ONE XLSX report
```

Running pytest again generates another report with a new timestamp.

---

# 11. Excel Report Content

Execution reports contain information such as:

* Test Case ID
* Module
* Test Case Title
* Test Type
* Preconditions
* Test Steps
* Test Data
* Expected Result
* Actual Result
* Status
* Execution Date/Time

Status is based on the actual automation execution:

```text
PASS
FAIL
```

A failed automated test must never be artificially recorded as PASS.

The Actual Result should describe the behavior observed during execution.

---

# 12. Test Case Documentation Standard

Test cases should be documented in simple and professional English.

Preferred format:

```text
Verify that if [condition/action], then [expected behavior].
```

Examples:

```text
Verify that if the user enters a valid email and an incorrect password, then an incorrect password error is displayed.
```

```text
Verify that if the user enters an invalid email format, then the user is not allowed to proceed with login.
```

```text
Verify that if valid user details are provided, then the user is created successfully.
```

Requirements may initially be provided in Hinglish, but test case documentation and execution reports should use clear English.

---

# 13. Data-Driven Regression Testing

The framework is being prepared to support Excel-driven regression automation.

Test data is maintained separately from execution reports.

## Input Workbook

The expected input workbook is:

```text
test_data/OrthoLegal_Test_Data.xlsx
```

Do not store the input workbook inside `.venv`.

The input workbook acts only as a test data source.

Execution results must not be written into the input workbook.

---

# 14. Input Workbook Structure

The workbook can contain module-specific worksheets such as:

```text
OrthoLegal_Test_Data.xlsx
│
├── Users
├── Physicians
├── Clinic_Locations
├── Attorneys
└── Insurance_Companies
```

This allows a single input workbook to support multiple regression modules.

---

# 15. Users Input Sheet

The `Users` worksheet is the first planned data-driven workflow.

Example:

| Execute | Action | Test Data ID  | Role      | First Name | Last Name | Email                                         |
| ------- | ------ | ------------- | --------- | ---------- | --------- | --------------------------------------------- |
| Yes     | Create | USER_DATA_001 | Historian | John       | Smith     | [john@example.com](mailto:john@example.com)   |
| Yes     | Create | USER_DATA_002 | Editor    | Sarah      | Lee       | [sarah@example.com](mailto:sarah@example.com) |
| No      | Create | USER_DATA_003 | Historian | Test       | User      | [test@example.com](mailto:test@example.com)   |

Each row represents one data set.

---

# 16. Execute Column

The `Execute` column controls whether an input row should be processed.

```text
Execute = Yes
```

The row should be processed.

```text
Execute = No
```

The row should be skipped.

This allows test data to remain in the workbook without requiring every row to execute during every regression run.

---

# 17. Action Column

The input workbook is designed to support actions such as:

```text
Create
Edit
Deactivate
Activate
```

The first implementation will focus on:

```text
Action = Create
```

Future automation will progressively add Edit, Deactivate, and Activate workflows.

---

# 18. Test Data ID

Every input data row should have a unique identifier.

Examples:

```text
USER_DATA_001
USER_DATA_002
USER_DATA_003
```

This makes individual data sets easier to identify during execution and reporting.

---

# 19. Excel Field Mapping

Excel column headers should represent the corresponding application fields.

Example:

```text
Excel Column              Application

First Name        →       First Name field
Last Name         →       Last Name field
Email             →       Email field
Role              →       Role dropdown
Language          →       Language field
```

The Excel reader itself should not contain Playwright UI locators.

Its responsibility is:

```text
Excel
   ↓
Read Headers
   ↓
Read Row
   ↓
Convert Row to Dictionary
```

Example:

```python
{
    "Execute": "Yes",
    "Action": "Create",
    "Test Data ID": "USER_DATA_001",
    "Role": "Historian",
    "First Name": "John",
    "Last Name": "Smith",
    "Email": "john@example.com"
}
```

A separate automation/mapping layer is responsible for mapping those values to actual application controls.

This separation allows application locators to change without requiring the input workbook format to be redesigned.

---

# 20. Excel Reader

A reusable Excel reader is used/planned under:

```text
utils/excel_reader.py
```

Its responsibility is to:

* Open the input workbook.
* Read a worksheet by name.
* Treat the first row as headers.
* Convert subsequent rows into dictionaries.
* Skip rows where `Execute = No`.
* Allow filtering by `Action`.
* Handle blank cells safely.
* Report meaningful errors when the workbook, worksheet, or required columns are missing.

The same reader should eventually work for every supported module.

Conceptually:

```python
get_test_data("Users", action="Create")
```

returns enabled User Creation datasets.

The same mechanism can later support:

```python
get_test_data("Physicians", action="Create")
get_test_data("Clinic_Locations", action="Create")
get_test_data("Attorneys", action="Create")
get_test_data("Insurance_Companies", action="Create")
```

---

# 21. Planned Data-Driven User Creation Flow

The planned User Creation regression flow is:

```text
Login to OrthoLegal
        ↓
Navigate to User Management
        ↓
Open Create User
        ↓
Read Users Sheet
        ↓
Filter Execute = Yes
        ↓
Filter Action = Create
        ↓
Read One Dataset
        ↓
Map Excel Columns to Application Fields
        ↓
Populate User Form
        ↓
Submit Create Action
        ↓
Verify User Creation
        ↓
Record PASS / FAIL
        ↓
Continue With Next Dataset
```

A failure for one dataset should ideally not prevent the remaining independent datasets from being executed.

---

# 22. Future Regression Coverage

The framework is intended to progressively support data-driven automation for:

### User Management

* Create User
* Edit User
* Deactivate User
* Activate User

### Physician Management

* Create Physician
* Edit Physician
* Deactivate Physician
* Activate Physician

### Clinic Location Management

* Create Clinic Location
* Edit Clinic Location
* Deactivate Clinic Location
* Activate Clinic Location

### Attorney Management

* Create Attorney
* Edit Attorney
* Deactivate Attorney
* Activate Attorney

### Insurance Company Management

* Create Insurance Company
* Edit Insurance Company
* Deactivate Insurance Company
* Activate Insurance Company

Additional OrthoLegal workflows can be added incrementally.

---

# 23. Input Data vs Execution Report

These are two completely different Excel files.

## Input

```text
test_data/OrthoLegal_Test_Data.xlsx
```

Purpose:

```text
Provide data to automation
```

## Output

```text
reports/OrthoLegal_Automation_Test_Report_<IST_TIMESTAMP>.xlsx
```

Purpose:

```text
Record automation execution results
```

The framework must never confuse or combine these responsibilities.

---

# 24. Automation Development Principles

When extending this project:

1. Inspect existing automation before making changes.
2. Preserve working tests.
3. Avoid duplicate test implementations.
4. Prefer stable Playwright locators:

   * `get_by_role`
   * `get_by_label`
   * `get_by_placeholder`
   * `get_by_text`
   * `get_by_test_id`
5. Avoid arbitrary `time.sleep()`.
6. Use Playwright auto-waiting and appropriate assertions.
7. Never hardcode credentials or verification PINs.
8. Keep test data separate from automation logic.
9. Keep execution reporting separate from input test data.
10. Avoid unnecessary framework complexity.
11. Add reusable abstractions only when repeated automation creates a genuine need.
12. Keep test cases understandable and maintainable.

---

# 25. Current Development Status

### Completed / Implemented

* Standalone external automation repository
* Python virtual environment
* Playwright installation
* pytest integration
* pytest-playwright integration
* Playwright browser installation
* Environment-based configuration
* Basic smoke testing
* OrthoLegal login automation
* Dynamic six-digit verification PIN handling
* Successful-login validation
* Incorrect-password negative scenario
* Invalid-email-format negative scenario
* Excel execution reporting foundation
* New timestamped XLSX execution report per pytest run
* IST-based execution timestamps
* Excel input-data template design
* Excel-driven regression foundation / reader

### Next Major Implementation

**Excel-driven User Creation automation**

The next phase will:

1. Read enabled User Creation datasets from the `Users` worksheet.
2. Log in to OrthoLegal.
3. Navigate to User Management.
4. Open the User Creation screen.
5. Map Excel fields to application fields.
6. Populate the form.
7. Create each user.
8. Verify the creation result.
9. Record individual dataset execution results in the timestamped Excel report.

---

# 26. Development Approach

This automation framework is intentionally being developed incrementally.

The approach is:

```text
Understand Manual Workflow
        ↓
Define Test Case
        ↓
Prepare Test Data
        ↓
Automate with Playwright
        ↓
Execute with pytest
        ↓
Validate Actual Behavior
        ↓
Generate Excel Report
        ↓
Refactor Repeated Logic
        ↓
Add Next Regression Scenario
```

The goal is to progressively build a maintainable OrthoLegal regression automation suite without introducing unnecessary complexity before it is required.
>>>>>>> 9243ee2e67300ee59df14f2596af153068686234
