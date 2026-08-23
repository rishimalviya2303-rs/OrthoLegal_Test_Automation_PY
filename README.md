# OrthoLegal Playwright Tests

Simple end-to-end tests for the OrthoLegal web application using Playwright, pytest, and Python.

## Excel Report

Every pytest run creates one new workbook in `reports/`. The filename includes the run start time in IST, for example:

`reports/OrthoLegal_Automation_Test_Report_23_Aug_2026_03-24-38_PM_IST.xlsx`

The pytest session hook creates the report path once. Each reportable test uses `@report_case(...)` metadata, and the hook records all results from that run in the same workbook. Existing Test Case IDs are updated within that run instead of duplicated. To add a future test, add the decorator with a new ID and module name. A new module name creates a worksheet while preserving existing worksheets.

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
