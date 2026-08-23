from openpyxl import Workbook

from utils.excel_reader import get_test_data, read_worksheet


def test_get_test_data_filters_enabled_create_rows(tmp_path) -> None:
    workbook_path = tmp_path / "OrthoLegal_Test_Data.xlsx"
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Users"
    worksheet.append(["Execute", "Action", "Test Data ID", "First Name", "Optional"])
    worksheet.append(["Yes", "Create", "USER_DATA_001", "John", None])
    worksheet.append(["no", "Create", "USER_DATA_002", "Jane", "value"])
    worksheet.append(["YES", "Edit", "USER_DATA_003", "Alex", "value"])
    workbook.save(workbook_path)

    all_rows = read_worksheet("Users", workbook_path)
    create_rows = get_test_data("Users", action="create", workbook_path=workbook_path)

    assert all_rows[0]["Optional"] is None
    assert [row["Test Data ID"] for row in create_rows] == ["USER_DATA_001"]