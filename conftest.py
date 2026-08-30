from collections.abc import Callable
from typing import Any

import pytest

from reporting import initialize_report, record_test_case


def test_case(**metadata: Any) -> Callable:
    def decorator(test_function: Callable) -> Callable:
        test_function.report_metadata = metadata
        return test_function

    return decorator


def pytest_sessionstart(session: pytest.Session) -> None:
    initialize_report()


@pytest.hookimpl(hookwrapper=True)
def pytest_runtest_makereport(item: pytest.Item, call: pytest.CallInfo[Any]) -> Any:
    outcome = yield
    report = outcome.get_result()
    if report.when != "call":
        return

    test_case_data = getattr(item.obj, "report_metadata", None)
    if test_case_data is None:
        return

    passed = report.passed
    if passed:
        actual_result = test_case_data["actual_result"]
    elif call.excinfo is not None:
        actual_result = f"Test failed during execution: {call.excinfo.typename}: {call.excinfo.value}"
    else:
        actual_result = "Test failed during execution."
    record_test_case(test_case_data, "PASS" if passed else "FAIL", actual_result)
