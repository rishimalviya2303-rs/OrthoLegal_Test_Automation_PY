"""Business-day utilities for OrthoLegal Encounter due-date calculations.

Provides:
- Dynamic U.S. federal holiday generation for any calendar year.
- Report Due Date calculation: evaluation date + 30 calendar days, adjusted
  backward for weekends and federal holidays.
"""

from __future__ import annotations

import calendar
from datetime import date, timedelta


def _nth_weekday_of_month(year: int, month: int, weekday: int, n: int) -> date:
    """Return the *n*-th occurrence of *weekday* in the given month/year.

    *weekday* uses Monday=0 … Sunday=6 (matching ``date.weekday()``).
    *n* is 1-based (1 = first, 2 = second, …).
    """
    first_day, days_in_month = calendar.monthrange(year, month)
    # Offset from the 1st of the month to the first target weekday
    first_occurrence = 1 + (weekday - first_day) % 7
    target = first_occurrence + 7 * (n - 1)
    if target > days_in_month:
        raise ValueError(
            f"There is no {n}-th weekday={weekday} in {year}-{month:02d}."
        )
    return date(year, month, target)


def _last_weekday_of_month(year: int, month: int, weekday: int) -> date:
    """Return the last occurrence of *weekday* in the given month/year."""
    _, days_in_month = calendar.monthrange(year, month)
    last_day = date(year, month, days_in_month)
    offset = (last_day.weekday() - weekday) % 7
    return last_day - timedelta(days=offset)


def _observed(holiday: date) -> date:
    """Apply the U.S. federal observed-holiday rule.

    If the actual holiday falls on Saturday the observed date is the
    preceding Friday; if it falls on Sunday the observed date is the
    following Monday.
    """
    if holiday.weekday() == 5:  # Saturday
        return holiday - timedelta(days=1)
    if holiday.weekday() == 6:  # Sunday
        return holiday + timedelta(days=1)
    return holiday


def us_federal_holidays(year: int) -> set[date]:
    """Return the set of **observed** U.S. federal holidays for *year*.

    Covers all 11 federally recognized holidays:

    1.  New Year's Day              — January 1
    2.  Martin Luther King Jr. Day  — 3rd Monday of January
    3.  Presidents' Day             — 3rd Monday of February
    4.  Memorial Day                — Last Monday of May
    5.  Juneteenth                  — June 19
    6.  Independence Day            — July 4
    7.  Labor Day                   — 1st Monday of September
    8.  Columbus Day                — 2nd Monday of October
    9.  Veterans Day                — November 11
    10. Thanksgiving Day            — 4th Thursday of November
    11. Christmas Day               — December 25
    """
    holidays: set[date] = set()

    # Fixed-date holidays (use observed rule)
    holidays.add(_observed(date(year, 1, 1)))    # New Year's Day
    holidays.add(_observed(date(year, 6, 19)))   # Juneteenth
    holidays.add(_observed(date(year, 7, 4)))    # Independence Day
    holidays.add(_observed(date(year, 11, 11)))  # Veterans Day
    holidays.add(_observed(date(year, 12, 25)))  # Christmas Day

    # Floating holidays (always fall on the specified weekday — no observed shift)
    holidays.add(_nth_weekday_of_month(year, 1, 0, 3))   # MLK Day — 3rd Mon Jan
    holidays.add(_nth_weekday_of_month(year, 2, 0, 3))   # Presidents' Day — 3rd Mon Feb
    holidays.add(_last_weekday_of_month(year, 5, 0))      # Memorial Day — last Mon May
    holidays.add(_nth_weekday_of_month(year, 9, 0, 1))   # Labor Day — 1st Mon Sep
    holidays.add(_nth_weekday_of_month(year, 10, 0, 2))  # Columbus Day — 2nd Mon Oct
    holidays.add(_nth_weekday_of_month(year, 11, 3, 4))  # Thanksgiving — 4th Thu Nov

    return holidays


def is_business_day(d: date, holidays: set[date] | None = None) -> bool:
    """Return ``True`` if *d* is neither a weekend nor a federal holiday."""
    if d.weekday() >= 5:  # Saturday or Sunday
        return False
    if holidays is None:
        holidays = us_federal_holidays(d.year)
    return d not in holidays


def calculate_report_due_date(evaluation_date: date) -> date:
    """Calculate the expected PQME Report Due Date.

    Business rule:
        Base due date = evaluation_date + 30 calendar days.
        If the base date falls on a weekend or federal holiday, move
        **backward** one day at a time until a valid business day is reached.

    Returns the adjusted Report Due Date.
    """
    base_due_date = evaluation_date + timedelta(days=30)

    # Pre-compute holidays for the year(s) that the due date may span
    holidays = us_federal_holidays(base_due_date.year)
    # Edge case: backward adjustment might cross a year boundary
    if base_due_date.month == 1 and base_due_date.day <= 5:
        holidays |= us_federal_holidays(base_due_date.year - 1)

    adjusted = base_due_date
    while not is_business_day(adjusted, holidays):
        adjusted -= timedelta(days=1)
        # If we've crossed into a new year, refresh the holiday set
        if adjusted.year != base_due_date.year and adjusted.year not in {
            h.year for h in holidays
        }:
            holidays |= us_federal_holidays(adjusted.year)

    return adjusted
