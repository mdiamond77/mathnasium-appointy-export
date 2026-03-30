from datetime import date
from scripts.appointy_export import get_week_start, get_date_range, to_utc_iso_start, to_utc_iso_end


def test_get_week_start_on_sunday():
    assert get_week_start(date(2026, 3, 29)) == date(2026, 3, 29)  # Sunday → itself


def test_get_week_start_on_wednesday():
    assert get_week_start(date(2026, 4, 1)) == date(2026, 3, 29)  # Wed → prior Sunday


def test_get_week_start_on_saturday():
    assert get_week_start(date(2026, 4, 4)) == date(2026, 3, 29)  # Sat → prior Sunday


def test_get_date_range_start_is_sunday():
    start, end = get_date_range(date(2026, 4, 1))  # Wednesday
    assert start == date(2026, 3, 29)


def test_get_date_range_end_is_70_days_out():
    start, end = get_date_range(date(2026, 4, 1))
    assert end == date(2026, 6, 7)


def test_utc_iso_start_format():
    # March 29 midnight ET (EDT = UTC-4) → 04:00 UTC
    result = to_utc_iso_start(date(2026, 3, 29))
    assert result == "2026-03-29T04:00:00.000Z"


def test_utc_iso_end_format():
    # June 7 11:59 PM ET (EDT = UTC-4) → June 8 03:59 UTC
    result = to_utc_iso_end(date(2026, 6, 7))
    assert result == "2026-06-08T03:59:00.000Z"
