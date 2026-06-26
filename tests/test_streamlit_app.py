"""Unit tests for streamlit_app.py utility functions.

The module depends on routesampler_utils and streamlit which aren't
available in the test environment, so we mock them before import.
"""
from __future__ import annotations

import sys
import types
from pathlib import Path
from unittest.mock import MagicMock

import pytest


@pytest.fixture(autouse=True, scope="module")
def _mock_streamlit_deps():
    """Stub out unavailable imports so streamlit_app can be loaded."""
    stubs = {}

    # routesampler_utils
    ru = types.ModuleType("routesampler_utils")
    ru.build_location_edge_map = MagicMock()
    ru.infer_route_sampler_path = MagicMock(return_value=None)
    ru.parse_time_value = lambda value: _parse_time_value_stub(value)
    ru.write_edge_data_xml = MagicMock()
    stubs["routesampler_utils"] = ru

    # streamlit
    st_mod = MagicMock()
    st_mod.__name__ = "streamlit"
    stubs["streamlit"] = st_mod

    for name, mod in stubs.items():
        if name not in sys.modules:
            sys.modules[name] = mod

    yield

    for name in stubs:
        if name in sys.modules and sys.modules[name] is stubs[name]:
            del sys.modules[name]


def _parse_time_value_stub(value) -> int | None:
    text = str(value).strip()
    parts = text.split(":")
    if len(parts) != 3:
        return None
    try:
        h, m, s = (int(part) for part in parts)
    except ValueError:
        return None
    if not (0 <= h <= 23 and 0 <= m <= 59 and 0 <= s <= 59):
        return None
    return h * 3600 + m * 60 + s


def _get_app():
    import streamlit_app
    return streamlit_app


class TestNormalizeKey:
    def test_basic(self):
        app = _get_app()
        assert app.normalize_key("Direction") == "direction"

    def test_strips_non_alnum(self):
        app = _get_app()
        assert app.normalize_key("  Vehicle Class  ") == "vehicleclass"

    def test_numeric(self):
        app = _get_app()
        assert app.normalize_key("Count123") == "count123"

    def test_special_chars_removed(self):
        app = _get_app()
        assert app.normalize_key("Hello-World!") == "helloworld"


class TestNormalizeDateToken:
    def test_dash_to_space(self):
        app = _get_app()
        assert app.normalize_date_token("2026-02-22") == "2026 02 22"

    def test_strips_whitespace(self):
        app = _get_app()
        assert app.normalize_date_token("  2026-02-22  ") == "2026 02 22"

    def test_no_dashes(self):
        app = _get_app()
        assert app.normalize_date_token("20260222") == "20260222"


class TestDateMatchKey:
    def test_compact_eight_digits(self):
        app = _get_app()
        assert app.date_match_key("20260222") == "2026-02-22"

    def test_already_formatted(self):
        app = _get_app()
        assert app.date_match_key("2026-02-22") == "2026-02-22"

    def test_space_separated(self):
        app = _get_app()
        assert app.date_match_key("2026 02 22") == "2026-02-22"

    def test_slash_separated(self):
        app = _get_app()
        assert app.date_match_key("2026/02/22") == "2026-02-22"


class TestParseIntervalStartCode:
    def test_four_digit(self):
        app = _get_app()
        result = app.parse_interval_start_code("0905")
        assert result == "09:15:00"

    def test_three_digit(self):
        app = _get_app()
        result = app.parse_interval_start_code("905")
        assert result == "09:15:00"

    def test_invalid_text(self):
        app = _get_app()
        assert app.parse_interval_start_code("abc") is None

    def test_hour_out_of_range(self):
        app = _get_app()
        assert app.parse_interval_start_code("2500") is None

    def test_minute_out_of_range(self):
        app = _get_app()
        assert app.parse_interval_start_code("0960") is None

    def test_midnight(self):
        app = _get_app()
        result = app.parse_interval_start_code("0000")
        assert result == "00:10:00"


class TestParseIntervalEndCode:
    def test_basic(self):
        app = _get_app()
        result = app.parse_interval_end_code("09:15:00")
        assert result == "09:20:00"

    def test_hour_boundary(self):
        app = _get_app()
        result = app.parse_interval_end_code("09:55:00")
        assert result == "10:00:00"


class TestParseClockToken:
    def test_colon_separated(self):
        app = _get_app()
        assert app.parse_clock_token("09:15:30") == "09:15:30"

    def test_dot_separated(self):
        app = _get_app()
        assert app.parse_clock_token("09.15.30") == "09:15:30"

    def test_underscore_separated(self):
        app = _get_app()
        assert app.parse_clock_token("09_15_30") == "09:15:30"

    def test_two_parts(self):
        app = _get_app()
        assert app.parse_clock_token("09:15") == "09:15:00"

    def test_invalid(self):
        app = _get_app()
        assert app.parse_clock_token("abc") is None

    def test_out_of_range(self):
        app = _get_app()
        assert app.parse_clock_token("25:00:00") is None


class TestParseIntervalFromFilename:
    def test_time_range(self):
        app = _get_app()
        start, end = app.parse_interval_from_filename(Path("data_09.00.00-09.05.00.xlsx"))
        assert start == "09:00:00"
        assert end == "09:05:00"

    def test_numeric_code(self):
        app = _get_app()
        start, end = app.parse_interval_from_filename(Path("0905.xlsx"))
        assert start == "09:15:00"
        assert end == "09:20:00"

    def test_no_interval(self):
        app = _get_app()
        start, end = app.parse_interval_from_filename(Path("random.xlsx"))
        assert start is None
        assert end is None


class TestParseDirectionalHeaderName:
    def test_valid(self):
        app = _get_app()
        result = app.parse_directional_header_name("North-to-South | Car")
        assert result == ("North-to-South", "Car")

    def test_no_pipe(self):
        app = _get_app()
        result = app.parse_directional_header_name("Direction")
        assert result is None

    def test_empty_parts(self):
        app = _get_app()
        result = app.parse_directional_header_name("| Car")
        assert result is None


class TestAddRow:
    def test_basic(self):
        app = _get_app()
        records = []
        base = {
            "Location": "Test",
            "SessionFolder": "20260222",
            "IntervalFolder": "9",
            "TimeStart": "09:00:00",
            "TimeEnd": "09:05:00",
            "SourceWorkbook": "test.xlsx",
        }
        app.add_row(records, base, "North-to-South", "Car", 50)
        assert len(records) == 1
        assert records[0]["Direction"] == "North-to-South"
        assert records[0]["VehicleClass"] == "Car"
        assert records[0]["Count"] == 50

    def test_none_count_skipped(self):
        app = _get_app()
        records = []
        base = {"Location": "Test", "SessionFolder": "", "IntervalFolder": "", "TimeStart": "", "TimeEnd": "", "SourceWorkbook": ""}
        app.add_row(records, base, "North", "Car", None)
        assert len(records) == 0

    def test_negative_count_skipped(self):
        app = _get_app()
        records = []
        base = {"Location": "Test", "SessionFolder": "", "IntervalFolder": "", "TimeStart": "", "TimeEnd": "", "SourceWorkbook": ""}
        app.add_row(records, base, "North", "Car", -5)
        assert len(records) == 0

    def test_approach_derived_from_direction(self):
        app = _get_app()
        records = []
        base = {"Location": "Test", "SessionFolder": "", "IntervalFolder": "", "TimeStart": "", "TimeEnd": "", "SourceWorkbook": ""}
        app.add_row(records, base, "East-to-West", "Bus", 10)
        assert records[0]["Approach"] == "East"


class TestLocateMetadata:
    def test_four_level_path(self):
        app = _get_app()
        root = Path("/data")
        path = Path("/data/Location/20260222/9/file.xlsx")
        loc, session, interval = app.locate_metadata(path, root)
        assert loc == "Location"
        assert session == "20260222"
        assert interval == "9"

    def test_short_path(self):
        app = _get_app()
        root = Path("/data")
        path = Path("/data/file.xlsx")
        loc, session, interval = app.locate_metadata(path, root)
        assert loc is not None


class TestPathRelativeToRoot:
    def test_under_root(self):
        app = _get_app()
        assert app.path_relative_to_root(Path("/a/b/c.txt"), Path("/a")) == "b/c.txt"

    def test_not_under_root(self):
        app = _get_app()
        result = app.path_relative_to_root(Path("/x/y/z.txt"), Path("/a"))
        assert "z.txt" in result


class TestDiscoverExcelFiles:
    def test_empty_root(self, tmp_path):
        app = _get_app()
        result = app.discover_excel_files(tmp_path)
        assert result == []

    def test_nonexistent_root(self):
        app = _get_app()
        result = app.discover_excel_files(Path("/nonexistent"))
        assert result == []
