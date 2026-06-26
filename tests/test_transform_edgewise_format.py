"""Unit tests for transform_edgewise_format.py."""
from __future__ import annotations

import pytest

import transform_edgewise_format as tef


class TestExtractTimeFromInterval:
    def test_valid_interval(self):
        start, end = tef.extract_time_from_interval("08:00:00-08:05:00")
        assert start == "08:00:00"
        assert end == "08:05:00"

    def test_different_interval(self):
        start, end = tef.extract_time_from_interval("09:15:00-09:20:00")
        assert start == "09:15:00"
        assert end == "09:20:00"

    def test_invalid_format_no_dash(self):
        start, end = tef.extract_time_from_interval("08:00:00")
        assert start is None
        assert end is None

    def test_midnight_boundary(self):
        start, end = tef.extract_time_from_interval("23:55:00-00:00:00")
        assert start == "23:55:00"
        assert end == "00:00:00"

    def test_empty_string(self):
        start, end = tef.extract_time_from_interval("")
        assert start is None
        assert end is None

    def test_multiple_dashes(self):
        start, end = tef.extract_time_from_interval("08:00:00-08:05:00-extra")
        # Function splits on '-' and checks for exactly 2 parts
        assert start is None
        assert end is None
