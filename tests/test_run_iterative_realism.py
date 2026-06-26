"""Unit tests for _run_iterative_realism.py utility functions.

The module has top-level code that runs on import, so we extract
the testable functions via AST/exec rather than normal import.
"""
from __future__ import annotations

import ast
import textwrap
from pathlib import Path

import pandas as pd
import pytest


def _load_functions():
    """Load only the function definitions from _run_iterative_realism.py."""
    source = Path("_run_iterative_realism.py").read_text(encoding="utf-8")
    tree = ast.parse(source)

    func_lines = []
    for node in ast.walk(tree):
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            start = node.lineno - 1
            end = node.end_lineno
            func_lines.append((start, end))

    source_lines = source.splitlines()
    import_block = [
        "import re",
        "import pandas as pd",
        "import openpyxl",
    ]
    extracted = "\n".join(import_block) + "\n"
    for start, end in func_lines:
        extracted += "\n" + "\n".join(source_lines[start:end]) + "\n"

    namespace = {}
    exec(extracted, namespace)
    return namespace


_NS = _load_functions()
parse_report = _NS["parse_report"]
realism_rank = _NS["realism_rank"]
scale_traffic = _NS["scale_traffic"]


class TestParseReport:
    def test_full_report(self, tmp_path):
        report_text = textwrap.dedent("""\
            Some header
            - Realism classification: Moderate congestion
            - Mean waiting time: 12.5 seconds
            - Mean speed relative (edge weighted): 0.850
            - Edges with teleports: 3
            - GEH < 5 share: 85.0%
        """)
        report_file = tmp_path / "report.txt"
        report_file.write_text(report_text, encoding="utf-8")
        result = parse_report(report_file)
        assert result["realism"] == "Moderate congestion"
        assert result["wait"] == 12.5
        assert result["speed"] == 0.850
        assert result["tele"] == 3.0
        assert result["geh"] == 5.0

    def test_missing_keys(self, tmp_path):
        report_text = "Nothing useful here\n"
        report_file = tmp_path / "report.txt"
        report_file.write_text(report_text, encoding="utf-8")
        result = parse_report(report_file)
        assert result["realism"] == "Unknown"
        assert result["wait"] == 1e9
        assert result["speed"] == 0.0
        assert result["tele"] == 1e9
        assert result["geh"] == 0.0

    def test_partial_report(self, tmp_path):
        report_text = textwrap.dedent("""\
            - Realism classification: Near free-flow
            - Mean waiting time: 3.2 seconds
        """)
        report_file = tmp_path / "report.txt"
        report_file.write_text(report_text, encoding="utf-8")
        result = parse_report(report_file)
        assert result["realism"] == "Near free-flow"
        assert result["wait"] == 3.2
        assert result["speed"] == 0.0


class TestRealismRank:
    def test_near_free_flow(self):
        assert realism_rank("Near free-flow") == 0

    def test_moderate_congestion(self):
        assert realism_rank("Moderate congestion") == 1

    def test_heavy_congestion(self):
        assert realism_rank("Heavy congestion") == 2

    def test_severely_congested(self):
        assert realism_rank("Severely congested") == 3

    def test_unknown(self):
        assert realism_rank("Unknown") == 4

    def test_unrecognized_label(self):
        assert realism_rank("Something else") == 4

    def test_ordering(self):
        assert realism_rank("Near free-flow") < realism_rank("Moderate congestion")
        assert realism_rank("Moderate congestion") < realism_rank("Heavy congestion")
        assert realism_rank("Heavy congestion") < realism_rank("Severely congested")


class TestScaleTraffic:
    def test_scaling(self, tmp_path):
        src = tmp_path / "input.xlsx"
        dst = tmp_path / "output.xlsx"
        df = pd.DataFrame({"count": [10, 20, 30], "other": ["a", "b", "c"]})
        df.to_excel(src, index=False)
        scale_traffic(src, dst, 0.5)
        result = pd.read_excel(dst)
        assert list(result["count"]) == [5.0, 10.0, 15.0]
        assert list(result["other"]) == ["a", "b", "c"]

    def test_scaling_clamps_to_minimum_1(self, tmp_path):
        src = tmp_path / "input.xlsx"
        dst = tmp_path / "output.xlsx"
        df = pd.DataFrame({"count": [1, 2, 0]})
        df.to_excel(src, index=False)
        scale_traffic(src, dst, 0.1)
        result = pd.read_excel(dst)
        assert all(v >= 1.0 for v in result["count"])
