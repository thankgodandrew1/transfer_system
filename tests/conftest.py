from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest


NEWS_COLUMNS = [
    "Name of Missionary",
    "Assignment",
    "New/Existing Zone",
    "New/Existing Area",
    "New/Existing Companion(s)",
    "Status",
    "Previous Zone",
    "Previous Area",
    "Previous Assignment",
]


@pytest.fixture
def sample_news() -> pd.DataFrame:
    return pd.DataFrame(
        [
            ["ELDER ALPHA", "TR", "UYO", "CENTRAL", "BETA & NEW MISSIONARY", "ROLE CHANGE", "UYO", "CENTRAL", "SC"],
            ["ELDER BETA", "DL", "UYO", "CENTRAL", "ALPHA & GAMMA", "SAME", "UYO", "CENTRAL", "DL"],
            ["ELDER GAMMA", "JC", "UYO", "CENTRAL", "ALPHA & BETA", "SAME", "UYO", "CENTRAL", "JC"],
        ],
        columns=NEWS_COLUMNS,
    )


@pytest.fixture
def sample_workbook(tmp_path: Path, sample_news: pd.DataFrame) -> Path:
    workbook = tmp_path / "transfer.xlsx"
    transfer = pd.DataFrame(
        [{"Zone": "UYO", "District": "CENTRAL", "Area": "CENTRAL", "Missionary 1": "ALPHA"}]
    )
    meta = pd.DataFrame(
        [
            {"Metric": "New Missionaries (Incoming)", "Value": 1},
            {"Metric": "Transfer Title", "Value": "AUGUST 2026 TRANSFER NEWS"},
            {"Metric": "Zone Order", "Value": "UYO"},
        ]
    )
    with pd.ExcelWriter(workbook, engine="openpyxl") as writer:
        sample_news.to_excel(writer, sheet_name="News Format", index=False)
        transfer.to_excel(writer, sheet_name="Transfer Sheet", index=False)
        meta.to_excel(writer, sheet_name="Meta", index=False)
    return workbook

