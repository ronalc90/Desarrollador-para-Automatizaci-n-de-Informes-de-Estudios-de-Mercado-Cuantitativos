"""Tests unitarios para ``engine.excel_reader``."""

from __future__ import annotations

from pathlib import Path

import pytest

from engine.excel_reader import (
    ExcelFileNotFoundError,
    ExcelReader,
    InvalidRangeError,
    RangeNotFoundError,
    SheetNotFoundError,
)


def test_sheet_names(sample_xlsx: Path) -> None:
    with ExcelReader(sample_xlsx) as reader:
        assert "P1_satisfaccion" in reader.sheet_names
        assert "P3_nps" in reader.sheet_names


def test_get_table_with_header(sample_xlsx: Path) -> None:
    with ExcelReader(sample_xlsx) as reader:
        df = reader.get_table("P1_satisfaccion", "A1:E6")
    assert list(df.columns) == ["Dimension", "2022", "2023", "2024", "2025"]
    assert df.shape == (5, 5)
    # Primera fila de datos.
    row = df.iloc[0]
    assert row["Dimension"] == "Producto"
    assert row["2022"] == 70
    assert row["2025"] == 78


def test_get_table_without_header(sample_xlsx: Path) -> None:
    with ExcelReader(sample_xlsx) as reader:
        df = reader.get_table("P1_satisfaccion", "A2:E2", header=False)
    assert df.shape == (1, 5)
    # Con header=False los nombres son col_0..col_N.
    assert df.iloc[0]["col_0"] == "Producto"
    assert df.iloc[0]["col_4"] == 78


def test_missing_sheet_raises(sample_xlsx: Path) -> None:
    with ExcelReader(sample_xlsx) as reader:
        with pytest.raises(SheetNotFoundError):
            reader.get_table("hoja_inexistente", "A1:B2")


def test_invalid_range_raises(sample_xlsx: Path) -> None:
    with ExcelReader(sample_xlsx) as reader:
        with pytest.raises(InvalidRangeError):
            reader.get_table("P1_satisfaccion", "not_a_range")


def test_empty_range_raises(sample_xlsx: Path) -> None:
    with ExcelReader(sample_xlsx) as reader:
        with pytest.raises(RangeNotFoundError):
            reader.get_table("P99_vacia", "A1:D5")


def test_missing_file_raises(tmp_path: Path) -> None:
    missing = tmp_path / "nope.xlsx"
    with pytest.raises(ExcelFileNotFoundError):
        ExcelReader(missing)


def test_requires_range_or_table(sample_xlsx: Path) -> None:
    with ExcelReader(sample_xlsx) as reader:
        with pytest.raises(ValueError):
            reader.get_table("P1_satisfaccion")
        with pytest.raises(ValueError):
            reader.get_table(
                "P1_satisfaccion", data_range="A1:B2", table_id="T1"
            )
