"""Tests unitarios para ``engine.validator``."""

from __future__ import annotations

from pathlib import Path

import pytest

from engine.validator import (
    Mapping,
    load_mapping,
    validate_all,
    validate_excel_against_mapping,
    validate_template_against_mapping,
)


def _write_yaml(path: Path, content: str) -> Path:
    path.write_text(content, encoding="utf-8")
    return path


def test_load_mapping_valid(sample_mapping: Path) -> None:
    mapping = load_mapping(sample_mapping)
    assert isinstance(mapping, Mapping)
    assert len(mapping.slides) == 1
    slide = mapping.slides[0]
    assert slide.slide_index == 2
    assert slide.charts[0].excel_sheet == "P1_satisfaccion"
    assert slide.charts[0].data_range == "A1:E6"


def test_load_mapping_without_slides(tmp_path: Path) -> None:
    bad = _write_yaml(tmp_path / "bad.yaml", "slides: []\n")
    with pytest.raises(ValueError):
        load_mapping(bad)


def test_load_mapping_missing_fields(tmp_path: Path) -> None:
    bad = _write_yaml(
        tmp_path / "bad.yaml",
        "slides:\n  - slide_index: 1\n    charts:\n      - {}\n",
    )
    with pytest.raises(ValueError):
        load_mapping(bad)


def test_validate_excel_ok(sample_xlsx: Path, sample_mapping: Path) -> None:
    mapping = load_mapping(sample_mapping)
    result = validate_excel_against_mapping(sample_xlsx, mapping)
    assert result.ok
    assert not result.errors


def test_validate_excel_missing_sheet(
    tmp_path: Path, sample_xlsx: Path
) -> None:
    bad = _write_yaml(
        tmp_path / "bad.yaml",
        (
            "slides:\n"
            "  - slide_index: 2\n"
            "    charts:\n"
            "      - chart_name: x\n"
            "        excel_sheet: hoja_inexistente\n"
            "        data_range: A1:B2\n"
        ),
    )
    mapping = load_mapping(bad)
    result = validate_excel_against_mapping(sample_xlsx, mapping)
    assert not result.ok
    assert any("hoja_inexistente" in e for e in result.errors)


def test_validate_template_ok(
    sample_pptx: Path, sample_mapping: Path
) -> None:
    mapping = load_mapping(sample_mapping)
    result = validate_template_against_mapping(sample_pptx, mapping)
    assert result.ok


def test_validate_template_bad_slide_index(
    tmp_path: Path, sample_pptx: Path
) -> None:
    bad = _write_yaml(
        tmp_path / "bad.yaml",
        (
            "slides:\n"
            "  - slide_index: 999\n"
            "    charts:\n"
            "      - chart_name: x\n"
            "        excel_sheet: P1_satisfaccion\n"
            "        data_range: A1:B2\n"
        ),
    )
    mapping = load_mapping(bad)
    result = validate_template_against_mapping(sample_pptx, mapping)
    assert not result.ok
    assert any("999" in e for e in result.errors)


def test_validate_all(
    sample_pptx: Path, sample_xlsx: Path, sample_mapping: Path
) -> None:
    result = validate_all(sample_pptx, sample_xlsx, sample_mapping)
    assert result.ok
