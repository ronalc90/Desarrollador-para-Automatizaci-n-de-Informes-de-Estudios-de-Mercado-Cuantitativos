"""Tests para ``engine.inspector``."""

from __future__ import annotations

from pathlib import Path

import pytest

from engine.inspector import TemplateReport, inspect_template


def test_inspect_basic(sample_pptx: Path) -> None:
    report = inspect_template(sample_pptx)
    assert isinstance(report, TemplateReport)
    # El fixture tiene 3 slides: portada (0 charts), slide con 1 chart
    # por nombre y slide con 2 charts para chart_index.
    assert report.total_slides == 3
    assert report.total_charts == 3

    portada, slide_named, slide_multi = report.slides
    assert portada.slide_index == 1
    assert portada.charts == []

    assert slide_named.slide_index == 2
    assert len(slide_named.charts) == 1
    chart_info = slide_named.charts[0]
    assert chart_info.shape_name == "Grafico satisfaccion"
    assert chart_info.chart_index == 0
    assert chart_info.chart_type  # algun string no vacio

    assert slide_multi.slide_index == 3
    assert len(slide_multi.charts) == 2
    assert [c.chart_index for c in slide_multi.charts] == [0, 1]


def test_as_text_contains_slide_numbers(sample_pptx: Path) -> None:
    report = inspect_template(sample_pptx)
    text = report.as_text()
    assert "Slide 1" in text
    assert "Slide 2" in text
    assert "Slide 3" in text
    assert "Grafico satisfaccion" in text


def test_as_mapping_stub_only_includes_slides_with_charts(
    sample_pptx: Path,
) -> None:
    stub = inspect_template(sample_pptx).as_mapping_stub()
    assert "slide_index: 2" in stub
    assert "slide_index: 3" in stub
    # La portada no tiene graficos: no debe aparecer.
    assert "slide_index: 1" not in stub
    assert "Grafico satisfaccion" in stub
    assert "TODO" in stub  # placeholders listos para editar


def test_inspect_missing_template(tmp_path: Path) -> None:
    with pytest.raises(FileNotFoundError):
        inspect_template(tmp_path / "nope.pptx")
