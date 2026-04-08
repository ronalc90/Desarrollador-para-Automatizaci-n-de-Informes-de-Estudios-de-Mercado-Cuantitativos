"""Tests para ``engine.inspector``."""

from __future__ import annotations

from pathlib import Path

import pytest

from engine.inspector import TemplateReport, inspect_template


def test_inspect_basic(sample_pptx: Path) -> None:
    report = inspect_template(sample_pptx)
    assert isinstance(report, TemplateReport)
    # El fixture tiene 2 slides: portada (sin grafico) + slide con 1 grafico.
    assert report.total_slides == 2
    assert report.total_charts == 1

    portada, slide_chart = report.slides
    assert portada.slide_index == 1
    assert portada.charts == []
    assert slide_chart.slide_index == 2
    assert len(slide_chart.charts) == 1
    chart_info = slide_chart.charts[0]
    assert chart_info.shape_name == "Grafico satisfaccion"
    assert chart_info.chart_index == 0
    assert chart_info.chart_type  # algun string no vacio


def test_as_text_contains_slide_numbers(sample_pptx: Path) -> None:
    report = inspect_template(sample_pptx)
    text = report.as_text()
    assert "Slide 1" in text
    assert "Slide 2" in text
    assert "Grafico satisfaccion" in text


def test_as_mapping_stub_only_includes_slides_with_charts(
    sample_pptx: Path,
) -> None:
    stub = inspect_template(sample_pptx).as_mapping_stub()
    assert "slide_index: 2" in stub
    # La portada no tiene graficos: no debe aparecer.
    assert "slide_index: 1" not in stub
    assert "Grafico satisfaccion" in stub
    assert "TODO" in stub  # placeholders listos para editar


def test_inspect_missing_template(tmp_path: Path) -> None:
    with pytest.raises(FileNotFoundError):
        inspect_template(tmp_path / "nope.pptx")
