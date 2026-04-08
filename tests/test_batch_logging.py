"""Tests para las features de logging y reporte JSON del batch."""

from __future__ import annotations

import json
from pathlib import Path

from engine.batch_processor import process_batch, write_batch_report_json


def _seed_data_folder(dest: Path, source_xlsx: Path, names: list[str]) -> None:
    dest.mkdir(parents=True, exist_ok=True)
    blob = source_xlsx.read_bytes()
    for name in names:
        (dest / name).write_bytes(blob)


def test_log_file_is_created_and_contains_expected_markers(
    tmp_path: Path,
    sample_pptx: Path,
    sample_xlsx: Path,
    sample_mapping: Path,
) -> None:
    data_folder = tmp_path / "data"
    _seed_data_folder(
        data_folder,
        sample_xlsx,
        ["estudio_1.xlsx", "estudio_2.xlsx"],
    )
    log_file = tmp_path / "logs" / "batch.log"

    result = process_batch(
        template_path=sample_pptx,
        data_folder=data_folder,
        mapping=sample_mapping,
        output_dir=tmp_path / "out",
        log_file=log_file,
    )

    assert len(result.successful) == 2
    assert log_file.exists()
    content = log_file.read_text(encoding="utf-8")
    # Cabeceras y resumen.
    assert "batch run started" in content
    assert "batch run finished" in content
    assert "Resumen: total=2 ok=2 ko=0" in content
    # Una linea OK por archivo.
    assert content.count("OK estudio_1.xlsx") >= 1
    assert content.count("OK estudio_2.xlsx") >= 1


def test_report_json_structure(
    tmp_path: Path,
    sample_pptx: Path,
    sample_xlsx: Path,
    sample_mapping: Path,
) -> None:
    data_folder = tmp_path / "data"
    _seed_data_folder(data_folder, sample_xlsx, ["only.xlsx"])

    result = process_batch(
        template_path=sample_pptx,
        data_folder=data_folder,
        mapping=sample_mapping,
        output_dir=tmp_path / "out",
    )

    report_path = write_batch_report_json(result, tmp_path / "report.json")
    assert report_path.exists()

    data = json.loads(report_path.read_text(encoding="utf-8"))
    assert set(data.keys()) == {"generated_at", "summary", "items"}
    assert data["summary"] == {"total": 1, "successful": 1, "failed": 0}
    assert len(data["items"]) == 1
    item = data["items"][0]
    assert item["ok"] is True
    assert item["charts_updated"] == 1
    assert item["charts_failed"] == 0
    assert item["error"] is None
    assert item["output"].endswith("output_only.pptx")


def test_report_json_with_failure(
    tmp_path: Path,
    sample_pptx: Path,
    sample_mapping: Path,
) -> None:
    """El JSON debe reflejar correctamente un item fallido."""
    data_folder = tmp_path / "data"
    data_folder.mkdir()
    # xlsx corrupto para forzar el error.
    bad = data_folder / "broken.xlsx"
    bad.write_bytes(b"not an xlsx")

    result = process_batch(
        template_path=sample_pptx,
        data_folder=data_folder,
        mapping=sample_mapping,
        output_dir=tmp_path / "out",
    )

    report_path = write_batch_report_json(result, tmp_path / "report.json")
    data = json.loads(report_path.read_text(encoding="utf-8"))
    assert data["summary"]["failed"] == 1
    item = data["items"][0]
    assert item["ok"] is False
    assert item["error"] is not None
