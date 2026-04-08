"""Procesamiento por lotes de varios .xlsx con el mismo template.

Recibe una carpeta con múltiples archivos de datos y genera un .pptx
por cada uno aplicando el mismo template y mapping. Se reporta un
``BatchResult`` con los archivos exitosos, los fallidos y las razones.

El ``BatchResult`` puede serializarse a JSON con
:func:`write_batch_report_json` para consumo programático.
"""

from __future__ import annotations

import json
import logging
import time
from contextlib import contextmanager
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Iterator, Optional, Union

from engine.excel_reader import ExcelReaderError
from engine.ppt_builder import BuildResult, PPTBuilderError, build_presentation
from engine.validator import Mapping, load_mapping

logger = logging.getLogger(__name__)


@dataclass
class BatchItemResult:
    """Resultado del procesamiento de un archivo individual."""

    input_path: Path
    build_result: Optional[BuildResult] = None
    error: Optional[str] = None
    duration_s: float = 0.0

    @property
    def ok(self) -> bool:
        return self.error is None and (
            self.build_result is not None and self.build_result.ok
        )


@dataclass
class BatchResult:
    """Resultado agregado de un proceso batch."""

    items: list[BatchItemResult] = field(default_factory=list)

    @property
    def successful(self) -> list[BatchItemResult]:
        return [i for i in self.items if i.ok]

    @property
    def failed(self) -> list[BatchItemResult]:
        return [i for i in self.items if not i.ok]

    def summary(self) -> str:
        parts = [
            f"Total: {len(self.items)}",
            f"Exitosos: {len(self.successful)}",
            f"Fallidos: {len(self.failed)}",
            "",
        ]
        if self.successful:
            parts.append("Exitosos:")
            for item in self.successful:
                assert item.build_result is not None  # por el property ok
                parts.append(
                    f"  - {item.input_path.name} -> "
                    f"{item.build_result.output_path.name} "
                    f"({item.build_result.charts_updated} graficos, "
                    f"{item.duration_s:.2f}s)"
                )
        if self.failed:
            parts.append("")
            parts.append("Fallidos:")
            for item in self.failed:
                reason = item.error or (
                    "errores durante el build"
                    if item.build_result
                    else "desconocido"
                )
                parts.append(f"  - {item.input_path.name}: {reason}")
        return "\n".join(parts)


def _discover_excel_files(data_folder: Path, pattern: str) -> list[Path]:
    if not data_folder.exists():
        raise FileNotFoundError(f"Carpeta de datos no encontrada: {data_folder}")
    if not data_folder.is_dir():
        raise NotADirectoryError(f"{data_folder} no es un directorio.")
    files = sorted(data_folder.glob(pattern))
    # Ignorar archivos de lock de Office (~$nombre.xlsx).
    files = [f for f in files if not f.name.startswith("~$")]
    return files


@contextmanager
def _engine_file_logging(log_file: Optional[Path]) -> Iterator[None]:
    """Agrega temporalmente un FileHandler al logger raíz de ``engine``.

    Todos los mensajes emitidos durante el ``with`` se escriben también
    al archivo indicado. Al salir se remueve el handler y se cierra el
    archivo para no dejar file descriptors colgados.
    """
    if log_file is None:
        yield
        return

    log_file = Path(log_file)
    log_file.parent.mkdir(parents=True, exist_ok=True)

    handler = logging.FileHandler(str(log_file), mode="a", encoding="utf-8")
    handler.setLevel(logging.INFO)
    handler.setFormatter(
        logging.Formatter(
            "%(asctime)s [%(levelname)s] %(name)s: %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        )
    )
    engine_logger = logging.getLogger("engine")
    prev_level = engine_logger.level
    if prev_level > logging.INFO or prev_level == logging.NOTSET:
        engine_logger.setLevel(logging.INFO)
    engine_logger.addHandler(handler)

    start = datetime.now().isoformat(timespec="seconds")
    engine_logger.info("=== batch run started %s ===", start)
    try:
        yield
    finally:
        end = datetime.now().isoformat(timespec="seconds")
        engine_logger.info("=== batch run finished %s ===", end)
        engine_logger.removeHandler(handler)
        handler.close()
        if prev_level != logging.NOTSET:
            engine_logger.setLevel(prev_level)


def _item_to_dict(item: BatchItemResult) -> dict:
    data = {
        "input": str(item.input_path),
        "ok": item.ok,
        "duration_s": round(item.duration_s, 4),
        "error": item.error,
    }
    if item.build_result is not None:
        data["output"] = str(item.build_result.output_path)
        data["charts_updated"] = item.build_result.charts_updated
        data["charts_failed"] = item.build_result.charts_failed
        data["build_errors"] = list(item.build_result.errors)
        data["build_warnings"] = list(item.build_result.warnings)
    else:
        data["output"] = None
    return data


def write_batch_report_json(
    result: BatchResult, output_path: Union[str, Path]
) -> Path:
    """Serializa un ``BatchResult`` a un archivo JSON estructurado.

    El JSON tiene la forma::

        {
          "generated_at": "2026-04-08T12:34:56",
          "summary": {"total": N, "successful": M, "failed": K},
          "items": [ ... ]
        }
    """
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "summary": {
            "total": len(result.items),
            "successful": len(result.successful),
            "failed": len(result.failed),
        },
        "items": [_item_to_dict(it) for it in result.items],
    }
    output_path.write_text(
        json.dumps(payload, indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    return output_path


def process_batch(
    template_path: Union[str, Path],
    data_folder: Union[str, Path],
    mapping: Union[str, Path, Mapping],
    output_dir: Union[str, Path],
    *,
    pattern: str = "*.xlsx",
    log_file: Optional[Union[str, Path]] = None,
) -> BatchResult:
    """Procesa todos los archivos Excel de ``data_folder``.

    Por cada archivo se llama a :func:`engine.ppt_builder.build_presentation`
    con el mismo template y mapping. Los errores por archivo se capturan
    y se reportan en el ``BatchResult`` sin interrumpir los demás.

    Parameters
    ----------
    template_path:
        Template .pptx común a todos los archivos.
    data_folder:
        Carpeta con los archivos de datos.
    mapping:
        Ruta al YAML o instancia :class:`Mapping` ya cargada.
    output_dir:
        Carpeta donde se dejarán los .pptx generados.
    pattern:
        Patrón glob para descubrir archivos (por defecto ``*.xlsx``).
    log_file:
        Si se indica, todos los mensajes de log del namespace ``engine``
        emitidos durante el batch se escriben también a este archivo
        (modo append). Útil para auditar corridas largas.

    Returns
    -------
    BatchResult
        Resultado agregado con exitosos, fallidos y duración por item.
    """
    template_path = Path(template_path)
    data_folder = Path(data_folder)
    output_dir = Path(output_dir)
    log_file_path = Path(log_file) if log_file is not None else None

    # Cargamos el mapping una sola vez para no releer YAML por archivo.
    if isinstance(mapping, (str, Path)):
        mapping_obj = load_mapping(mapping)
    else:
        mapping_obj = mapping

    files = _discover_excel_files(data_folder, pattern)

    batch = BatchResult()

    with _engine_file_logging(log_file_path):
        if not files:
            logger.warning(
                "No se encontraron archivos con patron %s en %s",
                pattern,
                data_folder,
            )

        for file_path in files:
            logger.info("Procesando %s", file_path.name)
            started = time.perf_counter()
            item = BatchItemResult(input_path=file_path)
            try:
                item.build_result = build_presentation(
                    template_path=template_path,
                    excel_path=file_path,
                    mapping=mapping_obj,
                    output_dir=output_dir,
                )
            except (PPTBuilderError, ExcelReaderError) as exc:
                item.error = str(exc)
                logger.error("Fallo %s: %s", file_path.name, exc)
            except Exception as exc:  # noqa: BLE001 - robustez batch
                item.error = f"Error inesperado: {exc}"
                logger.exception(
                    "Error inesperado procesando %s", file_path.name
                )
            finally:
                item.duration_s = time.perf_counter() - started

            batch.items.append(item)

            if item.ok:
                assert item.build_result is not None
                logger.info(
                    "OK %s -> %s (%s graficos, %.2fs)",
                    file_path.name,
                    item.build_result.output_path.name,
                    item.build_result.charts_updated,
                    item.duration_s,
                )
            else:
                logger.error(
                    "KO %s (%.2fs): %s",
                    file_path.name,
                    item.duration_s,
                    item.error or "errores durante build",
                )

        logger.info(
            "Resumen: total=%s ok=%s ko=%s",
            len(batch.items),
            len(batch.successful),
            len(batch.failed),
        )

    return batch


__all__ = [
    "BatchItemResult",
    "BatchResult",
    "process_batch",
    "write_batch_report_json",
]
