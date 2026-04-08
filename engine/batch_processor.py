"""Procesamiento por lotes de varios .xlsx con el mismo template.

Recibe una carpeta con múltiples archivos de datos y genera un .pptx
por cada uno aplicando el mismo template y mapping. Se reporta un
``BatchResult`` con los archivos exitosos, los fallidos y las razones.
"""

from __future__ import annotations

import logging
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional, Union

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


def process_batch(
    template_path: Union[str, Path],
    data_folder: Union[str, Path],
    mapping: Union[str, Path, Mapping],
    output_dir: Union[str, Path],
    *,
    pattern: str = "*.xlsx",
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

    Returns
    -------
    BatchResult
        Resultado agregado con exitosos, fallidos y duración por item.
    """
    template_path = Path(template_path)
    data_folder = Path(data_folder)
    output_dir = Path(output_dir)

    # Cargamos el mapping una sola vez para no releer YAML por archivo.
    if isinstance(mapping, (str, Path)):
        mapping_obj = load_mapping(mapping)
    else:
        mapping_obj = mapping

    files = _discover_excel_files(data_folder, pattern)
    if not files:
        logger.warning(
            "No se encontraron archivos con patrón %s en %s",
            pattern,
            data_folder,
        )

    batch = BatchResult()
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
            logger.exception("Error inesperado procesando %s", file_path.name)
        finally:
            item.duration_s = time.perf_counter() - started

        batch.items.append(item)

    return batch


__all__ = ["BatchItemResult", "BatchResult", "process_batch"]
