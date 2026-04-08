"""CLI del motor de generación automática de presentaciones.

Uso::

    python main.py build \\
        --template template.pptx \\
        --data data.xlsx \\
        --mapping config/mapping.yaml \\
        --output output/

    python main.py batch \\
        --template template.pptx \\
        --data-folder data/ \\
        --mapping config/mapping.yaml \\
        --output output/

    python main.py validate \\
        --template template.pptx \\
        --data data.xlsx \\
        --mapping config/mapping.yaml
"""

from __future__ import annotations

import logging
import sys
from pathlib import Path
from typing import Optional

import typer
from rich.console import Console
from rich.logging import RichHandler

from engine.batch_processor import process_batch, write_batch_report_json
from engine.excel_reader import ExcelReaderError
from engine.inspector import inspect_template
from engine.ppt_builder import PPTBuilderError, build_presentation
from engine.validator import validate_all

app = typer.Typer(
    add_completion=False,
    help="Motor de generacion automatica de PPT a partir de Excel.",
)
console = Console()


def _configure_logging(verbose: bool) -> None:
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(message)s",
        datefmt="[%X]",
        handlers=[RichHandler(console=console, show_path=False)],
    )


@app.command()
def build(
    template: Path = typer.Option(
        ..., "--template", "-t", exists=True, help="Template .pptx."
    ),
    data: Path = typer.Option(
        ..., "--data", "-d", exists=True, help="Archivo Excel .xlsx de datos."
    ),
    mapping: Path = typer.Option(
        Path("config/mapping.yaml"),
        "--mapping",
        "-m",
        exists=True,
        help="YAML con el mapping slide -> chart -> hoja -> rango.",
    ),
    output: Path = typer.Option(
        Path("output"),
        "--output",
        "-o",
        help="Directorio de salida (se crea si no existe).",
    ),
    name: Optional[str] = typer.Option(
        None,
        "--name",
        "-n",
        help="Nombre del .pptx de salida. Por defecto se deriva del Excel.",
    ),
    skip_validation: bool = typer.Option(
        False,
        "--skip-validation",
        help="Omite la validacion previa.",
    ),
    verbose: bool = typer.Option(False, "--verbose", "-v"),
) -> None:
    """Genera un PPT a partir de un template, un Excel y un mapping."""
    _configure_logging(verbose)

    if not skip_validation:
        result = validate_all(template, data, mapping)
        if not result.ok:
            console.print("[bold red]Validacion fallida[/bold red]")
            console.print(result.as_report())
            raise typer.Exit(code=2)
        if result.warnings:
            console.print("[yellow]Warnings de validacion[/yellow]")
            console.print(result.as_report())

    try:
        build_result = build_presentation(
            template_path=template,
            excel_path=data,
            mapping=mapping,
            output_dir=output,
            output_name=name,
        )
    except (PPTBuilderError, ExcelReaderError) as exc:
        console.print(f"[bold red]Error:[/bold red] {exc}")
        raise typer.Exit(code=1) from exc

    console.print(build_result.summary())
    if not build_result.ok:
        raise typer.Exit(code=1)


@app.command()
def batch(
    template: Path = typer.Option(
        ..., "--template", "-t", exists=True, help="Template .pptx."
    ),
    data_folder: Path = typer.Option(
        ...,
        "--data-folder",
        "-d",
        exists=True,
        file_okay=False,
        help="Carpeta con archivos .xlsx.",
    ),
    mapping: Path = typer.Option(
        Path("config/mapping.yaml"),
        "--mapping",
        "-m",
        exists=True,
    ),
    output: Path = typer.Option(Path("output"), "--output", "-o"),
    pattern: str = typer.Option("*.xlsx", "--pattern"),
    log_file: Optional[Path] = typer.Option(
        None,
        "--log-file",
        help="Archivo donde se escribira el log detallado del batch.",
    ),
    report_json: Optional[Path] = typer.Option(
        None,
        "--report-json",
        help="Archivo JSON donde se escribira el reporte estructurado.",
    ),
    verbose: bool = typer.Option(False, "--verbose", "-v"),
) -> None:
    """Procesa todos los .xlsx de una carpeta con el mismo template."""
    _configure_logging(verbose)
    result = process_batch(
        template_path=template,
        data_folder=data_folder,
        mapping=mapping,
        output_dir=output,
        pattern=pattern,
        log_file=log_file,
    )
    console.print(result.summary())
    if report_json is not None:
        written = write_batch_report_json(result, report_json)
        console.print(f"[green]Reporte JSON:[/green] {written}")
    if result.failed:
        raise typer.Exit(code=1)


@app.command()
def validate(
    template: Path = typer.Option(..., "--template", "-t", exists=True),
    data: Path = typer.Option(..., "--data", "-d", exists=True),
    mapping: Path = typer.Option(
        Path("config/mapping.yaml"),
        "--mapping",
        "-m",
        exists=True,
    ),
    verbose: bool = typer.Option(False, "--verbose", "-v"),
) -> None:
    """Valida template + Excel + mapping sin generar archivos."""
    _configure_logging(verbose)
    result = validate_all(template, data, mapping)
    console.print(result.as_report())
    if not result.ok:
        raise typer.Exit(code=2)


@app.command()
def inspect(
    template: Path = typer.Option(
        ..., "--template", "-t", exists=True, help="Template .pptx a inspeccionar."
    ),
    mapping_stub: bool = typer.Option(
        False,
        "--mapping-stub",
        help="Imprime un stub de mapping.yaml listo para editar.",
    ),
) -> None:
    """Lista slides y graficos del template para armar el mapping.yaml."""
    report = inspect_template(template)
    if mapping_stub:
        console.print(report.as_mapping_stub())
    else:
        console.print(report.as_text())


def main() -> None:
    app()


if __name__ == "__main__":
    main()
