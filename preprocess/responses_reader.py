"""Lector unificado de archivos de respuestas.

Soporta los formatos que aparecen tipicamente en estudios de mercado:

- ``.sav`` (SPSS)
- ``.por`` (SPSS portable)
- ``.dta`` (Stata)
- ``.sas7bdat`` (SAS)
- ``.xlsx`` (Excel)
- ``.csv`` (texto delimitado)
- ``.dat`` (texto delimitado; heurística con pandas)

Todos devuelven la misma estructura :class:`ResponsesData` con un
``pandas.DataFrame`` y un diccionario de metadata (labels de variables,
labels de valores, tipos, etc.) cuando el formato lo provee.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Optional, Union

import pandas as pd


class ResponsesReaderError(Exception):
    """Error base del modulo responses_reader."""


class UnsupportedFormatError(ResponsesReaderError):
    """El formato del archivo no esta soportado."""


@dataclass
class ResponsesData:
    """Resultado de cargar un archivo de respuestas."""

    df: pd.DataFrame
    source_path: Path
    format: str
    variable_labels: dict[str, str] = field(default_factory=dict)
    value_labels: dict[str, dict[Any, str]] = field(default_factory=dict)
    notes: list[str] = field(default_factory=list)

    @property
    def columns(self) -> list[str]:
        return list(self.df.columns)

    @property
    def n_rows(self) -> int:
        return int(len(self.df))

    def describe(self) -> str:
        lines = [
            f"Archivo: {self.source_path}",
            f"Formato: {self.format}",
            f"Filas: {self.n_rows}",
            f"Columnas ({len(self.columns)}): {self.columns}",
        ]
        if self.variable_labels:
            lines.append("Labels de variables:")
            for k, v in list(self.variable_labels.items())[:10]:
                lines.append(f"  {k}: {v}")
            if len(self.variable_labels) > 10:
                lines.append(
                    f"  ... ({len(self.variable_labels) - 10} mas)"
                )
        return "\n".join(lines)


# ---------------------------------------------------------------------- #
# Loaders por formato                                                    #
# ---------------------------------------------------------------------- #


def _load_with_pyreadstat(path: Path, format_name: str) -> ResponsesData:
    try:
        import pyreadstat
    except ImportError as exc:  # pragma: no cover
        raise ResponsesReaderError(
            f"Se requiere pyreadstat para leer {format_name}. "
            "Instala con: pip install pyreadstat"
        ) from exc

    readers = {
        "spss": pyreadstat.read_sav,
        "spss_por": pyreadstat.read_por,
        "stata": pyreadstat.read_dta,
        "sas": pyreadstat.read_sas7bdat,
    }
    reader = readers.get(format_name)
    if reader is None:
        raise UnsupportedFormatError(
            f"Formato pyreadstat no soportado: {format_name}"
        )

    df, meta = reader(str(path))
    return ResponsesData(
        df=df,
        source_path=path,
        format=format_name,
        variable_labels=dict(meta.column_names_to_labels or {}),
        value_labels={
            col: dict(labels)
            for col, labels in (meta.variable_value_labels or {}).items()
        },
    )


def _load_xlsx(path: Path, sheet: Optional[str] = None) -> ResponsesData:
    df = pd.read_excel(path, sheet_name=sheet or 0)
    return ResponsesData(
        df=df,
        source_path=path,
        format="xlsx",
        notes=[f"hoja: {sheet or 0}"],
    )


def _load_csv(path: Path, sep: Optional[str] = None) -> ResponsesData:
    if sep is None:
        # Auto-detect comma vs semicolon vs tab.
        with path.open("r", encoding="utf-8", errors="replace") as fh:
            sample = fh.read(4096)
        counts = {d: sample.count(d) for d in [",", ";", "\t", "|"]}
        sep = max(counts, key=counts.get) if any(counts.values()) else ","
    df = pd.read_csv(path, sep=sep)
    return ResponsesData(
        df=df,
        source_path=path,
        format="csv",
        notes=[f"separator: {sep!r}"],
    )


def _load_dat(path: Path) -> ResponsesData:
    """Heurística: intenta pyreadstat, luego cae a pandas CSV."""
    try:
        return _load_with_pyreadstat(path, "spss")
    except Exception:
        pass
    try:
        data = _load_csv(path)
        data.format = "dat (csv-like)"
        return data
    except Exception as exc:
        raise ResponsesReaderError(
            f"No se pudo interpretar {path}: {exc}"
        ) from exc


# ---------------------------------------------------------------------- #
# API publica                                                            #
# ---------------------------------------------------------------------- #


_FORMATS_PYREADSTAT = {
    ".sav": "spss",
    ".por": "spss_por",
    ".dta": "stata",
    ".sas7bdat": "sas",
}


class ResponsesReader:
    """Fachada para cargar archivos de respuestas en formato unificado."""

    @staticmethod
    def load(
        path: Union[str, Path],
        *,
        sheet: Optional[str] = None,
        sep: Optional[str] = None,
    ) -> ResponsesData:
        """Carga un archivo de respuestas autodetectando el formato.

        Parameters
        ----------
        path:
            Ruta al archivo.
        sheet:
            Nombre o indice de hoja (solo para .xlsx).
        sep:
            Separador (solo para .csv/.dat).
        """
        path = Path(path)
        if not path.exists():
            raise ResponsesReaderError(f"Archivo no encontrado: {path}")

        suffix = path.suffix.lower()
        if suffix in _FORMATS_PYREADSTAT:
            return _load_with_pyreadstat(path, _FORMATS_PYREADSTAT[suffix])
        if suffix in {".xlsx", ".xlsm"}:
            return _load_xlsx(path, sheet=sheet)
        if suffix == ".csv":
            return _load_csv(path, sep=sep)
        if suffix == ".dat":
            return _load_dat(path)
        raise UnsupportedFormatError(
            f"Extension no soportada: {suffix}. "
            "Formatos validos: .sav .por .dta .sas7bdat .xlsx .csv .dat"
        )


__all__ = [
    "ResponsesData",
    "ResponsesReader",
    "ResponsesReaderError",
    "UnsupportedFormatError",
]
