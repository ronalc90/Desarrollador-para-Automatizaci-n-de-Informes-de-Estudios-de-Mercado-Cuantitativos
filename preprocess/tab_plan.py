"""Modelo y loaders del Tab Plan.

Un Tab Plan describe que tablas cruzadas hay que producir a partir
de los datos crudos. Por ejemplo::

    version: 1
    crosses:
      - name: "Satisfaccion por segmento"
        rows: ["satisfaccion"]
        columns: ["segmento"]
        values: "respondent_id"
        aggregate: "count"
        percentage: "column"
      - name: "NPS por edad"
        rows: ["grupo_edad"]
        columns: ["nps_categoria"]

Cada cross produce una tabla pivot que despues se escribe como una hoja
del xlsx de tablas cruzadas, con el nombre de cross como nombre de hoja.

El Tab Plan se puede cargar desde:

- YAML estructurado (``load_tab_plan_yaml``).
- Excel con una hoja de "instrucciones" (``load_tab_plan_excel``).
- Texto libre semi-estructurado via ``preprocess.llm_interpreter``.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Optional, Union

import pandas as pd
import yaml


# Valores permitidos para las columnas opcionales.
ALLOWED_AGGREGATES = {"count", "sum", "mean", "median", "min", "max"}
ALLOWED_PERCENTAGE = {"none", "row", "column", "total"}


class TabPlanError(Exception):
    """Error base del modulo tab_plan."""


# ---------------------------------------------------------------------- #
# Modelos                                                                 #
# ---------------------------------------------------------------------- #


@dataclass
class CrossSpec:
    """Especificacion de una tabla cruzada individual."""

    name: str
    rows: list[str]
    columns: list[str] = field(default_factory=list)
    values: Optional[str] = None
    aggregate: str = "count"
    percentage: str = "none"
    filter: Optional[str] = None  # expresion pandas.query()

    def validate(self) -> None:
        if not self.name:
            raise TabPlanError("Cada cross debe tener un 'name'.")
        if not self.rows:
            raise TabPlanError(
                f"Cross '{self.name}': debe tener al menos una columna en 'rows'."
            )
        if self.aggregate not in ALLOWED_AGGREGATES:
            raise TabPlanError(
                f"Cross '{self.name}': aggregate '{self.aggregate}' no es valido. "
                f"Validos: {sorted(ALLOWED_AGGREGATES)}"
            )
        if self.percentage not in ALLOWED_PERCENTAGE:
            raise TabPlanError(
                f"Cross '{self.name}': percentage '{self.percentage}' no es valido. "
                f"Validos: {sorted(ALLOWED_PERCENTAGE)}"
            )
        if self.aggregate != "count" and self.values is None:
            raise TabPlanError(
                f"Cross '{self.name}': aggregate '{self.aggregate}' "
                "requiere 'values'."
            )


@dataclass
class TabPlan:
    """Tab Plan completo."""

    version: int
    crosses: list[CrossSpec]
    metadata: dict[str, Any] = field(default_factory=dict)

    def validate(self) -> None:
        if not self.crosses:
            raise TabPlanError("El Tab Plan debe tener al menos un cross.")
        names = set()
        for cross in self.crosses:
            cross.validate()
            if cross.name in names:
                raise TabPlanError(
                    f"Nombre de cross duplicado: {cross.name!r}"
                )
            names.add(cross.name)

    def sheet_names(self) -> list[str]:
        """Nombres de hoja que produciria este plan al escribir a xlsx."""
        return [_safe_sheet_name(c.name) for c in self.crosses]

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "TabPlan":
        if not isinstance(data, dict):
            raise TabPlanError("El Tab Plan debe ser un diccionario.")
        version = int(data.get("version", 1))
        raw_crosses = data.get("crosses")
        if not isinstance(raw_crosses, list) or not raw_crosses:
            raise TabPlanError(
                "El Tab Plan debe tener una lista 'crosses' no vacia."
            )

        crosses: list[CrossSpec] = []
        for idx, rc in enumerate(raw_crosses):
            if not isinstance(rc, dict):
                raise TabPlanError(f"crosses[{idx}] debe ser un diccionario.")
            try:
                name = str(rc["name"])
                rows = rc["rows"]
            except KeyError as exc:
                raise TabPlanError(
                    f"crosses[{idx}] debe tener 'name' y 'rows'."
                ) from exc

            if isinstance(rows, str):
                rows = [rows]
            if not isinstance(rows, list) or not rows:
                raise TabPlanError(
                    f"crosses[{idx}].rows debe ser una lista no vacia."
                )
            rows_list = [str(r) for r in rows]

            columns = rc.get("columns", []) or []
            if isinstance(columns, str):
                columns = [columns]
            columns_list = [str(c) for c in columns]

            cross = CrossSpec(
                name=name,
                rows=rows_list,
                columns=columns_list,
                values=str(rc["values"]) if rc.get("values") else None,
                aggregate=str(rc.get("aggregate", "count")),
                percentage=str(rc.get("percentage", "none")),
                filter=str(rc["filter"]) if rc.get("filter") else None,
            )
            crosses.append(cross)

        plan = cls(
            version=version,
            crosses=crosses,
            metadata=dict(data.get("metadata") or {}),
        )
        plan.validate()
        return plan


# ---------------------------------------------------------------------- #
# Loaders                                                                 #
# ---------------------------------------------------------------------- #


def load_tab_plan_yaml(path: Union[str, Path]) -> TabPlan:
    """Carga un Tab Plan desde un archivo YAML."""
    path = Path(path)
    if not path.exists():
        raise TabPlanError(f"Tab Plan no encontrado: {path}")
    with path.open("r", encoding="utf-8") as fh:
        try:
            data = yaml.safe_load(fh)
        except yaml.YAMLError as exc:
            raise TabPlanError(f"YAML invalido: {exc}") from exc
    return TabPlan.from_dict(data)


def load_tab_plan_excel(
    path: Union[str, Path], sheet: str = "TabPlan"
) -> TabPlan:
    """Carga un Tab Plan desde una hoja 'TabPlan' de un xlsx.

    La hoja debe tener las columnas::

        name | rows | columns | values | aggregate | percentage | filter

    ``rows`` y ``columns`` pueden contener varios nombres separados por
    coma o ``+``. Los campos vacios se interpretan como defaults.
    """
    path = Path(path)
    if not path.exists():
        raise TabPlanError(f"Archivo no encontrado: {path}")
    try:
        df = pd.read_excel(path, sheet_name=sheet)
    except Exception as exc:
        raise TabPlanError(
            f"No se pudo leer la hoja {sheet!r} de {path}: {exc}"
        ) from exc

    required = {"name", "rows"}
    missing = required - set(str(c).lower() for c in df.columns)
    if missing:
        raise TabPlanError(
            f"La hoja {sheet!r} debe tener las columnas: {sorted(required)}. "
            f"Faltan: {sorted(missing)}"
        )

    def _split(value: Any) -> list[str]:
        if value is None or (isinstance(value, float) and pd.isna(value)):
            return []
        return [
            s.strip()
            for s in re.split(r"[,\+]", str(value))
            if s.strip()
        ]

    def _opt(value: Any, default: str) -> str:
        if value is None or (isinstance(value, float) and pd.isna(value)):
            return default
        return str(value).strip() or default

    raw_crosses: list[dict[str, Any]] = []
    for _, row in df.iterrows():
        row_dict = {str(k).lower(): v for k, v in row.items()}
        cross_dict = {
            "name": str(row_dict.get("name") or "").strip(),
            "rows": _split(row_dict.get("rows")),
            "columns": _split(row_dict.get("columns")),
            "values": (
                str(row_dict["values"]).strip()
                if row_dict.get("values") and not pd.isna(row_dict["values"])
                else None
            ),
            "aggregate": _opt(row_dict.get("aggregate"), "count"),
            "percentage": _opt(row_dict.get("percentage"), "none"),
            "filter": (
                str(row_dict["filter"]).strip()
                if row_dict.get("filter") and not pd.isna(row_dict["filter"])
                else None
            ),
        }
        if not cross_dict["name"]:
            continue
        raw_crosses.append(cross_dict)

    if not raw_crosses:
        raise TabPlanError(
            f"La hoja {sheet!r} no tiene ningun cross definido."
        )

    return TabPlan.from_dict({"version": 1, "crosses": raw_crosses})


def _safe_sheet_name(name: str) -> str:
    """Convierte un nombre arbitrario en un nombre de hoja xlsx valido."""
    # Excel limita nombres a 31 chars y prohibe algunos caracteres.
    forbidden = set(r":\/?*[]")
    cleaned = "".join(c if c not in forbidden else "_" for c in name)
    return cleaned[:31].strip() or "Sheet1"


__all__ = [
    "CrossSpec",
    "TabPlan",
    "TabPlanError",
    "ALLOWED_AGGREGATES",
    "ALLOWED_PERCENTAGE",
    "load_tab_plan_yaml",
    "load_tab_plan_excel",
]
