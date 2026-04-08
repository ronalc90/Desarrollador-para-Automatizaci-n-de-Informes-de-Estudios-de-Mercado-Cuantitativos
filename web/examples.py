"""Catalogo de ejemplos pre-cargados para la interfaz web.

Apunta a los fixtures de ``tests/fixtures/`` para no duplicar archivos.
Cada ejemplo expone:

- ``id``         identificador unico (slug)
- ``name``       nombre legible
- ``description`` descripcion corta para mostrar en la UI
- ``files``      dict {kind -> Path absoluta}
- ``stats``      dict con metadata adicional para la UI

``kind`` puede ser ``template``, ``data`` o ``mapping``.
"""

from __future__ import annotations

from pathlib import Path
from typing import Optional

# La carpeta de fixtures se resuelve relativa a la raiz del repo,
# no relativa al cwd, para que funcione independientemente de donde
# se levante el server.
_REPO_ROOT = Path(__file__).resolve().parent.parent
_FIXTURES = _REPO_ROOT / "tests" / "fixtures"


def _ensure_fixtures_present() -> None:
    """Garantiza que los fixtures de los ejemplos existan en disco.

    Los fixtures estan en .gitignore (se regeneran via make_fixtures.py),
    asi que en un clone limpio no existen al levantar el server. Esta
    funcion los genera de forma idempotente al importar el modulo, para
    que el catalogo de ejemplos funcione siempre.
    """
    try:
        from tests.fixtures.make_fixtures import ensure_fixtures
    except Exception:  # pragma: no cover - solo defensa
        return
    try:
        ensure_fixtures(_FIXTURES)
    except Exception:  # pragma: no cover - solo defensa
        # Si la generacion falla, los endpoints van a devolver 503 con
        # un mensaje claro. Mejor no romper el import del web.
        pass


_ensure_fixtures_present()


EXAMPLES: dict[str, dict] = {
    "basico": {
        "id": "basico",
        "name": "Básico — un gráfico",
        "description": (
            "Template con un único gráfico de columnas en el slide 2. "
            "Ideal para entender el flujo mínimo: template + datos + mapping."
        ),
        "files": {
            "template": _FIXTURES / "sample_template.pptx",
            "data": _FIXTURES / "sample_data.xlsx",
            "mapping": _FIXTURES / "sample_mapping.yaml",
        },
        "stats": {
            "slides": 3,
            "charts": 1,
            "level": "Principiante",
        },
    },
    "multi": {
        "id": "multi",
        "name": "Multi-gráfico — 3 charts en 2 slides",
        "description": (
            "Demuestra matching por chart_name en el slide 2 y por "
            "chart_index (0 y 1) en el slide 3 con dos rangos lado a lado."
        ),
        "files": {
            "template": _FIXTURES / "sample_template.pptx",
            "data": _FIXTURES / "sample_data.xlsx",
            "mapping": _FIXTURES / "sample_mapping_multi.yaml",
        },
        "stats": {
            "slides": 3,
            "charts": 3,
            "level": "Intermedio",
        },
    },
    "con-error": {
        "id": "con-error",
        "name": "Con error intencional — para ver la validación",
        "description": (
            "Mapping a propósito mal armado: referencia un slide_index "
            "fuera de rango y una hoja Excel inexistente. Sirve para ver "
            "cómo el motor reporta errores antes de generar nada."
        ),
        "files": {
            "template": _FIXTURES / "sample_template.pptx",
            "data": _FIXTURES / "sample_data.xlsx",
            "mapping": _FIXTURES / "sample_mapping_broken.yaml",
        },
        "stats": {
            "slides": 3,
            "charts": 0,
            "level": "Demo de errores",
        },
    },
}


# Filename "amigable" que se le muestra al usuario en la descarga.
_FILENAMES = {
    "template": "template.pptx",
    "data": "datos.xlsx",
    "mapping": "mapping.yaml",
}


def get_example(example_id: str) -> dict:
    """Devuelve el ejemplo o lanza KeyError."""
    if example_id not in EXAMPLES:
        raise KeyError(example_id)
    return EXAMPLES[example_id]


def get_example_file(example_id: str, kind: str) -> Path:
    """Devuelve la ruta absoluta del archivo del ejemplo.

    Lanza ``KeyError`` si el ejemplo o el kind no existen, y
    ``FileNotFoundError`` si el archivo no esta en disco.
    """
    example = get_example(example_id)
    if kind not in example["files"]:
        raise KeyError(kind)
    path = example["files"][kind]
    if not path.exists():
        raise FileNotFoundError(str(path))
    return path


def filename_for(kind: str) -> str:
    """Nombre legible para descargas."""
    return _FILENAMES.get(kind, kind)


def list_examples() -> list[dict]:
    """Devuelve la lista para el endpoint /api/examples.

    Cada ejemplo incluye, ademas de la metadata, el tamano de cada
    archivo en bytes para mostrarlo en la UI sin tener que pegarle al
    endpoint de descarga.
    """
    out = []
    for ex in EXAMPLES.values():
        files_info = {}
        for kind, path in ex["files"].items():
            try:
                size = path.stat().st_size if path.exists() else None
            except OSError:
                size = None
            files_info[kind] = {
                "filename": filename_for(kind),
                "size_bytes": size,
                "available": path.exists(),
            }
        out.append({
            "id": ex["id"],
            "name": ex["name"],
            "description": ex["description"],
            "stats": ex["stats"],
            "files": files_info,
        })
    return out


def resolve_example_files(example_id: Optional[str]) -> Optional[dict[str, Path]]:
    """Helper usado por los endpoints para resolver los archivos de un ejemplo.

    Devuelve un dict ``{kind: Path}`` o ``None`` si ``example_id`` es None.
    Lanza ``KeyError`` si el id no existe.
    """
    if not example_id:
        return None
    example = get_example(example_id)
    return dict(example["files"])
