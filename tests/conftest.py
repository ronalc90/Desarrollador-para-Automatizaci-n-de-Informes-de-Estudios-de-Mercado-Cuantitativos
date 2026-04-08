"""Configuración común de pytest.

- Agrega la raíz del proyecto al ``sys.path`` para poder importar
  ``engine`` sin tener que instalar el paquete.
- Expone un fixture ``fixtures_dir`` con la ruta absoluta a la carpeta
  de fixtures y se asegura de generar los archivos si no existen.
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from tests.fixtures.make_fixtures import ensure_fixtures  # noqa: E402


@pytest.fixture(scope="session")
def fixtures_dir() -> Path:
    fixtures = Path(__file__).parent / "fixtures"
    ensure_fixtures(fixtures)
    return fixtures


@pytest.fixture(scope="session")
def sample_xlsx(fixtures_dir: Path) -> Path:
    return fixtures_dir / "sample_data.xlsx"


@pytest.fixture(scope="session")
def sample_pptx(fixtures_dir: Path) -> Path:
    return fixtures_dir / "sample_template.pptx"


@pytest.fixture(scope="session")
def sample_mapping(fixtures_dir: Path) -> Path:
    return fixtures_dir / "sample_mapping.yaml"
