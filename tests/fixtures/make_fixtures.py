"""Genera fixtures mínimos para los tests.

Crea programáticamente:
- ``sample_data.xlsx`` con dos hojas y dos tablas.
- ``sample_template.pptx`` con un slide que contiene un gráfico de
  columnas cuyo nombre es ``"Grafico satisfaccion"``.
- ``sample_mapping.yaml`` consistente con los dos archivos anteriores.

Se exportan las funciones ``ensure_fixtures`` y ``build_sample_xlsx``
para reutilizarlas desde los tests cuando sea necesario.
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.util import Inches


# ---------------------------------------------------------------------- #
# Datos canónicos que usan los tests                                     #
# ---------------------------------------------------------------------- #

SATISFACCION_HEADER = ["Dimension", "2022", "2023", "2024", "2025"]
SATISFACCION_ROWS = [
    ["Producto", 70, 72, 75, 78],
    ["Precio", 60, 62, 65, 66],
    ["Atencion", 80, 83, 85, 87],
    ["Entrega", 55, 58, 60, 64],
    ["Post-venta", 50, 53, 57, 60],
]

NPS_HEADER = ["Segmento", "Promotores", "Detractores"]
NPS_ROWS = [
    ["General", 55, 20],
    ["Jovenes", 60, 15],
    ["Adultos", 50, 25],
]


def build_sample_xlsx(path: Path) -> Path:
    """Crea el Excel de prueba con dos hojas y tablas conocidas."""
    wb = Workbook()

    # Hoja 1: satisfacción.
    ws1 = wb.active
    ws1.title = "P1_satisfaccion"
    ws1.append(SATISFACCION_HEADER)
    for row in SATISFACCION_ROWS:
        ws1.append(row)

    # Hoja 2: NPS.
    ws2 = wb.create_sheet("P3_nps")
    ws2.append(NPS_HEADER)
    for row in NPS_ROWS:
        ws2.append(row)

    # Hoja 3 vacía para probar validaciones.
    wb.create_sheet("P99_vacia")

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


def build_sample_pptx(path: Path) -> Path:
    """Crea un template mínimo con un gráfico de columnas.

    El slide 1 es una portada vacía.
    El slide 2 contiene un gráfico con nombre conocido.
    """
    prs = Presentation()

    # Slide 1: portada sin gráficos.
    portada_layout = prs.slide_layouts[0]
    slide_portada = prs.slides.add_slide(portada_layout)
    if slide_portada.shapes.title:
        slide_portada.shapes.title.text = "Estudio de ejemplo"

    # Slide 2: slide con gráfico.
    blank_layout = prs.slide_layouts[5]
    slide2 = prs.slides.add_slide(blank_layout)

    chart_data = CategoryChartData()
    chart_data.categories = ["A", "B", "C"]
    chart_data.add_series("Serie inicial", (1, 2, 3))

    chart_shape = slide2.shapes.add_chart(
        XL_CHART_TYPE.COLUMN_CLUSTERED,
        Inches(1),
        Inches(1.5),
        Inches(8),
        Inches(4.5),
        chart_data,
    )
    # El nombre de la shape permite identificar el gráfico desde el mapping.
    chart_shape.name = "Grafico satisfaccion"

    path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(path)
    return path


def build_sample_mapping(path: Path) -> Path:
    """Crea un mapping consistente con los fixtures de Excel y pptx."""
    content = (
        "slides:\n"
        "  - slide_index: 2\n"
        "    charts:\n"
        "      - chart_name: \"Grafico satisfaccion\"\n"
        "        excel_sheet: \"P1_satisfaccion\"\n"
        "        data_range: \"A1:E6\"\n"
    )
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8")
    return path


def ensure_fixtures(fixtures_dir: Path) -> None:
    """Se asegura de que existan los tres fixtures, creándolos si faltan."""
    xlsx = fixtures_dir / "sample_data.xlsx"
    pptx = fixtures_dir / "sample_template.pptx"
    mapping = fixtures_dir / "sample_mapping.yaml"

    if not xlsx.exists():
        build_sample_xlsx(xlsx)
    if not pptx.exists():
        build_sample_pptx(pptx)
    if not mapping.exists():
        build_sample_mapping(mapping)


if __name__ == "__main__":
    target = Path(__file__).parent
    ensure_fixtures(target)
    print(f"Fixtures generados en {target}")
