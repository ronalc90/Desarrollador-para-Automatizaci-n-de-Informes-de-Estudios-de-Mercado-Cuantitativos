# PPT Engine

Motor de generación automática de presentaciones PowerPoint a partir de
archivos Excel con tablas cruzadas y un template `.pptx` de referencia.

Dado un template y uno o varios archivos de datos, el motor actualiza
los gráficos del template con los nuevos valores manteniendo intactos
el diseño visual y el Excel embebido dentro de cada gráfico.

## Características

- Lectura de tablas desde Excel por rango A1 o por nombre de tabla
- Actualización de gráficos preservando colores, fuentes y layout
- Sincronización automática del `.xlsx` embebido en `ppt/embeddings/`
- Procesamiento batch de múltiples archivos con logging y reporte JSON
- Validación fail-fast del mapping, el template y el Excel
- Comando `inspect` para descubrir slides y gráficos del template
- CLI con `typer`, output con `rich`
- 32 tests unitarios y de integración end-to-end

## Instalación

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

O con `make`:

```bash
make venv
make install
```

## Flujo de uso recomendado

### 1. Inspeccionar el template para descubrir los gráficos

Antes de armar el `mapping.yaml` corré `inspect` sobre el template
para ver qué slides contienen gráficos y cómo se llaman:

```bash
python main.py inspect --template templates/base.pptx
```

Salida típica:

```
Template: templates/base.pptx
Total slides: 12
Total graficos: 5

  Slide 1 [Title Slide] - 0 grafico(s)
* Slide 2 [Title and Content] - 1 grafico(s)
    - chart_index=0, chart_name='Grafico satisfaccion', tipo=COLUMN_CLUSTERED
  Slide 3 [Section Header] - 0 grafico(s)
* Slide 5 [Two Content] - 2 grafico(s)
    - chart_index=0, chart_name='Chart 4', tipo=BAR_CLUSTERED
    - chart_index=1, chart_name='Chart 5', tipo=LINE
...
```

Podés generar un stub de `mapping.yaml` listo para editar:

```bash
python main.py inspect -t templates/base.pptx --mapping-stub > config/mapping.yaml
```

### 2. Completar el mapping

Editá `config/mapping.yaml` para asociar cada gráfico a una hoja y
rango del Excel:

```yaml
slides:
  - slide_index: 2
    charts:
      - chart_name: "Grafico satisfaccion"
        excel_sheet: "P1_satisfaccion"
        data_range: "A1:E6"

  - slide_index: 5
    charts:
      - chart_index: 0
        excel_sheet: "P4_recomendacion"
        data_range: "A1:B8"
      - chart_index: 1
        excel_sheet: "P4_recomendacion"
        data_range: "D1:E8"
```

Reglas:

- `slide_index` es **1-based** (el primer slide es `1`).
- `chart_name` tiene prioridad sobre `chart_index`.
- `data_range` incluye la fila de encabezado (se usa como nombres de
  serie del gráfico).
- La primera columna del rango se trata como **categorías**; el resto
  como **series numéricas**.

### 3. Validar

```bash
python main.py validate \
    --template templates/base.pptx \
    --data data/estudio_chile.xlsx \
    --mapping config/mapping.yaml
```

Reporta errores bloqueantes (hoja inexistente, slide fuera de rango,
rango vacío) y warnings informativos antes de procesar.

### 4. Generar un PPT

Archivo único:

```bash
python main.py build \
    --template templates/base.pptx \
    --data data/estudio_chile.xlsx \
    --mapping config/mapping.yaml \
    --output output/
```

Carpeta completa (uno por país/estudio):

```bash
python main.py batch \
    --template templates/base.pptx \
    --data-folder data/ \
    --mapping config/mapping.yaml \
    --output output/ \
    --log-file output/batch.log \
    --report-json output/batch_report.json
```

El `--log-file` guarda todos los mensajes del engine (incluidas las
líneas `OK`/`KO` por archivo) y el `--report-json` deja un reporte
estructurado listo para consumir desde otras herramientas.

## Estructura del proyecto

```
ppt-engine/
├── engine/
│   ├── __init__.py
│   ├── excel_reader.py      # Lectura de Excel con API get_table
│   ├── validator.py         # Validación fail-fast del mapping
│   ├── chart_updater.py     # Actualización de gráfico + xlsx embebido
│   ├── ppt_builder.py       # Orquestador principal build_presentation
│   ├── batch_processor.py   # process_batch + write_batch_report_json
│   └── inspector.py         # inspect_template para descubrir gráficos
├── config/
│   └── mapping.yaml         # Esquema y ejemplo del mapping
├── tests/
│   ├── conftest.py
│   ├── fixtures/
│   │   └── make_fixtures.py  # Genera fixtures reproducibles
│   ├── test_excel_reader.py
│   ├── test_validator.py
│   ├── test_chart_updater.py
│   ├── test_inspector.py
│   ├── test_batch_logging.py
│   └── test_integration.py
├── .github/workflows/ci.yml  # GitHub Actions (matrix 3.10/3.11/3.12)
├── Makefile
├── main.py                   # CLI con typer
├── requirements.txt
└── README.md
```

## Tests

```bash
pytest tests/
# o
make test
```

La suite incluye 32 tests: unitarios por módulo y dos tests de
integración end-to-end que validan que el PPT generado contiene los
datos correctos tanto en el XML del gráfico como en el Excel embebido.

## Troubleshooting

### "El rango 'X' en la hoja 'Y' está vacío o fuera de límites"

- Verificá que el rango incluya la fila de encabezado.
- En Excel, `A1:E6` son 5 columnas y 6 filas (1 header + 5 de datos).
- Usá `python main.py validate ...` para detectar rangos problemáticos.

### "no se encontró un gráfico con nombre 'X'"

- Corré `python main.py inspect -t template.pptx` para ver los nombres
  reales de los gráficos.
- PowerPoint asigna nombres como `Chart 2`, `Chart 3` por defecto.
  Renombrá el gráfico desde PowerPoint (Ctrl+F6 en el panel de
  selección) o usá `chart_index` como alternativa.

### El gráfico no muestra los nuevos datos al abrir en PowerPoint

- El engine actualiza tanto el XML del gráfico como el Excel embebido,
  pero PowerPoint puede mostrar una caché vieja si el archivo se abrió
  antes de regenerarlo. Cerrá el archivo y reabrilo.
- Si el gráfico tiene formato condicional o fórmulas que apuntan a un
  rango específico, actualizalas para que abarquen el nuevo tamaño.

### Los colores del gráfico cambiaron después del update

- `chart.replace_data(...)` preserva estilos a nivel **serie** siempre
  que el número de series no cambie. Si tu tabla nueva tiene más o
  menos columnas que el gráfico original, el estilo de las series
  nuevas será el default del tema.

### "ERROR Archivo Excel no encontrado: ..."

- Las rutas relativas se resuelven desde el directorio donde corrés el
  comando, no desde el directorio donde está el `.py`. Usá rutas
  absolutas o `cd` a la raíz del proyecto antes de ejecutar.

## Desarrollo

Makefile con los comandos más comunes:

```bash
make help          # lista todos los targets
make test          # corre la suite de tests
make lint          # py_compile sobre engine/ y main.py
make fixtures      # regenera los fixtures de prueba
make inspect       # corre inspect sobre el template de fixture
make run-example   # genera un PPT usando los fixtures
make clean         # borra caches, outputs y artefactos
```

CI corre automáticamente con GitHub Actions en cada push a `main` y
en cada PR, usando matrix con Python 3.10, 3.11 y 3.12.

## Licencia

Autor: Ronald.
