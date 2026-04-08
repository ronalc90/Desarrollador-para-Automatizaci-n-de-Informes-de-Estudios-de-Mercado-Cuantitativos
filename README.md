# PPT Engine

Motor de generación automática de presentaciones PowerPoint a partir de
archivos Excel con tablas cruzadas y un template `.pptx` de referencia.

Dado un template y uno o varios archivos de datos, el motor actualiza los
gráficos del template con los nuevos valores, manteniendo intactos el
diseño visual y el Excel embebido dentro de cada gráfico.

## Características

- Lectura de archivos Excel con múltiples hojas y tablas
- Actualización de gráficos preservando el diseño del template
- Actualización consistente del Excel embebido dentro del `.pptx`
- Procesamiento batch de múltiples archivos
- Validación fail-fast antes de ejecutar el pipeline
- CLI simple con modo single y modo batch

## Instalación

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Uso

Modo archivo único:

```bash
python main.py build \
    --template templates/base.pptx \
    --data data/estudio_chile.xlsx \
    --mapping config/mapping.yaml \
    --output output/
```

Modo batch (una carpeta con varios `.xlsx`):

```bash
python main.py batch \
    --template templates/base.pptx \
    --data-folder data/ \
    --mapping config/mapping.yaml \
    --output output/
```

Solo validar sin generar archivos:

```bash
python main.py validate \
    --template templates/base.pptx \
    --data data/estudio_chile.xlsx \
    --mapping config/mapping.yaml
```

## Estructura del proyecto

```
ppt-engine/
├── engine/
│   ├── __init__.py
│   ├── excel_reader.py
│   ├── ppt_builder.py
│   ├── chart_updater.py
│   ├── batch_processor.py
│   └── validator.py
├── config/
│   └── mapping.yaml
├── tests/
│   ├── test_excel_reader.py
│   ├── test_chart_updater.py
│   └── fixtures/
├── main.py
├── requirements.txt
└── README.md
```

## Configuración del mapping

El archivo `config/mapping.yaml` define qué tabla del Excel alimenta a
qué gráfico del template. Ver `config/mapping.yaml` para un ejemplo
completo.

## Tests

```bash
pytest tests/
```
