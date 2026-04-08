# Guía de Usuario — PPT Engine

> Motor de generación automática de presentaciones PowerPoint a partir de
> archivos Excel y un template `.pptx` de referencia.

---

## 1. ¿Qué hace este sistema?

PPT Engine resuelve un problema concreto y repetitivo:

> *"Tengo un PowerPoint con el diseño aprobado, y cada mes/país/estudio
> tengo que regenerarlo cambiando solamente los datos de los gráficos."*

En lugar de copiar y pegar valores a mano dentro de PowerPoint, el motor:

1. Toma un **template `.pptx`** que ya tiene el diseño, los colores, las
   tipografías y los gráficos en su lugar.
2. Toma un **archivo Excel** con los datos nuevos.
3. Toma un **mapping YAML** que dice *"el gráfico X del slide Y se llena
   con el rango Z de la hoja W"*.
4. Genera un **`.pptx` nuevo idéntico al original** pero con los gráficos
   actualizados — preservando colores, fuentes, formato condicional y el
   Excel embebido dentro de cada gráfico.

### Lo que NO hace
- No diseña presentaciones desde cero.
- No genera gráficos en slides que no los tengan.
- No edita texto, títulos ni imágenes — solo datos de gráficos.

### Lo que SÍ garantiza
- El archivo de salida es **byte-compatible** con PowerPoint, Keynote y
  LibreOffice.
- Si algo no encaja (hoja inexistente, slide fuera de rango, rango vacío),
  el motor **falla rápido** y te dice exactamente qué está mal antes de
  generar nada.
- Procesamiento batch: podés tirarle una carpeta entera de Excels y
  generar 30 presentaciones de un saque, con log y reporte JSON.

---

## 2. Conceptos clave

| Concepto | Qué es |
|---|---|
| **Template** | Archivo `.pptx` con el diseño aprobado. Los gráficos ya existen y tienen formato; solo se actualizan los datos. |
| **Datos** | Archivo `.xlsx` con tablas cruzadas (una por pregunta o métrica). |
| **Mapping** | Archivo `.yaml` que conecta cada gráfico del template con un rango del Excel. |
| **Slide index** | Posición del slide en el template, **base 1** (el primero es `1`). |
| **Chart name** | Nombre del gráfico tal como aparece en el panel de selección de PowerPoint. |
| **Data range** | Rango A1 del Excel (incluye la fila de encabezado). |

---

## 3. Las tres operaciones

El sistema tiene **tres operaciones principales**, accesibles tanto desde
la interfaz web como desde la línea de comandos:

### 3.1 Inspeccionar (`Inspect`)
> *"¿Qué slides tiene este template y cómo se llaman los gráficos?"*

Lee el template y devuelve la lista de slides, layouts y gráficos. Sirve
para descubrir los nombres reales de los gráficos antes de armar el
mapping. Además, puede generar un **stub de mapping.yaml** listo para
editar.

### 3.2 Validar (`Validate`)
> *"¿Mi mapping es coherente con el template y el Excel?"*

Verifica fail-fast antes de procesar:
- ¿Existen los slides referenciados?
- ¿Existen los gráficos por nombre/index?
- ¿Existen las hojas del Excel?
- ¿Los rangos tienen datos y el formato esperado?
- Reporta **errores bloqueantes** y **warnings informativos**.

Validar es opcional pero **muy recomendado** antes de generar.

### 3.3 Generar (`Build`)
> *"Dame el .pptx con los datos nuevos."*

Ejecuta la validación, actualiza los gráficos (XML + Excel embebido) y
escribe el archivo final. Devuelve cuántos gráficos se actualizaron y un
link de descarga.

---

## 4. Cómo usar la interfaz web

### 4.1 Levantar el servidor

```bash
# Una sola vez
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt

# Cada vez que quieras correrlo
.venv/bin/python main.py serve --host 127.0.0.1 --port 8765
```

> Por defecto el comando `serve` usa el puerto **8000**. Pasale `--port`
> si querés otro.

Después, abrí en el navegador: **http://127.0.0.1:8765**

### 4.2 Pasos en la UI

1. **Elegí la pestaña** según lo que quieras hacer (Generar / Inspeccionar / Validar).
2. **Subí el template** arrastrando o haciendo click en la zona de drop.
3. **Subí los datos** Excel (no aplica para Inspeccionar).
4. **Subí el mapping** YAML (opcional — si el servidor tiene un default
   configurado, podés omitirlo).
5. *(Opcional)* Ponele un **nombre de salida** al `.pptx`.
6. Click en el botón principal.
7. Mirá el resultado en el panel inferior:
    - **Verde** → todo OK, podés descargar.
    - **Amarillo** → OK con warnings (revisá igual).
    - **Rojo** → falló, expandí los detalles para ver qué pasó.

---

## 5. Cómo usar la línea de comandos

Para flujos automatizados o para procesar múltiples archivos, la CLI es
más práctica que la web.

### 5.1 Inspeccionar el template

```bash
python main.py inspect --template templates/base.pptx
```

Generar un stub de mapping listo para editar:

```bash
python main.py inspect -t templates/base.pptx --mapping-stub > config/mapping.yaml
```

### 5.2 Validar

```bash
python main.py validate \
    --template templates/base.pptx \
    --data data/estudio_chile.xlsx \
    --mapping config/mapping.yaml
```

### 5.3 Generar un PPT (archivo único)

```bash
python main.py build \
    --template templates/base.pptx \
    --data data/estudio_chile.xlsx \
    --mapping config/mapping.yaml \
    --output output/
```

### 5.4 Generar batch (carpeta completa)

```bash
python main.py batch \
    --template templates/base.pptx \
    --data-folder data/ \
    --mapping config/mapping.yaml \
    --output output/ \
    --log-file output/batch.log \
    --report-json output/batch_report.json
```

El batch procesa **un Excel = un PPT**, te deja el log completo y un
reporte JSON estructurado por si querés consumirlo desde otra
herramienta.

---

## 6. Cómo escribir el mapping

El `mapping.yaml` es el corazón del sistema. Estructura:

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

### Reglas
- `slide_index` es **base 1** (el primer slide es `1`, no `0`).
- `chart_name` tiene **prioridad** sobre `chart_index`. Si un nombre no
  matchea, el motor falla — no cae a `chart_index` silenciosamente.
- `data_range` **incluye la fila de encabezado**. Esa primera fila se usa
  como nombres de las series del gráfico.
- La **primera columna** del rango se interpreta como **categorías**; el
  resto como **series numéricas**.
- Ej: `A1:E6` → 5 columnas (1 de categoría + 4 series) y 6 filas (1 header + 5 datos).

### Tip
Antes de armarlo a mano, generá un stub con:

```bash
python main.py inspect -t templates/base.pptx --mapping-stub > config/mapping.yaml
```

Vas a obtener un YAML pre-poblado con todos los gráficos detectados, y
solo tenés que editar `excel_sheet` y `data_range` en cada uno.

---

## 7. Pipeline de procesamiento previo (Etapa 3)

Si tus datos no vienen ya en formato de tablas cruzadas sino como un
**dataset crudo de respuestas**, hay una etapa adicional de
preprocesamiento que vive en el módulo `preprocess/`:

- `preprocess/responses_reader.py` → lee el dataset de respuestas.
- `preprocess/crosstabs.py` → arma las tablas cruzadas (preguntas × cortes).
- `preprocess/tab_plan.py` → aplica un *Tab Plan* declarativo que define
  qué cortes correr para cada pregunta.
- `preprocess/llm_interpreter.py` → opcionalmente usa un LLM para
  interpretar el Tab Plan en lenguaje natural y traducirlo a la
  estructura formal.

El output de este pipeline es un `.xlsx` con las tablas cruzadas listas
para alimentar al motor principal (paso 5.3 / 5.4).

> Nota: la pipeline de preprocesamiento todavía no está expuesta en la
> interfaz web; se opera por código o por scripts.

---

## 8. Troubleshooting

### "El rango 'X' en la hoja 'Y' está vacío o fuera de límites"
- Verificá que el rango incluya la fila de encabezado.
- En Excel, `A1:E6` son 5 columnas y 6 filas (1 header + 5 de datos).
- Corré `validate` para detectar rangos problemáticos antes de generar.

### "no se encontró un gráfico con nombre 'X'"
- Corré `inspect` para ver los nombres reales.
- PowerPoint asigna nombres genéricos como `Chart 2`, `Chart 3` por
  defecto. Renombralos desde el panel de selección (Ctrl+F6 en PowerPoint)
  o usá `chart_index` como alternativa.

### El gráfico no muestra los nuevos datos al abrirlo
- El motor actualiza tanto el XML del gráfico como el Excel embebido,
  pero PowerPoint puede mostrar caché vieja. Cerrá y reabrí el archivo.
- Si el gráfico tiene fórmulas que apuntan a un rango fijo, ampliá esos
  rangos para que cubran el tamaño nuevo.

### Los colores cambiaron después del update
- `chart.replace_data(...)` preserva estilos a nivel **serie** siempre
  que el número de series no cambie. Si tu tabla nueva tiene más o menos
  columnas que el gráfico original, las series nuevas usan el default
  del tema.

### "ERROR Archivo Excel no encontrado: ..."
- Las rutas relativas se resuelven desde el directorio donde corrés el
  comando, no desde donde está el `.py`. Usá rutas absolutas o `cd` a la
  raíz del proyecto antes de ejecutar.

### El servidor web no responde
- Verificá que el puerto sea el correcto (default es **8000**, no 8765).
- Mirá el log: `.venv/bin/python main.py serve --port 8765` debería
  imprimir `Uvicorn running on http://127.0.0.1:8765`.
- ¿Otro proceso usando el mismo puerto? `lsof -iTCP:8765 -sTCP:LISTEN`.

---

## 9. Endpoints de la API

Todos bajo `http://127.0.0.1:<port>`:

| Método | Endpoint | Qué hace |
|---|---|---|
| `GET`  | `/` | UI HTML |
| `GET`  | `/api/health` | Health check |
| `POST` | `/api/inspect` | Inspecciona un template |
| `POST` | `/api/validate` | Valida template + datos + mapping |
| `POST` | `/api/jobs` | Genera un PPT |
| `GET`  | `/api/jobs/{id}` | Estado de un job |
| `GET`  | `/api/jobs/{id}/download` | Descarga el `.pptx` generado |
| `GET`  | `/docs` | Swagger UI generada por FastAPI |

---

## 10. Estructura del proyecto

```
ppt-engine/
├── engine/                  # Motor principal
│   ├── excel_reader.py      # Lectura de tablas Excel por rango
│   ├── validator.py         # Validación fail-fast
│   ├── chart_updater.py     # Update de gráficos + Excel embebido
│   ├── ppt_builder.py       # Orquestador build_presentation
│   ├── batch_processor.py   # Procesamiento batch + reporte JSON
│   └── inspector.py         # inspect_template
├── preprocess/              # Etapa 3: pipeline de procesamiento previo
│   ├── responses_reader.py
│   ├── crosstabs.py
│   ├── tab_plan.py
│   └── llm_interpreter.py
├── web/
│   └── app.py               # Backend FastAPI + UI
├── config/
│   └── mapping.yaml         # Esquema y ejemplo del mapping
├── tests/                   # 32 tests unitarios + integración
├── main.py                  # CLI con typer
├── requirements.txt
├── README.md                # Doc técnica
└── GUIA_USUARIO.md          # Este archivo
```

---

**Autor:** Ronald
