# Makefile del proyecto ppt-engine
#
# Comandos comunes durante el desarrollo. Todas las tareas usan el
# entorno virtual local ./.venv si existe; si no, cae al python del
# sistema. Sobrescribi PYTHON=... al invocar make para forzar un
# interprete especifico (por ejemplo en CI).

PYTHON ?= $(shell if [ -x .venv/bin/python ]; then echo .venv/bin/python; else echo python3; fi)
PIP    := $(PYTHON) -m pip

.DEFAULT_GOAL := help

.PHONY: help venv install test lint clean fixtures inspect run-example

help: ## Lista los comandos disponibles
	@grep -E '^[a-zA-Z_-]+:.*?## .*$$' $(MAKEFILE_LIST) | \
		awk 'BEGIN {FS = ":.*?## "}; {printf "  %-15s %s\n", $$1, $$2}'

venv: ## Crea un entorno virtual local en .venv
	python3 -m venv .venv
	.venv/bin/pip install --upgrade pip

install: ## Instala las dependencias en el venv activo
	$(PIP) install -r requirements.txt

test: ## Corre toda la suite de tests
	$(PYTHON) -m pytest tests/ -q

lint: ## Verificacion rapida de sintaxis de todos los modulos
	$(PYTHON) -m py_compile engine/*.py main.py

fixtures: ## Regenera los fixtures de prueba
	$(PYTHON) tests/fixtures/make_fixtures.py

inspect: ## Inspecciona el fixture de template (demo del comando inspect)
	$(PYTHON) main.py inspect -t tests/fixtures/sample_template.pptx

run-example: fixtures ## Genera un PPT de ejemplo sobre los fixtures
	$(PYTHON) main.py build \
		-t tests/fixtures/sample_template.pptx \
		-d tests/fixtures/sample_data.xlsx \
		-m tests/fixtures/sample_mapping.yaml \
		-o output/

clean: ## Limpia artefactos de build, cache y output
	rm -rf .pytest_cache .coverage htmlcov build dist *.egg-info
	rm -rf output/
	find . -type d -name "__pycache__" -exec rm -rf {} + 2>/dev/null || true
