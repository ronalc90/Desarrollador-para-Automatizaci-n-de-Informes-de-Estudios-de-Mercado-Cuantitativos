"""Interprete pluggable de Tab Plans en lenguaje libre.

El brief menciona explicitamente que en Etapa 3 "se valorara
especialmente experiencia en IA / LLMs para interpretar instrucciones
semi-estructuradas y convertirlas en logica de procesamiento".

Este modulo define una interfaz abstracta ``LLMInterpreter`` que puede
ser implementada por distintos proveedores (Anthropic, OpenAI, etc.).
La implementacion por defecto ``MockLLMInterpreter`` es determinista
y no llama a ninguna API externa: extrae crosses del texto usando
expresiones regulares. Es lo suficiente para demostrar la integracion
end-to-end y para que los tests sean reproducibles sin claves de API.

Uso::

    from preprocess.llm_interpreter import interpret_tab_plan_text

    plan = interpret_tab_plan_text(
        '''
        1. Cruzar satisfaccion por segmento
        2. NPS por edad
        '''
    )

Para usar un LLM real en produccion, se pasa una instancia de
``AnthropicLLMInterpreter`` (o equivalente) via el parametro
``interpreter``.
"""

from __future__ import annotations

import os
import re
from abc import ABC, abstractmethod
from typing import Optional

from preprocess.tab_plan import TabPlan, TabPlanError


class LLMInterpreter(ABC):
    """Interfaz base para interpretar un Tab Plan en lenguaje libre."""

    @abstractmethod
    def interpret(self, text: str) -> TabPlan:
        """Convierte un texto en un ``TabPlan`` estructurado.

        Debe lanzar ``TabPlanError`` si el texto no es interpretable.
        """


# ---------------------------------------------------------------------- #
# MockLLMInterpreter: implementacion determinista sin API                #
# ---------------------------------------------------------------------- #


class MockLLMInterpreter(LLMInterpreter):
    """Implementacion determinista basada en regex.

    Reconoce patrones comunes como:

    - ``cruzar X por Y``
    - ``X por Y`` (sin el verbo)
    - ``X vs Y``
    - ``X contra Y``

    Cada linea no vacia genera un cross. El primer sustantivo es la
    fila (``rows``), el segundo es la columna (``columns``). Cuenta
    registros por defecto.
    """

    # Prefijos que se descartan antes de aplicar los patrones (numeros
    # de lista, bullets, verbos de accion).
    PREFIX_STRIP = re.compile(
        r"^[\d\.\)\-\s]*"  # 1. / 1) / - / espacios
        r"(?:cruzar|cruzo|crear tabla de|crear|generar|armar)?\s*",
        re.IGNORECASE,
    )

    # Patron principal: "X por Y" / "X vs Y" / "X contra Y".
    PATTERN = re.compile(
        r"^(.+?)\s+(?:por|vs\.?|contra)\s+(.+)$",
        re.IGNORECASE,
    )

    def interpret(self, text: str) -> TabPlan:
        if not text or not text.strip():
            raise TabPlanError("Texto vacio, no se puede interpretar.")

        crosses: list[dict] = []
        for raw_line in text.splitlines():
            line = raw_line.strip().rstrip(".")
            if not line:
                continue
            if line.startswith("#"):  # comentarios
                continue

            # Remover prefijo (numero de lista + verbo) antes de matchear.
            cleaned = self.PREFIX_STRIP.sub("", line).strip()
            match = self.PATTERN.match(cleaned)
            if match is None:
                continue

            row_part = match.group(1).strip()
            col_part = match.group(2).strip()
            row_var = _slug(row_part)
            col_var = _slug(col_part)
            name = f"{row_part} por {col_part}".strip().capitalize()

            crosses.append(
                {
                    "name": name,
                    "rows": [row_var],
                    "columns": [col_var],
                    "aggregate": "count",
                }
            )

        if not crosses:
            raise TabPlanError(
                "No se pudo extraer ningun cross del texto. "
                "Usa 'X por Y' o 'cruzar X por Y'."
            )

        # Deduplicar nombres (por si el texto tenia repeticiones).
        seen: set[str] = set()
        unique: list[dict] = []
        for c in crosses:
            name = c["name"]
            if name in seen:
                continue
            seen.add(name)
            unique.append(c)

        return TabPlan.from_dict({"version": 1, "crosses": unique})


def _slug(name: str) -> str:
    """Convierte 'Nivel de Satisfaccion' en 'nivel_de_satisfaccion'."""
    cleaned = re.sub(r"[^\w\s]", "", name, flags=re.UNICODE)
    cleaned = re.sub(r"\s+", "_", cleaned.strip())
    return cleaned.lower()


# ---------------------------------------------------------------------- #
# AnthropicLLMInterpreter: hook pluggable para un LLM real               #
# ---------------------------------------------------------------------- #


class AnthropicLLMInterpreter(LLMInterpreter):
    """Interprete que usa un modelo de Anthropic via el SDK oficial.

    Solo se puede instanciar si ``ANTHROPIC_API_KEY`` esta seteada en
    el entorno y el paquete ``anthropic`` esta instalado. El objetivo
    de este hook es mostrar como se integraria un LLM real; en los
    tests se usa siempre ``MockLLMInterpreter`` para que sean rapidos
    y reproducibles.
    """

    SYSTEM_PROMPT = (
        "Eres un asistente que convierte instrucciones de analisis de "
        "datos en un Tab Plan estructurado en YAML. "
        "Cada linea del usuario describe una tabla cruzada. "
        "Debes responder UNICAMENTE con un bloque YAML valido que "
        "contenga una clave 'crosses' con una lista de objetos con "
        "'name', 'rows', 'columns' y opcionalmente 'aggregate'."
    )

    def __init__(self, model: str = "claude-3-5-sonnet-latest") -> None:
        self.model = model
        self._client = None

    def _ensure_client(self):
        if self._client is not None:
            return self._client
        try:
            import anthropic  # type: ignore
        except ImportError as exc:  # pragma: no cover
            raise TabPlanError(
                "Se requiere el paquete 'anthropic' para usar "
                "AnthropicLLMInterpreter. Instala con: pip install anthropic"
            ) from exc
        if not os.environ.get("ANTHROPIC_API_KEY"):
            raise TabPlanError(
                "No se encontro ANTHROPIC_API_KEY en el entorno."
            )
        self._client = anthropic.Anthropic()
        return self._client

    def interpret(self, text: str) -> TabPlan:  # pragma: no cover
        client = self._ensure_client()
        msg = client.messages.create(
            model=self.model,
            max_tokens=1024,
            system=self.SYSTEM_PROMPT,
            messages=[{"role": "user", "content": text}],
        )
        yaml_text = _extract_yaml_block(msg.content[0].text)

        import yaml

        data = yaml.safe_load(yaml_text)
        return TabPlan.from_dict(data)


def _extract_yaml_block(text: str) -> str:  # pragma: no cover
    """Extrae el primer bloque YAML de un texto con markdown."""
    match = re.search(r"```(?:yaml|yml)?\s*\n(.*?)```", text, re.DOTALL)
    if match:
        return match.group(1)
    return text


# ---------------------------------------------------------------------- #
# API publica                                                             #
# ---------------------------------------------------------------------- #


def interpret_tab_plan_text(
    text: str,
    interpreter: Optional[LLMInterpreter] = None,
) -> TabPlan:
    """Convierte un texto libre en un ``TabPlan`` estructurado.

    Si no se pasa ``interpreter`` se usa ``MockLLMInterpreter``, que
    es determinista y no depende de APIs externas.
    """
    if interpreter is None:
        interpreter = MockLLMInterpreter()
    return interpreter.interpret(text)


__all__ = [
    "LLMInterpreter",
    "MockLLMInterpreter",
    "AnthropicLLMInterpreter",
    "interpret_tab_plan_text",
]
