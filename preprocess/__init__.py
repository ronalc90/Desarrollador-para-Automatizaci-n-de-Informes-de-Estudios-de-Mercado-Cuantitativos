"""Pipeline de procesamiento previo (Etapa 3).

Toma archivos de respuestas crudos (.sav/.dta/.csv/.xlsx/.dat) y un
Tab Plan con la logica de cruces, y produce un xlsx de tablas
cruzadas listo para alimentar el motor de generacion de PPT (Etapa 1).

Modulos principales:

- responses_reader: lector unificado de formatos de respuestas.
- tab_plan: modelo y loaders del Tab Plan (YAML o Excel).
- crosstabs: motor de tabulaciones cruzadas basado en pandas.
- llm_interpreter: interprete pluggable de Tab Plans en lenguaje libre.
"""

__version__ = "0.1.0"
