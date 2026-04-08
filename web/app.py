"""Backend FastAPI para operar el motor desde una interfaz web.

Expone una API REST minima que envuelve los modulos del motor:

- POST /api/inspect     -> inspecciona un template .pptx
- POST /api/validate    -> valida template + xlsx + mapping
- POST /api/jobs        -> genera un PPT a partir de template + xlsx + mapping
- GET  /api/jobs/{id}   -> estado de un job
- GET  /api/jobs/{id}/download -> descarga el .pptx generado
- GET  /                -> UI HTML minima para operar sin dev

El store de jobs es en memoria (dict). Para uso real con varias
instancias se reemplazaria por Redis o una base de datos, pero para
Etapa 2 alcanza con esto.
"""

from __future__ import annotations

import io
import shutil
import uuid
import zipfile
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, Response

from engine.excel_reader import ExcelReaderError
from engine.inspector import inspect_template
from engine.ppt_builder import BuildResult, PPTBuilderError, build_presentation
from engine.validator import load_mapping, validate_all
from web.examples import (
    filename_for as example_filename_for,
    get_example_file,
    list_examples,
    resolve_example_files,
)


# ---------------------------------------------------------------------- #
# Samples: mapeo logico -> ruta de archivo real en tests/fixtures         #
# ---------------------------------------------------------------------- #


def _samples_dir() -> Path:
    """Carpeta donde viven los fixtures/ejemplos listos para descargar."""
    return Path(__file__).resolve().parent.parent / "tests" / "fixtures"


_MT_PPTX = (
    "application/vnd.openxmlformats-officedocument.presentationml.presentation"
)
_MT_XLSX = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

SAMPLE_FILES: dict[str, tuple[str, str, str]] = {
    # slug -> (filename in fixtures, media type, description)
    "template": (
        "sample_template.pptx",
        _MT_PPTX,
        "Template base con 3 slides y 3 graficos",
    ),
    "data": (
        "sample_data.xlsx",
        _MT_XLSX,
        "Tablas cruzadas listas para Etapa 1",
    ),
    "mapping": (
        "sample_mapping_multi.yaml",
        "text/yaml",
        "Mapping slide -> chart -> tabla",
    ),
    "responses": (
        "sample_responses.csv",
        "text/csv",
        "15 respuestas crudas (Etapa 3)",
    ),
    "tab-plan": (
        "sample_tab_plan.yaml",
        "text/yaml",
        "Plan con 3 crosses (count/mean/sum)",
    ),
    "tab-plan-xlsx": (
        "sample_tab_plan.xlsx",
        _MT_XLSX,
        "Tab Plan embebido en Excel",
    ),
}

# ---------------------------------------------------------------------- #
# Job store y tipos                                                       #
# ---------------------------------------------------------------------- #


@dataclass
class Job:
    """Estado de un job de generacion."""

    job_id: str
    created_at: datetime
    status: str  # "running" | "success" | "error"
    output_path: Optional[Path] = None
    charts_updated: int = 0
    charts_failed: int = 0
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    error_message: Optional[str] = None

    def to_dict(self) -> dict:
        return {
            "job_id": self.job_id,
            "created_at": self.created_at.isoformat(timespec="seconds"),
            "status": self.status,
            "charts_updated": self.charts_updated,
            "charts_failed": self.charts_failed,
            "errors": list(self.errors),
            "warnings": list(self.warnings),
            "error_message": self.error_message,
            "output_available": (
                self.output_path is not None and self.output_path.exists()
            ),
        }


# ---------------------------------------------------------------------- #
# Factory                                                                 #
# ---------------------------------------------------------------------- #


def create_app(
    workdir: Optional[Path] = None,
    default_mapping: Optional[Path] = None,
) -> FastAPI:
    """Crea una instancia de FastAPI con estado aislado.

    Parameters
    ----------
    workdir:
        Directorio donde se guardan uploads, outputs y estado de jobs.
        Si es ``None`` se usa ``./web_workdir``.
    default_mapping:
        Ruta al mapping YAML por defecto cuando el request no envia uno.
    """
    workdir = Path(workdir) if workdir else Path("web_workdir")
    workdir.mkdir(parents=True, exist_ok=True)
    (workdir / "uploads").mkdir(exist_ok=True)
    (workdir / "outputs").mkdir(exist_ok=True)

    jobs: dict[str, Job] = {}

    app = FastAPI(
        title="PPT Engine - Etapa 2",
        version="0.1.0",
        description=(
            "Interfaz web simple para el motor de generacion de PPT. "
            "Operacion sin desarrollador."
        ),
    )

    # ------------------------------------------------------------------ #
    # Helpers                                                            #
    # ------------------------------------------------------------------ #

    def _save_upload(upload: UploadFile, dest: Path) -> Path:
        with dest.open("wb") as fh:
            shutil.copyfileobj(upload.file, fh)
        return dest

    def _resolve_mapping(
        mapping_upload: Optional[UploadFile], job_dir: Path
    ) -> Path:
        if mapping_upload is not None and mapping_upload.filename:
            path = _save_upload(mapping_upload, job_dir / "mapping.yaml")
            return path
        if default_mapping is not None and default_mapping.exists():
            return default_mapping
        raise HTTPException(
            status_code=400,
            detail=(
                "Debe enviarse un archivo mapping.yaml o el servidor debe "
                "tener un mapping default configurado."
            ),
        )

    # ------------------------------------------------------------------ #
    # Health                                                             #
    # ------------------------------------------------------------------ #

    @app.get("/api/health")
    def health() -> dict:
        return {
            "status": "ok",
            "workdir": str(workdir),
            "jobs_in_memory": len(jobs),
            "samples_available": sorted(SAMPLE_FILES.keys()),
        }

    @app.get("/favicon.ico")
    def favicon() -> Response:
        """Favicon SVG inline con el logo P del brand."""
        svg = (
            '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32">'
            '<defs><linearGradient id="g" x1="0%" y1="0%" x2="100%" y2="100%">'
            '<stop offset="0%" stop-color="#6366f1"/>'
            '<stop offset="50%" stop-color="#8b5cf6"/>'
            '<stop offset="100%" stop-color="#ec4899"/>'
            '</linearGradient></defs>'
            '<rect width="32" height="32" rx="7" fill="url(#g)"/>'
            '<text x="16" y="23" font-family="Inter,sans-serif" '
            'font-size="20" font-weight="800" fill="white" '
            'text-anchor="middle">P</text>'
            '</svg>'
        )
        return Response(content=svg, media_type="image/svg+xml")

    # ------------------------------------------------------------------ #
    # Samples (descarga de archivos de ejemplo)                          #
    # ------------------------------------------------------------------ #

    @app.get("/api/samples")
    def samples_zip() -> Response:
        """Devuelve un zip con TODOS los samples listos para probar."""
        buf = io.BytesIO()
        missing: list[str] = []
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for _slug, (filename, _mt, _desc) in SAMPLE_FILES.items():
                path = _samples_dir() / filename
                if not path.exists():
                    missing.append(filename)
                    continue
                zf.write(str(path), arcname=filename)
            zf.writestr("README.txt", _render_samples_readme())
        if missing:
            raise HTTPException(
                status_code=503,
                detail=(
                    "Faltan fixtures: "
                    + ", ".join(missing)
                    + ". Corre `make fixtures`."
                ),
            )
        return Response(
            content=buf.getvalue(),
            media_type="application/zip",
            headers={
                "Content-Disposition": (
                    'attachment; filename="ppt-engine-samples.zip"'
                )
            },
        )

    # ------------------------------------------------------------------ #
    # Examples (sets pre-cargados de template + data + mapping)          #
    # ------------------------------------------------------------------ #

    @app.get("/api/examples")
    def examples_list() -> JSONResponse:
        """Lista los ejemplos pre-cargados disponibles para probar."""
        return JSONResponse({"examples": list_examples()})

    @app.get("/api/examples/{example_id}/files/{kind}")
    def example_download(example_id: str, kind: str) -> FileResponse:
        """Descarga uno de los archivos de un ejemplo (template/data/mapping)."""
        try:
            path = get_example_file(example_id, kind)
        except KeyError:
            raise HTTPException(
                status_code=404,
                detail=f"Ejemplo o tipo no valido: {example_id}/{kind}",
            )
        except FileNotFoundError as exc:
            raise HTTPException(
                status_code=503,
                detail=(
                    f"El archivo del ejemplo no esta disponible: {exc}. "
                    "Corre `make fixtures`."
                ),
            )
        media = {
            "template": _MT_PPTX,
            "data": _MT_XLSX,
            "mapping": "text/yaml",
        }.get(kind, "application/octet-stream")
        return FileResponse(
            str(path), filename=example_filename_for(kind), media_type=media
        )

    # Helpers para resolver inputs cuando viene example_id en lugar de uploads.

    def _resolve_inputs(
        template: Optional[UploadFile],
        data: Optional[UploadFile],
        mapping: Optional[UploadFile],
        example_id: Optional[str],
        job_dir: Path,
        require_data: bool,
    ) -> tuple[Path, Optional[Path], Path]:
        """Resuelve template/data/mapping desde uploads o desde un ejemplo.

        Si ``example_id`` esta presente, se usan los archivos de ese
        ejemplo (los uploads se ignoran). Si no, se exige el upload.
        """
        if example_id:
            try:
                files = resolve_example_files(example_id) or {}
            except KeyError:
                raise HTTPException(
                    status_code=404,
                    detail=f"Ejemplo no encontrado: {example_id}",
                )
            return (
                files["template"],
                files.get("data") if require_data else files.get("data"),
                files["mapping"],
            )

        if template is None or not template.filename:
            raise HTTPException(
                status_code=400,
                detail="Falta el archivo template (o un example_id).",
            )
        template_path = _save_upload(template, job_dir / template.filename)

        data_path: Optional[Path] = None
        if require_data:
            if data is None or not data.filename:
                raise HTTPException(
                    status_code=400,
                    detail="Falta el archivo de datos (o un example_id).",
                )
            data_path = _save_upload(data, job_dir / data.filename)
        elif data is not None and data.filename:
            data_path = _save_upload(data, job_dir / data.filename)

        mapping_path = _resolve_mapping(mapping, job_dir)
        return template_path, data_path, mapping_path

    @app.get("/api/samples/{slug}")
    def download_sample(slug: str) -> FileResponse:
        if slug not in SAMPLE_FILES:
            raise HTTPException(
                status_code=404,
                detail=(
                    f"Sample '{slug}' no existe. "
                    f"Validos: {sorted(SAMPLE_FILES.keys())}"
                ),
            )
        filename, media_type, _desc = SAMPLE_FILES[slug]
        path = _samples_dir() / filename
        if not path.exists():
            raise HTTPException(
                status_code=503,
                detail=(
                    f"El sample '{filename}' no fue generado aun. "
                    "Corre `make fixtures`."
                ),
            )
        return FileResponse(
            str(path), filename=filename, media_type=media_type
        )

    # ------------------------------------------------------------------ #
    # Inspect                                                            #
    # ------------------------------------------------------------------ #

    @app.post("/api/inspect")
    async def inspect_endpoint(
        template: Optional[UploadFile] = File(
            None, description="Template .pptx (omitir si se usa example_id)"
        ),
        example_id: Optional[str] = Form(
            None, description="Id de ejemplo pre-cargado a usar"
        ),
    ) -> JSONResponse:
        job_id = uuid.uuid4().hex
        job_dir = workdir / "uploads" / job_id
        job_dir.mkdir(parents=True, exist_ok=True)
        if example_id:
            try:
                template_path = get_example_file(example_id, "template")
            except KeyError:
                raise HTTPException(
                    status_code=404,
                    detail=f"Ejemplo no encontrado: {example_id}",
                )
        else:
            if template is None or not template.filename:
                raise HTTPException(
                    status_code=400,
                    detail="Debe enviarse template o example_id.",
                )
            if not template.filename.lower().endswith((".pptx", ".pptm")):
                raise HTTPException(
                    status_code=400, detail="El archivo debe ser .pptx/.pptm"
                )
            template_path = _save_upload(template, job_dir / template.filename)
        try:
            report = inspect_template(template_path)
        except (FileNotFoundError, ValueError) as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc

        payload = {
            "template": template.filename,
            "total_slides": report.total_slides,
            "total_charts": report.total_charts,
            "slides": [
                {
                    "slide_index": s.slide_index,
                    "layout_name": s.layout_name,
                    "has_charts": s.has_charts,
                    "charts": [
                        {
                            "chart_index": c.chart_index,
                            "shape_name": c.shape_name,
                            "chart_type": c.chart_type,
                        }
                        for c in s.charts
                    ],
                }
                for s in report.slides
            ],
            "mapping_stub": report.as_mapping_stub(),
        }
        return JSONResponse(payload)

    # ------------------------------------------------------------------ #
    # Validate                                                           #
    # ------------------------------------------------------------------ #

    @app.post("/api/validate")
    async def validate_endpoint(
        template: Optional[UploadFile] = File(None),
        data: Optional[UploadFile] = File(None),
        mapping: Optional[UploadFile] = File(None),
        example_id: Optional[str] = Form(None),
    ) -> JSONResponse:
        job_id = uuid.uuid4().hex
        job_dir = workdir / "uploads" / job_id
        job_dir.mkdir(parents=True, exist_ok=True)

        template_path, data_path, mapping_path = _resolve_inputs(
            template, data, mapping, example_id, job_dir, require_data=True
        )
        assert data_path is not None  # require_data=True garantiza no-None

        result = validate_all(template_path, data_path, mapping_path)
        return JSONResponse(
            {
                "ok": result.ok,
                "errors": list(result.errors),
                "warnings": list(result.warnings),
            }
        )

    # ------------------------------------------------------------------ #
    # Jobs                                                               #
    # ------------------------------------------------------------------ #

    @app.post("/api/jobs")
    async def create_job(
        template: Optional[UploadFile] = File(None),
        data: Optional[UploadFile] = File(None),
        mapping: Optional[UploadFile] = File(None),
        output_name: Optional[str] = Form(None),
        example_id: Optional[str] = Form(None),
    ) -> JSONResponse:
        job_id = uuid.uuid4().hex
        job_dir = workdir / "uploads" / job_id
        job_dir.mkdir(parents=True, exist_ok=True)
        output_dir = workdir / "outputs" / job_id
        output_dir.mkdir(parents=True, exist_ok=True)

        template_path, data_path, mapping_path = _resolve_inputs(
            template, data, mapping, example_id, job_dir, require_data=True
        )
        assert data_path is not None

        job = Job(
            job_id=job_id,
            created_at=datetime.now(),
            status="running",
        )
        jobs[job_id] = job

        # Validacion previa.
        validation = validate_all(template_path, data_path, mapping_path)
        if not validation.ok:
            job.status = "error"
            job.errors = list(validation.errors)
            job.warnings = list(validation.warnings)
            job.error_message = "Validacion fallida"
            return JSONResponse(job.to_dict(), status_code=400)
        job.warnings.extend(validation.warnings)

        try:
            build_result: BuildResult = build_presentation(
                template_path=template_path,
                excel_path=data_path,
                mapping=mapping_path,
                output_dir=output_dir,
                output_name=output_name,
            )
        except (PPTBuilderError, ExcelReaderError) as exc:
            job.status = "error"
            job.error_message = str(exc)
            return JSONResponse(job.to_dict(), status_code=500)

        job.status = "success" if build_result.ok else "error"
        job.output_path = build_result.output_path
        job.charts_updated = build_result.charts_updated
        job.charts_failed = build_result.charts_failed
        job.errors.extend(build_result.errors)
        job.warnings.extend(build_result.warnings)
        if not build_result.ok:
            job.error_message = "Errores durante el build. Ver 'errors'."
        return JSONResponse(job.to_dict())

    @app.get("/api/jobs/{job_id}")
    def get_job(job_id: str) -> JSONResponse:
        job = jobs.get(job_id)
        if job is None:
            raise HTTPException(status_code=404, detail="Job no encontrado")
        return JSONResponse(job.to_dict())

    @app.get("/api/jobs/{job_id}/download")
    def download_job(job_id: str):
        job = jobs.get(job_id)
        if job is None:
            raise HTTPException(status_code=404, detail="Job no encontrado")
        if job.output_path is None or not job.output_path.exists():
            raise HTTPException(
                status_code=404, detail="Output no disponible para este job"
            )
        return FileResponse(
            str(job.output_path),
            filename=job.output_path.name,
            media_type=(
                "application/vnd.openxmlformats-officedocument."
                "presentationml.presentation"
            ),
        )

    # ------------------------------------------------------------------ #
    # Preprocess (Etapa 3)                                               #
    # ------------------------------------------------------------------ #

    @app.post("/api/preprocess")
    async def preprocess_endpoint(
        data: UploadFile = File(..., description="Archivo de respuestas"),
        tab_plan: Optional[UploadFile] = File(
            None, description="Tab Plan YAML (opcional)"
        ),
        tab_plan_text: Optional[str] = Form(
            None,
            description="Texto libre para el LLM interpreter (opcional)",
        ),
    ) -> Response:
        """Genera un xlsx de tablas cruzadas y lo devuelve como binario."""
        from preprocess.crosstabs import run_tab_plan
        from preprocess.llm_interpreter import interpret_tab_plan_text
        from preprocess.responses_reader import (
            ResponsesReader,
            ResponsesReaderError,
        )
        from preprocess.tab_plan import TabPlanError, load_tab_plan_yaml

        if (tab_plan is None or not tab_plan.filename) and not tab_plan_text:
            raise HTTPException(
                status_code=400,
                detail="Debes enviar un archivo tab_plan o texto tab_plan_text.",
            )

        job_id = uuid.uuid4().hex
        job_dir = workdir / "uploads" / job_id
        job_dir.mkdir(parents=True, exist_ok=True)
        data_path = _save_upload(data, job_dir / data.filename)

        try:
            if tab_plan is not None and tab_plan.filename:
                tp_path = _save_upload(tab_plan, job_dir / tab_plan.filename)
                plan = load_tab_plan_yaml(tp_path)
            else:
                assert tab_plan_text is not None
                plan = interpret_tab_plan_text(tab_plan_text)
        except TabPlanError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc

        try:
            responses = ResponsesReader.load(data_path)
        except ResponsesReaderError as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc

        out_path = workdir / "outputs" / job_id / "crosstabs.xlsx"
        out_path.parent.mkdir(parents=True, exist_ok=True)
        result = run_tab_plan(plan, responses, out_path)

        return FileResponse(
            str(result.output_path),
            filename="crosstabs.xlsx",
            media_type=_MT_XLSX,
            headers={
                "X-Tables-Produced": str(result.n_tables),
                "X-Warnings": str(len(result.warnings)),
            },
        )

    # ------------------------------------------------------------------ #
    # UI                                                                  #
    # ------------------------------------------------------------------ #

    @app.get("/", response_class=HTMLResponse)
    def ui_home() -> HTMLResponse:
        return HTMLResponse(_render_home())

    @app.get("/app", response_class=HTMLResponse)
    def ui_app() -> HTMLResponse:
        return HTMLResponse(_render_ui())

    @app.get("/docs-ui", response_class=HTMLResponse)
    def ui_docs() -> HTMLResponse:
        return HTMLResponse(_render_docs())

    return app


# ---------------------------------------------------------------------- #
# UI HTML minima                                                          #
# ---------------------------------------------------------------------- #


def _render_ui() -> str:
    return """<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>PPT Engine — Generador de presentaciones</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">
<style>
  :root {
    --bg: #f5f7fb;
    --surface: #ffffff;
    --surface-soft: #f8fafc;
    --surface-softer: #f1f5f9;
    --border: #e5e7eb;
    --border-strong: #cbd5e1;
    --text: #0f172a;
    --text-soft: #475569;
    --text-mute: #94a3b8;
    --primary: #6366f1;
    --primary-2: #8b5cf6;
    --primary-3: #ec4899;
    --primary-soft: #eef2ff;
    --primary-softer: #f5f3ff;
    --accent-cyan: #06b6d4;
    --accent-emerald: #10b981;
    --success: #059669;
    --success-soft: #ecfdf5;
    --danger: #ef4444;
    --danger-soft: #fef2f2;
    --warning: #f59e0b;
    --warning-soft: #fffbeb;
    --shadow-xl: 0 20px 60px -15px rgba(99, 102, 241, 0.25);
    --shadow-lg: 0 10px 40px -10px rgba(15, 23, 42, 0.12);
    --shadow-md: 0 4px 16px -4px rgba(15, 23, 42, 0.08);
    --shadow-sm: 0 1px 3px rgba(15, 23, 42, 0.06);
    --radius: 16px;
    --radius-sm: 10px;
  }
  * { box-sizing: border-box; }
  html, body { margin: 0; padding: 0; }
  body {
    font-family: 'Inter', system-ui, -apple-system, sans-serif;
    background:
      radial-gradient(800px 400px at 0% 0%, rgba(139, 92, 246, 0.10), transparent 60%),
      radial-gradient(800px 400px at 100% 0%, rgba(6, 182, 212, 0.10), transparent 60%),
      radial-gradient(600px 300px at 50% 100%, rgba(236, 72, 153, 0.06), transparent 60%),
      var(--bg);
    min-height: 100vh;
    color: var(--text);
    -webkit-font-smoothing: antialiased;
    text-rendering: optimizeLegibility;
  }
  .page { max-width: 920px; margin: 0 auto; padding: 2rem 1.25rem 4rem; }

  /* Topnav sticky */
  .topnav {
    position: sticky; top: 0; z-index: 50;
    background: rgba(255,255,255,0.85);
    backdrop-filter: saturate(180%) blur(12px);
    -webkit-backdrop-filter: saturate(180%) blur(12px);
    border-bottom: 1px solid var(--border);
  }
  .topnav-inner {
    display: flex; align-items: center; justify-content: space-between;
    padding: 0.85rem 1.5rem; max-width: 1100px; margin: 0 auto;
    gap: 1rem; flex-wrap: wrap;
  }
  .topnav-brand {
    display: flex; align-items: center; gap: 0.65rem;
    text-decoration: none; color: var(--text);
    font-weight: 800; font-size: 1.05rem; letter-spacing: -0.01em;
  }
  .topnav-logo {
    width: 36px; height: 36px; border-radius: 10px;
    background: linear-gradient(135deg, #6366f1 0%, #8b5cf6 50%, #ec4899 100%);
    display: grid; place-items: center;
    color: white; box-shadow: 0 10px 24px -8px rgba(139,92,246,0.5);
  }
  .topnav-logo svg { width: 20px; height: 20px; }
  .topnav-title {
    background: linear-gradient(135deg, #6366f1 0%, #8b5cf6 50%, #ec4899 100%);
    -webkit-background-clip: text; background-clip: text; color: transparent;
  }
  .topnav-links { display: flex; gap: 0.25rem; align-items: center; flex-wrap: wrap; }
  .nav-link {
    padding: 0.5rem 0.95rem; border-radius: 8px;
    color: var(--text-soft); font-weight: 600; font-size: 0.88rem;
    text-decoration: none;
    transition: all .15s ease;
  }
  .nav-link:hover { background: var(--surface-softer); color: var(--text); }
  .nav-link.active { background: var(--primary-soft); color: var(--primary); }

  .pill {
    background: white;
    border: 1px solid var(--border);
    color: var(--text); font-size: 0.78rem; font-weight: 600;
    padding: 0.45rem 0.9rem; border-radius: 999px;
    box-shadow: var(--shadow-sm);
    margin-left: 0.4rem;
  }
  .pill .dot {
    display: inline-block; width: 8px; height: 8px; border-radius: 50%;
    background: var(--accent-emerald); margin-right: 7px;
    box-shadow: 0 0 0 3px rgba(16, 185, 129, 0.18);
    animation: pulse 2s ease-in-out infinite;
  }
  @keyframes pulse { 0%,100% { opacity: 1; } 50% { opacity: 0.55; } }

  /* Page intro */
  .page-intro { margin-bottom: 1.5rem; }
  .page-title {
    font-size: 1.9rem; font-weight: 800; letter-spacing: -0.025em;
    margin: 0 0 0.4rem;
    background: linear-gradient(135deg, #0f172a 0%, #475569 100%);
    -webkit-background-clip: text; background-clip: text; color: transparent;
  }
  .page-sub { color: var(--text-soft); font-size: 0.95rem; margin: 0; line-height: 1.55; max-width: 640px; }

  /* Examples section */
  .examples-card {
    background: white;
    border-radius: var(--radius);
    box-shadow: var(--shadow-lg);
    border: 1px solid var(--border);
    overflow: hidden;
    margin-bottom: 1.5rem;
  }
  .examples-head {
    padding: 1.25rem 1.5rem 0.75rem;
    background: linear-gradient(180deg, #fdf4ff 0%, #fafbff 100%);
    border-bottom: 1px solid var(--border);
  }
  .examples-head h2 {
    margin: 0 0 0.3rem; font-size: 1.05rem; font-weight: 700;
    display: flex; align-items: center; gap: 0.55rem;
    color: var(--text);
  }
  .examples-head h2 svg { width: 18px; height: 18px; color: var(--primary); }
  .examples-head p {
    margin: 0 0 1rem; color: var(--text-soft);
    font-size: 0.86rem; line-height: 1.5;
  }
  .examples-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(260px, 1fr));
    gap: 1rem;
    padding: 1.25rem 1.5rem;
  }
  .ex-card {
    background: var(--surface-soft);
    border: 1.5px solid var(--border);
    border-radius: var(--radius-sm);
    padding: 1rem 1.05rem 1.1rem;
    display: flex; flex-direction: column;
    transition: all 0.18s ease;
    position: relative;
  }
  .ex-card:hover {
    border-color: var(--primary);
    transform: translateY(-2px);
    box-shadow: var(--shadow-md);
    background: white;
  }
  .ex-card.selected {
    border-color: var(--primary);
    border-width: 2px;
    background: linear-gradient(180deg, #fafbff 0%, white 100%);
    box-shadow: 0 0 0 4px rgba(99,102,241,0.12), var(--shadow-md);
  }
  .ex-card.selected::after {
    content: "✓";
    position: absolute;
    top: 0.75rem; right: 0.75rem;
    width: 22px; height: 22px;
    background: var(--primary); color: white;
    border-radius: 50%;
    display: grid; place-items: center;
    font-size: 0.78rem; font-weight: 700;
  }
  .ex-name {
    font-weight: 700; font-size: 0.92rem;
    color: var(--text); margin: 0 0 0.35rem;
    padding-right: 1.5rem;
  }
  .ex-desc {
    font-size: 0.79rem; color: var(--text-soft);
    line-height: 1.45; margin: 0 0 0.75rem; flex: 1;
  }
  .ex-stats {
    display: flex; gap: 0.4rem; flex-wrap: wrap;
    margin-bottom: 0.85rem;
  }
  .ex-stat {
    background: white;
    border: 1px solid var(--border);
    color: var(--text-soft);
    font-size: 0.7rem; font-weight: 600;
    padding: 0.22rem 0.55rem; border-radius: 999px;
  }
  .ex-stat.level { background: var(--primary-soft); border-color: #c7d2fe; color: var(--primary); }
  .ex-stat.error { background: var(--danger-soft); border-color: #fecaca; color: var(--danger); }
  .ex-actions { display: flex; gap: 0.4rem; flex-wrap: wrap; }
  .ex-btn {
    flex: 1; min-width: 0;
    border: 0; cursor: pointer;
    font-family: inherit; font-weight: 600; font-size: 0.78rem;
    padding: 0.55rem 0.7rem; border-radius: 7px;
    display: inline-flex; align-items: center; justify-content: center; gap: 0.35rem;
    text-decoration: none;
    transition: all 0.15s ease;
  }
  .ex-btn svg { width: 13px; height: 13px; }
  .ex-btn-use {
    background: linear-gradient(135deg, #6366f1 0%, #8b5cf6 100%);
    color: white;
    box-shadow: 0 4px 12px -4px rgba(99,102,241,0.5);
  }
  .ex-btn-use:hover { transform: translateY(-1px); box-shadow: 0 6px 16px -4px rgba(99,102,241,0.6); }
  .ex-btn-dl {
    background: white; color: var(--text-soft);
    border: 1px solid var(--border);
  }
  .ex-btn-dl:hover { background: var(--surface-softer); color: var(--text); border-color: var(--border-strong); }

  /* Banner ejemplo cargado */
  .example-banner {
    margin: 0 0 1rem;
    padding: 0.85rem 1.1rem;
    background: linear-gradient(135deg, #eef2ff 0%, #faf5ff 100%);
    border: 1.5px solid #c7d2fe;
    border-radius: var(--radius-sm);
    display: flex; align-items: center; justify-content: space-between; gap: 0.75rem;
    font-size: 0.86rem;
  }
  .example-banner-text { color: var(--text); }
  .example-banner-text strong { color: var(--primary); }
  .example-banner-clear {
    border: 0; background: white; cursor: pointer;
    padding: 0.4rem 0.8rem; border-radius: 6px;
    color: var(--text-soft); font-size: 0.78rem; font-weight: 600;
    border: 1px solid var(--border);
    transition: all 0.15s ease;
  }
  .example-banner-clear:hover { color: var(--danger); border-color: var(--danger); }

  /* Card */
  .card {
    background: var(--surface);
    border-radius: var(--radius);
    box-shadow: var(--shadow-xl);
    overflow: hidden;
    border: 1px solid var(--border);
  }
  .card-head {
    padding: 1.75rem 1.75rem 0;
    background: linear-gradient(180deg, #fafbff 0%, white 100%);
  }
  .card-head h2 {
    margin: 0 0 0.35rem; font-size: 1.2rem; font-weight: 700;
    letter-spacing: -0.015em; color: var(--text);
  }
  .card-head p { margin: 0 0 1.25rem; color: var(--text-soft); font-size: 0.9rem; line-height: 1.5; }
  .card-head p code {
    background: var(--surface-softer); padding: 1px 6px;
    border-radius: 4px; font-size: 0.82em;
    font-family: 'JetBrains Mono', ui-monospace, monospace;
    color: var(--primary);
  }
  .card-body { padding: 1.5rem 1.75rem 1.75rem; }

  /* Tabs */
  .tabs {
    display: inline-flex; gap: 4px; padding: 5px;
    background: var(--surface-softer);
    border-radius: 12px;
    margin: 0;
  }
  .tab {
    border: 0; background: transparent; cursor: pointer;
    padding: 0.6rem 1.15rem; border-radius: 9px;
    font-family: inherit; font-size: 0.87rem; font-weight: 600;
    color: var(--text-soft);
    display: inline-flex; align-items: center; gap: 0.5rem;
    transition: all 0.18s ease;
  }
  .tab:hover { color: var(--text); background: rgba(255,255,255,0.6); }
  .tab.active {
    background: white;
    color: var(--primary);
    box-shadow: var(--shadow-sm), 0 0 0 1px rgba(99, 102, 241, 0.12);
  }
  .tab svg { width: 16px; height: 16px; }

  /* Form fields */
  .field { margin-top: 1.15rem; }
  .field:first-child { margin-top: 0.25rem; }
  .field-label {
    display: flex; align-items: baseline; justify-content: space-between;
    margin-bottom: 0.5rem;
  }
  .field-label .name { font-size: 0.84rem; font-weight: 600; color: var(--text); }
  .field-label .hint { font-size: 0.74rem; color: var(--text-mute); font-weight: 500; }
  .required { color: var(--danger); margin-left: 3px; font-weight: 700; }

  /* Drop zone — IMPORTANT: label is inline by default, must be block */
  .dropzone {
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    width: 100%;
    min-height: 110px;
    position: relative;
    border: 1.5px dashed var(--border-strong);
    border-radius: var(--radius-sm);
    background: var(--surface-soft);
    padding: 1.1rem 1rem;
    text-align: center;
    cursor: pointer;
    transition: all 0.18s ease;
  }
  .dropzone:hover {
    border-color: var(--primary);
    background: var(--primary-softer);
    transform: translateY(-1px);
    box-shadow: var(--shadow-md);
  }
  .dropzone.dragover {
    border-color: var(--primary);
    background: var(--primary-soft);
    transform: scale(1.005);
    box-shadow: 0 0 0 4px rgba(99,102,241,0.15);
  }
  .dropzone.has-file {
    border-style: solid;
    border-color: var(--accent-emerald);
    background: linear-gradient(135deg, #ecfdf5 0%, #f0fdfa 100%);
  }
  .dropzone input[type=file] {
    position: absolute; inset: 0; width: 100%; height: 100%;
    opacity: 0; cursor: pointer;
  }
  .dz-icon {
    display: block;
    width: 34px; height: 34px;
    margin: 0 auto 0.5rem;
    color: var(--text-mute);
    flex-shrink: 0;
  }
  .dropzone:hover .dz-icon { color: var(--primary); }
  .dropzone.has-file .dz-icon { color: var(--accent-emerald); }
  .dz-text {
    font-size: 0.86rem; color: var(--text-soft); font-weight: 500;
    line-height: 1.4;
  }
  .dz-text strong { color: var(--text); font-weight: 700; }
  .dz-file {
    font-size: 0.86rem; color: var(--success); font-weight: 600;
    word-break: break-all; line-height: 1.4;
  }
  .dz-file small {
    display: block; color: var(--text-mute);
    font-weight: 500; margin-top: 3px; font-size: 0.76rem;
  }

  input[type=text] {
    display: block; width: 100%;
    padding: 0.7rem 0.9rem;
    border: 1.5px solid var(--border);
    border-radius: var(--radius-sm);
    font-family: inherit; font-size: 0.9rem;
    color: var(--text);
    transition: all 0.15s ease;
    background: white;
  }
  input[type=text]:focus {
    outline: none;
    border-color: var(--primary);
    box-shadow: 0 0 0 4px rgba(79,70,229,0.12);
  }
  input[type=text]::placeholder { color: var(--text-mute); }

  /* Submit */
  .actions { margin-top: 1.75rem; display: flex; gap: 0.75rem; align-items: center; }
  .btn {
    border: 0; cursor: pointer;
    font-family: inherit; font-weight: 600; font-size: 0.94rem;
    padding: 0.85rem 1.7rem; border-radius: var(--radius-sm);
    display: inline-flex; align-items: center; gap: 0.55rem;
    transition: all 0.18s ease;
  }
  .btn-primary {
    background: linear-gradient(135deg, #6366f1 0%, #8b5cf6 50%, #ec4899 100%);
    background-size: 150% 150%;
    color: white;
    box-shadow: 0 10px 24px -8px rgba(139, 92, 246, 0.55);
  }
  .btn-primary:hover { transform: translateY(-2px); box-shadow: 0 14px 30px -8px rgba(139, 92, 246, 0.65); background-position: 100% 100%; }
  .btn-primary:active { transform: translateY(0); }
  .btn-primary:disabled { background: var(--text-mute); cursor: wait; transform: none; box-shadow: none; }
  .btn svg { width: 16px; height: 16px; }
  .spinner {
    width: 16px; height: 16px; border: 2px solid rgba(255,255,255,0.4);
    border-top-color: white; border-radius: 50%;
    animation: spin 0.7s linear infinite;
  }
  @keyframes spin { to { transform: rotate(360deg); } }

  /* Result panel */
  .result {
    margin-top: 1.5rem;
    border-radius: var(--radius-sm);
    border: 1px solid var(--border);
    background: white;
    overflow: hidden;
    animation: slideIn 0.25s ease;
  }
  @keyframes slideIn { from { opacity: 0; transform: translateY(8px); } to { opacity: 1; transform: translateY(0); } }
  .result-header {
    display: flex; align-items: center; justify-content: space-between;
    padding: 0.9rem 1.1rem;
    background: var(--surface-soft);
    border-bottom: 1px solid var(--border);
  }
  .result-title {
    display: flex; align-items: center; gap: 0.55rem;
    font-weight: 700; font-size: 0.92rem;
  }
  .result-title svg { width: 18px; height: 18px; }
  .result.ok .result-header { background: var(--success-soft); border-bottom-color: #a7f3d0; }
  .result.ok .result-title { color: var(--success); }
  .result.error .result-header { background: var(--danger-soft); border-bottom-color: #fecaca; }
  .result.error .result-title { color: var(--danger); }
  .result.warn .result-header { background: var(--warning-soft); border-bottom-color: #fde68a; }
  .result.warn .result-title { color: var(--warning); }
  .result-body { padding: 1.1rem; }
  .stat-grid {
    display: grid; grid-template-columns: repeat(auto-fit, minmax(140px, 1fr));
    gap: 0.6rem; margin-bottom: 0.9rem;
  }
  .stat {
    padding: 0.7rem 0.9rem;
    background: var(--surface-soft);
    border-radius: var(--radius-sm);
    border: 1px solid var(--border);
  }
  .stat .label { font-size: 0.7rem; text-transform: uppercase; letter-spacing: 0.05em; color: var(--text-mute); font-weight: 600; }
  .stat .value { font-size: 1.4rem; font-weight: 700; color: var(--text); margin-top: 2px; }
  .msg-list { margin: 0.5rem 0 0; padding: 0; list-style: none; }
  .msg-list li {
    padding: 0.55rem 0.75rem;
    border-radius: 6px;
    font-size: 0.83rem;
    margin-bottom: 0.3rem;
    border-left: 3px solid;
    display: flex; gap: 0.55rem; align-items: flex-start;
  }
  .msg-list li.err { background: var(--danger-soft); border-color: var(--danger); color: #7f1d1d; }
  .msg-list li.warn { background: var(--warning-soft); border-color: var(--warning); color: #78350f; }
  .msg-list li svg { flex-shrink: 0; width: 14px; height: 14px; margin-top: 2px; }
  .download-btn {
    display: inline-flex; align-items: center; gap: 0.5rem;
    margin-top: 1rem;
    padding: 0.7rem 1.2rem;
    background: linear-gradient(135deg, #059669 0%, #10b981 100%);
    color: white; text-decoration: none; font-weight: 600; font-size: 0.88rem;
    border-radius: var(--radius-sm);
    box-shadow: 0 8px 20px -8px rgba(5,150,105,0.6);
    transition: all 0.15s ease;
  }
  .download-btn:hover { transform: translateY(-1px); box-shadow: 0 12px 24px -8px rgba(5,150,105,0.7); }
  .download-btn svg { width: 16px; height: 16px; }

  /* Inspect report */
  .slide-list { margin-top: 0.8rem; }
  .slide-item {
    padding: 0.75rem 0.9rem;
    border: 1px solid var(--border);
    border-radius: var(--radius-sm);
    margin-bottom: 0.5rem;
    background: white;
  }
  .slide-item.has-charts { border-color: #c7d2fe; background: #fafbff; }
  .slide-item-head { display: flex; justify-content: space-between; align-items: center; }
  .slide-num { font-weight: 700; color: var(--text); font-size: 0.9rem; }
  .slide-layout { color: var(--text-mute); font-size: 0.78rem; font-weight: 500; }
  .badge {
    background: var(--primary); color: white;
    font-size: 0.7rem; font-weight: 700;
    padding: 0.2rem 0.55rem; border-radius: 999px;
  }
  .badge.empty { background: #e5e7eb; color: var(--text-mute); }
  .chart-row {
    margin-top: 0.5rem; padding-left: 0.9rem;
    font-family: 'JetBrains Mono', ui-monospace, monospace;
    font-size: 0.76rem; color: var(--text-soft);
    border-left: 2px solid #c7d2fe;
  }
  .chart-row .cname { color: var(--primary); font-weight: 600; }
  .chart-row .ctype { color: var(--text-mute); }

  details.raw { margin-top: 0.9rem; }
  details.raw summary {
    cursor: pointer; font-size: 0.78rem; color: var(--text-mute);
    font-weight: 600; padding: 0.4rem 0;
  }
  details.raw pre {
    background: #0f172a; color: #e2e8f0;
    padding: 0.9rem; border-radius: var(--radius-sm);
    font-family: 'JetBrains Mono', ui-monospace, monospace;
    font-size: 0.74rem; overflow-x: auto;
    margin: 0.4rem 0 0;
  }

  /* Footer */
  footer {
    text-align: center; margin-top: 2rem;
    color: var(--text-mute); font-size: 0.8rem;
  }
  footer a { color: var(--primary); text-decoration: none; font-weight: 600; }
  footer a:hover { color: var(--primary-2); }

  @media (max-width: 600px) {
    .page { padding: 1.5rem 1rem 3rem; }
    .card-head, .card-body { padding-left: 1.25rem; padding-right: 1.25rem; }
    .brand h1 { font-size: 1.2rem; }
    .tabs { width: 100%; }
    .tab { flex: 1; justify-content: center; }
  }
</style>
</head>
<body>
<nav class="topnav">
  <div class="topnav-inner">
    <a href="/" class="topnav-brand">
      <span class="topnav-logo">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round">
          <rect x="3" y="3" width="18" height="18" rx="2"/>
          <path d="M3 9h18"/>
          <path d="M9 21V9"/>
        </svg>
      </span>
      <span class="topnav-title">PPT Engine</span>
    </a>
    <div class="topnav-links">
      <a class="nav-link" href="/">Home</a>
      <a class="nav-link active" href="/app">Operar</a>
      <a class="nav-link" href="/docs-ui">API</a>
      <a class="nav-link" href="/docs" target="_blank">Swagger</a>
      <span class="pill"><span class="dot"></span>Activo</span>
    </div>
  </div>
</nav>
<div class="page">
  <div class="page-intro">
    <h1 class="page-title">Operar el motor</h1>
    <p class="page-sub">Subí un template, los datos en Excel y generá la presentación con un click. O probá uno de los ejemplos pre-cargados.</p>
  </div>

  <section class="examples-card" id="examples-section">
    <div class="examples-head">
      <h2>
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 2 2 7l10 5 10-5-10-5z"/><path d="M2 17l10 5 10-5"/><path d="M2 12l10 5 10-5"/></svg>
        Ejemplos pre-cargados
      </h2>
      <p>Seleccioná un ejemplo para usarlo directamente (no hace falta subir archivos), o descargá los archivos para inspeccionarlos. Después tocá <b>Validar</b> o <b>Generar PPT</b> en el formulario de abajo.</p>
    </div>
    <div class="examples-grid" id="examples-grid">
      <div style="grid-column:1/-1;text-align:center;color:var(--text-mute);font-size:0.86rem;padding:1rem 0;">Cargando ejemplos…</div>
    </div>
  </section>

  <div class="card">
    <div class="card-head">
      <h2>Procesar archivos</h2>
      <p>Subí un template <code>.pptx</code>, los datos en Excel y opcionalmente un mapping. El motor actualiza los gráficos manteniendo el diseño visual original.</p>
      <div class="tabs" role="tablist">
        <button type="button" class="tab active" data-mode="build" role="tab">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 2v20M2 12h20"/></svg>
          Generar PPT
        </button>
        <button type="button" class="tab" data-mode="inspect" role="tab">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="8"/><path d="m21 21-4.3-4.3"/></svg>
          Inspeccionar
        </button>
        <button type="button" class="tab" data-mode="validate" role="tab">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 6 9 17l-5-5"/></svg>
          Validar
        </button>
      </div>
    </div>

    <div class="card-body">
      <form id="main-form" enctype="multipart/form-data">

        <div id="example-banner" class="example-banner" style="display:none;">
          <div class="example-banner-text">
            Usando ejemplo: <strong id="example-banner-name">—</strong>
          </div>
          <button type="button" class="example-banner-clear" id="example-banner-clear">Quitar ejemplo</button>
        </div>

        <div class="field">
          <div class="field-label">
            <span class="name">Template PowerPoint<span class="required">*</span></span>
            <span class="hint">.pptx / .pptm</span>
          </div>
          <label class="dropzone" data-target="template">
            <svg class="dz-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
              <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="17 8 12 3 7 8"/><line x1="12" y1="3" x2="12" y2="15"/>
            </svg>
            <div class="dz-text"><strong>Arrastrá tu .pptx</strong> o hacé click para seleccionar</div>
            <input type="file" name="template" accept=".pptx,.pptm" required>
          </label>
        </div>

        <div class="extras">
          <div class="field">
            <div class="field-label">
              <span class="name" id="data-label">Datos Excel<span class="required">*</span></span>
              <span class="hint">.xlsx / .xlsm</span>
            </div>
            <label class="dropzone" data-target="data">
              <svg class="dz-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="8" y1="13" x2="16" y2="13"/><line x1="8" y1="17" x2="16" y2="17"/>
              </svg>
              <div class="dz-text"><strong>Arrastrá tu .xlsx</strong> o hacé click para seleccionar</div>
              <input type="file" name="data" accept=".xlsx,.xlsm" required>
            </label>
          </div>

          <div class="field">
            <div class="field-label">
              <span class="name">Mapping YAML</span>
              <span class="hint">Opcional · usa default si se omite</span>
            </div>
            <label class="dropzone" data-target="mapping">
              <svg class="dz-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/>
              </svg>
              <div class="dz-text"><strong>Arrastrá tu mapping.yaml</strong> o dejalo en blanco</div>
              <input type="file" name="mapping" accept=".yaml,.yml">
            </label>
          </div>

          <div class="field" id="output-name-field">
            <div class="field-label">
              <span class="name">Nombre del archivo de salida</span>
              <span class="hint">Opcional</span>
            </div>
            <input type="text" name="output_name" placeholder="ej: estudio_chile_2026.pptx">
          </div>
        </div>

        <div class="actions">
          <button type="submit" id="submit-btn" class="btn btn-primary">
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12h14M13 5l7 7-7 7"/></svg>
            <span id="submit-label">Generar PPT</span>
          </button>
        </div>
      </form>

      <div id="result-container"></div>
    </div>
  </div>

  <footer>
    PPT Engine · <a href="/api/health" target="_blank">/api/health</a> · <a href="/docs" target="_blank">API docs</a>
  </footer>
</div>

<script>
const MODES = {
  build:    { url: "/api/jobs",     label: "Generar PPT",   active: "Generando PPT...",  needsData: true,  needsExtras: true  },
  inspect:  { url: "/api/inspect",  label: "Inspeccionar",  active: "Inspeccionando...", needsData: false, needsExtras: false },
  validate: { url: "/api/validate", label: "Validar",       active: "Validando...",      needsData: true,  needsExtras: false },
};

let currentMode = "build";
let activeExampleId = null;
const tabs = document.querySelectorAll(".tab");
const form = document.getElementById("main-form");
const extras = document.querySelector(".extras");
const outputNameField = document.getElementById("output-name-field");
const submitBtn = document.getElementById("submit-btn");
const submitLabel = document.getElementById("submit-label");
const resultContainer = document.getElementById("result-container");
const examplesGrid = document.getElementById("examples-grid");
const exampleBanner = document.getElementById("example-banner");
const exampleBannerName = document.getElementById("example-banner-name");
const exampleBannerClear = document.getElementById("example-banner-clear");

// ---- Ejemplos pre-cargados --------------------------------------------
function fmtBytes(n) {
  if (n == null) return "—";
  if (n < 1024) return n + " B";
  if (n < 1024*1024) return (n/1024).toFixed(1) + " KB";
  return (n/(1024*1024)).toFixed(1) + " MB";
}

async function loadExamples() {
  try {
    const r = await fetch("/api/examples");
    const json = await r.json();
    renderExamples(json.examples || []);
  } catch (err) {
    examplesGrid.innerHTML = '<div style="grid-column:1/-1;color:var(--danger);font-size:0.85rem">No se pudieron cargar los ejemplos: ' + err.message + '</div>';
  }
}

function renderExamples(examples) {
  if (!examples.length) {
    examplesGrid.innerHTML = '<div style="grid-column:1/-1;color:var(--text-mute);font-size:0.85rem">No hay ejemplos disponibles.</div>';
    return;
  }
  examplesGrid.innerHTML = "";
  examples.forEach(ex => {
    const card = document.createElement("div");
    card.className = "ex-card";
    card.dataset.exampleId = ex.id;

    const stats = ex.stats || {};
    const isErrorDemo = ex.id === "con-error";

    card.innerHTML = `
      <div class="ex-name">${escapeHtml(ex.name)}</div>
      <p class="ex-desc">${escapeHtml(ex.description)}</p>
      <div class="ex-stats">
        ${stats.slides != null ? `<span class="ex-stat">${stats.slides} slides</span>` : ""}
        ${stats.charts != null ? `<span class="ex-stat">${stats.charts} gráfico${stats.charts === 1 ? '' : 's'}</span>` : ""}
        ${stats.level ? `<span class="ex-stat ${isErrorDemo ? 'error' : 'level'}">${escapeHtml(stats.level)}</span>` : ""}
      </div>
      <div class="ex-actions">
        <button type="button" class="ex-btn ex-btn-use" data-action="use">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><path d="M20 6 9 17l-5-5"/></svg>
          Usar este
        </button>
        <a class="ex-btn ex-btn-dl" href="/api/examples/${ex.id}/files/template" download title="Descargar template (${fmtBytes(ex.files?.template?.size_bytes)})">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
          PPT
        </a>
        <a class="ex-btn ex-btn-dl" href="/api/examples/${ex.id}/files/data" download title="Descargar datos (${fmtBytes(ex.files?.data?.size_bytes)})">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
          XLSX
        </a>
        <a class="ex-btn ex-btn-dl" href="/api/examples/${ex.id}/files/mapping" download title="Descargar mapping (${fmtBytes(ex.files?.mapping?.size_bytes)})">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
          YAML
        </a>
      </div>
    `;
    card.querySelector('[data-action="use"]').addEventListener("click", () => selectExample(ex));
    examplesGrid.appendChild(card);
  });
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

function selectExample(ex) {
  activeExampleId = ex.id;
  document.querySelectorAll(".ex-card").forEach(c => c.classList.toggle("selected", c.dataset.exampleId === ex.id));
  exampleBannerName.textContent = ex.name;
  exampleBanner.style.display = "flex";
  // Hacer no required los inputs porque vamos a mandar example_id
  form.querySelectorAll('input[type=file]').forEach(i => i.required = false);
  resultContainer.innerHTML = "";
  // Scroll suave al form
  document.getElementById("main-form").scrollIntoView({ behavior: "smooth", block: "center" });
}

function clearExample() {
  activeExampleId = null;
  document.querySelectorAll(".ex-card").forEach(c => c.classList.remove("selected"));
  exampleBanner.style.display = "none";
  // Restaurar required del template
  const tpl = form.querySelector('input[name="template"]');
  if (tpl) tpl.required = true;
  // El required de "data" lo maneja setMode segun el modo
  setMode(currentMode);
}

exampleBannerClear.addEventListener("click", clearExample);
loadExamples();
// -----------------------------------------------------------------------
const dataInput = form.querySelector('input[name="data"]');
const dataDz = form.querySelector('[data-target="data"]');

function setMode(mode) {
  currentMode = mode;
  tabs.forEach(t => t.classList.toggle("active", t.dataset.mode === mode));
  const cfg = MODES[mode];
  submitLabel.textContent = cfg.label;
  // Si hay un ejemplo activo, ningun input es required (los archivos vienen del server)
  dataInput.required = activeExampleId ? false : cfg.needsData;
  const tplInput = form.querySelector('input[name="template"]');
  if (tplInput) tplInput.required = activeExampleId ? false : true;
  dataDz.style.display = cfg.needsData ? "" : "none";
  outputNameField.style.display = cfg.needsExtras ? "" : "none";
  const mappingDz = form.querySelector('[data-target="mapping"]').parentElement;
  mappingDz.style.display = (mode === "inspect") ? "none" : "";
  resultContainer.innerHTML = "";
}
tabs.forEach(t => t.addEventListener("click", () => setMode(t.dataset.mode)));

// Dropzone behavior
function fmtSize(b) {
  if (b < 1024) return b + " B";
  if (b < 1024*1024) return (b/1024).toFixed(1) + " KB";
  return (b/(1024*1024)).toFixed(1) + " MB";
}
function bindDz(dz) {
  const input = dz.querySelector('input[type=file]');
  const updateUI = () => {
    if (input.files && input.files[0]) {
      const f = input.files[0];
      dz.classList.add("has-file");
      dz.querySelector(".dz-text") && dz.querySelector(".dz-text").remove();
      let info = dz.querySelector(".dz-file");
      if (!info) {
        info = document.createElement("div");
        info.className = "dz-file";
        dz.appendChild(info);
      }
      info.innerHTML = f.name + "<small>" + fmtSize(f.size) + "</small>";
    }
  };
  input.addEventListener("change", updateUI);
  ["dragenter","dragover"].forEach(ev => dz.addEventListener(ev, e => {
    e.preventDefault();
    e.stopPropagation();
    if (e.dataTransfer) e.dataTransfer.dropEffect = "copy";
    dz.classList.add("dragover");
  }));
  dz.addEventListener("dragleave", e => {
    e.preventDefault();
    e.stopPropagation();
    if (e.target === dz) dz.classList.remove("dragover");
  });
  dz.addEventListener("drop", e => {
    e.preventDefault();
    e.stopPropagation();
    dz.classList.remove("dragover");
    const files = e.dataTransfer && e.dataTransfer.files;
    if (!files || !files.length) return;
    // Asignar los archivos dropeados al input via DataTransfer
    try {
      const dt = new DataTransfer();
      dt.items.add(files[0]);
      input.files = dt.files;
    } catch (_) {
      // Fallback: asignacion directa (funciona en algunos navegadores)
      input.files = files;
    }
    updateUI();
    input.dispatchEvent(new Event("change", { bubbles: true }));
  });
}
document.querySelectorAll(".dropzone").forEach(bindDz);
// Bloquear el drop por defecto en el resto del documento para que el browser
// no abra el archivo si el usuario suelta fuera de un dropzone.
["dragover","drop"].forEach(ev => document.addEventListener(ev, e => {
  if (!e.target.closest(".dropzone")) e.preventDefault();
}));

// Result rendering
function el(tag, attrs={}, ...children) {
  const e = document.createElement(tag);
  for (const [k,v] of Object.entries(attrs)) {
    if (k === "class") e.className = v;
    else if (k === "html") e.innerHTML = v;
    else e.setAttribute(k, v);
  }
  for (const c of children.flat()) {
    if (c == null) continue;
    e.appendChild(typeof c === "string" ? document.createTextNode(c) : c);
  }
  return e;
}
function iconCheck() { return el("span", {html: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M20 6 9 17l-5-5"/></svg>'}); }
function iconX()     { return el("span", {html: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M18 6 6 18M6 6l12 12"/></svg>'}); }
function iconWarn()  { return el("span", {html: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><path d="m21.73 18-8-14a2 2 0 0 0-3.46 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.73-3z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>'}); }
function iconDl()    { return el("span", {html: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>'}); }

function renderMessages(items, kind) {
  if (!items || !items.length) return null;
  const ul = el("ul", {class: "msg-list"});
  items.forEach(m => ul.appendChild(el("li", {class: kind}, kind === "err" ? iconX() : iconWarn(), el("span", {}, String(m)))));
  return ul;
}

function renderBuildResult(json, mode) {
  const errors = json.errors || [];
  const warnings = json.warnings || [];
  const isOk = json.status === "success" && errors.length === 0;
  const isWarn = isOk && warnings.length > 0;
  const cls = isOk ? (isWarn ? "warn" : "ok") : "error";
  const title = isOk ? "PPT generado correctamente" : "El job falló";
  const titleIcon = isOk ? iconCheck() : iconX();

  const stats = el("div", {class: "stat-grid"},
    el("div", {class: "stat"}, el("div", {class: "label"}, "Gráficos OK"), el("div", {class: "value"}, String(json.charts_updated ?? 0))),
    el("div", {class: "stat"}, el("div", {class: "label"}, "Gráficos KO"), el("div", {class: "value"}, String(json.charts_failed ?? 0))),
    el("div", {class: "stat"}, el("div", {class: "label"}, "Estado"), el("div", {class: "value"}, json.status || "—")),
  );

  const body = el("div", {class: "result-body"}, stats);
  if (json.error_message) body.appendChild(el("p", {class: "msg-list"}, el("li", {class: "err"}, iconX(), el("span", {}, json.error_message))));
  const errMsgs = renderMessages(errors, "err"); if (errMsgs) body.appendChild(errMsgs);
  const warnMsgs = renderMessages(warnings, "warn"); if (warnMsgs) body.appendChild(warnMsgs);

  if (json.output_available && json.job_id) {
    const a = el("a", {class: "download-btn", href: "/api/jobs/" + json.job_id + "/download"}, iconDl(), "Descargar presentación");
    body.appendChild(a);
  }
  body.appendChild(rawDetails(json));

  return el("div", {class: "result " + cls},
    el("div", {class: "result-header"}, el("div", {class: "result-title"}, titleIcon, title)),
    body
  );
}

function renderValidateResult(json) {
  const isOk = json.ok === true;
  const warnings = json.warnings || [];
  const errors = json.errors || [];
  const cls = isOk ? (warnings.length ? "warn" : "ok") : "error";
  const title = isOk ? (warnings.length ? "Validación OK con warnings" : "Validación correcta") : "Validación fallida";

  const body = el("div", {class: "result-body"});
  body.appendChild(el("p", {class: "stat-grid"},
    el("div", {class: "stat"}, el("div", {class: "label"}, "Errores"), el("div", {class: "value"}, String(errors.length))),
    el("div", {class: "stat"}, el("div", {class: "label"}, "Warnings"), el("div", {class: "value"}, String(warnings.length))),
    el("div", {class: "stat"}, el("div", {class: "label"}, "Resultado"), el("div", {class: "value"}, isOk ? "OK" : "KO")),
  ));
  const errMsgs = renderMessages(errors, "err"); if (errMsgs) body.appendChild(errMsgs);
  const warnMsgs = renderMessages(warnings, "warn"); if (warnMsgs) body.appendChild(warnMsgs);
  body.appendChild(rawDetails(json));

  return el("div", {class: "result " + cls},
    el("div", {class: "result-header"}, el("div", {class: "result-title"}, isOk ? iconCheck() : iconX(), title)),
    body
  );
}

function renderInspectResult(json) {
  const body = el("div", {class: "result-body"});
  body.appendChild(el("div", {class: "stat-grid"},
    el("div", {class: "stat"}, el("div", {class: "label"}, "Slides"), el("div", {class: "value"}, String(json.total_slides ?? 0))),
    el("div", {class: "stat"}, el("div", {class: "label"}, "Gráficos"), el("div", {class: "value"}, String(json.total_charts ?? 0))),
    el("div", {class: "stat"}, el("div", {class: "label"}, "Template"), el("div", {class: "value", style: "font-size:0.9rem;word-break:break-all"}, json.template || "—")),
  ));

  const list = el("div", {class: "slide-list"});
  (json.slides || []).forEach(s => {
    const item = el("div", {class: "slide-item" + (s.has_charts ? " has-charts" : "")});
    item.appendChild(el("div", {class: "slide-item-head"},
      el("div", {},
        el("span", {class: "slide-num"}, "Slide " + s.slide_index),
        el("span", {class: "slide-layout"}, "  ·  " + (s.layout_name || ""))
      ),
      el("span", {class: "badge" + (s.has_charts ? "" : " empty")}, (s.charts || []).length + " gráfico(s)")
    ));
    (s.charts || []).forEach(c => {
      item.appendChild(el("div", {class: "chart-row"},
        "index=" + c.chart_index + "  ",
        el("span", {class: "cname"}, "'" + (c.shape_name || "") + "'"),
        "  ",
        el("span", {class: "ctype"}, c.chart_type || "")
      ));
    });
    list.appendChild(item);
  });
  body.appendChild(list);
  body.appendChild(rawDetails(json));

  return el("div", {class: "result ok"},
    el("div", {class: "result-header"}, el("div", {class: "result-title"}, iconCheck(), "Template inspeccionado")),
    body
  );
}

function rawDetails(json) {
  const d = el("details", {class: "raw"});
  d.appendChild(el("summary", {}, "Ver respuesta JSON"));
  d.appendChild(el("pre", {}, JSON.stringify(json, null, 2)));
  return d;
}

function renderError(status, json) {
  const body = el("div", {class: "result-body"},
    el("p", {style: "margin:0 0 0.5rem;color:#7f1d1d;font-weight:600"}, "HTTP " + status),
    el("pre", {style: "background:#0f172a;color:#e2e8f0;padding:0.9rem;border-radius:8px;font-family:'JetBrains Mono',monospace;font-size:0.78rem;overflow-x:auto;margin:0"}, JSON.stringify(json, null, 2))
  );
  return el("div", {class: "result error"},
    el("div", {class: "result-header"}, el("div", {class: "result-title"}, iconX(), "Error en la petición")),
    body
  );
}

form.addEventListener("submit", async (e) => {
  e.preventDefault();
  const cfg = MODES[currentMode];
  submitBtn.disabled = true;
  submitLabel.textContent = cfg.active;
  // swap arrow icon for spinner
  const oldSvg = submitBtn.querySelector("svg");
  const spinner = document.createElement("div");
  spinner.className = "spinner";
  if (oldSvg) submitBtn.replaceChild(spinner, oldSvg);
  resultContainer.innerHTML = "";

  const fd = new FormData(form);
  // Limpiar campos vacios para que FastAPI no los reciba como UploadFile vacios
  for (const name of ["template","data","mapping"]) {
    const f = fd.get(name);
    if (f instanceof File && f.size === 0 && !f.name) fd.delete(name);
  }
  if (currentMode === "inspect") { fd.delete("data"); fd.delete("mapping"); fd.delete("output_name"); }
  if (currentMode === "validate") { fd.delete("output_name"); }
  // Si hay un ejemplo activo, mandarlo. Esto le dice al backend que use
  // los archivos pre-cargados en lugar de los uploads (que pueden estar vacios).
  if (activeExampleId) {
    fd.append("example_id", activeExampleId);
    // Borrar uploads vacios para evitar el "filename required" error
    for (const name of ["template","data","mapping"]) {
      const f = fd.get(name);
      if (f instanceof File && (f.size === 0 || !f.name)) fd.delete(name);
    }
  }

  try {
    const resp = await fetch(cfg.url, { method: "POST", body: fd });
    const json = await resp.json();
    let node;
    if (!resp.ok) node = renderError(resp.status, json);
    else if (currentMode === "build") node = renderBuildResult(json, currentMode);
    else if (currentMode === "validate") node = renderValidateResult(json);
    else node = renderInspectResult(json);
    resultContainer.appendChild(node);
  } catch (err) {
    resultContainer.appendChild(renderError("network", { error: err.message }));
  } finally {
    submitBtn.disabled = false;
    submitLabel.textContent = cfg.label;
    if (oldSvg) submitBtn.replaceChild(oldSvg, spinner);
  }
});

setMode("build");
</script>
</body>
</html>
"""


# ---------------------------------------------------------------------- #
# Homepage (/) y API reference (/docs-ui)                                 #
# ---------------------------------------------------------------------- #


_HOME_CSS = """
  :root {
    --bg: #f5f7fb;
    --surface: #ffffff;
    --surface-soft: #f8fafc;
    --surface-softer: #f1f5f9;
    --border: #e5e7eb;
    --border-strong: #cbd5e1;
    --text: #0f172a;
    --text-soft: #475569;
    --text-mute: #94a3b8;
    --primary: #6366f1;
    --primary-2: #8b5cf6;
    --primary-3: #ec4899;
    --primary-soft: #eef2ff;
    --accent-cyan: #06b6d4;
    --accent-emerald: #10b981;
    --success: #059669;
    --success-soft: #ecfdf5;
    --shadow-xl: 0 20px 60px -15px rgba(99, 102, 241, 0.25);
    --shadow-lg: 0 10px 40px -10px rgba(15, 23, 42, 0.12);
    --shadow-md: 0 4px 16px -4px rgba(15, 23, 42, 0.08);
    --shadow-sm: 0 1px 3px rgba(15, 23, 42, 0.06);
    --radius: 16px;
    --radius-sm: 10px;
    --mono: 'JetBrains Mono', ui-monospace, Menlo, monospace;
    --sans: 'Inter', system-ui, -apple-system, sans-serif;
  }
  * { box-sizing: border-box; }
  html, body { margin: 0; padding: 0; }
  body {
    font-family: var(--sans);
    background:
      radial-gradient(900px 500px at 0% 0%, rgba(139, 92, 246, 0.12), transparent 60%),
      radial-gradient(900px 500px at 100% 0%, rgba(6, 182, 212, 0.10), transparent 60%),
      radial-gradient(700px 400px at 50% 100%, rgba(236, 72, 153, 0.06), transparent 60%),
      var(--bg);
    min-height: 100vh;
    color: var(--text);
    -webkit-font-smoothing: antialiased;
    text-rendering: optimizeLegibility;
    line-height: 1.55;
  }
  a { color: var(--primary); text-decoration: none; }
  a:hover { color: var(--primary-2); }

  .container { max-width: 1100px; margin: 0 auto; padding: 0 1.5rem; }

  /* Topnav */
  .topnav {
    position: sticky; top: 0; z-index: 20;
    background: rgba(255,255,255,0.85);
    backdrop-filter: saturate(180%) blur(12px);
    border-bottom: 1px solid var(--border);
  }
  .topnav-inner {
    display: flex; align-items: center; justify-content: space-between;
    padding: 0.85rem 1.5rem; max-width: 1100px; margin: 0 auto;
    gap: 1rem; flex-wrap: wrap;
  }
  .brand {
    display: flex; align-items: center; gap: 0.7rem;
    font-weight: 800; font-size: 1.1rem; color: var(--text);
    letter-spacing: -0.01em;
  }
  .brand-logo {
    width: 36px; height: 36px; border-radius: 10px;
    background: linear-gradient(135deg, #6366f1 0%, #8b5cf6 50%, #ec4899 100%);
    display: grid; place-items: center;
    color: white; box-shadow: 0 10px 24px -8px rgba(139,92,246,0.5);
  }
  .brand-logo svg { width: 20px; height: 20px; }
  .nav-links { display: flex; gap: 0.25rem; flex-wrap: wrap; }
  .nav-link {
    padding: 0.5rem 0.95rem; border-radius: 8px;
    color: var(--text-soft); font-weight: 600; font-size: 0.88rem;
    transition: all .15s ease;
  }
  .nav-link:hover { background: var(--surface-softer); color: var(--text); }
  .nav-link.active { background: var(--primary-soft); color: var(--primary); }

  /* Hero */
  .hero { padding: 4.5rem 1.5rem 3.5rem; text-align: center; }
  .hero-pill {
    display: inline-block;
    padding: 0.35rem 1rem;
    border: 1px solid var(--border);
    background: white;
    border-radius: 100px;
    font-size: 0.8rem; color: var(--text-soft); font-weight: 500;
    margin-bottom: 1.75rem;
    box-shadow: var(--shadow-sm);
  }
  .hero-pill b {
    background: linear-gradient(135deg, #6366f1 0%, #8b5cf6 50%, #ec4899 100%);
    -webkit-background-clip: text; background-clip: text; color: transparent;
    font-weight: 700;
  }
  .hero h1 {
    font-size: clamp(2rem, 5vw, 3.25rem);
    line-height: 1.08; font-weight: 800;
    margin: 0 auto 1.3rem; max-width: 800px;
    letter-spacing: -0.025em;
    color: var(--text);
  }
  .hero h1 .grad {
    background: linear-gradient(135deg, #6366f1 0%, #8b5cf6 50%, #ec4899 100%);
    -webkit-background-clip: text; background-clip: text; color: transparent;
  }
  .hero p.lead {
    font-size: 1.1rem; color: var(--text-soft);
    max-width: 640px; margin: 0 auto 2.2rem;
    line-height: 1.6;
  }
  .cta-row {
    display: flex; gap: 0.75rem; justify-content: center; flex-wrap: wrap;
  }
  .btn {
    display: inline-flex; align-items: center; gap: 0.55rem;
    padding: 0.85rem 1.5rem;
    border-radius: 10px; border: 0; cursor: pointer;
    font-family: var(--sans); font-size: 0.95rem; font-weight: 600;
    text-decoration: none;
    transition: all .18s ease;
  }
  .btn svg { width: 18px; height: 18px; }
  .btn-primary {
    background: linear-gradient(135deg, #6366f1 0%, #8b5cf6 50%, #ec4899 100%);
    color: white;
    box-shadow: 0 10px 30px -8px rgba(139,92,246,0.55);
  }
  .btn-primary:hover {
    transform: translateY(-2px);
    box-shadow: 0 14px 34px -8px rgba(139,92,246,0.65);
    color: white;
  }
  .btn-ghost {
    background: white; color: var(--text);
    border: 1px solid var(--border);
    box-shadow: var(--shadow-sm);
  }
  .btn-ghost:hover {
    border-color: var(--primary);
    color: var(--primary);
    transform: translateY(-1px);
  }

  /* Stats */
  .stats {
    display: grid; gap: 1rem;
    grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
    margin-top: 3rem;
  }
  .stat {
    background: white; border: 1px solid var(--border);
    border-radius: var(--radius-sm); padding: 1.2rem 1.35rem;
    box-shadow: var(--shadow-sm);
  }
  .stat .big {
    font-size: 2.1rem; font-weight: 800; letter-spacing: -0.03em;
    background: linear-gradient(135deg, #6366f1 0%, #8b5cf6 50%, #ec4899 100%);
    -webkit-background-clip: text; background-clip: text; color: transparent;
  }
  .stat .lbl { color: var(--text-soft); font-size: 0.85rem; margin-top: 0.15rem; font-weight: 500; }

  /* Sections */
  section { padding: 3.5rem 0; }
  section.alt { background: white; border-top: 1px solid var(--border); border-bottom: 1px solid var(--border); }
  .section-kicker {
    font-size: 0.78rem; text-transform: uppercase;
    color: var(--primary); letter-spacing: 0.14em; font-weight: 700;
    margin-bottom: 0.6rem;
  }
  section h2 {
    font-size: clamp(1.6rem, 3.5vw, 2.1rem);
    margin: 0 0 0.6rem; max-width: 600px;
    letter-spacing: -0.02em; line-height: 1.2;
  }
  section p.subtitle {
    color: var(--text-soft); max-width: 660px; margin: 0 0 2.5rem;
    font-size: 1.02rem; line-height: 1.6;
  }

  /* Stage cards */
  .stages {
    display: grid; gap: 1.25rem;
    grid-template-columns: repeat(auto-fit, minmax(290px, 1fr));
  }
  .stage-card {
    background: white;
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 1.75rem;
    box-shadow: var(--shadow-sm);
    transition: all .2s ease;
    position: relative; overflow: hidden;
  }
  .stage-card::before {
    content: ""; position: absolute; top: 0; left: 0; right: 0; height: 3px;
    background: linear-gradient(90deg, #6366f1 0%, #8b5cf6 50%, #ec4899 100%);
    opacity: 0; transition: opacity .2s;
  }
  .stage-card:hover {
    transform: translateY(-4px);
    box-shadow: var(--shadow-lg);
  }
  .stage-card:hover::before { opacity: 1; }
  .stage-card .num {
    width: 42px; height: 42px; border-radius: 10px;
    background: linear-gradient(135deg, #6366f1 0%, #8b5cf6 50%, #ec4899 100%);
    display: grid; place-items: center;
    font-weight: 800; color: white; font-size: 1.15rem;
    margin-bottom: 1rem;
    box-shadow: 0 8px 20px -6px rgba(139,92,246,0.4);
  }
  .stage-card h3 {
    margin: 0 0 0.45rem;
    font-size: 1.2rem; font-weight: 700;
    letter-spacing: -0.01em;
  }
  .stage-card p { color: var(--text-soft); margin: 0 0 1rem; font-size: 0.94rem; line-height: 1.55; }
  .stage-card ul {
    list-style: none; padding: 0; margin: 0;
    font-size: 0.87rem; color: var(--text-soft);
  }
  .stage-card ul li {
    padding: 0.3rem 0; padding-left: 1.5rem; position: relative;
  }
  .stage-card ul li::before {
    content: ""; position: absolute; left: 0; top: 0.65rem;
    width: 14px; height: 14px;
    background-image: url("data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='%23059669' stroke-width='3' stroke-linecap='round' stroke-linejoin='round'><path d='M20 6 9 17l-5-5'/></svg>");
    background-repeat: no-repeat; background-size: contain;
  }

  /* Pipeline */
  .pipeline {
    display: grid; gap: 1rem;
    grid-template-columns: 1fr;
    margin-top: 1.5rem;
  }
  @media (min-width: 820px) {
    .pipeline { grid-template-columns: 1fr auto 1fr auto 1fr; align-items: center; }
  }
  .pipe-box {
    background: white;
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 1.4rem 1.2rem;
    text-align: center;
    box-shadow: var(--shadow-sm);
    transition: all .2s ease;
  }
  .pipe-box:hover { transform: translateY(-2px); box-shadow: var(--shadow-md); }
  .pipe-box .icon { font-size: 2rem; margin-bottom: 0.55rem; }
  .pipe-box h4 { margin: 0 0 0.3rem; font-size: 1rem; font-weight: 700; }
  .pipe-box p { margin: 0; color: var(--text-mute); font-size: 0.82rem; font-family: var(--mono); }
  .pipe-arrow {
    color: var(--primary); font-size: 1.8rem; font-weight: 300;
    text-align: center;
    transform: rotate(90deg);
  }
  @media (min-width: 820px) { .pipe-arrow { transform: none; } }

  /* Samples grid */
  .samples-grid {
    display: grid; gap: 0.9rem;
    grid-template-columns: repeat(auto-fit, minmax(270px, 1fr));
    margin-top: 1.75rem;
  }
  .sample-item {
    display: flex; align-items: center; gap: 1rem;
    background: white; border: 1px solid var(--border);
    border-radius: var(--radius-sm); padding: 1rem 1.1rem;
    color: var(--text); text-decoration: none;
    transition: all .15s ease;
    box-shadow: var(--shadow-sm);
  }
  .sample-item:hover {
    border-color: var(--primary);
    transform: translateY(-2px);
    box-shadow: var(--shadow-md);
    color: var(--text);
  }
  .sample-item .ext {
    flex-shrink: 0;
    background: var(--primary-soft);
    color: var(--primary);
    padding: 0.3rem 0.6rem; border-radius: 6px;
    font-size: 0.7rem; font-weight: 800; font-family: var(--mono);
    letter-spacing: 0.02em;
  }
  .sample-item .info { min-width: 0; }
  .sample-item .name {
    font-weight: 600; font-size: 0.92rem;
    color: var(--text); word-break: break-all;
  }
  .sample-item .hint {
    color: var(--text-soft); font-size: 0.8rem;
    margin-top: 0.15rem; line-height: 1.4;
  }

  /* Docs page */
  .endpoint {
    background: white; border: 1px solid var(--border);
    border-radius: var(--radius-sm); padding: 1.2rem 1.4rem;
    margin-bottom: 1rem;
    box-shadow: var(--shadow-sm);
  }
  .endpoint .method {
    display: inline-block; padding: 0.22rem 0.65rem;
    border-radius: 6px; font-family: var(--mono); font-size: 0.72rem;
    font-weight: 800; text-transform: uppercase;
    margin-right: 0.6rem; letter-spacing: 0.03em;
  }
  .endpoint .method.get { background: #ecfeff; color: var(--accent-cyan); }
  .endpoint .method.post { background: var(--primary-soft); color: var(--primary); }
  .endpoint code {
    font-family: var(--mono); font-size: 0.9rem;
    color: var(--text); font-weight: 600;
  }
  .endpoint p { margin: 0.55rem 0 0; color: var(--text-soft); font-size: 0.92rem; line-height: 1.55; }
  .endpoint pre {
    background: #0f172a; color: #e2e8f0;
    border-radius: 8px; padding: 0.9rem 1.1rem;
    font-family: var(--mono); font-size: 0.8rem;
    overflow-x: auto; margin: 0.8rem 0 0;
    line-height: 1.5;
  }

  /* Footer */
  footer {
    border-top: 1px solid var(--border);
    padding: 2.25rem 1.5rem;
    color: var(--text-mute);
    text-align: center;
    font-size: 0.88rem;
    background: white;
  }
  footer a { color: var(--primary); font-weight: 600; }
"""


def _render_topnav(active: str) -> str:
    def cls(name: str) -> str:
        return "nav-link active" if name == active else "nav-link"
    return f"""
<nav class="topnav">
  <div class="topnav-inner">
    <a href="/" class="brand">
      <span class="brand-logo">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round">
          <rect x="3" y="3" width="18" height="18" rx="2"/>
          <path d="M3 9h18"/>
          <path d="M9 21V9"/>
        </svg>
      </span>
      PPT Engine
    </a>
    <div class="nav-links">
      <a class="{cls('home')}" href="/">Home</a>
      <a class="{cls('app')}" href="/app">Operar</a>
      <a class="{cls('docs')}" href="/docs-ui">API</a>
      <a class="nav-link" href="/docs" target="_blank">Swagger</a>
    </div>
  </div>
</nav>
"""


def _render_home() -> str:
    nav = _render_topnav("home")
    return f"""<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>PPT Engine — Motor de generacion automatica de PowerPoint</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;500;700&display=swap" rel="stylesheet">
<style>{_HOME_CSS}</style>
</head>
<body>
{nav}

<header class="hero">
  <div class="container">
    <div class="hero-pill">Proyecto <b>v1.0</b> · Etapas 1, 2 y 3 integradas</div>
    <h1>Del Excel al PowerPoint final <span class="grad">en un click</span></h1>
    <p class="lead">
      Motor que reemplaza el flujo manual de cargar tablas en graficos embebidos
      de PowerPoint. Acepta datos crudos, tablas cruzadas o ambos, y genera
      presentaciones listas para entregar preservando el diseno del template.
    </p>
    <div class="cta-row">
      <a class="btn btn-primary" href="/app">
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12h14M13 5l7 7-7 7"/></svg>
        Abrir la app
      </a>
      <a class="btn btn-ghost" href="/api/samples" download>
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
        Descargar ejemplos (.zip)
      </a>
    </div>

    <div class="container stats">
      <div class="stat"><div class="big">61</div><div class="lbl">tests pasando</div></div>
      <div class="stat"><div class="big">6</div><div class="lbl">modulos del engine</div></div>
      <div class="stat"><div class="big">3</div><div class="lbl">etapas integradas</div></div>
      <div class="stat"><div class="big">~2s</div><div class="lbl">por batch de 5 PPTs</div></div>
    </div>
  </div>
</header>

<section class="alt">
  <div class="container">
    <div class="section-kicker">Problema</div>
    <h2>Un flujo manual que no escala</h2>
    <p class="subtitle">
      Hoy el equipo arma presentaciones pegando datos en Exceles embebidos
      dentro de graficos de PowerPoint, replicando el mismo proceso para
      cada pais o estudio. Horas de trabajo operativo por cada entregable.
    </p>
  </div>
</section>

<section>
  <div class="container">
    <div class="section-kicker">Solucion</div>
    <h2>Tres etapas, una sola API</h2>
    <p class="subtitle">
      El motor esta dividido en etapas que podes usar por separado o
      encadenadas. Cada etapa resuelve una parte del pipeline y puede
      ejecutarse desde CLI, desde la interfaz web, o desde Python.
    </p>

    <div class="stages">
      <div class="stage-card">
        <div class="num">1</div>
        <h3>Motor de PPT</h3>
        <p>Toma un template .pptx y un Excel con tablas, y actualiza los graficos preservando el diseno.</p>
        <ul>
          <li>Actualiza series y categorias</li>
          <li>Reescribe el xlsx embebido en cada grafico</li>
          <li>Procesamiento batch por paises</li>
          <li>Validacion fail-fast con errores claros</li>
        </ul>
      </div>

      <div class="stage-card">
        <div class="num">2</div>
        <h3>Interfaz web</h3>
        <p>FastAPI + UI moderna para operar sin depender de un desarrollador. Upload, build, download.</p>
        <ul>
          <li>Subir archivos y generar con un click</li>
          <li>Inspect del template</li>
          <li>Validacion antes de procesar</li>
          <li>Descarga directa del PPT generado</li>
        </ul>
      </div>

      <div class="stage-card">
        <div class="num">3</div>
        <h3>Procesamiento previo</h3>
        <p>Lee .sav/.dta/.csv/.xlsx, aplica un Tab Plan y produce tablas cruzadas listas para Etapa 1.</p>
        <ul>
          <li>Lector unificado con pyreadstat</li>
          <li>Tab Plans en YAML o Excel</li>
          <li>Motor de crosstabs con pandas</li>
          <li>LLM interpreter pluggable</li>
        </ul>
      </div>
    </div>
  </div>
</section>

<section class="alt">
  <div class="container">
    <div class="section-kicker">Como funciona</div>
    <h2>Pipeline de datos</h2>
    <p class="subtitle">
      Podes entrar en cualquier punto del pipeline. Si ya tenes las tablas
      cruzadas, arrancas directo en la Etapa 1. Si tenes las respuestas
      crudas, arrancas en la 3.
    </p>

    <div class="pipeline">
      <div class="pipe-box">
        <div class="icon">📊</div>
        <h4>Respuestas crudas</h4>
        <p>.sav .dta .csv .xlsx</p>
      </div>
      <div class="pipe-arrow">→</div>
      <div class="pipe-box">
        <div class="icon">🔀</div>
        <h4>Tablas cruzadas</h4>
        <p>crosstabs.xlsx</p>
      </div>
      <div class="pipe-arrow">→</div>
      <div class="pipe-box">
        <div class="icon">📽️</div>
        <h4>PowerPoint final</h4>
        <p>output.pptx</p>
      </div>
    </div>
  </div>
</section>

<section id="samples">
  <div class="container">
    <div class="section-kicker">Pruebalo</div>
    <h2>Archivos de ejemplo listos para descargar</h2>
    <p class="subtitle">
      Bajate cualquiera de estos archivos y subilos en la pagina
      <a href="/app">Operar</a>. O bajate el zip entero para tener
      el set completo con un README.
    </p>

    <div class="samples-grid">
      {_render_samples_cards()}
    </div>

    <div style="margin-top: 2rem; text-align: center;">
      <a class="btn btn-primary" href="/api/samples" download>
        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
        Descargar todos en un .zip
      </a>
    </div>
  </div>
</section>

<section class="alt">
  <div class="container">
    <div class="section-kicker">Stack</div>
    <h2>Tecnologias</h2>
    <p class="subtitle">
      Python moderno, sin frameworks pesados. Cada modulo se puede usar
      de forma independiente desde Python o desde la API REST.
    </p>
    <div class="stages">
      <div class="stage-card">
        <h3>Etapa 1 — Engine</h3>
        <p>python-pptx · openpyxl · pandas · typer · PyYAML</p>
      </div>
      <div class="stage-card">
        <h3>Etapa 2 — Web</h3>
        <p>FastAPI · Uvicorn · HTML/CSS/JS vanilla (sin build)</p>
      </div>
      <div class="stage-card">
        <h3>Etapa 3 — Preprocess</h3>
        <p>pyreadstat · pandas · LLM interpreter (mock + Anthropic)</p>
      </div>
    </div>
  </div>
</section>

<footer>
  <div class="container">
    PPT Engine · Autor: Ronald · 61 tests · Python 3.10+
    &nbsp;·&nbsp; <a href="/docs-ui">API</a>
    &nbsp;·&nbsp; <a href="/docs" target="_blank">Swagger</a>
    &nbsp;·&nbsp; <a href="/app">Operar</a>
  </div>
</footer>
</body>
</html>"""


def _render_samples_cards() -> str:
    """Genera las tarjetas de samples desde SAMPLE_FILES."""
    cards = []
    for slug, (filename, _mt, desc) in SAMPLE_FILES.items():
        ext = filename.rsplit(".", 1)[-1].upper() if "." in filename else "FILE"
        cards.append(
            f"""<a class="sample-item" href="/api/samples/{slug}" download>
  <span class="ext">{ext}</span>
  <div class="info">
    <div class="name">{filename}</div>
    <div class="hint">{desc}</div>
  </div>
</a>"""
        )
    return "\n".join(cards)


def _render_docs() -> str:
    nav = _render_topnav("docs")
    return f"""<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>PPT Engine — API Reference</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;500;700&display=swap" rel="stylesheet">
<style>{_HOME_CSS}</style>
</head>
<body>
{nav}

<main class="container" style="padding-top: 2.8rem; padding-bottom: 4rem;">
  <div class="section-kicker">API Reference</div>
  <h2 style="margin: 0 0 0.6rem;">Endpoints REST</h2>
  <p style="color: var(--text-soft); max-width: 700px;">
    Todos los endpoints aceptan <code>multipart/form-data</code>.
    Para la version auto-generada con modelos Pydantic, ver
    <a href="/docs" target="_blank">Swagger UI</a>.
  </p>

  <h3 style="margin: 2.2rem 0 1rem; font-size: 1.1rem;">Etapa 1 — Motor de PPT</h3>

  <div class="endpoint">
    <span class="method post">POST</span><code>/api/inspect</code>
    <p>Analiza un template .pptx y devuelve slides, layouts y graficos detectados.</p>
    <pre>curl -X POST http://127.0.0.1:8765/api/inspect \\
  -F "template=@template.pptx"</pre>
  </div>

  <div class="endpoint">
    <span class="method post">POST</span><code>/api/validate</code>
    <p>Valida template + xlsx + mapping sin generar archivos. Fail-fast.</p>
    <pre>curl -X POST http://127.0.0.1:8765/api/validate \\
  -F "template=@template.pptx" \\
  -F "data=@data.xlsx" \\
  -F "mapping=@mapping.yaml"</pre>
  </div>

  <div class="endpoint">
    <span class="method post">POST</span><code>/api/jobs</code>
    <p>Genera un PPT a partir de template + xlsx + mapping. Devuelve un job con un id.</p>
    <pre>curl -X POST http://127.0.0.1:8765/api/jobs \\
  -F "template=@template.pptx" \\
  -F "data=@data.xlsx" \\
  -F "mapping=@mapping.yaml"</pre>
  </div>

  <div class="endpoint">
    <span class="method get">GET</span><code>/api/jobs/{{id}}</code>
    <p>Estado del job. Devuelve status, charts_updated, errors, warnings.</p>
  </div>

  <div class="endpoint">
    <span class="method get">GET</span><code>/api/jobs/{{id}}/download</code>
    <p>Descarga el .pptx generado.</p>
  </div>

  <h3 style="margin: 2.2rem 0 1rem; font-size: 1.1rem;">Etapa 3 — Procesamiento previo</h3>

  <div class="endpoint">
    <span class="method post">POST</span><code>/api/preprocess</code>
    <p>Genera un xlsx de tablas cruzadas desde respuestas + tab plan (YAML o texto libre).</p>
    <pre>curl -X POST http://127.0.0.1:8765/api/preprocess \\
  -F "data=@responses.csv" \\
  -F "tab_plan=@plan.yaml" \\
  -o crosstabs.xlsx</pre>
  </div>

  <h3 style="margin: 2.2rem 0 1rem; font-size: 1.1rem;">Utilidad</h3>

  <div class="endpoint">
    <span class="method get">GET</span><code>/api/samples</code>
    <p>Descarga un zip con todos los archivos de ejemplo listos para probar.</p>
  </div>

  <div class="endpoint">
    <span class="method get">GET</span><code>/api/samples/{{slug}}</code>
    <p>Descarga un archivo individual. Slugs:
    <code>template</code>, <code>data</code>, <code>mapping</code>,
    <code>responses</code>, <code>tab-plan</code>, <code>tab-plan-xlsx</code>.</p>
  </div>

  <div class="endpoint">
    <span class="method get">GET</span><code>/api/health</code>
    <p>Health check. Devuelve status, workdir, jobs en memoria y samples disponibles.</p>
  </div>
</main>

<footer>
  <div class="container">
    <a href="/">Home</a> · <a href="/app">Operar</a> · <a href="/docs" target="_blank">Swagger UI</a>
  </div>
</footer>
</body>
</html>"""


def _render_samples_readme() -> str:
    return (
        "PPT Engine - Archivos de ejemplo\n"
        "=================================\n"
        "\n"
        "Este zip contiene todos los fixtures listos para probar el motor.\n"
        "\n"
        "Contenido:\n"
        "\n"
        "- sample_template.pptx       Template base con 3 slides y 3 graficos\n"
        "- sample_data.xlsx           Tablas cruzadas listas para Etapa 1\n"
        "- sample_mapping_multi.yaml  Mapping slide -> chart -> tabla\n"
        "- sample_responses.csv       15 respuestas crudas para Etapa 3\n"
        "- sample_tab_plan.yaml       Tab Plan con 3 crosses\n"
        "- sample_tab_plan.xlsx       Mismo plan embebido en Excel\n"
        "\n"
        "Como probar desde la UI:\n"
        "\n"
        "1. Abri http://127.0.0.1:8765 y anda a 'Operar'\n"
        "2. Elegi 'Generar PPT'\n"
        "3. Subi sample_template.pptx como Template\n"
        "4. Subi sample_data.xlsx como Datos\n"
        "5. Subi sample_mapping_multi.yaml como Mapping\n"
        "6. Click en 'Generar PPT' y descarga el resultado\n"
    )


# ---------------------------------------------------------------------- #
# ASGI app por defecto (para uvicorn web.app:app)                         #
# ---------------------------------------------------------------------- #

app = create_app()

__all__ = ["create_app", "Job", "app"]
