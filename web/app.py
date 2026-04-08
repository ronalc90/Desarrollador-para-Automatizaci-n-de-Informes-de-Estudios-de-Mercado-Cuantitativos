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

import shutil
import uuid
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse

from engine.excel_reader import ExcelReaderError
from engine.inspector import inspect_template
from engine.ppt_builder import BuildResult, PPTBuilderError, build_presentation
from engine.validator import load_mapping, validate_all

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
        }

    # ------------------------------------------------------------------ #
    # Inspect                                                            #
    # ------------------------------------------------------------------ #

    @app.post("/api/inspect")
    async def inspect_endpoint(
        template: UploadFile = File(..., description="Template .pptx"),
    ) -> JSONResponse:
        if not template.filename or not template.filename.lower().endswith(
            (".pptx", ".pptm")
        ):
            raise HTTPException(
                status_code=400, detail="El archivo debe ser .pptx/.pptm"
            )
        job_id = uuid.uuid4().hex
        job_dir = workdir / "uploads" / job_id
        job_dir.mkdir(parents=True, exist_ok=True)
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
        template: UploadFile = File(...),
        data: UploadFile = File(...),
        mapping: Optional[UploadFile] = File(None),
    ) -> JSONResponse:
        job_id = uuid.uuid4().hex
        job_dir = workdir / "uploads" / job_id
        job_dir.mkdir(parents=True, exist_ok=True)

        template_path = _save_upload(template, job_dir / template.filename)
        data_path = _save_upload(data, job_dir / data.filename)
        mapping_path = _resolve_mapping(mapping, job_dir)

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
        template: UploadFile = File(...),
        data: UploadFile = File(...),
        mapping: Optional[UploadFile] = File(None),
        output_name: Optional[str] = Form(None),
    ) -> JSONResponse:
        job_id = uuid.uuid4().hex
        job_dir = workdir / "uploads" / job_id
        job_dir.mkdir(parents=True, exist_ok=True)
        output_dir = workdir / "outputs" / job_id
        output_dir.mkdir(parents=True, exist_ok=True)

        template_path = _save_upload(template, job_dir / template.filename)
        data_path = _save_upload(data, job_dir / data.filename)
        mapping_path = _resolve_mapping(mapping, job_dir)

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
    # UI                                                                  #
    # ------------------------------------------------------------------ #

    @app.get("/", response_class=HTMLResponse)
    def ui_root() -> HTMLResponse:
        return HTMLResponse(_render_ui())

    return app


# ---------------------------------------------------------------------- #
# UI HTML minima                                                          #
# ---------------------------------------------------------------------- #


def _render_ui() -> str:
    return """<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>PPT Engine</title>
<style>
  body { font-family: system-ui, sans-serif; max-width: 720px; margin: 2rem auto; padding: 0 1rem; color: #222; }
  h1 { margin-bottom: 0.2rem; }
  .subtitle { color: #666; margin-top: 0; }
  form { background: #f7f7f9; padding: 1rem 1.25rem; border-radius: 8px; border: 1px solid #e3e3e8; }
  label { display: block; margin-top: 0.75rem; font-weight: 600; font-size: 0.9rem; }
  input[type=file], input[type=text] { display: block; width: 100%; margin-top: 0.25rem; padding: 0.4rem; box-sizing: border-box; }
  button { margin-top: 1rem; padding: 0.6rem 1.2rem; background: #2563eb; color: white; border: 0; border-radius: 6px; font-weight: 600; cursor: pointer; }
  button:hover { background: #1d4ed8; }
  button:disabled { background: #9ca3af; cursor: wait; }
  .tabs { display: flex; gap: 0.5rem; margin: 1rem 0 0.5rem; }
  .tab { padding: 0.4rem 0.9rem; border: 1px solid #d1d5db; border-radius: 6px; background: white; cursor: pointer; font-weight: 600; }
  .tab.active { background: #2563eb; color: white; border-color: #2563eb; }
  #result { margin-top: 1.5rem; padding: 1rem; background: #f0f9ff; border-left: 4px solid #2563eb; border-radius: 4px; white-space: pre-wrap; font-family: ui-monospace, monospace; font-size: 0.85rem; }
  .error { background: #fef2f2; border-left-color: #dc2626; }
  .ok { background: #f0fdf4; border-left-color: #16a34a; }
  a.download { display: inline-block; margin-top: 0.75rem; padding: 0.5rem 1rem; background: #16a34a; color: white; text-decoration: none; border-radius: 6px; font-weight: 600; }
</style>
</head>
<body>
  <h1>PPT Engine</h1>
  <p class="subtitle">Motor de generacion automatica de PowerPoint desde Excel</p>

  <div class="tabs">
    <button type="button" class="tab active" data-mode="build">Generar PPT</button>
    <button type="button" class="tab" data-mode="inspect">Inspeccionar</button>
    <button type="button" class="tab" data-mode="validate">Validar</button>
  </div>

  <form id="main-form" enctype="multipart/form-data">
    <label>Template .pptx <input type="file" name="template" accept=".pptx,.pptm" required></label>
    <div class="extras">
      <label>Datos .xlsx <input type="file" name="data" accept=".xlsx,.xlsm" required></label>
      <label>Mapping .yaml (opcional) <input type="file" name="mapping" accept=".yaml,.yml"></label>
      <label>Nombre de salida (opcional) <input type="text" name="output_name" placeholder="output_chile.pptx"></label>
    </div>
    <button type="submit" id="submit-btn">Generar</button>
  </form>

  <div id="result" style="display:none;"></div>

<script>
let currentMode = "build";
const tabs = document.querySelectorAll(".tab");
const form = document.getElementById("main-form");
const extras = document.querySelector(".extras");
const submitBtn = document.getElementById("submit-btn");
const result = document.getElementById("result");

tabs.forEach(t => t.addEventListener("click", () => {
  tabs.forEach(x => x.classList.remove("active"));
  t.classList.add("active");
  currentMode = t.dataset.mode;
  const isInspect = currentMode === "inspect";
  extras.style.display = isInspect ? "none" : "block";
  const dataInput = form.querySelector('input[name="data"]');
  dataInput.required = !isInspect;
  submitBtn.textContent = {
    build: "Generar",
    inspect: "Inspeccionar",
    validate: "Validar"
  }[currentMode];
  result.style.display = "none";
}));

form.addEventListener("submit", async (e) => {
  e.preventDefault();
  submitBtn.disabled = true;
  submitBtn.textContent = "Procesando...";
  result.style.display = "block";
  result.className = "";
  result.textContent = "Enviando...";

  const fd = new FormData(form);
  let url, method = "POST";
  if (currentMode === "build") url = "/api/jobs";
  else if (currentMode === "inspect") { url = "/api/inspect"; fd.delete("data"); fd.delete("mapping"); fd.delete("output_name"); }
  else url = "/api/validate";

  try {
    const resp = await fetch(url, { method, body: fd });
    const json = await resp.json();
    if (!resp.ok) {
      result.className = "error";
      result.textContent = "Error " + resp.status + ":\\n" + JSON.stringify(json, null, 2);
    } else {
      const success = json.ok !== false && json.status !== "error";
      result.className = success ? "ok" : "error";
      result.textContent = JSON.stringify(json, null, 2);
      if (currentMode === "build" && json.output_available && json.job_id) {
        const a = document.createElement("a");
        a.className = "download";
        a.href = "/api/jobs/" + json.job_id + "/download";
        a.textContent = "Descargar PPT";
        result.appendChild(document.createElement("br"));
        result.appendChild(a);
      }
    }
  } catch (err) {
    result.className = "error";
    result.textContent = "Error de red: " + err.message;
  } finally {
    submitBtn.disabled = false;
    submitBtn.textContent = {
      build: "Generar",
      inspect: "Inspeccionar",
      validate: "Validar"
    }[currentMode];
  }
});
</script>
</body>
</html>
"""


# ---------------------------------------------------------------------- #
# ASGI app por defecto (para uvicorn web.app:app)                         #
# ---------------------------------------------------------------------- #

app = create_app()

__all__ = ["create_app", "Job", "app"]
