"""Servidor HTTP (FastAPI) para convertir .docx -> PDF por la red local.

Reutiliza la API de conversión de la PR (varios motores). Es un módulo aparte
de ``api.py`` (que aquí es la API de conversión, no el servidor web).

    python serve.py            # escucha en 0.0.0.0:8000
    # luego abrir http://<IP-de-esta-maquina>:8000/docs

Variables de entorno:
  DOCX2PDF_ENGINE  motor por defecto (weasyprint|libreoffice|word|auto). Por
                   defecto 'weasyprint' (Python puro). Se puede sobreescribir
                   por petición con ?engine=...
  DOCX2PDF_SAVE    '1/true/yes/on' para GUARDAR cada .docx y su PDF (def. false)
  DOCX2PDF_SAVE_DIR  carpeta de guardado (def. ./recibidos)
"""
import os
import re
import tempfile
import time

from fastapi import FastAPI, File, HTTPException, Query, UploadFile
from fastapi.responses import FileResponse
from starlette.background import BackgroundTask

from .api import convert_detailed
from .engines import default_engine

DEFAULT_ENGINE = os.environ.get("DOCX2PDF_ENGINE", "weasyprint")
SAVE_ENABLED = os.environ.get("DOCX2PDF_SAVE", "false").lower() in ("1", "true", "yes", "on")
SAVE_DIR = os.environ.get(
    "DOCX2PDF_SAVE_DIR",
    os.path.join(os.path.dirname(os.path.dirname(__file__)), "recibidos"),
)

app = FastAPI(
    title="docx2pdf-py API",
    description="Convierte documentos .docx a PDF. Motor por defecto: "
                f"{DEFAULT_ENGINE} (configurable por ?engine= o DOCX2PDF_ENGINE).",
    version="0.2.0",
)


@app.get("/health")
def health():
    """Comprobación rápida + motor que elegiría 'auto' en esta máquina."""
    return {"status": "ok", "default_engine_auto": default_engine(),
            "engine_por_defecto": DEFAULT_ENGINE, "guardado": SAVE_ENABLED}


def _slug(nombre):
    base = os.path.basename(nombre)
    return re.sub(r"[^A-Za-z0-9._-]+", "_", base) or "documento.docx"


def _limpiar(*rutas):
    for ruta in rutas:
        try:
            os.remove(ruta)
        except OSError:
            pass


@app.post("/convert")
async def convert_endpoint(
    file: UploadFile = File(...),
    engine: str = Query(DEFAULT_ENGINE, description="weasyprint|libreoffice|word|auto"),
):
    """Recibe un .docx y devuelve el PDF. ?engine= elige el motor."""
    nombre = file.filename or "documento.docx"
    if not nombre.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="El fichero debe tener extensión .docx")

    if SAVE_ENABLED:
        os.makedirs(SAVE_DIR, exist_ok=True)
        base = time.strftime("%Y%m%d-%H%M%S") + "_" + _slug(nombre)
        ruta_docx = os.path.join(SAVE_DIR, base)
        ruta_pdf = os.path.join(SAVE_DIR, base[:-5] + ".pdf")
        cleanup = None
    else:
        fd, ruta_docx = tempfile.mkstemp(suffix=".docx")
        os.close(fd)
        ruta_pdf = ruta_docx[:-5] + ".pdf"
        cleanup = BackgroundTask(_limpiar, ruta_docx, ruta_pdf)

    try:
        with open(ruta_docx, "wb") as f:
            f.write(await file.read())
        result = convert_detailed(ruta_docx, ruta_pdf, engine=engine)
    except Exception as exc:  # noqa: BLE001 — devolvemos el error al cliente
        _limpiar(ruta_docx, ruta_pdf)
        raise HTTPException(status_code=500, detail=f"Error al convertir: {exc}") from exc

    return FileResponse(
        ruta_pdf,
        media_type="application/pdf",
        filename=nombre[:-5] + ".pdf",
        headers={"X-Engine": result.engine,
                 "X-Page-Count": str(result.page_count or "")},
        background=cleanup,
    )


def run(host="0.0.0.0", port=8000):
    """Arranca el servidor escuchando en toda la red local."""
    import uvicorn

    uvicorn.run("docx2pdf_py.webserver:app", host=host, port=port)


if __name__ == "__main__":
    run()
