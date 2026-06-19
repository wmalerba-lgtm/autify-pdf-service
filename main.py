#!/usr/bin/env python3
"""
AUTIFY — PDF Generation Service
FastAPI wrapper for prenda (UVA/FIJA) and formulario 03 generators.
"""

import io
import json
import os
import tempfile
import zipfile
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, StreamingResponse

from generar_prenda_autify_v7_working import generar_pdf as generar_pdf_uva
from generar_prenda_fija_autify import generar_prenda_fija
from generar_form03_autify import generar_form03

# ── Static paths ────────────────────────────────────────────
BASE_DIR = Path(__file__).parent
STATIC_DIR = BASE_DIR / "static"
XLSX_PATH = STATIC_DIR / "parametros_contrato_prenda_v8.xlsx"
TEMPLATE_PRENDA_UVA  = STATIC_DIR / "template_prenda.pdf"
TEMPLATE_PRENDA_FIJA = STATIC_DIR / "template_prenda_fija.pdf"
TEMPLATE_FORM03      = STATIC_DIR / "template_form03.pdf"

VALID_TIPOS = {"UVA_PI", "UVA_PRE", "FIJA_PI", "FIJA_PRE"}

# ── App ──────────────────────────────────────────────────────
app = FastAPI(title="Autify PDF Service", version="2.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "https://admin.autify.com.ar",
        "http://localhost:3000",
    ],
    allow_methods=["GET", "POST"],
    allow_headers=["*"],
)


@app.get("/health")
def health():
    return {"status": "ok"}


@app.post("/generar")
async def generar(
    solicitud: UploadFile = File(..., description="Solicitud sin TyC (PDF)"),
    aprobado: UploadFile = File(..., description="Carta de aprobación (PDF)"),
    tipo_op: str = Form(..., description="UVA_PI | UVA_PRE | FIJA_PI | FIJA_PRE"),
    mutuo: Optional[UploadFile] = File(None, description="Mutuo prendario (solo _PRE)"),
    documentos: str = Form('["prenda", "form03"]', description='JSON array: ["prenda","form03"]'),
):
    # ── Validaciones ─────────────────────────────────────────
    tipo_op = tipo_op.strip().upper()
    if tipo_op not in VALID_TIPOS:
        raise HTTPException(400, f"tipo_op inválido: {tipo_op!r}. Debe ser uno de {sorted(VALID_TIPOS)}")

    if tipo_op.endswith("_PRE") and mutuo is None:
        raise HTTPException(400, "mutuo es requerido cuando tipo_op termina en _PRE")

    try:
        docs = json.loads(documentos)
        if not isinstance(docs, list):
            raise ValueError
        docs = [d.strip().lower() for d in docs]
        invalidos = set(docs) - {"prenda", "form03"}
        if invalidos:
            raise HTTPException(400, f"documentos inválidos: {invalidos}")
        if not docs:
            raise HTTPException(400, "documentos no puede estar vacío")
    except (json.JSONDecodeError, ValueError):
        raise HTTPException(400, 'documentos debe ser un JSON array, ej: ["prenda","form03"]')

    if not XLSX_PATH.exists():
        raise HTTPException(500, f"Archivo de configuración no encontrado: {XLSX_PATH}")

    es_fija = tipo_op.startswith("FIJA")
    template_prenda = TEMPLATE_PRENDA_FIJA if es_fija else TEMPLATE_PRENDA_UVA

    if "prenda" in docs and not template_prenda.exists():
        raise HTTPException(500, f"{template_prenda.name} no encontrado en static/")
    if "form03" in docs and not TEMPLATE_FORM03.exists():
        raise HTTPException(500, "template_form03.pdf no encontrado en static/")

    # ── Archivos temporales ───────────────────────────────────
    tmp_files = []

    def save_upload(upload: UploadFile) -> str:
        content = upload.file.read()
        suffix = Path(upload.filename or "file.pdf").suffix or ".pdf"
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
        tmp.write(content)
        tmp.flush()
        tmp.close()
        tmp_files.append(tmp.name)
        return tmp.name

    def make_output_tmp() -> str:
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
        tmp.close()
        tmp_files.append(tmp.name)
        return tmp.name

    try:
        solicitud_path = save_upload(solicitud)
        carta_path     = save_upload(aprobado)
        mutuo_path     = save_upload(mutuo) if mutuo is not None else None

        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", compression=zipfile.ZIP_DEFLATED) as zf:

            if "prenda" in docs:
                out_prenda = make_output_tmp()
                if es_fija:
                    generar_prenda_fija(
                        solicitud_path=solicitud_path,
                        template_path=str(template_prenda),
                        output_path=out_prenda,
                        xlsx_path=str(XLSX_PATH),
                        tipo_op=tipo_op,
                        carta_path=carta_path,
                    )
                else:
                    generar_pdf_uva(
                        solicitud_path=solicitud_path,
                        template_path=str(template_prenda),
                        output_path=out_prenda,
                        xlsx_path=str(XLSX_PATH),
                        tipo_op=tipo_op,
                        fecha_firma=None,
                        carta_path=carta_path,
                        mutuo_path=mutuo_path,
                    )
                zf.write(out_prenda, "contrato_prenda.pdf")

            if "form03" in docs:
                out_form03 = make_output_tmp()
                generar_form03(
                    solicitud_path=solicitud_path,
                    template_path=str(TEMPLATE_FORM03),
                    output_path=out_form03,
                    xlsx_path=str(XLSX_PATH),
                    tipo_op=tipo_op,
                    carta_path=carta_path,
                )
                zf.write(out_form03, "formulario_03.pdf")

        zip_buffer.seek(0)
        return StreamingResponse(
            zip_buffer,
            media_type="application/zip",
            headers={"Content-Disposition": "attachment; filename=documentos_prenda.zip"},
        )

    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(500, f"Error generando documentos: {exc}") from exc

    finally:
        for path in tmp_files:
            try:
                os.unlink(path)
            except OSError:
                pass
