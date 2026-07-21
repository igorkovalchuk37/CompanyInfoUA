from __future__ import annotations

import base64
import io
import json
import os
import tempfile
from pathlib import Path
from typing import Any

from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel, Field

from company_reporting import ReportConfig, generate_report_bundle
from ua_business_map_release.render_business_map import render as render_business_map


BASE_DIR = Path(__file__).resolve().parent
MAP_DIR = BASE_DIR / "ua_business_map_release"
DOCX_DIR = BASE_DIR / "CompanyInfo_UA_docx_setup"


def read_text(path: Path) -> str:
    for encoding in ("utf-8", "utf-8-sig", "cp1251"):
        try:
            return path.read_text(encoding=encoding)
        except UnicodeDecodeError:
            continue
    return path.read_text(encoding="utf-8", errors="replace")


def build_agent_context() -> dict[str, Any]:
    docx_instruction = read_text(DOCX_DIR / "CompanyInfo_UA_instruction_optimized_v2.txt")
    map_instruction = read_text(MAP_DIR / "AGENT_INSTRUCTION.txt")
    manifest = json.loads((MAP_DIR / "ua_business_map_manifest.json").read_text(encoding="utf-8"))
    map_example = json.loads((MAP_DIR / "business_group_map_data.example.json").read_text(encoding="utf-8"))

    return {
        "project": "CompanyInfo_UA",
        "docx_template": "CompanyInfo_UA_report_template.docx",
        "instructions": {
            "docx": docx_instruction,
            "map": map_instruction,
        },
        "map_assets": {
            "manifest_version": manifest.get("version"),
            "supported_asset_types": sorted(manifest.get("asset_types", {}).keys()),
            "example_payload": map_example,
        },
    }


class MapRenderRequest(BaseModel):
    data: dict[str, Any] = Field(..., description="Map JSON payload compatible with business_group_map_data.example.json")
    svg_filename: str = Field(default="business_map.svg")
    png_filename: str = Field(default="business_map.png")
    include_png: bool = Field(default=True)


class ReportGenerateRequest(BaseModel):
    edrpou: str = Field(..., min_length=8, max_length=10, description="ЄДРПОУ компанії")
    company_name_hint: str | None = Field(default=None, description="Необов'язкова підказка по назві компанії")
    additional_instructions: str | None = Field(default=None, description="Додаткові уточнення для серверного агента")


class ReportGenerateResponse(BaseModel):
    filename: str
    docx_base64: str
    map_png_base64: str
    report_data: dict[str, Any]


app = FastAPI(
    title="CompanyInfo UA Agent Assets",
    version="0.1.0",
    description="Render-ready host for CompanyInfo UA report instructions, templates, and business-map rendering.",
)

app.mount("/static", StaticFiles(directory=str(BASE_DIR), html=False), name="static")


@app.get("/")
def root() -> dict[str, Any]:
    return {
        "service": "companyinfo-ua-agent-assets",
        "status": "ok",
        "endpoints": {
            "health": "/health",
            "agent_context": "/api/agent-context",
            "render_map": "/api/render-map",
            "generate_report": "/api/generate-report",
            "generate_report_file": "/api/generate-report-file",
            "docx_template": "/downloads/docx-template",
        },
    }


@app.get("/health")
def health() -> dict[str, str]:
    return {
        "status": "ok",
        "openai_api_key_configured": "true" if bool(os.getenv("OPENAI_API_KEY")) else "false",
    }


@app.get("/api/agent-context")
def agent_context() -> dict[str, Any]:
    return build_agent_context()


@app.get("/downloads/docx-template")
def download_docx_template() -> FileResponse:
    path = DOCX_DIR / "CompanyInfo_UA_report_template.docx"
    if not path.exists():
        raise HTTPException(status_code=404, detail="DOCX template not found")
    return FileResponse(path, filename=path.name, media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


@app.post("/api/render-map")
def render_map(payload: MapRenderRequest) -> dict[str, Any]:
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_path = Path(tmp_dir)
        data_path = tmp_path / "map_data.json"
        svg_path = tmp_path / payload.svg_filename
        png_path = tmp_path / payload.png_filename if payload.include_png else None

        data_path.write_text(json.dumps(payload.data, ensure_ascii=False, indent=2), encoding="utf-8")

        try:
            render_business_map(
                data_path=data_path,
                output_svg=svg_path,
                template_path=MAP_DIR / "ua_business_map_template.svg",
                manifest_path=MAP_DIR / "ua_business_map_manifest.json",
                output_png=png_path,
            )
        except Exception as exc:
            raise HTTPException(status_code=400, detail=str(exc)) from exc

        result = {
            "svg_filename": svg_path.name,
            "svg_base64": base64.b64encode(svg_path.read_bytes()).decode("ascii"),
        }
        if png_path and png_path.exists():
            result["png_filename"] = png_path.name
            result["png_base64"] = base64.b64encode(png_path.read_bytes()).decode("ascii")
        return result


@app.post("/api/generate-report", response_model=ReportGenerateResponse)
def generate_report(payload: ReportGenerateRequest) -> ReportGenerateResponse:
    try:
        bundle = generate_report_bundle(
            edrpou=payload.edrpou.strip(),
            company_name_hint=payload.company_name_hint,
            additional_instructions=payload.additional_instructions,
            base_dir=BASE_DIR,
            config=ReportConfig(),
        )
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    return ReportGenerateResponse(
        filename=bundle["filename"],
        docx_base64=bundle["docx_base64"],
        map_png_base64=bundle["map_png_base64"],
        report_data=bundle["report_data"],
    )


@app.post("/api/generate-report-file")
def generate_report_file(payload: ReportGenerateRequest) -> StreamingResponse:
    try:
        bundle = generate_report_bundle(
            edrpou=payload.edrpou.strip(),
            company_name_hint=payload.company_name_hint,
            additional_instructions=payload.additional_instructions,
            base_dir=BASE_DIR,
            config=ReportConfig(),
        )
    except Exception as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    headers = {
        "Content-Disposition": f"attachment; filename*=UTF-8''{bundle['download_name']}",
    }
    return StreamingResponse(
        io.BytesIO(bundle["docx_bytes"]),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers=headers,
    )
