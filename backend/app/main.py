from __future__ import annotations

import json
import os
import shutil
import sys
import tempfile
from pathlib import Path
from typing import Annotated
from urllib.parse import quote
from uuid import uuid4

from fastapi import FastAPI, File, HTTPException, Query, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from backend.app.excel_processor import (
    SummaryExportRequest,
    SummaryConfig,
    detect_header_row,
    export_summary_results_workbook,
    export_summary_workbook,
    preview_workbook,
    summarize_workbook,
)


DATA_DIR = Path(tempfile.gettempdir()) / "excel-edit-tool"
UPLOAD_DIR = DATA_DIR / "uploads"
OUTPUT_DIR = DATA_DIR / "outputs"


class GuidePreference(BaseModel):
    guideHidden: bool = False


def create_app(static_dir: str | Path | None = None) -> FastAPI:
    app_instance = FastAPI(title="ExcelEditTool")
    app_instance.add_middleware(
        CORSMiddleware,
        allow_origins=["http://localhost:5173", "http://127.0.0.1:5173"],
        allow_credentials=True,
        allow_methods=["*"],
        allow_headers=["*"],
    )
    _register_api_routes(app_instance)

    resolved_static_dir = Path(static_dir) if static_dir else _default_static_dir()
    if resolved_static_dir.exists():
        app_instance.mount(
            "/",
            StaticFiles(directory=resolved_static_dir, html=True),
            name="frontend",
        )
    return app_instance


def _register_api_routes(app_instance: FastAPI) -> None:
    app_instance.post("/api/workbooks")(upload_workbook)
    app_instance.get("/api/workbooks/{file_id}/preview")(preview_sheet)
    app_instance.post("/api/workbooks/{file_id}/summary")(create_summary)
    app_instance.post("/api/workbooks/{file_id}/export")(export_summary)
    app_instance.post("/api/workbooks/{file_id}/export-results")(export_summary_results)
    app_instance.get("/api/preferences/guide")(read_guide_preference)
    app_instance.post("/api/preferences/guide")(save_guide_preference)


def _default_static_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS")) / "frontend" / "dist"
    return Path(__file__).resolve().parents[2] / "frontend" / "dist"


async def upload_workbook(file: Annotated[UploadFile, File()]) -> dict:
    if not file.filename or not file.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="请上传 .xlsx 文件")

    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    file_id = uuid4().hex
    workbook_path = UPLOAD_DIR / f"{file_id}.xlsx"
    with workbook_path.open("wb") as target:
        shutil.copyfileobj(file.file, target)

    metadata = {"originalFileName": file.filename}
    _metadata_path(file_id).write_text(json.dumps(metadata, ensure_ascii=False), encoding="utf-8")

    header_row = detect_header_row(workbook_path)
    preview = preview_workbook(workbook_path, header_row=header_row)

    return {
        "fileId": file_id,
        "fileName": file.filename,
        "sheets": preview.sheets,
        "preview": preview.model_dump(),
    }


async def preview_sheet(
    file_id: str,
    sheet_name: Annotated[str | None, Query(alias="sheetName")] = None,
    header_row: Annotated[int | None, Query(alias="headerRow", ge=1)] = None,
) -> dict:
    workbook_path = _workbook_path(file_id)
    if not workbook_path.exists():
        raise HTTPException(status_code=404, detail="文件不存在或已过期")

    resolved_header_row = header_row or detect_header_row(workbook_path, sheet_name)
    preview = preview_workbook(
        workbook_path,
        sheet_name=sheet_name,
        header_row=resolved_header_row,
    )
    return preview.model_dump()


async def create_summary(file_id: str, config: SummaryConfig) -> dict:
    workbook_path = _workbook_path(file_id)
    if not workbook_path.exists():
        raise HTTPException(status_code=404, detail="文件不存在或已过期")

    try:
        return summarize_workbook(workbook_path, config).model_dump()
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


async def export_summary(file_id: str, config: SummaryConfig) -> FileResponse:
    workbook_path = _workbook_path(file_id)
    if not workbook_path.exists():
        raise HTTPException(status_code=404, detail="文件不存在或已过期")

    try:
        result = summarize_workbook(workbook_path, config)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    metadata = _read_metadata(file_id)
    original_file_name = metadata["originalFileName"]
    output_path = export_summary_workbook(result, original_file_name, OUTPUT_DIR)
    download_name = output_path.name
    disposition = f"attachment; filename*=UTF-8''{quote(download_name)}"
    return FileResponse(
        output_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=download_name,
        headers={"Content-Disposition": disposition},
    )


async def export_summary_results(file_id: str, payload: SummaryExportRequest) -> FileResponse:
    workbook_path = _workbook_path(file_id)
    if not workbook_path.exists():
        raise HTTPException(status_code=404, detail="文件不存在或已过期")

    metadata = _read_metadata(file_id)
    original_file_name = metadata["originalFileName"]
    try:
        output_path = export_summary_results_workbook(payload.results, original_file_name, OUTPUT_DIR)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    download_name = output_path.name
    disposition = f"attachment; filename*=UTF-8''{quote(download_name)}"
    return FileResponse(
        output_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=download_name,
        headers={"Content-Disposition": disposition},
    )


async def read_guide_preference() -> dict:
    return _read_guide_preference().model_dump()


async def save_guide_preference(preference: GuidePreference) -> dict:
    path = _preferences_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(preference.model_dump_json(), encoding="utf-8")
    return preference.model_dump()


def _workbook_path(file_id: str) -> Path:
    return UPLOAD_DIR / f"{file_id}.xlsx"


def _metadata_path(file_id: str) -> Path:
    return UPLOAD_DIR / f"{file_id}.json"


def _preferences_path() -> Path:
    configured_dir = os.getenv("EXCEL_EDIT_TOOL_CONFIG_DIR")
    if configured_dir:
        return Path(configured_dir) / "preferences.json"

    root = Path(os.getenv("APPDATA") or Path.home())
    return root / "ExcelEditTool" / "preferences.json"


def _read_guide_preference() -> GuidePreference:
    path = _preferences_path()
    if not path.exists():
        return GuidePreference()
    try:
        return GuidePreference.model_validate_json(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, ValueError):
        return GuidePreference()


def _read_metadata(file_id: str) -> dict:
    metadata_path = _metadata_path(file_id)
    if not metadata_path.exists():
        raise HTTPException(status_code=404, detail="文件元数据不存在或已过期")
    return json.loads(metadata_path.read_text(encoding="utf-8"))


app = create_app()
