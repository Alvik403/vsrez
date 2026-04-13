from __future__ import annotations

import json
from pathlib import Path
from urllib.parse import quote

from fastapi import FastAPI, File, HTTPException, Request, UploadFile
from fastapi.responses import HTMLResponse, JSONResponse, StreamingResponse
from fastapi.templating import Jinja2Templates

from app.config import DEFAULT_CONFIG_PATH, ROOT_DIR, load_app_config
from app.excel_service import ConfigError, consolidate_workbooks


app = FastAPI(title="Консолидация Excel в шаблон")
templates = Jinja2Templates(directory="templates")

TEST_TEMPLATE_PATH = ROOT_DIR / "Файлы" / "Шаблон ВСР.xlsm"
TEST_SOURCE_PATH = ROOT_DIR / "Файлы" / "РД С24 от 06.03.2026.xlsx"


@app.middleware("http")
async def add_no_store_headers(request: Request, call_next):
    response = await call_next(request)
    _apply_no_store_headers(response)
    return response


@app.get("/", response_class=HTMLResponse)
async def index(request: Request) -> HTMLResponse:
    return templates.TemplateResponse(
        request=request,
        name="index.html",
        context={
            "default_config_path": str(DEFAULT_CONFIG_PATH).replace("\\", "/"),
            "has_local_test_files": TEST_TEMPLATE_PATH.exists() and TEST_SOURCE_PATH.exists(),
        },
    )


@app.get("/api/config/default")
async def get_default_config() -> JSONResponse:
    try:
        config = load_app_config()
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    return JSONResponse(config)


@app.post("/api/consolidate")
async def consolidate(
    template_file: UploadFile | None = File(default=None),
    source_files: list[UploadFile] = File(...),
    config_file: UploadFile | None = File(default=None),
) -> StreamingResponse:
    try:
        config_payload = await config_file.read() if config_file is not None else None
        config = load_app_config(config_payload or None)
        template_name = template_file.filename if template_file is not None else None
        _validate_upload_names(template_name=template_name, source_files=source_files)
        template_bytes = await template_file.read() if template_file is not None else None
        source_payloads = [
            (upload.filename or f"source_{idx + 1}.xlsx", await upload.read())
            for idx, upload in enumerate(source_files)
        ]
        result_bytes, output_name, report = _run_consolidation(
            template_bytes=template_bytes,
            template_name=template_name,
            sources=source_payloads,
            config=config,
        )
    except (ValueError, ConfigError) as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Ошибка обработки Excel: {exc}") from exc
    finally:
        if template_file is not None:
            await template_file.close()
        if config_file is not None:
            await config_file.close()
        for upload in source_files:
            await upload.close()

    return _build_result_response(result_bytes=result_bytes, output_name=output_name, report=report)


@app.post("/api/consolidate/test")
async def consolidate_test() -> StreamingResponse:
    try:
        config = load_app_config()
        if not TEST_TEMPLATE_PATH.exists():
            raise HTTPException(status_code=400, detail=f"Тестовый шаблон не найден: {TEST_TEMPLATE_PATH}")
        if not TEST_SOURCE_PATH.exists():
            raise HTTPException(status_code=400, detail=f"Тестовый исходник не найден: {TEST_SOURCE_PATH}")

        result_bytes, output_name, report = _run_consolidation(
            template_bytes=TEST_TEMPLATE_PATH.read_bytes(),
            template_name=TEST_TEMPLATE_PATH.name,
            sources=[(TEST_SOURCE_PATH.name, TEST_SOURCE_PATH.read_bytes())],
            config=config,
        )
    except HTTPException:
        raise
    except (ValueError, ConfigError) as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Ошибка обработки Excel: {exc}") from exc

    return _build_result_response(result_bytes=result_bytes, output_name=output_name, report=report)


def _run_consolidation(
    template_bytes: bytes | None,
    template_name: str | None,
    sources: list[tuple[str, bytes]],
    config: dict,
) -> tuple[bytes, str, dict]:
    return consolidate_workbooks(
        template_bytes=template_bytes,
        template_name=template_name,
        sources=sources,
        config=config,
    )


def _validate_upload_names(template_name: str | None, source_files: list[UploadFile]) -> None:
    if template_name is not None and not _is_excel_name(template_name):
        raise HTTPException(status_code=400, detail="Шаблон должен быть .xlsx или .xlsm файлом.")

    if not source_files:
        raise HTTPException(status_code=400, detail="Нужно загрузить хотя бы один исходный Excel-файл.")

    invalid_sources = [
        upload.filename or "<без имени>"
        for upload in source_files
        if not _is_excel_name(upload.filename or "")
    ]
    if invalid_sources:
        raise HTTPException(
            status_code=400,
            detail=f"Исходные файлы должны быть .xlsx или .xlsm: {', '.join(invalid_sources)}.",
        )


def _build_result_response(result_bytes: bytes, output_name: str, report: dict) -> StreamingResponse:
    ascii_filename = _build_ascii_download_name(output_name)
    headers = {
        "Content-Disposition": (
            f'attachment; filename="{ascii_filename}"; '
            f"filename*=UTF-8''{quote(output_name)}"
        ),
        "X-Consolidation-Report": json.dumps(report, ensure_ascii=True),
    }
    media_type = (
        "application/vnd.ms-excel.sheet.macroEnabled.12"
        if output_name.lower().endswith(".xlsm")
        else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    return StreamingResponse(
        content=iter([result_bytes]),
        media_type=media_type,
        headers=headers,
    )


def _is_excel_name(filename: str) -> bool:
    lower = filename.lower()
    return lower.endswith(".xlsx") or lower.endswith(".xlsm")


def _build_ascii_download_name(filename: str) -> str:
    suffix = Path(filename).suffix.lower() or ".xlsx"
    stem = Path(filename).stem
    ascii_stem = "".join(ch if ch.isascii() and (ch.isalnum() or ch in "-_.") else "_" for ch in stem)
    ascii_stem = ascii_stem.strip("._") or "result"
    return f"{ascii_stem}{suffix}"


def _apply_no_store_headers(response) -> None:
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0, private"
    response.headers["Pragma"] = "no-cache"
    response.headers["Expires"] = "0"
