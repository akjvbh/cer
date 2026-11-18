"""文档工坊：本机运行的多格式文档转换服务。"""

from __future__ import annotations

import logging
import mimetypes
import os
import platform
import re
import shutil
import sys
import threading
import time
import uuid
import webbrowser
from contextlib import asynccontextmanager
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from typing import Any
from urllib.parse import quote

import uvicorn
from dotenv import load_dotenv
from fastapi import FastAPI, File, Form, Request, UploadFile
from fastapi.exceptions import RequestValidationError
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from starlette.exceptions import HTTPException as StarletteHTTPException

import converter

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")

UPLOADS_DIR = BASE_DIR / "uploads"
OUTPUTS_DIR = BASE_DIR / "outputs"
STATIC_DIR = BASE_DIR / "static"
GLOSSARY_PATH = BASE_DIR / "glossary.txt"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(name)s: %(message)s",
)
LOGGER = logging.getLogger("docworkshop")


def _env_int(name: str, default: int) -> int:
    try:
        return int(os.getenv(name, str(default)))
    except (TypeError, ValueError):
        return default


MAX_FILE_SIZE_MB = max(1, _env_int("MAX_FILE_SIZE_MB", 200))
MAX_FILE_BYTES = MAX_FILE_SIZE_MB * 1024 * 1024
TASK_TTL_SECONDS = max(60, _env_int("TASK_TTL_SECONDS", 3600))
CLEANUP_INTERVAL_SECONDS = max(60, _env_int("CLEANUP_INTERVAL_SECONDS", 600))
VERSION = "0.1.0"

TASKS: dict[str, dict[str, Any]] = {}
TASK_LOCK = threading.RLock()
EXECUTOR = ThreadPoolExecutor(max_workers=4, thread_name_prefix="convert")
ILLEGAL_FILENAME = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


class ApiError(Exception):
    def __init__(self, status_code: int, message: str):
        super().__init__(message)
        self.status_code = status_code
        self.message = message


# ---------------------------------------------------------------------------
# 生命周期与异常
# ---------------------------------------------------------------------------

def _cleanup_once() -> None:
    cutoff = time.time() - TASK_TTL_SECONDS
    expired: list[str] = []
    with TASK_LOCK:
        for task_id, task in list(TASKS.items()):
            if task["created_at"] < cutoff and task["status"] in {"done", "error"}:
                expired.append(task_id)
        for task_id in expired:
            TASKS.pop(task_id, None)

    for task_id in expired:
        shutil.rmtree(OUTPUTS_DIR / task_id, ignore_errors=True)

    for path in UPLOADS_DIR.glob("*"):
        if path.is_file() and path.stat().st_mtime < cutoff:
            path.unlink(missing_ok=True)

    active_ids = set(TASKS)
    for path in OUTPUTS_DIR.glob("*"):
        if path.is_dir() and path.name not in active_ids and path.stat().st_mtime < cutoff:
            shutil.rmtree(path, ignore_errors=True)


def _cleanup_loop() -> None:
    while True:
        time.sleep(CLEANUP_INTERVAL_SECONDS)
        try:
            _cleanup_once()
        except Exception:
            LOGGER.exception("清理临时文件失败")


@asynccontextmanager
async def lifespan(app: FastAPI):
    UPLOADS_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUTS_DIR.mkdir(parents=True, exist_ok=True)
    _cleanup_once()
    threading.Thread(target=_cleanup_loop, daemon=True, name="temp-cleaner").start()
    LOGGER.info(
        "服务启动 | Python %s | Office %s | 监听 127.0.0.1:%s",
        platform.python_version(),
        "可用" if converter.office_available() else "不可用",
        os.getenv("PORT", "8000"),
    )
    yield
    EXECUTOR.shutdown(wait=False, cancel_futures=True)


app = FastAPI(title="文档工坊", version=VERSION, lifespan=lifespan)


@app.exception_handler(ApiError)
async def handle_api_error(request: Request, exc: ApiError):
    return JSONResponse(status_code=exc.status_code, content={"error": exc.message})


@app.exception_handler(RequestValidationError)
async def handle_validation_error(request: Request, exc: RequestValidationError):
    LOGGER.warning("请求参数校验失败：%s", exc.errors())
    return JSONResponse(status_code=400, content={"error": "请求参数不完整或格式错误"})


@app.exception_handler(StarletteHTTPException)
async def handle_http_error(request: Request, exc: StarletteHTTPException):
    if exc.status_code == 404:
        message = "请求的地址不存在"
    elif exc.status_code == 405:
        message = "请求方法不允许"
    else:
        message = str(exc.detail) if exc.detail else "请求失败"
    return JSONResponse(status_code=exc.status_code, content={"error": message})


@app.exception_handler(Exception)
async def handle_unexpected_error(request: Request, exc: Exception):
    LOGGER.exception("未捕获异常：%s %s", request.method, request.url.path)
    return JSONResponse(status_code=500, content={"error": "服务器内部错误"})


# ---------------------------------------------------------------------------
# 上传与任务辅助
# ---------------------------------------------------------------------------

def _extension(filename: str | None) -> str:
    name = Path(filename or "").name
    if "." not in name:
        return ""
    return name.rsplit(".", 1)[1].lower().lstrip(".")


def _repair_upload_filename(filename: str | None) -> str:
    """兼容 multipart 把 UTF-8 或 GBK 文件名按 latin-1 解码的情况。"""
    name = filename or ""
    if not name:
        return name
    try:
        raw = name.encode("latin-1")
    except UnicodeEncodeError:
        return name

    candidates: list[str] = []
    for encoding in ("utf-8", "gb18030"):
        try:
            decoded = raw.decode(encoding)
        except UnicodeDecodeError:
            continue
        if decoded and decoded != name:
            candidates.append(decoded)
    if not candidates:
        return name

    def score(value: str) -> tuple[int, int]:
        cjk = sum(1 for char in value if "\u4e00" <= char <= "\u9fff")
        bad = sum(1 for char in value if char == "\ufffd" or ord(char) < 32)
        return cjk, -bad

    best = max(candidates, key=score)
    return best if score(best)[0] > 0 else name


def _sanitize_filename(filename: str | None) -> str:
    name = Path(_repair_upload_filename(filename)).name
    name = ILLEGAL_FILENAME.sub("_", name).strip().strip(".")
    return (name or "未命名文件")[:120]


def _safe_stem(filename: str | None) -> str:
    stem = ILLEGAL_FILENAME.sub("_", Path(filename or "输出文件").stem).strip().strip(".")
    return (stem or "输出文件")[:80]


def _magic_ok(ext: str, head: bytes) -> bool:
    if ext == "pdf":
        return b"%PDF" in head[:1024]
    if ext in {"docx", "xlsx", "pptx"}:
        return head.startswith(b"PK\x03\x04")
    if ext in {"doc", "xls", "ppt"}:
        return head.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1")
    if ext == "png":
        return head.startswith(b"\x89PNG\r\n\x1a\n")
    if ext in {"jpg", "jpeg"}:
        return head.startswith(b"\xff\xd8\xff")
    if ext == "gif":
        return head.startswith((b"GIF87a", b"GIF89a"))
    if ext == "bmp":
        return head.startswith(b"BM")
    if ext == "tiff":
        return head.startswith((b"II*\x00", b"MM\x00*"))
    if ext == "webp":
        return len(head) >= 12 and head[:4] == b"RIFF" and head[8:12] == b"WEBP"
    if ext in {"txt", "md", "csv"}:
        return b"\x00" not in head[:4096]
    return True


async def _save_upload(upload: UploadFile, allowed_exts: set[str]) -> tuple[Path, str, int]:
    display = _sanitize_filename(upload.filename)
    ext = _extension(display)
    if not ext:
        raise ApiError(415, "文件缺少扩展名，无法判断格式")
    if ext not in allowed_exts:
        raise ApiError(415, f"不支持 .{ext} 文件，请检查支持的格式列表")

    UPLOADS_DIR.mkdir(parents=True, exist_ok=True)
    destination = UPLOADS_DIR / f"{uuid.uuid4().hex}.{ext}"
    total = 0
    head = bytearray()
    try:
        with destination.open("wb") as handle:
            while True:
                chunk = await upload.read(1024 * 1024)
                if not chunk:
                    break
                total += len(chunk)
                if total > MAX_FILE_BYTES:
                    raise ApiError(413, f"文件超过 {MAX_FILE_SIZE_MB}MB 上限")
                if len(head) < 8192:
                    head.extend(chunk[: 8192 - len(head)])
                handle.write(chunk)
    except ApiError:
        destination.unlink(missing_ok=True)
        raise
    except Exception as exc:
        destination.unlink(missing_ok=True)
        LOGGER.exception("保存上传文件失败")
        raise ApiError(400, "文件上传失败，请重试") from exc
    finally:
        await upload.close()

    if total == 0:
        destination.unlink(missing_ok=True)
        raise ApiError(400, "上传文件为空")
    if not _magic_ok(ext, bytes(head)):
        destination.unlink(missing_ok=True)
        raise ApiError(415, f"文件内容与 .{ext} 扩展名不匹配")
    return destination, display, total


def _new_task(task_id: str, display: str, size: int) -> dict[str, Any]:
    task = {
        "task_id": task_id,
        "status": "pending",
        "progress": 0,
        "message": "排队中",
        "error": "",
        "result": None,
        "created_at": time.time(),
        "source_name": display,
        "source_size": size,
        "_path": None,
        "_download_name": None,
    }
    with TASK_LOCK:
        TASKS[task_id] = task
    return task


def _update_task(task_id: str, **fields: Any) -> None:
    with TASK_LOCK:
        task = TASKS.get(task_id)
        if task is not None:
            task.update(fields)


def _public_task(task: dict[str, Any]) -> dict[str, Any]:
    return {
        "task_id": task["task_id"],
        "status": task["status"],
        "progress": task["progress"],
        "message": task["message"],
        "error": task["error"],
        "result": task["result"],
        "created_at": task["created_at"],
        "source_name": task["source_name"],
        "source_size": task["source_size"],
    }


def _finalize_output(output: Path, display: str, target: str, task_id: str) -> Path:
    stem = _safe_stem(display)
    suffix = output.suffix.lower()
    if suffix == ".zip" and target == "jpg":
        name = f"{stem}_jpg.zip"
    elif suffix == ".zip":
        name = f"{stem}.zip"
    else:
        name = f"{stem}{suffix}"
    final = output.parent / name
    if final.exists():
        final = output.parent / f"{stem}_{task_id[:6]}{suffix}"
    output.replace(final)
    return final


def _run_conversion(task_id: str, src: Path, display: str, target: str) -> None:
    _update_task(task_id, status="running", progress=5, message="正在准备转换")

    def progress(value: int, message: str) -> None:
        _update_task(task_id, progress=max(5, min(95, int(value))), message=message)

    try:
        output, note = converter.convert(src, OUTPUTS_DIR / task_id, target, progress)
        final = _finalize_output(output, display, target, task_id)
        result = {
            "filename": final.name,
            "size": final.stat().st_size,
            "download_url": f"/api/download/{task_id}",
            "note": note,
            "target": target,
        }
        _update_task(
            task_id,
            status="done",
            progress=100,
            message="转换完成",
            result=result,
            _path=str(final),
            _download_name=final.name,
        )
    except converter.ConversionError as exc:
        LOGGER.warning("转换失败 task=%s: %s", task_id, exc)
        shutil.rmtree(OUTPUTS_DIR / task_id, ignore_errors=True)
        _update_task(task_id, status="error", progress=100, message="转换失败", error=str(exc))
    except Exception:
        LOGGER.exception("转换出现未预期错误 task=%s", task_id)
        shutil.rmtree(OUTPUTS_DIR / task_id, ignore_errors=True)
        _update_task(
            task_id,
            status="error",
            progress=100,
            message="转换失败",
            error="转换失败，请检查文件是否损坏或稍后重试",
        )
    finally:
        src.unlink(missing_ok=True)


# ---------------------------------------------------------------------------
# 接口
# ---------------------------------------------------------------------------

@app.get("/")
def index():
    return FileResponse(STATIC_DIR / "index.html")


@app.head("/")
def index_head():
    return FileResponse(STATIC_DIR / "index.html")


@app.get("/api/health")
def health():
    return {
        "status": "ok",
        "office": converter.office_available(),
        "python": platform.python_version(),
        "version": VERSION,
    }


@app.head("/api/health")
def health_head():
    return health()


@app.get("/api/config")
def config():
    glossary = ""
    if GLOSSARY_PATH.exists():
        try:
            glossary = GLOSSARY_PATH.read_text(encoding="utf-8-sig")
        except Exception:
            LOGGER.warning("读取术语表失败", exc_info=True)
    return {
        "max_file_size_mb": MAX_FILE_SIZE_MB,
        "translate_configured": bool(os.getenv("TRANSLATE_API_KEY", "").strip()),
        "glossary": glossary,
    }


@app.get("/api/formats")
def formats():
    return converter.get_formats()


@app.post("/api/convert")
async def convert_api(file: UploadFile = File(...), target: str = Form(...)):
    requested = target.strip().lower().lstrip(".")
    display = _sanitize_filename(file.filename)
    ext = _extension(display)
    all_formats = converter.get_formats()
    if ext not in all_formats:
        raise ApiError(415, f"不支持 .{ext or '未知'} 文件，请检查支持的格式列表")
    if requested not in all_formats[ext]:
        raise ApiError(400, f"不支持从 .{ext} 转换到 {requested}")
    src, display, size = await _save_upload(file, set(all_formats))
    task_id = uuid.uuid4().hex
    _new_task(task_id, display, size)
    EXECUTOR.submit(_run_conversion, task_id, src, display, requested)
    return {
        "task_id": task_id,
        "status": "pending",
        "filename": display,
        "size": size,
        "download_url": f"/api/download/{task_id}",
        "note": "",
    }


@app.get("/api/tasks/{task_id}")
def task_status(task_id: str):
    with TASK_LOCK:
        task = TASKS.get(task_id)
        if task is None:
            raise ApiError(404, "任务不存在或已过期")
        return _public_task(task)


@app.get("/api/download/{task_id}")
def download(task_id: str):
    with TASK_LOCK:
        task = TASKS.get(task_id)
        if task is None:
            raise ApiError(404, "任务不存在或结果已过期")
        if task["status"] == "error":
            raise ApiError(409, task["error"] or "任务转换失败")
        if task["status"] != "done":
            raise ApiError(409, "任务尚未完成，请稍后再试")
        path = Path(task["_path"]) if task["_path"] else None
        filename = task["_download_name"] or (path.name if path else "download")
    if path is None or not path.exists():
        raise ApiError(404, "结果文件已过期或被清理")

    media_type = mimetypes.guess_type(filename)[0] or "application/octet-stream"
    ascii_name = f"download{path.suffix}"
    encoded = quote(filename)
    headers = {
        "Content-Disposition": f"attachment; filename=\"{ascii_name}\"; filename*=UTF-8''{encoded}"
    }
    return FileResponse(path, media_type=media_type, headers=headers)


if STATIC_DIR.exists():
    app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")


def _open_browser_later(port: int) -> None:
    time.sleep(1.5)
    try:
        webbrowser.open(f"http://127.0.0.1:{port}/")
    except Exception:
        LOGGER.warning("自动打开浏览器失败，请手动访问 http://127.0.0.1:%s/", port)


if __name__ == "__main__":
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass
    port = _env_int("PORT", 8000)
    if os.getenv("AUTO_OPEN_BROWSER", "0") == "1":
        threading.Thread(target=_open_browser_later, args=(port,), daemon=True).start()
    uvicorn.run(app, host="127.0.0.1", port=port, log_level="info")