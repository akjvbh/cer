"""文档工坊：本机运行的多格式文档转换服务。"""

from __future__ import annotations

import json
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
import zipfile
from contextlib import asynccontextmanager
from concurrent.futures import ThreadPoolExecutor, as_completed
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
import pdf_tools
import scanner
import translator

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")

UPLOADS_DIR = BASE_DIR / "uploads"
OUTPUTS_DIR = BASE_DIR / "outputs"
STATIC_DIR = BASE_DIR / "static"

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
MAX_BATCH_FILES = max(1, _env_int("MAX_BATCH_FILES", 200))
MAX_BATCH_SIZE_MB = max(1, _env_int("MAX_BATCH_SIZE_MB", 500))
MAX_BATCH_BYTES = MAX_BATCH_SIZE_MB * 1024 * 1024
TASK_TTL_SECONDS = max(60, _env_int("TASK_TTL_SECONDS", 3600))
CLEANUP_INTERVAL_SECONDS = max(60, _env_int("CLEANUP_INTERVAL_SECONDS", 600))
VERSION = "0.4.0"

TASKS: dict[str, dict[str, Any]] = {}
TASK_LOCK = threading.RLock()
EXECUTOR = ThreadPoolExecutor(max_workers=4, thread_name_prefix="convert")
BATCH_EXECUTOR = ThreadPoolExecutor(max_workers=4, thread_name_prefix="batch")
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
    BATCH_EXECUTOR.shutdown(wait=False, cancel_futures=True)


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


async def _save_upload(
    upload: UploadFile,
    allowed_exts: set[str],
    max_bytes: int | None = None,
    limit_message: str | None = None,
) -> tuple[Path, str, int]:
    display = _sanitize_filename(upload.filename)
    ext = _extension(display)
    if not ext:
        raise ApiError(415, "文件缺少扩展名，无法判断格式")
    if ext not in allowed_exts:
        raise ApiError(415, f"不支持 .{ext} 文件，请检查支持的格式列表")

    limit = MAX_FILE_BYTES if max_bytes is None else max(1, max_bytes)
    over_limit_message = limit_message or f"文件超过 {MAX_FILE_SIZE_MB}MB 上限"
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
                if total > limit:
                    raise ApiError(413, over_limit_message)
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


def _new_task(task_id: str, display: str, size: int, **extra: Any) -> dict[str, Any]:
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
        "_batch_files": {},
    }
    task.update(extra)
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


def _rename_zip_members(path: Path, old_prefix: str, new_prefix: str) -> None:
    """把 ZIP 内以临时文件名为前缀的成员改成用户文件名，避免暴露 UUID。"""
    if not old_prefix or old_prefix == new_prefix:
        return
    temp = path.with_name(f".{path.stem}_rename.zip")
    try:
        rewritten = False
        with zipfile.ZipFile(path) as source:
            infos = source.infolist()
            if not any(info.filename.startswith(f"{old_prefix}_") for info in infos):
                return
            with zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED) as target:
                for info in infos:
                    name = info.filename
                    if name.startswith(f"{old_prefix}_"):
                        name = f"{new_prefix}{name[len(old_prefix):]}"
                    if info.is_dir():
                        target.writestr(name, b"")
                        continue
                    target_info = zipfile.ZipInfo(name, date_time=info.date_time)
                    target_info.compress_type = zipfile.ZIP_DEFLATED
                    target_info.external_attr = info.external_attr
                    with source.open(info) as reader, target.open(target_info, "w") as writer:
                        shutil.copyfileobj(reader, writer)
            rewritten = True
        if rewritten:
            temp.replace(path)
    finally:
        temp.unlink(missing_ok=True)


def _finalize_output(output: Path, display: str, target: str, task_id: str) -> Path:
    stem = _safe_stem(display)
    suffix = output.suffix.lower()
    if suffix == ".zip":
        _rename_zip_members(output, output.stem, stem)
    if suffix == ".zip" and target in {"jpg", "png"}:
        name = f"{stem}_{target}.zip"
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


def _batch_result(
    filename: str,
    ok: bool,
    download_url: str = "",
    error: str = "",
    note: str = "",
) -> dict[str, Any]:
    return {
        "filename": filename,
        "ok": ok,
        "download_url": download_url,
        "error": error,
        "note": note,
    }


def _batch_convert_one(entry: dict[str, Any], task_id: str, target: str) -> dict[str, Any]:
    index = entry["index"]
    try:
        output, note = converter.convert(
            entry["src"],
            OUTPUTS_DIR / task_id / f"{index:03d}",
            target,
        )
        final = _finalize_output(output, entry["display"], target, f"{task_id}{index}")
        return {**entry, "path": final, "note": note, "error": ""}
    except converter.ConversionError as exc:
        LOGGER.warning("批量转换失败 task=%s index=%s: %s", task_id, index, exc)
        return {**entry, "path": None, "note": "", "error": str(exc)}
    except Exception:
        LOGGER.exception("批量转换出现未预期错误 task=%s index=%s", task_id, index)
        return {
            **entry,
            "path": None,
            "note": "",
            "error": "转换失败，请检查文件是否损坏或稍后重试",
        }
    finally:
        entry["src"].unlink(missing_ok=True)


def _unique_arcname(name: str, used: set[str]) -> str:
    if name not in used:
        used.add(name)
        return name
    stem = Path(name).stem
    suffix = Path(name).suffix
    counter = 2
    while f"{stem} ({counter}){suffix}" in used:
        counter += 1
    candidate = f"{stem} ({counter}){suffix}"
    used.add(candidate)
    return candidate


def _build_batch_zip(task_id: str, entries: list[dict[str, Any]]) -> Path | None:
    zip_path = OUTPUTS_DIR / task_id / "批量转换结果.zip"
    used: set[str] = set()
    written = 0
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as archive:
        for entry in entries:
            path = entry.get("path")
            if path is None or not Path(path).exists():
                continue
            archive.write(path, _unique_arcname(Path(path).name, used))
            written += 1
    if written == 0:
        zip_path.unlink(missing_ok=True)
        return None
    return zip_path


def _run_batch(
    task_id: str,
    saved: list[dict[str, Any]],
    total_files: int,
    target: str,
    skipped: list[dict[str, Any]],
) -> None:
    _update_task(
        task_id,
        status="running",
        progress=2,
        message=f"正在处理 0/{total_files} 个文件",
    )
    results: list[dict[str, Any] | None] = [None] * total_files
    for item in skipped:
        results[item["index"]] = _batch_result(item["filename"], False, error=item["error"])

    done_entries: list[dict[str, Any]] = []
    completed = len(skipped)
    try:
        futures = {
            BATCH_EXECUTOR.submit(_batch_convert_one, entry, task_id, target): entry
            for entry in saved
        }
        for future in as_completed(futures):
            entry = future.result()
            path = entry.get("path")
            results[entry["index"]] = _batch_result(
                entry["display"],
                path is not None,
                f"/api/download/{task_id}?file={entry['index']}" if path is not None else "",
                entry["error"],
                entry["note"],
            )
            if path is not None:
                done_entries.append(entry)
            completed += 1
            progress = min(98, 2 + int(completed * 96 / total_files))
            _update_task(
                task_id,
                progress=progress,
                message=f"已完成 {completed}/{total_files} 个文件",
            )
    except Exception:
        LOGGER.exception("批量任务异常 task=%s", task_id)
        for entry in saved:
            entry["src"].unlink(missing_ok=True)
        shutil.rmtree(OUTPUTS_DIR / task_id, ignore_errors=True)
        _update_task(
            task_id,
            status="error",
            progress=100,
            message="批量处理失败",
            error="批量处理失败，请重试或减少文件数量",
        )
        return

    actual_results = [item for item in results if item is not None]
    success_files = sum(1 for item in actual_results if item["ok"])
    failed_files = total_files - success_files

    zip_path: Path | None = None
    if done_entries:
        try:
            zip_path = _build_batch_zip(task_id, done_entries)
        except Exception:
            LOGGER.exception("批量结果打包失败 task=%s", task_id)

    batch_files = {
        entry["index"]: (str(entry["path"]), Path(entry["path"]).name)
        for entry in done_entries
        if entry.get("path") is not None
    }
    result = {
        "total_files": total_files,
        "success_files": success_files,
        "failed_files": failed_files,
        "results": actual_results,
    }
    message = f"批量完成：成功 {success_files}，失败 {failed_files}"
    _update_task(
        task_id,
        status="done",
        progress=100,
        message=message,
        result=result,
        _path=str(zip_path) if zip_path is not None else None,
        _download_name=zip_path.name if zip_path is not None else None,
        _batch_files=batch_files,
    )


# ---------------------------------------------------------------------------
# 压缩与 PDF 工具
# ---------------------------------------------------------------------------

def _json_arg(value: str | None, label: str):
    text = (value or "").strip()
    if not text:
        return None
    try:
        return json.loads(text)
    except json.JSONDecodeError as exc:
        raise ApiError(400, f"{label}格式错误，应为 JSON") from exc


def _json_list_arg(value: str | None, label: str) -> list:
    data = _json_arg(value, label)
    if data is None:
        return []
    if not isinstance(data, list):
        raise ApiError(400, f"{label}应为 JSON 数组")
    return data


def _json_dict_arg(value: str | None, label: str) -> dict:
    data = _json_arg(value, label)
    if data is None:
        return {}
    if not isinstance(data, dict):
        raise ApiError(400, f"{label}应为 JSON 对象")
    return data


def _discard_uploads(paths: list[Path]) -> None:
    for path in paths:
        try:
            path.unlink(missing_ok=True)
        except OSError:
            LOGGER.debug("删除上传文件失败：%s", path, exc_info=True)


def _check_pdf(src: Path, display: str) -> int:
    try:
        return pdf_tools.validate_pdf(src)
    except converter.ConversionError as exc:
        src.unlink(missing_ok=True)
        raise ApiError(400, f"文件「{display}」：{exc}") from None


def _rename_output(output: Path, base: str, task_id: str) -> Path:
    suffix = output.suffix.lower()
    final = output.parent / f"{base}{suffix}"
    if final.exists():
        final = output.parent / f"{base}_{task_id[:6]}{suffix}"
    output.replace(final)
    return final


def _finish_output_task(
    task_id: str,
    output: Path,
    base: str,
    note: str,
    target: str,
    message: str,
) -> Path:
    final = _rename_output(output, base, task_id)
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
        message=message,
        result=result,
        _path=str(final),
        _download_name=final.name,
    )
    return final


def _fail_task(task_id: str, message: str, error: str) -> None:
    shutil.rmtree(OUTPUTS_DIR / task_id, ignore_errors=True)
    _update_task(task_id, status="error", progress=100, message=message, error=error)


def _start_task_thread(target, *args) -> None:
    threading.Thread(target=target, args=args, daemon=True).start()


def _run_compress(task_id: str, src: Path, display: str, level: str) -> None:
    _update_task(task_id, status="running", progress=5, message="正在准备压缩")

    def progress(value: int, message: str) -> None:
        _update_task(task_id, progress=max(5, min(95, int(value))), message=message)

    try:
        output, note = converter.compress_file(src, OUTPUTS_DIR / task_id, level, progress)
        _finish_output_task(
            task_id,
            output,
            f"{_safe_stem(display)}_压缩",
            note,
            output.suffix.lower().lstrip("."),
            "压缩完成",
        )
    except converter.ConversionError as exc:
        LOGGER.warning("压缩失败 task=%s: %s", task_id, exc)
        _fail_task(task_id, "压缩失败", str(exc))
    except Exception:
        LOGGER.exception("压缩出现未预期错误 task=%s", task_id)
        _fail_task(task_id, "压缩失败", "压缩失败，请检查文件是否损坏或稍后重试")
    finally:
        src.unlink(missing_ok=True)


def _run_pdf_merge(
    task_id: str,
    files: list[tuple[Path, str]],
    mode: str,
    insert_index: int,
) -> None:
    _update_task(task_id, status="running", progress=5, message="正在读取 PDF")

    def progress(value: int, message: str) -> None:
        _update_task(task_id, progress=max(5, min(95, int(value))), message=message)

    try:
        output, note = pdf_tools.merge(files, OUTPUTS_DIR / task_id, mode, insert_index, progress)
        _finish_output_task(task_id, output, "合并结果", note, "pdf", "合并完成")
    except converter.ConversionError as exc:
        LOGGER.warning("PDF 合并失败 task=%s: %s", task_id, exc)
        _fail_task(task_id, "合并失败", str(exc))
    except Exception:
        LOGGER.exception("PDF 合并出现未预期错误 task=%s", task_id)
        _fail_task(task_id, "合并失败", "合并失败，请检查文件是否损坏或稍后重试")
    finally:
        for path, _ in files:
            path.unlink(missing_ok=True)


def _run_pdf_ops(
    task_id: str,
    src: Path,
    display: str,
    delete_pages: list,
    rotations: dict,
    page_order: list,
) -> None:
    _update_task(task_id, status="running", progress=5, message="正在读取 PDF")

    def progress(value: int, message: str) -> None:
        _update_task(task_id, progress=max(5, min(95, int(value))), message=message)

    try:
        output, note = pdf_tools.page_ops(
            src,
            OUTPUTS_DIR / task_id,
            delete_pages,
            rotations,
            page_order,
            progress,
        )
        _finish_output_task(
            task_id,
            output,
            f"{_safe_stem(display)}_页面操作",
            note,
            "pdf",
            "页面处理完成",
        )
    except converter.ConversionError as exc:
        LOGGER.warning("PDF 页面操作失败 task=%s: %s", task_id, exc)
        _fail_task(task_id, "页面处理失败", str(exc))
    except Exception:
        LOGGER.exception("PDF 页面操作出现未预期错误 task=%s", task_id)
        _fail_task(task_id, "页面处理失败", "页面处理失败，请检查文件是否损坏或稍后重试")
    finally:
        src.unlink(missing_ok=True)


def _run_pdf_to_jpg(
    task_id: str,
    files: list[tuple[Path, str]],
    max_kb: int,
) -> None:
    _update_task(
        task_id,
        status="running",
        progress=3,
        message=f"正在导出 0/{len(files)} 个文件",
    )

    def progress(value: int, message: str) -> None:
        _update_task(task_id, progress=max(5, min(95, int(value))), message=message)

    try:
        zip_path, note, failures = pdf_tools.export_jpg(
            files,
            OUTPUTS_DIR / task_id,
            max_kb,
            progress,
        )
        success = len(files) - len(failures)
        base = f"{_safe_stem(files[0][1])}_JPG" if len(files) == 1 else "PDF转JPG"
        final = _rename_output(zip_path, base, task_id)
        results = [
            {
                "filename": item["filename"],
                "ok": False,
                "download_url": "",
                "error": item["error"],
                "note": "",
            }
            for item in failures
        ]
        result = {
            "filename": final.name,
            "size": final.stat().st_size,
            "download_url": f"/api/download/{task_id}",
            "note": note,
            "target": "jpg",
            "total_files": len(files),
            "success_files": success,
            "failed_files": len(failures),
            "results": results,
        }
        _update_task(
            task_id,
            status="done",
            progress=100,
            message=f"JPG 导出完成：成功 {success}，失败 {len(failures)}",
            result=result,
            _path=str(final),
            _download_name=final.name,
        )
    except converter.ConversionError as exc:
        LOGGER.warning("JPG 导出失败 task=%s: %s", task_id, exc)
        _fail_task(task_id, "JPG 导出失败", str(exc))
    except Exception:
        LOGGER.exception("JPG 导出出现未预期错误 task=%s", task_id)
        _fail_task(task_id, "JPG 导出失败", "JPG 导出失败，请检查文件是否损坏或稍后重试")
    finally:
        for path, _ in files:
            path.unlink(missing_ok=True)


# ---------------------------------------------------------------------------
# AI 增强
# ---------------------------------------------------------------------------

def _run_translate(
    task_id: str,
    src: Path,
    display: str,
    mode: str,
    target_lang: str,
    glossary: str,
) -> None:
    _update_task(task_id, status="running", progress=1, message="正在提取段落")

    def progress(value: int, message: str) -> None:
        _update_task(task_id, progress=max(1, min(95, int(value))), message=message)

    try:
        output, note = translator.translate_document(
            src,
            display,
            OUTPUTS_DIR / task_id,
            mode,
            target_lang,
            glossary,
            progress,
        )
        stem = _safe_stem(display)
        base = {
            "bilingual": f"{stem}_中英对照",
            "translated": f"{stem}_译文",
            "layout": f"{stem}_版式译文",
        }[mode]
        _finish_output_task(
            task_id,
            output,
            base,
            note,
            output.suffix.lower().lstrip("."),
            "翻译完成",
        )
    except (translator.TranslateError, converter.ConversionError) as exc:
        LOGGER.warning("翻译失败 task=%s: %s", task_id, exc)
        _fail_task(task_id, "翻译失败", str(exc))
    except Exception:
        LOGGER.exception("翻译出现未预期错误 task=%s", task_id)
        _fail_task(task_id, "翻译失败", "翻译失败，请检查文件是否损坏或稍后重试")
    finally:
        src.unlink(missing_ok=True)


def _run_scan(task_id: str, src: Path, display: str) -> None:
    _update_task(task_id, status="running", progress=5, message="正在读取图片")

    def progress(value: int, message: str) -> None:
        _update_task(task_id, progress=max(5, min(98, int(value))), message=message)

    try:
        output, note = scanner.enhance(src, OUTPUTS_DIR / task_id, progress)
        _finish_output_task(
            task_id,
            output,
            f"{_safe_stem(display)}_扫描增强",
            note,
            "png",
            "扫描件增强完成",
        )
    except converter.ConversionError as exc:
        LOGGER.warning("扫描件增强失败 task=%s: %s", task_id, exc)
        _fail_task(task_id, "扫描件增强失败", str(exc))
    except Exception:
        LOGGER.exception("扫描件增强出现未预期错误 task=%s", task_id)
        _fail_task(task_id, "扫描件增强失败", "扫描件增强失败，请换一张图片重试")
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
        "excel": converter.excel_available(),
        "powerpoint": converter.powerpoint_available(),
        "python": platform.python_version(),
        "version": VERSION,
    }


@app.head("/api/health")
def health_head():
    return health()


@app.get("/api/config")
def config():
    return {
        "max_file_size_mb": MAX_FILE_SIZE_MB,
        "max_batch_files": MAX_BATCH_FILES,
        "max_batch_size_mb": MAX_BATCH_SIZE_MB,
        "translate_configured": translator.is_configured(),
        "glossary": translator.read_glossary_file(),
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


@app.post("/api/batch")
async def batch_api(
    files: list[UploadFile] = File(...),
    mode: str = Form("convert"),
    target: str = Form(""),
):
    mode = (mode or "convert").strip().lower()
    if mode != "convert":
        for upload in files:
            await upload.close()
        raise ApiError(400, "当前阶段仅支持批量格式转换，批量压缩将在后续开放")
    if not files:
        raise ApiError(400, "请至少选择一个文件")
    if len(files) > MAX_BATCH_FILES:
        for upload in files:
            await upload.close()
        raise ApiError(413, f"批量最多支持 {MAX_BATCH_FILES} 个文件，当前 {len(files)} 个")

    requested = target.strip().lower().lstrip(".")
    if not requested:
        for upload in files:
            await upload.close()
        raise ApiError(400, "请选择批量转换的目标格式")

    all_formats = converter.get_formats()
    allowed = set(all_formats)
    saved: list[dict[str, Any]] = []
    skipped: list[dict[str, Any]] = []
    total_bytes = 0
    try:
        for index, upload in enumerate(files):
            display = _sanitize_filename(upload.filename)
            ext = _extension(display)
            declared = int(upload.size or 0)
            reason = ""
            if ext not in all_formats:
                reason = f"不支持 .{ext or '未知'} 文件"
            elif requested not in all_formats[ext]:
                reason = f"该格式不支持转换为 {requested.upper()}"
            elif declared > MAX_FILE_BYTES:
                reason = f"文件超过 {MAX_FILE_SIZE_MB}MB 上限"
            elif declared and total_bytes + declared > MAX_BATCH_BYTES:
                raise ApiError(413, f"批量文件总量超过 {MAX_BATCH_SIZE_MB}MB 上限")

            if reason:
                skipped.append({"index": index, "filename": display, "error": reason})
                await upload.close()
                continue

            remaining = MAX_BATCH_BYTES - total_bytes
            limit = min(MAX_FILE_BYTES, remaining)
            try:
                src, display, size = await _save_upload(
                    upload,
                    allowed,
                    max_bytes=limit,
                    limit_message=(
                        f"批量文件总量超过 {MAX_BATCH_SIZE_MB}MB 上限"
                        if limit < MAX_FILE_BYTES
                        else f"文件超过 {MAX_FILE_SIZE_MB}MB 上限"
                    ),
                )
            except ApiError as exc:
                if limit < MAX_FILE_BYTES and exc.status_code == 413:
                    raise
                skipped.append({"index": index, "filename": display, "error": exc.message})
                continue
            total_bytes += size
            saved.append({"index": index, "src": src, "display": display, "size": size})
    except ApiError:
        for item in saved:
            item["src"].unlink(missing_ok=True)
        for upload in files:
            await upload.close()
        raise
    except Exception:
        for item in saved:
            item["src"].unlink(missing_ok=True)
        LOGGER.exception("批量上传处理失败")
        for upload in files:
            await upload.close()
        raise ApiError(500, "批量上传处理失败，请重试") from None

    task_id = uuid.uuid4().hex
    _new_task(
        task_id,
        f"批量 {len(files)} 个文件",
        total_bytes,
        total_files=len(files),
        target=requested,
    )
    EXECUTOR.submit(_run_batch, task_id, saved, len(files), requested, skipped)
    return {"task_id": task_id}


@app.post("/api/compress")
async def compress_api(file: UploadFile = File(...), level: str = Form("80")):
    display = _sanitize_filename(file.filename)
    ext = _extension(display)
    if ext != "pdf" and ext not in converter.IMAGE_EXTS:
        await file.close()
        raise ApiError(400, converter.COMPRESS_UNSUPPORTED)
    level = (level or "80").strip()
    if level not in converter.COMPRESS_LEVELS:
        await file.close()
        raise ApiError(400, "压缩档位无效，只支持 80 / 60 / 40 / 20")

    src, display, size = await _save_upload(file, {"pdf", *converter.IMAGE_EXTS})
    if ext == "pdf":
        _check_pdf(src, display)

    task_id = uuid.uuid4().hex
    _new_task(task_id, display, size, level=level)
    _start_task_thread(_run_compress, task_id, src, display, level)
    return {
        "task_id": task_id,
        "status": "pending",
        "filename": display,
        "size": size,
        "download_url": f"/api/download/{task_id}",
        "note": "",
    }


@app.post("/api/compress/estimate")
async def compress_estimate_api(file: UploadFile = File(...)):
    display = _sanitize_filename(file.filename)
    ext = _extension(display)
    if ext != "pdf" and ext not in converter.IMAGE_EXTS:
        await file.close()
        raise ApiError(400, converter.COMPRESS_UNSUPPORTED)

    src, display, _ = await _save_upload(file, {"pdf", *converter.IMAGE_EXTS})
    try:
        if ext == "pdf":
            _check_pdf(src, display)
        return converter.estimate_compression(src)
    except converter.ConversionError as exc:
        raise ApiError(400, str(exc)) from None
    finally:
        src.unlink(missing_ok=True)


@app.post("/api/pdf/merge")
async def pdf_merge_api(
    files: list[UploadFile] = File(...),
    mode: str = Form("append"),
    insert_index: int = Form(0),
):
    mode = (mode or "append").strip().lower()
    if len(files) < 2:
        for upload in files:
            await upload.close()
        raise ApiError(400, "合并至少需要 2 个 PDF 文件")
    if mode not in {"append", "interleave", "insert"}:
        for upload in files:
            await upload.close()
        raise ApiError(400, "合并方式无效，只支持 append / interleave / insert")

    saved: list[tuple[Path, str]] = []
    try:
        for upload in files:
            display = _sanitize_filename(upload.filename)
            if _extension(display) != "pdf":
                await upload.close()
                raise ApiError(400, f"文件「{display}」不是 PDF")
            src, display, _ = await _save_upload(upload, {"pdf"})
            _check_pdf(src, display)
            saved.append((src, display))

        if mode == "insert":
            base_count = pdf_tools.validate_pdf(saved[0][0])
            if insert_index < 0 or insert_index > base_count:
                raise ApiError(
                    400,
                    f"插入位置应在 0 到 {base_count} 之间（0 表示插到第 1 页之前）",
                )
    except ApiError:
        _discard_uploads([path for path, _ in saved])
        for upload in files:
            await upload.close()
        raise
    except Exception:
        _discard_uploads([path for path, _ in saved])
        for upload in files:
            await upload.close()
        LOGGER.exception("PDF 合并上传处理失败")
        raise ApiError(500, "合并文件上传处理失败，请重试") from None

    total_size = sum(path.stat().st_size for path, _ in saved)
    task_id = uuid.uuid4().hex
    _new_task(task_id, f"合并 {len(saved)} 个 PDF", total_size, mode=mode)
    _start_task_thread(_run_pdf_merge, task_id, saved, mode, insert_index)
    return {
        "task_id": task_id,
        "status": "pending",
        "filename": f"合并 {len(saved)} 个 PDF",
        "size": total_size,
        "download_url": f"/api/download/{task_id}",
        "note": "",
    }


@app.post("/api/pdf/ops")
async def pdf_ops_api(
    file: UploadFile = File(...),
    delete_pages: str = Form(""),
    rotations: str = Form(""),
    page_order: str = Form(""),
):
    deletes = _json_list_arg(delete_pages, "delete_pages")
    rotation_map = _json_dict_arg(rotations, "rotations")
    order = _json_list_arg(page_order, "page_order")
    if not deletes and not rotation_map and not order:
        await file.close()
        raise ApiError(400, "请至少指定删除页、页面顺序或旋转角度")

    src, display, size = await _save_upload(file, {"pdf"})
    _check_pdf(src, display)

    task_id = uuid.uuid4().hex
    _new_task(task_id, display, size)
    _start_task_thread(_run_pdf_ops, task_id, src, display, deletes, rotation_map, order)
    return {
        "task_id": task_id,
        "status": "pending",
        "filename": display,
        "size": size,
        "download_url": f"/api/download/{task_id}",
        "note": "",
    }


@app.post("/api/pdf/preview")
async def pdf_preview_api(file: UploadFile = File(...), limit: int = Form(50)):
    src, display, _ = await _save_upload(file, {"pdf"})
    try:
        _check_pdf(src, display)
        return pdf_tools.preview(src, limit)
    except converter.ConversionError as exc:
        raise ApiError(400, str(exc)) from None
    finally:
        src.unlink(missing_ok=True)


@app.post("/api/pdf/to-jpg")
async def pdf_to_jpg_api(
    files: list[UploadFile] = File(...),
    max_kb: int = Form(300),
):
    if not files:
        raise ApiError(400, "请至少选择一个 PDF 文件")
    if max_kb < 10 or max_kb > 10240:
        for upload in files:
            await upload.close()
        raise ApiError(400, "每张大小上限应在 10KB 到 10240KB 之间")

    saved: list[tuple[Path, str]] = []
    try:
        for upload in files:
            display = _sanitize_filename(upload.filename)
            if _extension(display) != "pdf":
                await upload.close()
                raise ApiError(400, f"文件「{display}」不是 PDF")
            src, display, _ = await _save_upload(upload, {"pdf"})
            _check_pdf(src, display)
            saved.append((src, display))
    except ApiError:
        _discard_uploads([path for path, _ in saved])
        for upload in files:
            await upload.close()
        raise
    except Exception:
        _discard_uploads([path for path, _ in saved])
        for upload in files:
            await upload.close()
        LOGGER.exception("JPG 导出上传处理失败")
        raise ApiError(500, "JPG 导出文件上传处理失败，请重试") from None

    total_size = sum(path.stat().st_size for path, _ in saved)
    task_id = uuid.uuid4().hex
    _new_task(task_id, f"JPG 导出 {len(saved)} 个 PDF", total_size, max_kb=max_kb)
    _start_task_thread(_run_pdf_to_jpg, task_id, saved, max_kb)
    return {
        "task_id": task_id,
        "status": "pending",
        "filename": f"JPG 导出 {len(saved)} 个 PDF",
        "size": total_size,
        "download_url": f"/api/download/{task_id}",
        "note": "",
    }


@app.post("/api/translate")
async def translate_api(
    file: UploadFile = File(...),
    target_lang: str = Form("zh"),
    mode: str = Form("bilingual"),
    glossary: str = Form(""),
):
    display = _sanitize_filename(file.filename)
    ext = _extension(display)
    mode = (mode or "bilingual").strip().lower()
    target_lang = (target_lang or "zh").strip().lower()

    if not translator.is_configured():
        await file.close()
        raise ApiError(400, "未配置翻译密钥，请在 .env 里填写 TRANSLATE_API_KEY")
    if ext not in {"pdf", "docx", "txt"}:
        await file.close()
        raise ApiError(415, f"文献翻译只支持 PDF、DOCX、TXT，不支持 .{ext or '未知'} 文件")
    if target_lang not in {"zh", "en"}:
        await file.close()
        raise ApiError(400, "目标语言无效，只支持 zh / en")
    if mode not in {"bilingual", "translated", "layout"}:
        await file.close()
        raise ApiError(400, "输出方式无效，只支持 bilingual / translated / layout")
    if mode == "layout" and ext != "pdf":
        await file.close()
        raise ApiError(400, "保留版式输出只支持 PDF，请先用格式转换转成 PDF")

    src, display, size = await _save_upload(file, {"pdf", "docx", "txt"})
    if ext == "pdf":
        try:
            _check_pdf(src, display)
        except ApiError:
            src.unlink(missing_ok=True)
            raise

    task_id = uuid.uuid4().hex
    _new_task(task_id, display, size, mode=mode, target_lang=target_lang)
    _start_task_thread(_run_translate, task_id, src, display, mode, target_lang, glossary)
    return {
        "task_id": task_id,
        "status": "pending",
        "filename": display,
        "size": size,
        "download_url": f"/api/download/{task_id}",
        "note": "",
    }


@app.post("/api/scan-photo")
async def scan_photo_api(file: UploadFile = File(...)):
    src, display, size = await _save_upload(file, set(converter.IMAGE_EXTS))
    task_id = uuid.uuid4().hex
    _new_task(task_id, display, size)
    _start_task_thread(_run_scan, task_id, src, display)
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
def download(task_id: str, file: int | None = None):
    with TASK_LOCK:
        task = TASKS.get(task_id)
        if task is None:
            raise ApiError(404, "任务不存在或结果已过期")
        if task["status"] == "error":
            raise ApiError(409, task["error"] or "任务转换失败")
        if task["status"] != "done":
            raise ApiError(409, "任务尚未完成，请稍后再试")
        if file is not None:
            entry = (task.get("_batch_files") or {}).get(file)
            if entry is None:
                raise ApiError(404, "指定的文件不存在或转换失败")
            path = Path(entry[0])
            filename = entry[1]
        else:
            path = Path(task["_path"]) if task["_path"] else None
            filename = task["_download_name"] or (path.name if path else "download")
    if path is None:
        raise ApiError(404, "没有可下载的结果文件")
    if not path.exists():
        raise ApiError(404, "结果文件已过期或被清理")

    media_type = "application/zip" if path.suffix.lower() == ".zip" else (
        mimetypes.guess_type(filename)[0] or "application/octet-stream"
    )
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
                LOGGER.debug("重设标准输出编码失败", exc_info=True)
    port = _env_int("PORT", 8000)
    if os.getenv("AUTO_OPEN_BROWSER", "0") == "1":
        threading.Thread(target=_open_browser_later, args=(port,), daemon=True).start()
    uvicorn.run(app, host="127.0.0.1", port=port, log_level="info")