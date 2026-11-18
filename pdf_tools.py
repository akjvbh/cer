"""PDF 工具：合并、页面操作、缩略图预览、逐页导出 JPG。"""

from __future__ import annotations

import base64
import logging
import math
import zipfile
from pathlib import Path
from typing import Callable

from converter import ConversionError

LOGGER = logging.getLogger(__name__)

PREVIEW_MAX_EDGE = 260
PREVIEW_LIMIT = 50
JPG_DPI_LADDER = (150, 120, 96, 72)
JPG_QUALITY_LADDER = (95, 85, 75, 65)
JPG_FALLBACK_DPI = 72
JPG_FALLBACK_QUALITY = 50
VALID_ROTATIONS = {0, 90, 180, 270}

ProgressCallback = Callable[[int, str], None]


def _fitz():
    try:
        import fitz
    except ImportError as exc:
        raise ConversionError("缺少 PyMuPDF，请重新安装依赖") from exc
    return fitz


def _open_pdf(src: Path):
    fitz = _fitz()
    try:
        doc = fitz.open(str(src))
    except Exception as exc:
        raise ConversionError("无法打开 PDF，文件可能已损坏") from exc
    if getattr(doc, "needs_pass", False):
        doc.close()
        raise ConversionError("不支持加密 PDF，请先解除密码")
    if doc.page_count == 0:
        doc.close()
        raise ConversionError("PDF 中没有可处理的页面")
    return doc


def validate_pdf(src: Path) -> int:
    """打开一次做加密与结构校验，返回页数。"""
    doc = _open_pdf(src)
    try:
        return doc.page_count
    finally:
        doc.close()


def _page_error(index: int, total: int) -> ConversionError:
    return ConversionError(f"第 {index} 页不存在，该文件共 {total} 页")


def _check_pages(values: list[int], total: int, label: str) -> list[int]:
    checked: list[int] = []
    for value in values:
        if not isinstance(value, int) or isinstance(value, bool):
            raise ConversionError(f"{label}中的页码必须是整数")
        if value < 1 or value > total:
            if label == "删除页":
                raise _page_error(value, total)
            raise ConversionError(f"{label}中的第 {value} 页不存在，该文件共 {total} 页")
        checked.append(value)
    return checked


def merge(
    files: list[tuple[Path, str]],
    dst_dir: Path,
    mode: str,
    insert_index: int = 0,
    progress: ProgressCallback | None = None,
) -> tuple[Path, str]:
    """合并多个 PDF。files 为 (路径, 展示名) 列表。"""
    if len(files) < 2:
        raise ConversionError("合并至少需要 2 个 PDF 文件")
    mode = (mode or "append").strip().lower()
    if mode not in {"append", "interleave", "insert"}:
        raise ConversionError("合并方式无效，只支持 append / interleave / insert")

    fitz = _fitz()
    docs = []
    try:
        for path, display in files:
            try:
                docs.append(_open_pdf(path))
            except ConversionError as exc:
                raise ConversionError(f"文件「{display}」：{exc}") from exc

        counts = [doc.page_count for doc in docs]
        total_pages = sum(counts)
        if mode == "insert":
            base_count = counts[0]
            if insert_index < 0 or insert_index > base_count:
                raise ConversionError(
                    f"插入位置应在 0 到 {base_count} 之间（0 表示插到第 1 页之前）"
                )

        merged = fitz.open()
        written = 0

        def report(message: str) -> None:
            if progress is None:
                return
            percent = 5 + int(written / max(1, total_pages) * 90)
            progress(min(95, percent), message)

        try:
            if mode == "append":
                for doc, (_, display) in zip(docs, files):
                    merged.insert_pdf(doc)
                    written += doc.page_count
                    report(f"已合并「{display}」，共 {written}/{total_pages} 页")
            elif mode == "interleave":
                for index in range(max(counts)):
                    for doc, count in zip(docs, counts):
                        if index >= count:
                            continue
                        merged.insert_pdf(doc, from_page=index, to_page=index)
                        written += 1
                    report(f"逐页交替合并，共 {written}/{total_pages} 页")
            else:
                cursor = insert_index
                merged.insert_pdf(docs[0])
                written = counts[0]
                report(f"已放入主文件，共 {written}/{total_pages} 页")
                for doc in docs[1:]:
                    merged.insert_pdf(doc, start_at=cursor)
                    cursor += doc.page_count
                    written += doc.page_count
                    report(f"已插入附加文件，共 {written}/{total_pages} 页")

            if merged.page_count != total_pages:
                raise ConversionError("合并后的页数与原文件之和不一致，请检查源文件")

            dst_dir.mkdir(parents=True, exist_ok=True)
            out = dst_dir / "merged.pdf"
            merged.save(str(out), garbage=3, deflate=True)
        finally:
            merged.close()
    finally:
        for doc in docs:
            doc.close()

    if progress:
        progress(98, "正在整理合并结果")
    return out, f"共合并 {len(files)} 个文件、{total_pages} 页"


def page_ops(
    src: Path,
    dst_dir: Path,
    delete_pages: list[int],
    rotations: dict[int, int],
    page_order: list[int],
    progress: ProgressCallback | None = None,
) -> tuple[Path, str]:
    """按白名单/删除页/最终旋转角度生成新 PDF。"""
    fitz = _fitz()
    doc = _open_pdf(src)
    total = doc.page_count
    try:
        deletes = _check_pages(delete_pages or [], total, "删除页")
        order = _check_pages(page_order or [], total, "页面顺序")
        rotation_map: dict[int, int] = {}
        for key, value in (rotations or {}).items():
            try:
                page_number = int(key)
                angle = int(value)
            except (TypeError, ValueError) as exc:
                raise ConversionError("旋转角度必须是 0 / 90 / 180 / 270") from exc
            if page_number < 1 or page_number > total:
                raise _page_error(page_number, total)
            if angle not in VALID_ROTATIONS:
                raise ConversionError("旋转角度只允许 0 / 90 / 180 / 270")
            rotation_map[page_number] = angle

        if not deletes and not order and not rotation_map:
            raise ConversionError("请至少指定删除页、页面顺序或旋转角度")

        delete_set = set(deletes)
        if order:
            final_order = order
        elif deletes:
            final_order = [index for index in range(1, total + 1) if index not in delete_set]
        else:
            final_order = list(range(1, total + 1))

        if not final_order:
            raise ConversionError("所有页面都被删除，结果不能为空")

        out_doc = fitz.open()
        try:
            for position, page_number in enumerate(final_order, start=1):
                out_doc.insert_pdf(doc, from_page=page_number - 1, to_page=page_number - 1)
                if page_number in rotation_map:
                    out_doc[out_doc.page_count - 1].set_rotation(rotation_map[page_number])
                if progress:
                    percent = 5 + int(position / len(final_order) * 90)
                    progress(min(95, percent), f"正在处理第 {position}/{len(final_order)} 页")

            if out_doc.page_count != len(final_order):
                raise ConversionError("页面操作结果页数异常，请重试")

            dst_dir.mkdir(parents=True, exist_ok=True)
            out = dst_dir / "pages.pdf"
            out_doc.save(str(out), garbage=3, deflate=True)
        finally:
            out_doc.close()
    finally:
        doc.close()

    parts = []
    if deletes:
        parts.append(f"删除 {len(deletes)} 页")
    if order:
        parts.append(f"按新顺序保留 {len(order)} 页")
    if rotation_map:
        parts.append(f"旋转 {len(rotation_map)} 页")
    note = "；".join(parts) if parts else "页面未变化"
    return out, f"已处理：{note}，结果共 {len(final_order)} 页"


def preview(src: Path, limit: int = PREVIEW_LIMIT) -> dict:
    """返回前 limit 页的缩略图，长边 260px；rotation 为原始角度。"""
    fitz = _fitz()
    doc = _open_pdf(src)
    try:
        total = doc.page_count
        limit = max(1, min(int(limit or PREVIEW_LIMIT), PREVIEW_LIMIT))
        shown = min(limit, total)
        pages: list[dict] = []
        for index in range(shown):
            page = doc[index]
            original_rotation = int(page.rotation or 0)
            if original_rotation:
                page.set_rotation(0)
            try:
                rect = page.rect
                longest = max(rect.width, rect.height, 1)
                scale = PREVIEW_MAX_EDGE / longest
                pix = page.get_pixmap(
                    matrix=fitz.Matrix(scale, scale),
                    colorspace=fitz.csRGB,
                    alpha=False,
                )
                png = pix.tobytes("png")
            finally:
                if original_rotation:
                    page.set_rotation(original_rotation)
            pages.append(
                {
                    "index": index + 1,
                    "width": round(rect.width),
                    "height": round(rect.height),
                    "rotation": original_rotation,
                    "png_base64": base64.b64encode(png).decode("ascii"),
                }
            )
        return {"pages": pages, "truncated": total > shown, "total": total}
    finally:
        doc.close()


def _render_jpeg(page, dpi: int, quality: int) -> bytes:
    fitz = _fitz()
    zoom = dpi / 72
    pix = page.get_pixmap(
        matrix=fitz.Matrix(zoom, zoom),
        colorspace=fitz.csRGB,
        alpha=False,
    )
    return pix.tobytes("jpeg", jpg_quality=quality)


def _render_under_limit(page, max_bytes: int, page_number: int) -> bytes:
    last_size = 0
    for dpi in JPG_DPI_LADDER:
        for quality in JPG_QUALITY_LADDER:
            data = _render_jpeg(page, dpi, quality)
            last_size = len(data)
            if last_size <= max_bytes:
                return data
    data = _render_jpeg(page, JPG_FALLBACK_DPI, JPG_FALLBACK_QUALITY)
    last_size = len(data)
    if last_size <= max_bytes:
        return data
    current_kb = math.ceil(last_size / 1024)
    limit_kb = max(1, round(max_bytes / 1024))
    raise ConversionError(f"第 {page_number} 页无法压到 {limit_kb}KB 以内（当前 {current_kb}KB）")


def _unique_dir(name: str, used: set[str]) -> str:
    if name not in used:
        used.add(name)
        return name
    counter = 2
    while f"{name} ({counter})" in used:
        counter += 1
    final = f"{name} ({counter})"
    used.add(final)
    return final


def export_jpg(
    files: list[tuple[Path, str]],
    dst_dir: Path,
    max_kb: int = 300,
    progress: ProgressCallback | None = None,
) -> tuple[Path, str, list[dict[str, str]]]:
    """逐页导出 JPG 并打包 ZIP，每张不超过 max_kb。

    返回 (zip, note, failures)。failures 中每个元素为
    {"filename": 展示名, "error": 中文原因}。
    """
    if not files:
        raise ConversionError("请至少选择一个 PDF 文件")
    max_kb = max(10, min(int(max_kb or 300), 10240))
    max_bytes = max_kb * 1024

    docs: list[tuple] = []
    try:
        for path, display in files:
            try:
                docs.append((path, display, _open_pdf(path)))
            except ConversionError as exc:
                raise ConversionError(f"文件「{display}」：{exc}") from exc

        total_pages = sum(doc.page_count for _, _, doc in docs)
        done_pages = 0
        root = dst_dir / "jpg"
        root.mkdir(parents=True, exist_ok=True)
        used_dirs: set[str] = set()
        failures: list[dict[str, str]] = []
        written_files = 0

        for path, display, doc in docs:
            folder = _unique_dir(_safe_folder(display), used_dirs)
            target_dir = root / folder
            target_dir.mkdir(parents=True, exist_ok=True)
            file_failed = False
            for index in range(doc.page_count):
                page_number = index + 1
                try:
                    data = _render_under_limit(doc[index], max_bytes, page_number)
                except ConversionError as exc:
                    LOGGER.warning("JPG 导出失败 file=%s: %s", display, exc)
                    failures.append({"filename": display, "error": str(exc)})
                    file_failed = True
                    done_pages += doc.page_count - index
                    break
                (target_dir / f"{page_number:03d}.jpg").write_bytes(data)
                written_files += 1
                done_pages += 1
                if progress:
                    percent = 5 + int(done_pages / max(1, total_pages) * 88)
                    progress(
                        min(93, percent),
                        f"正在导出「{display}」第 {page_number}/{doc.page_count} 页",
                    )
            if file_failed:
                for leftover in target_dir.glob("*"):
                    leftover.unlink(missing_ok=True)
                target_dir.rmdir()

        if written_files == 0:
            first_error = failures[0]["error"] if failures else "没有可导出的页面"
            raise ConversionError(f"所有文件都未能导出：{first_error}")

        zip_path = dst_dir / "pdf_jpg.zip"
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_STORED) as archive:
            for image_path in sorted(root.rglob("*.jpg")):
                arcname = image_path.relative_to(root).as_posix()
                archive.write(image_path, arcname)
        note = f"共导出 {written_files} 张 JPG，每张不超过 {max_kb}KB"
        if failures:
            note = f"{note}；{len(failures)} 个文件失败"
        return zip_path, note, failures
    finally:
        for _, _, doc in docs:
            doc.close()


def _safe_folder(display: str) -> str:
    stem = Path(display or "PDF").stem
    cleaned = "".join("_" if char in '<>:"/\\|?*' else char for char in stem).strip().strip(".")
    return (cleaned or "PDF")[:80]
