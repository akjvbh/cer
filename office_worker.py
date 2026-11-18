"""独立进程执行 Office COM 转换，主进程负责加锁、超时和清理。"""

from __future__ import annotations

import gc
import json
import logging
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

LOGGER = logging.getLogger(__name__)


def _configure_stdio() -> None:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                LOGGER.debug("忽略 Office 清理或可选设置失败", exc_info=True)


def _word(src: Path, dst: Path) -> None:
    import win32com.client

    word = win32com.client.DispatchEx("Word.Application")
    doc = None
    try:
        word.Visible = False
        word.DisplayAlerts = 0
        try:
            word.AutomationSecurity = 3
        except Exception:
            LOGGER.debug("忽略 Office 清理或可选设置失败", exc_info=True)
        doc = word.Documents.Open(
            FileName=str(src),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
            Visible=False,
        )
        if dst.suffix.lower() == ".docx":
            doc.SaveAs2(FileName=str(dst), FileFormat=16)
        else:
            try:
                doc.SaveAs2(FileName=str(dst), FileFormat=17)
            except Exception:
                doc.ExportAsFixedFormat(
                    OutputFileName=str(dst),
                    ExportFormat=17,
                    OpenAfterExport=False,
                    OptimizeFor=0,
                    Range=0,
                    Item=0,
                    IncludeDocProps=True,
                    KeepIRM=True,
                    CreateBookmarks=0,
                    DocStructureTags=True,
                    BitmapMissingFonts=True,
                    UseISO19005_1=False,
                )
    finally:
        if doc is not None:
            try:
                doc.Close(SaveChanges=0)
            except Exception:
                LOGGER.debug("忽略 Office 清理或可选设置失败", exc_info=True)
        try:
            word.Quit()
        except Exception:
            LOGGER.debug("忽略 Office 清理或可选设置失败", exc_info=True)
        del doc
        del word
        gc.collect()


def _excel(src: Path, dst: Path) -> None:
    import win32com.client

    app = win32com.client.DispatchEx("Excel.Application")
    workbook = None
    try:
        app.Visible = False
        app.DisplayAlerts = False
        workbook = app.Workbooks.Open(str(src), ReadOnly=True, UpdateLinks=0)
        workbook.ExportAsFixedFormat(0, str(dst))
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                LOGGER.debug("忽略 Office 清理或可选设置失败", exc_info=True)
        try:
            app.Quit()
        except Exception:
            LOGGER.debug("忽略 Office 清理或可选设置失败", exc_info=True)
        del workbook
        del app
        gc.collect()


def _powerpoint(src: Path, dst: Path) -> None:
    import win32com.client

    app = win32com.client.DispatchEx("PowerPoint.Application")
    presentation = None
    try:
        try:
            app.DisplayAlerts = 1
        except Exception:
            LOGGER.debug("忽略 Office 清理或可选设置失败", exc_info=True)
        app.Visible = True
        presentation = app.Presentations.Open(str(src), ReadOnly=True, WithWindow=False)
        try:
            presentation.SaveAs(str(dst), 32)
        except Exception:
            presentation.ExportAsFixedFormat(str(dst), 2)
    finally:
        if presentation is not None:
            try:
                presentation.Close()
            except Exception:
                LOGGER.debug("忽略 Office 清理或可选设置失败", exc_info=True)
        try:
            app.Quit()
        except Exception:
            LOGGER.debug("忽略 Office 清理或可选设置失败", exc_info=True)
        del presentation
        del app
        gc.collect()


def _powerpoint_png(src: Path, dst: Path) -> None:
    import win32com.client

    app = win32com.client.DispatchEx("PowerPoint.Application")
    presentation = None
    folder = Path(tempfile.mkdtemp(prefix="ppt-png-", dir=str(dst.parent)))
    try:
        try:
            app.DisplayAlerts = 1
        except Exception:
            LOGGER.debug("忽略 Office 清理或可选设置失败", exc_info=True)
        app.Visible = True
        presentation = app.Presentations.Open(str(src), ReadOnly=True, WithWindow=False)
        presentation.Export(str(folder), "PNG", 1600, 1200)
        images = sorted(
            (path for path in folder.iterdir() if path.suffix.lower() in {".png", ".jpg"}),
            key=lambda path: path.name.lower(),
        )
        if not images:
            raise RuntimeError("PowerPoint 没有导出任何幻灯片图片")
        with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as archive:
            for index, image in enumerate(images, start=1):
                suffix = ".png" if image.suffix.lower() == ".png" else image.suffix.lower()
                archive.write(image, f"{index:03d}{suffix}")
    finally:
        if presentation is not None:
            try:
                presentation.Close()
            except Exception:
                LOGGER.debug("忽略 Office 清理或可选设置失败", exc_info=True)
        try:
            app.Quit()
        except Exception:
            LOGGER.debug("忽略 Office 清理或可选设置失败", exc_info=True)
        del presentation
        del app
        gc.collect()
        shutil.rmtree(folder, ignore_errors=True)


def main() -> int:
    _configure_stdio()
    if len(sys.argv) != 4:
        print(json.dumps({"ok": False, "error": "参数应为：kind src dst"}, ensure_ascii=False))
        return 2
    kind, src, dst = sys.argv[1].lower(), Path(sys.argv[2]), Path(sys.argv[3])
    handlers = {
        "word": _word,
        "excel": _excel,
        "powerpoint": _powerpoint,
        "powerpoint_png": _powerpoint_png,
    }
    handler = handlers.get(kind)
    if handler is None:
        print(json.dumps({"ok": False, "error": f"不支持的 Office 类型：{kind}"}, ensure_ascii=False))
        return 2

    import pythoncom

    pythoncom.CoInitialize()
    try:
        handler(src, dst)
    except Exception as exc:
        print(
            json.dumps({"ok": False, "error": f"{type(exc).__name__}: {exc}"}, ensure_ascii=False),
            file=sys.stderr,
        )
        return 1
    finally:
        pythoncom.CoUninitialize()

    if not dst.exists() or dst.stat().st_size == 0:
        print(json.dumps({"ok": False, "error": "转换结果为空"}, ensure_ascii=False), file=sys.stderr)
        return 1
    print(json.dumps({"ok": True, "path": str(dst)}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())