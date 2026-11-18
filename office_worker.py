"""独立进程执行 Office COM 转换，主进程负责加锁、超时和清理。"""

from __future__ import annotations

import gc
import json
import sys
from pathlib import Path


def _configure_stdio() -> None:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass


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
            pass
        doc = word.Documents.Open(
            FileName=str(src),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
            Visible=False,
        )
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
                pass
        try:
            word.Quit()
        except Exception:
            pass
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
                pass
        try:
            app.Quit()
        except Exception:
            pass
        del workbook
        del app
        gc.collect()


def _powerpoint(src: Path, dst: Path) -> None:
    import win32com.client

    app = win32com.client.DispatchEx("PowerPoint.Application")
    presentation = None
    try:
        app.Visible = True
        presentation = app.Presentations.Open(str(src), ReadOnly=True, WithWindow=False)
        presentation.SaveAs(str(dst), 32)
    finally:
        if presentation is not None:
            try:
                presentation.Close()
            except Exception:
                pass
        try:
            app.Quit()
        except Exception:
            pass
        del presentation
        del app
        gc.collect()


def main() -> int:
    _configure_stdio()
    if len(sys.argv) != 4:
        print(json.dumps({"ok": False, "error": "参数应为：kind src dst"}, ensure_ascii=False))
        return 2
    kind, src, dst = sys.argv[1].lower(), Path(sys.argv[2]), Path(sys.argv[3])
    handlers = {"word": _word, "excel": _excel, "powerpoint": _powerpoint}
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