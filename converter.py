"""转换器注册表：所有「源格式 → 目标格式」的实现都放在这里。"""

from __future__ import annotations

import base64
import csv
import datetime as dt
import html
import importlib.util
import io
import json
import logging
import os
import re
import subprocess
import sys
import tempfile
import threading
import time
import zipfile
from pathlib import Path
from typing import Callable

LOGGER = logging.getLogger(__name__)
BASE_DIR = Path(__file__).resolve().parent
IMAGE_EXTS = ("png", "jpg", "jpeg", "webp", "bmp", "gif", "tiff")
IMAGE_TARGETS = ("png", "jpg", "webp", "bmp", "tiff")
OFFICE_LOCK = threading.RLock()
_APP_CACHE: dict[str, tuple[float, bool]] = {}
_APP_CACHE_LOCK = threading.Lock()
MAX_PDF_ROWS = 3000
MAX_PDF_COLS = 30
OFFICE_EXES = {
    "word": "WINWORD.EXE",
    "excel": "EXCEL.EXE",
    "powerpoint": "POWERPNT.EXE",
    "powerpoint_png": "POWERPNT.EXE",
}


class ConversionError(Exception):
    """可直接展示给用户的中文转换错误。"""


ProgressCallback = Callable[[int, str], None]


# ---------------------------------------------------------------------------
# 注册表
# ---------------------------------------------------------------------------

HANDLERS: dict[tuple[str, str], Callable[[Path, Path, str], tuple[Path, str]]] = {}


def register(src: str, target: str):
    def decorator(func: Callable[[Path, Path, str], tuple[Path, str]]):
        HANDLERS[(src.lower(), target.lower())] = func
        return func

    return decorator


def get_formats() -> dict[str, list[str]]:
    """由注册表反推格式表，避免维护第二份映射。"""
    result: dict[str, set[str]] = {}
    for (src, target) in HANDLERS:
        result.setdefault(src, set()).add(target)
    return {src: sorted(targets) for src, targets in sorted(result.items())}


def convert(
    src: Path,
    dst_dir: Path,
    target: str,
    progress: ProgressCallback | None = None,
) -> tuple[Path, str]:
    src = Path(src)
    target = target.lower().lstrip(".")
    ext = src.suffix.lower().lstrip(".")
    handler = HANDLERS.get((ext, target))
    if handler is None:
        raise ConversionError(f"不支持从 {ext or '未知格式'} 转换到 {target}")

    dst_dir = Path(dst_dir)
    dst_dir.mkdir(parents=True, exist_ok=True)
    if progress:
        progress(8, "正在读取文件")
    output, note = handler(src, dst_dir, target)
    output = Path(output)
    if not output.exists() or output.stat().st_size == 0:
        raise ConversionError("转换没有产生有效文件，请检查源文件是否损坏")
    if progress:
        progress(92, "正在整理结果")
    return output, note


# ---------------------------------------------------------------------------
# 文本与 Office 辅助
# ---------------------------------------------------------------------------

def _read_text(path: Path) -> str:
    raw = path.read_bytes()
    for encoding in ("utf-8-sig", "utf-8", "gb18030"):
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def _text_output(src: Path, dst_dir: Path, content: str, note: str) -> tuple[Path, str]:
    out = dst_dir / f"{src.stem}.txt"
    out.write_text(content, encoding="utf-8-sig", newline="\n")
    return out, note


def _is_pdf_encrypted(doc) -> bool:
    return bool(getattr(doc, "needs_pass", False))


def _open_pdf(src: Path):
    try:
        import fitz
    except ImportError as exc:
        raise ConversionError("缺少 PyMuPDF，请重新安装依赖") from exc
    try:
        doc = fitz.open(str(src))
    except Exception as exc:
        raise ConversionError(f"无法打开 PDF：{exc}") from exc
    if _is_pdf_encrypted(doc):
        doc.close()
        raise ConversionError("不支持加密或需要密码的 PDF，请先解除密码")
    return doc


def _office_timeout() -> int:
    try:
        return max(10, int(os.getenv("OFFICE_TIMEOUT_SECONDS", "600")))
    except ValueError:
        return 600


def _office_enabled() -> bool:
    enabled = os.getenv("OFFICE_ENABLED", "1").lower() not in {"0", "false", "no"}
    if not enabled or sys.platform != "win32":
        return False
    return importlib.util.find_spec("win32com") is not None


def _probe_executable(exe: str) -> bool:
    """查 App Paths 注册表与常见安装目录，不启动 Office 进程。"""
    try:
        import winreg

        for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            try:
                with winreg.OpenKey(
                    root,
                    rf"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\{exe}",
                ) as key:
                    value, _ = winreg.QueryValueEx(key, "")
                    if value and Path(value).exists():
                        return True
            except OSError:
                continue
    except Exception:
        LOGGER.debug("读取 Office 注册表失败", exc_info=True)

    for env_name in ("ProgramFiles", "ProgramFiles(x86)"):
        base = os.environ.get(env_name)
        if not base:
            continue
        for version in ("Office16", "Office15", "Office14"):
            if (Path(base) / "Microsoft Office" / "root" / version / exe).exists():
                return True
    return False


def _app_available(exe: str, force: bool = False) -> bool:
    if not _office_enabled():
        return False
    now = time.monotonic()
    with _APP_CACHE_LOCK:
        cached = _APP_CACHE.get(exe)
        if not force and cached and now - cached[0] < 60:
            return cached[1]
    found = _probe_executable(exe)
    with _APP_CACHE_LOCK:
        _APP_CACHE[exe] = (now, found)
    return found


def office_available(force: bool = False) -> bool:
    """Word 是否可用；注册表/安装路径探测，不启动 Word。"""
    return _app_available("WINWORD.EXE", force)


def excel_available(force: bool = False) -> bool:
    return _app_available("EXCEL.EXE", force)


def powerpoint_available(force: bool = False) -> bool:
    return _app_available("POWERPNT.EXE", force)


def _list_office_pids(image: str) -> set[int]:
    try:
        completed = subprocess.run(
            ["tasklist", "/FI", f"IMAGENAME eq {image}", "/FO", "CSV", "/NH"],
            capture_output=True,
            text=True,
            encoding="mbcs",
            errors="replace",
            timeout=10,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except Exception:
        return set()
    pids: set[int] = set()
    for row in csv.reader(io.StringIO(completed.stdout)):
        if len(row) > 1 and row[1].strip().isdigit() and row[0].strip().lower() == image.lower():
            pids.add(int(row[1]))
    return pids


def _kill_office(image: str, pids: set[int] | None = None) -> None:
    if pids:
        for pid in pids:
            subprocess.run(
                ["taskkill", "/F", "/PID", str(pid)],
                capture_output=True,
                timeout=10,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
    else:
        subprocess.run(
            ["taskkill", "/F", "/IM", image],
            capture_output=True,
            timeout=10,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )


def _office_error_detail(stderr: str) -> str:
    lines = [line for line in (stderr or "").strip().splitlines() if line.strip()]
    if not lines:
        return ""
    try:
        payload = json.loads(lines[-1])
        return str(payload.get("error", ""))[:300]
    except (ValueError, TypeError):
        return lines[-1][:300]


def _office_convert(kind: str, src: Path, dst: Path) -> None:
    image = OFFICE_EXES.get(kind, "WINWORD.EXE")
    src = Path(src).resolve()
    dst = Path(dst).resolve()
    dst.parent.mkdir(parents=True, exist_ok=True)
    before = _list_office_pids(image)
    timeout = _office_timeout()
    command = [sys.executable, str(BASE_DIR / "office_worker.py"), kind, str(src), str(dst)]
    with OFFICE_LOCK:
        try:
            completed = subprocess.run(
                command,
                cwd=str(BASE_DIR),
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=timeout,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
        except subprocess.TimeoutExpired as exc:
            LOGGER.error("Office 转换超时（%s 秒），结束 %s", timeout, image)
            _kill_office(image)
            raise ConversionError(f"Office 转换超过 {timeout} 秒，已结束残留进程") from exc
        finally:
            new_pids = _list_office_pids(image) - before
            if new_pids:
                LOGGER.warning("清理残留 Office 进程：%s", sorted(new_pids))
                _kill_office(image, new_pids)

    if completed.returncode != 0:
        detail = _office_error_detail(completed.stderr)
        LOGGER.error("Office 转换失败：%s", detail or completed.stderr.strip())
        raise ConversionError(f"Office 转换失败：{detail or '未知原因'}")
    if not dst.exists() or dst.stat().st_size == 0:
        raise ConversionError("Office 没有生成有效文件")


HTML_CSS = """
body { font-family: "Microsoft YaHei", "PingFang SC", "Noto Sans SC", sans-serif;
       font-size: 10.5pt; line-height: 1.65; color: #1d2733; }
h1, h2, h3, h4 { color: #102a43; line-height: 1.35; margin: 14pt 0 7pt; }
h1 { font-size: 20pt; } h2 { font-size: 16pt; } h3 { font-size: 13pt; }
p { margin: 0 0 6pt; }
table { border-collapse: collapse; width: 100%; margin: 8pt 0; font-size: 9pt; }
th, td { border: 1px solid #9aa7b4; padding: 3pt 5pt; vertical-align: top; }
th { background: #eef2f7; font-weight: bold; }
pre { background: #f5f7fa; padding: 7pt; font-family: Consolas, monospace; font-size: 8.5pt; }
code { font-family: Consolas, monospace; font-size: 9pt; }
blockquote { margin: 8pt 0; padding-left: 10pt; border-left: 3px solid #9db4c8; color: #4a5b6c; }
img { max-width: 100%; }
ul, ol { margin: 0 0 6pt 18pt; padding: 0; }
"""


def _html_document(body: str, title: str, extra_css: str = "") -> str:
    return f"""<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{html.escape(title)}</title>
<style>
body{{max-width:860px;margin:40px auto;padding:0 24px;}}
{HTML_CSS}
{extra_css}
</style>
</head>
<body>
{body}
</body>
</html>
"""


def _html_to_pdf(body: str, out: Path, title: str) -> int:
    """用 PyMuPDF 的 HTML 布局引擎分页生成 PDF。"""
    try:
        import fitz
    except ImportError as exc:
        raise ConversionError("缺少 PyMuPDF，请重新安装依赖") from exc

    page_rect = fitz.paper_rect("a4")
    where = page_rect + (54, 54, -54, -54)
    story = fitz.Story(html=f"<html><body>{body}</body></html>", user_css=HTML_CSS)
    writer = fitz.DocumentWriter(str(out))
    pages = 0
    try:
        while pages < 200:
            device = writer.begin_page(page_rect)
            more, _ = story.place(where)
            story.draw(device)
            writer.end_page()
            pages += 1
            if not more:
                break
    except Exception as exc:
        LOGGER.exception("HTML 排版失败")
        raise ConversionError(f"内置排版失败：{exc}") from exc
    finally:
        writer.close()

    if pages >= 200:
        out.unlink(missing_ok=True)
        raise ConversionError("内容超过 200 页，已停止生成，请拆分文件后重试")
    if not out.exists() or out.stat().st_size == 0:
        raise ConversionError("内置排版没有生成有效 PDF")
    return pages


# ---------------------------------------------------------------------------
# docx
# ---------------------------------------------------------------------------

@register("docx", "txt")
def docx_to_txt(src: Path, dst_dir: Path, target: str) -> tuple[Path, str]:
    try:
        from docx import Document
        from docx.oxml.ns import qn
        from docx.table import Table
        from docx.text.paragraph import Paragraph
    except ImportError as exc:
        raise ConversionError("缺少 python-docx，请重新安装依赖") from exc

    document = Document(str(src))
    lines: list[str] = []

    def walk(parent_elm, parent_obj):
        for child in parent_elm.iterchildren():
            if child.tag == qn("w:p"):
                text = Paragraph(child, parent_obj).text.strip()
                if text:
                    lines.append(text)
            elif child.tag == qn("w:tbl"):
                table = Table(child, parent_obj)
                for row in table.rows:
                    cells = [cell.text.replace("\n", " ").strip() for cell in row.cells]
                    if any(cells):
                        lines.append("\t".join(cells))
                lines.append("")

    walk(document.element.body, document)
    return _text_output(src, dst_dir, "\n".join(lines), "已提取正文段落与表格")


@register("docx", "html")
def docx_to_html(src: Path, dst_dir: Path, target: str) -> tuple[Path, str]:
    try:
        import mammoth
    except ImportError as exc:
        raise ConversionError("缺少 mammoth，请重新安装依赖") from exc

    def embed_image(image):
        with image.open() as handle:
            data = base64.b64encode(handle.read()).decode("ascii")
        return {"src": f"data:{image.content_type};base64,{data}"}

    try:
        result = mammoth.convert_to_html(
            str(src), convert_image=mammoth.images.img_element(embed_image)
        )
    except Exception as exc:
        LOGGER.exception("docx 转 HTML 失败")
        raise ConversionError(f"Word 转 HTML 失败：{exc}") from exc

    body = result.value or "<p>（空文档）</p>"
    out = dst_dir / f"{src.stem}.html"
    out.write_text(_html_document(body, src.stem), encoding="utf-8", newline="\n")
    warnings = [item.message for item in result.messages if getattr(item, "message", "")]
    note = "已保留标题、列表与表格，图片已内嵌"
    if warnings:
        note = f"{note}；{len(warnings)} 条版式提示"
    return out, note


def _md_inline(paragraph) -> str:
    parts: list[str] = []
    for run in paragraph.runs:
        text = run.text
        if not text:
            continue
        if run.bold and run.italic:
            text = f"***{text}***"
        elif run.bold:
            text = f"**{text}**"
        elif run.italic:
            text = f"*{text}*"
        parts.append(text)
    return "".join(parts).strip()


def _md_table(table) -> str:
    rows: list[list[str]] = []
    for row in table.rows:
        cells = [cell.text.replace("\n", " ").replace("|", "\\|").strip() for cell in row.cells]
        rows.append(cells)
    if not rows:
        return ""
    width = max(len(row) for row in rows)
    rows = [row + [""] * (width - len(row)) for row in rows]
    lines = ["| " + " | ".join(rows[0]) + " |", "| " + " | ".join(["---"] * width) + " |"]
    lines.extend("| " + " | ".join(row) + " |" for row in rows[1:])
    return "\n".join(lines)


@register("docx", "md")
def docx_to_md(src: Path, dst_dir: Path, target: str) -> tuple[Path, str]:
    try:
        from docx import Document
        from docx.oxml.ns import qn
        from docx.table import Table
        from docx.text.paragraph import Paragraph
    except ImportError as exc:
        raise ConversionError("缺少 python-docx，请重新安装依赖") from exc

    document = Document(str(src))
    blocks: list[str] = []
    for child in document.element.body.iterchildren():
        if child.tag == qn("w:p"):
            paragraph = Paragraph(child, document)
            text = _md_inline(paragraph)
            if not text:
                continue
            style = (paragraph.style.name or "").lower()
            match = re.match(r"heading (\d+)", style)
            if match:
                level = min(6, max(1, int(match.group(1))))
                blocks.append(f"{'#' * level} {text}")
            elif style in {"title", "标题"}:
                blocks.append(f"# {text}")
            elif "list bullet" in style or "列表段落" in style:
                blocks.append(f"- {text}")
            elif "list number" in style:
                blocks.append(f"1. {text}")
            else:
                blocks.append(text)
        elif child.tag == qn("w:tbl"):
            table = _md_table(Table(child, document))
            if table:
                blocks.append(table)

    content = re.sub(r"\n{3,}", "\n\n", "\n\n".join(blocks)).strip() + "\n"
    out = dst_dir / f"{src.stem}.md"
    out.write_text(content, encoding="utf-8", newline="\n")
    return out, "已保留标题、列表与表格；复杂浮动版式不会出现在 Markdown 中"


@register("docx", "pdf")
def docx_to_pdf(src: Path, dst_dir: Path, target: str) -> tuple[Path, str]:
    out = dst_dir / f"{src.stem}.pdf"
    reason = "未检测到可用的 Microsoft Word"
    if office_available():
        try:
            _office_convert("word", src, out)
            return out, "已使用本机 Microsoft Word 转换，保留原始版式"
        except ConversionError as exc:
            reason = str(exc)
            LOGGER.warning("Word COM 失败，改用内置排版：%s", exc)
        except Exception as exc:
            LOGGER.exception("Word COM 出现未预期错误")
            reason = f"Word 转换异常：{exc}"
    out.unlink(missing_ok=True)
    try:
        _docx_to_pdf_fallback(src, out)
    except ConversionError:
        raise
    except Exception as exc:
        LOGGER.exception("内置排版失败")
        raise ConversionError("无法解析该 Word 文件，可能已损坏或不是有效的 .docx") from exc
    return out, f"{reason}；已用内置排版引擎生成，字体与行距可能与 Word 有出入"


def _cjk_font() -> str:
    try:
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
    except ImportError as exc:
        raise ConversionError("缺少 reportlab，无法使用内置排版引擎") from exc

    candidates = [
        BASE_DIR / "assets" / "fonts" / "NotoSansSC-Regular.ttf",
        Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts" / "msyh.ttc",
        Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts" / "simsun.ttc",
        Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts" / "simhei.ttf",
    ]
    for candidate in candidates:
        if not candidate.exists():
            continue
        try:
            kwargs = {"subfontIndex": 0} if candidate.suffix.lower() == ".ttc" else {}
            pdfmetrics.registerFont(TTFont("CJK", str(candidate), **kwargs))
            return "CJK"
        except Exception:
            LOGGER.debug("字体注册失败：%s", candidate, exc_info=True)

    from reportlab.pdfbase.cidfonts import UnicodeCIDFont

    pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
    return "STSong-Light"


def _docx_to_pdf_fallback(src: Path, out: Path) -> None:
    try:
        from docx import Document
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.oxml.ns import qn
        from docx.table import Table
        from docx.text.paragraph import Paragraph
        from reportlab.lib import colors
        from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT, TA_RIGHT
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.styles import ParagraphStyle
        from reportlab.lib.units import mm
        from reportlab.platypus import Paragraph as RLParagraph
        from reportlab.platypus import SimpleDocTemplate, Spacer, Table as RLTable, TableStyle
    except ImportError as exc:
        raise ConversionError("缺少内置排版所需依赖，请重新安装 requirements.txt") from exc

    font = _cjk_font()
    document = Document(str(src))
    styles = {
        "normal": ParagraphStyle("normal", fontName=font, fontSize=10.5, leading=18, spaceAfter=5, wordWrap="CJK"),
        "h1": ParagraphStyle("h1", fontName=font, fontSize=18, leading=26, spaceBefore=10, spaceAfter=10, textColor=colors.HexColor("#15202b")),
        "h2": ParagraphStyle("h2", fontName=font, fontSize=15, leading=22, spaceBefore=8, spaceAfter=7, textColor=colors.HexColor("#1f3a5f")),
        "h3": ParagraphStyle("h3", fontName=font, fontSize=12.5, leading=20, spaceBefore=6, spaceAfter=5, textColor=colors.HexColor("#2f4f6f")),
        "cell": ParagraphStyle("cell", fontName=font, fontSize=8.5, leading=12, wordWrap="CJK"),
    }
    alignments = {
        WD_ALIGN_PARAGRAPH.CENTER: TA_CENTER,
        WD_ALIGN_PARAGRAPH.RIGHT: TA_RIGHT,
        WD_ALIGN_PARAGRAPH.JUSTIFY: TA_JUSTIFY,
        WD_ALIGN_PARAGRAPH.LEFT: TA_LEFT,
    }

    def paragraph_style(paragraph) -> ParagraphStyle:
        name = (paragraph.style.name or "").lower()
        if "heading 1" in name or name in {"title", "标题"}:
            return styles["h1"]
        if "heading 2" in name or "标题 2" in name:
            return styles["h2"]
        if "heading 3" in name or "标题 3" in name:
            return styles["h3"]
        return styles["normal"]

    def make_paragraph(paragraph) -> object | None:
        text = paragraph.text.strip()
        if not text:
            return Spacer(1, 4)
        style = paragraph_style(paragraph)
        style = ParagraphStyle(
            f"p{id(paragraph)}",
            parent=style,
            alignment=alignments.get(paragraph.alignment, style.alignment),
        )
        safe = html.escape(text).replace("\n", "<br/>")
        if (paragraph.style.name or "").lower().startswith("list"):
            safe = "• " + safe
        return RLParagraph(safe, style)

    def make_table(table: Table) -> RLTable:
        rows: list[list[object]] = []
        max_columns = max((len(row.cells) for row in table.rows), default=0)
        for row in table.rows:
            cells = [cell.text.replace("\n", " ").strip() for cell in row.cells]
            while len(cells) < max_columns:
                cells.append("")
            rows.append([RLParagraph(html.escape(text) or " ", styles["cell"]) for text in cells])
        if not rows:
            return RLTable([[" "]], colWidths=[120])
        widths = [max(45, 480 / max(1, max_columns))] * max_columns
        result = RLTable(rows, colWidths=widths, repeatRows=1)
        result.setStyle(
            TableStyle(
                [
                    ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#b8c2cc")),
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#eef2f7")),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 4),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                    ("TOPPADDING", (0, 0), (-1, -1), 3),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                ]
            )
        )
        return result

    story: list[object] = []
    for child in document.element.body.iterchildren():
        if child.tag == qn("w:p"):
            paragraph = Paragraph(child, document)
            if paragraph.paragraph_format.page_break_before and story:
                from reportlab.platypus import PageBreak

                story.append(PageBreak())
            item = make_paragraph(paragraph)
            if item is not None:
                story.append(item)
        elif child.tag == qn("w:tbl"):
            story.append(make_table(Table(child, document)))
            story.append(Spacer(1, 6))

    if not story:
        story.append(RLParagraph("（空文档）", styles["normal"]))

    doc = SimpleDocTemplate(
        str(out),
        pagesize=A4,
        leftMargin=18 * mm,
        rightMargin=18 * mm,
        topMargin=18 * mm,
        bottomMargin=18 * mm,
        title=src.stem,
        author="文档工坊",
    )
    doc.build(story)


# ---------------------------------------------------------------------------
# excel / csv
# ---------------------------------------------------------------------------

ILLEGAL_NAME = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def _sheet_filename(name: str, fallback: str) -> str:
    clean = ILLEGAL_NAME.sub("_", name or "").strip().strip(".")
    return (clean or fallback)[:60]


def _safe_sheet_title(name: str, used: set[str]) -> str:
    base = ILLEGAL_NAME.sub("_", name or "Sheet").strip().strip("'")[:31] or "Sheet"
    title = base
    index = 2
    while title.lower() in used:
        suffix = f"_{index}"
        title = base[: 31 - len(suffix)] + suffix
        index += 1
    used.add(title.lower())
    return title


def _cell_text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, (dt.datetime, dt.date, dt.time)):
        return value.isoformat(sep=" ") if isinstance(value, dt.datetime) else value.isoformat()
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    return str(value)


def _write_csv(path: Path, rows) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle, lineterminator="\n")
        for row in rows:
            writer.writerow([_cell_text(value) for value in row])


def _csv_bytes(rows) -> bytes:
    buffer = io.StringIO(newline="")
    writer = csv.writer(buffer, lineterminator="\n")
    for row in rows:
        writer.writerow([_cell_text(value) for value in row])
    return buffer.getvalue().encode("utf-8-sig")


@register("xlsx", "csv")
def xlsx_to_csv(src: Path, dst_dir: Path, target: str) -> tuple[Path, str]:
    try:
        import openpyxl
    except ImportError as exc:
        raise ConversionError("缺少 openpyxl，请重新安装依赖") from exc

    workbook = openpyxl.load_workbook(str(src), read_only=True, data_only=True)
    try:
        names = list(workbook.sheetnames)
        if not names:
            raise ConversionError("Excel 文件中没有工作表")
        if len(names) == 1:
            out = dst_dir / f"{src.stem}.csv"
            sheet = workbook[names[0]]
            _write_csv(out, sheet.iter_rows(values_only=True))
            return out, "已导出唯一工作表，编码为 UTF-8 带 BOM"

        out = dst_dir / f"{src.stem}.zip"
        used: set[str] = set()
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as archive:
            for index, name in enumerate(names, start=1):
                sheet = workbook[name]
                filename = f"{src.stem}_{_sheet_filename(name, f'Sheet{index}')}.csv"
                if filename in used:
                    filename = f"{src.stem}_{index}.csv"
                used.add(filename)
                archive.writestr(filename, _csv_bytes(sheet.iter_rows(values_only=True)))
        return out, f"共 {len(names)} 个工作表，已分别导出并打包为 ZIP"
    finally:
        workbook.close()


def _xls_value(book, cell):
    import xlrd

    if cell.ctype == xlrd.XL_CELL_DATE:
        try:
            return xlrd.xldate.xldate_as_datetime(cell.value, book.datemode)
        except (ValueError, OverflowError):
            return cell.value
    if cell.ctype == xlrd.XL_CELL_BOOLEAN:
        return bool(cell.value)
    if cell.ctype in {xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK, xlrd.XL_CELL_ERROR}:
        return None
    return cell.value


def _read_xls_sheets(src: Path) -> list[tuple[str, list[list[object]]]]:
    try:
        import xlrd
    except ImportError as exc:
        raise ConversionError("缺少 xlrd，无法读取旧版 .xls，请重新安装依赖") from exc
    try:
        book = xlrd.open_workbook(str(src), on_demand=True)
    except Exception as exc:
        raise ConversionError(f"无法读取 .xls 文件：{exc}") from exc
    try:
        sheets: list[tuple[str, list[list[object]]]] = []
        for index in range(book.nsheets):
            sheet = book.sheet_by_index(index)
            rows = [
                [_xls_value(book, cell) for cell in sheet.row(row_index)]
                for row_index in range(sheet.nrows)
            ]
            sheets.append((sheet.name, rows))
        return sheets
    finally:
        book.release_resources()


@register("xls", "xlsx")
def xls_to_xlsx(src: Path, dst_dir: Path, target: str) -> tuple[Path, str]:
    try:
        from openpyxl import Workbook
    except ImportError as exc:
        raise ConversionError("缺少 openpyxl，请重新安装依赖") from exc

    sheets = _read_xls_sheets(src)
    if not sheets:
        raise ConversionError("Excel 文件中没有工作表")
    workbook = Workbook()
    workbook.remove(workbook.active)
    used: set[str] = set()
    for name, rows in sheets:
        sheet = workbook.create_sheet(_safe_sheet_title(name, used))
        for row in rows:
            sheet.append(list(row))
    out = dst_dir / f"{src.stem}.xlsx"
    workbook.save(str(out))
    return out, f"已转换 {len(sheets)} 个工作表；公式转换为当前计算值"


@register("xls", "csv")
def xls_to_csv(src: Path, dst_dir: Path, target: str) -> tuple[Path, str]:
    sheets = _read_xls_sheets(src)
    if not sheets:
        raise ConversionError("Excel 文件中没有工作表")
    if len(sheets) == 1:
        out = dst_dir / f"{src.stem}.csv"
        _write_csv(out, sheets[0][1])
        return out, "已导出唯一工作表，编码为 UTF-8 带 BOM"

    out = dst_dir / f"{src.stem}.zip"
    used: set[str] = set()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as archive:
        for index, (name, rows) in enumerate(sheets, start=1):
            filename = f"{src.stem}_{_sheet_filename(name, f'Sheet{index}')}.csv"
            if filename in used:
                filename = f"{src.stem}_{index}.csv"
            used.add(filename)
            archive.writestr(filename, _csv_bytes(rows))
    return out, f"共 {len(sheets)} 个工作表，已分别导出并打包为 ZIP"


@register("csv", "xlsx")
def csv_to_xlsx(src: Path, dst_dir: Path, target: str) -> tuple[Path, str]:
    try:
        from openpyxl import Workbook
    except ImportError as exc:
        raise ConversionError("缺少 openpyxl，请重新安装依赖") from exc

    text = _read_text(src)
    rows = list(csv.reader(io.StringIO(text)))
    if not rows:
        raise ConversionError("CSV 文件为空，无法转换")
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = _safe_sheet_title(src.stem, set())
    for row in rows:
        sheet.append(row)
    out = dst_dir / f"{src.stem}.xlsx"
    workbook.save(str(out))
    return out, "已写入 Excel；单元格按原文本保存，编码自动识别"


def _xlsx_html_rows(sheet) -> tuple[list[str], bool]:
    rows: list[str] = []
    truncated = False
    for index, row in enumerate(sheet.iter_rows(values_only=True), start=1):
        if index > MAX_PDF_ROWS:
            truncated = True
            break
        values = list(row)
        if len(values) > MAX_PDF_COLS:
            truncated = True
            values = values[:MAX_PDF_COLS]
        cells = [html.escape(_cell_text(value)) for value in values]
        if not any(cells):
            continue
        tag = "th" if not rows else "td"
        rows.append("<tr>" + "".join(f"<{tag}>{cell or ' '}</{tag}>" for cell in cells) + "</tr>")
    return rows, truncated


def _xlsx_to_pdf_fallback(src: Path, out: Path) -> bool:
    try:
        import openpyxl
    except ImportError as exc:
        raise ConversionError("缺少 openpyxl，请重新安装依赖") from exc

    workbook = openpyxl.load_workbook(str(src), read_only=True, data_only=True)
    truncated = False
    blocks: list[str] = [f"<h1>{html.escape(src.stem)}</h1>"]
    try:
        for sheet in workbook.worksheets:
            rows, sheet_truncated = _xlsx_html_rows(sheet)
            truncated = truncated or sheet_truncated
            blocks.append(f"<h2>{html.escape(sheet.title)}</h2>")
            if rows:
                blocks.append("<table>" + "".join(rows) + "</table>")
            else:
                blocks.append("<p>（空工作表）</p>")
    finally:
        workbook.close()
    _html_to_pdf("".join(blocks), out, src.stem)
    return truncated


@register("xlsx", "pdf")
def xlsx_to_pdf(src: Path, dst_dir: Path, target: str) -> tuple[Path, str]:
    out = dst_dir / f"{src.stem}.pdf"
    reason = "未检测到 Microsoft Excel"
    if excel_available():
        try:
            _office_convert("excel", src, out)
            return out, "已使用本机 Microsoft Excel 转换，保留原始工作表版式"
        except ConversionError as exc:
            reason = str(exc)
            LOGGER.warning("Excel COM 失败，改用 HTML 表格排版：%s", exc)
        except Exception as exc:
            LOGGER.exception("Excel COM 出现未预期错误")
            reason = f"Excel 转换异常：{exc}"
    out.unlink(missing_ok=True)
    truncated = _xlsx_to_pdf_fallback(src, out)
    note = f"{reason}；已用 HTML 表格近似排版"
    if truncated:
        note = f"{note}；超长工作表仅渲染前 {MAX_PDF_ROWS} 行、{MAX_PDF_COLS} 列"
    return out, note


# ---------------------------------------------------------------------------
# pdf
# ---------------------------------------------------------------------------

@register("pdf", "txt")
def pdf_to_txt(src: Path, dst_dir: Path, target: str) -> tuple[Path, str]:
    doc = _open_pdf(src)
    try:
        pages = [page.get_text("text") for page in doc]
    finally:
        doc.close()
    content = "\n\n".join(part.strip() for part in pages if part.strip())
    note = "已按页提取文字"
    if not content:
        note = "未提取到文字，可能是扫描件；本版本不包含 OCR"
    return _text_output(src, dst_dir, content, note)


@register("pdf", "docx")
def pdf_to_docx(src: Path, dst_dir: Path, target: str) -> tuple[Path, str]:
    doc = _open_pdf(src)
    doc.close()
    try:
        from pdf2docx import Converter
    except ImportError as exc:
        raise ConversionError("缺少 pdf2docx，请重新安装依赖") from exc

    out = dst_dir / f"{src.stem}.docx"
    converter = None
    try:
        converter = Converter(str(src))
        converter.convert(str(out), multi_processing=False)
    except Exception as exc:
        LOGGER.exception("pdf2docx 转换失败")
        raise ConversionError(f"PDF 转 Word 失败：{exc}") from exc
    finally:
        if converter is not None:
            try:
                converter.close()
            except Exception:
                LOGGER.debug("关闭 pdf2docx 失败", exc_info=True)
    return out, "版式由 pdf2docx 推断，复杂表格或扫描件可能需要手工调整"


def _pdf_to_images(src: Path, dst_dir: Path, image_ext: str) -> tuple[Path, str]:
    try:
        import fitz
    except ImportError as exc:
        raise ConversionError("缺少 PyMuPDF，请重新安装依赖") from exc

    doc = _open_pdf(src)
    count = doc.page_count
    if count == 0:
        doc.close()
        raise ConversionError("PDF 中没有可导出的页面")
    matrix = fitz.Matrix(150 / 72, 150 / 72)
    try:
        if count == 1:
            pix = doc[0].get_pixmap(matrix=matrix, colorspace=fitz.csRGB, alpha=False)
            out = dst_dir / f"{src.stem}.{image_ext}"
            out.write_bytes(pix.tobytes(image_ext, jpg_quality=92))
            return out, f"已按 150 DPI 导出第 1 页为 {image_ext.upper()}"

        out = dst_dir / f"{src.stem}.zip"
        with zipfile.ZipFile(out, "w", zipfile.ZIP_STORED) as archive:
            for index, page in enumerate(doc, start=1):
                pix = page.get_pixmap(matrix=matrix, colorspace=fitz.csRGB, alpha=False)
                archive.writestr(f"{index:03d}.{image_ext}", pix.tobytes(image_ext, jpg_quality=92))
        return out, f"共 {count} 页，已按 150 DPI 导出为 {image_ext.upper()} 并打包为 ZIP"
    finally:
        doc.close()


@register("pdf", "jpg")
def pdf_to_jpg(src: Path, dst_dir: Path, target: str) -> tuple[Path, str]:
    return _pdf_to_images(src, dst_dir, "jpg")


@register("pdf", "png")
def pdf_to_png(src: Path, dst_dir: Path, target: str) -> tuple[Path, str]:
    return _pdf_to_images(src, dst_dir, "png")


# ---------------------------------------------------------------------------
# 图片
# ---------------------------------------------------------------------------

def _image_to_pdf(src: Path, dst_dir: Path, target: str) -> tuple[Path, str]:
    try:
        import fitz
        from PIL import Image, ImageOps
    except ImportError as exc:
        raise ConversionError("缺少 Pillow/PyMuPDF，请重新安装依赖") from exc

    out = dst_dir / f"{src.stem}.pdf"
    note = "已按页面比例居中放入 PDF"
    tmp_path: Path | None = None
    image = None
    doc = None
    try:
        image = Image.open(src)
        frames = getattr(image, "n_frames", 1)
        if frames > 1:
            note = f"检测到多帧图片，仅处理第一帧；{note}"
        image.seek(0)
        image = ImageOps.exif_transpose(image)
        if image.mode in {"RGBA", "LA", "P"}:
            rgba = image.convert("RGBA")
            background = Image.new("RGB", rgba.size, "white")
            background.paste(rgba, mask=rgba.split()[-1])
            image = background
        else:
            image = image.convert("RGB")

        with tempfile.NamedTemporaryFile(suffix=".jpg", dir=dst_dir, delete=False) as temp:
            tmp_path = Path(temp.name)
        image.save(tmp_path, "JPEG", quality=92)

        landscape = image.width > image.height
        page_width, page_height = (842, 595) if landscape else (595, 842)
        doc = fitz.open()
        page = doc.new_page(width=page_width, height=page_height)
        margin = 36
        scale = min(
            (page_width - margin * 2) / image.width,
            (page_height - margin * 2) / image.height,
        )
        width = image.width * scale
        height = image.height * scale
        rect = fitz.Rect(
            (page_width - width) / 2,
            (page_height - height) / 2,
            (page_width + width) / 2,
            (page_height + height) / 2,
        )
        page.insert_image(rect, filename=str(tmp_path))
        doc.save(str(out), garbage=3, deflate=True)
        return out, note
    finally:
        if doc is not None:
            doc.close()
        if image is not None:
            image.close()
        if tmp_path is not None:
            tmp_path.unlink(missing_ok=True)


for _ext in IMAGE_EXTS:
    HANDLERS[(_ext, "pdf")] = _image_to_pdf


def _image_convert(src: Path, dst_dir: Path, target: str) -> tuple[Path, str]:
    try:
        from PIL import Image, ImageOps
    except ImportError as exc:
        raise ConversionError("缺少 Pillow，请重新安装依赖") from exc

    target = "jpg" if target == "jpeg" else target
    image = None
    try:
        image = Image.open(src)
        frames = getattr(image, "n_frames", 1)
        image.seek(0)
        image = ImageOps.exif_transpose(image)
        note = f"已转换为 {target.upper()}"
        if frames > 1:
            note = f"检测到多帧图片，仅处理第一帧；{note}"

        out = dst_dir / f"{src.stem}.{target}"
        if target == "jpg":
            rgba = image.convert("RGBA")
            background = Image.new("RGB", rgba.size, "white")
            background.paste(rgba, mask=rgba.split()[-1])
            background.save(out, "JPEG", quality=92, optimize=True)
            background.close()
            rgba.close()
        elif target == "png":
            if image.mode in {"RGBA", "LA", "P"}:
                converted = image.convert("RGBA")
            else:
                converted = image.convert("RGB")
            converted.save(out, "PNG", optimize=True)
            converted.close()
        elif target == "webp":
            converted = image.convert("RGBA") if image.mode in {"RGBA", "LA", "P"} else image.convert("RGB")
            converted.save(out, "WEBP", quality=90, method=6)
            converted.close()
        elif target == "bmp":
            converted = image.convert("RGB")
            converted.save(out, "BMP")
            converted.close()
        elif target == "tiff":
            converted = image.convert("RGBA") if image.mode in {"RGBA", "LA", "P"} else image.convert("RGB")
            converted.save(out, "TIFF", compression="tiff_deflate")
            converted.close()
        else:
            raise ConversionError(f"不支持的图片目标格式：{target}")
        return out, note
    finally:
        if image is not None:
            image.close()


for _source_ext in IMAGE_EXTS:
    for _target_ext in IMAGE_TARGETS:
        if _source_ext != _target_ext:
            HANDLERS[(_source_ext, _target_ext)] = _image_convert


@register("png", "docx")
def _image_to_docx_impl(src: Path, dst_dir: Path, target: str) -> tuple[Path, str]:
    try:
        from PIL import Image, ImageOps
        from docx import Document
        from docx.shared import Inches
    except ImportError as exc:
        raise ConversionError("缺少 Pillow/python-docx，请重新安装依赖") from exc

    tmp_path = dst_dir / f".{src.stem}_embed.png"
    image = None
    try:
        image = Image.open(src)
        frames = getattr(image, "n_frames", 1)
        image.seek(0)
        image = ImageOps.exif_transpose(image)
        converted = image.convert("RGBA") if image.mode in {"RGBA", "LA", "P"} else image.convert("RGB")
        converted.save(tmp_path, "PNG")
        width, height = converted.size
        converted.close()

        document = Document()
        if width / max(1, height) > 8.9 / 6.2:
            document.add_picture(str(tmp_path), height=Inches(8.9))
        else:
            document.add_picture(str(tmp_path), width=Inches(6.2))
        document.add_paragraph("")
        out = dst_dir / f"{src.stem}.docx"
        document.save(str(out))
        note = "已按页面宽度等比缩放嵌入 Word"
        if frames > 1:
            note = f"检测到多帧图片，仅处理第一帧；{note}"
        return out, note
    finally:
        if image is not None:
            image.close()
        tmp_path.unlink(missing_ok=True)


for _ext in IMAGE_EXTS:
    HANDLERS[(_ext, "docx")] = _image_to_docx_impl


# ---------------------------------------------------------------------------
# markdown / txt
# ---------------------------------------------------------------------------

@register("md", "html")
def md_to_html(src: Path, dst_dir: Path, target: str) -> tuple[Path, str]:
    try:
        import markdown
    except ImportError as exc:
        raise ConversionError("缺少 markdown，请重新安装依赖") from exc

    text = _read_text(src)
    body = markdown.markdown(text, extensions=["tables", "fenced_code", "toc"])
    title = html.escape(src.stem)
    out = dst_dir / f"{src.stem}.html"
    out.write_text(
        f"""<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{title}</title>
<style>
body{{max-width:860px;margin:40px auto;padding:0 24px;font:16px/1.75 -apple-system,"Segoe UI","Microsoft YaHei",sans-serif;color:#1d2733;background:#fff}}
h1,h2,h3{{line-height:1.35;color:#102a43}}
table{{border-collapse:collapse;width:100%;margin:16px 0}}
th,td{{border:1px solid #d8e0e8;padding:8px 10px;text-align:left}}
th{{background:#f1f5f9}}
pre{{background:#f6f8fa;padding:14px;overflow:auto;border-radius:8px}}
code{{font-family:Consolas,monospace;background:#f6f8fa;padding:2px 4px;border-radius:4px}}
blockquote{{margin:16px 0;padding:8px 16px;border-left:4px solid #9db4c8;color:#4a5b6c;background:#f8fafc}}
img{{max-width:100%}}
</style>
</head>
<body>
{body}
</body>
</html>
""",
        encoding="utf-8",
        newline="\n",
    )
    return out, "已生成独立 HTML 文件，样式内嵌"


@register("md", "pdf")
def md_to_pdf(src: Path, dst_dir: Path, target: str) -> tuple[Path, str]:
    try:
        import markdown
    except ImportError as exc:
        raise ConversionError("缺少 markdown，请重新安装依赖") from exc

    text = _read_text(src)
    body = markdown.markdown(text, extensions=["tables", "fenced_code", "toc"])
    out = dst_dir / f"{src.stem}.pdf"
    pages = _html_to_pdf(body or "<p>（空文档）</p>", out, src.stem)
    return out, f"已用内置 HTML 排版生成 {pages} 页 PDF；本地图片路径不会自动嵌入"


@register("txt", "pdf")
def txt_to_pdf(src: Path, dst_dir: Path, target: str) -> tuple[Path, str]:
    text = _read_text(src)
    paragraphs = [
        "<p>&nbsp;</p>" if not line.strip() else f"<p>{html.escape(line)}</p>"
        for line in text.splitlines()
    ]
    out = dst_dir / f"{src.stem}.pdf"
    pages = _html_to_pdf("".join(paragraphs) or "<p>（空文件）</p>", out, src.stem)
    return out, f"已按原段落生成 {pages} 页 PDF"


@register("txt", "docx")
def txt_to_docx(src: Path, dst_dir: Path, target: str) -> tuple[Path, str]:
    try:
        from docx import Document
        from docx.oxml.ns import qn
        from docx.shared import Pt
    except ImportError as exc:
        raise ConversionError("缺少 python-docx，请重新安装依赖") from exc

    text = _read_text(src)
    document = Document()
    normal = document.styles["Normal"]
    normal.font.name = "Microsoft YaHei"
    normal.font.size = Pt(11)
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
    for line in text.splitlines():
        document.add_paragraph(line)
    out = dst_dir / f"{src.stem}.docx"
    document.save(str(out))
    return out, "已按原段落写入 Word"


# ---------------------------------------------------------------------------
# 旧版 Office / PowerPoint
# ---------------------------------------------------------------------------

@register("doc", "pdf")
def doc_to_pdf(src: Path, dst_dir: Path, target: str) -> tuple[Path, str]:
    if not office_available():
        raise ConversionError("本机 Word 不可用，.doc 转 PDF 需要安装 Microsoft Office")
    out = dst_dir / f"{src.stem}.pdf"
    _office_convert("word", src, out)
    return out, "已使用本机 Microsoft Word 转换，保留原始版式"


@register("doc", "docx")
def doc_to_docx(src: Path, dst_dir: Path, target: str) -> tuple[Path, str]:
    if not office_available():
        raise ConversionError("本机 Word 不可用，.doc 转 .docx 需要安装 Microsoft Office")
    out = dst_dir / f"{src.stem}.docx"
    _office_convert("word", src, out)
    return out, "已使用本机 Microsoft Word 转换为 DOCX"


@register("pptx", "pdf")
@register("ppt", "pdf")
def powerpoint_to_pdf(src: Path, dst_dir: Path, target: str) -> tuple[Path, str]:
    if not powerpoint_available():
        raise ConversionError("本机 PowerPoint 不可用，该格式需要安装 Microsoft Office PowerPoint")
    out = dst_dir / f"{src.stem}.pdf"
    _office_convert("powerpoint", src, out)
    return out, "已使用本机 Microsoft PowerPoint 转换"


@register("pptx", "png")
def pptx_to_png(src: Path, dst_dir: Path, target: str) -> tuple[Path, str]:
    if not powerpoint_available():
        raise ConversionError("本机 PowerPoint 不可用，PPTX 转图片需要安装 Microsoft Office PowerPoint")
    out = dst_dir / f"{src.stem}.zip"
    _office_convert("powerpoint_png", src, out)
    return out, "已导出全部幻灯片为 PNG，并打包为 ZIP"