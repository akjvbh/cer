"""转换器注册表：所有「源格式 → 目标格式」的实现都放在这里。"""

from __future__ import annotations

import csv
import html
import importlib.util
import io
import logging
import os
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
OFFICE_LOCK = threading.RLock()
_OFFICE_CACHE: tuple[float, bool] | None = None
_OFFICE_CACHE_LOCK = threading.Lock()


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


def office_available(force: bool = False) -> bool:
    """注册表/安装路径探测，不启动 Word；结果缓存 60 秒。"""
    global _OFFICE_CACHE
    enabled = os.getenv("OFFICE_ENABLED", "1").lower() not in {"0", "false", "no"}
    if not enabled or sys.platform != "win32":
        return False
    if importlib.util.find_spec("win32com") is None:
        return False

    now = time.monotonic()
    with _OFFICE_CACHE_LOCK:
        if not force and _OFFICE_CACHE and now - _OFFICE_CACHE[0] < 60:
            return _OFFICE_CACHE[1]

    found = False
    try:
        import winreg

        for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            try:
                with winreg.OpenKey(
                    root,
                    r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE",
                ) as key:
                    value, _ = winreg.QueryValueEx(key, "")
                    if value and Path(value).exists():
                        found = True
                        break
            except OSError:
                continue
    except Exception:
        LOGGER.debug("读取 Office 注册表失败", exc_info=True)

    if not found:
        for env_name in ("ProgramFiles", "ProgramFiles(x86)"):
            base = os.environ.get(env_name)
            if base and (Path(base) / "Microsoft Office" / "root" / "Office16" / "WINWORD.EXE").exists():
                found = True
                break

    with _OFFICE_CACHE_LOCK:
        _OFFICE_CACHE = (now, found)
    return found


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


def _office_convert(kind: str, src: Path, dst: Path) -> None:
    image = "WINWORD.EXE" if kind == "word" else "EXCEL.EXE"
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
        LOGGER.error("Office 转换失败：%s", completed.stderr.strip())
        raise ConversionError("Office 转换失败，已切换到内置排版引擎")
    if not dst.exists() or dst.stat().st_size == 0:
        raise ConversionError("Office 没有生成有效 PDF")


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
        raise ConversionError(f"内置排版失败：{exc}") from exc
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


@register("pdf", "jpg")
def pdf_to_jpg(src: Path, dst_dir: Path, target: str) -> tuple[Path, str]:
    try:
        import fitz
        from PIL import Image
    except ImportError as exc:
        raise ConversionError("缺少 PDF/图片处理依赖，请重新安装 requirements.txt") from exc

    doc = _open_pdf(src)
    count = doc.page_count
    if count == 0:
        doc.close()
        raise ConversionError("PDF 中没有可导出的页面")
    matrix = fitz.Matrix(150 / 72, 150 / 72)
    try:
        if count == 1:
            page = doc[0]
            pix = page.get_pixmap(matrix=matrix, colorspace=fitz.csRGB, alpha=False)
            out = dst_dir / f"{src.stem}.jpg"
            Image.frombytes("RGB", (pix.width, pix.height), pix.samples).save(out, "JPEG", quality=92)
            return out, "已按 150 DPI 导出第 1 页"
        out = dst_dir / f"{src.stem}.zip"
        with zipfile.ZipFile(out, "w", zipfile.ZIP_STORED) as archive:
            for index, page in enumerate(doc, start=1):
                pix = page.get_pixmap(matrix=matrix, colorspace=fitz.csRGB, alpha=False)
                buffer = io.BytesIO()
                Image.frombytes("RGB", (pix.width, pix.height), pix.samples).save(
                    buffer, "JPEG", quality=92
                )
                archive.writestr(f"{index:03d}.jpg", buffer.getvalue())
        return out, f"共 {count} 页，已按 150 DPI 导出并打包为 ZIP"
    finally:
        doc.close()


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