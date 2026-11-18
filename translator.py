"""文献翻译：段落抽取、模型调用与成品生成。"""

from __future__ import annotations

import logging
import os
import re
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from pathlib import Path
from statistics import median
from typing import Callable

import fitz
import requests
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from docx.table import Table
from docx.text.paragraph import Paragraph

from converter import ConversionError

LOGGER = logging.getLogger("docworkshop.translator")

BASE_DIR = Path(__file__).resolve().parent
FONT_DIR = BASE_DIR / "assets" / "fonts"
GLOSSARY_PATH = BASE_DIR / "glossary.txt"

CHUNK_CHAR_LIMIT = 2000
MAX_WORKERS = 3
REQUEST_TIMEOUT_SECONDS = 120
RETRY_DELAYS = (2, 4, 8)
SEGMENT_MARKER = "<<<SEG>>>"
MAX_GLOSSARY_CHARS = 8000
LAYOUT_MIN_FONT_SIZE = 6.0
LAYOUT_FONT_STEP = 0.5
LAYOUT_LINE_HEIGHT = 1.22
LAYOUT_FONT_ALIAS = "DocWorkshopSC"
DEFAULT_BASE_URL = "https://api.deepseek.com/v1"
DEFAULT_MODEL = "deepseek-chat"

ProgressFn = Callable[[int, str], None]

_PAGE_NUMBER_RE = re.compile(r"^[\s\-–—·•]*(?:\d{1,4}|[ivxlcdmIVXLCDM]{1,8})[\s\-–—·•]*$")
_TEXT_RE = re.compile(r"[A-Za-z\u4e00-\u9fff]")
_CJK_RE = re.compile(r"[\u4e00-\u9fff]")
_SENTENCE_RE = re.compile(r"(?<=[。！？!?；;])|(?<=\.)\s+")


class TranslateError(Exception):
    """可直接展示给用户的翻译错误。"""


@dataclass
class Segment:
    text: str
    style: str = ""
    page: int = 0
    rect: tuple[float, float, float, float] | None = None
    font_size: float = 0.0
    translated: str = ""


# ---------------------------------------------------------------------------
# 配置
# ---------------------------------------------------------------------------

def api_key() -> str:
    return os.getenv("TRANSLATE_API_KEY", "").strip()


def is_configured() -> bool:
    return bool(api_key())


def base_url() -> str:
    return (os.getenv("TRANSLATE_BASE_URL", "").strip() or DEFAULT_BASE_URL).rstrip("/")


def model_name() -> str:
    return os.getenv("TRANSLATE_MODEL", "").strip() or DEFAULT_MODEL


def mask_key(value: str) -> str:
    if not value:
        return "(未配置)"
    if len(value) <= 6:
        return "sk-****"
    return f"{value[:3]}****{value[-2:]}"


def read_glossary_file() -> str:
    if not GLOSSARY_PATH.exists():
        return ""
    try:
        return GLOSSARY_PATH.read_text(encoding="utf-8-sig")
    except OSError:
        LOGGER.warning("读取术语表失败：%s", GLOSSARY_PATH, exc_info=True)
        return ""


def parse_glossary(text: str) -> dict[str, str]:
    terms: dict[str, str] = {}
    for raw_line in (text or "").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        source, target = line.split("=", 1)
        source = source.strip()
        target = target.strip()
        if source and target:
            terms[source] = target
    return terms


def merged_glossary(overrides: str) -> dict[str, str]:
    terms = parse_glossary(read_glossary_file())
    terms.update(parse_glossary(overrides))
    return terms


# ---------------------------------------------------------------------------
# 段落抽取
# ---------------------------------------------------------------------------

def _is_cjk(char: str) -> bool:
    return bool(_CJK_RE.match(char))


def _keep_text(text: str) -> bool:
    value = (text or "").strip()
    if len(value) < 2:
        return False
    if _PAGE_NUMBER_RE.match(value):
        return False
    return bool(_TEXT_RE.search(value))


def _merge_lines(text: str) -> str:
    lines = [line.strip() for line in (text or "").splitlines() if line.strip()]
    merged = ""
    for line in lines:
        if not merged:
            merged = line
        elif merged.endswith("-") and line[:1].islower():
            merged = f"{merged[:-1]}{line}"
        elif _is_cjk(merged[-1]) or _is_cjk(line[0]):
            merged = f"{merged}{line}"
        else:
            merged = f"{merged} {line}"
    return re.sub(r"\s+", " ", merged).strip()


def _bbox_key(bbox) -> tuple[float, float, float, float]:
    return tuple(round(float(value), 1) for value in bbox)  # type: ignore[return-value]


def _block_font_sizes(page: fitz.Page) -> dict[tuple[float, float, float, float], float]:
    sizes: dict[tuple[float, float, float, float], float] = {}
    data = page.get_text("dict")
    for block in data.get("blocks", []):
        if block.get("type") != 0:
            continue
        values = [
            float(span.get("size", 0))
            for line in block.get("lines", [])
            for span in line.get("spans", [])
            if str(span.get("text", "")).strip() and float(span.get("size", 0)) > 0
        ]
        if values:
            sizes[_bbox_key(block.get("bbox", (0, 0, 0, 0)))] = max(6.0, min(24.0, median(values)))
    return sizes


def _estimate_font_size(block) -> float:
    height = max(0.0, float(block[3]) - float(block[1]))
    lines = max(1, len([line for line in str(block[4]).splitlines() if line.strip()]))
    return max(6.0, min(24.0, height / lines * 0.82))


def _extract_pdf(src: Path) -> list[Segment]:
    document = fitz.open(str(src))
    try:
        if document.needs_pass:
            raise TranslateError("不支持加密 PDF，请先解除密码保护")
        segments: list[Segment] = []
        for page_number, page in enumerate(document, start=1):
            size_map = _block_font_sizes(page)
            blocks = [
                block
                for block in page.get_text("blocks")
                if len(block) > 6 and block[6] == 0
            ]
            blocks.sort(key=lambda block: (round(float(block[1]), 1), round(float(block[0]), 1)))
            for block in blocks:
                text = _merge_lines(block[4])
                if not _keep_text(text):
                    continue
                bbox = (float(block[0]), float(block[1]), float(block[2]), float(block[3]))
                font_size = size_map.get(_bbox_key(bbox)) or _estimate_font_size(block)
                segments.append(
                    Segment(
                        text=text,
                        page=page_number,
                        rect=bbox,
                        font_size=font_size,
                    )
                )
        return segments
    finally:
        document.close()


def _extract_docx(src: Path) -> list[Segment]:
    document = Document(str(src))
    segments: list[Segment] = []
    for child in document.element.body.iterchildren():
        if child.tag == qn("w:p"):
            paragraph = Paragraph(child, document)
            text = _merge_lines(paragraph.text)
            if not _keep_text(text):
                continue
            style = ""
            try:
                style = paragraph.style.name if paragraph.style is not None else ""
            except (AttributeError, KeyError):
                style = ""
            segments.append(Segment(text=text, style=style))
        elif child.tag == qn("w:tbl"):
            table = Table(child, document)
            for row in table.rows:
                seen: set[int] = set()
                for cell in row.cells:
                    marker = id(cell._tc)
                    if marker in seen:
                        continue
                    seen.add(marker)
                    for paragraph in cell.paragraphs:
                        text = _merge_lines(paragraph.text)
                        if _keep_text(text):
                            segments.append(Segment(text=text))
    return segments


def _read_text_file(src: Path) -> str:
    data = src.read_bytes()
    for encoding in ("utf-8-sig", "gb18030"):
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue
    raise TranslateError("无法识别文本编码，请先转成 UTF-8")


def _extract_txt(src: Path) -> list[Segment]:
    text = _read_text_file(src).replace("\r\n", "\n").replace("\r", "\n")
    segments: list[Segment] = []
    for block in re.split(r"\n\s*\n", text):
        merged = _merge_lines(block)
        if _keep_text(merged):
            segments.append(Segment(text=merged))
    return segments


def extract_segments(src: Path) -> list[Segment]:
    suffix = src.suffix.lower().lstrip(".")
    if suffix == "pdf":
        segments = _extract_pdf(src)
    elif suffix == "docx":
        segments = _extract_docx(src)
    elif suffix == "txt":
        segments = _extract_txt(src)
    else:
        raise TranslateError("文献翻译只支持 PDF、DOCX 和 TXT")
    if not segments:
        raise TranslateError("没有提取到可翻译的文字；扫描版 PDF 需要先做 OCR（本版未内置）")
    return segments


# ---------------------------------------------------------------------------
# 分块与模型调用
# ---------------------------------------------------------------------------

def _split_long_text(text: str, limit: int) -> list[str]:
    if len(text) <= limit:
        return [text]
    parts: list[str] = []
    current = ""
    for sentence in _SENTENCE_RE.split(text):
        piece = sentence or ""
        if not piece:
            continue
        if len(piece) > limit:
            if current:
                parts.append(current)
                current = ""
            for index in range(0, len(piece), limit):
                parts.append(piece[index : index + limit])
            continue
        if current and len(current) + len(piece) > limit:
            parts.append(current)
            current = piece
        else:
            current = f"{current}{piece}"
    if current:
        parts.append(current)
    return parts or [text]


def _build_chunks(segments: list[Segment]) -> list[list[tuple[int, str]]]:
    pieces: list[tuple[int, str]] = []
    for index, segment in enumerate(segments):
        for part in _split_long_text(segment.text, CHUNK_CHAR_LIMIT):
            pieces.append((index, part))

    chunks: list[list[tuple[int, str]]] = []
    current: list[tuple[int, str]] = []
    length = 0
    for index, part in pieces:
        extra = len(part) + (len(SEGMENT_MARKER) if current else 0)
        if current and length + extra > CHUNK_CHAR_LIMIT:
            chunks.append(current)
            current = []
            length = 0
            extra = len(part)
        current.append((index, part))
        length += extra
    if current:
        chunks.append(current)
    return chunks


def _target_label(target_lang: str) -> str:
    return "简体中文" if target_lang == "zh" else "英文"


def _system_prompt(target_lang: str, glossary: dict[str, str]) -> str:
    lines = [
        f"你是一名专业的学术文献翻译引擎，请把用户提供的文本翻译成{_target_label(target_lang)}。",
        "规则：",
        "1. 只输出译文，不要解释、不要添加引号、不要保留原文。",
        f"2. 文本用 {SEGMENT_MARKER} 分隔为若干段；输出必须使用同样的分隔符，顺序和段数完全一致，不能合并或拆分段落。",
        "3. 保留数字、公式、专有名词的准确性，不要输出原文。",
    ]
    if glossary:
        body: list[str] = []
        used = 0
        for source, target in glossary.items():
            line = f"{source}={target}"
            if used + len(line) + 1 > MAX_GLOSSARY_CHARS:
                break
            body.append(line)
            used += len(line) + 1
        if body:
            lines.append("术语表（原文=译文，必须优先采用）：")
            lines.extend(body)
    return "\n".join(lines)


def _provider_error_message(response: requests.Response) -> str:
    try:
        payload = response.json()
    except ValueError:
        payload = None
    if isinstance(payload, dict):
        error = payload.get("error")
        if isinstance(error, dict) and error.get("message"):
            return str(error["message"])
        if isinstance(error, str) and error:
            return error
        if payload.get("message"):
            return str(payload["message"])
    text = (response.text or "").strip()
    return text[:300] if text else "接口没有返回错误详情"


def _request_translation(system: str, user: str) -> str:
    key = api_key()
    url = f"{base_url()}/chat/completions"
    headers = {
        "Authorization": f"Bearer {key}",
        "Content-Type": "application/json",
    }
    payload = {
        "model": model_name(),
        "temperature": 0.2,
        "messages": [
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ],
    }
    last_status = 0
    for attempt in range(len(RETRY_DELAYS) + 1):
        try:
            response = requests.post(
                url,
                json=payload,
                headers=headers,
                timeout=REQUEST_TIMEOUT_SECONDS,
            )
        except requests.Timeout as exc:
            raise TranslateError(
                f"翻译接口请求超时（{REQUEST_TIMEOUT_SECONDS} 秒），请检查网络或接口地址"
            ) from exc
        except requests.RequestException as exc:
            raise TranslateError(f"无法连接翻译接口：{exc}") from exc

        last_status = response.status_code
        if response.status_code == 429 or response.status_code >= 500:
            if attempt < len(RETRY_DELAYS):
                delay = RETRY_DELAYS[attempt]
                LOGGER.warning(
                    "翻译接口返回 HTTP %s，%s 秒后重试（%s/3）",
                    response.status_code,
                    delay,
                    attempt + 1,
                )
                time.sleep(delay)
                continue
            raise TranslateError(
                f"翻译接口持续返回 HTTP {response.status_code}，已重试 3 次，请稍后再试"
            )

        if response.status_code >= 400:
            detail = _provider_error_message(response)
            raise TranslateError(
                f"翻译接口返回错误（HTTP {response.status_code}）：{detail}"
            )

        try:
            data = response.json()
            content = data["choices"][0]["message"]["content"]
        except (ValueError, KeyError, IndexError, TypeError) as exc:
            raise TranslateError("翻译接口返回格式不正确，缺少 choices[0].message.content") from exc
        if not isinstance(content, str) or not content.strip():
            raise TranslateError("翻译接口返回了空译文")
        return content.strip()

    raise TranslateError(f"翻译接口返回 HTTP {last_status}，请稍后再试")


def _translate_chunk(
    index: int,
    total: int,
    chunk: list[tuple[int, str]],
    target_lang: str,
    glossary: dict[str, str],
) -> list[str]:
    texts = [part for _, part in chunk]
    user = f"\n{SEGMENT_MARKER}\n".join(texts)
    try:
        raw = _request_translation(_system_prompt(target_lang, glossary), user)
    except TranslateError as exc:
        first = min(index for index, _ in chunk) + 1
        last = max(index for index, _ in chunk) + 1
        label = f"第 {first} 段" if first == last else f"第 {first}–{last} 段"
        raise TranslateError(f"第 {index + 1}/{total} 块（{label}）翻译失败：{exc}") from exc

    parts = [part.strip() for part in re.split(r"<<<\s*SEG\s*>>>", raw, flags=re.IGNORECASE)]
    if len(parts) != len(texts):
        first = min(index for index, _ in chunk) + 1
        last = max(index for index, _ in chunk) + 1
        label = f"第 {first} 段" if first == last else f"第 {first}–{last} 段"
        raise TranslateError(
            f"第 {index + 1}/{total} 块（{label}）翻译失败："
            f"接口返回 {len(parts)} 段，预期 {len(texts)} 段，请调整术语后重试"
        )
    if any(not part for part in parts):
        raise TranslateError(
            f"第 {index + 1}/{total} 块翻译失败：接口返回了空段落，请重试"
        )
    return parts


def _translate_segments(
    segments: list[Segment],
    target_lang: str,
    glossary: dict[str, str],
    progress: ProgressFn,
) -> None:
    chunks = _build_chunks(segments)
    total = len(chunks)
    results: list[list[str] | None] = [None] * total
    done = 0
    done_lock = threading.Lock()

    def work(chunk_index: int) -> list[str]:
        return _translate_chunk(chunk_index, total, chunks[chunk_index], target_lang, glossary)

    executor = ThreadPoolExecutor(max_workers=MAX_WORKERS, thread_name_prefix="translate")
    try:
        futures = {executor.submit(work, index): index for index in range(total)}
        for future in as_completed(futures):
            chunk_index = futures[future]
            results[chunk_index] = future.result()
            with done_lock:
                done += 1
                current = done
            percent = max(1, min(99, int(current * 100 / total)))
            progress(percent, f"正在翻译第 {current}/{total} 段")
    except Exception:
        for future in futures:
            future.cancel()
        raise
    finally:
        executor.shutdown(wait=False, cancel_futures=True)

    buffers: dict[int, list[str]] = {index: [] for index in range(len(segments))}
    joiner = " " if target_lang == "en" else ""
    for chunk_index, chunk in enumerate(chunks):
        parts = results[chunk_index] or []
        if len(parts) != len(chunk):
            raise TranslateError(f"第 {chunk_index + 1}/{total} 块译文缺失，请重试")
        for (segment_index, _), part in zip(chunk, parts):
            buffers[segment_index].append(part)
    for index, segment in enumerate(segments):
        segment.translated = joiner.join(buffers[index]).strip()
        if not segment.translated:
            raise TranslateError(f"第 {index + 1} 段译文为空，请重试")


# ---------------------------------------------------------------------------
# 成品生成
# ---------------------------------------------------------------------------

def _set_east_asian_font(document: Document, chinese: bool) -> None:
    style = document.styles["Normal"]
    style.font.name = "Segoe UI"
    style.font.size = Pt(10.5)
    rpr = style.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = rpr.makeelement(qn("w:rFonts"), {})
        rpr.append(rfonts)
    rfonts.set(qn("w:eastAsia"), "微软雅黑" if chinese else "Segoe UI")


def _write_bilingual(segments: list[Segment], dst: Path, target_lang: str) -> None:
    document = Document()
    _set_east_asian_font(document, target_lang == "zh")
    document.add_heading("中英对照译文" if target_lang == "zh" else "原文与英文对照", level=1)
    table = document.add_table(rows=1, cols=2)
    table.style = "Table Grid"
    table.autofit = False
    header = table.rows[0].cells
    header[0].text = "原文"
    header[1].text = "译文"
    for cell in header:
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.bold = True
    for segment in segments:
        cells = table.add_row().cells
        cells[0].text = segment.text
        cells[1].text = segment.translated
        cells[0].width = Cm(7.6)
        cells[1].width = Cm(7.6)
        for cell in cells:
            for paragraph in cell.paragraphs:
                paragraph.paragraph_format.space_after = Pt(2)
                for run in paragraph.runs:
                    run.font.size = Pt(9.5)
    document.save(str(dst))


def _write_translated(segments: list[Segment], dst: Path, target_lang: str) -> None:
    document = Document()
    _set_east_asian_font(document, target_lang == "zh")
    for segment in segments:
        style = segment.style or ""
        paragraph = None
        if style.lower().startswith("heading") or style.lower() == "title":
            try:
                paragraph = document.add_paragraph(segment.translated, style=style)
            except KeyError:
                paragraph = document.add_paragraph(segment.translated)
        else:
            try:
                paragraph = document.add_paragraph(segment.translated, style=style) if style else None
            except KeyError:
                paragraph = None
            if paragraph is None:
                paragraph = document.add_paragraph(segment.translated)
    document.save(str(dst))


def _font_candidates() -> list[tuple[Path, str]]:
    candidates: list[tuple[Path, str]] = []
    if FONT_DIR.exists():
        noto_files = sorted(
            [
                path
                for path in FONT_DIR.iterdir()
                if path.is_file()
                and path.suffix.lower() in {".ttf", ".otf", ".ttc"}
                and "noto" in path.name.lower()
            ],
            key=lambda path: (("subset" not in path.name.lower()), path.name.lower()),
        )
        candidates.extend((path, "Noto Sans SC 子集") for path in noto_files)
    system_font = Path(os.getenv("WINDIR", r"C:\Windows")) / "Fonts" / "msyh.ttc"
    if system_font.exists():
        candidates.append((system_font, "系统微软雅黑"))
    return candidates


def _missing_glyphs(font: fitz.Font, text: str) -> list[str]:
    missing: list[str] = []
    for char in text:
        if char.isspace():
            continue
        if not (_CJK_RE.match(char) or char.isascii()):
            continue
        if not font.has_glyph(ord(char)) and char not in missing:
            missing.append(char)
    return missing


def _resolve_font(segments: list[Segment]) -> tuple[Path, str, fitz.Font]:
    text = "".join(segment.translated for segment in segments)
    candidates = _font_candidates()
    if not candidates:
        raise TranslateError(
            "未找到中文字体：请在 assets/fonts/ 放入 Noto Sans SC 字体，或安装微软雅黑后重试"
        )
    missing: list[str] = []
    for path, label in candidates:
        try:
            font = fitz.Font(fontfile=str(path))
        except Exception:
            LOGGER.warning("无法加载字体：%s", path, exc_info=True)
            continue
        missing = _missing_glyphs(font, text)
        if not missing:
            return path, label, font
    detail = "、".join(missing[:8]) or "未知字符"
    raise TranslateError(
        f"译文包含当前字体无法显示的生僻字（{detail}），请安装覆盖这些字符的中文字体后重试"
    )


def _wrap_text(text: str, font: fitz.Font, size: float, width: float) -> list[str]:
    lines: list[str] = []
    for hard_line in (text or "").splitlines() or [""]:
        if not hard_line:
            lines.append("")
            continue
        current = ""
        for char in hard_line:
            candidate = f"{current}{char}"
            if current and font.text_length(candidate, fontsize=size) > width:
                lines.append(current)
                current = char
            else:
                current = candidate
        if current:
            lines.append(current)
    return lines


def _fits(text: str, font: fitz.Font, size: float, width: float, height: float) -> list[str] | None:
    lines = _wrap_text(text, font, size, max(1.0, width - 1.0))
    needed = len(lines) * size * LAYOUT_LINE_HEIGHT
    if needed <= max(1.0, height - 1.0):
        return lines
    return None


def _fit_text(
    text: str,
    font: fitz.Font,
    start_size: float,
    width: float,
    height: float,
) -> tuple[str, float] | None:
    size = max(LAYOUT_MIN_FONT_SIZE, min(18.0, float(start_size or 10.0)))
    while size >= LAYOUT_MIN_FONT_SIZE:
        if _fits(text, font, size, width, height):
            return text, size
        size -= LAYOUT_FONT_STEP
    clipped = text
    while clipped:
        candidate = f"{clipped}…"
        if _fits(candidate, font, LAYOUT_MIN_FONT_SIZE, width, height):
            return candidate, LAYOUT_MIN_FONT_SIZE
        clipped = clipped[:-1]
    return None


def _write_layout(
    src: Path,
    segments: list[Segment],
    dst: Path,
    target_lang: str,
    font_path: Path,
    font_label: str,
    font: fitz.Font,
) -> str:
    document = fitz.open(str(src))
    try:
        if document.needs_pass:
            raise TranslateError("不支持加密 PDF，请先解除密码保护")
        by_page: dict[int, list[Segment]] = {}
        for segment in segments:
            if segment.rect is not None:
                by_page.setdefault(segment.page, []).append(segment)

        for page_number, page_segments in sorted(by_page.items()):
            page = document[page_number - 1]
            for segment in page_segments:
                if segment.rect is None:
                    continue
                page.add_redact_annot(fitz.Rect(segment.rect), fill=(1, 1, 1))
            page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)
            page.insert_font(fontname=LAYOUT_FONT_ALIAS, fontfile=str(font_path))
            for segment in page_segments:
                if segment.rect is None:
                    continue
                rect = fitz.Rect(segment.rect)
                if rect.width < 4 or rect.height < 4:
                    continue
                placed = _fit_text(
                    segment.translated,
                    font,
                    segment.font_size,
                    rect.width,
                    rect.height,
                )
                if placed is None:
                    placed = ("…", LAYOUT_MIN_FONT_SIZE)
                text, size = placed
                page.insert_textbox(
                    rect,
                    text,
                    fontname=LAYOUT_FONT_ALIAS,
                    fontsize=size,
                    color=(0, 0, 0),
                    overlay=True,
                )
        try:
            document.subset_fonts()
        except Exception:
            LOGGER.warning("子集化嵌入字体失败，保留完整字体", exc_info=True)
        document.save(str(dst), garbage=3, deflate=True)
    finally:
        document.close()
    return f"已翻译 {len(segments)} 段并覆盖原文字块；嵌入字体：{font_label}"


# ---------------------------------------------------------------------------
# 入口
# ---------------------------------------------------------------------------

def translate_document(
    src: Path,
    display: str,
    output_dir: Path,
    mode: str,
    target_lang: str,
    ui_glossary: str,
    progress: ProgressFn,
) -> tuple[Path, str]:
    if mode not in {"bilingual", "translated", "layout"}:
        raise TranslateError("输出方式无效，只支持 bilingual / translated / layout")
    if target_lang not in {"zh", "en"}:
        raise TranslateError("目标语言无效，只支持 zh / en")
    if mode == "layout" and src.suffix.lower() != ".pdf":
        raise TranslateError("保留版式输出只支持 PDF，请先用格式转换转成 PDF")
    if not is_configured():
        raise TranslateError("未配置翻译密钥，请在 .env 里填写 TRANSLATE_API_KEY")

    output_dir.mkdir(parents=True, exist_ok=True)
    LOGGER.info(
        "开始翻译 | %s | mode=%s | %s → %s | key=%s | model=%s",
        display,
        mode,
        src.suffix.lower().lstrip("."),
        target_lang,
        mask_key(api_key()),
        model_name(),
    )
    progress(2, "正在提取段落")
    segments = extract_segments(src)
    glossary = merged_glossary(ui_glossary)
    progress(5, f"已提取 {len(segments)} 段，正在调用翻译接口")
    _translate_segments(segments, target_lang, glossary, progress)
    progress(96, "正在生成成品文件")

    stem = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", src.stem).strip() or "译文"
    term_note = f"，术语 {len(glossary)} 条" if glossary else ""
    if mode == "bilingual":
        dst = output_dir / f"{stem}_中英对照.docx"
        _write_bilingual(segments, dst, target_lang)
        note = f"已翻译 {len(segments)} 段，左右两列逐段对应{term_note}"
    elif mode == "translated":
        dst = output_dir / f"{stem}_译文.docx"
        _write_translated(segments, dst, target_lang)
        note = f"已翻译 {len(segments)} 段，标题层级尽量沿用原文档{term_note}"
    else:
        font_path, font_label, font = _resolve_font(segments)
        dst = output_dir / f"{stem}_版式译文.pdf"
        note = _write_layout(src, segments, dst, target_lang, font_path, font_label, font)
        note = f"{note}{term_note}"
    progress(99, "译文已生成")
    return dst, note
