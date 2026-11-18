"""端到端冒烟测试：自起服务，真实上传、转换、翻译、扫描增强、轮询、下载。"""

from __future__ import annotations

import base64
import io
import json
import os
import socket
import subprocess
import sys
import tempfile
import threading
import time
import urllib.error
import urllib.request
import uuid
import zipfile
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

import fitz
from PIL import Image, ImageDraw
from docx import Document
from openpyxl import Workbook, load_workbook

ROOT = Path(__file__).resolve().parents[1]
PYTHON = Path(sys.executable)
for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        _stream.reconfigure(encoding="utf-8", errors="replace")
PASSED: list[str] = []
FAILED: list[str] = []
SERVER_LOG: Path | None = None
BASE_URL = ""


class MockChatHandler(BaseHTTPRequestHandler):
    """本地假模型接口：按 <<<SEG>>> 数量返回等量译文。"""

    last_system = ""
    last_user = ""
    calls = 0
    lock = threading.Lock()

    def do_POST(self) -> None:
        length = int(self.headers.get("Content-Length", "0"))
        try:
            payload = json.loads(self.rfile.read(length).decode("utf-8"))
        except Exception:
            self.send_error(400)
            return
        messages = payload.get("messages") or []
        system = "\n".join(
            str(item.get("content", "")) for item in messages if item.get("role") == "system"
        )
        user = "\n".join(
            str(item.get("content", "")) for item in messages if item.get("role") == "user"
        )
        parts = user.split("<<<SEG>>>")
        translated = [f"译文{index + 1}" for index in range(len(parts))]
        with self.lock:
            call_index = type(self).calls
            type(self).calls += 1
            type(self).last_system = system
            type(self).last_user = user
        # 让第一次请求明显更慢，保证轮询能观察到块级进度
        time.sleep(2.0 if call_index == 0 else 0.3)
        response = json.dumps(
            {
                "id": "mock",
                "object": "chat.completion",
                "choices": [
                    {"index": 0, "message": {"role": "assistant", "content": "<<<SEG>>>".join(translated)}}
                ],
                "usage": {"prompt_tokens": 10, "completion_tokens": 10, "total_tokens": 20},
            },
            ensure_ascii=False,
        ).encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(response)))
        self.end_headers()
        self.wfile.write(response)

    def log_message(self, format: str, *args) -> None:
        return


def start_mock_chat() -> tuple[ThreadingHTTPServer, int]:
    server = ThreadingHTTPServer(("127.0.0.1", 0), MockChatHandler)
    server.daemon_threads = True
    thread = threading.Thread(target=server.serve_forever, name="mock-chat", daemon=True)
    thread.start()
    return server, int(server.server_address[1])


def pick_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind(("127.0.0.1", 0))
        return int(sock.getsockname()[1])


def http_request(
    method: str,
    path: str,
    data: bytes | None = None,
    headers: dict[str, str] | None = None,
    timeout: int = 120,
) -> tuple[int, dict[str, str], bytes]:
    request = urllib.request.Request(
        BASE_URL + path,
        data=data,
        method=method,
        headers=headers or {},
    )
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            return response.status, {k.lower(): v for k, v in response.headers.items()}, response.read()
    except urllib.error.HTTPError as exc:
        return exc.code, {k.lower(): v for k, v in exc.headers.items()}, exc.read()


def json_request(method: str, path: str, timeout: int = 30) -> tuple[int, dict, bytes]:
    status, headers, body = http_request(method, path, timeout=timeout)
    try:
        payload = json.loads(body.decode("utf-8"))
    except Exception:
        payload = {}
    return status, payload, body


def multipart_batch(
    fields: dict[str, str],
    files: list[tuple[str, bytes, str]],
    timeout: int = 120,
) -> tuple[int, dict, bytes]:
    boundary = "----smokebatch" + uuid.uuid4().hex
    chunks: list[bytes] = []
    for name, value in fields.items():
        chunks.append(
            f'--{boundary}\r\nContent-Disposition: form-data; name="{name}"\r\n\r\n{value}\r\n'.encode(
                "utf-8"
            )
        )
    for filename, content, content_type in files:
        chunks.append(
            (
                f'--{boundary}\r\nContent-Disposition: form-data; name="files"; '
                f'filename="{filename}"\r\nContent-Type: {content_type}\r\n\r\n'
            ).encode("utf-8")
        )
        chunks.append(content)
        chunks.append(b"\r\n")
    chunks.append(f"--{boundary}--\r\n".encode("utf-8"))
    body = b"".join(chunks)
    return http_request(
        "POST",
        "/api/batch",
        data=body,
        headers={"Content-Type": f"multipart/form-data; boundary={boundary}", "Content-Length": str(len(body))},
        timeout=timeout,
    )


def poll_task_tracked(task_id: str, timeout_seconds: int = 120) -> tuple[dict, list[str], list[int]]:
    """轮询直到结束，同时收集出现过的进度文案与百分比。"""
    deadline = time.time() + timeout_seconds
    messages: list[str] = []
    percents: list[int] = []
    while time.time() < deadline:
        status, payload, _ = json_request("GET", f"/api/tasks/{task_id}")
        if status != 200:
            raise AssertionError(f"轮询失败 HTTP {status}: {payload}")
        message = str(payload.get("message") or "")
        if message and (not messages or messages[-1] != message):
            messages.append(message)
        percent = int(payload.get("progress") or 0)
        percents.append(percent)
        if payload["status"] == "done":
            return payload, messages, percents
        if payload["status"] == "error":
            raise AssertionError(f"任务失败：{payload.get('error')}")
        time.sleep(0.15)
    raise AssertionError("任务在 120 秒内没有完成")


def submit_translate(
    filename: str,
    content: bytes,
    mode: str,
    lang: str = "zh",
    glossary: str = "",
    content_type: str = "application/pdf",
) -> str:
    body, boundary = multipart_named(
        {"target_lang": lang, "mode": mode, "glossary": glossary},
        [(filename, content, content_type)],
        field_name="file",
    )
    status, _, response = http_request(
        "POST",
        "/api/translate",
        data=body,
        headers={"Content-Type": boundary, "Content-Length": str(len(body))},
    )
    if status != 200:
        raise AssertionError(f"翻译提交失败 HTTP {status}: {response.decode('utf-8', 'replace')}")
    return json.loads(response.decode("utf-8"))["task_id"]


def submit_scan(filename: str, content: bytes, content_type: str = "image/jpeg") -> str:
    body, boundary = multipart_named({}, [(filename, content, content_type)], field_name="file")
    status, _, response = http_request(
        "POST",
        "/api/scan-photo",
        data=body,
        headers={"Content-Type": boundary, "Content-Length": str(len(body))},
    )
    if status != 200:
        raise AssertionError(f"扫描件提交失败 HTTP {status}: {response.decode('utf-8', 'replace')}")
    return json.loads(response.decode("utf-8"))["task_id"]


def poll_task(task_id: str, timeout_seconds: int = 120) -> dict:
    deadline = time.time() + timeout_seconds
    while time.time() < deadline:
        status, payload, _ = json_request("GET", f"/api/tasks/{task_id}")
        if status != 200:
            raise AssertionError(f"轮询失败 HTTP {status}: {payload}")
        if payload["status"] == "done":
            return payload
        if payload["status"] == "error":
            raise AssertionError(f"任务失败：{payload.get('error')}")
        time.sleep(0.2)
    raise AssertionError("任务在 120 秒内没有完成")


def submit_batch_and_wait(
    files: list[tuple[str, bytes, str]],
    target: str,
    mode: str = "convert",
) -> tuple[dict, dict, bytes]:
    status, headers, body = multipart_batch({"mode": mode, "target": target}, files)
    if status != 200:
        raise AssertionError(f"批量提交失败 HTTP {status}: {body.decode('utf-8', 'replace')}")
    task = json.loads(body.decode("utf-8"))
    task_id = task["task_id"]
    payload = poll_task(task_id)
    download_status, download_headers, zip_bytes = http_request("GET", f"/api/download/{task_id}")
    if download_status != 200:
        raise AssertionError(f"批量 ZIP 下载失败 HTTP {download_status}")
    return payload, download_headers, zip_bytes


def submit_and_expect_error(filename: str, content: bytes, target: str) -> dict:
    body, content_type = multipart(
        {"target": target},
        {"file": (filename, content, "application/octet-stream")},
    )
    status, _, response = http_request(
        "POST",
        "/api/convert",
        data=body,
        headers={"Content-Type": content_type, "Content-Length": str(len(body))},
    )
    if status != 200:
        raise AssertionError(f"提交失败 HTTP {status}: {response.decode('utf-8', 'replace')}")
    task_id = json.loads(response.decode("utf-8"))["task_id"]
    deadline = time.time() + 120
    while time.time() < deadline:
        status, payload, _ = json_request("GET", f"/api/tasks/{task_id}")
        if status != 200:
            raise AssertionError(f"轮询失败 HTTP {status}: {payload}")
        if payload["status"] == "error":
            return payload
        if payload["status"] == "done":
            raise AssertionError("任务本应失败，但返回了成功")
        time.sleep(0.2)
    raise AssertionError("错误任务在 120 秒内没有结束")


def multipart(fields: dict[str, str], files: dict[str, tuple[str, bytes, str]]) -> tuple[bytes, str]:
    boundary = "----smoke" + uuid.uuid4().hex
    chunks: list[bytes] = []
    for name, value in fields.items():
        chunks.append(
            f'--{boundary}\r\nContent-Disposition: form-data; name="{name}"\r\n\r\n{value}\r\n'.encode(
                "utf-8"
            )
        )
    for name, (filename, content, content_type) in files.items():
        chunks.append(
            (
                f'--{boundary}\r\nContent-Disposition: form-data; name="{name}"; '
                f'filename="{filename}"\r\nContent-Type: {content_type}\r\n\r\n'
            ).encode("utf-8")
        )
        chunks.append(content)
        chunks.append(b"\r\n")
    chunks.append(f"--{boundary}--\r\n".encode("utf-8"))
    return b"".join(chunks), f"multipart/form-data; boundary={boundary}"


def submit_and_wait(filename: str, content: bytes, target: str) -> tuple[dict, dict, bytes]:
    body, content_type = multipart(
        {"target": target},
        {"file": (filename, content, "application/octet-stream")},
    )
    status, headers, response = http_request(
        "POST",
        "/api/convert",
        data=body,
        headers={"Content-Type": content_type, "Content-Length": str(len(body))},
    )
    if status != 200:
        raise AssertionError(f"提交失败 HTTP {status}: {response.decode('utf-8', 'replace')}")
    task = json.loads(response.decode("utf-8"))
    task_id = task["task_id"]
    payload = poll_task(task_id)
    result = payload["result"]
    download_status, headers, content_bytes = http_request("GET", result["download_url"])
    if download_status != 200:
        raise AssertionError(f"下载失败 HTTP {download_status}")
    return payload, result, content_bytes


def multipart_named(
    fields: dict[str, str],
    files: list[tuple[str, bytes, str]],
    field_name: str = "files",
) -> tuple[bytes, str]:
    boundary = "----smokenamed" + uuid.uuid4().hex
    chunks: list[bytes] = []
    for name, value in fields.items():
        chunks.append(
            f'--{boundary}\r\nContent-Disposition: form-data; name="{name}"\r\n\r\n{value}\r\n'.encode(
                "utf-8"
            )
        )
    for filename, content, content_type in files:
        chunks.append(
            (
                f'--{boundary}\r\nContent-Disposition: form-data; name="{field_name}"; '
                f'filename="{filename}"\r\nContent-Type: {content_type}\r\n\r\n'
            ).encode("utf-8")
        )
        chunks.append(content)
        chunks.append(b"\r\n")
    chunks.append(f"--{boundary}--\r\n".encode("utf-8"))
    return b"".join(chunks), f"multipart/form-data; boundary={boundary}"


def post_task(
    path: str,
    fields: dict[str, str],
    files: list[tuple[str, bytes, str]],
    field_name: str = "files",
) -> str:
    body, content_type = multipart_named(fields, files, field_name)
    status, _, response = http_request(
        "POST",
        path,
        data=body,
        headers={"Content-Type": content_type, "Content-Length": str(len(body))},
    )
    if status != 200:
        raise AssertionError(f"{path} 提交失败 HTTP {status}: {response.decode('utf-8', 'replace')}")
    return json.loads(response.decode("utf-8"))["task_id"]


def wait_task_error(task_id: str, timeout: int = 120) -> dict:
    deadline = time.time() + timeout
    while time.time() < deadline:
        status, payload, _ = json_request("GET", f"/api/tasks/{task_id}")
        if status != 200:
            raise AssertionError(f"轮询失败 HTTP {status}: {payload}")
        if payload["status"] == "error":
            return payload
        if payload["status"] == "done":
            raise AssertionError("任务本应失败，但返回了成功")
        time.sleep(0.2)
    raise AssertionError("错误任务在 120 秒内没有结束")


def make_pdf(path: Path, texts: list[str], rotation: int = 0) -> None:
    document = fitz.open()
    try:
        for text in texts:
            page = document.new_page()
            page.insert_text((72, 72), text, fontsize=14)
        if rotation:
            document[0].set_rotation(rotation)
        document.save(str(path))
    finally:
        document.close()


def check(name: str, condition: bool, detail: str = "") -> None:
    if condition:
        PASSED.append(name)
        print(f"PASS  {name}")
    else:
        FAILED.append(name)
        print(f"FAIL  {name}  {detail}")


def run(name: str, func) -> None:
    try:
        func()
        check(name, True)
    except Exception as exc:
        FAILED.append(name)
        print(f"FAIL  {name}  {type(exc).__name__}: {exc}")


def make_samples(directory: Path) -> dict[str, Path]:
    docx_path = directory / "烟雾测试.docx"
    document = Document()
    document.add_heading("烟雾测试文档", level=1)
    document.add_paragraph("这是一个用于验证 DOCX 转 TXT 和 DOCX 转 PDF 的中文段落。")
    document.add_paragraph("第二段包含 English words 和数字 12345。")
    table = document.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "项目"
    table.cell(0, 1).text = "状态"
    table.cell(1, 0).text = "转换"
    table.cell(1, 1).text = "通过"
    document.save(str(docx_path))

    pdf_path = directory / "smoke.pdf"
    pdf = fitz.open()
    for text in ("PDF smoke page one", "PDF smoke page two"):
        page = pdf.new_page()
        page.insert_text((72, 72), text, fontsize=14)
    pdf.save(str(pdf_path))
    pdf.close()

    rotated_path = directory / "rotated.pdf"
    make_pdf(rotated_path, ["rotated page one", "rotated page two"], rotation=90)

    encrypted_path = directory / "encrypted.pdf"
    encrypted = fitz.open()
    try:
        encrypted.new_page().insert_text((72, 72), "secret", fontsize=14)
        encrypted.save(
            str(encrypted_path),
            encryption=fitz.PDF_ENCRYPT_AES_256,
            owner_pw="owner",
            user_pw="secret",
        )
    finally:
        encrypted.close()

    letters_a = directory / "letters_a.pdf"
    make_pdf(letters_a, ["A1", "A2"])
    letters_b = directory / "letters_b.pdf"
    make_pdf(letters_b, ["B1", "B2"])
    pages3_path = directory / "pages3.pdf"
    make_pdf(pages3_path, ["P1", "P2", "P3"])

    noise_path = directory / "noise.png"
    noise = Image.effect_noise((900, 700), 120).convert("RGB")
    noise.save(str(noise_path), "PNG")
    noise.close()
    noisy_pdf_path = directory / "noisy.pdf"
    noisy = fitz.open()
    try:
        page = noisy.new_page()
        page.insert_image(page.rect, filename=str(noise_path))
        noisy.save(str(noisy_pdf_path))
    finally:
        noisy.close()

    image_path = directory / "smoke.png"
    image = Image.new("RGB", (480, 320), "white")
    draw = ImageDraw.Draw(image)
    draw.rectangle((40, 40, 440, 280), outline="#2f63f4", width=5)
    draw.ellipse((190, 110, 290, 210), fill="#2f63f4")
    image.save(str(image_path), "PNG")
    image.close()

    english_path = directory / "english.pdf"
    english = fitz.open()
    try:
        for page_index in range(3):
            page = english.new_page()
            for row in range(6):
                text = (
                    f"Section {page_index + 1}.{row + 1}. This study evaluates document conversion "
                    "workflows for multilingual academic papers. The proposed pipeline preserves "
                    "paragraph order, heading levels and table structure while keeping every file "
                    "on the local machine. Results show stable output across repeated runs."
                )
                page.insert_textbox(fitz.Rect(56, 60 + row * 118, 539, 168 + row * 118), text, fontsize=10)
        english.save(str(english_path))
    finally:
        english.close()

    scan_path = directory / "photo.jpg"
    width, height = 1000, 800
    canvas = Image.new("RGB", (width, height), (52, 55, 62))
    canvas_draw = ImageDraw.Draw(canvas)
    paper = [(230, 120), (850, 180), (800, 700), (180, 640)]
    canvas_draw.polygon(paper, fill=(248, 246, 238))
    for line in range(9):
        y = 220 + line * 48
        canvas_draw.line((280, y, 740, y + 8), fill=(48, 48, 52), width=3)
    gradient = Image.linear_gradient("L").resize((width, height))
    shade = Image.new("RGB", (width, height), (18, 18, 24))
    alpha = gradient.point(lambda value: int(value * 0.38))
    canvas = Image.composite(shade, canvas, alpha)
    grain = Image.effect_noise((width, height), 28).convert("RGB")
    canvas = Image.blend(canvas, grain, 0.05)
    canvas.save(str(scan_path), "JPEG", quality=88)
    canvas.close()

    md_path = directory / "smoke.md"
    md_path.write_text("# 烟雾测试标题\n\n这是 **Markdown** 正文。\n\n| A | B |\n|---|---|\n| 1 | 2 |\n", encoding="utf-8")

    txt_path = directory / "smoke.txt"
    txt_path.write_text("第一行中文\n第二行 smoke text\n\n第三段\n", encoding="utf-8")

    exe_path = directory / "bad.exe"
    exe_path.write_bytes(b"MZ" + b"\x00" * 128)

    csv_path = directory / "smoke.csv"
    csv_path.write_text("姓名,分数\n张三,95\n李四,88\n", encoding="utf-8-sig")

    table_path = directory / "table.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "数据"
    sheet.append(["姓名", "分数"])
    sheet.append(["张三", 95])
    sheet.append(["李四", 88])
    workbook.save(str(table_path))

    multi_path = directory / "sheets.xlsx"
    workbook = Workbook()
    first = workbook.active
    first.title = "第一页"
    first.append(["项目", "状态"])
    first.append(["转换", "通过"])
    second = workbook.create_sheet("第二页")
    second.append(["名称", "数量"])
    second.append(["测试", 3])
    workbook.save(str(multi_path))

    broken_docx = directory / "broken.docx"
    broken_docx.write_bytes(b"PK\x03\x04" + b"broken-package" * 20)

    fake_doc = directory / "fake.doc"
    fake_doc.write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 512)

    fake_pptx = directory / "fake.pptx"
    fake_pptx.write_bytes(b"PK\x03\x04" + b"\x00" * 512)

    xls_source = ROOT / "samples" / "示例_旧版表格.xls"
    if not xls_source.exists():
        raise AssertionError(f"缺少旧版 XLS 样例：{xls_source}")

    return {
        "docx": docx_path,
        "pdf": pdf_path,
        "rotated": rotated_path,
        "encrypted": encrypted_path,
        "letters_a": letters_a,
        "letters_b": letters_b,
        "pages3": pages3_path,
        "noisy": noisy_pdf_path,
        "png": image_path,
        "english": english_path,
        "photo": scan_path,
        "md": md_path,
        "txt": txt_path,
        "exe": exe_path,
        "csv": csv_path,
        "table": table_path,
        "multi": multi_path,
        "broken_docx": broken_docx,
        "fake_doc": fake_doc,
        "fake_pptx": fake_pptx,
        "xls": xls_source,
    }


def main() -> int:
    global BASE_URL, SERVER_LOG
    port = pick_port()
    BASE_URL = f"http://127.0.0.1:{port}"
    temp_dir = Path(tempfile.mkdtemp(prefix="docworkshop-smoke-"))
    SERVER_LOG = temp_dir / "server.log"
    samples = make_samples(temp_dir)
    log_handle = SERVER_LOG.open("wb")
    mock_server, mock_port = start_mock_chat()

    env = os.environ.copy()
    env.update(
        {
            "PORT": str(port),
            "OFFICE_ENABLED": "0",
            "AUTO_OPEN_BROWSER": "0",
            "PYTHONUNBUFFERED": "1",
            "TASK_TTL_SECONDS": "3600",
            "CLEANUP_INTERVAL_SECONDS": "60",
            "TRANSLATE_API_KEY": "sk-smoke-mock-key",
            "TRANSLATE_BASE_URL": f"http://127.0.0.1:{mock_port}/v1",
            "TRANSLATE_MODEL": "mock-model",
        }
    )
    creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    process = subprocess.Popen(
        [str(PYTHON), str(ROOT / "main.py")],
        cwd=str(ROOT),
        env=env,
        stdout=log_handle,
        stderr=subprocess.STDOUT,
        creationflags=creationflags,
    )

    try:
        deadline = time.time() + 25
        health = None
        while time.time() < deadline:
            if process.poll() is not None:
                break
            try:
                status, payload, _ = json_request("GET", "/api/health", timeout=3)
                if status == 200 and payload.get("status") == "ok":
                    health = payload
                    break
            except Exception:
                pass
            time.sleep(0.25)
        if health is None:
            log_handle.flush()
            print("服务没有启动成功，日志：")
            print(SERVER_LOG.read_text(encoding="utf-8", errors="replace")[-4000:])
            return 1

        def test_health():
            status, payload, _ = json_request("GET", "/api/health")
            assert status == 200 and payload.get("status") == "ok"
            assert payload.get("office") is False
            assert payload.get("excel") is False
            assert payload.get("powerpoint") is False
            assert payload.get("version") == "0.4.0"
            head_status, _, _ = http_request("HEAD", "/api/health")
            assert head_status == 200

        run("health 与 HEAD 健康检查", test_health)

        def test_config():
            status, payload, _ = json_request("GET", "/api/config")
            assert status == 200
            assert payload["max_file_size_mb"] == 200
            assert payload["max_batch_files"] == 200
            assert payload["max_batch_size_mb"] == 500
            assert payload["translate_configured"] is True
            assert "大语言模型" in payload["glossary"]
            assert "英文=中文" in payload["glossary"]

        run("配置接口含翻译状态与术语表", test_config)

        def test_formats():
            status, payload, _ = json_request("GET", "/api/formats")
            assert status == 200
            assert set(payload["docx"]) == {"html", "md", "pdf", "txt"}
            assert set(payload["pdf"]) == {"docx", "jpg", "png", "txt"}
            assert set(payload["doc"]) == {"docx", "pdf"}
            assert set(payload["xlsx"]) == {"csv", "pdf"}
            assert set(payload["xls"]) == {"csv", "xlsx"}
            assert set(payload["pptx"]) == {"pdf", "png"}
            assert set(payload["csv"]) == {"xlsx"}
            assert set(payload["md"]) == {"html", "pdf"}
            assert set(payload["txt"]) == {"docx", "pdf"}
            assert {"pdf", "docx", "jpg", "webp", "bmp", "tiff"} <= set(payload["png"])

        run("格式表由注册表反推", test_formats)

        def test_homepage():
            status, headers, body = http_request("GET", "/")
            text = body.decode("utf-8")
            assert status == 200
            assert "格式转换" in text and "拖拽文件" in text
            assert "批量处理" in text and "batch-dropzone" in text
            assert "compress-confirm-dialog" in text and "pdf-grid" in text
            assert "merge-mode" in text and "pdf-jpg-btn" in text
            assert "translate-dropzone" in text and "translate-modes" in text
            assert "translate-config-warning" in text and "scan-dropzone" in text
            assert "scan-compare" in text and "scan-result" in text
            assert "app.js" in text and "app.css" in text
            status, _, css = http_request("GET", "/static/app.css")
            assert status == 200 and len(css) > 1000
            status, _, js = http_request("GET", "/static/app.js")
            assert status == 200 and len(js) > 1000
            script = js.decode("utf-8")
            assert "askCompressConfirm" in script and "window.confirm" not in script

        run("首页与静态资源", test_homepage)

        def test_docx_txt():
            docx = samples["docx"].read_bytes()
            _, result, content = submit_and_wait("烟雾测试.docx", docx, "txt")
            assert result["filename"] == "烟雾测试.txt"
            assert content.startswith(b"\xef\xbb\xbf"), "TXT 应为 UTF-8 BOM"
            text = content.decode("utf-8-sig")
            assert "烟雾测试文档" in text
            assert "English words" in text
            assert "项目\t状态" in text
            assert "内置排版" not in result.get("note", "")

        run("DOCX → TXT（中文文件名 + 表格）", test_docx_txt)

        def test_pdf_txt():
            _, result, content = submit_and_wait("smoke.pdf", samples["pdf"].read_bytes(), "txt")
            text = content.decode("utf-8-sig")
            assert "PDF smoke page one" in text
            assert "PDF smoke page two" in text
            assert result["note"]

        run("PDF → TXT 两页文字", test_pdf_txt)

        def test_pdf_docx():
            _, result, content = submit_and_wait("smoke.pdf", samples["pdf"].read_bytes(), "docx")
            with zipfile.ZipFile(io.BytesIO(content)) as archive:
                names = archive.namelist()
                assert "word/document.xml" in names
                document_xml = archive.read("word/document.xml").decode("utf-8", "replace")
            assert "PDF smoke page one" in document_xml
            assert result["filename"] == "smoke.docx"

        run("PDF → DOCX", test_pdf_docx)

        def test_pdf_jpg():
            _, result, content = submit_and_wait("smoke.pdf", samples["pdf"].read_bytes(), "jpg")
            assert result["filename"] == "smoke_jpg.zip"
            with zipfile.ZipFile(io.BytesIO(content)) as archive:
                names = sorted(archive.namelist())
                assert names == ["001.jpg", "002.jpg"]
                for name in names:
                    data = archive.read(name)
                    assert data.startswith(b"\xff\xd8\xff")
                    assert len(data) > 1000

        run("PDF → JPG（两页 ZIP）", test_pdf_jpg)

        def test_png_pdf():
            _, result, content = submit_and_wait("smoke.png", samples["png"].read_bytes(), "pdf")
            assert content.startswith(b"%PDF")
            doc = fitz.open(stream=content, filetype="pdf")
            try:
                assert doc.page_count == 1
                assert doc[0].rect.width > 0 and doc[0].rect.height > 0
            finally:
                doc.close()
            assert result["note"]

        run("PNG → PDF", test_png_pdf)

        def test_md_html():
            _, result, content = submit_and_wait("smoke.md", samples["md"].read_bytes(), "html")
            text = content.decode("utf-8")
            assert "<h1" in text and "烟雾测试标题" in text
            assert "<table>" in text and "<td>1</td>" in text
            assert result["filename"] == "smoke.html"

        run("Markdown → HTML", test_md_html)

        def test_txt_docx():
            _, result, content = submit_and_wait("smoke.txt", samples["txt"].read_bytes(), "docx")
            document = Document(io.BytesIO(content))
            text = "\n".join(paragraph.text for paragraph in document.paragraphs)
            assert "第一行中文" in text
            assert "第二行 smoke text" in text
            assert result["filename"] == "smoke.docx"

        run("TXT → DOCX", test_txt_docx)

        def test_docx_pdf_fallback():
            _, result, content = submit_and_wait("烟雾测试.docx", samples["docx"].read_bytes(), "pdf")
            assert content.startswith(b"%PDF")
            assert "内置排版" in result["note"]
            assert result["filename"] == "烟雾测试.pdf"
            doc = fitz.open(stream=content, filetype="pdf")
            try:
                assert doc.page_count >= 1
                text = "".join(page.get_text() for page in doc)
                assert "烟雾测试文档" in text, f"PDF 文字层缺失：{text[:200]!r}"
            finally:
                doc.close()

        run("DOCX → PDF（无 Office 内置排版）", test_docx_pdf_fallback)

        def test_missing_download():
            status, payload, _ = json_request("GET", "/api/download/not-a-real-task")
            assert status == 404
            assert "error" in payload
            status, payload, _ = json_request("GET", "/api/tasks/not-a-real-task")
            assert status == 404 and "error" in payload

        run("不存在的任务与下载地址返回 404", test_missing_download)

        def test_exe_rejected():
            body, content_type = multipart(
                {"target": "pdf"},
                {"file": ("bad.exe", samples["exe"].read_bytes(), "application/octet-stream")},
            )
            status, _, response = http_request(
                "POST",
                "/api/convert",
                data=body,
                headers={"Content-Type": content_type, "Content-Length": str(len(body))},
            )
            assert status == 415
            payload = json.loads(response.decode("utf-8"))
            assert "error" in payload and "exe" in payload["error"].lower()

        run(".exe 上传返回 415", test_exe_rejected)

        def test_unsupported_pair():
            body, content_type = multipart(
                {"target": "txt"},
                {"file": ("smoke.md", samples["md"].read_bytes(), "text/markdown")},
            )
            status, _, response = http_request(
                "POST",
                "/api/convert",
                data=body,
                headers={"Content-Type": content_type, "Content-Length": str(len(body))},
            )
            assert status == 400
            payload = json.loads(response.decode("utf-8"))
            assert "error" in payload

        run("不支持的转换组合返回 400", test_unsupported_pair)

        def test_download_header():
            _, result, _ = submit_and_wait("烟雾测试.docx", samples["docx"].read_bytes(), "txt")
            status, headers, _ = http_request("GET", result["download_url"])
            disposition = headers.get("content-disposition", "")
            assert status == 200
            assert "filename*=UTF-8''" in disposition
            assert "%E7%83%9F" in disposition

        run("下载响应含中文文件名", test_download_header)

        def test_docx_html():
            _, result, content = submit_and_wait("烟雾测试.docx", samples["docx"].read_bytes(), "html")
            text = content.decode("utf-8")
            assert result["filename"] == "烟雾测试.html"
            assert "<h1" in text and "烟雾测试文档" in text
            assert "<table" in text and "项目" in text

        run("DOCX → HTML", test_docx_html)

        def test_docx_md():
            _, result, content = submit_and_wait("烟雾测试.docx", samples["docx"].read_bytes(), "md")
            text = content.decode("utf-8")
            assert result["filename"] == "烟雾测试.md"
            assert "烟雾测试文档" in text
            assert "| 项目 | 状态 |" in text
            assert "English words" in text

        run("DOCX → Markdown（标题、表格、正文）", test_docx_md)

        def test_xlsx_csv_single():
            _, result, content = submit_and_wait("table.xlsx", samples["table"].read_bytes(), "csv")
            assert result["filename"] == "table.csv"
            assert content.startswith(b"\xef\xbb\xbf"), "CSV 应为 UTF-8 BOM"
            text = content.decode("utf-8-sig")
            assert "姓名,分数" in text
            assert "张三,95" in text

        run("XLSX → CSV（单工作表 + BOM）", test_xlsx_csv_single)

        def test_xlsx_csv_multi():
            _, result, content = submit_and_wait("sheets.xlsx", samples["multi"].read_bytes(), "csv")
            assert result["filename"] == "sheets.zip"
            assert "2 个工作表" in result["note"]
            with zipfile.ZipFile(io.BytesIO(content)) as archive:
                names = sorted(archive.namelist())
                assert len(names) == 2
                assert all(name.startswith("sheets_") and name.endswith(".csv") for name in names)
                assert any("第一页" in name for name in names)
                for name in names:
                    data = archive.read(name)
                    assert data.startswith(b"\xef\xbb\xbf"), f"{name} 缺少 BOM"

        run("XLSX → CSV（多工作表 ZIP）", test_xlsx_csv_multi)

        def test_csv_xlsx():
            _, result, content = submit_and_wait("smoke.csv", samples["csv"].read_bytes(), "xlsx")
            workbook = load_workbook(io.BytesIO(content), data_only=True)
            try:
                sheet = workbook.active
                assert sheet["A1"].value == "姓名"
                # CSV 无类型信息，转换器按原文本写入，避免前导零被破坏
                assert sheet["B2"].value == "95"
            finally:
                workbook.close()
            assert result["filename"] == "smoke.xlsx"

        run("CSV → XLSX", test_csv_xlsx)

        def test_xls_to_xlsx():
            _, result, content = submit_and_wait(
                "示例_旧版表格.xls",
                samples["xls"].read_bytes(),
                "xlsx",
            )
            workbook = load_workbook(io.BytesIO(content), data_only=True)
            try:
                assert len(workbook.sheetnames) == 2
                assert workbook.sheetnames[0] == "汇总"
            finally:
                workbook.close()
            assert result["filename"] == "示例_旧版表格.xlsx"

        run("XLS → XLSX（xlrd）", test_xls_to_xlsx)

        def test_xls_to_csv():
            _, result, content = submit_and_wait(
                "示例_旧版表格.xls",
                samples["xls"].read_bytes(),
                "csv",
            )
            assert result["filename"] == "示例_旧版表格.zip"
            with zipfile.ZipFile(io.BytesIO(content)) as archive:
                names = archive.namelist()
                assert len(names) == 2
                for name in names:
                    assert name.startswith("示例_旧版表格_")
                    assert archive.read(name).startswith(b"\xef\xbb\xbf")

        run("XLS → CSV（多工作表 ZIP）", test_xls_to_csv)

        def test_md_pdf():
            _, result, content = submit_and_wait("smoke.md", samples["md"].read_bytes(), "pdf")
            assert content.startswith(b"%PDF")
            document = fitz.open(stream=content, filetype="pdf")
            try:
                assert document.page_count >= 1
            finally:
                document.close()
            assert result["filename"] == "smoke.pdf"

        run("Markdown → PDF", test_md_pdf)

        def test_txt_pdf():
            _, result, content = submit_and_wait("smoke.txt", samples["txt"].read_bytes(), "pdf")
            assert content.startswith(b"%PDF")
            document = fitz.open(stream=content, filetype="pdf")
            try:
                assert document.page_count >= 1
            finally:
                document.close()
            assert result["filename"] == "smoke.pdf"

        run("TXT → PDF", test_txt_pdf)

        def test_pdf_png():
            _, result, content = submit_and_wait("smoke.pdf", samples["pdf"].read_bytes(), "png")
            assert result["filename"] == "smoke_png.zip"
            with zipfile.ZipFile(io.BytesIO(content)) as archive:
                names = sorted(archive.namelist())
                assert names == ["001.png", "002.png"]
                for name in names:
                    assert archive.read(name).startswith(b"\x89PNG\r\n\x1a\n")

        run("PDF → PNG（页面 ZIP）", test_pdf_png)

        def test_png_jpg():
            _, result, content = submit_and_wait("smoke.png", samples["png"].read_bytes(), "jpg")
            assert result["filename"] == "smoke.jpg"
            assert content.startswith(b"\xff\xd8\xff")

        run("PNG → JPG", test_png_jpg)

        def test_png_webp():
            _, result, content = submit_and_wait("smoke.png", samples["png"].read_bytes(), "webp")
            assert result["filename"] == "smoke.webp"
            assert content[:4] == b"RIFF" and content[8:12] == b"WEBP"

        run("PNG → WebP", test_png_webp)

        def test_png_docx():
            _, result, content = submit_and_wait("smoke.png", samples["png"].read_bytes(), "docx")
            document = Document(io.BytesIO(content))
            assert len(document.inline_shapes) == 1
            assert result["filename"] == "smoke.docx"

        run("PNG → DOCX（图片嵌入）", test_png_docx)

        def test_office_disabled_errors():
            doc_task = submit_and_expect_error("fake.doc", samples["fake_doc"].read_bytes(), "pdf")
            assert "Word" in doc_task["error"] and "Office" in doc_task["error"]
            ppt_task = submit_and_expect_error("fake.pptx", samples["fake_pptx"].read_bytes(), "pdf")
            assert "PowerPoint" in ppt_task["error"] and "Office" in ppt_task["error"]

        run("Office 不可用时 .doc / .pptx 返回明确错误", test_office_disabled_errors)

        def test_batch_three_formats():
            files = [
                ("烟雾测试.docx", samples["docx"].read_bytes(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document"),
                ("smoke.md", samples["md"].read_bytes(), "text/markdown"),
                ("smoke.png", samples["png"].read_bytes(), "image/png"),
            ]
            payload, headers, content = submit_batch_and_wait(files, "pdf")
            result = payload["result"]
            assert result["total_files"] == 3
            assert result["success_files"] == 3
            assert result["failed_files"] == 0
            assert all(item["ok"] for item in result["results"])
            assert headers.get("content-type") == "application/zip"
            with zipfile.ZipFile(io.BytesIO(content)) as archive:
                names = archive.namelist()
                assert len(names) == 3
                assert all(name.endswith(".pdf") for name in names)
                for name in names:
                    assert archive.read(name).startswith(b"%PDF")

            first_url = result["results"][0]["download_url"]
            assert first_url.startswith(f"/api/download/{payload['task_id']}?file=")
            status, item_headers, item_content = http_request("GET", first_url)
            assert status == 200
            assert item_headers.get("content-type") == "application/pdf"
            assert item_content.startswith(b"%PDF")

        run("批量 3 种格式并行转换 + ZIP + 单文件下载", test_batch_three_formats)

        def test_batch_failure_isolation():
            files = [
                ("smoke.pdf", samples["pdf"].read_bytes(), "application/pdf"),
                ("broken.docx", samples["broken_docx"].read_bytes(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document"),
                ("bad.exe", samples["exe"].read_bytes(), "application/octet-stream"),
            ]
            payload, headers, content = submit_batch_and_wait(files, "txt")
            result = payload["result"]
            assert result["total_files"] == 3
            assert result["success_files"] == 1
            assert result["failed_files"] == 2
            assert result["results"][0]["ok"] is True
            assert result["results"][0]["download_url"]
            assert result["results"][1]["ok"] is False and result["results"][1]["error"]
            assert result["results"][2]["ok"] is False
            assert "exe" in result["results"][2]["error"].lower()
            assert headers.get("content-type") == "application/zip"
            with zipfile.ZipFile(io.BytesIO(content)) as archive:
                names = archive.namelist()
                assert len(names) == 1 and names[0].endswith(".txt")
                assert archive.read(names[0]).startswith(b"\xef\xbb\xbf")

        run("批量失败隔离（坏文件 + 不支持格式）", test_batch_failure_isolation)

        def test_batch_duplicate_names():
            files = [
                ("same.txt", "第一份".encode("utf-8"), "text/plain"),
                ("same.txt", "第二份".encode("utf-8"), "text/plain"),
            ]
            payload, _, content = submit_batch_and_wait(files, "docx")
            assert payload["result"]["success_files"] == 2
            with zipfile.ZipFile(io.BytesIO(content)) as archive:
                names = sorted(archive.namelist())
                assert names == ["same (2).docx", "same.docx"]

        run("批量 ZIP 内重名自动加序号", test_batch_duplicate_names)

        def test_batch_validation():
            status, _, body = multipart_batch(
                {"mode": "convert"},
                [("smoke.txt", samples["txt"].read_bytes(), "text/plain")],
            )
            assert status == 400 and "目标格式" in body.decode("utf-8")
            status, _, body = multipart_batch(
                {"mode": "compress", "target": "docx"},
                [("smoke.txt", samples["txt"].read_bytes(), "text/plain")],
            )
            assert status == 400 and "批量格式转换" in body.decode("utf-8")
            too_many = [(f"f{index}.txt", b"x", "text/plain") for index in range(201)]
            status, _, body = multipart_batch({"mode": "convert", "target": "docx"}, too_many)
            assert status == 413
            assert "200" in body.decode("utf-8")

        run("批量参数与数量上限校验", test_batch_validation)

        def test_pdf_preview():
            body, content_type = multipart(
                {"limit": "1"},
                {"file": ("smoke.pdf", samples["pdf"].read_bytes(), "application/pdf")},
            )
            status, _, response = http_request(
                "POST",
                "/api/pdf/preview",
                data=body,
                headers={"Content-Type": content_type, "Content-Length": str(len(body))},
            )
            assert status == 200
            data = json.loads(response.decode("utf-8"))
            assert len(data["pages"]) == 1
            assert data["truncated"] is True
            page = data["pages"][0]
            assert page["index"] == 1
            assert page["width"] == 595 and page["height"] == 842
            assert page["rotation"] == 0
            assert base64.b64decode(page["png_base64"]).startswith(b"\x89PNG\r\n\x1a\n")

        run("PDF 预览（前 N 页 + PNG base64）", test_pdf_preview)

        def test_pdf_preview_rotation():
            body, content_type = multipart(
                {"limit": "50"},
                {"file": ("rotated.pdf", samples["rotated"].read_bytes(), "application/pdf")},
            )
            status, _, response = http_request(
                "POST",
                "/api/pdf/preview",
                data=body,
                headers={"Content-Type": content_type, "Content-Length": str(len(body))},
            )
            assert status == 200
            data = json.loads(response.decode("utf-8"))
            assert data["truncated"] is False
            assert data["pages"][0]["rotation"] == 90
            assert data["pages"][0]["width"] == 595

        run("PDF 预览保留原 rotation", test_pdf_preview_rotation)

        def test_pdf_encrypted_rejected():
            body, content_type = multipart(
                {"limit": "10"},
                {"file": ("encrypted.pdf", samples["encrypted"].read_bytes(), "application/pdf")},
            )
            status, _, response = http_request(
                "POST",
                "/api/pdf/preview",
                data=body,
                headers={"Content-Type": content_type, "Content-Length": str(len(body))},
            )
            assert status == 400
            assert "不支持加密" in response.decode("utf-8")

        run("加密 PDF 返回 400", test_pdf_encrypted_rejected)

        def read_pdf(content: bytes):
            document = fitz.open(stream=content, filetype="pdf")
            try:
                return document.page_count, [page.get_text().strip() for page in document], [
                    page.rotation for page in document
                ]
            finally:
                document.close()

        def test_pdf_ops_delete_rotate():
            task_id = post_task(
                "/api/pdf/ops",
                {"delete_pages": "[2]", "rotations": '{"1": 90}'},
                [("smoke.pdf", samples["pdf"].read_bytes(), "application/pdf")],
                field_name="file",
            )
            payload = poll_task(task_id)
            assert "删除 1 页" in payload["result"]["note"]
            assert "旋转 1 页" in payload["result"]["note"]
            status, _, content = http_request("GET", payload["result"]["download_url"])
            assert status == 200
            count, texts, rotations = read_pdf(content)
            assert count == 1
            assert "PDF smoke page one" in texts[0]
            assert rotations == [90]

        run("PDF 删页 + 旋转（1-based 原页码）", test_pdf_ops_delete_rotate)

        def test_pdf_ops_preserve_rotation():
            task_id = post_task(
                "/api/pdf/ops",
                {"delete_pages": "[2]"},
                [("rotated.pdf", samples["rotated"].read_bytes(), "application/pdf")],
                field_name="file",
            )
            payload = poll_task(task_id)
            status, _, content = http_request("GET", payload["result"]["download_url"])
            assert status == 200
            count, _, rotations = read_pdf(content)
            assert count == 1
            assert rotations == [90], "未指定旋转时应保留原 rotation"

        run("PDF 删页时保留原 rotation", test_pdf_ops_preserve_rotation)

        def test_pdf_ops_page_order_duplicates():
            task_id = post_task(
                "/api/pdf/ops",
                {"page_order": "[3, 1, 1]"},
                [("pages3.pdf", samples["pages3"].read_bytes(), "application/pdf")],
                field_name="file",
            )
            payload = poll_task(task_id)
            status, _, content = http_request("GET", payload["result"]["download_url"])
            assert status == 200
            count, texts, _ = read_pdf(content)
            assert count == 3
            assert texts == ["P3", "P1", "P1"]

        run("PDF page_order 白名单允许重复页", test_pdf_ops_page_order_duplicates)

        def test_pdf_ops_errors():
            task_id = post_task(
                "/api/pdf/ops",
                {"delete_pages": "[99]"},
                [("pages3.pdf", samples["pages3"].read_bytes(), "application/pdf")],
                field_name="file",
            )
            payload = wait_task_error(task_id)
            assert "第 99 页" in payload["error"] and "3 页" in payload["error"]

            body, content_type = multipart(
                {},
                {"file": ("pages3.pdf", samples["pages3"].read_bytes(), "application/pdf")},
            )
            status, _, response = http_request(
                "POST",
                "/api/pdf/ops",
                data=body,
                headers={"Content-Type": content_type, "Content-Length": str(len(body))},
            )
            assert status == 400
            assert "至少指定" in response.decode("utf-8")

        run("PDF 页面操作越界与空参数报错", test_pdf_ops_errors)

        def test_pdf_merge_modes():
            files = [
                ("letters_a.pdf", samples["letters_a"].read_bytes(), "application/pdf"),
                ("letters_b.pdf", samples["letters_b"].read_bytes(), "application/pdf"),
            ]
            expected = {
                "append": ["A1", "A2", "B1", "B2"],
                "interleave": ["A1", "B1", "A2", "B2"],
                "insert": ["A1", "B1", "B2", "A2"],
            }
            for mode, order in expected.items():
                task_id = post_task(
                    "/api/pdf/merge",
                    {"mode": mode, "insert_index": "1"},
                    files,
                )
                payload = poll_task(task_id)
                assert "共合并 2 个文件、4 页" in payload["result"]["note"]
                status, _, content = http_request("GET", payload["result"]["download_url"])
                assert status == 200
                count, texts, _ = read_pdf(content)
                assert count == 4
                assert texts == order, f"{mode} 页序错误：{texts}"

        run("PDF 合并 append / interleave / insert", test_pdf_merge_modes)

        def test_compress_estimate():
            body, content_type = multipart(
                {},
                {"file": ("smoke.pdf", samples["pdf"].read_bytes(), "application/pdf")},
            )
            status, _, response = http_request(
                "POST",
                "/api/compress/estimate",
                data=body,
                headers={"Content-Type": content_type, "Content-Length": str(len(body))},
            )
            assert status == 200
            data = json.loads(response.decode("utf-8"))
            assert data["original_size"] == samples["pdf"].stat().st_size
            assert set(data["levels"]) == {"80", "60", "40", "20"}
            assert data["sampled_pages"] == 2

        run("PDF 压缩四档预估", test_compress_estimate)

        def test_compress_pdf():
            task_id = post_task(
                "/api/compress",
                {"level": "20"},
                [("smoke.pdf", samples["pdf"].read_bytes(), "application/pdf")],
                field_name="file",
            )
            payload = poll_task(task_id)
            result = payload["result"]
            assert "不可再搜索" in result["note"]
            status, _, content = http_request("GET", result["download_url"])
            assert status == 200 and content.startswith(b"%PDF")
            count, texts, _ = read_pdf(content)
            assert count == 2
            assert all(not text for text in texts), "压缩后不应保留文字层"

        run("PDF 压缩（文字层丢失提示）", test_compress_pdf)

        def test_compress_image():
            task_id = post_task(
                "/api/compress",
                {"level": "20"},
                [("smoke.png", samples["png"].read_bytes(), "image/png")],
                field_name="file",
            )
            payload = poll_task(task_id)
            status, _, content = http_request("GET", payload["result"]["download_url"])
            assert status == 200 and content.startswith(b"\x89PNG")
            assert len(content) < len(samples["png"].read_bytes())
            assert payload["result"]["filename"] == "smoke_压缩.png"

        run("图片压缩（20 档）", test_compress_image)

        def test_compress_unsupported():
            body, content_type = multipart(
                {"level": "60"},
                {"file": ("smoke.md", samples["md"].read_bytes(), "text/markdown")},
            )
            status, _, response = http_request(
                "POST",
                "/api/compress",
                data=body,
                headers={"Content-Type": content_type, "Content-Length": str(len(body))},
            )
            assert status == 400
            payload = json.loads(response.decode("utf-8"))
            assert payload["error"] == "该格式不支持压缩，本版只支持 PDF 和图片"

        run("非 PDF/图片压缩返回 400", test_compress_unsupported)

        def test_pdf_to_jpg_limit():
            task_id = post_task(
                "/api/pdf/to-jpg",
                {"max_kb": "300"},
                [("smoke.pdf", samples["pdf"].read_bytes(), "application/pdf")],
            )
            payload = poll_task(task_id)
            result = payload["result"]
            assert result["success_files"] == 1 and result["failed_files"] == 0
            assert "每张不超过 300KB" in result["note"]
            status, _, content = http_request("GET", result["download_url"])
            assert status == 200
            with zipfile.ZipFile(io.BytesIO(content)) as archive:
                names = [name for name in archive.namelist() if name.endswith(".jpg")]
                assert len(names) == 2
                for name in names:
                    data = archive.read(name)
                    assert data.startswith(b"\xff\xd8\xff")
                    assert len(data) <= 300 * 1024

        run("PDF 逐页 JPG ≤300KB", test_pdf_to_jpg_limit)

        def test_pdf_to_jpg_too_small_limit():
            task_id = post_task(
                "/api/pdf/to-jpg",
                {"max_kb": "10"},
                [("noisy.pdf", samples["noisy"].read_bytes(), "application/pdf")],
            )
            payload = wait_task_error(task_id)
            assert "无法压到 10KB 以内" in payload["error"]

        run("JPG 上限无法达成时明确报错", test_pdf_to_jpg_too_small_limit)

        def test_translate_bilingual():
            task_id = submit_translate(
                "english.pdf",
                samples["english"].read_bytes(),
                "bilingual",
                glossary="large language model=大语言模型\n",
            )
            payload, messages, percents = poll_task_tracked(task_id)
            result = payload["result"]
            assert result["filename"] == "english_中英对照.docx"
            assert "术语" in result["note"]
            assert any("正在翻译第" in item and "/" in item for item in messages), messages
            assert max(percents) >= 40
            status, _, content = http_request("GET", result["download_url"])
            assert status == 200
            with zipfile.ZipFile(io.BytesIO(content)) as archive:
                document_xml = archive.read("word/document.xml").decode("utf-8", "replace")
            assert document_xml.count("<w:tbl>") == 1
            assert "Section 1.1." in document_xml
            assert "译文" in document_xml
            assert "大语言模型" in MockChatHandler.last_system
            assert "只输出译文" in MockChatHandler.last_system

        run("翻译：中英对照 DOCX + 术语表 + 真实进度", test_translate_bilingual)

        def test_translate_translated():
            task_id = submit_translate("english.pdf", samples["english"].read_bytes(), "translated")
            payload, _, _ = poll_task_tracked(task_id)
            result = payload["result"]
            assert result["filename"] == "english_译文.docx"
            status, _, content = http_request("GET", result["download_url"])
            assert status == 200
            with zipfile.ZipFile(io.BytesIO(content)) as archive:
                document_xml = archive.read("word/document.xml").decode("utf-8", "replace")
            assert "<w:tbl>" not in document_xml
            assert document_xml.count("译文") >= 6
            assert "Section 1.1." not in document_xml

        run("翻译：纯译文 DOCX", test_translate_translated)

        def test_translate_layout():
            task_id = submit_translate("english.pdf", samples["english"].read_bytes(), "layout")
            payload, _, _ = poll_task_tracked(task_id)
            result = payload["result"]
            assert result["filename"] == "english_版式译文.pdf"
            status, _, content = http_request("GET", result["download_url"])
            assert status == 200 and content.startswith(b"%PDF")
            document = fitz.open(stream=content, filetype="pdf")
            try:
                assert document.page_count == 3
                text = "\n".join(page.get_text() for page in document)
            finally:
                document.close()
            assert "译文" in text
            assert "Section 1.1." not in text
            assert "嵌入字体" in result["note"]

        run("翻译：保留版式 PDF 原位覆盖", test_translate_layout)

        def test_translate_layout_docx_rejected():
            body, content_type = multipart_named(
                {"target_lang": "zh", "mode": "layout", "glossary": ""},
                [
                    (
                        "烟雾测试.docx",
                        samples["docx"].read_bytes(),
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    )
                ],
                field_name="file",
            )
            status, _, response = http_request(
                "POST",
                "/api/translate",
                data=body,
                headers={"Content-Type": content_type, "Content-Length": str(len(body))},
            )
            assert status == 400
            payload = json.loads(response.decode("utf-8"))
            assert payload["error"] == "保留版式输出只支持 PDF，请先用格式转换转成 PDF"

        run("翻译：保留版式拒绝 DOCX（400）", test_translate_layout_docx_rejected)

        def test_translate_without_key():
            global BASE_URL
            no_key_port = pick_port()
            no_key_log = temp_dir / "no-key-server.log"
            no_key_env = env.copy()
            no_key_env.update(
                {
                    "PORT": str(no_key_port),
                    "TRANSLATE_API_KEY": "",
                    "TRANSLATE_BASE_URL": "",
                    "TRANSLATE_MODEL": "",
                }
            )
            handle = no_key_log.open("wb")
            no_key_process = subprocess.Popen(
                [str(PYTHON), str(ROOT / "main.py")],
                cwd=str(ROOT),
                env=no_key_env,
                stdout=handle,
                stderr=subprocess.STDOUT,
                creationflags=creationflags,
            )
            previous = BASE_URL
            BASE_URL = f"http://127.0.0.1:{no_key_port}"
            try:
                deadline = time.time() + 25
                ready = False
                while time.time() < deadline:
                    if no_key_process.poll() is not None:
                        break
                    try:
                        status, payload, _ = json_request("GET", "/api/health", timeout=3)
                        if status == 200 and payload.get("status") == "ok":
                            ready = True
                            break
                    except Exception:
                        pass
                    time.sleep(0.25)
                assert ready, "无密钥服务未启动"
                status, config, _ = json_request("GET", "/api/config")
                assert status == 200 and config["translate_configured"] is False
                body, content_type = multipart_named(
                    {"target_lang": "zh", "mode": "bilingual", "glossary": ""},
                    [("smoke.pdf", samples["pdf"].read_bytes(), "application/pdf")],
                    field_name="file",
                )
                status, _, response = http_request(
                    "POST",
                    "/api/translate",
                    data=body,
                    headers={"Content-Type": content_type, "Content-Length": str(len(body))},
                )
                assert status == 400
                payload = json.loads(response.decode("utf-8"))
                assert payload["error"] == "未配置翻译密钥，请在 .env 里填写 TRANSLATE_API_KEY"
            finally:
                BASE_URL = previous
                if no_key_process.poll() is None:
                    no_key_process.terminate()
                    try:
                        no_key_process.wait(timeout=5)
                    except subprocess.TimeoutExpired:
                        no_key_process.kill()
                        no_key_process.wait(timeout=5)
                handle.close()

        run("翻译：未配置密钥返回 400", test_translate_without_key)

        def test_scan_photo():
            task_id = submit_scan("photo.jpg", samples["photo"].read_bytes())
            payload, messages, percents = poll_task_tracked(task_id)
            result = payload["result"]
            assert result["filename"] == "photo_扫描增强.png"
            status, _, content = http_request("GET", result["download_url"])
            assert status == 200 and content.startswith(b"\x89PNG")
            output = Image.open(io.BytesIO(content))
            try:
                assert output.width > 300 and output.height > 300
                assert output.mode in {"RGB", "RGBA", "L"}
            finally:
                output.close()
            assert "未检测到纸张边界" not in result["note"], result["note"]
            assert max(percents) >= 40

        run("扫描件增强：倾斜 + 阴影照片", test_scan_photo)

        def test_scan_rejects_pdf():
            body, content_type = multipart_named(
                {},
                [("smoke.pdf", samples["pdf"].read_bytes(), "application/pdf")],
                field_name="file",
            )
            status, _, response = http_request(
                "POST",
                "/api/scan-photo",
                data=body,
                headers={"Content-Type": content_type, "Content-Length": str(len(body))},
            )
            assert status == 415
            payload = json.loads(response.decode("utf-8"))
            assert "不支持" in payload["error"]

        run("扫描件增强：非图片返回 415", test_scan_rejects_pdf)

    finally:
        mock_server.shutdown()
        mock_server.server_close()
        if process.poll() is None:
            process.terminate()
            try:
                process.wait(timeout=5)
            except subprocess.TimeoutExpired:
                process.kill()
                process.wait(timeout=5)
        log_handle.close()

    print()
    print(f"通过 {len(PASSED)} 项，失败 {len(FAILED)} 项")
    if FAILED:
        print("失败项：" + "、".join(FAILED))
        log_text = SERVER_LOG.read_text(encoding="utf-8", errors="replace") if SERVER_LOG.exists() else ""
        if log_text:
            print("\n服务日志最后 4000 字符：")
            print(log_text[-4000:])
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())