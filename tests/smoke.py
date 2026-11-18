"""阶段 1 端到端冒烟测试：自起服务，真实上传、转换、轮询、下载。"""

from __future__ import annotations

import io
import json
import os
import socket
import subprocess
import sys
import tempfile
import time
import urllib.error
import urllib.request
import uuid
import zipfile
from pathlib import Path

import fitz
from PIL import Image, ImageDraw
from docx import Document

ROOT = Path(__file__).resolve().parents[1]
PYTHON = Path(sys.executable)
for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        _stream.reconfigure(encoding="utf-8", errors="replace")
PASSED: list[str] = []
FAILED: list[str] = []
SERVER_LOG: Path | None = None
BASE_URL = ""


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
    for _ in range(300):
        status, payload, _ = json_request("GET", f"/api/tasks/{task_id}")
        if status != 200:
            raise AssertionError(f"轮询失败 HTTP {status}: {payload}")
        if payload["status"] == "done":
            result = payload["result"]
            download_status, headers, content_bytes = http_request("GET", result["download_url"])
            if download_status != 200:
                raise AssertionError(f"下载失败 HTTP {download_status}")
            return payload, result, content_bytes
        if payload["status"] == "error":
            raise AssertionError(f"任务失败：{payload.get('error')}")
        time.sleep(0.2)
    raise AssertionError("任务在 60 秒内没有完成")


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

    image_path = directory / "smoke.png"
    image = Image.new("RGB", (480, 320), "white")
    draw = ImageDraw.Draw(image)
    draw.rectangle((40, 40, 440, 280), outline="#2f63f4", width=5)
    draw.ellipse((190, 110, 290, 210), fill="#2f63f4")
    image.save(str(image_path), "PNG")
    image.close()

    md_path = directory / "smoke.md"
    md_path.write_text("# 烟雾测试标题\n\n这是 **Markdown** 正文。\n\n| A | B |\n|---|---|\n| 1 | 2 |\n", encoding="utf-8")

    txt_path = directory / "smoke.txt"
    txt_path.write_text("第一行中文\n第二行 smoke text\n\n第三段\n", encoding="utf-8")

    exe_path = directory / "bad.exe"
    exe_path.write_bytes(b"MZ" + b"\x00" * 128)

    return {
        "docx": docx_path,
        "pdf": pdf_path,
        "png": image_path,
        "md": md_path,
        "txt": txt_path,
        "exe": exe_path,
    }


def main() -> int:
    global BASE_URL, SERVER_LOG
    port = pick_port()
    BASE_URL = f"http://127.0.0.1:{port}"
    temp_dir = Path(tempfile.mkdtemp(prefix="docworkshop-smoke-"))
    SERVER_LOG = temp_dir / "server.log"
    samples = make_samples(temp_dir)
    log_handle = SERVER_LOG.open("wb")

    env = os.environ.copy()
    env.update(
        {
            "PORT": str(port),
            "OFFICE_ENABLED": "0",
            "AUTO_OPEN_BROWSER": "0",
            "PYTHONUNBUFFERED": "1",
            "TASK_TTL_SECONDS": "3600",
            "CLEANUP_INTERVAL_SECONDS": "60",
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
            assert payload.get("version") == "0.2.0"
            head_status, _, _ = http_request("HEAD", "/api/health")
            assert head_status == 200

        run("health 与 HEAD 健康检查", test_health)

        def test_formats():
            status, payload, _ = json_request("GET", "/api/formats")
            assert status == 200
            assert "pdf" in payload["docx"] and "txt" in payload["docx"]
            assert "docx" in payload["pdf"] and "txt" in payload["pdf"] and "jpg" in payload["pdf"]
            assert payload["png"] == ["pdf"]
            assert payload["md"] == ["html"]
            assert payload["txt"] == ["docx"]

        run("格式表由注册表反推", test_formats)

        def test_homepage():
            status, headers, body = http_request("GET", "/")
            text = body.decode("utf-8")
            assert status == 200
            assert "格式转换" in text and "拖拽文件" in text
            assert "app.js" in text and "app.css" in text
            status, _, css = http_request("GET", "/static/app.css")
            assert status == 200 and len(css) > 1000
            status, _, js = http_request("GET", "/static/app.js")
            assert status == 200 and len(js) > 1000

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
                {"target": "pdf"},
                {"file": ("smoke.txt", samples["txt"].read_bytes(), "text/plain")},
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

    finally:
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