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
        "png": image_path,
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
            assert payload.get("excel") is False
            assert payload.get("powerpoint") is False
            assert payload.get("version") == "0.2.0"
            head_status, _, _ = http_request("HEAD", "/api/health")
            assert head_status == 200

        run("health 与 HEAD 健康检查", test_health)

        def test_config():
            status, payload, _ = json_request("GET", "/api/config")
            assert status == 200
            assert payload["max_file_size_mb"] == 200
            assert payload["max_batch_files"] == 200
            assert payload["max_batch_size_mb"] == 500

        run("配置接口含批量上限", test_config)

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