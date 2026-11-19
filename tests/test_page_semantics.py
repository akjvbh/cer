"""页码语义测试：删页、重排复制、最终旋转角度。"""

from __future__ import annotations

import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        _stream.reconfigure(encoding="utf-8", errors="replace")

import fitz  # noqa: E402

from pdf_tools import page_ops  # noqa: E402

PASSED: list[str] = []
FAILED: list[str] = []


def check(name: str, ok: bool, detail: str = "") -> None:
    if ok:
        PASSED.append(name)
        print(f"  [OK] {name}")
    else:
        FAILED.append(f"{name}：{detail}")
        print(f"  [FAIL] {name}：{detail}")


def make_pdf(path: Path, pages: int = 5) -> None:
    doc = fitz.open()
    try:
        for index in range(1, pages + 1):
            page = doc.new_page(width=300 + index * 10, height=500 + index * 10)
            page.insert_text((40, 80), f"page-{index}", fontsize=20)
        doc.save(str(path))
    finally:
        doc.close()


def page_texts(path: Path) -> list[str]:
    doc = fitz.open(str(path))
    try:
        return [page.get_text().strip() for page in doc]
    finally:
        doc.close()


def main() -> int:
    print("页码语义测试")
    with tempfile.TemporaryDirectory(prefix="pagesem-") as temp:
        work = Path(temp)
        src = work / "five.pdf"
        make_pdf(src, 5)

        print("\n[1] 删除第 2 页")
        out, note = page_ops(src, work / "deleted", [2], {}, [])
        texts = page_texts(out)
        check("删除后共 4 页", len(texts) == 4, f"实际 {len(texts)} 页")
        check("删除的是原文件第 2 页", texts == ["page-1", "page-3", "page-4", "page-5"], str(texts))
        check("note 说明删除 1 页", "删除 1 页" in note, note)

        print("\n[2] page_order 重复页码 = 复制页")
        out, note = page_ops(src, work / "copied", [], {}, [1, 1, 2, 3, 4, 5])
        texts = page_texts(out)
        check("重排后共 6 页", len(texts) == 6, f"实际 {len(texts)} 页")
        check("第 1、2 页都是原第 1 页", texts[:2] == ["page-1", "page-1"], str(texts))
        check("后续页按白名单顺序", texts[2:] == ["page-2", "page-3", "page-4", "page-5"], str(texts))

        print("\n[3] 旋转 90 度后宽高互换")
        src_doc = fitz.open(str(src))
        try:
            before = fitz.open()
            before.insert_pdf(src_doc, from_page=0, to_page=0)
            before_rect = before[0].rect
            before.close()
        finally:
            src_doc.close()
        out, note = page_ops(src, work / "rotated", [], {1: 90}, [])
        doc = fitz.open(str(out))
        try:
            after_rect = doc[0].rect
            rotation = doc[0].rotation
        finally:
            doc.close()
        check("旋转角度为 90", rotation == 90, str(rotation))
        check("宽高互换", abs(after_rect.width - before_rect.height) < 1 and abs(after_rect.height - before_rect.width) < 1,
              f"前 {before_rect.width}x{before_rect.height}，后 {after_rect.width}x{after_rect.height}")
        check("note 说明旋转 1 页", "旋转 1 页" in note, note)

        print("\n[4] 删除与旋转同时提交")
        out, note = page_ops(src, work / "mixed", [2], {1: 180}, [])
        doc = fitz.open(str(out))
        try:
            mixed_texts = [page.get_text().strip() for page in doc]
            mixed_rotation = doc[0].rotation
            mixed_pages = doc.page_count
        finally:
            doc.close()
        check("先删页再旋转：共 4 页", mixed_pages == 4, str(mixed_pages))
        check("第 1 页仍是原第 1 页", mixed_texts[0] == "page-1", str(mixed_texts))
        check("旋转角度为 180", mixed_rotation == 180, str(mixed_rotation))

    print(f"\n通过 {len(PASSED)} 项，失败 {len(FAILED)} 项")
    if FAILED:
        for item in FAILED:
            print(f"  - {item}")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
