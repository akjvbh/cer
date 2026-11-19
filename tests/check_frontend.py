"""前端静态自检：括号配平、id 引用、资源路径、HTML 标签闭合。"""

from __future__ import annotations

import re
import sys
from html.parser import HTMLParser
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
STATIC = ROOT / "static"
HTML_PATH = STATIC / "index.html"
CSS_PATH = STATIC / "app.css"
JS_PATH = STATIC / "app.js"

for _stream in (sys.stdout, sys.stderr):
    if hasattr(_stream, "reconfigure"):
        _stream.reconfigure(encoding="utf-8", errors="replace")

VOID_TAGS = {
    "area", "base", "br", "col", "embed", "hr", "img", "input",
    "link", "meta", "param", "source", "track", "wbr",
}
OPTIONAL_CLOSE = {"p", "li", "option", "tr", "td", "th", "thead", "tbody", "dt", "dd"}

PASSED: list[str] = []
FAILED: list[str] = []


def check(name: str, ok: bool, detail: str = "") -> None:
    if ok:
        PASSED.append(name)
        print(f"  [OK] {name}")
    else:
        FAILED.append(f"{name}：{detail}")
        print(f"  [FAIL] {name}：{detail}")


def strip_js_literals(source: str) -> str:
    """去掉注释、字符串、模板字符串，保留代码骨架，便于括号配平。"""
    out: list[str] = []
    index = 0
    length = len(source)
    while index < length:
        char = source[index]
        nxt = source[index + 1] if index + 1 < length else ""
        if char == "/" and nxt == "/":
            index += 2
            while index < length and source[index] not in "\r\n":
                index += 1
            continue
        if char == "/" and nxt == "*":
            index += 2
            while index + 1 < length and not (source[index] == "*" and source[index + 1] == "/"):
                index += 1
            index += 2
            continue
        if char in "\"'`":
            quote = char
            index += 1
            while index < length:
                if source[index] == "\\":
                    index += 2
                    continue
                if source[index] == quote:
                    index += 1
                    break
                index += 1
            out.append(" ")
            continue
        out.append(char)
        index += 1
    return "".join(out)


def check_brackets(source: str) -> None:
    skeleton = strip_js_literals(source)
    pairs = {")": "(", "]": "[", "}": "{"}
    stack: list[tuple[str, int]] = []
    line = 1
    for index, char in enumerate(skeleton):
        if char == "\n":
            line += 1
        if char in "([{":
            stack.append((char, line))
        elif char in ")]}":
            if not stack:
                check("括号配平", False, f"第 {line} 行出现多余的 {char}")
                return
            opener, opener_line = stack.pop()
            if opener != pairs[char]:
                check("括号配平", False, f"第 {opener_line} 行的 {opener} 与第 {line} 行的 {char} 不匹配")
                return
    if stack:
        opener, opener_line = stack[-1]
        check("括号配平", False, f"第 {opener_line} 行的 {opener} 没有闭合")
        return
    check("括号配平", True)


class TagChecker(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.stack: list[tuple[str, int]] = []
        self.errors: list[str] = []

    def handle_starttag(self, tag: str, attrs) -> None:
        if tag in VOID_TAGS:
            return
        self.stack.append((tag, self.getpos()[0]))

    def handle_startendtag(self, tag: str, attrs) -> None:
        return

    def handle_endtag(self, tag: str) -> None:
        if tag in VOID_TAGS:
            return
        if not self.stack:
            self.errors.append(f"第 {self.getpos()[0]} 行出现多余的 </{tag}>")
            return
        if self.stack[-1][0] == tag:
            self.stack.pop()
            return
        for depth in range(len(self.stack) - 1, -1, -1):
            if self.stack[depth][0] == tag:
                for unclosed, line in self.stack[depth + 1:]:
                    if unclosed not in OPTIONAL_CLOSE:
                        self.errors.append(f"第 {line} 行的 <{unclosed}> 没有闭合")
                del self.stack[depth:]
                return
        self.errors.append(f"第 {self.getpos()[0]} 行闭合了未打开的 </{tag}>")

    def finish(self) -> None:
        for tag, line in self.stack:
            if tag not in OPTIONAL_CLOSE:
                self.errors.append(f"第 {line} 行的 <{tag}> 没有闭合")


def main() -> int:
    print("前端静态自检")
    for path in (HTML_PATH, CSS_PATH, JS_PATH):
        if not path.exists():
            print(f"缺少文件：{path.relative_to(ROOT)}")
            return 1

    html = HTML_PATH.read_text(encoding="utf-8")
    js = JS_PATH.read_text(encoding="utf-8")
    css = CSS_PATH.read_text(encoding="utf-8")

    print("\n[1] JavaScript 括号")
    check_brackets(js)

    print("\n[2] HTML 标签闭合")
    parser = TagChecker()
    parser.feed(html)
    parser.close()
    parser.finish()
    if parser.errors:
        check("HTML 标签闭合", False, "；".join(parser.errors[:5]))
    else:
        check("HTML 标签闭合", True)

    print("\n[3] getElementById / $() 引用的 id")
    html_ids = set(re.findall(r'\bid="([^"]+)"', html))
    referenced = set(re.findall(r'\$\("([^"]+)"\)', js))
    referenced |= set(re.findall(r'getElementById\("([^"]+)"\)', js))
    missing = sorted(item for item in referenced if item not in html_ids)
    check("脚本引用的 id 都存在于 HTML", not missing, "缺少：" + "、".join(missing))

    print("\n[4] HTML 引用的静态资源")
    resources = re.findall(r'(?:href|src)="(/static/[^"]+)"', html)
    missing_resources = [item for item in resources if not (ROOT / item.lstrip("/")).exists()]
    check("HTML 静态资源存在", not missing_resources, "缺少：" + "、".join(missing_resources))
    check("HTML 引用了 app.css 与 app.js", "/static/app.css" in resources and "/static/app.js" in resources)

    print("\n[5] app.js 中引用的静态路径")
    js_paths = set(re.findall(r'"(/static/[^"?#]+)"', js))
    js_paths |= set(re.findall(r"'(/static/[^'?#]+)'", js))
    missing_js = sorted(item for item in js_paths if not (ROOT / item.lstrip("/")).exists())
    check("app.js 静态路径存在", not missing_js, "缺少：" + "、".join(missing_js))

    print("\n[6] CSS 括号")
    css_stack: list[int] = []
    line = 1
    for char in css:
        if char == "\n":
            line += 1
        elif char == "{":
            css_stack.append(line)
        elif char == "}":
            if not css_stack:
                check("CSS 括号配平", False, f"第 {line} 行出现多余的 }}")
                break
            css_stack.pop()
    else:
        check("CSS 括号配平", not css_stack, f"第 {css_stack[-1]} 行的 {{ 没有闭合" if css_stack else "")

    print(f"\n通过 {len(PASSED)} 项，失败 {len(FAILED)} 项")
    if FAILED:
        for item in FAILED:
            print(f"  - {item}")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
