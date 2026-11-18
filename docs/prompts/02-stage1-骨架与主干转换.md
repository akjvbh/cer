# 阶段 1：项目骨架 + 主干转换链路

## 当前状态

空目录，只有 `.qoder\rules\` 里的项目规则。没有任何代码。

## 本阶段目标

1. 建一个能启动的 FastAPI 骨架，浏览器能打开界面
2. 打通「上传 → 转换 → 下载」主干：`docx→pdf`、`docx→txt`、`pdf→docx`、`pdf→txt`、`pdf→jpg`、`图片→pdf`
3. 把后续阶段都要复用的公共设施先建好：上传校验、任务表、统一错误格式、健康检查、格式表
4. 写出 `tests/smoke.py`，之后每个阶段都跑它

## 要做的事（按这个顺序）

### 1. 环境与启动

- `requirements.txt` 先只放本阶段需要的：

```
fastapi
uvicorn[standard]
python-multipart
python-dotenv
PyMuPDF
pdf2docx
Pillow
python-docx
openpyxl
pywin32; sys_platform == "win32"
```

- 创建虚拟环境：`python -m venv .venv`，目录名必须是 `.venv`
- `start.bat`：
  - 不存在 `.venv` 就创建并 `pip install -r requirements.txt`
  - 用 `.venv\Scripts\python.exe` 启动 `main.py`
  - 启动后自动打开 `http://127.0.0.1:8000`
  - 端口被占用时换 8001 并提示（用 `netstat` 或启动失败后重试）
- `main.py` 里用 `uvicorn.run(app, host="127.0.0.1", port=8000)`

### 2. `main.py`

- FastAPI 实例；`/static` 静态挂载；`GET /` 返回 `static/index.html`
- 用 `python-dotenv` 读 `.env`，缺失时用默认值：`MAX_FILE_SIZE_MB=200`
- 任务表：`TASKS: dict[str, dict] = {}`，字段 `status`（pending/running/done/error）、`progress`、`message`、`result`、`error`、`created_at`
- 上传校验函数（所有上传接口共用）：
  - 扩展名白名单
  - 大小上限（默认 200MB），超限返回 413 + `{"error":"文件超过 200MB 上限"}`
  - 文件头校验：PDF 判断 `%PDF`，docx/xlsx/pptx 判断 `PK\x03\x04`，图片判断对应 magic
  - 落盘到 `uploads/<uuid4().hex>.<ext>`，**不使用用户提供的文件名作为路径**
- 路由：
  - `GET /api/health` → `{"status":"ok","office":true/false,"python":"3.11.x","version":"0.2.0"}`。office 字段用**注册表或可执行文件探测**得出（查 `winword.exe` 是否安装 / COM 类是否注册），**不要启动 Word 探活**——拉起一个 WINWORD 进程既慢又会残留。结果缓存 60 秒
  - `GET /api/config` → `{"max_file_size_mb":200,"translate_configured":false,"glossary":""}`
  - `GET /api/formats` → **由转换器注册表自动生成**，不要手写第二份映射
  - `POST /api/convert`（`file`, `target`）→ 校验 → 生成 `task_id` → 同步执行 → 返回 `{"task_id","filename","size","download_url","note"}`
  - `GET /api/tasks/{task_id}` → 未找到返回 404
  - `GET /api/download/{task_id}` → `FileResponse`；`Content-Disposition` 中文文件名用 `filename*=UTF-8''<urlencode>`

### 3. `converter.py`

- 注册表：`HANDLERS: dict[tuple[str, str], Callable]`，键是 `(源扩展名, 目标扩展名)`
- 统一签名：`def _x(src: Path, dst_dir: Path, target: str) -> tuple[Path, str]`，返回 `(输出路径, note)`
- 提供 `get_formats() -> dict[str, list[str]]` 给 `/api/formats` 用，由注册表反推
- 本阶段实现的 handler：
  - `docx→txt`：python-docx 读段落 + 表格
  - `pdf→txt`：PyMuPDF `page.get_text()`
  - `pdf→docx`：pdf2docx
  - `pdf→jpg`：PyMuPDF 按 150 DPI 渲染第 1 页（多页留到阶段 3）
  - 图片→pdf：Pillow
  - `docx→pdf`：**优先**走 COM（`office_worker.py`），失败则走兜底
- COM 单独放 `office_worker.py`：`pythoncom.CoInitialize()`、`DisplayAlerts=False`、`Visible=False`、全局 `threading.Lock`、超时（默认 600 秒，读 `OFFICE_TIMEOUT_SECONDS`）后 `taskkill /F /IM WINWORD.EXE`
- `docx→pdf` 兜底：用 PyMuPDF 的 `page.insert_htmlbox()` 把 docx 的段落、标题、列表、表格排成 PDF；note 写「Office 不可用，已用内置排版引擎，字体与行距可能与 Word 有出入」
- 所有 handler 失败时抛 `ConversionError("中文原因")`，由路由层统一转成 `{"error": ...}`

### 4. `static/`

- `index.html`：
  - 顶栏：产品名 + 主题切换按钮 + 健康状态徽标（Office 可用 / 不可用）
  - 左侧：5 个功能标签（格式转换 / 批量处理 / 压缩 / PDF 工具 / AI 增强），本阶段只做「格式转换」，其它标签显示「即将上线」占位
  - 中间：拖拽上传区（支持点击选择、拖文件、拖文件夹）+ 已选文件列表
  - 右侧：结果卡片（文件名、大小、下载按钮、note）
- `app.css`：CSS 变量定义配色，`:root` 浅色、`[data-theme="dark"]` 深色，默认跟随 `prefers-color-scheme`；不要用框架类名
- `app.js`：
  - 全局 `state` 对象
  - 拖拽 + 选择文件、显示列表
  - 从 `/api/formats` 拉目标格式，联动源格式生成下拉
  - 上传用 `XMLHttpRequest` 显示上传进度
  - 结果卡片渲染，文件名一律用 `textContent`
  - 主题切换写 `localStorage`

### 5. `tests/smoke.py`

用标准库 `urllib` 或 `requests`，启动服务后依次检查并打印 `PASS/FAIL`，任一失败退出码 1：

1. `GET /api/health` 返回 200 且 `status == "ok"`
2. `GET /api/formats` 里 `docx` 包含 `pdf`
3. 上传 `samples\参考论文_红鳍东方鲀.pdf`，`target=txt`，拿 `download_url` 下载，文件非空且包含中文
4. `GET /api/download/<乱写的id>` 返回 404
5. 上传一个 `.exe` 文件返回 415

## 硬约束

- 不引入数据库、不引入前端框架、不出现 `package.json`
- 不允许 `except: pass`
- 中文文件名在上传、日志、下载三处都不能出错
- 没有 Office 时 `docx→pdf` 必须能出文件并带 `note`，不能 500
- 本阶段只做同步转换，但接口形态保留 `task_id`，且**只保留一套任务模型**（`POST` 返回 `task_id` → `GET /api/tasks/{id}` → `GET /api/download/{id}`），阶段 3 直接把同步换成异步，不改契约

## 验收（请真实执行，把命令和输出贴给我）

1. `.\.venv\Scripts\python.exe main.py` 启动成功，日志无 traceback
2. 浏览器打开 `http://127.0.0.1:8000`，能看到界面，拖入 `samples\参考论文_红鳍东方鲀.pdf` 选 txt 能下载
3. 命令行验证（PowerShell）：

```powershell
curl.exe -F "file=@samples\参考论文_红鳍东方鲀.pdf" -F "target=txt" http://127.0.0.1:8000/api/convert
curl.exe -F "file=@samples\初稿_田螺厣盖个体识别.docx" -F "target=pdf" http://127.0.0.1:8000/api/convert
curl.exe -I http://127.0.0.1:8000/api/health
```

4. `.\.venv\Scripts\python.exe tests\smoke.py` 全 PASS
5. `git init`，`.gitignore` 排除 `.venv/ __pycache__/ uploads/ outputs/ .env`，首次提交

## 完成后告诉我

- 新增/修改了哪些文件，各自职责一句话
- 哪些是降级实现（尤其没有 Office 时的 `docx→pdf`）
- 你实际跑过的命令与结果
- 留给下一阶段的 TODO 或已知问题
