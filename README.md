# 文档工坊

一个**在本机运行的办公文档转换与处理工具箱**：拖入文件、选择目标格式，转换在本机完成，结果直接下载。不需要登录、不上传云端、不依赖 Docker 或数据库。

![Python](https://img.shields.io/badge/Python-3.11-3776AB?logo=python&logoColor=white)
![FastAPI](https://img.shields.io/badge/FastAPI-0.1x-009688?logo=fastapi&logoColor=white)
![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-0078D4?logo=windows&logoColor=white)
![License](https://img.shields.io/badge/License-Private-lightgrey)

> 当前版本 **0.2.0 / 阶段 2（格式矩阵与批量处理）**：17 种源格式键（含 `jpg`/`jpeg` 别名）、单文件转换、按格式分组的批量转换、失败隔离与重试均已可用；压缩、PDF 工具、AI 增强为后续阶段。

## 快速开始

**方式一：双击**

```
start.bat
```

脚本会自动创建 `.venv`、安装依赖、选择 8000（被占用则 8001）端口并打开浏览器。关闭命令窗口即停止服务。

**方式二：手动**

```bat
py -3.11 -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe main.py
```

然后打开 <http://127.0.0.1:8000>。

环境要求：

- Windows 10/11（Office COM 相关能力仅 Windows 可用）
- Python 3.11
- 可选：本机安装 Microsoft Office（用于高保真 `docx → pdf`）。未安装时自动降级，不影响启动与其他功能。

## 已支持的转换

| 源格式 | 目标格式 | 实现方式 | 说明 |
| --- | --- | --- | --- |
| DOCX | PDF / HTML / MD / TXT | Word COM；PDF 失败时 ReportLab 降级；其余 python-docx + mammoth + markdown | PDF 有 Office 时保留原版式；HTML/MD 保留标题、段落与表格 |
| PDF | DOCX / TXT / JPG / PNG | pdf2docx、PyMuPDF | 多页图片打包 ZIP，逐页 `001.jpg` / `001.png` |
| XLSX | CSV / PDF | openpyxl；Excel COM，失败时 HTML 表格降级 | CSV 为 UTF-8 BOM；多工作表分别导出并打包 ZIP；PDF 降级只渲染前 3000 行、30 列 |
| XLS | XLSX / CSV | xlrd + openpyxl | 保留全部工作表；公式转为当前计算值；CSV 多表打包 ZIP |
| CSV | XLSX | csv + openpyxl | 按原文本写入，避免前导零被当成数字破坏 |
| DOC | PDF / DOCX | Word COM | 需要本机 Microsoft Word，否则返回明确错误 |
| PPTX / PPT | PDF | PowerPoint COM | 需要本机 Microsoft PowerPoint |
| PPTX | PNG | PowerPoint COM | 每页一张 PNG，打包 ZIP |
| Markdown | HTML / PDF | markdown；ReportLab Story | HTML 带表格、代码块与目录；PDF 支持中文与基本排版 |
| TXT | DOCX / PDF | python-docx；ReportLab | 保留空行与段落，正文使用微软雅黑 |
| PNG / JPG / JPEG / WEBP / BMP / GIF / TIFF | PDF / DOCX / 其他图片格式 | Pillow + PyMuPDF | 多帧图片只取首帧；转 PDF 时按 A4 居中，不裁切 |

格式表由 `converter.py` 的注册表自动生成，前端与 `/api/formats` 共用同一份数据，没有第二份手写映射。

## 批量处理

在左侧选择「批量处理」：

1. 一次选择多个文件，或点击「选择文件夹」；也可以把文件夹直接拖进虚线区域（浏览器支持时递归读取子目录）。
2. 文件按扩展名自动分组，每组选择一个目标格式；不同组可以给不同目标。
3. 点击「开始批量处理」，每组各发一个 `/api/batch` 请求并行执行。
4. 每个文件独立显示成功、失败或排队状态；单个文件失败不会中断整批。
5. 全部结束后可下载整批 ZIP，也可在单行点击「下载」；失败行会显示原因，可「重试失败项」只重跑失败文件。

默认限制：单文件 200MB、单批最多 200 个文件、单批总大小 500MB。超限或不支持的扩展名会在开始前标红并说明原因，不会静默丢弃。

## 使用方式

1. 打开页面，把文件拖进虚线区域，或点击选择文件。
2. 目标格式下拉框会根据文件类型自动列出可选项。
3. 点击「开始转换」，进度条按真实任务状态更新。
4. 完成后在「转换结果」中下载；失败会显示可复制的错误信息与重试按钮。

上传限制：默认单文件不超过 **200MB**，批量总计不超过 **500MB**、最多 **200 个文件**；文件名、扩展名、文件头（magic）都会校验，服务端使用 UUID 保存，不信任任何用户提供的路径。

## 配置

复制 `.env.example` 为 `.env` 后按需修改，重启生效：

| 变量 | 默认值 | 作用 |
| --- | --- | --- |
| `PORT` | `8000` | 手动运行 `main.py` 时的监听端口 |
| `AUTO_OPEN_BROWSER` | `1` | 启动后是否自动打开浏览器 |
| `MAX_FILE_SIZE_MB` | `200` | 单文件大小上限 |
| `MAX_BATCH_FILES` | `200` | 单批最多文件数 |
| `MAX_BATCH_SIZE_MB` | `500` | 单批总大小上限 |
| `OFFICE_ENABLED` | `1` | 是否允许调用本机 Office COM |
| `OFFICE_TIMEOUT_SECONDS` | `600` | 单次 Office 转换超时；`start.bat` 的演示默认值为 120 |
| `TRANSLATE_API_KEY` | 空 | 翻译能力预留（阶段 4） |
| `TRANSLATE_BASE_URL` | DeepSeek 兼容地址 | 翻译接口地址（阶段 4） |
| `TRANSLATE_MODEL` | `deepseek-chat` | 翻译模型（阶段 4） |

翻译是唯一可能把**提取出的文本**发送到外部地址的能力，且只发送到你配置的接口；其余转换全部在本机完成。

## API

所有转换都是异步任务：先上传拿 `task_id`，再轮询状态，完成后通过下载接口取文件。

| 方法 | 路径 | 说明 |
| --- | --- | --- |
| `GET` | `/api/health` | 健康检查，返回 Python 版本、版本号、Office 是否可用 |
| `GET` | `/api/config` | 前端配置：大小上限、翻译是否配置、术语表 |
| `GET` | `/api/formats` | 由转换注册表反推的格式表 |
| `POST` | `/api/convert` | multipart：`file` + `target`，返回 `task_id` |
| `POST` | `/api/batch` | multipart：`files`（可重复）+ `target`，返回 `task_id`；结果在任务 `result` 中逐文件列出 |
| `GET` | `/api/tasks/{task_id}` | 任务状态：`pending / running / done / error`、进度、结果 |
| `GET` | `/api/download/{task_id}` | 仅完成任务可下载，中文文件名用 `filename*` 传递；批量任务用 `?file=<序号>` 下载单个文件 |

错误统一返回 `{"error": "中文说明"}`，并使用有意义的 HTTP 状态码（`400` 参数不支持、`413` 过大、`415` 文件类型不符、`404` 任务不存在、`409` 未完成）。

示例：

```bat
curl.exe -F "file=@samples\示例_论文.pdf" -F "target=txt" http://127.0.0.1:8000/api/convert
curl.exe http://127.0.0.1:8000/api/tasks/<task_id>
curl.exe -OJ http://127.0.0.1:8000/api/download/<task_id>

rem 批量：同一 target 下可重复 files 字段
curl.exe -F "files=@a.docx" -F "files=@b.docx" -F "target=pdf" http://127.0.0.1:8000/api/batch
curl.exe -OJ "http://127.0.0.1:8000/api/download/<task_id>?file=0"
```

## 降级与已知限制

- **未安装 Office**：`docx → pdf` 自动使用 ReportLab 重新排版，文字可选中，但字体、行距、分页与原文件可能不同；页面会提示当前为降级模式。
- **Office COM 串行执行**：同一时间只允许一个 Word/Excel/PowerPoint 实例执行转换，避免 COM 冲突；Office 超时后会被强制结束并清理残留进程。
- **不提供 OCR**：扫描版 PDF 没有文字层时，转 TXT/DOCX 只能得到空内容，需等待阶段 4 的扫描增强能力。
- **PDF → DOCX 是推断版式**：表格、文本框、多栏可能错位，不保证与原文像素级一致。
- **服务仅监听 `127.0.0.1`**：没有登录、权限与 HTTPS，请不要把端口暴露到局域网或公网。
- **任务存在内存中**：重启服务后任务记录消失；结果默认保留 1 小时，过期自动清理。
- **仅 Windows**：Office COM 与 `pywin32` 绑定 Windows；其他系统只能使用纯 Python 链路（且当前以 Windows 为主要支持平台）。
- **Office 专属格式**：`.doc → pdf/docx` 需要 Word，`.ppt/.pptx → pdf` 与 `.pptx → png` 需要 PowerPoint。未安装时这些任务会失败并给出明确提示，不会返回 500。
- **批量中的 Office 任务串行**：同一时刻仍只允许一个 Word/Excel/PowerPoint 实例；纯 Python 任务仍按 4 线程并行。
- **批量 ZIP 会扁平化目录**：所有成功结果放在 ZIP 根目录，重名自动加 `(2)`、`(3)`；不还原原始文件夹层级。
- 阶段 3–4 尚未实现：压缩、PDF 合并/拆分/旋转/重排、文献翻译、扫描增强与 OCR。

## 项目结构

```
main.py            FastAPI 应用、上传校验、任务与下载
converter.py       转换注册表与全部转换实现
office_worker.py   独立子进程执行 Word/Excel/PPT COM
static/            原生 HTML/CSS/JS 前端（无框架、无构建）
tests/smoke.py     自启动服务的端到端冒烟测试
samples/           可直接拖入的示例文件
uploads/ outputs/  运行时目录（已忽略，不提交）
docs/prompts/      需求、决策与阶段文档
requirements.txt   Python 依赖
start.bat          Windows 一键启动
```

## 测试

```bat
.venv\Scripts\python.exe tests\smoke.py
```

测试会在随机端口自行启动服务（关闭 Office），用真实 HTTP 上传完成 34 项检查，覆盖：

- 健康检查、配置上限、格式注册表、首页静态资源；
- DOCX → TXT/HTML/MD/PDF、PDF → TXT/DOCX/JPG/PNG、XLSX → CSV/单表与多表、CSV → XLSX、XLS → XLSX/CSV、Markdown/TXT → PDF、图片互转与图片 → Word；
- Office 不可用时 `.doc`/`.pptx` 的明确错误；
- 批量 3 种格式的 ZIP 与单文件下载、坏文件失败隔离、ZIP 重名加序号、批量数量与参数校验；
- 415 扩展名、400 不支持的组合、404、中文下载头与 CSV 的 UTF-8 BOM。

前端验证：桌面 1440px、移动 600px、浅色/深色主题均已人工检查；浏览器控制台无报错。

## 阶段规划

| 阶段 | 内容 | 状态 |
| --- | --- | --- |
| 1 | 骨架、异步任务、上传校验、6 类主干转换、界面、冒烟测试 | 已完成 |
| 2 | 完整格式矩阵、批量与文件夹、失败清单与重试 | 已完成 |
| 3 | 压缩、PDF 合并/拆分/旋转/重排、缩略图编辑器 | 计划中 |
| 4 | 翻译（双语/仅译文/版式叠加）、术语表、扫描增强 | 计划中 |
| 5 | 整体打磨、异常与性能验收、演示材料与交付 | 计划中 |

## 隐私

- 文件只保存在本机 `uploads/` 与 `outputs/`，转换完成后源上传文件即删除，结果按 TTL 清理。
- 不写日志记录文件内容；不记录 API Key。
- 除你主动配置并使用的翻译接口外，应用不发起任何外部网络请求。