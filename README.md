# 文档工坊

一个**在本机运行的办公文档转换与处理工具箱**：拖入文件、选择目标格式，转换在本机完成，结果直接下载。不需要登录、不上传云端、不依赖 Docker 或数据库。

![Python](https://img.shields.io/badge/Python-3.11-3776AB?logo=python&logoColor=white)
![FastAPI](https://img.shields.io/badge/FastAPI-0.115+-009688?logo=fastapi&logoColor=white)
![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-0078D4?logo=windows&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green)

> 当前版本 **1.0.0 / 阶段 5（打磨与交付）**：转换、批量、压缩、PDF 工具、文献翻译与扫描件增强全部可用，并补齐三态主题、页面操作撤销/重做、后台清理、启动日志、三项自检脚本与交付文档。除翻译调用你配置的 OpenAI 兼容接口外，所有文件处理都在本机完成。

## 它解决什么问题

- 收到一份 `.doc` 或 `.pptx`，手边只有浏览器，不想上传到在线转换站：本机打开就能转 PDF。
- 论文、合同、说明书需要把几十页 PDF 拆页、删页、旋转或重新排序，又不想学专业 PDF 编辑器：缩略图里点选、拖拽、撤销即可。
- 扫描/拍照的纸质表格发灰、倾斜、有阴影：直接获得矫正、裁边、去阴影后的 PNG。
- 英文文献需要快速读懂，又希望保留原文对照或原版式：中英对照 DOCX、纯译文 DOCX、覆盖原文的 PDF 三种输出任选。
- 一批格式混杂的文件夹要统一处理：按扩展名自动分组、各组选目标格式、失败项单独重试，不因一个坏文件中断整批。

## 快速开始

**方式一：双击**

```
start.bat
```

脚本会自动创建 `.venv`、安装依赖，并从 8000 开始寻找空闲端口（8000 被占用则依次尝试 8001–8005），在窗口和控制台打印实际地址，同时把启动过程写入 `_start_log.txt`。关闭命令窗口即停止服务。

**方式二：手动**

```bat
py -3.11 -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe main.py
```

然后打开 <http://127.0.0.1:8000>（若 `.env` 里改了 `PORT`，以实际打印的地址为准）。

环境要求：

- Windows 10/11（Office COM 相关能力仅 Windows 可用）
- Python 3.11
- 可选：本机安装 Microsoft Office（用于高保真 `docx → pdf`）。未安装时自动降级，不影响启动与其他功能。

**Windows 上 `py` / `python` 不可用怎么办**

如果双击或输入 `python` 弹到 Microsoft Store，说明系统里的 `python.exe` 是应用商店占位程序。请从 [python.org](https://www.python.org/downloads/windows/) 安装 3.11，并在安装时勾选 “Add python.exe to PATH”；或直接使用已安装的 `py -3.11`。`start.bat` 优先调用项目内的 `.venv\Scripts\python.exe`，装好 `.venv` 后不依赖 PATH 里的 `python`。

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

## 压缩

左侧选择「压缩」，只接受 **PDF 与图片**，四档 `80 / 60 / 40 / 20` 对应不同的 DPI、JPEG 质量与最长边限制：

- 选中文件后先显示四档预估体积；PDF 只抽前 3 页估算，不为了预估跑完整份文件。
- PDF 压缩会把每页重渲染为 JPEG 图片，**文字层会丢失**，提交前会弹出页面内确认框，结果卡片也会再次说明。
- 图片压缩按最长边等比缩小，JPG 会先铺白底，避免透明区域变黑；动图与多帧 TIFF 只处理首帧并在结果中说明。
- 纯文字或本身很小的 PDF 重渲染后体积可能**变大**，结果说明会明确提示；原文件不会被修改。

## PDF 工具

左侧选择「PDF 工具」，页面分三块：

1. **页面编辑**：加载 PDF 后生成缩略图网格（最多显示前 50 页），可勾选、旋转、删除、复制、拖拽排序，修改只在本地状态中累积，点击「应用并下载」后一次性提交。删除、旋转、排序、复制共用同一个撤销栈（最多 50 步），支持页面按钮与 `Ctrl+Z` / `Ctrl+Shift+Z`，无可撤销/重做内容时按钮自动置灰。
2. **合并**：按选择顺序合并至少两个 PDF，支持首尾相接、逐页交替、插入到指定位置（`0` 表示插到第 1 页之前）。
3. **逐页导出 JPG**：为当前 PDF 的每一页生成 JPG 并打包 ZIP；每张图片必须小于等于设定上限，先降质量再降 DPI，到 `72 DPI + q50` 仍超限时该文件明确失败并给出当前体积，不会静默超限。

页码始终是 **1-based 且指原文件页码**；旋转提交的是**最终角度**，会覆盖该页原有 rotation；`page_order` 非空时作为白名单，重复页码表示复制该页。提交后后端按「删页 / 重排 → 旋转」的顺序处理。

## AI 增强

左侧选择「AI 增强」，页面分上下两块。

**文献翻译**（PDF / DOCX / TXT，一次一个文件）：

- PDF 按文字块提取并按「页码 → 纵坐标 → 横坐标」排序，折行会合并；页码、纯符号与过短片段会被过滤。DOCX 按正文顺序读取段落与表格单元格，TXT 按空行分段。扫描版 PDF 没有文字层时会明确提示，不会输出空文件。
- 段落按顺序打包成**不超过 2000 字符**的块，最多 3 个块并发请求 `/chat/completions`；单次超时 120 秒，429/5xx 按 2/4/8 秒重试 3 次，其他 4xx 立即失败。任何一块失败则整个任务失败，不生成半成品，错误会说明是第几块、第几段。
- **中英对照**：DOCX 两列表格，一段一行，左原文右译文。
- **纯译文**：DOCX 只保留译文，尽量沿用原标题级别、列表与表格结构。
- **保留版式**：只接受 PDF。先按段落矩形抹除原文，再把译文写回原位置；字号在 18pt 到 6pt 之间以 0.5pt 步进自动缩小，缩到 6pt 仍放不下才截断并加省略号。
- 版式模式优先嵌入 `assets/fonts/` 的中文字体子集；遇到子集缺字时自动改用系统微软雅黑，两者都缺字则明确报错，不会画出方框。
- 术语表由 `glossary.txt` 预填到界面，每行 `English=中文`，`#` 开头为注释。界面里追加的条目只对本次任务生效，写在后面的同名条目会覆盖前面的。
- 进度按已完成块数计算，例如「正在翻译第 12/58 段」。

**扫描件增强**（PNG / JPG / WEBP / BMP / GIF / TIFF）：先缩放长边到 1600px，再做灰度、模糊与 Canny 边缘检测，取面积最大的近似四边形做透视矫正；随后用大核高斯估计背景光并去除阴影，CLAHE 增强对比度，输出彩色 PNG。找不到纸张边界时不会报错，只做 1% 裁边并说明「未检测到纸张边界」。界面上原图与结果并排展示。

本版不做 OCR。翻译只会把**提取出的段落文本**发送到你配置的接口；图片、PDF 文件本身与扫描件增强都不会离开本机。

## 使用方式

1. 打开页面，把文件拖进虚线区域，或点击选择文件。
2. 目标格式下拉框会根据文件类型自动列出可选项。
3. 点击「开始转换」，进度条按真实任务状态更新。
4. 完成后在「转换结果」中下载；失败会显示可复制的错误信息与重试按钮。

上传限制：默认单文件不超过 **200MB**，批量总计不超过 **500MB**、最多 **200 个文件**；文件名、扩展名、文件头（magic）都会校验，服务端使用 UUID 保存，不信任任何用户提供的路径。

界面细节：

- 主题按钮按「跟随系统 → 浅色 → 深色」循环，选择写入 `localStorage`；首屏在 `<head>` 中先恢复主题，不会闪白/闪黑。
- 窗口宽度小于 900px 时左侧导航变为顶部横向滚动标签，结果卡片单列；600px 宽仍可正常操作。
- 统一顶部 toast 约 3 秒自动消失；失败结果卡片提供可展开的「查看详细错误」与复制按钮。
- 选择超过 50MB 的文件会提示「文件较大，转换可能需要 1-2 分钟」；上传/转换/压缩/导出期间按钮禁用并切换文案。
- `uploads/` 与 `outputs/` 由后台线程每小时清理一次超过 1 小时的临时文件；服务重启后任务记录清空属于预期行为。

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
| `TASK_TTL_SECONDS` | `3600` | 任务与结果保留秒数，过期由后台清理 |
| `CLEANUP_INTERVAL_SECONDS` | `3600` | 后台清理间隔秒数，最小值 60 |
| `START_LOG` | 空 | 由 `start.bat` 设为 `1` 时，把日志同时写入 `_start_log.txt` |
| `TRANSLATE_API_KEY` | 空 | 翻译接口密钥；为空时「AI 增强」中的翻译区域置灰 |
| `TRANSLATE_BASE_URL` | `https://api.deepseek.com/v1` | OpenAI 兼容接口前缀，实际请求 `<base>/chat/completions` |
| `TRANSLATE_MODEL` | `deepseek-chat` | 翻译使用的模型名 |

翻译是唯一可能把**提取出的文本**发送到外部地址的能力，且只发送到你配置的接口；其余转换与扫描件增强全部在本机完成。密钥只从环境变量 / `.env` 读取，日志中只出现掩码（如 `sk-****ab`）。术语表放在项目根目录的 `glossary.txt`，由 `/api/config` 原样返回给前端预填。

## API

所有转换都是异步任务：先上传拿 `task_id`，再轮询状态，完成后通过下载接口取文件。

| 方法 | 路径 | 说明 |
| --- | --- | --- |
| `GET` | `/api/health` | 健康检查，返回 Python 版本、版本号、Office 是否可用 |
| `GET` | `/api/config` | 前端配置：大小上限、翻译是否配置、术语表 |
| `GET` | `/api/formats` | 由转换注册表反推的格式表 |
| `POST` | `/api/convert` | multipart：`file` + `target`，返回 `task_id` |
| `POST` | `/api/batch` | multipart：`files`（可重复）+ `target`，返回 `task_id`；结果在任务 `result` 中逐文件列出 |
| `POST` | `/api/compress` | multipart：`file` + `level`（`80/60/40/20`），返回 `task_id`；只接受 PDF / 图片 |
| `POST` | `/api/compress/estimate` | multipart：`file`，直接返回四档预估体积；PDF 抽样前 3 页 |
| `POST` | `/api/pdf/merge` | multipart：`files`（至少 2 个）+ `mode`（`append/interleave/insert`）+ `insert_index`，返回 `task_id` |
| `POST` | `/api/pdf/ops` | multipart：`file` + `delete_pages` / `rotations` / `page_order`（JSON 字符串），返回 `task_id` |
| `POST` | `/api/pdf/preview` | multipart：`file` + `limit`（默认 50），直接返回每页 PNG base64、原 rotation 与尺寸 |
| `POST` | `/api/pdf/to-jpg` | multipart：`files` + `max_kb`（默认 300），返回 `task_id`；完成后下载 ZIP |
| `POST` | `/api/translate` | multipart：`file` + `target_lang`（`zh/en`）+ `mode`（`bilingual/translated/layout`）+ `glossary`，返回 `task_id`；`layout` 只接受 PDF |
| `POST` | `/api/scan-photo` | multipart：`file`（图片），返回 `task_id`；结果统一为 PNG |
| `GET` | `/api/tasks/{task_id}` | 任务状态：`pending / running / done / error`、进度、结果 |
| `GET` | `/api/download/{task_id}` | 仅完成任务可下载，中文文件名用 `filename*` 传递；批量任务用 `?file=<序号>` 下载单个文件 |

错误统一返回 `{"error": "中文说明"}`，并使用有意义的 HTTP 状态码（`400` 参数不支持、`413` 过大、`415` 文件类型不符、`404` 任务不存在、`409` 未完成）。

示例：

```bat
curl.exe -F "file=@samples\示例_论文.pdf" -F "target=txt" http://127.0.0.1:8000/api/convert
curl.exe http://127.0.0.1:8000/api/tasks/<task_id>
curl.exe -OJ http://127.0.0.1:8000/api/download/<task_id>

rem 批量：同一 target 下可重复 files 字段
curl.exe -F "files=@samples\示例_项目说明.docx" -F "files=@samples\示例_文本.txt" -F "target=pdf" http://127.0.0.1:8000/api/batch
curl.exe -OJ "http://127.0.0.1:8000/api/download/<task_id>?file=0"

rem 压缩与 PDF 工具
curl.exe -F "file=@samples\示例_论文.pdf" -F "level=60" http://127.0.0.1:8000/api/compress
curl.exe -F "file=@samples\示例_论文.pdf" http://127.0.0.1:8000/api/compress/estimate
curl.exe -F "files=@samples\示例_论文.pdf" -F "files=@samples\示例_英文摘要.pdf" -F "mode=append" http://127.0.0.1:8000/api/pdf/merge
curl.exe -F "file=@samples\示例_论文.pdf" -F "delete_pages=[2]" -F "rotations={\"1\":90}" http://127.0.0.1:8000/api/pdf/ops
curl.exe -F "file=@samples\示例_论文.pdf" -F "limit=10" http://127.0.0.1:8000/api/pdf/preview
curl.exe -F "files=@samples\示例_论文.pdf" -F "max_kb=300" http://127.0.0.1:8000/api/pdf/to-jpg

rem AI 增强（翻译需先在 .env 填 TRANSLATE_API_KEY）
curl.exe -F "file=@samples\示例_英文摘要.pdf" -F "target_lang=zh" -F "mode=bilingual" http://127.0.0.1:8000/api/translate
curl.exe -F "file=@samples\示例_英文摘要.pdf" -F "target_lang=zh" -F "mode=layout" http://127.0.0.1:8000/api/translate
curl.exe -F "file=@samples\示例_扫描照片.jpg" http://127.0.0.1:8000/api/scan-photo
```

`<task_id>` 替换为上一条命令返回的 `task_id`。`-OJ` 会尝试使用服务端文件名；如果本机 curl 未识别 `filename*`，文件会保存为 `download.*`，可用 `-o 结果.txt` 指定名字。翻译命令在没有 `.env` 密钥时返回 `400 未配置翻译密钥`，属于预期行为。

## 降级与已知限制

- **未安装 Office**：`docx → pdf` 自动使用 ReportLab 重新排版，文字可选中，但字体、行距、分页与原文件可能不同；页面会提示当前为降级模式。
- **Office COM 串行执行**：同一时间只允许一个 Word/Excel/PowerPoint 实例执行转换，避免 COM 冲突；Office 超时后会被强制结束并清理残留进程。
- **不提供 OCR**：扫描版 PDF 没有文字层时，翻译会明确提示提取不到文字；扫描件增强只改善图片质量，不识别文字。需要文字时请先用外部 OCR 工具识别。
- **PDF 翻译按文字块提取**：多栏排版、文本框、脚注可能按坐标顺序重排；图片中的文字不会翻译。DOCX 只读取段落与表格，页眉页脚、文本框不参与翻译。
- **保留版式是覆盖而非重排**：译文直接写回原段落矩形，字号最小 6pt；超出矩形且无法再缩时会截断并加省略号，不保证与原文逐像素一致。
- **翻译依赖外部接口**：耗时与费用由你配置的服务和模型决定；本版不内置任何在线转换或免费翻译服务。
- **中文字体子集缺字时改用系统字体**：先尝试项目内置的 Noto Sans SC 子集，缺字时回退到 Windows 微软雅黑；两者都没有的字形会明确报错，不会输出方框。
- **PDF → DOCX 是推断版式**：表格、文本框、多栏可能错位，不保证与原文像素级一致。
- **服务仅监听 `127.0.0.1`**：没有登录、权限与 HTTPS，请不要把端口暴露到局域网或公网。
- **任务存在内存中**：重启服务后任务记录消失；结果默认保留 1 小时，过期自动清理。
- **仅 Windows**：Office COM 与 `pywin32` 绑定 Windows；其他系统只能使用纯 Python 链路（且当前以 Windows 为主要支持平台）。
- **Office 专属格式**：`.doc → pdf/docx` 需要 Word，`.ppt/.pptx → pdf` 与 `.pptx → png` 需要 PowerPoint。未安装时这些任务会失败并给出明确提示，不会返回 500。
- **批量中的 Office 任务串行**：同一时刻仍只允许一个 Word/Excel/PowerPoint 实例；纯 Python 任务仍按 4 线程并行。
- **批量 ZIP 会扁平化目录**：所有成功结果放在 ZIP 根目录，重名自动加 `(2)`、`(3)`；不还原原始文件夹层级。
- **PDF 压缩会丢文字层**：压缩结果是逐页 JPEG 重新封装的 PDF，不能再搜索、复制或选中文字；纯文字 PDF 可能变得更大。
- **PDF 预览只渲染前 50 页**：长文档的后续页面仍会参与保存与导出，不会丢页，只是编辑界面不显示缩略图。
- **逐页 JPG 有硬上限**：到 `72 DPI + q50` 仍超过设定 KB 时，该文件明确失败并给出当前体积，不会悄悄超限。
- **本版没有 PDF 批注 / 拆分接口**：页面编辑通过 `delete_pages / rotations / page_order` 组合完成；拆分可理解为白名单式 `page_order`。
- **Known limitations（汇总）**：仅支持 Windows；Office COM 不支持并发；PDF 压缩会丢文字层；服务无鉴权、无 HTTPS，仅限本机或可信局域网；任务与结果不持久化；不做 OCR。

## 常见问题

**Word 转换卡住不动，或结束后 WINWORD.EXE 还在**

Office COM 转换受 `OFFICE_TIMEOUT_SECONDS` 限制（`start.bat` 演示默认 120 秒）。超时后服务会执行 `taskkill /F /IM WINWORD.EXE` 并把任务标记为失败；如果仍残留，请关闭所有 Word 窗口后在任务管理器结束 `WINWORD.EXE`，再重试。Office 任务串行执行，同时提交多个 Word/PPT 任务时后面的会排队。

**双击 `start.bat` 后窗口里没有文字，或一闪而过**

先看根目录的 `_start_log.txt`。常见原因是 Python 未安装/不是 3.11、`.venv` 创建失败或依赖安装失败。窗口底部会给出具体阶段（`[1/3]` 创建环境、`[2/3]` 安装依赖、`[3/3]` 启动）。也可以在项目根目录手动执行 `.venv\Scripts\python.exe main.py` 查看完整报错。若端口 8000–8005 全部被占用，脚本会提示并退出。

**翻译按钮是灰的，或提示未配置翻译密钥**

复制 `.env.example` 为 `.env`，填写 `TRANSLATE_API_KEY`（以及需要时的 `TRANSLATE_BASE_URL`、`TRANSLATE_MODEL`），重启服务。密钥留空时界面会显示配置说明并禁用翻译；服务端接口也会返回 `400 未配置翻译密钥…`。其他转换、压缩与 PDF 工具不需要密钥。

**双击 `python` 弹出 Microsoft Store**

这是 Windows 的应用执行别名占位程序。安装 python.org 的 3.11 并勾选 “Add python.exe to PATH”，或使用 `py -3.11` 创建虚拟环境；`start.bat` 启动服务时始终使用 `.venv\Scripts\python.exe`，不依赖 PATH 中的 `python`。

## 项目结构

```
main.py            FastAPI 应用、上传校验、任务与下载
converter.py       转换注册表、全部转换实现与压缩
pdf_tools.py       PDF 合并、页面操作、预览与逐页 JPG
translator.py      翻译分段、并发调用、对照 / 纯译文 / 版式覆盖
scanner.py         扫描件透视矫正、裁边、去阴影与对比度增强
office_worker.py   独立子进程执行 Word/Excel/PPT COM
assets/fonts/      中文 Noto Sans SC 子集（SIL OFL 1.1，许可见 OFL.txt）
glossary.txt       术语表（英文=中文，界面自动预填）
static/            原生 HTML/CSS/JS 前端（无框架、无构建）
tests/smoke.py     自启动服务的端到端冒烟测试
tests/check_frontend.py    前端静态自检（括号、id、资源、标签）
tests/test_page_semantics.py  页码语义测试（删页/复制/旋转）
samples/           可直接拖入的示例文件
uploads/ outputs/  运行时目录（已忽略，不提交）
docs/prompts/      需求、决策与阶段文档
docs/DEMO.md       5 分钟演示脚本与翻车兜底
docs/QA.md         答辩问答与代码证据
docs/AI-USAGE.md   Qoder 使用方式、规则与真实纠偏案例
requirements.txt   Python 依赖
start.bat          Windows 一键启动（自动选端口、写 _start_log.txt）
CHANGELOG.md       分阶段变更记录
LICENSE            MIT 许可证
```

## 测试

```bat
.venv\Scripts\python.exe tests\check_frontend.py
.venv\Scripts\python.exe tests\test_page_semantics.py
.venv\Scripts\python.exe tests\smoke.py
```

三个脚本都可单独运行，退出码 `0` 表示通过、`1` 表示失败：

- `check_frontend.py`：纯静态检查 `static/` 的 JS/CSS 括号配平、HTML 标签闭合、`$()` / `getElementById` 引用的 id 是否存在、HTML 与 JS 引用的 `/static/` 资源是否存在。
- `test_page_semantics.py`：构造 5 页 PDF，验证删第 2 页后共 4 页、`page_order` 重复页码得到复制页、旋转 90° 后 `page.rect` 宽高互换、删除与旋转同时提交的顺序。
- `smoke.py`：在随机端口自行启动服务（关闭 Office），并启动一个本地假模型接口，用真实 HTTP 上传完成 56 项检查，覆盖：
  - 健康检查、配置上限、翻译配置与术语表、格式注册表、首页静态资源与阶段 5 前端标记；
  - DOCX → TXT/HTML/MD/PDF、PDF → TXT/DOCX/JPG/PNG、XLSX → CSV/单表与多表、CSV → XLSX、XLS → XLSX/CSV、Markdown/TXT → PDF、图片互转与图片 → Word；
  - Office 不可用时 `.doc`/`.pptx` 的明确错误；
  - 批量 3 种格式的 ZIP 与单文件下载、坏文件失败隔离、ZIP 重名加序号、批量数量与参数校验；
  - 415 扩展名、400 不支持的组合、404、中文下载头与 CSV 的 UTF-8 BOM；
  - PDF 预览（PNG base64、原 rotation、50 页截断）、300 页 PDF 只预览 50 页且导出后 300 页完整保留、加密 PDF 400；
  - PDF 删页 + 旋转、保留原 rotation、`page_order` 白名单与重复页、越界页码与空参数报错；
  - 合并三种模式的实际页序、四档压缩预估、PDF 文字层丢失、图片压缩、非 PDF/图片 400；
  - 逐页 JPG 全部 ≤300KB，以及上限无法达成时的明确失败；
  - 翻译三种输出模式的成品结构、术语表生效、进度文案与百分比；保留版式拒绝 DOCX 返回明确 400；未配置密钥时 400；
  - 扫描件增强对倾斜 + 阴影照片完成四边形检测与透视矫正，非图片返回 415。

浏览器验收：浅色/深色/跟随系统三态主题切换与刷新后保持、600px 窄屏布局、PDF 删页/复制后 `Ctrl+Z` 撤销与 `Ctrl+Shift+Z` 重做、连续操作超过 50 步时只保留最近 50 步、300 页 PDF 预览提示、超过 50MB 文件提示、错误详情折叠展开均已实际操作，控制台无报错。未配置密钥时翻译区域置灰并显示配置说明。

## 阶段规划

| 阶段 | 内容 | 状态 |
| --- | --- | --- |
| 1 | 骨架、异步任务、上传校验、6 类主干转换、界面、冒烟测试 | 已完成 |
| 2 | 完整格式矩阵、批量与文件夹、失败清单与重试 | 已完成 |
| 3 | 压缩、PDF 合并/删页/旋转/重排、缩略图编辑器、逐页 JPG | 已完成 |
| 4 | 翻译（双语/仅译文/版式叠加）、术语表、扫描增强 | 已完成 |
| 5 | 三态主题、响应式、撤销/重做、空/加载/错误状态、后台清理、启动兜底、三项测试与交付文档 | 已完成 |

## 隐私

- 文件只保存在本机 `uploads/` 与 `outputs/`，转换完成后源上传文件即删除，结果按 TTL 清理。
- 不写日志记录文件内容；不记录 API Key。
- 除你主动配置并使用的翻译接口外，应用不发起任何外部网络请求。