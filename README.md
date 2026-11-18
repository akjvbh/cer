# 文档工坊

一个**在本机运行的办公文档转换与处理工具箱**：拖入文件、选择目标格式，转换在本机完成，结果直接下载。不需要登录、不上传云端、不依赖 Docker 或数据库。

![Python](https://img.shields.io/badge/Python-3.11-3776AB?logo=python&logoColor=white)
![FastAPI](https://img.shields.io/badge/FastAPI-0.1x-009688?logo=fastapi&logoColor=white)
![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11-0078D4?logo=windows&logoColor=white)
![License](https://img.shields.io/badge/License-Private-lightgrey)

> 当前为 **阶段 1（骨架与主干转换）**：界面、异步任务、上传校验、6 类主干链路已可用；批量、压缩、PDF 工具、AI 增强为后续阶段。

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
| DOCX | PDF | Word COM；失败时 ReportLab 降级 | 有 Office 时保留原版式；降级排版可能有差异 |
| DOCX | TXT | python-docx | 按正文顺序输出段落与表格，UTF-8 BOM |
| PDF | DOCX | pdf2docx | 版式为程序推断，复杂排版可能与原文不同 |
| PDF | TXT | PyMuPDF | 按页提取文字层，UTF-8 BOM |
| PDF | JPG | PyMuPDF 150 DPI | 单页输出 JPG；多页打包 ZIP，逐页 `001.jpg` |
| PNG / JPG / JPEG / WEBP / BMP / GIF / TIFF | PDF | PyMuPDF + Pillow | 取首帧，按 A4 居中排版，不裁切 |
| Markdown | HTML | markdown | 支持表格、代码块、目录，生成独立带样式的 HTML |
| TXT | DOCX | python-docx | 保留空行与段落，微软雅黑 |

其他格式（`doc/xls/xlsx/pptx/ppt/csv`、图片互转、图片转 Word、压缩、PDF 合并拆分、翻译、扫描增强）将在阶段 2–4 补齐，界面中已预留入口。

## 使用方式

1. 打开页面，把文件拖进虚线区域，或点击选择文件。
2. 目标格式下拉框会根据文件类型自动列出可选项。
3. 点击「开始转换」，进度条按真实任务状态更新。
4. 完成后在「转换结果」中下载；失败会显示可复制的错误信息与重试按钮。

上传限制：默认单文件不超过 **200MB**；文件名、扩展名、文件头（magic）都会校验，服务端使用 UUID 保存，不信任任何用户提供的路径。

## 配置

复制 `.env.example` 为 `.env` 后按需修改，重启生效：

| 变量 | 默认值 | 作用 |
| --- | --- | --- |
| `PORT` | `8000` | 手动运行 `main.py` 时的监听端口 |
| `AUTO_OPEN_BROWSER` | `1` | 启动后是否自动打开浏览器 |
| `MAX_FILE_SIZE_MB` | `200` | 单文件大小上限 |
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
| `GET` | `/api/tasks/{task_id}` | 任务状态：`pending / running / done / error`、进度、结果 |
| `GET` | `/api/download/{task_id}` | 仅完成任务可下载，中文文件名用 `filename*` 传递 |

错误统一返回 `{"error": "中文说明"}`，并使用有意义的 HTTP 状态码（`400` 参数不支持、`413` 过大、`415` 文件类型不符、`404` 任务不存在、`409` 未完成）。

示例：

```bat
curl.exe -F "file=@samples\示例_论文.pdf" -F "target=txt" http://127.0.0.1:8000/api/convert
curl.exe http://127.0.0.1:8000/api/tasks/<task_id>
curl.exe -OJ http://127.0.0.1:8000/api/download/<task_id>
```

## 降级与已知限制

- **未安装 Office**：`docx → pdf` 自动使用 ReportLab 重新排版，文字可选中，但字体、行距、分页与原文件可能不同；页面会提示当前为降级模式。
- **Office COM 串行执行**：同一时间只允许一个 Word/Excel/PowerPoint 实例执行转换，避免 COM 冲突；Office 超时后会被强制结束并清理残留进程。
- **不提供 OCR**：扫描版 PDF 没有文字层时，转 TXT/DOCX 只能得到空内容，需等待阶段 4 的扫描增强能力。
- **PDF → DOCX 是推断版式**：表格、文本框、多栏可能错位，不保证与原文像素级一致。
- **服务仅监听 `127.0.0.1`**：没有登录、权限与 HTTPS，请不要把端口暴露到局域网或公网。
- **任务存在内存中**：重启服务后任务记录消失；结果默认保留 1 小时，过期自动清理。
- **仅 Windows**：Office COM 与 `pywin32` 绑定 Windows；其他系统只能使用纯 Python 链路（且当前以 Windows 为主要支持平台）。
- 阶段 1 尚未实现：批量/文件夹、压缩、PDF 页面工具、翻译、扫描增强，以及 `E4 单页 JPG ≤300KB`、`绝对角度旋转` 等后续约定。

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

测试会在随机端口自行启动服务（关闭 Office），用真实 HTTP 上传完成 15 项检查：健康检查、格式注册表、首页静态资源、DOCX→TXT（中文文件名/表格/BOM）、PDF→TXT、PDF→DOCX、PDF→JPG 多页 ZIP、图片→PDF、Markdown→HTML、TXT→DOCX、DOCX→PDF 降级、404、415、400 与中文下载头。

前端验证：桌面 1440px、移动 600px、浅色/深色主题均已人工检查；浏览器控制台无报错。

## 阶段规划

| 阶段 | 内容 | 状态 |
| --- | --- | --- |
| 1 | 骨架、异步任务、上传校验、6 类主干转换、界面、冒烟测试 | 已完成 |
| 2 | 完整格式矩阵、批量与文件夹、失败清单与重试 | 计划中 |
| 3 | 压缩、PDF 合并/拆分/旋转/重排、缩略图编辑器 | 计划中 |
| 4 | 翻译（双语/仅译文/版式叠加）、术语表、扫描增强 | 计划中 |
| 5 | 整体打磨、异常与性能验收、演示材料与交付 | 计划中 |

## 隐私

- 文件只保存在本机 `uploads/` 与 `outputs/`，转换完成后源上传文件即删除，结果按 TTL 清理。
- 不写日志记录文件内容；不记录 API Key。
- 除你主动配置并使用的翻译接口外，应用不发起任何外部网络请求。