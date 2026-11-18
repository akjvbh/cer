# 阶段 5：打磨、文档与交付

## 当前状态

阶段 1-4 已完成：转换、批量、压缩、PDF 工具、翻译、扫描件增强全部可用。现在功能齐了但还像「能跑的 demo」，这一步要把它变成「能给人看的产品」。

## 本阶段目标

1. 体验细节：主题、响应式、撤销、空状态、错误提示
2. 稳定性：大文件、并发、端口占用、异常兜底
3. 文档：README、CHANGELOG、.env.example、LICENSE
4. 自检：测试脚本补全，跑一遍全绿再提交

## 要做的事

### 1. 界面与交互

- 主题：浅色/深色/跟随系统三态切换，存 `localStorage`，首屏不闪（在 `<head>` 里先读 localStorage 设置 `data-theme`）
- 响应式：`< 900px` 时左侧导航变顶部横向滚动标签，结果卡片单列
- 撤销：PDF 页面操作（删除、旋转、排序、复制）进同一个撤销栈，最多 50 步，`Ctrl+Z` 撤销、`Ctrl+Shift+Z` 重做，按钮置灰状态要正确
- 空状态：没选文件时每个标签都有引导文案 + 支持的格式说明，不要一片空白
- 加载态：上传/转换中按钮禁用 + 文案变化（「转换中…」），不要能重复点
- 错误提示：统一一个顶部 toast（3 秒自动消失）+ 结果卡片里的详细错误（可展开）
- 大文件提示：超过 50MB 时提示「文件较大，转换可能需要 1-2 分钟」

### 2. 稳定性

- 所有 `fitz.Document` / `PIL.Image` / `openpyxl` 工作簿用完立即关闭
- `uploads/` 和 `outputs/` 加后台清理：每小时清一次超过 1 小时的临时文件（`threading.Timer` 或启动时的守护线程）
- 端口占用：`start.bat` 检测 8000 被占用时自动用 8001，并在控制台明确打印实际地址
- `start.bat` 里加启动日志 `_start_log.txt`，方便排查起不来的原因
- 全局异常处理器：未捕获异常返回 500 + `{"error":"服务器内部错误"} `，同时把完整 traceback 写日志（不要返回给前端）
- 服务启动时打印：Python 版本、Office 是否可用、监听地址

### 3. 测试补齐

`tests/` 下至少三个脚本，都能用 `python tests/xxx.py` 单独跑，退出码 0/1：

- `smoke.py`：端到端（现有）
- `check_frontend.py`：纯静态检查 `static/`——括号配平、`getElementById` 的 id 都在 HTML 里存在、`app.js` 里引用的资源路径存在、标签闭合
- `test_page_semantics.py`：只测页码语义——构造一个 5 页 PDF，验证「删第 2 页后共 4 页」「page_order 重复页码得到复制页」「旋转 90 度后 `page.rect` 宽高互换」

### 4. 文档

- `README.md`：
  - 一句话定位 + 徽章（Python / FastAPI / Platform / License）
  - 它解决什么问题（4-5 条真实场景）
  - 功能一览表
  - 快速开始：前置条件、安装、启动（含 Windows 上 `python` 是应用商店占位程序怎么办）
  - 配置：`.env` 各字段说明
  - 支持的转换矩阵表
  - API 表（和 `main.py` 保持一致）
  - 项目结构
  - 已知限制：仅 Windows、Office COM 不支持并发、压缩丢文字层、无鉴权
  - 常见问题：Word 卡死、双击文字没反应、翻译未配置
- `CHANGELOG.md`：按阶段记录
- `.env.example`、`.gitignore`、`LICENSE`（MIT）

### 5. 提交前自检

逐条确认并贴出结果：

```powershell
.\.venv\Scripts\python.exe tests\check_frontend.py
.\.venv\Scripts\python.exe tests\test_page_semantics.py
.\.venv\Scripts\python.exe tests\smoke.py
git status --short
```

- 工作区里没有 `uploads/`、`outputs/`、`.env`、`__pycache__`
- `README.md` 里的每条命令都能照抄执行
- 换一个没有 Office 的假设（把 COM 探测强制返回 False）后，核心功能仍可用

## 硬约束

- 本阶段**不加新功能**，只修体验和稳定性；要加功能先单独说
- 不为了好看引入任何前端库或图标库
- 不改接口契约；确实要改就同步 README 和前端
- 每改完一项就跑一次 `smoke.py`，不要攒到最后

## 验收

- 窄屏（浏览器窗口拖到 600px 宽）布局不崩
- `Ctrl+Z` 能撤销刚才删掉的页，撤销栈超过 50 步后最老的被丢弃
- 300 页 PDF：预览只出 50 页并提示，导出时其余页面仍在
- 端口被占用时 `start.bat` 自动换端口并在窗口里显示新地址
- 三个测试脚本全绿
- `git commit -m "stage 5: 打磨与文档"` 并打 tag `v1.0-demo`

## 完成后告诉我

- 改动了哪些文件
- 三条测试脚本的真实输出
- 还剩下哪些已知问题（写进 README 的「已知限制」）
