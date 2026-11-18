# 阶段 3：压缩 + PDF 工具

## 当前状态

阶段 1、2 已完成：转换矩阵完整，批量与失败清单可用，`tests/smoke.py` 覆盖转换与批量。所有耗时操作目前仍是同步返回，靠 `task_id` 占位。

## 本阶段目标

1. 四档压缩 + 压缩前预估
2. PDF 合并、页面操作（删/旋/重排/复制页）、逐页导出 JPG（每张 ≤300KB）、页面缩略图预览
3. 把耗时操作真正改成异步任务 + 轮询进度（`/api/tasks/{task_id}` 已经定义好，这次把它用起来）

## 要做的事

### 1. 压缩 `POST /api/compress`

- `level` 取 `80`/`60`/`40`/`20`，含义是质量档位（80 最轻，20 最强）
- PDF：用 PyMuPDF 每页按目标 DPI 渲染成 JPEG 再写回新 PDF
  - `80`：150 DPI / quality 85；`60`：120 DPI / 75；`40`：100 DPI / 65；`20`：80 DPI / 50
  - note 必须写「已重渲染为图片，文字不可再搜索」
- 图片：Pillow 重采样 + quality 递减，尺寸按档位上限缩放（如 80 档最长边 2400，20 档 1280）
- docx/xlsx/pptx：先尝试只重压 ZIP 内部结构、图片降采样；做不到就明确返回「该格式不支持压缩」而不是返回原文件冒充成功
- `POST /api/compress/estimate`：**抽样**前 3 页（或 3 张图）估算四档大小，返回 `{"original_size","levels":{"80":..,"60":..,"40":..,"20":..}}`，不要为了预估把整个文件跑一遍

### 2. PDF 合并 `POST /api/pdf/merge`

- `mode`：
  - `append`：按 `files[]` 顺序首尾相接
  - `interleave`：逐页交替（以最长的那份为准，短的用完后只追加长的剩余页）
  - `insert`：把第 2 份及之后的文件插进第 1 份的 `insert_index` 位置之后
- `insert_index` 是 0-based：`0` = 插到首页之前
- 合并后页数必须等于各文件页数之和（`insert` 模式也一样）
- 校验：所有文件都必须是 PDF，否则 400 并指出是哪个文件

### 3. 页面操作 `POST /api/pdf/ops`

按 API 规则里的语义实现，处理顺序固定：

1. `page_order`（白名单 + 可重复 = 复制页）→ 决定页序；为空或未传则保持原顺序
2. `delete_pages`（仅当 `page_order` 未传时生效）
3. `rotations`：绝对角度，0/90/180/270，加到该页原有旋转上，最终值归一化到 0-359
4. 返回新 PDF

要点：

- 页码越界要返回明确错误（`第 12 页不存在，该文件共 10 页`），不要静默忽略
- 复制页用 PyMuPDF 的 `insert_pdf(src, from_page, to_page)` 实现
- 空请求（三个参数都没传）直接返回 400，不要生成一个一模一样的文件

### 4. 逐页导出 JPG `POST /api/pdf/to-jpg`

- 每个 PDF 逐页渲染，每张控制在 `max_kb` 以内（默认 300）
- 策略：先 150 DPI / quality 95；超过就降 quality（95→85→75→65），仍超过就降 DPI（150→120→96）；到下限仍超就接受并记 note
- 输出 ZIP，内部结构 `原文件名/001.jpg`
- 多文件时一个 ZIP 里按文件分目录

### 5. 预览 `POST /api/pdf/preview`

- 返回前 `limit` 页（默认 50）的缩略图：`{"pages":[{"index":1,"width":595,"height":842,"png_base64":"..."}]}`
- 缩略图长边 260px 左右即可，不要渲染原尺寸（会很慢）
- 超过 50 页时返回 `"truncated": true`，前端提示「仅显示前 50 页，其余页面保存时会保留」

### 6. 异步任务

- 新增 `POST /api/tasks` 的写入路径：提交时立刻返回 `task_id`，实际转换放 `threading.Thread`
- 进度更新：写 `TASKS[task_id]["progress"]`，PDF 类按页数推进，批量按文件数推进
- 前端统一 `pollTask(taskId, onProgress, onDone)` 函数，轮询间隔 800ms，`done`/`error` 后停止

### 7. 前端「PDF 工具」标签

- 上传一个 PDF → 调 `/api/pdf/preview` → 展示页面缩略图网格
- 每页卡片：页码、选中框、旋转按钮（每次 +90）、删除按钮
- 选中的页可以拖拽调整顺序（用原生 `draggable`，不要引库）
- 顶部工具条：`全选 / 反选 / 顺时针旋转选中 / 删除选中 / 恢复原顺序 / 导出`
- 操作只改前端 state，点「导出」时一次性拼成 `page_order` + `rotations` + `delete_pages` 提交
- 缩略图网格上方显示「共 N 页，已选 M 页」

## 硬约束

- 页码 1-based，前端界面写「按原文件页码」
- 旋转是绝对角度，不要写成增量
- 大文件不能把整个 PDF 读进内存两次；用完立刻 `doc.close()`
- 压缩的 note 必须体现「丢文字层」这个事实，不能只说「压缩完成」
- 不要为了赶进度把页面操作做成「只支持前 10 页」

## 验收

```powershell
# 页面操作：删第 2 页 + 第 1 页旋转 90 度
curl.exe -F "file=@samples\参考论文_红鳍东方鲀.pdf" -F "delete_pages=[2]" -F "rotations={\"1\":90}" http://127.0.0.1:8000/api/pdf/ops
# 合并
curl.exe -F "files=@samples\参考论文_红鳍东方鲀.pdf" -F "files=@samples\参考论文_红鳍东方鲀.pdf" -F "mode=append" http://127.0.0.1:8000/api/pdf/merge
# 压缩预估
curl.exe -F "file=@samples\参考论文_红鳍东方鲀.pdf" http://127.0.0.1:8000/api/compress/estimate
```

用 Python 断言（请真实执行，把输出贴给我）：

```powershell
.\.venv\Scripts\python.exe -c "import fitz; d=fitz.open('outputs/<结果文件>'); print('页数', d.page_count); print('首页旋转', d[0].rotation)"
```

- 原文件页数 - 1，首页旋转 90，这两个数对不上就说明页码语义错了
- 浏览器里：上传 PDF，缩略图能拖拽排序，旋转后缩略图方向跟着变，导出后页数正确
- `.\.venv\Scripts\python.exe tests\smoke.py` 全 PASS（新增：preview、ops、merge、compress 四项）
- `git commit -m "stage 3: 压缩与 PDF 工具"`

## 完成后告诉我

- 页码语义你是怎么实现的（哪几行负责白名单、复制页、旋转叠加）
- 压缩四档的实际压缩率（拿 `参考论文_红鳍东方鲀.pdf` 测）
- 50 页截断的边界怎么处理的
