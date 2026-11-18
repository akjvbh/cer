# 阶段 3：压缩 + PDF 工具

## 当前状态

阶段 1、2 已完成：转换矩阵完整，批量按任务模型跑通（`POST /api/batch` → `task_id` → 轮询 → 下载 ZIP），`tests/smoke.py` 覆盖转换与批量。

## 本阶段目标

1. 四档压缩（只做 PDF 和图片）+ 压缩前预估 + 丢文字层提示
2. PDF 合并、页面操作、逐页导出 JPG、缩略图预览
3. 前端「PDF 工具」标签：缩略图网格，选中删除/旋转/拖拽排序后一次性提交

## 要做的事

### 1. 压缩（范围：PDF + 图片，别的不做）

`POST /api/compress`，`level` 取 `80`/`60`/`40`/`20`：

- **PDF**：PyMuPDF 每页按目标 DPI 渲染成 JPEG 再重新打包
  - `80`：150 DPI / quality 85；`60`：120 / 75；`40`：100 / 65；`20`：80 / 50
- **图片**：Pillow 重采样，按档位限制最长边（`80`→2400、`60`：1920、`40`：1600、`20`：1280），quality 同上
- **其它格式**（docx/xlsx/pptx）：直接返回 400 `{"error":"该格式不支持压缩，本版只支持 PDF 和图片"}`，不要返回原文件冒充成功
- **提交前提示**：前端在压缩 PDF 时弹确认框「压缩后文字会变成图片、不可再搜索，继续吗？」；结果卡片 `note` 里再写一次
- `POST /api/compress/estimate`：**抽样**前 3 页（图片取 1 张）估算四档大小，返回 `{"original_size","levels":{"80":..,"60":..,"40":..,"20":..}}`。不要为了预估把整个文件跑一遍

### 2. PDF 合并 `POST /api/pdf/merge`

- `mode`：
  - `append`：按 `files[]` 顺序首尾相接
  - `interleave`：逐页交替（以最长的那份为准，短的用完后只追加剩余页）
  - `insert`：把第 2 份及之后的文件插进第 1 份的 `insert_index` 位置之后
- `insert_index` 是 0-based：`0` = 插到首页之前。**界面上让用户填「第 N 页之后」（1-based），提交时减 1 转成 0-based**，转换发生在 JS 里，后端只认 0-based
- 合并后页数必须等于各文件页数之和
- 任一文件不是 PDF 或加密 → 400，并指出是哪个文件

### 3. 页面操作 `POST /api/pdf/ops`

语义按 API 规则，处理顺序固定：

1. `page_order` **非空**时决定页序（白名单，没列出的页丢弃，重复页码 = 复制该页）；此时忽略 `delete_pages`
2. `page_order` **为空数组或未传**时视为没传，`delete_pages` 生效
3. `rotations`：`{"1": 90}` 是**最终角度**，直接覆盖该页原 rotation（不是叠加）；范围校验只允许 0/90/180/270
4. 返回新 PDF

要点：

- 页码越界返回明确错误（`第 12 页不存在，该文件共 10 页`），不要静默忽略
- 复制页用 `insert_pdf(src, from_page, to_page)`
- 三个参数都没传 → 400，不要生成一个一模一样的文件
- 加密 PDF → 400「不支持加密 PDF」
- **本阶段不做标注/批注**，不要加 `annotations` 参数（翻译覆盖层在阶段 4 单独实现）

### 4. 预览 `POST /api/pdf/preview`

- 返回前 `limit` 页（默认 50）：`{"pages":[{"index":1,"width":595,"height":842,"rotation":0,"png_base64":"..."}],"truncated":false}`
- 缩略图 **PNG，长边 260px**，不要渲染原尺寸
- `rotation` 必须返回该页原始旋转角度，前端按这个角度初始化显示（否则原文件本身带旋转时缩略图是歪的）
- 超过 50 页时 `truncated: true`，前端提示「仅显示前 50 页，其余页面保存时原样保留」

### 5. 逐页导出 JPG `POST /api/pdf/to-jpg`

- 每个 PDF 逐页渲染，**每张必须 ≤ `max_kb`（默认 300）**
- 策略：150 DPI / q95 → 降 quality（95→85→75→65）→ 降 DPI（150→120→96→72）→ 最低 q50
- 到 72 DPI + q50 仍超 300KB：**该文件判失败**，`error` 写「第 3 页无法压到 300KB 以内（当前 412KB）」，不要无限降、也不要静默超限
- 输出 ZIP，结构按源文件分目录：`原文件名/001.jpg`
- 多文件时一个 ZIP 里按文件分目录

### 6. 异步任务

- 所有 PDF 操作立刻返回 `task_id`，处理放 `threading.Thread`
- PDF 类按页数推进进度（`已完成页 / 总页数`），批量按文件数推进
- 前端统一 `pollTask(taskId, onProgress, onDone)`，间隔 800ms，`done`/`error` 后停止

### 7. 前端「PDF 工具」标签

- 上传一个 PDF → 调 `/api/pdf/preview` → 缩略图网格
- 每页卡片：页码、选中框、旋转按钮（每次 +90）、删除按钮；缩略图按返回的 `rotation` 初始旋转
- 选中的页可拖拽调整顺序（原生 `draggable`，不要引库）
- 顶部工具条：`全选 / 反选 / 顺时针旋转选中 / 删除选中 / 恢复原顺序 / 导出`
- 操作只改前端 `state`，点导出时一次性拼成 `page_order` + `rotations`（累加后的最终角度）+ `delete_pages` 提交
- 网格上方显示「共 N 页，已选 M 页」

## 硬约束

- 页码 1-based，界面写「按原文件页码」
- 旋转是最终角度（覆盖），前端负责把用户点过的 +90 累加成最终值
- 大文件不能把整个 PDF 读进内存两次；用完立刻 `doc.close()`
- 300KB 是硬指标，达不到就报错，不要写一句 note 蒙混过去
- 不要为了赶进度只支持前 10 页

## 验收

```powershell
# 页面操作：删第 2 页 + 第 1 页转 90 度
curl.exe -F "file=@samples\参考论文_红鳍东方鲀.pdf" -F "delete_pages=[2]" -F "rotations={\"1\":90}" http://127.0.0.1:8000/api/pdf/ops
# 合并
curl.exe -F "files=@samples\参考论文_红鳍东方鲀.pdf" -F "files=@samples\参考论文_红鳍东方鲀.pdf" -F "mode=append" http://127.0.0.1:8000/api/pdf/merge
# 压缩预估
curl.exe -F "file=@samples\参考论文_红鳍东方鲀.pdf" http://127.0.0.1:8000/api/compress/estimate
```

用 Python 断言（请真实执行并贴输出）：

```powershell
.\.venv\Scripts\python.exe -c "import fitz; d=fitz.open('outputs/<结果文件>'); print('页数', d.page_count); print('首页旋转', d[0].rotation)"
```

- 原页数 − 1、首页旋转 90，两个数对不上就说明页码语义错了
- 再验一次「原文件本身带旋转」的情况：找一个 rotation ≠ 0 的 PDF，不传 rotations 原样导出，结果页的 rotation 必须保持不变
- 加密 PDF：随便给一个设了密码的 PDF，断言返回 400 且 error 含「不支持加密」
- 300KB：`GET /api/pdf/to-jpg` 后解压，逐张断言 `size <= 300*1024`
- `.\.venv\Scripts\python.exe tests\smoke.py` 全 PASS（新增 preview、ops、merge、compress、to-jpg 五项）
- `git commit -m "stage 3: 压缩与 PDF 工具"`

## 完成后告诉我

- 页码语义的实现位置：哪几行负责白名单、复制页、旋转覆盖
- 压缩四档在 `参考论文_红鳍东方鲀.pdf` 上的实际压缩率
- 300KB 那个下限在真实文件上有没有触发
