# 阶段 2：完整格式矩阵 + 批量处理

## 当前状态

阶段 1 已完成：服务能启动，`/api/convert` 能处理 `docx→pdf`、`docx→txt`、`pdf→docx`、`pdf→txt`、`pdf→jpg`、`图片→pdf`，前端能拖拽上传和下载。`tests/smoke.py` 已有 5 个检查项。

## 本阶段目标

1. 把转换矩阵补全，`/api/formats` 自动反映全部真实可用的能力
2. 加批量接口与前端批量列表（逐行状态、失败清单、只重试失败项、结果打包下载）
3. 让「失败」变成一等公民：每个文件失败都有可读原因，整体不因单个文件失败而中断

## 要做的事

### 1. 补全 `converter.py` 的 handler

| 源 | 目标 | 实现要点 |
|---|---|---|
| docx | html | 用 `mammoth` 或自己遍历 python-docx 生成简单 HTML；表格、标题、列表要保留 |
| docx | md | 同上，输出 Markdown |
| xlsx | csv | openpyxl 逐表导出；多 sheet 时打包 ZIP |
| xlsx | pdf | 优先 Excel COM；不可用时把表格渲染成 HTML 再 `insert_htmlbox` |
| xls | xlsx | 需要 `xlrd` 时再引入，先检查依赖，不要硬写死 |
| pptx | pdf | 优先 PowerPoint COM；不可用时提示「需要本机 PowerPoint」 |
| pptx | png | 有 COM 就用导出；没有就明确报不支持 |
| csv | xlsx | openpyxl，注意编码探测（utf-8-sig → gbk） |
| md | html | `markdown` 库，`extensions=["tables","fenced_code"]` |
| md | pdf | 先 md→html，再用 `insert_htmlbox` 排版成 PDF |
| txt | docx / pdf | docx 用 python-docx；pdf 用 `insert_htmlbox`，按段落分行 |
| pdf | png | 同 pdf→jpg，PNG 无损 |
| 图片互转 | png/jpg/webp/bmp/tiff | Pillow，带透明通道时 PNG 保留 alpha，转 JPG 时先填白底 |
| 图片 | docx | 嵌入到页面宽度自适应 |

要求：

- 每个 handler 都返回 `(路径, note)`，note 说明降级或特殊行为
- 不支持的组合不要硬编码进注册表；注册表里有什么，`/api/formats` 就返回什么
- 所有 handler 都要能被单独调用测试（不要只通过 HTTP 才能验证）

### 2. 批量接口 `POST /api/batch`

- 入参：`files[]`、`mode`（`convert` | `compress`）、`target` 或 `level`
- 用 `ThreadPoolExecutor(max_workers=4)` 并发；Office 相关任务进全局串行队列
- 每个文件独立 try/except，单个失败不影响其它
- 返回 `{"batch_id","total_files","success_files","failed_files","results":[{"filename","ok","download_url","error","note"}]}`
- 结果打包：`GET /api/batch/{batch_id}/zip` 返回所有成功的文件，ZIP 内重名自动加序号
- 批量上限读 `MAX_BATCH_SIZE_MB`（默认 500MB）

### 3. 前端

- 「格式转换」标签下支持多选和整文件夹拖入
- 文件列表每行：文件名、大小、状态（等待/处理中/成功/失败）、操作（下载/重试）
- 底部汇总条：`共 N 个，成功 X，失败 Y`，失败时出现「复制失败清单」按钮（复制成 `文件名\t错误原因` 的文本）
- 「只重试失败项」按钮
- 所有文件成功后显示「打包下载 ZIP」
- 单文件仍走 `/api/convert`，多文件自动走 `/api/batch`，不要出现两个入口

### 4. 真进度

- 批量任务返回后立即给每个文件一个 `pending` 状态
- 前端逐文件轮询（或后端提供一个 `GET /api/batch/{batch_id}` 返回整体进度）——选一种，写进代码注释
- 不要用假进度条

### 5. 扩充 `tests/smoke.py`

新增：

- 批量转换 3 个不同格式的文件，断言 `success_files == 3`
- 批量里故意混入 1 个不支持转换的文件，断言 `failed_files == 1` 且 `results` 里对应项 `ok == false` 且有 `error`
- `GET /api/batch/{batch_id}/zip` 返回 `application/zip` 且 size > 0

## 硬约束

- 一个文件失败不能让整批 500
- 批量并发时 Office 任务必须串行（全局锁），不要出现多个 WINWORD
- 前端不用任何框架；失败清单用 `textContent` + 剪贴板 API
- 新增依赖必须写进 `requirements.txt` 并说明用途（例如 `mammoth` 用于 docx→html）

## 验收

```powershell
# 批量
curl.exe -F "files=@samples\初稿_田螺厣盖个体识别.docx" -F "files=@samples\参考论文_红鳍东方鲀.pdf" -F "mode=convert" -F "target=txt" http://127.0.0.1:8000/api/batch
# 打包下载（把返回的 batch_id 填进去）
curl.exe -o out.zip http://127.0.0.1:8000/api/batch/<batch_id>/zip
```

- 浏览器里一次选 3 个文件，能看到逐行状态变化，故意放一个 `.exe` 进去，失败行显示原因，点「复制失败清单」能粘贴出内容
- `.\.venv\Scripts\python.exe tests\smoke.py` 全 PASS
- `git commit -m "stage 2: 格式矩阵与批量处理"`

## 完成后告诉我

- 新增了哪些格式组合，哪些做不到、原因是什么
- 批量并发模型（线程数、Office 串行是怎么保证的）
- 实际跑过的验证与输出
