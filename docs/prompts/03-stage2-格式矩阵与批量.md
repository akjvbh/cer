# 阶段 2：完整格式矩阵 + 批量处理

## 当前状态

阶段 1 已完成：服务能启动，`/api/convert` 能处理 `docx→pdf`、`docx→txt`、`pdf→docx`、`pdf→txt`、`pdf→jpg`、`图片→pdf`，前端能拖拽上传和下载，`tests/smoke.py` 能跑。

## 本阶段目标

1. 补全转换矩阵（含 `docx→md`），`/api/formats` 自动反映真实能力
2. 批量接口统一进任务模型：`POST /api/batch` 只返回 `task_id`，结果里带逐文件清单
3. 前端支持整文件夹、按格式分组提交、只重试失败项、结果打包下载

## 要做的事

### 1. 补全 `converter.py` 的 handler

| 源 | 目标 | 实现要点 |
|---|---|---|
| docx | html | `mammoth` 转 HTML，保留标题/列表/表格 |
| docx | md | 遍历 python-docx，标题转 `#`、列表转 `-`、表格转 Markdown 表格 |
| xlsx | csv | openpyxl 逐 sheet 导出；**多 sheet 打包 ZIP**，文件名 `<原名>_<sheet名>.csv`，编码 `utf-8-sig` |
| xlsx | pdf | 优先 Excel COM；不可用时用 HTML 表格近似排版 + `note` |
| xls | xlsx / csv | `xlrd` 读旧版；读不到就明确报错 |
| pptx | pdf / png | 优先 PowerPoint COM；不可用时**明确报错**「需要本机 PowerPoint」，不要输出残缺 PDF |
| ppt | pdf | 同上 |
| csv | xlsx | openpyxl；编码探测顺序 `utf-8-sig` → `utf-8` → `gbk` |
| md | html / pdf | `markdown` with `extensions=["tables","fenced_code"]`；pdf 走 `insert_htmlbox` |
| txt | docx / pdf | docx 用 python-docx；pdf 按段落分行排版 |
| pdf | png | 同 pdf→jpg，PNG 无损 |
| 图片互转 | png/jpg/webp/bmp/tiff | Pillow；保留 alpha 转 PNG，转 JPG 先铺白底 |
| 图片 | docx | 按页面宽度自适应嵌入 |

要求：

- 每个 handler 返回 `(path, note)`；`note` 用来写降级或特殊行为
- **多帧 GIF / 多页 TIFF 只处理第一帧/页**，并写 `note`「仅处理了第一帧」
- 注册表里有什么，`/api/formats` 就返回什么，不要手写第二份映射

### 2. 批量接口 `POST /api/batch`

- 入参：`files[]`、`mode`（`convert` | `compress`）、`target` 或 `level`
- **立刻返回 `{"task_id"}`**，实际处理放后台线程；进度写进同一个 `TASKS` 表
- 校验：最多 200 个文件、总量 500MB，超了直接 413 并说明
- 并发 `ThreadPoolExecutor(max_workers=4)`；Office 任务进全局串行队列
- 每个文件独立 try/except，单个失败不影响其它
- 任务完成后 `result` 为：

```json
{"total_files": 5, "success_files": 4, "failed_files": 1,
 "results": [{"filename": "a.docx", "ok": true, "download_url": "/api/download/xxx", "error": "", "note": ""}]}
```

- 批量下载：`GET /api/download/{task_id}` 直接返回 ZIP（所有成功的文件），ZIP 内重名自动加序号
- 不要再另立 `/api/batch/{batch_id}` 这类第二套轮询路径

### 3. 前端

- 支持多选 + 整文件夹拖入
- **混合格式分组**：拖进来的文件按源格式分组，每组一个目标格式下拉；点开始时每组各发一次 `/api/batch`，并行提交、并行轮询
- 文件列表每行：文件名、大小、状态（等待/处理中/成功/失败）、操作（下载/重试）
- 底部汇总：`共 N 个，成功 X，失败 Y`；失败时出现「复制失败清单」（格式 `文件名\t错误原因`）
- 「只重试失败项」按钮：只把失败的文件重新提交一次，成功的保留
- 全部完成后出现「打包下载 ZIP」

### 4. 真进度

- 后端按「已完成文件数 / 总数」更新 `progress`
- 前端轮询间隔 800ms，`done`/`error` 后停止；不要用假进度条

### 5. 扩充 `tests/smoke.py`

- 批量转换 3 个不同格式的文件，断言 `success_files == 3`
- 批量里混 1 个不支持转换的文件，断言 `failed_files == 1` 且对应项 `ok == false`、`error` 非空
- `GET /api/download/{task_id}` 返回 `application/zip` 且 size > 0
- 多 sheet xlsx → csv 返回 ZIP 且内含 2 个 csv
- txt/csv 输出编码是 `utf-8-sig`（读前 3 字节应为 `EF BB BF`）

## 硬约束

- 一个文件失败不能让整批 500
- 批量并发时 Office 任务必须串行（全局锁），不要出现多个 WINWORD
- 前端不用任何框架；失败清单用 `textContent` + 剪贴板 API
- 新增依赖写进 `requirements.txt` 并说明用途（`mammoth` 用于 docx→html，`xlrd` 用于读 .xls）

## 验收

```powershell
# 批量（注意：现在返回 task_id，需要再查 /api/tasks/{task_id}）
curl.exe -F "files=@samples\初稿_田螺厣盖个体识别.docx" -F "files=@samples\参考论文_红鳍东方鲀.pdf" -F "mode=convert" -F "target=txt" http://127.0.0.1:8000/api/batch
curl.exe http://127.0.0.1:8000/api/tasks/<task_id>
curl.exe -o out.zip http://127.0.0.1:8000/api/download/<task_id>
```

- 浏览器里一次拖入 3 种不同格式的文件，能分别选目标格式并分组提交；故意放一个 `.exe`，失败行显示原因，点「复制失败清单」能粘贴出内容，点「只重试失败项」只重跑那一个
- `.\.venv\Scripts\python.exe tests\smoke.py` 全 PASS
- `git commit -m "stage 2: 格式矩阵与批量处理"`

## 完成后告诉我

- 新增了哪些格式组合，哪些做不到、原因是什么
- 分组提交是怎么实现的（数据结构 + 几次请求）
- 实际跑过的验证与输出
