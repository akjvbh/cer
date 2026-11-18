# API 契约

前后端之间的硬约定。**改接口必须同时改三处**：`main.py` 的路由、`static/app.js` 的调用、`README.md` 的 API 表。三处不一致视为 bug。

## 接口清单

| 方法 | 路径 | 入参 | 返回 |
|---|---|---|---|
| GET | `/api/health` | — | `{"status":"ok","office":true,"python":"3.11","version":"0.2.0"}` |
| GET | `/api/config` | — | `{"max_file_size_mb":200,"max_batch_files":200,"translate_configured":true,"glossary":""}` |
| GET | `/api/formats` | — | `{"docx":["pdf","txt"],"pdf":["docx","png"],...}` |
| POST | `/api/convert` | `file`, `target` | 任务卡 |
| POST | `/api/compress` | `file`, `level`（`80`/`60`/`40`/`20`） | 任务卡 |
| POST | `/api/compress/estimate` | `file` | `{"original_size":123,"levels":{"80":100,"60":80,"40":60,"20":40}}` |
| POST | `/api/batch` | `files[]`, `mode`（`convert`/`compress`）, `target` 或 `level` | `{"task_id"}` |
| POST | `/api/pdf/merge` | `files[]`, `mode`（`append`/`interleave`/`insert`）, `insert_index` | 任务卡 |
| POST | `/api/pdf/ops` | `file`, `delete_pages`, `rotations`, `page_order` | 任务卡 |
| POST | `/api/pdf/preview` | `file`, `limit`（默认 50） | `{"pages":[{"index":1,"width":595,"height":842,"rotation":0,"png_base64":"..."}],"truncated":false}` |
| POST | `/api/pdf/to-jpg` | `files[]`, `max_kb`（默认 300） | 任务卡（ZIP） |
| POST | `/api/translate` | `file`, `target_lang`, `mode`（`bilingual`/`translated`/`layout`）, `glossary` | `{"task_id"}` |
| POST | `/api/scan-photo` | `file` | 任务卡 |
| GET | `/api/tasks/{task_id}` | — | `{"status":"pending\|running\|done\|error","progress":0-100,"message":"","result":{...}\|null,"error":""}` |
| GET | `/api/download/{task_id}` | — | 文件流 |

「任务卡」= `{"task_id","download_url","filename","size","note"}`。

## 任务模型（只有一套，不要再造 `batch_id`）

- 所有可能超过 3 秒的操作都立刻返回 `{"task_id"}`，前端轮询 `GET /api/tasks/{task_id}`
- 批量任务同样只返回 `task_id`；`result` 里是 `{"total_files","success_files","failed_files","results":[{"filename","ok","download_url","error","note"}]}`
- 下载统一 `GET /api/download/{task_id}`：单文件任务返回该文件，批量任务返回 ZIP
- 不要出现第二套 `/api/batch/{id}` 轮询路径

## 页码与坐标

- **所有页码都是 1-based，且指原文件的页码**，不是删除后的新页码
- `delete_pages`：`[2,5]`，1-based
- `rotations`：`{"1": 90}`，传的是**最终角度**，直接覆盖该页原有 rotation，不是叠加；只取 `0/90/180/270`
- `page_order`：非空时即白名单——没列进去的页丢弃，同一页码重复写 = 复制该页；此时 `delete_pages` 忽略
- `page_order` 为空数组或未传：**视为没传**，`delete_pages` 正常生效
- `insert_index`（合并）：0-based，`0` = 插到首页之前，`1` = 插到第 1 页之后
- 处理顺序固定为：删页/重排 → 旋转
- **本版没有 PDF 标注接口**。唯一用到坐标的是翻译的「保留版式 PDF」覆盖层：页码 1-based，坐标原点在页面左上角，单位 PDF 点（1pt = 1/72 英寸），写入时机是删页/旋转之后的最终页面

## 限制与错误码

- 单文件上限 200MB（`MAX_FILE_SIZE_MB`），批量总量 500MB、最多 200 个文件
- 超限 413；扩展名不在白名单 415；加密 / 需要密码的 PDF 返回 400 `{"error":"不支持加密 PDF"}`
- 不提供 OCR：扫描版 PDF 提取不到文字时返回明确说明，而不是空文件
- 下载中文文件名：`Content-Disposition: attachment; filename*=UTF-8''<urlencode(文件名)>`

## 格式键名

统一用不带点的小写扩展名：`pdf` `docx` `doc` `xlsx` `xls` `pptx` `ppt` `txt` `md` `csv` `html` `png` `jpg` `jpeg` `webp` `bmp` `gif` `tiff`。

`/api/formats` 由转换器注册表自动生成，禁止手写第二份映射表。
