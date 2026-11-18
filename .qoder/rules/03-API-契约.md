# API 契约

前后端之间的硬约定。**改接口必须同时改三处**：`main.py` 的路由、`static/app.js` 的调用、`README.md` 的 API 表。三处不一致视为 bug。

## 接口清单

| 方法 | 路径 | 入参 | 返回 |
|---|---|---|---|
| GET | `/api/health` | — | `{"status":"ok","office":true,"python":"3.11","version":"0.1.0"}` |
| GET | `/api/config` | — | `{"max_file_size_mb":200,"translate_configured":true,"glossary":""}` |
| GET | `/api/formats` | — | `{"docx":["pdf","txt"],"pdf":["docx","png"],...}` |
| POST | `/api/convert` | `file`, `target` | `{"task_id","filename","size","download_url","note"}` |
| POST | `/api/compress` | `file`, `level`（`80`/`60`/`40`/`20`） | 同上 |
| POST | `/api/compress/estimate` | `file` | `{"original_size":123,"levels":{"80":100,"60":80,"40":60,"20":40}}` |
| POST | `/api/batch` | `files[]`, `mode`（`convert`/`compress`）, `target` 或 `level` | `{"batch_id","total_files","success_files","failed_files","results":[...]}` |
| POST | `/api/pdf/merge` | `files[]`, `mode`（`append`/`interleave`/`insert`）, `insert_index` | 任务卡 |
| POST | `/api/pdf/ops` | `file`, `delete_pages`, `rotations`, `page_order` | 任务卡 |
| POST | `/api/pdf/preview` | `file`, `limit`（默认 50） | `{"pages":[{"index":1,"width":595,"height":842,"png_base64":"..."}]}` |
| POST | `/api/pdf/to-jpg` | `files[]`, `max_kb`（默认 300） | 任务卡（ZIP） |
| POST | `/api/translate` | `file`, `target_lang`, `mode`（`bilingual`/`translated`/`layout`）, `glossary` | `{"task_id"}` |
| POST | `/api/scan-photo` | `file` | 任务卡 |
| GET | `/api/tasks/{task_id}` | — | `{"status":"pending\|running\|done\|error","progress":0-100,"message":"","result":{...}\|null,"error":""}` |
| GET | `/api/download/{task_id}` | — | 文件流 |

「任务卡」= `{"task_id","download_url","filename","size","note"}`。

## 通用约定

- 全部返回 JSON；失败返回 `{"error": "中文说明"}`，HTTP 状态码 400/413/415/500 要对应准确
- 任何可能超过 3 秒的操作都返回 `task_id` 让前端轮询，不要阻塞请求直到超时
- 下载统一走 `GET /api/download/{task_id}`，不要暴露真实磁盘路径
- 下载中文文件名：`Content-Disposition: attachment; filename*=UTF-8''<urlencode(文件名)>`
- 上传大小超限返回 413，扩展名不在白名单返回 415
- 同一 `task_id` 的临时文件保留 1 小时后清理

## 页码与坐标（前后端写死，注释里也要写清）

- **所有页码都是 1-based，且指原文件的页码**，不是删除后的新页码
- `delete_pages`：`[2,5]`，1-based
- `rotations`：`{"1": 90}`，**绝对角度**，只取 `0/90/180/270`（不是增量）
- `page_order`：原页码的排列，同时是白名单——没列进去的页会被丢弃；同一页码重复写 = 复制该页；传空数组或不传 = 保持原顺序
- `delete_pages` 与 `page_order` 同时传时，以 `page_order` 为准，`delete_pages` 忽略
- `insert_index`（合并）：0-based，`0` = 插到首页之前，`1` = 插到第 1 页之后
- 标注对象：`page` 是 1-based；坐标原点在**页面左上角**，单位是 PDF 点（1pt = 1/72 英寸）
- 处理顺序固定为：删页/重排 → 旋转 → 写标注。标注坐标基于该页最终的内容框

## 格式键名

统一用不带点的小写扩展名：`pdf` `docx` `doc` `xlsx` `xls` `pptx` `ppt` `txt` `md` `csv` `html` `png` `jpg` `jpeg` `webp` `bmp` `gif` `tiff`。

`/api/formats` 由转换器注册表自动生成，禁止手写第二份映射表。
