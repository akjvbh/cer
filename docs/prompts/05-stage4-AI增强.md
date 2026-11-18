# 阶段 4：AI 增强（文献翻译 + 扫描件增强）

## 当前状态

阶段 1-3 已完成：转换、批量、压缩、PDF 工具可用，异步任务与进度轮询已经跑通。前端有 5 个标签，其中「AI 增强」还是占位。

## 本阶段目标

1. 文献翻译：PDF / DOCX / TXT 逐段翻译，输出中英对照 DOCX、纯译文 DOCX、保留原版式的 PDF
2. 扫描件增强：手机拍的照片 → 找纸张四角 → 裁切矫正 → 去阴影 → 输出干净扫描件
3. 两者都走异步任务 + 真实进度；没配置密钥时界面明确提示，不报错崩溃

## 要做的事

### 1. `translator.py`

**段落抽取**（三处实现，都要保留原顺序）：

- PDF：`page.get_text("blocks")`，按 `(page, y, x)` 排序，把同一段内被换行切开的行合并成段；过滤纯页码、纯符号、长度 < 2 的块
- DOCX：`doc.paragraphs` 逐个段落；表格按单元格遍历，保持行列顺序
- TXT：按空行分段

**分块与调用**：

- 每块 ≤ 2000 字符，按段落边界切，不要把一句话劈开
- 并发 3 个线程调用 `/chat/completions`，单个请求超时 120 秒
- 429 / 5xx 重试 3 次，指数退避（2s、4s、8s）；4xx（除 429）直接报错并把接口返回的 `error.message` 带给用户
- Prompt 要求「只输出译文，不要解释、不要加引号、不要保留原文」
- 术语表：`glossary.txt` 每行 `英文=中文`，`#` 开头是注释；拼进 system prompt；界面输入框的内容排在文件内容之后，可以覆盖

**三种输出**：

- `bilingual`：DOCX 两列表格，左原文右译文，段落一一对应
- `translated`：DOCX 纯译文，保留原文档的段落样式（标题层级尽量对应）
- `layout`：PDF 原位覆盖——对每个文本块 `add_redact_annot` 抹掉原文，再 `insert_textbox` 用中文字体写译文；字号按原块矩形自适应，放不下就缩小到 6pt 为止，仍放不下就截断并加省略号。**必须嵌入中文字体**（用项目内 `assets/fonts/` 放一个开源中文字体，或调用系统 `msyh.ttc`/`simsun.ttc`），否则会输出方框

**进度**：总块数 N，每完成一块 `TASKS[id]["progress"] = int(done/N*100)`；message 写「正在翻译第 12/58 段」。

### 2. `scanner.py`

- 读图 → 缩放到长边 1600 → 灰度 → 高斯模糊 → Canny → `findContours` → 取面积最大的近似四边形
- 找到四角：按左上/右上/右下/左下排序，`warpPerspective` 做透视矫正
- 找不到合理四边形：不报错，退回「整图轻微裁边 + 增强」，note 写「未检测到纸张边界，已按原图增强」
- 去阴影：用大核高斯模糊估背景，`cv2.divide(gray, background, scale=255)`，再按对比度做自适应阈值或 `CLAHE`
- 输出 PNG（保留细节）或 JPG（默认 PNG），单文件
- 前端展示「原图 / 结果」前后对比

### 3. 接口

- `POST /api/translate`：`file`、`target_lang`（默认 `zh`）、`mode`、`glossary` → `{"task_id"}`
- `POST /api/scan-photo`：`file` → 任务卡
- `/api/config` 增加 `translate_configured`（检查 `TRANSLATE_API_KEY` 是否非空）和 `glossary`（读 `glossary.txt` 内容）
- 未配置密钥时 `/api/translate` 返回 400 + `{"error":"未配置翻译密钥，请在 .env 里填写 TRANSLATE_API_KEY"}`

### 4. `.env.example`

```
TRANSLATE_API_KEY=
TRANSLATE_BASE_URL=https://api.deepseek.com/v1
TRANSLATE_MODEL=deepseek-chat
MAX_FILE_SIZE_MB=200
MAX_BATCH_SIZE_MB=500
OFFICE_TIMEOUT_SECONDS=600
```

### 5. 前端「AI 增强」标签

- 上半：文献翻译——上传文件、目标语言、输出方式（对照/纯译文/保留版式）、术语表输入框（从 `/api/config` 预填）、开始按钮、进度条、完成卡片（下载）
- 下半：扫描件增强——上传图片、前后对比、下载
- 未配置密钥时，翻译区域置灰并显示「未配置密钥」+ 一句怎么配

## 硬约束

- 翻译必须**逐段对应**，不允许把整篇拼成一段再翻（会丢结构）
- 中文 PDF 必须嵌入字体，不能在结果里出现「□□□」
- 不要打印或记录 API Key（日志里脱敏成 `sk-****`）
- 接口超时、限流要有明确提示，不能让用户干等
- 扫描件增强失败时返回原图 + note，不要返回 500
- 不引入新的重型依赖；OpenCV 只用 `opencv-python-headless`，不要装 GUI 版

## 验收

```powershell
# 翻译（先确认 .env 里配了 key）
curl.exe -F "file=@samples\初稿_田螺厣盖个体识别.docx" -F "target_lang=zh" -F "mode=bilingual" http://127.0.0.1:8000/api/translate
# 扫描件增强
curl.exe -F "file=@samples\photo.jpg" http://127.0.0.1:8000/api/scan-photo
```

- 找一个英文 PDF 或 DOCX，三种模式各跑一次：对照 DOCX 打开后左右两列能对上；`layout` 模式 PDF 里中文正常显示、没有方框、原文没有重叠残留
- 扫描件：拿手机拍一张有阴影、倾斜的纸，处理后的图边缘拉直、阴影基本消失
- 翻译过程中进度条真实推进（不是 0 → 100）
- `.\.venv\Scripts\python.exe tests\smoke.py` 全 PASS（新增：未配置密钥时的 400、术语表读取）
- `git commit -m "stage 4: AI 增强"`

## 完成后告诉我

- 段落抽取在 PDF 上遇到的实际情况（有没有顺序错乱、表格丢失），怎么处理的
- `layout` 模式字号自适应的规则
- 翻译一份 10 页 PDF 的大致耗时和成本估算
