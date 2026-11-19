"use strict";

const THEME_KEY = "docworkshop-theme";
const DEFAULT_TARGETS = {
    docx: "pdf",
    pdf: "txt",
    jpg: "pdf",
    jpeg: "pdf",
    png: "pdf",
    webp: "pdf",
    bmp: "pdf",
    gif: "pdf",
    tiff: "pdf",
    md: "html",
    txt: "docx",
};
const FORMAT_LABELS = {
    pdf: "PDF 文档",
    docx: "Word 文档",
    txt: "纯文本",
    jpg: "JPG 图片",
    jpeg: "JPEG 图片",
    png: "PNG 图片",
    webp: "WebP 图片",
    bmp: "BMP 图片",
    gif: "GIF 图片",
    tiff: "TIFF 图片",
    html: "HTML 网页",
};

const state = {
    formats: {},
    health: null,
    file: null,
    target: "",
    taskId: null,
    timer: null,
    resultCount: 0,
    maxFileMb: 200,
    busy: false,
    healthRetry: 0,
    batch: {
        items: [],
        groups: [],
        invalid: [],
        timer: null,
        nextId: 1,
        runId: 0,
        submitting: false,
        running: false,
        finished: false,
        maxFiles: 200,
        maxBatchMb: 500,
        maxFileMb: 200,
    },
    compress: {
        file: null,
        level: "60",
        busy: false,
        estimateToken: 0,
        resultCount: 0,
    },
    pdf: {
        file: null,
        total: 0,
        pages: {},
        order: [],
        selected: new Set(),
        rotations: {},
        deleted: new Set(),
        history: [],
        future: [],
        loading: false,
        exporting: false,
        dragIndex: null,
        mergeFiles: [],
        mergeNextId: 1,
        mergeBusy: false,
        jpgBusy: false,
        resultCount: {
            ops: 0,
            merge: 0,
            jpg: 0,
        },
    },
    ai: {
        translate: {
            file: null,
            mode: "bilingual",
            configured: false,
            busy: false,
        },
        scan: {
            file: null,
            busy: false,
            originalUrl: "",
        },
    },
};

const COMPRESS_LEVEL_SPECS = {
    "80": { dpi: 150, quality: 85, edge: 2400 },
    "60": { dpi: 120, quality: 75, edge: 1920 },
    "40": { dpi: 100, quality: 65, edge: 1600 },
    "20": { dpi: 80, quality: 50, edge: 1280 },
};
const COMPRESS_LEVELS = ["80", "60", "40", "20"];
const COMPRESS_IMAGE_EXTS = ["png", "jpg", "jpeg", "webp", "bmp", "gif", "tiff"];
const TRANSLATE_EXTS = ["pdf", "docx", "txt"];
const SCAN_EXTS = ["png", "jpg", "jpeg", "webp", "bmp", "gif", "tiff"];
const PDF_HISTORY_LIMIT = 50;

const $ = (id) => document.getElementById(id);

function getExtension(name) {
    const value = String(name || "");
    const index = value.lastIndexOf(".");
    return index >= 0 ? value.slice(index + 1).toLowerCase() : "";
}

function formatBytes(bytes) {
    const value = Number(bytes);
    if (!Number.isFinite(value) || value < 0) {
        return "—";
    }
    if (value < 1024) {
        return `${Math.round(value)} B`;
    }
    const units = ["KB", "MB", "GB"];
    let result = value / 1024;
    let index = 0;
    while (result >= 1024 && index < units.length - 1) {
        result /= 1024;
        index += 1;
    }
    return `${result.toFixed(result >= 10 ? 0 : 1)} ${units[index]}`;
}

function formatLabel(target) {
    return FORMAT_LABELS[target] || String(target || "").toUpperCase();
}

const LARGE_FILE_BYTES = 50 * 1024 * 1024;

function warnLargeFile(file) {
    if (file && Number(file.size) > LARGE_FILE_BYTES) {
        showToast("文件较大，转换可能需要 1-2 分钟", "info");
    }
}

function warnLargeFiles(files) {
    if ((files || []).some((file) => file && Number(file.size) > LARGE_FILE_BYTES)) {
        showToast("文件较大，转换可能需要 1-2 分钟", "info");
    }
}

function kindOf(ext) {
    if (ext === "pdf") {
        return "pdf";
    }
    if (["jpg", "jpeg", "png", "webp", "bmp", "gif", "tiff"].includes(ext)) {
        return "image";
    }
    if (["txt", "md"].includes(ext)) {
        return "text";
    }
    return "office";
}

const THEME_MODES = ["system", "light", "dark"];
const THEME_LABELS = {
    system: "主题：跟随系统（点击切到浅色）",
    light: "主题：浅色（点击切到深色）",
    dark: "主题：深色（点击切到跟随系统）",
};

function savedThemeMode() {
    let saved = null;
    try {
        saved = localStorage.getItem(THEME_KEY);
    } catch (error) {
        saved = null;
    }
    return THEME_MODES.includes(saved) ? saved : "system";
}

function updateThemeButton() {
    const button = $("theme-toggle");
    const mode = savedThemeMode();
    const label = THEME_LABELS[mode];
    button.setAttribute("aria-label", label);
    button.title = label;
    button.dataset.themeMode = mode;
}

function applyTheme(mode) {
    if (mode === "system") {
        document.documentElement.removeAttribute("data-theme");
    } else {
        document.documentElement.setAttribute("data-theme", mode);
    }
    try {
        localStorage.setItem(THEME_KEY, mode);
    } catch (error) {
        /* 隐私模式下 localStorage 不可用时，仅当前页面生效。 */
    }
    updateThemeButton();
}

function initTheme() {
    const mode = savedThemeMode();
    if (mode === "system") {
        document.documentElement.removeAttribute("data-theme");
    } else {
        document.documentElement.setAttribute("data-theme", mode);
    }
    updateThemeButton();
    const media = window.matchMedia("(prefers-color-scheme: dark)");
    if (typeof media.addEventListener === "function") {
        media.addEventListener("change", () => {
            if (savedThemeMode() === "system") {
                updateThemeButton();
            }
        });
    }
}

async function fetchJson(url, options = {}, timeoutMs = 15000) {
    const controller = new AbortController();
    const timer = window.setTimeout(() => controller.abort(), timeoutMs);
    try {
        const response = await fetch(url, { ...options, signal: controller.signal });
        let payload = null;
        try {
            payload = await response.json();
        } catch (error) {
            payload = null;
        }
        if (!response.ok) {
            const message = payload && payload.error ? payload.error : `请求失败（HTTP ${response.status}）`;
            throw new Error(message);
        }
        return payload;
    } catch (error) {
        if (error.name === "AbortError") {
            throw new Error("请求超时，请确认本地服务仍在运行");
        }
        throw error;
    } finally {
        window.clearTimeout(timer);
    }
}

async function loadHealth() {
    const badge = $("health-badge");
    const text = $("health-text");
    try {
        const data = await fetchJson("/api/health");
        state.health = data;
        badge.dataset.state = data.office ? "ok" : "warn";
        text.textContent = data.office ? "Office 可用" : "Office 未检测到";
        badge.title = `Python ${data.python} · 版本 ${data.version}`;
        updateOfficeWarning();
        updateBatchOfficeWarning();
    } catch (error) {
        badge.dataset.state = "error";
        text.textContent = "环境检测失败";
        badge.title = error.message;
        if (state.healthRetry < 1) {
            state.healthRetry += 1;
            window.setTimeout(loadHealth, 4000);
        }
    }
}

async function loadConfigAndFormats() {
    try {
        const [formats, config] = await Promise.all([
            fetchJson("/api/formats"),
            fetchJson("/api/config"),
        ]);
        state.formats = formats && typeof formats === "object" ? formats : {};
        state.maxFileMb = Number(config && config.max_file_size_mb) > 0
            ? Number(config.max_file_size_mb)
            : 200;
        state.batch.maxFileMb = state.maxFileMb;
        state.batch.maxFiles = Number(config && config.max_batch_files) > 0
            ? Number(config.max_batch_files)
            : 200;
        state.batch.maxBatchMb = Number(config && config.max_batch_size_mb) > 0
            ? Number(config.max_batch_size_mb)
            : 500;
        state.ai.translate.configured = Boolean(config && config.translate_configured);
        const glossary = config && typeof config.glossary === "string" ? config.glossary : "";
        state.ai.translate.glossary = glossary;
        const glossaryInput = $("translate-glossary");
        if (glossaryInput) {
            glossaryInput.value = glossary;
        }
        updateTranslateUi();
        const input = $("file-input");
        const extensions = Object.keys(state.formats).map((ext) => `.${ext}`).join(",");
        input.accept = extensions;
        $("batch-file-input").accept = extensions;
        $("batch-folder-input").accept = extensions;
        $("dropzone-sub").textContent = `支持 Word / PDF / 图片 / Markdown / TXT · 单文件 ≤${state.maxFileMb}MB`;
        $("batch-dropzone-sub").textContent = `按扩展名自动分组 · 单文件 ≤${state.batch.maxFileMb}MB · 最多 ${state.batch.maxFiles} 个`;
    } catch (error) {
        showToast(error.message || "无法读取格式表", "error");
        $("dropzone-sub").textContent = "服务格式表读取失败，请刷新页面重试";
    }
}

function updateOfficeWarning() {
    const warning = $("office-warning");
    const ext = state.file ? getExtension(state.file.name) : "";
    const visible = Boolean(
        state.health
        && state.health.office === false
        && ext === "docx"
        && state.target === "pdf"
    );
    warning.hidden = !visible;
}

function updateResultCount() {
    $("result-count").textContent = `${state.resultCount} 个完成`;
}

function hideEmptyStateIfNeeded() {
    const list = $("result-list");
    $("result-empty").hidden = list.children.length > 0;
}

function setBusy(busy, label = "转换中…") {
    state.busy = busy;
    const button = $("convert-btn");
    button.disabled = busy || !state.file || !state.target;
    $("convert-btn-text").textContent = busy ? label : "开始转换";
    $("target-select").disabled = busy || !state.file;
    $("remove-file").disabled = busy || !state.file;
    const dropzone = $("dropzone");
    dropzone.classList.toggle("is-disabled", busy);
    dropzone.setAttribute("aria-disabled", busy ? "true" : "false");
}

function setFile(file) {
    if (!file) {
        return;
    }
    if (state.busy) {
        showToast("转换进行中，请完成后再选择新文件", "warning");
        return;
    }
    const ext = getExtension(file.name);
    const targets = state.formats[ext];
    if (!targets || targets.length === 0) {
        showToast(`不支持 .${ext || "未知"} 文件，请查看支持的格式`, "error");
        return;
    }
    if (file.size === 0) {
        showToast("这个文件是空的，请选择有内容的文件", "error");
        return;
    }
    if (file.size > state.maxFileMb * 1024 * 1024) {
        showToast(`文件超过 ${state.maxFileMb}MB 上限`, "error");
        return;
    }

    state.file = file;
    state.target = targets.includes(DEFAULT_TARGETS[ext]) ? DEFAULT_TARGETS[ext] : targets[0];
    warnLargeFile(file);

    const icon = $("file-icon");
    icon.textContent = ext.toUpperCase().slice(0, 4);
    icon.className = `file-icon kind-${kindOf(ext)}`;
    $("file-name").textContent = file.name;
    $("file-size").textContent = `${formatBytes(file.size)} · ${ext.toUpperCase()} 文件`;
    $("selected-file").hidden = false;

    const select = $("target-select");
    const options = targets.map((target) => {
        const option = document.createElement("option");
        option.value = target;
        option.textContent = `${target.toUpperCase()} · ${formatLabel(target)}`;
        return option;
    });
    select.replaceChildren(...options);
    select.value = state.target;
    select.disabled = false;
    $("target-hint").textContent = `可转换为：${targets.map((target) => target.toUpperCase()).join(" / ")}`;
    $("convert-btn").disabled = false;
    updateOfficeWarning();
}

function clearFile() {
    if (state.busy) {
        showToast("转换进行中，请完成后再移除文件", "warning");
        return;
    }
    state.file = null;
    state.target = "";
    $("file-input").value = "";
    $("selected-file").hidden = true;
    const select = $("target-select");
    const option = document.createElement("option");
    option.textContent = "请先选择文件";
    select.replaceChildren(option);
    select.disabled = true;
    $("target-hint").textContent = "选择文件后会自动列出可转换的格式";
    $("convert-btn").disabled = true;
    updateOfficeWarning();
}

function showProgress(visible) {
    $("progress-panel").hidden = !visible;
}

function setProgress(percent, message) {
    const value = Math.max(0, Math.min(100, Math.round(Number(percent) || 0)));
    $("progress-bar").style.width = `${value}%`;
    $("progress-percent").textContent = `${value}%`;
    if (message) {
        $("progress-message").textContent = message;
    }
}

function iconSvg(name) {
    if (name === "check") {
        return '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="m5 12.5 4.2 4.2L19 7"></path></svg>';
    }
    if (name === "x") {
        return '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" aria-hidden="true"><path d="M7 7l10 10M17 7 7 17"></path></svg>';
    }
    return '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M12 4v11M8 11l4 4 4-4"></path><path d="M5 19h14"></path></svg>';
}

function appendResult(result) {
    if (!result) {
        appendError("转换完成，但服务没有返回结果信息");
        return;
    }
    const list = $("result-list");
    const card = document.createElement("article");
    card.className = "result-card success";

    const status = document.createElement("div");
    status.className = "result-status";
    status.innerHTML = iconSvg("check");

    const body = document.createElement("div");
    body.className = "result-body";

    const name = document.createElement("p");
    name.className = "result-name";
    name.textContent = result.filename || "转换结果";

    const meta = document.createElement("p");
    meta.className = "result-meta";
    const target = result.target ? String(result.target).toUpperCase() : "";
    meta.textContent = [formatBytes(result.size), target].filter(Boolean).join(" · ");

    body.append(name, meta);

    if (result.note) {
        const note = document.createElement("p");
        note.className = "result-note";
        note.textContent = result.note;
        body.append(note);
    }

    const actions = document.createElement("div");
    actions.className = "result-actions";
    const download = document.createElement("a");
    const url = typeof result.download_url === "string" && result.download_url.startsWith("/api/download/")
        ? result.download_url
        : "";
    if (!url) {
        appendError("下载地址无效，请重新转换");
        return;
    }
    download.className = "download-link";
    download.href = url;
    download.setAttribute("download", "");
    download.innerHTML = `${iconSvg("download")}<span>下载文件</span>`;
    actions.append(download);
    body.append(actions);

    card.append(status, body);
    list.append(card);
    state.resultCount += 1;
    updateResultCount();
    hideEmptyStateIfNeeded();
    card.scrollIntoView({ behavior: "smooth", block: "nearest" });
}

function errorDetails(message) {
    const details = document.createElement("details");
    details.className = "error-details";
    const summary = document.createElement("summary");
    summary.textContent = "查看详细错误";
    const note = document.createElement("p");
    note.className = "result-note";
    note.textContent = message || "处理失败，请重试";
    details.append(summary, note);
    return details;
}

function appendError(message) {
    const list = $("result-list");
    const card = document.createElement("article");
    card.className = "result-card error";

    const status = document.createElement("div");
    status.className = "result-status";
    status.innerHTML = iconSvg("x");

    const body = document.createElement("div");
    body.className = "result-body";

    const title = document.createElement("p");
    title.className = "result-name";
    title.textContent = "转换失败";

    const details = errorDetails(message);

    const actions = document.createElement("div");
    actions.className = "result-actions";
    const copy = document.createElement("button");
    copy.className = "copy-button";
    copy.type = "button";
    copy.dataset.action = "copy-error";
    copy.textContent = "复制错误信息";
    actions.append(copy);

    body.append(title, details, actions);
    card.append(status, body);
    list.append(card);
    hideEmptyStateIfNeeded();
    card.scrollIntoView({ behavior: "smooth", block: "nearest" });
}

function finishConvert() {
    setBusy(false);
    const taskId = state.taskId;
    state.taskId = null;
    window.setTimeout(() => {
        if (state.taskId === null && !state.busy) {
            showProgress(false);
        }
        if (taskId) {
            /* 保留结果卡片上的下载地址，进度条只负责本次任务。 */
        }
    }, 900);
}

function uploadFile(file, target, onProgress) {
    return new Promise((resolve, reject) => {
        const xhr = new XMLHttpRequest();
        xhr.open("POST", "/api/convert", true);
        xhr.upload.onprogress = (event) => {
            if (event.lengthComputable) {
                onProgress(Math.round((event.loaded / event.total) * 100));
            }
        };
        xhr.onload = () => {
            let payload = {};
            try {
                payload = JSON.parse(xhr.responseText || "{}");
            } catch (error) {
                payload = {};
            }
            if (xhr.status >= 200 && xhr.status < 300 && payload.task_id) {
                resolve(payload);
            } else {
                reject(new Error(payload.error || `上传失败（HTTP ${xhr.status}）`));
            }
        };
        xhr.onerror = () => reject(new Error("网络错误，无法连接本地服务"));
        xhr.onabort = () => reject(new Error("上传已取消"));
        const form = new FormData();
        form.append("file", file, file.name);
        form.append("target", target);
        xhr.send(form);
    });
}

function pollTask(taskId) {
    window.clearTimeout(state.timer);
    let failures = 0;
    const tick = async () => {
        if (state.taskId !== taskId) {
            return;
        }
        try {
            const task = await fetchJson(`/api/tasks/${taskId}`, {}, 20000);
            failures = 0;
            if (task.status === "done") {
                setProgress(100, "转换完成");
                appendResult(task.result);
                showToast("转换完成", "success");
                finishConvert();
                return;
            }
            if (task.status === "error") {
                appendError(task.error || "转换失败");
                showToast(task.error || "转换失败", "error");
                finishConvert();
                return;
            }
            const overall = 32 + Math.round((Number(task.progress) || 0) * 0.68);
            setProgress(overall, task.message || "正在转换");
            state.timer = window.setTimeout(tick, 800);
        } catch (error) {
            failures += 1;
            if (failures >= 3) {
                appendError(`无法获取任务状态：${error.message}`);
                showToast("任务状态获取失败", "error");
                finishConvert();
                return;
            }
            state.timer = window.setTimeout(tick, 1200);
        }
    };
    tick();
}

async function startConvert() {
    if (state.busy || !state.file || !state.target) {
        return;
    }
    const file = state.file;
    const target = state.target;
    setBusy(true, "上传中…");
    setProgress(0, "正在上传文件");
    showProgress(true);
    try {
        const data = await uploadFile(file, target, (percent) => {
            setProgress(Math.round(percent * 0.3), `正在上传 ${percent}%`);
        });
        state.taskId = data.task_id;
        $("convert-btn-text").textContent = "转换中…";
        setProgress(32, "已提交，正在转换");
        pollTask(data.task_id);
    } catch (error) {
        appendError(error.message || "上传失败");
        showToast(error.message || "上传失败", "error");
        finishConvert();
    }
}

function showToast(message, type = "info") {
    const stack = $("toast-stack");
    while (stack.children.length >= 3) {
        stack.firstElementChild.remove();
    }
    const toast = document.createElement("div");
    toast.className = `toast ${type}`;
    toast.textContent = message || "";
    stack.append(toast);
    window.setTimeout(() => {
        toast.classList.add("leaving");
        window.setTimeout(() => toast.remove(), 220);
    }, 2800);
}

async function copyText(text, successMessage = "错误信息已复制") {
    try {
        await navigator.clipboard.writeText(text);
        showToast(successMessage, "success");
    } catch (error) {
        const textarea = document.createElement("textarea");
        textarea.value = text;
        textarea.style.position = "fixed";
        textarea.style.opacity = "0";
        document.body.append(textarea);
        textarea.select();
        try {
            document.execCommand("copy");
            showToast(successMessage, "success");
        } catch (fallbackError) {
            showToast("复制失败，请手动选择文字", "error");
        } finally {
            textarea.remove();
        }
    }
}

function makeElement(tag, className, text) {
    const node = document.createElement(tag);
    if (className) {
        node.className = className;
    }
    if (text !== undefined) {
        node.textContent = text;
    }
    return node;
}

function safeDownloadUrl(value) {
    return typeof value === "string" && value.startsWith("/api/download/") ? value : "";
}

function batchGroupCounts(group) {
    let success = 0;
    let failed = 0;
    group.items.forEach((item) => {
        if (item.status === "success") {
            success += 1;
        } else if (item.status === "failed") {
            failed += 1;
        }
    });
    return { success, failed };
}

function batchStatusLabel(group) {
    const counts = batchGroupCounts(group);
    if (group.status === "done") {
        if (counts.failed && counts.success) {
            return "部分完成";
        }
        if (counts.failed) {
            return "全部失败";
        }
        return "已完成";
    }
    if (group.status === "error") {
        return "失败";
    }
    if (group.status === "running") {
        return "转换中";
    }
    if (group.status === "uploading") {
        return "上传中";
    }
    return "待提交";
}

function batchOverallProgress() {
    const groups = state.batch.groups;
    let total = 0;
    let weighted = 0;
    groups.forEach((group) => {
        total += group.items.length;
        weighted += group.items.length * (Number(group.progress) || 0);
    });
    return total ? Math.round(weighted / total) : 0;
}

function updateBatchProgressUI() {
    const batch = state.batch;
    const panel = $("batch-progress");
    const total = batch.groups.reduce((sum, group) => sum + group.items.length, 0);
    let progress = batchOverallProgress();
    let message = "等待开始";
    const success = batch.items.filter((item) => item.status === "success").length;
    const failed = batch.items.filter((item) => item.status === "failed").length;
    const completed = batch.groups.reduce((sum, group) => (
        sum + group.items.filter((item) => item.status === "success" || item.status === "failed").length
    ), 0);

    if (batch.submitting) {
        message = "正在上传并按组提交";
    } else if (batch.running) {
        message = total ? `已完成 ${completed}/${total} 个文件` : "正在准备";
    } else if (batch.finished) {
        progress = 100;
        message = `批量处理结束：成功 ${success}，失败 ${failed}`;
    }

    const visible = batch.items.length > 0 && (batch.running || batch.submitting || batch.finished);
    panel.hidden = !visible;
    if (!visible) {
        return;
    }
    $("batch-progress-bar").style.width = `${Math.max(0, Math.min(100, progress))}%`;
    $("batch-progress-percent").textContent = `${Math.max(0, Math.min(100, progress))}%`;
    $("batch-progress-message").textContent = message;
}

function renderBatchSummary() {
    const batch = state.batch;
    const panel = $("batch-summary");
    if (!batch.finished) {
        panel.hidden = true;
        return;
    }
    panel.hidden = false;
    const total = batch.items.length;
    const success = batch.items.filter((item) => item.status === "success").length;
    const failed = batch.items.filter((item) => item.status === "failed").length;
    const invalid = batch.invalid.length;
    const retryable = batch.items.filter((item) => item.status === "failed" && !item.invalid).length;

    $("batch-summary-title").textContent = failed
        ? `共 ${total} 个，成功 ${success}，失败 ${failed}`
        : `全部完成：共 ${total} 个文件`;
    const details = [];
    if (success) {
        details.push("每组成功结果可在下方下载 ZIP");
    }
    if (invalid) {
        details.push(`${invalid} 个文件因格式或大小未提交`);
    }
    if (failed) {
        details.push("失败原因可复制后逐条检查");
    }
    $("batch-summary-detail").textContent = details.join(" · ");

    const copyButton = $("batch-copy-failures");
    copyButton.disabled = failed === 0;
    const retryButton = $("batch-retry");
    retryButton.disabled = retryable === 0 || batch.running || batch.submitting;
    retryButton.textContent = retryable ? `重试失败项（${retryable}）` : "重试失败项";
}

function createBatchRow(item) {
    const row = makeElement("li", "batch-row");
    row.dataset.id = String(item.id);

    const ext = item.ext || "";
    const icon = makeElement("span", `row-ext kind-${kindOf(ext)}`, ext ? ext.toUpperCase().slice(0, 4) : "?");
    const main = makeElement("div", "row-main");
    const name = makeElement("p", "row-name", item.name || "未命名文件");
    name.title = item.name || "";
    const resultSuffix = item.status === "success" && item.resultName && item.resultName !== item.name
        ? ` · ${item.resultName}`
        : "";
    const meta = makeElement(
        "p",
        "row-meta",
        `${formatBytes(item.size)}${ext ? ` · ${ext.toUpperCase()}` : ""}${resultSuffix}`
    );
    main.append(name, meta);
    if (item.error) {
        main.append(makeElement("p", "row-error", item.error));
    } else if (item.note) {
        main.append(makeElement("p", "row-note", item.note));
    }

    const actions = makeElement("div", "row-actions");
    const statusMap = {
        processing: ["处理中", "processing"],
        success: ["成功", "success"],
        failed: ["失败", "failed"],
        ready: ["待处理", "ready"],
    };
    const status = statusMap[item.status] || statusMap.ready;
    const badge = makeElement("span", `row-state state-${status[1]}`, status[0]);
    actions.append(badge);

    const url = safeDownloadUrl(item.downloadUrl);
    if (item.status === "success" && url) {
        const link = document.createElement("a");
        link.className = "row-download";
        link.href = url;
        link.setAttribute("download", "");
        link.textContent = "下载";
        actions.append(link);
    }

    row.append(icon, main, actions);
    return row;
}

function createGroupCard(group) {
    const counts = batchGroupCounts(group);
    const card = makeElement("article", "batch-group");
    card.dataset.ext = group.ext;

    const head = makeElement("div", "batch-group-head");
    const title = makeElement("div", "group-title");
    const icon = makeElement("span", `file-icon kind-${kindOf(group.ext)}`, group.ext.toUpperCase().slice(0, 4));
    const titleCopy = makeElement("div");
    const heading = makeElement("h3", "", `${formatLabel(group.ext)} · ${group.items.length} 个文件`);
    const totalSize = group.items.reduce((sum, item) => sum + (Number(item.size) || 0), 0);
    const subtitle = makeElement("p", "", `${formatBytes(totalSize)} · 目标 ${group.target.toUpperCase()}`);
    titleCopy.append(heading, subtitle);
    title.append(icon, titleCopy);

    const control = makeElement("div", "group-control");
    const select = document.createElement("select");
    select.dataset.action = "group-target";
    select.dataset.ext = group.ext;
    const targets = state.formats[group.ext] || [];
    targets.forEach((target) => {
        const option = document.createElement("option");
        option.value = target;
        option.textContent = `${target.toUpperCase()} · ${formatLabel(target)}`;
        select.append(option);
    });
    select.value = group.target;
    select.disabled = Boolean(
        state.batch.running
        || state.batch.submitting
        || group.status === "uploading"
        || group.status === "running"
    );
    const badge = makeElement("span", "batch-status", batchStatusLabel(group));
    badge.dataset.state = group.status === "done" && counts.failed
        ? "partial"
        : group.status;
    control.append(select, badge);
    head.append(title, control);

    const progress = makeElement("div", "batch-group-progress");
    const progressBar = document.createElement("div");
    progressBar.style.width = `${Math.max(0, Math.min(100, Number(group.progress) || 0))}%`;
    progress.append(progressBar);

    const rows = makeElement("ul", "batch-rows");
    group.items.forEach((item) => rows.append(createBatchRow(item)));

    const footer = makeElement("div", "batch-group-foot");
    const message = makeElement("p", "group-message", group.message || (group.status === "ready" ? "等待开始" : ""));
    footer.append(message);
    if (group.zipUrl) {
        const zipLink = document.createElement("a");
        zipLink.className = "download-link";
        zipLink.href = group.zipUrl;
        zipLink.setAttribute("download", "");
        zipLink.textContent = "下载本组 ZIP";
        footer.append(zipLink);
    }
    if (counts.failed > 0) {
        const copyButton = makeElement("button", "copy-button", "复制本组失败原因");
        copyButton.type = "button";
        copyButton.dataset.action = "copy-group-failures";
        copyButton.dataset.ext = group.ext;
        footer.append(copyButton);
    }

    card.append(head, progress, rows, footer);
    return card;
}

function createInvalidCard() {
    const invalid = state.batch.invalid;
    const card = makeElement("article", "batch-group invalid");
    const head = makeElement("div", "batch-group-head");
    const title = makeElement("div", "group-title");
    const icon = makeElement("span", "file-icon", "!");
    const titleCopy = makeElement("div");
    const heading = makeElement("h3", "", `未加入处理 · ${invalid.length} 个文件`);
    const subtitle = makeElement("p", "", "格式、大小或批量总量不符合要求");
    titleCopy.append(heading, subtitle);
    title.append(icon, titleCopy);
    head.append(title);

    const rows = makeElement("ul", "batch-rows");
    invalid.forEach((item) => rows.append(createBatchRow(item)));

    const footer = makeElement("div", "batch-group-foot");
    const message = makeElement("p", "group-message", "这些文件不会提交给本地服务。");
    const copyButton = makeElement("button", "copy-button", "复制失败清单");
    copyButton.type = "button";
    copyButton.dataset.action = "copy-all-failures";
    footer.append(message, copyButton);

    card.append(head, rows, footer);
    return card;
}

function renderBatch() {
    const batch = state.batch;
    const total = batch.items.length;
    const invalidCount = batch.invalid.length;
    const validCount = total - invalidCount;
    const totalBytes = batch.items.reduce((sum, item) => sum + (Number(item.size) || 0), 0);

    $("batch-count").textContent = `${total} 个`;
    let stats = total
        ? `已选择 ${total} 个 · 可处理 ${validCount} 个 · ${formatBytes(totalBytes)}`
        : "还没有选择文件";
    if (invalidCount) {
        stats += ` · 无法处理 ${invalidCount} 个`;
    }
    $("batch-stats").textContent = stats;
    $("batch-clear").disabled = total === 0 || batch.running || batch.submitting;
    $("batch-start").disabled = batch.groups.length === 0 || batch.running || batch.submitting;
    $("batch-start-text").textContent = batch.running
        ? "处理中…"
        : (batch.finished ? "重新处理全部" : "开始批量处理");
    $("batch-pick-files").disabled = batch.running || batch.submitting;
    $("batch-pick-folder").disabled = batch.running || batch.submitting;
    $("batch-dropzone").classList.toggle("is-disabled", batch.running || batch.submitting);
    $("batch-dropzone").setAttribute("aria-disabled", batch.running || batch.submitting ? "true" : "false");
    $("batch-empty").hidden = total > 0;

    const container = $("batch-groups");
    container.replaceChildren();
    batch.groups.forEach((group) => container.append(createGroupCard(group)));
    if (invalidCount) {
        container.append(createInvalidCard());
    }

    updateBatchProgressUI();
    renderBatchSummary();
    updateBatchOfficeWarning();
}

function updateBatchOfficeWarning() {
    const warning = $("batch-office-warning");
    if (!warning) {
        return;
    }
    const health = state.health;
    const missing = Boolean(
        health
        && (health.office === false || health.excel === false || health.powerpoint === false)
    );
    warning.hidden = !missing;
}

function rebuildBatchGroups() {
    const batch = state.batch;
    const previous = new Map(batch.groups.map((group) => [group.ext, group.target]));
    const byExt = new Map();
    batch.invalid = [];

    batch.items.forEach((item) => {
        if (item.invalid) {
            batch.invalid.push(item);
            return;
        }
        const targets = state.formats[item.ext];
        if (!targets || targets.length === 0) {
            item.invalid = true;
            item.status = "failed";
            item.error = `不支持 .${item.ext || "未知"} 文件`;
            batch.invalid.push(item);
            return;
        }
        if (!byExt.has(item.ext)) {
            byExt.set(item.ext, []);
        }
        byExt.get(item.ext).push(item);
    });

    const groups = [];
    byExt.forEach((items, ext) => {
        const targets = state.formats[ext];
        let target = previous.get(ext);
        if (!target || !targets.includes(target)) {
            target = targets.includes(DEFAULT_TARGETS[ext]) ? DEFAULT_TARGETS[ext] : targets[0];
        }
        const old = batch.groups.find((group) => group.ext === ext);
        if (old) {
            old.items = items;
            old.target = target;
            groups.push(old);
        } else {
            groups.push({
                ext,
                items,
                target,
                status: "ready",
                progress: 0,
                message: "",
                taskId: null,
                zipUrl: "",
                error: "",
                pollFailures: 0,
            });
        }
    });
    batch.groups = groups;
}

function resetBatchResults() {
    const batch = state.batch;
    window.clearTimeout(batch.timer);
    batch.timer = null;
    batch.runId += 1;
    batch.submitting = false;
    batch.running = false;
    batch.finished = false;
    batch.items.forEach((item) => {
        if (item.invalid) {
            item.status = "failed";
            return;
        }
        item.status = "ready";
        item.error = "";
        item.note = "";
        item.downloadUrl = "";
        item.resultName = "";
    });
    batch.groups.forEach((group) => {
        group.status = "ready";
        group.progress = 0;
        group.message = "";
        group.taskId = null;
        group.zipUrl = "";
        group.error = "";
        group.pollFailures = 0;
    });
}

function addBatchFiles(fileList) {
    const files = Array.from(fileList || []);
    if (!files.length) {
        return;
    }
    const batch = state.batch;
    if (batch.running || batch.submitting) {
        showToast("批量处理进行中，请完成后再添加文件", "warning");
        return;
    }
    if (batch.finished) {
        resetBatchResults();
    }

    let totalBytes = batch.items
        .filter((item) => !item.invalid)
        .reduce((sum, item) => sum + (Number(item.size) || 0), 0);
    const limitBytes = batch.maxBatchMb * 1024 * 1024;
    const fileLimitBytes = batch.maxFileMb * 1024 * 1024;
    let added = 0;
    let invalid = 0;
    let truncated = 0;
    warnLargeFiles(files.filter((file) => file.size > 0 && file.size <= fileLimitBytes));

    for (const file of files) {
        if (batch.items.length >= batch.maxFiles) {
            truncated = files.length - added;
            break;
        }
        const ext = getExtension(file.name);
        const targets = state.formats[ext];
        let error = "";
        if (!targets || targets.length === 0) {
            error = `不支持 .${ext || "未知"} 文件`;
        } else if (!file.size) {
            error = "文件为空，无法处理";
        } else if (file.size > fileLimitBytes) {
            error = `文件超过 ${batch.maxFileMb}MB 上限`;
        } else if (totalBytes + file.size > limitBytes) {
            error = `批量文件总量超过 ${batch.maxBatchMb}MB 上限`;
        } else {
            totalBytes += file.size;
        }

        batch.items.push({
            id: batch.nextId,
            file,
            name: file.name,
            size: file.size,
            ext,
            status: error ? "failed" : "ready",
            invalid: Boolean(error),
            error,
            note: "",
            downloadUrl: "",
            resultName: "",
        });
        batch.nextId += 1;
        added += 1;
        if (error) {
            invalid += 1;
        }
    }

    if (!added) {
        showToast(`一次最多选择 ${batch.maxFiles} 个文件`, "warning");
        return;
    }

    rebuildBatchGroups();
    batch.finished = false;
    renderBatch();

    if (truncated > 0) {
        showToast(`一次最多 ${batch.maxFiles} 个文件，已忽略 ${truncated} 个`, "warning");
    } else if (invalid > 0) {
        showToast(`已加入 ${added} 个文件，其中 ${invalid} 个无法处理`, "warning");
    } else {
        showToast(`已加入 ${added} 个文件`, "success");
    }
}

function clearBatch() {
    const batch = state.batch;
    if (batch.running || batch.submitting) {
        showToast("批量处理进行中，请完成后再清空", "warning");
        return;
    }
    window.clearTimeout(batch.timer);
    batch.timer = null;
    batch.runId += 1;
    batch.items = [];
    batch.groups = [];
    batch.invalid = [];
    batch.nextId = 1;
    batch.finished = false;
    $("batch-file-input").value = "";
    $("batch-folder-input").value = "";
    renderBatch();
}

function setGroupTarget(ext, target) {
    const batch = state.batch;
    if (batch.running || batch.submitting) {
        return;
    }
    const group = batch.groups.find((item) => item.ext === ext);
    if (!group || group.target === target) {
        return;
    }
    group.target = target;
    if (group.status === "done" || group.status === "error") {
        group.status = "ready";
        group.progress = 0;
        group.message = "";
        group.taskId = null;
        group.zipUrl = "";
        group.error = "";
        group.pollFailures = 0;
        group.items.forEach((item) => {
            item.status = "ready";
            item.error = "";
            item.note = "";
            item.downloadUrl = "";
            item.resultName = "";
        });
        batch.finished = false;
    }
    renderBatch();
}

async function postBatchGroup(group) {
    const form = new FormData();
    group.items.forEach((item) => form.append("files", item.file, item.name));
    form.append("mode", "convert");
    form.append("target", group.target);

    let response;
    try {
        response = await fetch("/api/batch", { method: "POST", body: form });
    } catch (error) {
        throw new Error("网络错误，无法连接本地服务");
    }
    let payload = {};
    try {
        payload = await response.json();
    } catch (error) {
        payload = {};
    }
    if (!response.ok || !payload.task_id) {
        throw new Error(payload.error || `提交失败（HTTP ${response.status}）`);
    }
    return payload;
}

function failBatchGroup(group, message) {
    group.status = "error";
    group.progress = 100;
    group.message = message;
    group.error = message;
    group.zipUrl = "";
    group.items.forEach((item) => {
        item.status = "failed";
        item.error = message || "处理失败";
        item.note = "";
        item.downloadUrl = "";
    });
}

function applyBatchTask(group, task) {
    if (task.status === "done") {
        const payload = task.result && typeof task.result === "object" ? task.result : {};
        const results = Array.isArray(payload.results) ? payload.results : [];
        group.items.forEach((item, index) => {
            const result = results[index];
            if (!result) {
                item.status = "failed";
                item.error = "服务没有返回该文件的结果";
                item.note = "";
                item.downloadUrl = "";
                return;
            }
            const url = result.ok ? safeDownloadUrl(result.download_url) : "";
            if (result.ok && url) {
                item.status = "success";
                item.error = "";
                item.note = result.note || "";
                item.downloadUrl = url;
                item.resultName = result.filename || "";
            } else {
                item.status = "failed";
                item.error = result.error || (result.ok ? "结果下载地址无效" : "转换失败");
                item.note = "";
                item.downloadUrl = "";
                item.resultName = "";
            }
        });
        group.status = "done";
        group.progress = 100;
        group.message = task.message || "本组处理完成";
        group.taskId = group.taskId || task.task_id;
        const counts = batchGroupCounts(group);
        group.zipUrl = counts.success > 0 ? `/api/download/${group.taskId}` : "";
        return;
    }
    if (task.status === "error") {
        failBatchGroup(group, task.error || "批量处理失败");
        return;
    }
    group.status = "running";
    group.progress = Math.max(2, Math.min(99, Number(task.progress) || 0));
    group.message = task.message || "正在转换";
}

async function pollBatchGroup(group) {
    try {
        const task = await fetchJson(`/api/tasks/${group.taskId}`, {}, 20000);
        group.pollFailures = 0;
        applyBatchTask(group, task);
    } catch (error) {
        group.pollFailures = (group.pollFailures || 0) + 1;
        if (group.pollFailures >= 3) {
            failBatchGroup(group, `无法获取任务状态：${error.message}`);
        }
    }
}

function finishBatchIfIdle() {
    const batch = state.batch;
    if (batch.submitting) {
        return;
    }
    if (batch.groups.some((group) => group.status === "running" || group.status === "uploading")) {
        return;
    }
    if (batch.finished) {
        return;
    }
    batch.running = false;
    batch.finished = true;
    renderBatch();
    const success = batch.items.filter((item) => item.status === "success").length;
    const failed = batch.items.filter((item) => item.status === "failed").length;
    showToast(
        failed ? `批量处理完成：成功 ${success}，失败 ${failed}` : `批量处理完成：${success} 个文件全部成功`,
        failed ? "warning" : "success"
    );
}

function scheduleBatchPoll(runId) {
    const batch = state.batch;
    window.clearTimeout(batch.timer);
    const tick = async () => {
        if (runId !== batch.runId) {
            return;
        }
        const active = batch.groups.filter((group) => group.status === "running" && group.taskId);
        if (!active.length) {
            finishBatchIfIdle();
            return;
        }
        await Promise.all(active.map((group) => pollBatchGroup(group)));
        if (runId !== batch.runId) {
            return;
        }
        renderBatch();
        if (batch.groups.some((group) => group.status === "running" && group.taskId)) {
            batch.timer = window.setTimeout(tick, 800);
        } else {
            finishBatchIfIdle();
        }
    };
    tick();
}

async function startBatch() {
    const batch = state.batch;
    if (batch.running || batch.submitting) {
        return;
    }
    if (!batch.groups.length) {
        showToast("没有可提交的文件，请先选择文件", "warning");
        return;
    }
    const totalBytes = batch.groups.reduce((sum, group) => (
        sum + group.items.reduce((groupSum, item) => groupSum + (Number(item.size) || 0), 0)
    ), 0);
    if (totalBytes > batch.maxBatchMb * 1024 * 1024) {
        showToast(`批量文件总量超过 ${batch.maxBatchMb}MB 上限`, "error");
        return;
    }

    resetBatchResults();
    const runId = batch.runId;
    batch.submitting = true;
    batch.running = true;
    batch.finished = false;
    batch.groups.forEach((group) => {
        group.status = "uploading";
        group.progress = 2;
        group.message = "正在上传";
        group.error = "";
        group.zipUrl = "";
        group.pollFailures = 0;
        group.items.forEach((item) => {
            item.status = "processing";
            item.error = "";
            item.note = "";
            item.downloadUrl = "";
            item.resultName = "";
        });
    });
    renderBatch();

    await Promise.all(batch.groups.map(async (group) => {
        try {
            const data = await postBatchGroup(group);
            if (runId !== batch.runId) {
                return;
            }
            group.taskId = data.task_id;
            group.status = "running";
            group.progress = 4;
            group.message = "已提交，等待处理";
        } catch (error) {
            if (runId !== batch.runId) {
                return;
            }
            failBatchGroup(group, error.message || "批量提交失败");
        }
    }));

    if (runId !== batch.runId) {
        return;
    }
    batch.submitting = false;
    renderBatch();
    if (batch.groups.some((group) => group.status === "running" && group.taskId)) {
        scheduleBatchPoll(runId);
    } else {
        finishBatchIfIdle();
    }
}

function retryBatchFailed() {
    const batch = state.batch;
    if (batch.running || batch.submitting) {
        showToast("批量处理进行中，请稍后再试", "warning");
        return;
    }
    const retryable = batch.items.filter((item) => item.status === "failed" && !item.invalid);
    if (!retryable.length) {
        showToast("没有可重试的失败文件", "info");
        return;
    }
    const previousTargets = new Map(batch.groups.map((group) => [group.ext, group.target]));
    window.clearTimeout(batch.timer);
    batch.timer = null;
    batch.runId += 1;
    batch.items = retryable.map((item) => ({
        ...item,
        status: "ready",
        error: "",
        note: "",
        downloadUrl: "",
        resultName: "",
    }));
    batch.groups = [];
    batch.invalid = [];
    batch.finished = false;
    rebuildBatchGroups();
    batch.groups.forEach((group) => {
        if (previousTargets.has(group.ext)) {
            group.target = previousTargets.get(group.ext);
        }
    });
    renderBatch();
    showToast(`正在重试 ${retryable.length} 个失败文件`, "info");
    startBatch();
}

function batchFailureText(ext) {
    return state.batch.items
        .filter((item) => item.status === "failed" && (!ext || item.ext === ext))
        .map((item) => `${item.name}\t${item.error || "未知原因"}`)
        .join("\n");
}

function readDroppedEntry(entry, output, limit) {
    return new Promise((resolve) => {
        if (!entry || output.length >= limit) {
            resolve();
            return;
        }
        if (entry.isFile) {
            if (typeof entry.file !== "function") {
                resolve();
                return;
            }
            entry.file(
                (file) => {
                    if (output.length < limit) {
                        output.push(file);
                    }
                    resolve();
                },
                () => resolve()
            );
            return;
        }
        if (!entry.isDirectory || typeof entry.createReader !== "function") {
            resolve();
            return;
        }
        const reader = entry.createReader();
        const readBatch = () => {
            reader.readEntries(
                async (entries) => {
                    if (!entries.length || output.length >= limit) {
                        resolve();
                        return;
                    }
                    for (const child of entries) {
                        if (output.length >= limit) {
                            break;
                        }
                        await readDroppedEntry(child, output, limit);
                    }
                    readBatch();
                },
                () => resolve()
            );
        };
        readBatch();
    });
}

async function collectDroppedFiles(dataTransfer) {
    const items = dataTransfer && dataTransfer.items ? Array.from(dataTransfer.items) : [];
    const entries = items
        .map((item) => (
            item.kind === "file" && typeof item.webkitGetAsEntry === "function"
                ? item.webkitGetAsEntry()
                : null
        ))
        .filter(Boolean);
    if (entries.length) {
        const output = [];
        const scanLimit = Math.min(1000, Math.max(200, state.batch.maxFiles * 5));
        for (const entry of entries) {
            if (output.length >= scanLimit) {
                break;
            }
            await readDroppedEntry(entry, output, scanLimit);
        }
        if (output.length) {
            return output;
        }
    }
    return dataTransfer && dataTransfer.files ? Array.from(dataTransfer.files) : [];
}

function bindBatchEvents() {
    const batchDrop = $("batch-dropzone");
    batchDrop.addEventListener("click", (event) => {
        if (event.target.closest("button")) {
            return;
        }
        if (state.batch.running || state.batch.submitting) {
            return;
        }
        $("batch-file-input").click();
    });
    batchDrop.addEventListener("keydown", (event) => {
        if (event.key === "Enter" || event.key === " ") {
            event.preventDefault();
            batchDrop.click();
        }
    });

    $("batch-pick-files").addEventListener("click", () => {
        $("batch-file-input").click();
    });
    $("batch-pick-folder").addEventListener("click", () => {
        $("batch-folder-input").click();
    });
    $("batch-file-input").addEventListener("change", (event) => {
        const files = event.target.files ? Array.from(event.target.files) : [];
        event.target.value = "";
        addBatchFiles(files);
    });
    $("batch-folder-input").addEventListener("change", (event) => {
        const files = event.target.files ? Array.from(event.target.files) : [];
        event.target.value = "";
        addBatchFiles(files);
    });

    ["dragenter", "dragover"].forEach((name) => {
        batchDrop.addEventListener(name, (event) => {
            event.preventDefault();
            event.stopPropagation();
            if (!state.batch.running && !state.batch.submitting) {
                batchDrop.classList.add("dragover");
            }
        });
    });
    ["dragleave", "drop"].forEach((name) => {
        batchDrop.addEventListener(name, (event) => {
            event.preventDefault();
            event.stopPropagation();
            batchDrop.classList.remove("dragover");
        });
    });
    batchDrop.addEventListener("drop", async (event) => {
        if (state.batch.running || state.batch.submitting) {
            showToast("批量处理进行中，请稍候", "warning");
            return;
        }
        const files = await collectDroppedFiles(event.dataTransfer);
        addBatchFiles(files);
    });

    $("batch-clear").addEventListener("click", clearBatch);
    $("batch-start").addEventListener("click", startBatch);
    $("batch-retry").addEventListener("click", retryBatchFailed);
    $("batch-copy-failures").addEventListener("click", () => {
        const text = batchFailureText("");
        if (!text) {
            showToast("没有失败文件需要复制", "info");
            return;
        }
        copyText(text, "失败清单已复制");
    });

    $("batch-groups").addEventListener("change", (event) => {
        const select = event.target.closest('[data-action="group-target"]');
        if (!select) {
            return;
        }
        setGroupTarget(select.dataset.ext, select.value);
    });
    $("batch-groups").addEventListener("click", (event) => {
        const button = event.target.closest('[data-action="copy-group-failures"], [data-action="copy-all-failures"]');
        if (!button) {
            return;
        }
        const ext = button.dataset.action === "copy-all-failures" ? "" : button.dataset.ext;
        const text = batchFailureText(ext);
        if (!text) {
            showToast("没有失败文件需要复制", "info");
            return;
        }
        copyText(text, "失败原因已复制");
    });
}

/* ---------------------------------------------------------------------------
   通用任务轮询与结果卡片
   --------------------------------------------------------------------------- */

function watchTask(taskId, handlers = {}, interval = 800) {
    let stopped = false;
    let failures = 0;
    let timer = null;
    const stop = () => {
        stopped = true;
        window.clearTimeout(timer);
    };
    const tick = async () => {
        if (stopped) {
            return;
        }
        try {
            const task = await fetchJson(`/api/tasks/${taskId}`, {}, 20000);
            failures = 0;
            if (task.status === "done") {
                stop();
                if (handlers.onDone) {
                    handlers.onDone(task);
                }
                return;
            }
            if (task.status === "error") {
                stop();
                if (handlers.onError) {
                    handlers.onError(task.error || "任务处理失败");
                }
                return;
            }
            if (handlers.onProgress) {
                handlers.onProgress(task);
            }
            timer = window.setTimeout(tick, interval);
        } catch (error) {
            failures += 1;
            if (failures >= 3) {
                stop();
                if (handlers.onError) {
                    handlers.onError(`无法获取任务状态：${error.message}`);
                }
                return;
            }
            timer = window.setTimeout(tick, 1200);
        }
    };
    tick();
    return stop;
}

function showTaskProgress(prefix, visible) {
    const panel = $(`${prefix}-progress`);
    if (panel) {
        panel.hidden = !visible;
    }
}

function setTaskProgress(prefix, percent, message) {
    const value = Math.max(0, Math.min(100, Math.round(Number(percent) || 0)));
    const bar = $(`${prefix}-progress-bar`);
    const numeric = $(`${prefix}-progress-percent`);
    const label = $(`${prefix}-progress-message`);
    if (bar) {
        bar.style.width = `${value}%`;
    }
    if (numeric) {
        numeric.textContent = `${value}%`;
    }
    if (label && message) {
        label.textContent = message;
    }
}

function uploadTaskRequest(url, form, onProgress) {
    return new Promise((resolve, reject) => {
        const xhr = new XMLHttpRequest();
        xhr.open("POST", url, true);
        xhr.upload.onprogress = (event) => {
            if (event.lengthComputable && onProgress) {
                onProgress(Math.round((event.loaded / event.total) * 100));
            }
        };
        xhr.onload = () => {
            let payload = {};
            try {
                payload = JSON.parse(xhr.responseText || "{}");
            } catch (error) {
                payload = {};
            }
            if (xhr.status >= 200 && xhr.status < 300 && payload.task_id) {
                resolve(payload);
            } else {
                reject(new Error(payload.error || `上传失败（HTTP ${xhr.status}）`));
            }
        };
        xhr.onerror = () => reject(new Error("网络错误，无法连接本地服务"));
        xhr.onabort = () => reject(new Error("上传已取消"));
        xhr.send(form);
    });
}

function appendOperationResult(listId, result) {
    const list = $(listId);
    if (!list) {
        return;
    }
    if (!result) {
        appendOperationError(listId, "处理完成，但服务没有返回结果信息");
        return;
    }
    const card = document.createElement("article");
    card.className = "result-card success";

    const status = document.createElement("div");
    status.className = "result-status";
    status.innerHTML = iconSvg("check");

    const body = document.createElement("div");
    body.className = "result-body";

    const name = document.createElement("p");
    name.className = "result-name";
    name.textContent = result.filename || "处理结果";

    const meta = document.createElement("p");
    meta.className = "result-meta";
    const target = result.target ? String(result.target).toUpperCase() : "";
    const counts = Number(result.total_files) > 0
        ? `成功 ${Number(result.success_files) || 0} / 失败 ${Number(result.failed_files) || 0}`
        : "";
    meta.textContent = [formatBytes(result.size), target, counts].filter(Boolean).join(" · ");
    body.append(name, meta);

    if (result.note) {
        const note = document.createElement("p");
        note.className = "result-note";
        note.textContent = result.note;
        body.append(note);
    }

    const failures = Array.isArray(result.results)
        ? result.results.filter((item) => item && item.ok === false)
        : [];
    if (failures.length) {
        const failureList = document.createElement("div");
        failureList.className = "result-note";
        failures.forEach((item) => {
            const line = document.createElement("p");
            line.textContent = `${item.filename || "文件"}：${item.error || "处理失败"}`;
            failureList.append(line);
        });
        body.append(failureList);
    }

    const actions = document.createElement("div");
    actions.className = "result-actions";
    const url = safeDownloadUrl(result.download_url);
    if (url) {
        const download = document.createElement("a");
        download.className = "download-link";
        download.href = url;
        download.setAttribute("download", "");
        download.innerHTML = `${iconSvg("download")}<span>下载结果</span>`;
        actions.append(download);
    }
    if (failures.length) {
        const copy = document.createElement("button");
        copy.className = "copy-button";
        copy.type = "button";
        copy.dataset.action = "copy-error";
        copy.textContent = "复制失败原因";
        actions.append(copy);
    }
    if (actions.children.length) {
        body.append(actions);
    }

    card.append(status, body);
    list.append(card);
    card.scrollIntoView({ behavior: "smooth", block: "nearest" });
}

function appendOperationError(listId, message, title = "处理失败") {
    const list = $(listId);
    if (!list) {
        return;
    }
    const card = document.createElement("article");
    card.className = "result-card error";

    const status = document.createElement("div");
    status.className = "result-status";
    status.innerHTML = iconSvg("x");

    const body = document.createElement("div");
    body.className = "result-body";

    const heading = document.createElement("p");
    heading.className = "result-name";
    heading.textContent = title;

    const details = errorDetails(message);

    const actions = document.createElement("div");
    actions.className = "result-actions";
    const copy = document.createElement("button");
    copy.className = "copy-button";
    copy.type = "button";
    copy.dataset.action = "copy-error";
    copy.textContent = "复制错误信息";
    actions.append(copy);

    body.append(heading, details, actions);
    card.append(status, body);
    list.append(card);
    card.scrollIntoView({ behavior: "smooth", block: "nearest" });
}

/* ---------------------------------------------------------------------------
   压缩
   --------------------------------------------------------------------------- */

function isPdfFile(file) {
    return Boolean(file) && getExtension(file.name) === "pdf";
}

function isCompressibleFile(file) {
    if (!file) {
        return false;
    }
    const ext = getExtension(file.name);
    return ext === "pdf" || COMPRESS_IMAGE_EXTS.includes(ext);
}

function compressLevelSpec(level) {
    return COMPRESS_LEVEL_SPECS[String(level)] || COMPRESS_LEVEL_SPECS["60"];
}

function updateCompressLevelUi() {
    document.querySelectorAll("#compress-levels .level-card").forEach((button) => {
        button.classList.toggle("is-active", button.dataset.level === state.compress.level);
    });
    const spec = compressLevelSpec(state.compress.level);
    const ext = state.compress.file ? getExtension(state.compress.file.name) : "";
    if (ext === "pdf") {
        $("compress-level-hint").textContent =
            `PDF：${spec.dpi} DPI / quality ${spec.quality}；文字层会在压缩后丢失。`;
    } else if (ext === "png") {
        $("compress-level-hint").textContent =
            `PNG：最长边限制 ${spec.edge} 像素；PNG 为无损格式，档位主要影响尺寸。`;
    } else {
        $("compress-level-hint").textContent =
            `图片：最长边限制 ${spec.edge} 像素 / quality ${spec.quality}。`;
    }
}

function updateCompressUi() {
    const file = state.compress.file;
    const button = $("compress-btn");
    button.disabled = state.compress.busy || !file;
    $("compress-btn-text").textContent = state.compress.busy ? "压缩中…" : "开始压缩";
    $("compress-remove").disabled = state.compress.busy || !file;
    $("compress-dropzone").classList.toggle("is-disabled", state.compress.busy);
    $("compress-warning").hidden = !(file && isPdfFile(file));
    $("compress-dropzone").setAttribute("aria-disabled", state.compress.busy ? "true" : "false");
    updateCompressLevelUi();
}

function setCompressFile(file) {
    if (!file || state.compress.busy) {
        return;
    }
    if (!isCompressibleFile(file)) {
        showToast("该格式不支持压缩，本版只支持 PDF 和图片", "error");
        return;
    }
    if (file.size === 0) {
        showToast("这个文件是空的，请选择有内容的文件", "error");
        return;
    }
    if (file.size > state.maxFileMb * 1024 * 1024) {
        showToast(`文件超过 ${state.maxFileMb}MB 上限`, "error");
        return;
    }
    const ext = getExtension(file.name);
    state.compress.file = file;
    const icon = $("compress-file-icon");
    icon.textContent = ext.toUpperCase().slice(0, 4);
    icon.className = `file-icon kind-${ext === "pdf" ? "pdf" : "image"}`;
    $("compress-file-name").textContent = file.name;
    $("compress-file-size").textContent = `${formatBytes(file.size)} · ${ext.toUpperCase()} 文件`;
    $("compress-selected").hidden = false;
    warnLargeFile(file);
    updateCompressUi();
    estimateCompression(file);
}

function clearCompressFile() {
    if (state.compress.busy) {
        showToast("压缩进行中，请完成后再移除文件", "warning");
        return;
    }
    state.compress.file = null;
    state.compress.estimateToken += 1;
    $("compress-file-input").value = "";
    $("compress-selected").hidden = true;
    $("compress-estimate").hidden = true;
    $("estimate-rows").replaceChildren();
    updateCompressUi();
}

function renderEstimate(data) {
    const card = $("compress-estimate");
    const rows = $("estimate-rows");
    const original = Number(data && data.original_size) || 0;
    const levels = data && data.levels ? data.levels : {};
    const values = COMPRESS_LEVELS
        .map((level) => Number(levels[level]) || 0)
        .filter((value) => value > 0);
    const baseline = Math.max(original, ...values, 1);

    const makeRow = (label, value, ratioText, extraClass = "") => {
        const row = document.createElement("div");
        row.className = `estimate-row ${extraClass}`.trim();
        const level = document.createElement("span");
        level.className = "estimate-level";
        level.textContent = label;
        const size = document.createElement("span");
        size.className = "estimate-size";
        size.textContent = formatBytes(value);
        const track = document.createElement("div");
        track.className = "estimate-track";
        const bar = document.createElement("span");
        bar.style.width = `${Math.max(2, Math.min(100, Math.round((value / baseline) * 100)))}%`;
        track.append(bar);
        const ratio = document.createElement("span");
        ratio.className = "estimate-ratio";
        ratio.textContent = ratioText;
        row.append(level, size, track, ratio);
        return row;
    };

    const nodes = [makeRow("原文件", original, "100%", "original")];
    COMPRESS_LEVELS.forEach((level) => {
        const value = Number(levels[level]) || 0;
        const ratio = original > 0 ? `${Math.round((value / original) * 100)}%` : "—";
        nodes.push(makeRow(level, value, ratio));
    });
    rows.replaceChildren(...nodes);

    const sampled = Number(data && data.sampled_pages) || 0;
    const ext = state.compress.file ? getExtension(state.compress.file.name) : "";
    $("estimate-note").textContent = ext === "pdf"
        ? `按前 ${sampled} 页抽样估算，实际以完整压缩为准`
        : "按图片实际重采样计算";
    card.hidden = false;
}

async function estimateCompression(file) {
    const token = state.compress.estimateToken + 1;
    state.compress.estimateToken = token;
    $("compress-estimate").hidden = false;
    $("estimate-note").textContent = "正在估算…";
    const placeholder = document.createElement("p");
    placeholder.className = "form-hint";
    placeholder.textContent = "正在抽样计算四档体积，请稍候…";
    $("estimate-rows").replaceChildren(placeholder);

    const form = new FormData();
    form.append("file", file, file.name);
    try {
        const data = await fetchJson("/api/compress/estimate", { method: "POST", body: form }, 120000);
        if (state.compress.estimateToken !== token) {
            return;
        }
        renderEstimate(data);
    } catch (error) {
        if (state.compress.estimateToken !== token) {
            return;
        }
        const note = document.createElement("p");
        note.className = "form-hint";
        note.textContent = `体积预估失败：${error.message}`;
        $("estimate-rows").replaceChildren(note);
        $("estimate-note").textContent = "预估不可用";
    }
}

function askCompressConfirm() {
    const dialog = $("compress-confirm-dialog");
    if (!dialog || typeof dialog.showModal !== "function") {
        return Promise.resolve(true);
    }
    return new Promise((resolve) => {
        const controller = new AbortController();
        const finish = (value) => {
            controller.abort();
            if (dialog.open) {
                dialog.close();
            }
            resolve(value);
        };
        const options = { signal: controller.signal };
        $("compress-confirm-ok").addEventListener("click", () => finish(true), options);
        $("compress-confirm-cancel").addEventListener("click", () => finish(false), options);
        dialog.addEventListener("cancel", () => finish(false), options);
        dialog.showModal();
    });
}

async function startCompress() {
    const file = state.compress.file;
    if (!file || state.compress.busy) {
        return;
    }
    if (isPdfFile(file) && !(await askCompressConfirm())) {
        return;
    }
    state.compress.busy = true;
    updateCompressUi();
    showTaskProgress("compress", true);
    setTaskProgress("compress", 0, "正在上传文件");

    const form = new FormData();
    form.append("file", file, file.name);
    form.append("level", state.compress.level);
    try {
        const data = await uploadTaskRequest("/api/compress", form, (percent) => {
            setTaskProgress("compress", Math.round(percent * 0.3), `正在上传 ${percent}%`);
        });
        setTaskProgress("compress", 32, "已提交，正在压缩");
        watchTask(data.task_id, {
            onProgress: (task) => {
                setTaskProgress("compress", 32 + Math.round((Number(task.progress) || 0) * 0.68), task.message || "正在压缩");
            },
            onDone: (task) => {
                state.compress.busy = false;
                state.compress.resultCount += 1;
                $("compress-result-tag").textContent = `${state.compress.resultCount} 个`;
                appendOperationResult("compress-result-list", task.result);
                $("compress-empty").hidden = true;
                setTaskProgress("compress", 100, "压缩完成");
                showToast("压缩完成", "success");
                updateCompressUi();
                window.setTimeout(() => showTaskProgress("compress", false), 800);
            },
            onError: (message) => {
                state.compress.busy = false;
                appendOperationError("compress-result-list", message, "压缩失败");
                $("compress-empty").hidden = true;
                showToast(message || "压缩失败", "error");
                updateCompressUi();
                showTaskProgress("compress", false);
            },
        });
    } catch (error) {
        state.compress.busy = false;
        appendOperationError("compress-result-list", error.message || "上传失败", "压缩失败");
        showToast(error.message || "上传失败", "error");
        updateCompressUi();
        showTaskProgress("compress", false);
    }
}

/* ---------------------------------------------------------------------------
   PDF 页面编辑
   --------------------------------------------------------------------------- */

function pdfPreviewIndices() {
    return state.pdf.order.filter((index) => Boolean(state.pdf.pages[index]));
}

function pdfSnapshot() {
    return {
        order: state.pdf.order.slice(),
        rotations: { ...state.pdf.rotations },
        deleted: Array.from(state.pdf.deleted),
        selected: Array.from(state.pdf.selected),
    };
}

function pdfRestore(snapshot) {
    state.pdf.order = snapshot.order.slice();
    state.pdf.rotations = { ...snapshot.rotations };
    state.pdf.deleted = new Set(snapshot.deleted);
    state.pdf.selected = new Set(snapshot.selected);
}

function pushPdfHistory() {
    state.pdf.history.push(pdfSnapshot());
    if (state.pdf.history.length > PDF_HISTORY_LIMIT) {
        state.pdf.history.shift();
    }
    state.pdf.future = [];
}

function undoPdf() {
    if (state.pdf.loading || state.pdf.exporting) {
        return;
    }
    if (!state.pdf.history.length) {
        showToast("没有可撤销的操作", "info");
        return;
    }
    state.pdf.future.push(pdfSnapshot());
    if (state.pdf.future.length > PDF_HISTORY_LIMIT) {
        state.pdf.future.shift();
    }
    pdfRestore(state.pdf.history.pop());
    renderPdfGrid();
}

function redoPdf() {
    if (state.pdf.loading || state.pdf.exporting) {
        return;
    }
    if (!state.pdf.future.length) {
        showToast("没有可重做的操作", "info");
        return;
    }
    state.pdf.history.push(pdfSnapshot());
    if (state.pdf.history.length > PDF_HISTORY_LIMIT) {
        state.pdf.history.shift();
    }
    pdfRestore(state.pdf.future.pop());
    renderPdfGrid();
}

function updatePdfButtons() {
    const busy = state.pdf.loading || state.pdf.exporting || state.pdf.jpgBusy;
    [
        "pdf-select-all",
        "pdf-invert",
        "pdf-copy-selected",
        "pdf-rotate-selected",
        "pdf-delete-selected",
        "pdf-reset-order",
        "pdf-export",
    ].forEach((id) => {
        const button = $(id);
        if (button) {
            button.disabled = busy;
        }
    });
    $("pdf-undo").disabled = busy || !state.pdf.history.length;
    $("pdf-redo").disabled = busy || !state.pdf.future.length;
    $("pdf-export").textContent = state.pdf.exporting ? "处理中…" : "应用并下载";
    $("pdf-jpg-btn").disabled = busy || !state.pdf.file;
    $("pdf-jpg-btn-text").textContent = state.pdf.jpgBusy ? "导出中…" : "导出为 JPG";
    $("pdf-remove").disabled = busy || !state.pdf.file;
    $("pdf-dropzone").classList.toggle("is-disabled", busy);
}

function updatePdfStats() {
    const total = state.pdf.total;
    const remaining = state.pdf.order.length;
    const selected = state.pdf.selected.size;
    const extra = remaining !== total ? `（原文件 ${total} 页）` : "";
    $("pdf-stats").textContent = `共 ${remaining} 页${extra}，已选 ${selected} 页`;
    $("pdf-page-tag").textContent = state.pdf.file ? `共 ${total} 页` : "未加载";
}

function renderPdfGrid() {
    const grid = $("pdf-grid");
    const visible = pdfPreviewIndices();
    if (!visible.length) {
        const note = document.createElement("p");
        note.className = "pdf-empty-note";
        note.textContent = "所有页面都已删除，可以点「恢复原顺序」找回。";
        grid.replaceChildren(note);
        updatePdfStats();
        updatePdfButtons();
        return;
    }

    const nodes = visible.map((index) => {
        const meta = state.pdf.pages[index];
        const offset = Number(state.pdf.rotations[index]) || 0;
        const finalRotation = ((Number(meta.rotation) || 0) + offset + 360) % 360;
        const card = document.createElement("article");
        card.className = "pdf-page-card";
        card.dataset.index = String(index);
        card.draggable = state.pdf.selected.has(index);
        if (state.pdf.selected.has(index)) {
            card.classList.add("is-selected");
        }

        const thumbWrap = document.createElement("div");
        thumbWrap.className = "pdf-thumb-wrap";
        const image = document.createElement("img");
        image.className = "pdf-thumb";
        image.alt = `第 ${index} 页缩略图`;
        image.src = `data:image/png;base64,${meta.png_base64}`;
        image.style.transform = `rotate(${finalRotation}deg)`;
        thumbWrap.append(image);
        if (finalRotation !== 0) {
            const tag = document.createElement("span");
            tag.className = "pdf-rotation-tag";
            tag.textContent = `${finalRotation}°`;
            thumbWrap.append(tag);
        }

        const foot = document.createElement("div");
        foot.className = "pdf-page-foot";
        const label = document.createElement("label");
        label.className = "pdf-page-check";
        const checkbox = document.createElement("input");
        checkbox.type = "checkbox";
        checkbox.checked = state.pdf.selected.has(index);
        checkbox.dataset.action = "toggle-page";
        checkbox.dataset.index = String(index);
        label.append(checkbox, document.createTextNode(`第 ${index} 页`));

        const actions = document.createElement("div");
        actions.className = "pdf-page-actions";
        const rotate = document.createElement("button");
        rotate.type = "button";
        rotate.dataset.action = "rotate-page";
        rotate.dataset.index = String(index);
        rotate.title = "顺时针旋转 90°";
        rotate.textContent = "↻";
        const remove = document.createElement("button");
        remove.type = "button";
        remove.dataset.action = "delete-page";
        remove.dataset.index = String(index);
        remove.title = "删除这一页";
        remove.textContent = "×";
        actions.append(rotate, remove);
        foot.append(label, actions);

        card.append(thumbWrap, foot);
        return card;
    });
    grid.replaceChildren(...nodes);
    updatePdfStats();
    updatePdfButtons();
}

async function loadPdfForEditing(file) {
    if (!file || state.pdf.loading || state.pdf.exporting) {
        return;
    }
    if (!isPdfFile(file)) {
        showToast("PDF 工具只接受 PDF 文件", "error");
        return;
    }
    if (file.size === 0) {
        showToast("这个文件是空的，请选择有内容的文件", "error");
        return;
    }
    if (file.size > state.maxFileMb * 1024 * 1024) {
        showToast(`文件超过 ${state.maxFileMb}MB 上限`, "error");
        return;
    }

    state.pdf.loading = true;
    state.pdf.file = file;
    state.pdf.pages = {};
    state.pdf.order = [];
    state.pdf.selected = new Set();
    state.pdf.rotations = {};
    state.pdf.deleted = new Set();
    state.pdf.history = [];
    state.pdf.future = [];
    warnLargeFile(file);
    $("pdf-file-name").textContent = file.name;
    $("pdf-file-size").textContent = `${formatBytes(file.size)} · PDF 文件`;
    $("pdf-selected").hidden = false;
    $("pdf-empty").hidden = true;
    $("pdf-editor").hidden = true;
    $("pdf-page-tag").textContent = "正在加载";
    showTaskProgress("pdf", true);
    setTaskProgress("pdf", 10, "正在生成缩略图");
    updatePdfButtons();

    const form = new FormData();
    form.append("file", file, file.name);
    form.append("limit", "50");
    try {
        const data = await fetchJson("/api/pdf/preview", { method: "POST", body: form }, 180000);
        const pages = Array.isArray(data && data.pages) ? data.pages : [];
        if (!pages.length) {
            throw new Error("这个 PDF 没有可预览的页面");
        }
        state.pdf.pages = {};
        pages.forEach((page) => {
            state.pdf.pages[Number(page.index)] = page;
        });
        state.pdf.total = Number(data.total) > 0 ? Number(data.total) : pages.length;
        state.pdf.order = Array.from({ length: state.pdf.total }, (_, index) => index + 1);
        state.pdf.selected = new Set();
        state.pdf.rotations = {};
        state.pdf.deleted = new Set();
        state.pdf.history = [];
        state.pdf.future = [];
        $("pdf-editor").hidden = false;
        $("pdf-truncated").hidden = !data.truncated;
        renderPdfGrid();
        setTaskProgress("pdf", 100, "缩略图已生成");
        window.setTimeout(() => showTaskProgress("pdf", false), 600);
    } catch (error) {
        state.pdf.loading = false;
        clearPdfFile();
        showToast(error.message || "PDF 预览失败", "error");
    } finally {
        state.pdf.loading = false;
        updatePdfButtons();
    }
}

function clearPdfFile() {
    if (state.pdf.loading || state.pdf.exporting) {
        return;
    }
    state.pdf.file = null;
    state.pdf.pages = {};
    state.pdf.total = 0;
    state.pdf.order = [];
    state.pdf.selected = new Set();
    state.pdf.rotations = {};
    state.pdf.deleted = new Set();
    state.pdf.history = [];
    state.pdf.future = [];
    $("pdf-file-input").value = "";
    $("pdf-selected").hidden = true;
    $("pdf-editor").hidden = true;
    $("pdf-empty").hidden = false;
    $("pdf-grid").replaceChildren();
    $("pdf-result-list").replaceChildren();
    $("pdf-page-tag").textContent = "未加载";
    showTaskProgress("pdf", false);
    updatePdfStats();
    updatePdfButtons();
}

function rotatePdfPage(index) {
    const current = Number(state.pdf.rotations[index]) || 0;
    pushPdfHistory();
    state.pdf.rotations[index] = (current + 90) % 360;
    renderPdfGrid();
}

function deletePdfPage(index) {
    const position = state.pdf.order.indexOf(index);
    if (position < 0) {
        return;
    }
    pushPdfHistory();
    state.pdf.order.splice(position, 1);
    state.pdf.selected.delete(index);
    state.pdf.deleted.add(index);
    renderPdfGrid();
}

function togglePdfSelection(index, selected) {
    if (selected) {
        state.pdf.selected.add(index);
    } else {
        state.pdf.selected.delete(index);
    }
    const card = $("pdf-grid").querySelector(`.pdf-page-card[data-index="${index}"]`);
    if (card) {
        card.classList.toggle("is-selected", selected);
        card.draggable = selected;
    }
    updatePdfStats();
}

function selectAllPdfPages() {
    pdfPreviewIndices().forEach((index) => state.pdf.selected.add(index));
    renderPdfGrid();
}

function invertPdfSelection() {
    pdfPreviewIndices().forEach((index) => {
        if (state.pdf.selected.has(index)) {
            state.pdf.selected.delete(index);
        } else {
            state.pdf.selected.add(index);
        }
    });
    renderPdfGrid();
}

function rotateSelectedPdfPages() {
    if (!state.pdf.selected.size) {
        showToast("请先勾选要旋转的页面", "info");
        return;
    }
    pushPdfHistory();
    state.pdf.selected.forEach((index) => {
        const current = Number(state.pdf.rotations[index]) || 0;
        state.pdf.rotations[index] = (current + 90) % 360;
    });
    renderPdfGrid();
}

function copySelectedPdfPages() {
    if (!state.pdf.selected.size) {
        showToast("请先勾选要复制的页面", "info");
        return;
    }
    pushPdfHistory();
    const copied = [];
    state.pdf.order.forEach((index) => {
        copied.push(index);
        if (state.pdf.selected.has(index)) {
            copied.push(index);
        }
    });
    state.pdf.order = copied;
    renderPdfGrid();
    showToast(`已复制 ${state.pdf.selected.size} 页，可继续调整顺序`, "success");
}

function deleteSelectedPdfPages() {
    if (!state.pdf.selected.size) {
        showToast("请先勾选要删除的页面", "info");
        return;
    }
    pushPdfHistory();
    state.pdf.selected.forEach((index) => {
        const position = state.pdf.order.indexOf(index);
        if (position >= 0) {
            state.pdf.order.splice(position, 1);
        }
        state.pdf.deleted.add(index);
    });
    state.pdf.selected = new Set();
    renderPdfGrid();
}

function resetPdfOrder() {
    pushPdfHistory();
    state.pdf.order = Array.from({ length: state.pdf.total }, (_, index) => index + 1);
    state.pdf.deleted = new Set();
    state.pdf.selected = new Set(
        Array.from(state.pdf.selected).filter((index) => Boolean(state.pdf.pages[index]))
    );
    renderPdfGrid();
}

function reorderPdfPages(fromIndex, toIndex) {
    const order = state.pdf.order;
    const from = order.indexOf(fromIndex);
    const to = order.indexOf(toIndex);
    if (from < 0 || to < 0 || from === to) {
        return;
    }
    pushPdfHistory();
    order.splice(from, 1);
    const insert = from < to ? to - 1 : to;
    order.splice(insert, 0, fromIndex);
    renderPdfGrid();
}

async function exportPdfOps() {
    if (!state.pdf.file || state.pdf.exporting || state.pdf.loading) {
        return;
    }
    const rotationPayload = {};
    Object.entries(state.pdf.rotations).forEach(([key, offset]) => {
        const index = Number(key);
        const meta = state.pdf.pages[index];
        const angle = Number(offset) || 0;
        if (!meta || angle % 360 === 0) {
            return;
        }
        const finalAngle = ((Number(meta.rotation) || 0) + angle + 360) % 360;
        if (finalAngle !== (Number(meta.rotation) || 0)) {
            rotationPayload[String(index)] = finalAngle;
        }
    });
    const deleted = Array.from(state.pdf.deleted).sort((a, b) => a - b);
    const remaining = [];
    for (let index = 1; index <= state.pdf.total; index += 1) {
        if (!state.pdf.deleted.has(index)) {
            remaining.push(index);
        }
    }
    const orderChanged = state.pdf.order.length !== remaining.length
        || state.pdf.order.some((value, index) => value !== remaining[index]);
    if (!orderChanged && !deleted.length && !Object.keys(rotationPayload).length) {
        showToast("没有需要应用的修改", "info");
        return;
    }

    state.pdf.exporting = true;
    updatePdfButtons();
    showTaskProgress("pdf", true);
    setTaskProgress("pdf", 5, "正在提交页面操作");
    const form = new FormData();
    form.append("file", state.pdf.file, state.pdf.file.name);
    form.append("page_order", JSON.stringify(orderChanged ? state.pdf.order : []));
    form.append("rotations", JSON.stringify(rotationPayload));
    form.append("delete_pages", JSON.stringify(deleted));
    try {
        const data = await uploadTaskRequest("/api/pdf/ops", form, (percent) => {
            setTaskProgress("pdf", Math.round(percent * 0.2), `正在上传 ${percent}%`);
        });
        setTaskProgress("pdf", 25, "已提交，正在处理页面");
        watchTask(data.task_id, {
            onProgress: (task) => {
                setTaskProgress("pdf", 25 + Math.round((Number(task.progress) || 0) * 0.7), task.message || "正在处理页面");
            },
            onDone: (task) => {
                state.pdf.exporting = false;
                appendOperationResult("pdf-result-list", task.result);
                state.pdf.resultCount.ops += 1;
                setTaskProgress("pdf", 100, "页面处理完成");
                showToast("页面处理完成", "success");
                updatePdfButtons();
                window.setTimeout(() => showTaskProgress("pdf", false), 800);
            },
            onError: (message) => {
                state.pdf.exporting = false;
                appendOperationError("pdf-result-list", message, "页面处理失败");
                showToast(message || "页面处理失败", "error");
                updatePdfButtons();
                showTaskProgress("pdf", false);
            },
        });
    } catch (error) {
        state.pdf.exporting = false;
        appendOperationError("pdf-result-list", error.message || "上传失败", "页面处理失败");
        showToast(error.message || "上传失败", "error");
        updatePdfButtons();
        showTaskProgress("pdf", false);
    }
}

async function exportPdfJpg() {
    if (!state.pdf.file || state.pdf.jpgBusy || state.pdf.loading) {
        showToast("请先加载一个 PDF", "info");
        return;
    }
    const maxKb = Number($("jpg-max-kb").value);
    if (!Number.isFinite(maxKb) || maxKb < 10 || maxKb > 10240) {
        showToast("每张上限应在 10KB 到 10240KB 之间", "error");
        return;
    }
    state.pdf.jpgBusy = true;
    updatePdfButtons();
    showTaskProgress("jpg", true);
    setTaskProgress("jpg", 5, "正在上传 PDF");
    const form = new FormData();
    form.append("files", state.pdf.file, state.pdf.file.name);
    form.append("max_kb", String(Math.round(maxKb)));
    try {
        const data = await uploadTaskRequest("/api/pdf/to-jpg", form, (percent) => {
            setTaskProgress("jpg", Math.round(percent * 0.2), `正在上传 ${percent}%`);
        });
        setTaskProgress("jpg", 25, "已提交，正在导出 JPG");
        watchTask(data.task_id, {
            onProgress: (task) => {
                setTaskProgress("jpg", 25 + Math.round((Number(task.progress) || 0) * 0.7), task.message || "正在导出 JPG");
            },
            onDone: (task) => {
                state.pdf.jpgBusy = false;
                appendOperationResult("jpg-result-list", task.result);
                state.pdf.resultCount.jpg += 1;
                setTaskProgress("jpg", 100, "JPG 导出完成");
                showToast("JPG 导出完成", "success");
                updatePdfButtons();
                window.setTimeout(() => showTaskProgress("jpg", false), 800);
            },
            onError: (message) => {
                state.pdf.jpgBusy = false;
                appendOperationError("jpg-result-list", message, "JPG 导出失败");
                showToast(message || "JPG 导出失败", "error");
                updatePdfButtons();
                showTaskProgress("jpg", false);
            },
        });
    } catch (error) {
        state.pdf.jpgBusy = false;
        appendOperationError("jpg-result-list", error.message || "上传失败", "JPG 导出失败");
        showToast(error.message || "上传失败", "error");
        updatePdfButtons();
        showTaskProgress("jpg", false);
    }
}

/* ---------------------------------------------------------------------------
   PDF 合并
   --------------------------------------------------------------------------- */

function updateMergeUi() {
    $("merge-count").textContent = `${state.pdf.mergeFiles.length} 个`;
    $("merge-btn").disabled = state.pdf.mergeBusy || state.pdf.mergeFiles.length < 2;
    $("merge-btn-text").textContent = state.pdf.mergeBusy ? "合并中…" : "开始合并";
    $("merge-pick").disabled = state.pdf.mergeBusy;
    $("merge-mode").disabled = state.pdf.mergeBusy;
    $("merge-insert").disabled = state.pdf.mergeBusy;
    $("merge-insert-row").hidden = $("merge-mode").value !== "insert";
}

function renderMergeList() {
    const list = $("merge-list");
    if (!state.pdf.mergeFiles.length) {
        const empty = document.createElement("p");
        empty.className = "merge-empty";
        empty.textContent = "还没有选择 PDF";
        list.replaceChildren(empty);
        updateMergeUi();
        return;
    }
    const nodes = state.pdf.mergeFiles.map((item, index) => {
        const row = document.createElement("div");
        row.className = "merge-item";
        const order = document.createElement("span");
        order.className = "merge-index";
        order.textContent = String(index + 1);
        const name = document.createElement("span");
        name.className = "merge-name";
        name.textContent = item.file.name;
        const size = document.createElement("span");
        size.className = "merge-size";
        size.textContent = formatBytes(item.file.size);
        const remove = document.createElement("button");
        remove.type = "button";
        remove.dataset.action = "remove-merge-file";
        remove.dataset.id = String(item.id);
        remove.title = "移除";
        remove.textContent = "×";
        row.append(order, name, size, remove);
        return row;
    });
    list.replaceChildren(...nodes);
    updateMergeUi();
}

function addMergeFiles(files) {
    const incoming = Array.from(files || []);
    const invalid = incoming.filter((file) => !isPdfFile(file));
    if (invalid.length) {
        showToast(`已忽略 ${invalid.length} 个非 PDF 文件`, "warning");
    }
    const accepted = incoming.filter((file) => isPdfFile(file) && file.size > 0);
    warnLargeFiles(accepted);
    accepted.forEach((file) => {
        if (file.size > state.maxFileMb * 1024 * 1024) {
            showToast(`${file.name} 超过 ${state.maxFileMb}MB 上限`, "error");
            return;
        }
        state.pdf.mergeFiles.push({ id: state.pdf.mergeNextId, file });
        state.pdf.mergeNextId += 1;
    });
    renderMergeList();
}

function removeMergeFile(id) {
    state.pdf.mergeFiles = state.pdf.mergeFiles.filter((item) => item.id !== id);
    renderMergeList();
}

async function startMerge() {
    if (state.pdf.mergeBusy || state.pdf.mergeFiles.length < 2) {
        return;
    }
    const mode = $("merge-mode").value;
    const insertIndex = Number($("merge-insert").value);
    if (mode === "insert" && (!Number.isInteger(insertIndex) || insertIndex < 0)) {
        showToast("插入位置必须是 0 或更大的整数", "error");
        return;
    }

    state.pdf.mergeBusy = true;
    updateMergeUi();
    showTaskProgress("merge", true);
    setTaskProgress("merge", 3, "正在上传 PDF");
    const form = new FormData();
    state.pdf.mergeFiles.forEach((item) => {
        form.append("files", item.file, item.file.name);
    });
    form.append("mode", mode);
    form.append("insert_index", String(mode === "insert" ? insertIndex : 0));
    try {
        const data = await uploadTaskRequest("/api/pdf/merge", form, (percent) => {
            setTaskProgress("merge", Math.round(percent * 0.2), `正在上传 ${percent}%`);
        });
        setTaskProgress("merge", 25, "已提交，正在合并");
        watchTask(data.task_id, {
            onProgress: (task) => {
                setTaskProgress("merge", 25 + Math.round((Number(task.progress) || 0) * 0.7), task.message || "正在合并");
            },
            onDone: (task) => {
                state.pdf.mergeBusy = false;
                appendOperationResult("merge-result-list", task.result);
                state.pdf.resultCount.merge += 1;
                setTaskProgress("merge", 100, "合并完成");
                showToast("合并完成", "success");
                updateMergeUi();
                window.setTimeout(() => showTaskProgress("merge", false), 800);
            },
            onError: (message) => {
                state.pdf.mergeBusy = false;
                appendOperationError("merge-result-list", message, "合并失败");
                showToast(message || "合并失败", "error");
                updateMergeUi();
                showTaskProgress("merge", false);
            },
        });
    } catch (error) {
        state.pdf.mergeBusy = false;
        appendOperationError("merge-result-list", error.message || "上传失败", "合并失败");
        showToast(error.message || "上传失败", "error");
        updateMergeUi();
        showTaskProgress("merge", false);
    }
}

function resetAiResults(listId, emptyId) {
    const list = $(listId);
    if (list) {
        list.replaceChildren();
    }
    const empty = $(emptyId);
    if (empty) {
        empty.hidden = false;
    }
}

function updateTranslateModeHint() {
    const tr = state.ai.translate;
    const hint = $("translate-mode-hint");
    if (!hint) {
        return;
    }
    if (!tr.configured) {
        hint.textContent = "配置 TRANSLATE_API_KEY 后启用；格式转换与 PDF 工具不受影响。";
        return;
    }
    const ext = tr.file ? getExtension(tr.file.name) : "";
    if (tr.mode === "layout") {
        hint.textContent = ext && ext !== "pdf"
            ? "当前文件不是 PDF，保留版式只接受 PDF 输入。"
            : "按原页面位置覆盖译文；只接受 PDF，字号会在 6–18pt 之间自动缩放。";
        return;
    }
    hint.textContent = tr.mode === "translated"
        ? "只输出译文 DOCX，保留标题级别、列表与表格结构。"
        : "生成 DOCX 两列表格，左原文右译文，一段一行。";
}

function updateTranslateUi() {
    const tr = state.ai.translate;
    const configured = tr.configured;
    const file = tr.file;
    const ext = file ? getExtension(file.name) : "";
    const layoutBlocked = tr.mode === "layout" && Boolean(file) && ext !== "pdf";

    $("translate-panel").classList.toggle("is-locked", !configured);
    $("translate-panel").classList.toggle("is-busy", tr.busy);
    $("translate-config-warning").hidden = configured;
    $("translate-tag").textContent = configured ? "PDF / DOCX / TXT" : "未配置密钥";

    const zone = $("translate-dropzone");
    zone.classList.toggle("is-disabled", tr.busy || !configured);
    zone.setAttribute("aria-disabled", tr.busy || !configured ? "true" : "false");

    const button = $("translate-btn");
    button.disabled = !configured || !file || tr.busy || layoutBlocked;
    $("translate-btn-text").textContent = tr.busy ? "正在翻译…" : "开始翻译";
    $("translate-remove").disabled = tr.busy || !file;
    $("translate-lang").disabled = tr.busy || !configured;
    $("translate-glossary").disabled = tr.busy || !configured;

    document.querySelectorAll("#translate-modes .mode-card").forEach((card) => {
        card.classList.toggle("is-active", card.dataset.mode === tr.mode);
        card.disabled = tr.busy || !configured;
    });
    updateTranslateModeHint();
}

function setTranslateMode(mode) {
    const tr = state.ai.translate;
    if (tr.busy || !["bilingual", "translated", "layout"].includes(mode)) {
        return;
    }
    if (mode === "layout" && tr.file && getExtension(tr.file.name) !== "pdf") {
        showToast("保留版式输出只支持 PDF，请先用格式转换转成 PDF", "error");
        return;
    }
    tr.mode = mode;
    updateTranslateUi();
}

function setTranslateFile(file) {
    const tr = state.ai.translate;
    if (!file || tr.busy || !tr.configured) {
        return;
    }
    const ext = getExtension(file.name);
    if (!TRANSLATE_EXTS.includes(ext)) {
        showToast("翻译只支持 PDF、DOCX 和 TXT 文件", "error");
        return;
    }
    if (file.size === 0) {
        showToast("这个文件是空的，请选择有内容的文件", "error");
        return;
    }
    if (file.size > state.maxFileMb * 1024 * 1024) {
        showToast(`文件超过 ${state.maxFileMb}MB 上限`, "error");
        return;
    }
    tr.file = file;
    warnLargeFile(file);
    if (tr.mode === "layout" && ext !== "pdf") {
        tr.mode = "translated";
        showToast("保留版式只接受 PDF，已切换到纯译文", "info");
    }
    const icon = $("translate-file-icon");
    icon.textContent = ext.toUpperCase().slice(0, 4);
    icon.className = `file-icon ${ext === "pdf" ? "kind-pdf" : "kind-text"}`;
    $("translate-file-name").textContent = file.name;
    $("translate-file-size").textContent = `${formatBytes(file.size)} · ${ext.toUpperCase()} 文件`;
    $("translate-selected").hidden = false;
    resetAiResults("translate-result-list", "translate-empty");
    updateTranslateUi();
}

function clearTranslateFile() {
    const tr = state.ai.translate;
    if (tr.busy) {
        showToast("翻译进行中，请完成后再移除文件", "warning");
        return;
    }
    tr.file = null;
    $("translate-file-input").value = "";
    $("translate-selected").hidden = true;
    resetAiResults("translate-result-list", "translate-empty");
    showTaskProgress("translate", false);
    updateTranslateUi();
}

function finishTranslateTask() {
    state.ai.translate.busy = false;
    updateTranslateUi();
    window.setTimeout(() => showTaskProgress("translate", false), 900);
}

async function startTranslate() {
    const tr = state.ai.translate;
    if (tr.busy || !tr.file) {
        return;
    }
    if (!tr.configured) {
        showToast("未配置翻译密钥，请在 .env 里填写 TRANSLATE_API_KEY", "error");
        return;
    }
    const ext = getExtension(tr.file.name);
    if (tr.mode === "layout" && ext !== "pdf") {
        showToast("保留版式输出只支持 PDF，请先用格式转换转成 PDF", "error");
        return;
    }

    tr.busy = true;
    resetAiResults("translate-result-list", "translate-empty");
    updateTranslateUi();
    showTaskProgress("translate", true);
    setTaskProgress("translate", 0, "正在上传文件");

    const form = new FormData();
    form.append("file", tr.file);
    form.append("target_lang", $("translate-lang").value === "en" ? "en" : "zh");
    form.append("mode", tr.mode);
    form.append("glossary", $("translate-glossary").value || "");

    try {
        const data = await uploadTaskRequest("/api/translate", form, (percent) => {
            setTaskProgress("translate", Math.round(percent * 0.08), `正在上传 ${percent}%`);
        });
        setTaskProgress("translate", 8, "已提交，正在提取段落");
        watchTask(data.task_id, {
            onProgress: (task) => {
                const value = 8 + Math.round((Number(task.progress) || 0) * 0.92);
                setTaskProgress("translate", value, task.message || "正在翻译");
            },
            onDone: (task) => {
                setTaskProgress("translate", 100, "翻译完成");
                appendOperationResult("translate-result-list", task.result);
                $("translate-empty").hidden = true;
                finishTranslateTask();
                showToast("翻译完成，可下载结果", "success");
            },
            onError: (message) => {
                appendOperationError("translate-result-list", message, "翻译失败");
                $("translate-empty").hidden = true;
                finishTranslateTask();
                showToast(message || "翻译失败", "error");
            },
        });
    } catch (error) {
        appendOperationError("translate-result-list", error.message || "上传失败", "翻译失败");
        $("translate-empty").hidden = true;
        setTaskProgress("translate", 0, "上传失败");
        finishTranslateTask();
        showToast(error.message || "上传失败", "error");
    }
}

function updateScanUi() {
    const scan = state.ai.scan;
    const button = $("scan-btn");
    button.disabled = scan.busy || !scan.file;
    $("scan-btn-text").textContent = scan.busy ? "增强中…" : "开始增强";
    $("scan-remove").disabled = scan.busy || !scan.file;
    const zone = $("scan-dropzone");
    zone.classList.toggle("is-disabled", scan.busy);
    zone.setAttribute("aria-disabled", scan.busy ? "true" : "false");
}

function resetScanPreview() {
    const img = $("scan-result");
    img.hidden = true;
    img.removeAttribute("src");
    const placeholder = $("scan-result-placeholder");
    placeholder.hidden = false;
    placeholder.textContent = "增强结果会显示在这里";
}

function setScanFile(file) {
    const scan = state.ai.scan;
    if (!file || scan.busy) {
        return;
    }
    const ext = getExtension(file.name);
    if (!SCAN_EXTS.includes(ext)) {
        showToast("扫描件增强只接受图片文件", "error");
        return;
    }
    if (file.size === 0) {
        showToast("这个文件是空的，请选择有内容的照片", "error");
        return;
    }
    if (file.size > state.maxFileMb * 1024 * 1024) {
        showToast(`文件超过 ${state.maxFileMb}MB 上限`, "error");
        return;
    }
    if (scan.originalUrl) {
        URL.revokeObjectURL(scan.originalUrl);
    }
    scan.file = file;
    scan.originalUrl = URL.createObjectURL(file);
    warnLargeFile(file);
    const icon = $("scan-file-icon");
    icon.textContent = ext.toUpperCase().slice(0, 4);
    icon.className = "file-icon kind-image";
    $("scan-file-name").textContent = file.name;
    $("scan-file-size").textContent = `${formatBytes(file.size)} · ${ext.toUpperCase()} 图片`;
    $("scan-selected").hidden = false;
    $("scan-compare").hidden = false;
    $("scan-original").src = scan.originalUrl;
    resetAiResults("scan-result-list", "scan-empty");
    $("scan-empty").hidden = true;
    resetScanPreview();
    updateScanUi();
}

function clearScanFile() {
    const scan = state.ai.scan;
    if (scan.busy) {
        showToast("增强进行中，请完成后再移除照片", "warning");
        return;
    }
    if (scan.originalUrl) {
        URL.revokeObjectURL(scan.originalUrl);
        scan.originalUrl = "";
    }
    scan.file = null;
    $("scan-file-input").value = "";
    $("scan-selected").hidden = true;
    $("scan-compare").hidden = true;
    $("scan-original").removeAttribute("src");
    resetAiResults("scan-result-list", "scan-empty");
    showTaskProgress("scan", false);
    updateScanUi();
}

function finishScanTask() {
    state.ai.scan.busy = false;
    updateScanUi();
    window.setTimeout(() => showTaskProgress("scan", false), 900);
}

async function startScan() {
    const scan = state.ai.scan;
    if (scan.busy || !scan.file) {
        return;
    }
    scan.busy = true;
    resetAiResults("scan-result-list", "scan-empty");
    resetScanPreview();
    updateScanUi();
    showTaskProgress("scan", true);
    setTaskProgress("scan", 0, "正在上传照片");

    const form = new FormData();
    form.append("file", scan.file);

    try {
        const data = await uploadTaskRequest("/api/scan-photo", form, (percent) => {
            setTaskProgress("scan", Math.round(percent * 0.1), `正在上传 ${percent}%`);
        });
        setTaskProgress("scan", 10, "已提交，正在检测纸张边界");
        watchTask(data.task_id, {
            onProgress: (task) => {
                const value = 10 + Math.round((Number(task.progress) || 0) * 0.9);
                setTaskProgress("scan", value, task.message || "正在增强扫描件");
            },
            onDone: (task) => {
                setTaskProgress("scan", 100, "增强完成");
                appendOperationResult("scan-result-list", task.result);
                $("scan-empty").hidden = true;
                const url = safeDownloadUrl(task.result && task.result.download_url);
                if (url) {
                    const img = $("scan-result");
                    img.src = url;
                    img.hidden = false;
                    $("scan-result-placeholder").hidden = true;
                } else {
                    $("scan-result-placeholder").textContent = "增强完成，请从下方结果卡片下载";
                }
                finishScanTask();
                showToast("扫描件增强完成", "success");
            },
            onError: (message) => {
                appendOperationError("scan-result-list", message, "增强失败");
                $("scan-empty").hidden = true;
                $("scan-result-placeholder").textContent = "增强失败，可查看下方原因后重试";
                finishScanTask();
                showToast(message || "增强失败", "error");
            },
        });
    } catch (error) {
        appendOperationError("scan-result-list", error.message || "上传失败", "增强失败");
        $("scan-empty").hidden = true;
        $("scan-result-placeholder").textContent = "增强失败，可查看下方原因后重试";
        setTaskProgress("scan", 0, "上传失败");
        finishScanTask();
        showToast(error.message || "上传失败", "error");
    }
}

function bindAiDropzone(zoneId, inputId, isBusy, onFile) {
    const zone = $(zoneId);
    const input = $(inputId);
    zone.addEventListener("click", (event) => {
        if (event.target === input) {
            return;
        }
        if (isBusy()) {
            showToast("正在处理，请稍候", "warning");
            return;
        }
        input.click();
    });
    zone.addEventListener("keydown", (event) => {
        if (event.key === "Enter" || event.key === " ") {
            event.preventDefault();
            zone.click();
        }
    });
    input.addEventListener("change", () => {
        const file = input.files && input.files[0];
        input.value = "";
        if (file) {
            onFile(file);
        }
    });
    ["dragenter", "dragover"].forEach((name) => {
        zone.addEventListener(name, (event) => {
            event.preventDefault();
            event.stopPropagation();
            if (!isBusy()) {
                zone.classList.add("dragover");
            }
        });
    });
    ["dragleave", "drop"].forEach((name) => {
        zone.addEventListener(name, (event) => {
            event.preventDefault();
            event.stopPropagation();
            zone.classList.remove("dragover");
        });
    });
    zone.addEventListener("drop", (event) => {
        if (isBusy()) {
            showToast("正在处理，请稍候", "warning");
            return;
        }
        if (dropHasDirectory(event.dataTransfer)) {
            showToast("这里只接收单个文件，请直接拖入文件", "info");
            return;
        }
        const files = event.dataTransfer && event.dataTransfer.files
            ? Array.from(event.dataTransfer.files)
            : [];
        if (files.length > 1) {
            showToast("一次只处理一个文件，请只选择一个", "info");
            return;
        }
        if (files[0]) {
            onFile(files[0]);
        }
    });
}

function bindAiEvents() {
    bindAiDropzone(
        "translate-dropzone",
        "translate-file-input",
        () => state.ai.translate.busy,
        setTranslateFile
    );
    $("translate-remove").addEventListener("click", clearTranslateFile);
    $("translate-modes").addEventListener("click", (event) => {
        const card = event.target.closest("[data-mode]");
        if (card) {
            setTranslateMode(card.dataset.mode);
        }
    });
    $("translate-btn").addEventListener("click", startTranslate);

    bindAiDropzone(
        "scan-dropzone",
        "scan-file-input",
        () => state.ai.scan.busy,
        setScanFile
    );
    $("scan-remove").addEventListener("click", clearScanFile);
    $("scan-btn").addEventListener("click", startScan);

    ["translate-result-list", "scan-result-list"].forEach((id) => {
        $(id).addEventListener("click", (event) => {
            const button = event.target.closest('[data-action="copy-error"]');
            if (!button) {
                return;
            }
            const card = button.closest(".result-card");
            const notes = card ? Array.from(card.querySelectorAll(".result-note")) : [];
            const text = notes.map((node) => node.textContent || "").join("\n").trim();
            if (text) {
                copyText(text);
            }
        });
    });
}

function activatePage(page) {
    document.querySelectorAll(".nav-item").forEach((item) => {
        item.classList.toggle("active", item.dataset.page === page);
    });
    document.querySelectorAll(".page").forEach((section) => {
        const active = section.id === `page-${page}`;
        section.hidden = !active;
        section.classList.toggle("active", active);
    });
    if (page === "batch") {
        renderBatch();
    }
    if (page !== "convert" && state.busy) {
        showToast("转换仍在进行，返回「格式转换」可查看进度", "info");
    }
}

function dropHasDirectory(dataTransfer) {
    if (!dataTransfer || !dataTransfer.items) {
        return false;
    }
    return Array.from(dataTransfer.items).some((item) => {
        if (item.kind !== "file" || typeof item.webkitGetAsEntry !== "function") {
            return false;
        }
        const entry = item.webkitGetAsEntry();
        return Boolean(entry && entry.isDirectory);
    });
}

function bindCompressEvents() {
    const dropzone = $("compress-dropzone");
    const input = $("compress-file-input");
    dropzone.addEventListener("click", (event) => {
        if (event.target === input) {
            return;
        }
        if (state.compress.busy) {
            showToast("压缩进行中，请稍候", "warning");
            return;
        }
        input.click();
    });
    dropzone.addEventListener("keydown", (event) => {
        if (event.key === "Enter" || event.key === " ") {
            event.preventDefault();
            dropzone.click();
        }
    });
    input.addEventListener("change", () => {
        const file = input.files && input.files[0];
        input.value = "";
        if (file) {
            setCompressFile(file);
        }
    });
    ["dragenter", "dragover"].forEach((name) => {
        dropzone.addEventListener(name, (event) => {
            event.preventDefault();
            event.stopPropagation();
            if (!state.compress.busy) {
                dropzone.classList.add("dragover");
            }
        });
    });
    ["dragleave", "drop"].forEach((name) => {
        dropzone.addEventListener(name, (event) => {
            event.preventDefault();
            event.stopPropagation();
            dropzone.classList.remove("dragover");
        });
    });
    dropzone.addEventListener("drop", (event) => {
        if (state.compress.busy) {
            showToast("压缩进行中，请稍候", "warning");
            return;
        }
        const files = event.dataTransfer && event.dataTransfer.files
            ? Array.from(event.dataTransfer.files)
            : [];
        if (files.length > 1) {
            showToast("压缩一次只处理一个文件，请只选择一个", "info");
            return;
        }
        if (files[0]) {
            setCompressFile(files[0]);
        }
    });

    $("compress-remove").addEventListener("click", clearCompressFile);
    $("compress-levels").addEventListener("click", (event) => {
        const button = event.target.closest("[data-level]");
        if (!button || state.compress.busy) {
            return;
        }
        state.compress.level = button.dataset.level;
        updateCompressUi();
    });
    $("compress-btn").addEventListener("click", startCompress);
}

function bindPdfEvents() {
    const dropzone = $("pdf-dropzone");
    const input = $("pdf-file-input");
    dropzone.addEventListener("click", (event) => {
        if (event.target === input) {
            return;
        }
        if (state.pdf.loading || state.pdf.exporting) {
            showToast("PDF 正在处理，请稍候", "warning");
            return;
        }
        input.click();
    });
    dropzone.addEventListener("keydown", (event) => {
        if (event.key === "Enter" || event.key === " ") {
            event.preventDefault();
            dropzone.click();
        }
    });
    input.addEventListener("change", () => {
        const file = input.files && input.files[0];
        input.value = "";
        if (file) {
            loadPdfForEditing(file);
        }
    });
    ["dragenter", "dragover"].forEach((name) => {
        dropzone.addEventListener(name, (event) => {
            event.preventDefault();
            event.stopPropagation();
            if (!state.pdf.loading && !state.pdf.exporting) {
                dropzone.classList.add("dragover");
            }
        });
    });
    ["dragleave", "drop"].forEach((name) => {
        dropzone.addEventListener(name, (event) => {
            event.preventDefault();
            event.stopPropagation();
            dropzone.classList.remove("dragover");
        });
    });
    dropzone.addEventListener("drop", (event) => {
        if (state.pdf.loading || state.pdf.exporting) {
            return;
        }
        const files = event.dataTransfer && event.dataTransfer.files
            ? Array.from(event.dataTransfer.files)
            : [];
        if (files.length > 1) {
            showToast("页面编辑一次加载一个 PDF；合并请用右侧的合并面板", "info");
            return;
        }
        if (files[0]) {
            loadPdfForEditing(files[0]);
        }
    });

    $("pdf-remove").addEventListener("click", clearPdfFile);
    $("pdf-undo").addEventListener("click", undoPdf);
    $("pdf-redo").addEventListener("click", redoPdf);
    $("pdf-select-all").addEventListener("click", selectAllPdfPages);
    $("pdf-invert").addEventListener("click", invertPdfSelection);
    $("pdf-copy-selected").addEventListener("click", copySelectedPdfPages);
    $("pdf-rotate-selected").addEventListener("click", rotateSelectedPdfPages);
    $("pdf-delete-selected").addEventListener("click", deleteSelectedPdfPages);
    $("pdf-reset-order").addEventListener("click", resetPdfOrder);
    $("pdf-export").addEventListener("click", exportPdfOps);
    $("pdf-jpg-btn").addEventListener("click", exportPdfJpg);

    const grid = $("pdf-grid");
    grid.addEventListener("change", (event) => {
        const checkbox = event.target.closest('[data-action="toggle-page"]');
        if (!checkbox) {
            return;
        }
        togglePdfSelection(Number(checkbox.dataset.index), checkbox.checked);
    });
    grid.addEventListener("click", (event) => {
        const button = event.target.closest("[data-action]");
        if (!button || button.tagName === "INPUT") {
            return;
        }
        const index = Number(button.dataset.index);
        if (button.dataset.action === "rotate-page") {
            rotatePdfPage(index);
        } else if (button.dataset.action === "delete-page") {
            deletePdfPage(index);
        }
    });
    grid.addEventListener("dragstart", (event) => {
        const card = event.target.closest(".pdf-page-card");
        if (!card) {
            return;
        }
        const index = Number(card.dataset.index);
        if (!state.pdf.selected.has(index)) {
            event.preventDefault();
            return;
        }
        state.pdf.dragIndex = index;
        card.classList.add("is-dragging");
        if (event.dataTransfer) {
            event.dataTransfer.effectAllowed = "move";
            event.dataTransfer.setData("text/plain", String(index));
        }
    });
    grid.addEventListener("dragend", (event) => {
        const card = event.target.closest(".pdf-page-card");
        if (card) {
            card.classList.remove("is-dragging");
        }
        state.pdf.dragIndex = null;
    });
    grid.addEventListener("dragover", (event) => {
        if (state.pdf.dragIndex === null) {
            return;
        }
        const card = event.target.closest(".pdf-page-card");
        if (card) {
            event.preventDefault();
            if (event.dataTransfer) {
                event.dataTransfer.dropEffect = "move";
            }
        }
    });
    grid.addEventListener("drop", (event) => {
        if (state.pdf.dragIndex === null) {
            return;
        }
        const card = event.target.closest(".pdf-page-card");
        if (!card) {
            return;
        }
        event.preventDefault();
        const targetIndex = Number(card.dataset.index);
        const dragIndex = state.pdf.dragIndex;
        state.pdf.dragIndex = null;
        reorderPdfPages(dragIndex, targetIndex);
    });

    $("merge-pick").addEventListener("click", () => {
        if (!state.pdf.mergeBusy) {
            $("merge-file-input").click();
        }
    });
    $("merge-file-input").addEventListener("change", (event) => {
        const files = event.target.files ? Array.from(event.target.files) : [];
        event.target.value = "";
        addMergeFiles(files);
    });
    $("merge-list").addEventListener("click", (event) => {
        const button = event.target.closest('[data-action="remove-merge-file"]');
        if (!button || state.pdf.mergeBusy) {
            return;
        }
        removeMergeFile(Number(button.dataset.id));
    });
    $("merge-mode").addEventListener("change", updateMergeUi);
    $("merge-btn").addEventListener("click", startMerge);

    ["compress-result-list", "pdf-result-list", "merge-result-list", "jpg-result-list"].forEach((id) => {
        $(id).addEventListener("click", (event) => {
            const button = event.target.closest('[data-action="copy-error"]');
            if (!button) {
                return;
            }
            const card = button.closest(".result-card");
            const notes = card ? Array.from(card.querySelectorAll(".result-note")) : [];
            const text = notes.map((node) => node.textContent || "").join("\n").trim();
            if (text) {
                copyText(text);
            }
        });
    });
}

function bindEvents() {
    $("theme-toggle").addEventListener("click", () => {
        const index = THEME_MODES.indexOf(savedThemeMode());
        applyTheme(THEME_MODES[(index + 1) % THEME_MODES.length]);
    });

    $("nav").addEventListener("click", (event) => {
        const item = event.target.closest(".nav-item");
        if (item) {
            activatePage(item.dataset.page);
        }
    });

    document.addEventListener("keydown", (event) => {
        if (!(event.ctrlKey || event.metaKey) || event.altKey) {
            return;
        }
        if (String(event.key).toLowerCase() !== "z") {
            return;
        }
        const target = event.target;
        if (target instanceof HTMLElement
            && (target.matches("input, textarea, select") || target.isContentEditable)) {
            return;
        }
        if (!state.pdf.file || $("pdf-editor").hidden) {
            return;
        }
        event.preventDefault();
        if (event.shiftKey) {
            redoPdf();
        } else {
            undoPdf();
        }
    });

    const dropzone = $("dropzone");
    const input = $("file-input");
    dropzone.addEventListener("click", (event) => {
        if (event.target === input) {
            return;
        }
        if (state.busy) {
            showToast("转换进行中，请稍候", "warning");
            return;
        }
        input.click();
    });
    dropzone.addEventListener("keydown", (event) => {
        if (event.key === "Enter" || event.key === " ") {
            event.preventDefault();
            dropzone.click();
        }
    });
    input.addEventListener("change", () => {
        const file = input.files && input.files[0];
        if (file) {
            setFile(file);
        }
        input.value = "";
    });

    ["dragenter", "dragover"].forEach((name) => {
        dropzone.addEventListener(name, (event) => {
            event.preventDefault();
            event.stopPropagation();
            if (!state.busy) {
                dropzone.classList.add("dragover");
            }
        });
    });
    ["dragleave", "drop"].forEach((name) => {
        dropzone.addEventListener(name, (event) => {
            event.preventDefault();
            event.stopPropagation();
            dropzone.classList.remove("dragover");
        });
    });
    dropzone.addEventListener("drop", (event) => {
        if (state.busy) {
            showToast("转换进行中，请稍候", "warning");
            return;
        }
        if (dropHasDirectory(event.dataTransfer)) {
            showToast("请切换到「批量处理」中添加整个文件夹", "info");
            return;
        }
        const dropped = event.dataTransfer && event.dataTransfer.files
            ? Array.from(event.dataTransfer.files)
            : [];
        if (dropped.length > 1) {
            showToast("多个文件请切换到「批量处理」中添加", "info");
            return;
        }
        if (dropped.length === 1) {
            setFile(dropped[0]);
        }
    });
    window.addEventListener("dragover", (event) => event.preventDefault());
    window.addEventListener("drop", (event) => event.preventDefault());

    $("remove-file").addEventListener("click", clearFile);
    $("target-select").addEventListener("change", (event) => {
        state.target = event.target.value;
        updateOfficeWarning();
    });
    $("convert-btn").addEventListener("click", startConvert);

    $("result-list").addEventListener("click", (event) => {
        const button = event.target.closest('[data-action="copy-error"]');
        if (!button) {
            return;
        }
        const card = button.closest(".result-card");
        const note = card && card.querySelector(".result-note");
        if (note) {
            copyText(note.textContent || "");
        }
    });

    bindBatchEvents();
    bindCompressEvents();
    bindPdfEvents();
    bindAiEvents();
}

document.addEventListener("DOMContentLoaded", () => {
    initTheme();
    bindEvents();
    renderBatch();
    loadHealth();
    loadConfigAndFormats();
});