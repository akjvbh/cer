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
};

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

function currentTheme() {
    const explicit = document.documentElement.getAttribute("data-theme");
    if (explicit === "dark" || explicit === "light") {
        return explicit;
    }
    return window.matchMedia("(prefers-color-scheme: dark)").matches ? "dark" : "light";
}

function updateThemeButton() {
    const button = $("theme-toggle");
    const dark = currentTheme() === "dark";
    const label = dark ? "切换到浅色主题" : "切换到深色主题";
    button.setAttribute("aria-label", label);
    button.title = label;
}

function applyTheme(theme) {
    document.documentElement.setAttribute("data-theme", theme);
    try {
        localStorage.setItem(THEME_KEY, theme);
    } catch (error) {
        /* 隐私模式下 localStorage 不可用时，仅当前页面生效。 */
    }
    updateThemeButton();
}

function initTheme() {
    let saved = null;
    try {
        saved = localStorage.getItem(THEME_KEY);
    } catch (error) {
        saved = null;
    }
    if (saved === "dark" || saved === "light") {
        document.documentElement.setAttribute("data-theme", saved);
    }
    updateThemeButton();
    const media = window.matchMedia("(prefers-color-scheme: dark)");
    if (typeof media.addEventListener === "function") {
        media.addEventListener("change", () => {
            if (!document.documentElement.hasAttribute("data-theme")) {
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

    const note = document.createElement("p");
    note.className = "result-note";
    note.textContent = message || "转换失败，请重试";

    const actions = document.createElement("div");
    actions.className = "result-actions";
    const copy = document.createElement("button");
    copy.className = "copy-button";
    copy.type = "button";
    copy.dataset.action = "copy-error";
    copy.textContent = "复制错误信息";
    actions.append(copy);

    body.append(title, note, actions);
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
    }, 3400);
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

function bindEvents() {
    $("theme-toggle").addEventListener("click", () => {
        applyTheme(currentTheme() === "dark" ? "light" : "dark");
    });

    $("nav").addEventListener("click", (event) => {
        const item = event.target.closest(".nav-item");
        if (item) {
            activatePage(item.dataset.page);
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
}

document.addEventListener("DOMContentLoaded", () => {
    initTheme();
    bindEvents();
    renderBatch();
    loadHealth();
    loadConfigAndFormats();
});