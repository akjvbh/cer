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
        const input = $("file-input");
        const extensions = Object.keys(state.formats).map((ext) => `.${ext}`).join(",");
        input.accept = extensions;
        $("dropzone-sub").textContent = `支持 Word / PDF / 图片 / Markdown / TXT · 单文件 ≤${state.maxFileMb}MB`;
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

async function copyText(text) {
    try {
        await navigator.clipboard.writeText(text);
        showToast("错误信息已复制", "success");
    } catch (error) {
        const textarea = document.createElement("textarea");
        textarea.value = text;
        textarea.style.position = "fixed";
        textarea.style.opacity = "0";
        document.body.append(textarea);
        textarea.select();
        try {
            document.execCommand("copy");
            showToast("错误信息已复制", "success");
        } catch (fallbackError) {
            showToast("复制失败，请手动选择文字", "error");
        } finally {
            textarea.remove();
        }
    }
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
            showToast("文件夹拖拽会在「批量处理」中开放，当前请先拖入单个文件", "info");
            return;
        }
        const file = event.dataTransfer && event.dataTransfer.files && event.dataTransfer.files[0];
        if (file) {
            setFile(file);
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
}

document.addEventListener("DOMContentLoaded", () => {
    initTheme();
    bindEvents();
    loadHealth();
    loadConfigAndFormats();
});