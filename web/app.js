﻿﻿﻿/* 肥鸭文档转换 —— 前端交互逻辑 */
"use strict";

/* ---------------- 状态 ---------------- */
const state = {
  files: [],          // {id, name, size, status, tag}
  selected: new Set(),
  mode: "full",
  layout: "per_image",   // per_image：一图一表；merged：多图合并一张表
  output: "",
  running: false,
  jobId: null,
  ws: null,
  tab: "files",
  flags: [],
  flagsFilter: "all",
  flagStates: {},     // {flagIndex: "checked" | "retake"}，按任务持久化到本机
  lastOutPath: "",
  pendingSearchQuery: "",
  autoOpen: false,
  checkUpdates: true,
  lastUpdate: null,
};

const MODES = [
  { value: "full", name: "图片转 Excel", desc: "保留原来的表格识别与 Excel 导出" },
  { value: "find", name: "查找单号", desc: "粘贴单号，在所选图片中快速查找并定位原图" },
  { value: "waybill", name: "单号清单导出", desc: "只把识别到的运单号汇总成一列" },
];

const MODE_ICONS = {
  full: '<svg viewBox="0 0 24 24" aria-hidden="true"><rect x="3.5" y="4" width="17" height="16" rx="2"/><path d="M3.5 9h17M9 9v11M14.5 9v11"/><path d="M6.5 6.5h1"/></svg>',
  find: '<svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="10.8" cy="10.8" r="6.2"/><path d="m15.5 15.5 4.2 4.2"/><path d="M8.3 10.8h5"/></svg>',
  waybill: '<svg viewBox="0 0 24 24" aria-hidden="true"><rect x="4" y="4" width="16" height="16" rx="2"/><path d="M8 9h8M8 12h8M8 15h5"/></svg>',
};

const LAYOUTS = [
  { value: "per_image", name: "一图一表", desc: "每张图片一个工作表，工作表名即图片名" },
  { value: "merged", name: "多图一表", desc: "全部图片合成一张表，首列写明来源图片" },
];

const TAG_LABEL = {
  pending: ["待转换", "pill-pending"],
  run: ["识别中…", "pill-run"],
  ok: ["已完成", "pill-ok"],
  fail: ["识别失败", "pill-fail"],
  skip: ["已跳过", "pill-skip"],
};

const EXT_CLASS = {
  jpg: "jpg", jpeg: "jpg", png: "img", bmp: "img",
  tif: "img", tiff: "img", webp: "img", pdf: "pdf",
};
function fileType(name) {
  const m = /\.([a-z0-9]+)$/i.exec(name || "");
  const ext = m ? m[1].toUpperCase() : "FILE";
  return { ext, cls: EXT_CLASS[ext.toLowerCase()] || "file" };
}
const TAG_ICON = {
  pending: '<svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="8"/><path d="M12 8v5l3 2"/></svg>',
  run: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M19 7v5h-5M5 17v-5h5"/><path d="M7 7a7 7 0 0 1 11 2M17 17A7 7 0 0 1 6 14"/></svg>',
  ok: '<svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="9"/><path d="m8 12 3 3 5-6"/></svg>',
  fail: '<svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="9"/><path d="m9 9 6 6m0-6-6 6"/></svg>',
  skip: '<svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="8"/><path d="M12 8v5l3 2"/></svg>',
};

const $ = (id) => document.getElementById(id);
const MAX_FILE_BYTES = 100 * 1024 * 1024;   // 单文件上限：100 MB
const els = {
  dropzone: $("dropzone"), fileInput: $("fileInput"),
  btnAdd: $("btnAdd"), btnPaste: $("btnPaste"),
  btnRemove: $("btnRemove"), btnClear: $("btnClear"),
  fileCount: $("fileCount"), fileTbody: $("fileTbody"), checkAll: $("checkAll"),
  resultCard: $("resultCard"), rcTitle: $("rcTitle"), rcSubtitle: $("rcSubtitle"),
  rcFile: $("rcFile"), rcFileName: $("rcFileName"), rcNote: $("rcNote"),
  rcMetrics: $("rcMetrics"), rcActions: $("rcActions"),
  rcQuality: $("rcQuality"), rcQTier: $("rcQTier"), rcQScore: $("rcQScore"),
  rcQToggle: $("rcQToggle"), rcQList: $("rcQList"), rcQWhy: $("rcQWhy"), rcQSug: $("rcQSug"),
  rcFiles: $("rcFiles"), rcFilesBody: $("rcFilesBody"), rcFilesCount: $("rcFilesCount"),
  rcFails: $("rcFails"), rcFailCount: $("rcFailCount"), rcFailList: $("rcFailList"),
  statOk: $("statOk"), statWarn: $("statWarn"), statBad: $("statBad"),
  statOkLabel: $("statOkLabel"), statWarnLabel: $("statWarnLabel"),
  statBadLabel: $("statBadLabel"),
  btnOpen: $("btnOpen"), btnReveal: $("btnReveal"),
  btnDownload: $("btnDownload"), resultClose: $("resultClose"),
  btnRetryFailed: $("btnRetryFailed"), btnFullCheck: $("btnFullCheck"),
  logBox: $("logBox"),
  tabs: $("tabs"), tabFiles: $("tabFiles"), tabFlags: $("tabFlags"), tabLog: $("tabLog"),
  tabFilesN: $("tabFilesN"), tabFlagsN: $("tabFlagsN"),
  paneFiles: $("paneFiles"), paneFlags: $("paneFlags"), paneLog: $("paneLog"),
  flagHintTitle: $("flagHintTitle"), flagHintText: $("flagHintText"), btnFlagsOpen: $("btnFlagsOpen"),
  modeList: $("modeList"), layoutSection: $("layoutSection"), layoutList: $("layoutList"),
  outPath: $("outPath"), btnBrowse: $("btnBrowse"),
  outputSection: $("outputSection"), targetSection: $("targetSection"),
  sumMode: $("sumMode"), sumCount: $("sumCount"), sumFormat: $("sumFormat"),
  waybillInput: $("waybillInput"),
  btnStart: $("btnStart"), btnCancel: $("btnCancel"), btnOpenDir: $("btnOpenDir"),
  progressTrack: $("progressTrack"), progressFill: $("progressFill"), progressPct: $("progressPct"),
  statusText: $("statusText"), statusMsg: $("statusMsg"), elapsed: $("elapsed"),
  btnAbout: $("btnAbout"),
  firstRunModal: $("firstRunModal"), firstRunNotice: $("firstRunNotice"),
  firstRunVersion: $("firstRunVersion"), firstRunAck: $("firstRunAck"), firstRunConfirm: $("firstRunConfirm"),
  aboutModal: $("aboutModal"), aboutNotice: $("aboutNotice"),
  aboutVersion: $("aboutVersion"), aboutConfirm: $("aboutConfirm"),
  updateBanner: $("updateBanner"), updateVer: $("updateVer"),
  updateLink: $("updateLink"), updateClose: $("updateClose"),
  updateStatus: $("updateStatus"), btnCheckUpdate: $("btnCheckUpdate"),
  btnCloud: $("btnCloud"), cloudModal: $("cloudModal"), cloudToggle: $("cloudToggle"),
  cloudStatus: $("cloudStatus"), cloudConfirm: $("cloudConfirm"),

  btnFlags: $("btnFlags"), flagModal: $("flagModal"),
  // 页签内三组折叠表（S5）与弹窗表（旧入口，保留）分属两个 tbody
  flagBody: $("flagBody"), flagBodyModal: $("flagBodyModal"),
  flagPaneWrap: $("flagPaneWrap"), flagPaneSummary: $("flagPaneSummary"),
  btnFlagsModal: $("btnFlagsModal"),
  flagTools: $("flagTools"), flagFileFilter: $("flagFileFilter"),
  flagPrioFilter: $("flagPrioFilter"), flagSearch: $("flagSearch"),
  flagProgress: $("flagProgress"),
  flagSummary: $("flagSummary"), flagClose: $("flagClose"),

  btnSearch: $("btnSearch"), searchModal: $("searchModal"), searchInput: $("searchInput"),
  searchMeta: $("searchMeta"), searchResults: $("searchResults"),
  searchClear: $("searchClear"), searchClose: $("searchClose"),
  searchModeSingle: $("searchModeSingle"), searchModeBatch: $("searchModeBatch"),
  searchSingleRow: $("searchSingleRow"), searchBatchBox: $("searchBatchBox"),
  searchBatchInput: $("searchBatchInput"), searchBatchRun: $("searchBatchRun"),
  searchBatchExport: $("searchBatchExport"), searchBatchClear: $("searchBatchClear"),

  btnSettings: $("btnSettings"), settingsModal: $("settingsModal"),
  setDefaultMode: $("setDefaultMode"), setAutoOpen: $("setAutoOpen"),
  setJobKeep: $("setJobKeep"), setOutPath: $("setOutPath"),
  btnSettingsCloud: $("btnSettingsCloud"), settingsConfirm: $("settingsConfirm"),
  btnClearJobs: $("btnClearJobs"),

  btnJobs: $("btnJobs"), jobsModal: $("jobsModal"), jobsList: $("jobsList"),
  jobsMeta: $("jobsMeta"), jobsClose: $("jobsClose"),

  previewModal: $("previewModal"), previewTitle: $("previewTitle"),
  previewMeta: $("previewMeta"), previewImg: $("previewImg"),
  previewCanvas: $("previewCanvas"), previewMarks: $("previewMarks"),
  previewStage: $("previewStage"),
  previewHint: $("previewHint"), btnPreviewCopy: $("btnPreviewCopy"),
  btnPreviewCopyImage: $("btnPreviewCopyImage"),
  btnPreviewDelete: $("btnPreviewDelete"), btnPreviewClose: $("btnPreviewClose"),
  previewTools: $("previewTools"), previewAngle: $("previewAngle"),
  btnRotateLeft: $("btnRotateLeft"), btnRotateRight: $("btnRotateRight"),
  btnRotateReset: $("btnRotateReset"),
  previewZoomBar: $("previewZoomBar"), btnZoomIn: $("btnZoomIn"),
  btnZoomOut: $("btnZoomOut"), btnZoomReset: $("btnZoomReset"),
  btnZoomActual: $("btnZoomActual"),

  pasteChip: $("pasteChip"), pasteChipText: $("pasteChipText"),
  pasteChipClose: $("pasteChipClose"), dragOverlay: $("dragOverlay"),
  setHoverActivate: $("setHoverActivate"), setPasteHotkey: $("setPasteHotkey"),
  setCheckUpdates: $("setCheckUpdates"),
};

/* ---------------- 工具 ---------------- */
async function revealPath(path, mode = "select") {
  if (!path) return false;
  try {
    const res = await fetch("/api/reveal", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ path, mode }),
    });
    const data = await res.json().catch(() => ({}));
    return data.ok === true;
  } catch { return false; }
}

function humanSize(bytes) {
  if (bytes >= 1048576) return (bytes / 1048576).toFixed(1) + " MB";
  if (bytes >= 1024) return Math.round(bytes / 1024) + " KB";
  return bytes + " B";
}

let toastTimer = null;
function toast(msg) {
  let el = document.querySelector(".toast");
  if (!el) {
    el = document.createElement("div");
    el.className = "toast";
    document.body.appendChild(el);
  }
  el.textContent = msg;
  el.style.display = "block";
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => { el.style.display = "none"; }, 2600);
}

function log(text, withTime = true) {
  const time = withTime ? `[${new Date().toTimeString().slice(0, 8)}] ` : "";
  const span = document.createElement("span");
  span.className = "log-time";
  span.textContent = time;
  els.logBox.appendChild(span);
  els.logBox.appendChild(document.createTextNode(text + "\n"));
  els.logBox.scrollTop = els.logBox.scrollHeight;
}

/* ---------------- 模式卡片 ---------------- */
function renderModes() {
  els.modeList.innerHTML = "";
  for (const m of MODES) {
    const card = document.createElement("div");
    card.className = "mode-opt" + (state.mode === m.value ? " on" : "");
    card.dataset.value = m.value;
    card.setAttribute("role", "radio");
    card.setAttribute("aria-checked", state.mode === m.value ? "true" : "false");
    card.innerHTML =
      `<span class="mode-icon">${MODE_ICONS[m.value] || ""}</span>` +
      `<span class="mode-text"><b>${m.name}</b><i>${m.desc}</i></span>` +
      `<span class="mode-radio" aria-hidden="true"></span>`;
    card.addEventListener("click", () => {
      if (state.running) { toast("转换进行中，暂不能切换模式"); return; }
      state.mode = m.value;
      renderModes();
      syncModeControls();
    });
    els.modeList.appendChild(card);
  }
  syncModeControls();
}

function renderLayouts() {
  if (!els.layoutList) return;
  els.layoutList.innerHTML = "";
  for (const item of LAYOUTS) {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "layout-opt" + (state.layout === item.value ? " on" : "");
    btn.dataset.value = item.value;
    btn.setAttribute("role", "radio");
    btn.setAttribute("aria-checked", state.layout === item.value ? "true" : "false");
    btn.innerHTML = `<b>${item.name}</b><i>${item.desc}</i>`;
    btn.addEventListener("click", () => {
      if (state.running) { toast("转换进行中，暂不能切换输出方式"); return; }
      state.layout = item.value;
      renderLayouts();
      // 记住本次选择，下次启动沿用
      fetch("/api/config", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ last_layout: state.layout }),
      }).catch(() => {});
    });
    els.layoutList.appendChild(btn);
  }
}

function syncOutputSummary() {
  if (!els.sumMode) return;
  const m = MODES.find((x) => x.value === state.mode);
  els.sumMode.textContent = m ? m.name : "—";
  els.sumCount.textContent = `${state.files.length} 个`;
  els.sumFormat.textContent = state.mode === "find"
    ? "查找结果（不导出）"
    : state.mode === "waybill" ? "单号清单 (.xlsx)" : "Excel (.xlsx)";
}

function syncModeControls() {
  const finding = state.mode === "find";
  els.outputSection.hidden = finding;
  els.targetSection.hidden = !finding;
  if (els.layoutSection) els.layoutSection.hidden = state.mode !== "full";
  renderLayouts();
  syncOutputSummary();
  els.btnStart.innerHTML = finding
    ? '<svg class="bolt" viewBox="0 0 24 24" aria-hidden="true"><path d="M13 2 5 13h6l-1 9 9-12h-6V2Z"/></svg>开始查找'
    : state.mode === "waybill"
      ? '<svg class="bolt" viewBox="0 0 24 24" aria-hidden="true"><path d="M13 2 5 13h6l-1 9 9-12h-6V2Z"/></svg>导出单号清单'
      : '<svg class="bolt" viewBox="0 0 24 24" aria-hidden="true"><path d="M13 2 5 13h6l-1 9 9-12h-6V2Z"/></svg>开始转换';
  refreshStartButton();
}

/* ---------------- 页签：文件清单 / 待复核 / 运行日志 ---------------- */
function renderTabs() {
  const n = state.flags.length;
  els.tabFilesN.textContent = state.files.length;
  els.tabFlagsN.textContent = n;
  els.tabFlagsN.style.display = n ? "" : "none";
  els.tabFlags.classList.toggle("alert", n > 0);
  if (n) {
    els.flagHintTitle.textContent = `本次有 ${n} 处待复核`;
    els.flagHintText.textContent = "待复核项已在 Excel 中标黄，可在「待复核清单」中对照原图核对。";
    els.btnFlagsOpen.textContent = `打开待复核清单（${n}）`;
  } else {
    els.flagHintTitle.textContent = "暂无待复核项";
    els.flagHintText.textContent = "转换完成后，置信度不足、需人工核对的单元格会汇总到这里。";
    els.btnFlagsOpen.textContent = "打开待复核清单";
  }
}

const TAB_ORDER = ["files", "flags", "log"];
const TAB_BTN = () => ({ files: els.tabFiles, flags: els.tabFlags, log: els.tabLog });
const TAB_PANE = () => ({ files: els.paneFiles, flags: els.paneFlags, log: els.paneLog });

function applyTab(name) {
  const t = TAB_ORDER.includes(name) ? name : "files";
  state.tab = t;
  const btns = TAB_BTN();
  const panes = TAB_PANE();
  for (const key of TAB_ORDER) {
    const on = key === t;
    const btn = btns[key];
    btn.classList.toggle("on", on);
    btn.setAttribute("aria-selected", on ? "true" : "false");
    // roving tabindex：只有当前页签在 Tab 序列里，其余用方向键切换
    btn.tabIndex = on ? 0 : -1;
    panes[key].hidden = !on;
  }
}

/* ---------------- 文件列表 ---------------- */
function selectedCount() {
  return state.files.reduce((n, f) => n + (state.selected.has(f.id) ? 1 : 0), 0);
}

function syncFileSelection() {
  const total = state.files.length;
  const picked = selectedCount();
  els.checkAll.checked = total > 0 && picked === total;
  els.checkAll.indeterminate = picked > 0 && picked < total;
  els.checkAll.disabled = total === 0;
  els.btnRemove.disabled = picked === 0;
}

function setSelected(id, on) {
  if (on) state.selected.add(id);
  else state.selected.delete(id);
}

function removeFiles(ids) {
  if (state.running) { toast("转换进行中，请先取消或等待完成"); return; }
  const set = new Set(ids);
  const before = state.files.length;
  state.files = state.files.filter((f) => !set.has(f.id));
  ids.forEach((id) => state.selected.delete(id));
  const removed = before - state.files.length;
  if (removed) log(`已从列表删除 ${removed} 个文件（未删除磁盘原文件）。`);
  renderFiles();
}

function renderFiles() {
  els.fileTbody.innerHTML = "";
  if (state.files.length === 0) {
    const tr = document.createElement("tr");
    tr.className = "empty-row";
    tr.innerHTML = '<td colspan="5">暂无文件，请添加、拖入或粘贴图片</td>';
    els.fileTbody.appendChild(tr);
  }
  for (const f of state.files) {
    const [label, cls] = TAG_LABEL[f.tag] || TAG_LABEL.pending;
    const checked = state.selected.has(f.id);
    const { ext, cls: badgeCls } = fileType(f.name);
    const rotBadge = f.rotate
      ? `<span class="rot-badge" title="转换时按 ${f.rotate}° 方向识别">↻ ${f.rotate}°</span>`
      : "";
    const tr = document.createElement("tr");
    tr.dataset.id = f.id;
    if (checked) tr.classList.add("selected");
    tr.innerHTML = `
      <td class="c-check"><input type="checkbox" class="row-check" ${checked ? "checked" : ""} aria-label="选择 ${esc(f.name)}"></td>
      <td title="${esc(f.name)}">
        <div class="file-cell">
          <span class="file-badge ${badgeCls}">${ext}</span>
          <div class="file-info">
            <span class="file-name-text">${esc(f.name)}${rotBadge}</span>
            <span class="file-sub">${f.path ? "本地文件" : "拖入文件"}</span>
          </div>
        </div>
      </td>
      <td class="c-size">${humanSize(f.size)}</td>
      <td><span class="pill ${cls}">${TAG_ICON[f.tag] || TAG_ICON.pending}${label}</span></td>
      <td class="c-act"><button class="file-view" type="button" data-id="${esc(f.id)}" title="查看这张图片">查看</button><button class="file-del" type="button" data-id="${esc(f.id)}" title="从列表删除（不删除原文件）">删除</button></td>`;
    els.fileTbody.appendChild(tr);
  }
  const totalBytes = state.files.reduce((n, f) => n + (Number(f.size) || 0), 0);
  els.fileCount.textContent = state.files.length
    ? `${state.files.length} 个文件 · 共 ${humanSize(totalBytes)}`
    : "共 0 个文件";
  els.btnClear.disabled = state.files.length === 0;
  // 空态展示完整的拖拽引导，有文件后收成一条，把高度让给清单与结果面板
  const hasFiles = state.files.length > 0;
  els.paneFiles.classList.toggle("has-files", hasFiles);
  els.paneFiles.classList.toggle("is-empty", !hasFiles);
  syncFileSelection();
  syncOutputSummary();
  refreshStartButton();
  renderTabs();
}

// 行内交互统一走事件委托，避免每次重渲染重复绑定
els.fileTbody.addEventListener("click", (e) => {
  const view = e.target.closest(".file-view");
  if (view) {
    e.stopPropagation();
    const f = state.files.find((x) => x.id === view.dataset.id);
    if (f) openPreview(f);
    return;
  }
  const del = e.target.closest(".file-del");
  if (del) {
    e.stopPropagation();
    removeFiles([del.dataset.id]);
    return;
  }
  if (e.target.closest(".c-check")) return;  // 勾选交给 change 事件
  const tr = e.target.closest("tr[data-id]");
  if (!tr) return;
  const f = state.files.find((x) => x.id === tr.dataset.id);
  if (!f) return;
  if (e.ctrlKey || e.metaKey) {
    setSelected(f.id, !state.selected.has(f.id));
    renderFiles();
    return;
  }
  openPreview(f);
});

els.fileTbody.addEventListener("change", (e) => {
  const cb = e.target.closest(".row-check");
  if (!cb) return;
  const tr = cb.closest("tr[data-id]");
  if (!tr) return;
  setSelected(tr.dataset.id, cb.checked);
  tr.classList.toggle("selected", cb.checked);
  syncFileSelection();
});

els.checkAll.addEventListener("change", () => {
  if (els.checkAll.checked) state.files.forEach((f) => state.selected.add(f.id));
  else state.selected.clear();
  renderFiles();
});

function addFiles(fileList) {
  if (state.running) { toast("转换进行中，请先取消或等待完成"); return 0; }
  const exts = ["png", "jpg", "jpeg", "bmp", "tif", "tiff", "webp", "pdf"];
  let added = 0, tooBig = 0, badType = 0;
  for (const file of fileList) {
    const ext = (file.name.split(".").pop() || "").toLowerCase();
    if (!exts.includes(ext)) { badType++; continue; }
    if (file.size > MAX_FILE_BYTES) { tooBig++; continue; }
    state.files.push({
      id: "f" + Math.random().toString(36).slice(2, 9),
      name: file.name,
      size: file.size,
      tag: "pending",
      rotate: 0,
      file,
    });
    added++;
  }
  if (added) log(`已添加 ${added} 个文件。`);
  // 一次拖入很多文件时，把“超限”和“格式不支持”都说清楚；
  // 否则只跳过不提示，用户会以为文件已经加进去了（常见于 iPhone 的 HEIC）。
  const notes = [];
  if (tooBig) notes.push(`${tooBig} 个文件超过单文件上限 100 MB，已跳过`);
  if (badType) notes.push(`${badType} 个文件格式不支持，已跳过（支持 PNG/JPG/BMP/TIFF/WEBP/PDF）`);
  if (!added && !notes.length) notes.push("没有可添加的图片或 PDF");
  if (notes.length) toast(notes.join("；"));
  renderFiles();
  return added;
}

function addDiskFiles(items) {
  if (state.running) { toast("转换进行中，请先取消或等待完成"); return; }
  if (!items.length) return;
  for (const item of items) {
    state.files.push({
      id: "f" + Math.random().toString(36).slice(2, 9),
      name: item.name,
      size: item.size,
      tag: "pending",
      rotate: 0,
      path: item.path,
    });
  }
  log(`已添加 ${items.length} 个文件。`);
  renderFiles();
}

async function pickFilesFromDisk() {
  try {
    const r = await fetch("/api/pick-files", { method: "POST" });
    const d = await r.json();
    if (!d.ok) { toast("系统选择框不可用，改用普通选择"); return false; }
    if (d.skipped && d.skipped.length) {
      toast(`已跳过：${d.skipped.slice(0, 2).join("；")}${d.skipped.length > 2 ? "…" : ""}`);
    }
    if (d.files && d.files.length) addDiskFiles(d.files);
    return true;
  } catch {
    return false;
  }
}

function setFileTag(id, tag) {
  const f = state.files.find((x) => x.id === id);
  if (f) { f.tag = tag; renderFiles(); }
}

/* ---------------- 拖拽（含文件夹递归） ---------------- */
function collectDropped(dataTransfer) {
  const items = dataTransfer.items;
  if (!items || !items[0] || !items[0].webkitGetAsEntry) {
    return Promise.resolve([...dataTransfer.files]);
  }
  const entries = [...items].map((i) => i.webkitGetAsEntry()).filter(Boolean);
  const files = [];
  const walk = (entry) => new Promise((resolve) => {
    if (entry.isFile) {
      entry.file((f) => { files.push(f); resolve(); }, resolve);
    } else if (entry.isDirectory) {
      const reader = entry.createReader();
      const readBatch = () => reader.readEntries(async (list) => {
        if (!list.length) { resolve(); return; }
        await Promise.all([...list].map(walk));
        readBatch();
      }, resolve);
      readBatch();
    } else resolve();
  });
  return Promise.all(entries.map(walk)).then(() => files);
}

["dragenter", "dragover"].forEach((ev) =>
  els.paneFiles.addEventListener(ev, (e) => {
    e.preventDefault(); e.stopPropagation();
    els.dropzone.classList.add("dragover");
  })
);
["dragleave", "drop"].forEach((ev) =>
  els.paneFiles.addEventListener(ev, (e) => {
    e.preventDefault(); e.stopPropagation();
    els.dropzone.classList.remove("dragover");
  })
);
els.paneFiles.addEventListener("drop", (e) => {
  collectDropped(e.dataTransfer).then(addFiles);
});

/* 拖拽兜底：文件拖到窗口任意位置都按"添加到列表"处理。窗口其它区域原本没有
   拦截 drop，WebView 会直接把图片当成网页打开，把当前界面顶掉。 */
const dragHasFiles = (e) =>
  [...((e.dataTransfer && e.dataTransfer.types) || [])].includes("Files");

// 整窗拖拽提示：拖文件经过窗口任意位置都显示，离开或松手后收起
let dragOverlayTimer = null;
function showDragOverlay() {
  if (!els.dragOverlay) return;
  if (els.dragOverlay.hidden) els.dragOverlay.hidden = false;
  hidePasteChip();
  clearTimeout(dragOverlayTimer);
  // dragover 连续触发时会不断刷新；拖出窗口后不再触发，靠定时器收起
  dragOverlayTimer = setTimeout(() => { els.dragOverlay.hidden = true; }, 320);
}
// 用捕获阶段监听：文件页的 dragenter/over 会 stopPropagation，冒泡到不了 window
window.addEventListener("dragover", (e) => {
  if (!dragHasFiles(e)) return;
  showDragOverlay();
}, true);
window.addEventListener("dragenter", (e) => {
  if (!dragHasFiles(e)) return;
  showDragOverlay();
}, true);
window.addEventListener("dragleave", (e) => {
  // relatedTarget 为空表示真的离开窗口；在窗口内部元素间移动不收起
  if (e.relatedTarget) return;
  clearTimeout(dragOverlayTimer);
  if (els.dragOverlay) els.dragOverlay.hidden = true;
}, true);
window.addEventListener("drop", () => {
  clearTimeout(dragOverlayTimer);
  if (els.dragOverlay) els.dragOverlay.hidden = true;
}, true);

window.addEventListener("dragenter", (e) => {
  if (dragHasFiles(e)) e.preventDefault();
});
window.addEventListener("dragover", (e) => {
  if (!dragHasFiles(e)) return;
  e.preventDefault();
  if (e.dataTransfer) e.dataTransfer.dropEffect = "copy";
});
window.addEventListener("drop", (e) => {
  if (!dragHasFiles(e)) return;
  e.preventDefault();   // 文件页有自己的 drop（含文件夹递归），那里已 stopPropagation
  els.dropzone.classList.remove("dragover");
  collectDropped(e.dataTransfer).then(addFiles);
});
els.dropzone.addEventListener("click", () => els.btnAdd.click());
els.dropzone.addEventListener("keydown", (e) => {
  if (e.key === "Enter" || e.key === " ") { e.preventDefault(); els.btnAdd.click(); }
});
els.fileInput.addEventListener("change", () => {
  addFiles([...els.fileInput.files]);
  els.fileInput.value = "";
});

els.btnAdd.addEventListener("click", async () => {
  if (state.running) { toast("转换进行中，请先取消或等待完成"); return; }
  // 优先用本机系统选择框：能拿到真实路径，才能"查看原图 / 复制文件 / 删除文件"
  const ok = await pickFilesFromDisk();
  if (!ok) els.fileInput.click();
});
els.btnRemove.addEventListener("click", () => removeFiles([...state.selected]));
els.btnClear.addEventListener("click", () => {
  if (state.running) return;
  if (!state.files.length) return;
  const n = state.files.length;
  state.files = []; state.selected.clear();
  log(`已清空列表（${n} 个文件，未删除磁盘原文件）。`);
  renderFiles();
});

/* ---------------- 剪贴板粘贴图片 ---------------- */
const CLIP_EXT = {
  "image/png": "png", "image/jpeg": "jpg", "image/jpg": "jpg",
  "image/bmp": "bmp", "image/webp": "webp", "image/tiff": "tif",
  "application/pdf": "pdf",
};

function clipboardRawFiles(dt) {
  if (!dt) return [];
  if (dt.files && dt.files.length) return [...dt.files];
  if (!dt.items) return [];
  return [...dt.items]
    .filter((it) => it.kind === "file")
    .map((it) => it.getAsFile())
    .filter(Boolean);
}

// 剪贴板图片往往没有有意义的文件名（image.png），统一改成带时间戳的名字
let clipSeq = 0;
function stampClipboardFiles(files) {
  const d = new Date();
  const p = (n) => String(n).padStart(2, "0");
  const stamp = `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}${p(d.getSeconds())}`;
  return files.map((f) => {
    const ext = CLIP_EXT[(f.type || "").toLowerCase()]
      || (f.name.split(".").pop() || "").toLowerCase();
    if (!ext) return null;
    const name = `剪贴板图片_${stamp}_${++clipSeq}.${ext}`;
    try {
      return new File([f], name, { type: f.type || "" });
    } catch {
      return f;
    }
  }).filter(Boolean);
}

function addClipboardFiles(files) {
  if (!files.length) return 0;
  const added = addFiles(stampClipboardFiles(files));
  if (added) {
    toast(`已从剪贴板添加 ${added} 个文件`);
    if (state.tab !== "files") applyTab("files");
  }
  return added;
}

/* 粘贴统一入口：窗口任何位置都能用，不依赖焦点落在哪个控件上。
   1) 能读到剪贴板文件时走 paste 事件（最快，浏览器原生能力）；
   2) WebView2 在焦点不在输入框时不一定派发 paste 事件，从资源管理器复制的
      文件（CF_HDROP）也未必暴露给 JS，因此再由后端读一次系统剪贴板兜底。 */
let lastPasteAt = 0;
let nativePasteBusy = false;

const isEditableTarget = (el) =>
  !!el && (el.isContentEditable || /^(INPUT|TEXTAREA)$/.test(el.tagName || ""));

async function pasteNativeClipboard() {
  if (nativePasteBusy || state.running) return 0;
  nativePasteBusy = true;
  try {
    const r = await fetch("/api/clipboard-paste", { method: "POST" });
    const d = await r.json().catch(() => ({}));
    const files = (d && d.files) || [];
    if (!files.length) return 0;
    addDiskFiles(files);
    const notes = [`已从剪贴板添加 ${files.length} 个文件`];
    if (d.skipped && d.skipped.length) {
      notes.push(`已跳过 ${d.skipped.length} 个（${d.skipped[0]}）`);
    }
    toast(notes.join("；"));
    if (state.tab !== "files") applyTab("files");
    return files.length;
  } catch {
    return 0;   // 后端不可用时静默，交给 navigator.clipboard
  } finally {
    nativePasteBusy = false;
  }
}

window.addEventListener("paste", (e) => {
  lastPasteAt = Date.now();
  const files = clipboardRawFiles(e.clipboardData);
  if (files.length) {
    e.preventDefault();
    addClipboardFiles(files);
    return;
  }
  // 输入框里的纯文本粘贴交回浏览器（单号、搜索词等）
  if (isEditableTarget(e.target)) return;
  pasteNativeClipboard();
}, true);

/* 兜底：按下 Ctrl+V 后短暂等待 paste 事件，没等到就直接读系统剪贴板。 */
document.addEventListener("keydown", (e) => {
  if (!(e.ctrlKey || e.metaKey) || e.altKey || e.shiftKey) return;
  if (String(e.key || "").toLowerCase() !== "v") return;
  const pressedAt = Date.now();
  setTimeout(() => {
    if (lastPasteAt >= pressedAt) return;
    pasteNativeClipboard();
  }, 300);
}, true);

/* ---------------- 鼠标移入：激活窗口 + 剪贴板粘贴提示 ---------------- */
const desktopBridge = () => (window.pywebview && window.pywebview.api) || null;
let pasteChipTimer = null;
let pastePeekToken = 0;

function hidePasteChip() {
  if (!els.pasteChip) return;
  els.pasteChip.hidden = true;
  clearTimeout(pasteChipTimer);
  pasteChipTimer = null;
}

function showPasteChip(count) {
  if (!els.pasteChip) return;
  els.pasteChipText.textContent = count > 1
    ? `剪贴板里有 ${count} 个文件，点击粘贴`
    : "剪贴板里有图片，点击粘贴";
  els.pasteChip.hidden = false;
  clearTimeout(pasteChipTimer);
  pasteChipTimer = setTimeout(hidePasteChip, 15000);
}

// 只探测剪贴板有没有可识别的图片/文件，不落盘；鼠标进入窗口时问一次
async function peekClipboardForChip() {
  if (!els.pasteChip || state.running || !els.pasteChip.hidden) return;
  const token = ++pastePeekToken;
  try {
    const r = await fetch("/api/clipboard-peek", { method: "POST" });
    const d = await r.json();
    if (token !== pastePeekToken || state.running) return;
    if (d && d.ok && d.has) showPasteChip(Number(d.count) || 1);
  } catch { /* 剪贴板不可读时静默 */ }
}

document.addEventListener("mouseenter", () => {
  // 桌面壳：鼠标移入窗口即激活，随后直接按 Ctrl+V 才能落到本程序
  const api = desktopBridge();
  if (api && typeof api.win_activate === "function") {
    try { api.win_activate(); } catch { /* 忽略 */ }
  }
  peekClipboardForChip();
}, { passive: true });

els.pasteChip.addEventListener("click", async (e) => {
  if (e.target.closest(".paste-chip-close")) return;
  hidePasteChip();
  if (state.running) { toast("转换进行中，请等待完成后再粘贴"); return; }
  const n = await pasteNativeClipboard();
  if (!n) toast("剪贴板里没有可识别的图片或文件");
});

els.pasteChipClose.addEventListener("click", (e) => {
  e.stopPropagation();
  hidePasteChip();
});

// 桌面壳全局快捷键调起：窗口在后台/托盘也能把剪贴板图片贴进来
window.__pasteFromHotkey = async () => {
  if (state.running) { toast("转换进行中，请等待完成后再粘贴"); return; }
  hidePasteChip();
  const n = await pasteNativeClipboard();
  if (!n) toast("剪贴板里没有可识别的图片或文件");
  else if (state.tab !== "files") applyTab("files");
};

els.btnPaste.addEventListener("click", async () => {
  if (state.running) { toast("转换进行中，请先取消或等待完成"); return; }
  // 桌面程序优先直接读系统剪贴板，不依赖浏览器剪贴板权限
  if (await pasteNativeClipboard()) return;
  if (!navigator.clipboard || !navigator.clipboard.read) {
    toast("剪贴板里没有图片，先截图或复制图片再试");
    return;
  }
  try {
    const items = await navigator.clipboard.read();
    const files = [];
    for (const item of items) {
      const type = (item.types || []).find((t) => CLIP_EXT[t]);
      if (!type) continue;
      const blob = await item.getType(type);
      files.push(new File([blob], `剪贴板图片.${CLIP_EXT[type]}`, { type }));
    }
    if (!files.length) { toast("剪贴板里没有图片，先截图或复制图片再试"); return; }
    addClipboardFiles(files);
  } catch {
    toast("剪贴板里没有图片，先截图或复制图片再试");
  }
});

/* ---------------- 页签切换 ---------------- */
els.tabs.addEventListener("click", (e) => {
  const btn = e.target.closest(".tab");
  if (!btn || !btn.dataset.tab) return;
  applyTab(btn.dataset.tab);
});

// 键盘操作：方向键 / Home / End 在页签间移动并即时切换（WAI-ARIA tabs 自动激活模式）
els.tabs.addEventListener("keydown", (e) => {
  const btn = e.target.closest(".tab");
  if (!btn || !btn.dataset.tab) return;
  const cur = TAB_ORDER.indexOf(btn.dataset.tab);
  if (cur < 0) return;
  let next = cur;
  if (e.key === "ArrowRight" || e.key === "ArrowDown") next = (cur + 1) % TAB_ORDER.length;
  else if (e.key === "ArrowLeft" || e.key === "ArrowUp") next = (cur - 1 + TAB_ORDER.length) % TAB_ORDER.length;
  else if (e.key === "Home") next = 0;
  else if (e.key === "End") next = TAB_ORDER.length - 1;
  else if (e.key === "Enter" || e.key === " ") {
    e.preventDefault();
    applyTab(TAB_ORDER[cur]);
    TAB_BTN()[TAB_ORDER[cur]].focus();
    return;
  } else return;
  e.preventDefault();
  const target = TAB_ORDER[next];
  applyTab(target);
  TAB_BTN()[target].focus();
});
els.btnFlagsOpen.addEventListener("click", () => {
  if (state.flags.length) showFlagReview(state.flags);
  else toast("本次转换没有待复核项");
});

els.btnFlagsModal.addEventListener("click", () => {
  if (state.flags.length) showFlagReview(state.flags);
});

// 待复核筛选：文件 / 优先级 / 关键字；筛选只影响显示，进度始终按全量统计
els.flagFileFilter.addEventListener("change", () => renderFlagPane(state.flags));
els.flagPrioFilter.addEventListener("change", () => renderFlagPane(state.flags));
let flagSearchTimer = null;
els.flagSearch.addEventListener("input", () => {
  clearTimeout(flagSearchTimer);
  flagSearchTimer = setTimeout(() => renderFlagPane(state.flags), 200);
});

/* ---------------- 保存位置 ---------------- */
els.btnBrowse.addEventListener("click", async () => {
  try {
    const r = await fetch("/api/default-output");
    const d = await r.json();
    els.outPath.value = d.path;
    toast("已填入默认保存路径（网页版路径可直接编辑）");
  } catch { toast("获取默认路径失败"); }
});

els.btnOpenDir.addEventListener("click", async () => {
  const dir = els.outPath.value.replace(/\\[^\\]*$/, "");
  // 桌面壳里可以直接打开资源管理器，比"复制路径让用户自己粘贴"少两步
  if (dir && await revealPath(dir)) { toast("已在资源管理器中打开"); return; }
  try {
    await navigator.clipboard.writeText(dir);
    toast(`目录已复制到剪贴板：${dir}`);
  } catch { toast(`保存目录：${dir}`); }
});

/* ---------------- 进度与状态 ---------------- */
function setStatus(text, cls = "") {
  els.statusMsg.textContent = text;
  els.statusText.className = "status-text" + (cls ? " " + cls : "");
}

// 长批次（45 张要 8 分钟）没有时间感，加一个"已用 mm:ss"
let elapsedTimer = null;
let elapsedStart = 0;
function startElapsed() {
  elapsedStart = Date.now();
  els.elapsed.hidden = false;
  els.elapsed.textContent = "已用 0:00";
  clearInterval(elapsedTimer);
  elapsedTimer = setInterval(() => {
    const sec = Math.max(0, Math.floor((Date.now() - elapsedStart) / 1000));
    els.elapsed.textContent =
      `已用 ${Math.floor(sec / 60)}:${String(sec % 60).padStart(2, "0")}`;
  }, 1000);
}
function stopElapsed() { clearInterval(elapsedTimer); elapsedTimer = null; }
function setProgress(done, total, indeterminate) {
  els.progressTrack.classList.toggle("indeterminate", !!indeterminate);
  if (indeterminate) { els.progressPct.textContent = ""; return; }
  const pct = total ? Math.round((done / total) * 100) : 0;
  els.progressFill.style.width = pct + "%";
  els.progressPct.textContent = pct + "%";
}
function refreshStartButton() {
  const query = (els.waybillInput?.value || "").replace(/[^a-z0-9]/gi, "");
  // 没有文件时按钮保持可点：点击后直接打开文件选择框，避免用户找不到入口
  els.btnStart.disabled = state.running
    || (state.mode === "find" && query.length < 6);
  els.btnCancel.disabled = !state.running;
}

/* ---------------- 转换流程 ---------------- */
async function start(onlyIds) {
  if (state.running || state.files.length === 0) return;
  // 支持"只重试失败文件"：传一个 id 集合进来就只提交这些文件，
  // 其余文件保留在列表里，不参与本次转换
  const targets = onlyIds instanceof Set
    ? state.files.filter((f) => onlyIds.has(f.id))
    : state.files;
  if (targets.length === 0) {
    toast("没有可转换的文件");
    return;
  }
  const query = (els.waybillInput.value || "").trim();
  const out = els.outPath.value.trim();
  if (state.mode === "find" && query.replace(/[^a-z0-9]/gi, "").length < 6) {
    toast("请粘贴完整单号");
    els.waybillInput.focus();
    return;
  }
  if (state.mode !== "find" && !out.toLowerCase().endsWith(".xlsx")) {
    toast("保存路径必须以 .xlsx 结尾");
    return;
  }
  state.running = true;
  refreshStartButton();
  // 新一轮开跑先清场：结果卡/质量面板/失败条/待复核三组折叠/页签计数一起归零，
  // 避免上一轮的数字与表格残留在这一轮界面上。
  resetRunOutput();
  // 开跑时回到「文件」页签：本轮的看点在这里，避免停在已清空的待复核页
  applyTab("files");
  startElapsed();
  setStatus(state.mode === "find" ? "正在准备快速查找…" : "正在准备识别引擎…", "run");
  setProgress(0, targets.length, false);
  const modeName = MODES.find((m) => m.value === state.mode).name;
  const layoutName = (LAYOUTS.find((l) => l.value === state.layout) || LAYOUTS[0]).name;
  log(`开始转换 ${targets.length} 个文件，模式：${modeName}`
      + (state.mode === "full" ? `，表格输出：${layoutName}` : ""));

  const fd = new FormData();
  const plans = [];
  let uploadIndex = 0;
  for (const f of targets) {
    const rotate = Number(f.rotate) || 0;
    if (f.path) {
      // 磁盘文件直接给路径：免上传，且能预览/复制/删除
      plans.push({ id: f.id, name: f.name, path: f.path, rotate });
    } else {
      fd.append("files", f.file, f.name);
      plans.push({ id: f.id, name: f.name, upload: uploadIndex++, rotate });
    }
  }
  fd.append("items", JSON.stringify(plans));
  fd.append("mode", state.mode);
  fd.append("layout", state.layout);
  fd.append("output", out);
  if (state.mode === "find") fd.append("target", query);

  try {
    const r = await fetch("/api/convert", { method: "POST", body: fd });
    const d = await r.json();
    if (!r.ok) throw new Error(d.detail || "启动失败");
    state.jobId = d.job_id;
    if (d.renamed_from) {
      // 目标文件已存在：后端已自动另存为 "name (n).xlsx"，旧文件不动
      log(`目标文件已存在，已自动另存为：${d.out_path}`);
      toast("目标文件已存在，结果已自动改名保存");
    }
    openSocket(d.job_id);
  } catch (err) {
    fail(err.message || String(err));
  }
}

function openSocket(jobId) {
  const proto = location.protocol === "https:" ? "wss" : "ws";
  const ws = new WebSocket(`${proto}://${location.host}/ws/${jobId}`);
  state.ws = ws;
  ws.onmessage = (ev) => {
    let d;
    try { d = JSON.parse(ev.data); } catch { return; }
    handleEvent(d);
  };
  ws.onclose = () => {
    if (state.ws !== ws) return;
    state.ws = null;
    // 正常结束（done/error/cancelled）会先 resetRunning 再关连接；
    // 若此时仍在运行，说明连接异常断开，必须让界面回到可操作状态，
    // 否则进度条会一直停在“转换中”、开始按钮一直是禁用的。
    if (state.running) {
      fail("与本地识别服务的连接中断，本次任务可能已停止。请重新开始转换。");
    }
  };
}

function handleEvent(d) {
  switch (d.type) {
    case "log": log(d.text); break;
    case "phase": setStatus(d.text, "run"); break;
    case "progress":
      setProgress(d.done, d.total, d.indeterminate);
      break;
    case "file_status":
      setFileTag(d.id, d.tag);
      break;
    case "done":
      finish(d);
      break;
    case "error":
      fail(d.message);
      break;
    case "cancelled":
      state.pendingSearchQuery = "";
      setStatus("已取消转换", "warn");
      log("已取消。");
      resetRunning();
      showResult("info", "已取消转换。可重新点击「开始转换」。");
      break;
  }
}

function finish(d) {
  // 结果面板在「文件清单」下方：完成时回到该页签，避免用户停在日志页看不到结果
  applyTab("files");
  setProgress(100, 100, false);
  if (d.mode === "find") {
    setStatus("查找完成", "ok");
    log(`单号查找完成：扫描 ${d.scanned || 0} 张图片/页，用时 ${d.elapsed || 0} 秒。`);
    if (d.exact_matches) log(`完整命中 ${d.exact_matches} 处（无需复核）。`);
    if (d.review_matches) log(`后四位命中 ${d.review_matches} 处，请打开原图核对前几位。`);
    if (d.suspect_matches) log(`另有 ${d.suspect_matches} 处疑似 OCR 错读，请对照原图确认。`);
    if (!d.matches) log("快速识别没有匹配项；这不代表单号一定不存在，可用「图片转 Excel」进一步核对。");
    for (const w of d.warnings || []) log("提示：" + w);
    state.lastOutPath = "";
    state.flags = [];
    state.flagStates = {};
    renderTabs();
    renderFlagPane([]);
    showDoneCard(d);
    resetRunning();
    return;
  }
  setStatus("转换完成", "ok");
  const isWaybill = d.mode === "waybill";
  log(`已导出：${d.out_path}`);
  for (const w of d.warnings || []) log("提示：" + w);
  if (isWaybill) {
    log(`纯单号：共识别 ${d.waybills || 0} 条候选，去重后导出 ${d.unique_waybills || 0} 个；` +
        `已确认 ${d.confirmed_waybills || 0} 个，待核对 ${d.pending_waybills || 0} 个，` +
        `格式异常/非YT ${d.format_invalid_waybills || 0} 条未写入。`);
    if (d.pending_waybills) {
      log("注意：待核对单号已写入清单，请对照原图复核后再使用。");
    }
  } else if (d.waybills) {
    log(`单号：已确认 ${d.confirmed_waybills || 0} 处，待核对 ${d.pending_waybills || 0} 处，` +
        `格式异常/非YT ${d.format_invalid_waybills || 0} 处。`);
    if (d.pending_waybills || d.format_invalid_waybills) {
      log("只有「已确认」的 YT 单号可直接采信；待核对/格式异常项请对照原图，不能当作正确单号使用。");
    }
  }
  if (d.index_issues) {
    log(isWaybill
      ? `另有 ${d.index_issues} 页未提取到单号，请对照原图核查。`
      : `索引异常 ${d.index_issues} 项，详见 Excel 中的「索引异常」工作表。`);
  }
  if (d.flagged && !isWaybill) {
    log(`有 ${d.flagged} 处待复核，已在 Excel 数据页黄底标出。`);
  }
  if (d.failed) {
    log(`有 ${d.failed} 个文件未成功识别，可点结果卡片里的「重试失败文件」只重跑这几张。`);
  }

  state.lastOutPath = d.out_path || "";
  // 页签计数与待复核面板用的是 state.flags；showDoneCard 只负责结果卡，
  // 这里同步一次，避免「结果卡写了 N 处待复核、页签却是空的」。
  state.flags = isWaybill ? [] : (d.flags || []);
  state.flagStates = {};
  els.flagFileFilter.value = "";
  els.flagPrioFilter.value = "";
  els.flagSearch.value = "";
  renderTabs();
  renderFlagPane(state.flags);
  showDoneCard(d);
  resetRunning();
  if (state.autoOpen && d.out_path) {
    // 设置里开启「完成后自动打开 Excel」：让默认程序拉起刚导出的文件
    setTimeout(() => revealPath(d.out_path, "open"), 150);
  }
  if (state.pendingSearchQuery) {
    const query = state.pendingSearchQuery;
    state.pendingSearchQuery = "";
    if (d.searchable && d.waybills) setTimeout(() => openSearch(query), 0);
    else log("完整识别结果中没有可查询的单号，请在 Excel 中进一步核对。");
  }
}

/* 只重试上次失败的文件：列表里其余文件保持原状，结果另存为新工作簿 */
function retryFailed(items) {
  if (state.running) return;
  const ids = new Set((items || []).map((it) => it.id));
  const targets = state.files.filter((f) => ids.has(f.id));
  if (targets.length === 0) {
    toast("没找到可重试的文件，可能已被移除");
    return;
  }
  for (const f of targets) f.tag = "pending";
  renderFiles();
  log(`重试上次失败的 ${targets.length} 个文件；其余文件保持原状态，结果会另存为新文件。`);
  toast("正在重试失败文件，结果将另存为新文件");
  start(ids);
}

/* 转换完成：结构化结果卡片（摘要 / 输出文件 / 统计 / 下一步操作） */
function scrollResultIntoView() {
  const card = els.resultCard;
  if (!card || card.hidden) return;
  // 兼容两种落位：结果卡在右栏时滚右栏；在主区时文件列表会让出高度
  const rail = card.closest(".rail");
  if (rail) {
    const over = card.getBoundingClientRect().bottom
        - rail.getBoundingClientRect().bottom;
    if (over > 0) rail.scrollTop += Math.ceil(over) + 8;
    return;
  }
  const pane = card.closest(".pane");
  if (!pane) return;
  const over = card.getBoundingClientRect().bottom - pane.getBoundingClientRect().bottom;
  if (over > 0) card.scrollIntoView({ block: "nearest" });
}

function showDoneCard(d) {
  if (d.mode === "find") {
    const exact = d.exact_matches || 0;
    const review = d.review_matches || 0;
    const suspect = d.suspect_matches || 0;
    const needCheck = review + suspect;
    const failed = d.failed || 0;
    els.resultCard.hidden = false;
    els.resultCard.className = "result-card " + (exact ? "ok" : needCheck ? "warn" : "info");
    els.rcTitle.textContent = exact ? "找到单号" : needCheck ? "找到待复核单号" : "没有命中";
    els.rcSubtitle.textContent = `完整命中 ${exact} 处 · 待复核 ${needCheck} 处 · 已扫描 ${d.scanned || 0} 张`;
    els.rcFile.hidden = true;
    els.rcMetrics.hidden = false;
    els.statOkLabel.textContent = "完整命中";
    els.statOk.textContent = exact;
    els.statWarnLabel.textContent = "待复核";
    els.statWarn.textContent = needCheck;
    els.statBadLabel.textContent = "识别失败";
    els.statBad.textContent = failed;
    const notes = [];
    if (review) notes.push(`后四位命中 ${review} 处已计入结果，请打开原图核对前几位`);
    if (suspect) notes.push(`疑似错读 ${suspect} 处需要打开原图核对，不能直接采信`);
    if (!exact && !needCheck) notes.push("快速识别没有找到匹配项，不代表单号不存在，可切换到「图片转 Excel」进一步核对");
    if (failed) notes.push(`${failed} 张图片未能识别`);
    els.rcNote.textContent = notes.join("；");
    els.rcNote.hidden = notes.length === 0;
    els.rcActions.hidden = false;
    for (const button of [els.btnOpen, els.btnReveal, els.btnFlags, els.btnDownload]) button.hidden = true;
    els.btnSearch.hidden = false;
    els.btnSearch.textContent = "查看命中位置";
    els.btnSearch.onclick = () => openSearch(d.query || "");
    els.btnFullCheck.hidden = false;
    els.btnFullCheck.onclick = () => {
      state.pendingSearchQuery = d.query || "";
      state.mode = "full";
      renderModes();
      syncModeControls();
      start();
    };
    const failedItems = Array.isArray(d.failed_items) ? d.failed_items : [];
    els.btnRetryFailed.hidden = failedItems.length === 0;
    if (failedItems.length) {
      els.btnRetryFailed.textContent = `重试失败文件（${failedItems.length}）`;
      els.btnRetryFailed.onclick = () => retryFailed(failedItems);
    }
    els.rcQuality.hidden = true;
    els.rcFails.hidden = true;
    els.rcFiles.hidden = true;
    if (d.history) {
      // 历史快照不提供「查看命中位置/重新识别」：内存索引可能已不在本任务上
      els.rcTitle.textContent = "历史任务 · 查找单号";
      els.btnSearch.hidden = true;
      els.btnFullCheck.hidden = true;
      els.btnRetryFailed.hidden = true;
    }
    scrollResultIntoView();
    return;
  }
  const isWaybill = d.mode === "waybill";
  const out = String(d.out_path || "");
  const fileName = out.split(/[\\/]/).pop() || "导出文件";
  const dir = out.replace(/[\\/][^\\/]*$/, "");

  els.resultCard.hidden = false;
  els.resultCard.className = "result-card ok";
  els.rcTitle.textContent = "转换完成";
  els.rcSubtitle.textContent = isWaybill
    ? `已提取 ${d.unique_waybills || 0} 个运单号（自动去重）`
    : `文档数据已写入 Excel · 工作表 ${d.sheets || 0} 个 / 表格 ${d.tables || 0} 个`;
  if (d.quality && d.quality.level) {
    els.rcSubtitle.textContent += ` · 拍摄质量：${d.quality.level}`;
  }
  els.rcFile.hidden = !out;
  els.rcFileName.textContent = fileName;
  els.rcFileName.title = out;
  els.rcMetrics.hidden = false;

  if (isWaybill) {
    els.statOkLabel.textContent = "导出单号";
    els.statWarnLabel.textContent = "已确认";
    els.statBadLabel.textContent = "待核对";
    els.statOk.textContent = d.unique_waybills || 0;
    els.statWarn.textContent = d.confirmed_waybills || 0;
    els.statBad.textContent = d.pending_waybills || 0;
  } else {
    els.statOkLabel.textContent = "已确认";
    els.statWarnLabel.textContent = "待核对";
    els.statBadLabel.textContent = "格式异常";
    els.statOk.textContent = d.confirmed_waybills || 0;
    els.statWarn.textContent = d.pending_waybills || 0;
    els.statBad.textContent = d.format_invalid_waybills || 0;
  }

  const notes = [];
  if (isWaybill) {
    if (d.format_invalid_waybills) {
      notes.push(`格式异常/非 YT ${d.format_invalid_waybills} 条未纳入清单`);
    }
    if (d.index_issues) notes.push(`${d.index_issues} 页未提取到运单号`);
  } else {
    if (d.flags && d.flags.length) notes.push(`${d.flags.length} 处待复核已在 Excel 中标黄`);
    if (d.index_issues) notes.push(`${d.index_issues} 项索引异常，详见「索引异常」工作表`);
  }
  const failedItems = Array.isArray(d.failed_items) ? d.failed_items : [];
  const failedCount = failedItems.length || d.failed || 0;
  if (failedCount) notes.push(`${failedCount} 个文件未识别成功`);
  els.rcNote.textContent = notes.join("；");
  els.rcNote.hidden = notes.length === 0;
  els.rcActions.hidden = false;

  els.btnOpen.hidden = !out;
  els.btnOpen.onclick = async () => {
    if (await revealPath(out, "open")) toast("已用默认程序打开");
    else toast("打开失败，请用「下载」或到保存目录查看");
  };
  els.btnReveal.hidden = !dir;
  els.btnReveal.onclick = async () => {
    if (await revealPath(dir, "select")) toast("已在资源管理器中定位文件");
    else toast("打开目录失败，请手动到保存位置查看");
  };
  els.btnDownload.hidden = !d.download_url;
  els.btnDownload.onclick = () => {
    if (d.download_url) location.href = d.download_url;
  };

  // 纯单号只有一个清单工作表，没有黄底单元格和待复核页，不提供该入口
  els.btnFlags.hidden = isWaybill || !(d.flags && d.flags.length);
  if (!isWaybill && d.flags && d.flags.length) {
    els.btnFlags.textContent = `待复核（${d.flags.length}）`;
    els.btnFlags.onclick = () => applyTab("flags");
  }
  els.btnFullCheck.hidden = true;
  els.btnSearch.hidden = !(d.searchable && d.waybills > 0);
  els.btnSearch.textContent = `单号查询（${d.waybills || 0}）`;
  els.btnSearch.onclick = () => openSearch(els.waybillInput.value.trim());

  // 只重试失败文件：保留已成功的成果，新结果按重名规则另存
  els.btnRetryFailed.hidden = failedItems.length === 0;
  if (failedItems.length) {
    els.btnRetryFailed.textContent = `重试失败文件（${failedItems.length}）`;
    els.btnRetryFailed.title = failedItems
      .slice(0, 3)
      .map((it) => `${it.name}：${it.reason || "识别失败"}`)
      .join("\n");
    els.btnRetryFailed.onclick = () => retryFailed(failedItems);
  }

  // 质量面板与失败明细（含 retryable 区分）
  renderQuality(d);
  renderFail(failedItems);
  renderFileSummary(d);

  if (d.history) {
    // 历史快照：标明来源时间；查询/重试依赖当时的内存索引与文件清单，不提供
    els.rcTitle.textContent = "历史任务快照";
    els.rcSubtitle.textContent = `${d.time || ""} · ${els.rcSubtitle.textContent}`;
    els.btnSearch.hidden = true;
    els.btnFullCheck.hidden = true;
    els.btnRetryFailed.hidden = true;
  }

  // 内容全部填完再滚动，此时高度才准确
  scrollResultIntoView();
}

/* ---------------- 结果区：拍摄质量 / 失败文件 ---------------- */
const QUALITY_SEVERITY = { "良好": 0, "可用": 0, "需注意": 1, "建议重拍": 2 };
const QUALITY_CLASS = { "良好": "ok", "可用": "ok", "需注意": "warn", "建议重拍": "bad" };

function qualityRank(level) {
  const r = QUALITY_SEVERITY[String(level || "")];
  return typeof r === "number" ? r : -1;
}

// 按 entry.file 聚合，取最差档；同档取最低分（与后端任务级"最差页"口径一致）
function worstByFile(perFile) {
  const groups = new Map();
  for (const e of perFile) {
    if (!e) continue;
    const key = String(e.file || "（未知文件）");
    const cur = groups.get(key);
    if (!cur) {
      groups.set(key, { file: key, level: e.level || "", score: e.score, pages: 1 });
      continue;
    }
    cur.pages += 1;
    if (qualityRank(e.level) > qualityRank(cur.level)) cur.level = e.level || cur.level;
    if (typeof e.score === "number" && (typeof cur.score !== "number" || e.score < cur.score)) {
      cur.score = e.score;
    }
  }
  return [...groups.values()].sort(
    (a, b) => qualityRank(b.level) - qualityRank(a.level)
      || (typeof a.score === "number" ? a.score : 999) - (typeof b.score === "number" ? b.score : 999)
  );
}

function renderQuality(d) {
  const q = d && d.quality;
  const groups = q && Array.isArray(q.per_file) ? worstByFile(q.per_file) : [];
  els.rcQList.innerHTML = "";
  els.rcQList.hidden = true;
  els.rcQToggle.setAttribute("aria-expanded", "false");
  els.rcQToggle.hidden = groups.length === 0;
  els.rcQToggle.textContent = `每图明细（${groups.length}）`;
  els.rcQToggle.onclick = null;
  els.rcQSug.hidden = true;
  els.rcQSug.innerHTML = "";

  if (!q || !q.level) {
    els.rcQuality.hidden = true;
    els.rcQWhy.hidden = true;
    return;
  }
  els.rcQuality.hidden = false;
  const cls = QUALITY_CLASS[q.level] || "warn";
  els.rcQTier.className = `rcq-tier ${cls}`;
  els.rcQTier.textContent = q.level;
  els.rcQScore.className = `rcq-score ${cls}`;
  els.rcQScore.textContent = typeof q.score === "number" ? `${q.score} 分` : "—";
  const reasons = Array.isArray(q.reasons) ? q.reasons.filter(Boolean) : [];
  els.rcQWhy.hidden = reasons.length === 0;
  // 每条原因单独一行：292px 窄栏里分号串会挤成一坨，逐行更好读。
  // 只按顶层数组拆分，不再二次切分字符串，避免误伤原因内部的标点。
  els.rcQWhy.innerHTML = reasons
    .map((r) => `<span class="rcq-why-line">${esc(r)}</span>`)
    .join("");

  // 可执行建议：后端 suggestions 是「怎么改进拍摄」，最多展示 3 条避免挤满窄栏
  const suggestions = (Array.isArray(q.suggestions) ? q.suggestions : [])
    .filter(Boolean).slice(0, 3);
  els.rcQSug.hidden = suggestions.length === 0;
  els.rcQSug.innerHTML = suggestions.length
    ? `<span class="rcq-sug-title">怎么改进</span>` +
      suggestions.map((s) => `<span class="rcq-sug-line">· ${esc(s)}</span>`).join("")
    : "";

  if (!groups.length) return;
  for (const g of groups) {
    const row = document.createElement("div");
    row.className = `rcq-row ${QUALITY_CLASS[g.level] || "warn"}`;
    const scoreText = typeof g.score === "number" ? `${g.score} 分` : "—";
    const pages = g.pages > 1 ? `<i class="rcq-pages">${g.pages} 页 · 取最差</i>` : "";
    row.innerHTML =
      `<span class="rcq-file" title="${esc(g.file)}">${esc(g.file)}</span>${pages}` +
      `<span class="rcq-lv">${esc(g.level || "—")}</span><span class="rcq-sc">${scoreText}</span>`;
    // 点击该图可回到原图核对（拖入文件用内存对象预览，磁盘文件走本机服务）
    row.title = "点击查看原图";
    row.style.cursor = "pointer";
    row.addEventListener("click", (e) => {
      e.stopPropagation();
      viewQualityFile(g.file);
    });
    els.rcQList.appendChild(row);
  }
  els.rcQToggle.onclick = (e) => {
    e.stopPropagation();
    const open = els.rcQList.hidden;
    els.rcQList.hidden = !open;
    els.rcQToggle.setAttribute("aria-expanded", String(open));
    els.rcQToggle.textContent = open ? `收起明细（${groups.length}）` : `每图明细（${groups.length}）`;
  };
}

function renderFail(items) {
  const list = Array.isArray(items) ? items : [];
  els.rcFailList.innerHTML = "";
  els.rcFails.hidden = list.length === 0;
  if (!list.length) { els.rcFailCount.textContent = ""; return; }
  const retryable = list.filter((it) => it.retryable === true).length;
  const broken = list.filter((it) => it.retryable === false).length;
  const unknown = list.length - retryable - broken;
  const parts = [];
  if (retryable) parts.push(`${retryable} 个重传可能有效`);
  if (broken) parts.push(`${broken} 个文件本身有问题`);
  if (unknown) parts.push(`${unknown} 个未标注`);
  els.rcFailCount.textContent = `共 ${list.length} 个 · ${parts.join(" / ")}`;
  for (const it of list) {
    const tag = it.retryable === true
      ? ["retry", "重传可能有效"]
      : (it.retryable === false ? ["fault", "文件本身有问题"] : ["unknown", "未标注"]);
    const li = document.createElement("li");
    li.className = `rcf-item ${tag[0]}`;
    li.innerHTML =
      `<span class="rcf-name" title="${esc(it.name)}">${esc(it.name)}</span>` +
      `<span class="rcf-tag">${tag[1]}</span>` +
      `<span class="rcf-reason" title="${esc(it.reason || "")}">${esc(it.reason || "识别失败")}</span>`;
    els.rcFailList.appendChild(li);
  }
}

// 质量明细点某一图 → 打开原图核对；拖入文件走内存对象，磁盘文件走本机预览接口
function viewQualityFile(name) {
  const target = state.files.find((f) => f.name === name);
  if (!target) {
    toast("该文件已不在清单中，无法打开原图");
    return;
  }
  openPreview(target);
}

// 结果卡「按文件汇总」：单号 / 行 / 待复核 / 耗时，按待复核数降序
function renderFileSummary(d) {
  const list = Array.isArray(d && d.file_summary) ? d.file_summary.slice() : [];
  els.rcFilesBody.innerHTML = "";
  els.rcFilesCount.textContent = "";
  els.rcFiles.hidden = list.length === 0;
  if (!list.length) return;
  list.sort((a, b) => (Number(b.flags) || 0) - (Number(a.flags) || 0)
    || (Number(b.waybills) || 0) - (Number(a.waybills) || 0)
    || (Number(b.rows) || 0) - (Number(a.rows) || 0));
  els.rcFilesCount.textContent = `${list.length} 个文件 · 按待复核排序`;
  for (const it of list) {
    const tr = document.createElement("tr");
    const seconds = typeof it.seconds === "number" ? `${Math.round(it.seconds)}s` : "—";
    const pages = (Number(it.pages) || 0) > 1
      ? ` <i class="rcs-pages">${Number(it.pages)}页</i>` : "";
    const flags = Number(it.flags) || 0;
    tr.innerHTML =
      `<td class="rcs-file" title="${esc(it.file)}">${esc(it.file)}${pages}</td>` +
      `<td class="num">${Number(it.waybills) || 0}</td>` +
      `<td class="num">${Number(it.rows) || 0}</td>` +
      `<td class="num${flags ? " warn" : ""}">${flags}</td>` +
      `<td class="num">${seconds}</td>`;
    els.rcFilesBody.appendChild(tr);
  }
}

function esc(v) {
  return String(v ?? "").replace(/[&<>"']/g, (c) => (
    { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]
  ));
}

/* ---------------- 待复核：优先级分组（S5） ---------------- */
const FLAG_PRIORITIES = ["高", "中", "低"];
const FLAG_PRIO_LABEL = { "高": "高优先", "中": "中优先", "低": "低优先" };
const FLAG_PRIO_CLASS = { "高": "prio-hi", "中": "prio-mid", "低": "prio-low" };

// 后端给的是「高 / 中 / 低」；缺失或越界一律落到「中」并标记为未分级，不丢条、不改数。
function normalizeFlagPriority(flag) {
  const raw = String((flag && flag.priority) || "").trim();
  if (FLAG_PRIORITIES.includes(raw)) return { prio: raw, unrated: false };
  return { prio: "中", unrated: true };
}

// 三条组天然加和 = entries.length，不存在「凑数」的中间桶。
// entries 保留原始下标，用于标记「已核对 / 需重拍」并持久化到本机任务目录。
function groupFlags(entries) {
  const buckets = new Map(FLAG_PRIORITIES.map((p) => [p, []]));
  for (const entry of entries || []) {
    const { prio, unrated } = normalizeFlagPriority(entry.flag);
    buckets.get(prio).push({ ...entry, unrated });
  }
  return FLAG_PRIORITIES.map((p) => ({
    priority: p,
    label: FLAG_PRIO_LABEL[p],
    cls: FLAG_PRIO_CLASS[p],
    entries: buckets.get(p),
  }));
}

function flagSearchText(flag) {
  return [
    flag.file, flag.kind, flag.reason, flag.rebuilt, flag.whole,
    flag.ocr, flag.local, flag.zoom,
  ].map((v) => String(v ?? "")).join(" ").toLowerCase();
}

function renderFlagPane(items) {
  const list = items || [];
  if (!els.flagPaneWrap) return;
  els.flagTools.hidden = list.length === 0;
  els.flagPaneWrap.hidden = list.length === 0;
  // 有数据时收起「本次有 N 处待复核」引导卡，把高度让给筛选条与表格
  els.paneFlags.classList.toggle("has-flags", list.length > 0);
  if (list.length === 0) {
    els.flagBody.innerHTML = "";
    els.flagPaneSummary.textContent = "";
    els.flagProgress.textContent = "已处理 0 / 0";
    els.flagFileFilter.innerHTML = '<option value="">全部文件</option>';
    return;
  }

  // 文件下拉：保持当前选择；上一任务的文件不在本次清单里时回落到「全部」
  const files = [...new Set(list.map((f) => String(f.file || "")).filter(Boolean))].sort();
  const currentFile = els.flagFileFilter.value;
  els.flagFileFilter.innerHTML = '<option value="">全部文件</option>' +
    files.map((f) => `<option value="${esc(f)}">${esc(f)}</option>`).join("");
  els.flagFileFilter.value = files.includes(currentFile) ? currentFile : "";

  const file = els.flagFileFilter.value;
  const prio = els.flagPrioFilter.value;
  const query = els.flagSearch.value.trim().toLowerCase();
  const filtered = [];
  list.forEach((flag, index) => {
    if (file && String(flag.file || "") !== file) return;
    if (prio && normalizeFlagPriority(flag).prio !== prio) return;
    if (query && !flagSearchText(flag).includes(query)) return;
    filtered.push({ flag, index });
  });

  // 处理进度始终按整个任务统计，不受筛选影响
  const processed = list.reduce(
    (acc, _flag, index) => acc + (state.flagStates[String(index)] ? 1 : 0), 0);
  els.flagProgress.textContent = `已处理 ${processed} / ${list.length}`;

  const groups = groupFlags(filtered);
  const sum = groups.reduce((acc, g) => acc + g.entries.length, 0);
  if (sum !== filtered.length) {
    console.warn(`[flags] 三组条数之和 ${sum} ≠ 当前显示 ${filtered.length}`);
  }
  const high = groups[0].entries.length;
  const mid = groups[1].entries.length;
  const low = groups[2].entries.length;
  // 文案保持原口径：未筛选时展示全量高/中/低；筛选时同时给出当前显示条数
  els.flagPaneSummary.textContent = (file || prio || query)
    ? `共 ${list.length} 处 · 当前筛选 ${filtered.length} 处 · 高 ${high} / 中 ${mid} / 低 ${low}`
    : `共 ${list.length} 处待复核 · 高 ${high} / 中 ${mid} / 低 ${low}`;

  const prevExpanded = new Set(
    [...els.flagBody.querySelectorAll('.fgroup-btn[aria-expanded="true"]')]
      .map((btn) => btn.closest("tr")?.dataset.prio)
      .filter(Boolean)
  );
  els.flagBody.innerHTML = "";
  // 默认展开第一个有条目的组；刷新时保留用户已展开的分组
  const defaultOpen = (groups.find((g) => g.entries.length > 0) || groups[0]).priority;

  for (const g of groups) {
    const expanded = g.entries.length > 0 && (
      prevExpanded.size ? prevExpanded.has(g.priority) : g.priority === defaultOpen);
    const head = document.createElement("tr");
    head.className = "fgroup-head";
    head.dataset.prio = g.priority;
    head.innerHTML =
      `<td colspan="8"><button class="fgroup-btn" type="button" aria-expanded="${expanded}"` +
      `${g.entries.length ? "" : " disabled"}>` +
      `<span class="fgroup-caret" aria-hidden="true">▸</span>` +
      `<span class="fgroup-name">${g.label}</span>` +
      `<span class="fgroup-n">${g.entries.length}</span></button></td>`;
    els.flagBody.appendChild(head);

    const btn = head.querySelector(".fgroup-btn");
    // 展开状态只有这一个入口：aria-expanded 与 tr.hidden 永远同写，避免两处漂移
    const setGroupExpanded = (expanded) => {
      btn.setAttribute("aria-expanded", expanded ? "true" : "false");
      els.flagBody
        .querySelectorAll(`tr.fgroup-item[data-prio="${g.priority}"]`)
        .forEach((tr) => { tr.hidden = !expanded; });
    };
    btn.addEventListener("click", () => {
      setGroupExpanded(btn.getAttribute("aria-expanded") !== "true");
    });
    // 空格键：阻止原生 keyup 触发的 click，统一走一条路径，避免双触发
    btn.addEventListener("keydown", (e) => {
      if (e.key !== " " && e.key !== "Spacebar") return;
      e.preventDefault();
      btn.click();
    });

    for (const { flag, index, unrated } of g.entries) {
      const tr = document.createElement("tr");
      const status = state.flagStates[String(index)] || "";
      tr.className = "fgroup-item" + (status ? ` is-${status}` : "");
      tr.dataset.prio = g.priority;
      tr.hidden = !expanded;
      const kind = String(flag.kind || "").trim() || "—";
      tr.innerHTML =
        `<td><span class="fl-kind">${esc(kind)}</span>` +
        `${unrated ? '<span class="fl-unrated">未分级</span>' : ""}</td>` +
        `<td title="${esc(flag.file)}">${esc(flag.file)}</td>` +
        `<td>第${(flag.row ?? 0) + 1}行 第${(flag.col ?? 0) + 1}列</td>` +
        `<td>${esc(flag.rebuilt)}</td><td>${esc(flag.ocr ?? flag.whole)}</td>` +
        `<td>${esc(flag.zoom ?? flag.local)}</td><td class="flag-reason">${esc(flag.reason)}</td>` +
        `<td class="fl-actions">` +
        `<button class="flag-act" type="button" data-act="locate" title="查看原图并定位">定位</button>` +
        `<button class="flag-act ok${status === "checked" ? " on" : ""}" type="button" ` +
        `data-act="checked" title="标记已核对">✓</button>` +
        `<button class="flag-act bad${status === "retake" ? " on" : ""}" type="button" ` +
        `data-act="retake" title="标记需重拍">↻</button>` +
        `</td>`;
      tr.querySelector('[data-act="locate"]').addEventListener("click", (e) => {
        e.stopPropagation();
        openFlagImage(index);
      });
      tr.querySelector('[data-act="checked"]').addEventListener("click", (e) => {
        e.stopPropagation();
        toggleFlagState(index, "checked");
      });
      tr.querySelector('[data-act="retake"]').addEventListener("click", (e) => {
        e.stopPropagation();
        toggleFlagState(index, "retake");
      });
      els.flagBody.appendChild(tr);
    }
  }
}

let flagStateTimer = null;

function toggleFlagState(index, value) {
  const key = String(index);
  if (state.flagStates[key] === value) delete state.flagStates[key];
  else state.flagStates[key] = value;
  renderFlagPane(state.flags);
  clearTimeout(flagStateTimer);
  flagStateTimer = setTimeout(() => {
    if (!state.jobId) return;
    fetch("/api/job/flags", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ job_id: state.jobId, states: state.flagStates }),
    }).catch(() => { /* 本机写盘失败不打断核对 */ });
  }, 250);
}

/* ---------------- 待复核预览 ---------------- */
function showFlagReview(items) {
  state.flags = items || [];
  renderTabs();
  renderFlagPane(state.flags);
  els.flagSummary.textContent =
    `共 ${items.length} 处存疑值。Excel 末尾「待复核清单」已按优先级分层：` +
    `高（红）/中（橙）请对照原图核对；低优先多为置信度噪音或单号列粘连` +
    `（单号已由格式规则兜底提取），集中在「待复核-低优先」工作表，仅供参考。`;
  els.flagBodyModal.innerHTML = "";
  for (const fl of items) {
    const tr = document.createElement("tr");
    tr.innerHTML =
      `<td title="${esc(fl.file)}">${esc(fl.file)}</td>` +
      `<td>第${(fl.row ?? 0) + 1}行 第${(fl.col ?? 0) + 1}列</td>` +
      `<td>${esc(fl.rebuilt)}</td><td>${esc(fl.ocr)}</td>` +
      `<td>${esc(fl.zoom)}</td><td class="flag-reason">${esc(fl.reason)}</td>`;
    els.flagBodyModal.appendChild(tr);
  }
  els.flagModal.hidden = false;
}
els.flagClose.addEventListener("click", () => { els.flagModal.hidden = true; });
els.flagModal.addEventListener("click", (e) => {
  if (e.target === els.flagModal) els.flagModal.hidden = true;
});

/* ---------------- 单号查询（P1 应用内检索） ---------------- */
let searchTimer = null;

const SEARCH_STATUS = {
  "已找到": "pill-ok",
  "待复核": "pill-pending",
  "疑似错读": "pill-pending",
  "已确认": "pill-ok",
  "待核对": "pill-pending",
  "格式异常": "pill-fail",
  "非YT": "pill-fail",
};

function openSearch(query = "") {
  els.searchModal.hidden = false;
  setSearchMode("single");
  els.searchInput.value = query;
  els.searchResults.innerHTML = "";
  els.searchMeta.textContent = query ? "正在显示本次查找结果…"
    : "输入完整或部分单号，即时在当前结果里查找";
  if (query) setTimeout(runSearch, 0);
  setTimeout(() => els.searchInput.focus(), 30);
}

function closeSearch() { els.searchModal.hidden = true; }

function setSearchMode(mode) {
  const batch = mode === "batch";
  els.searchModeSingle.classList.toggle("on", !batch);
  els.searchModeBatch.classList.toggle("on", batch);
  els.searchModeSingle.setAttribute("aria-selected", String(!batch));
  els.searchModeBatch.setAttribute("aria-selected", String(batch));
  els.searchSingleRow.hidden = batch;
  els.searchBatchBox.hidden = !batch;
  els.searchResults.innerHTML = "";
  if (batch) {
    els.searchMeta.textContent = "每行一个单号，也可用逗号分隔；结果可导出 CSV";
    setTimeout(() => els.searchBatchInput.focus(), 30);
  } else {
    els.searchMeta.textContent = "输入完整或部分单号，即时在当前结果里查找";
    setTimeout(() => els.searchInput.focus(), 30);
  }
}

let batchResult = null;

async function runBatchSearch() {
  const text = els.searchBatchInput.value.trim();
  if (!text) {
    els.searchMeta.textContent = "请先粘贴要查询的单号";
    els.searchBatchInput.focus();
    return;
  }
  els.searchMeta.textContent = "查询中…";
  els.searchBatchExport.disabled = true;
  batchResult = null;
  try {
    const r = await fetch("/api/search/batch", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ job_id: state.jobId || "", text }),
    });
    renderBatchSearch(await r.json());
  } catch (err) {
    els.searchMeta.textContent = "查询失败：" + (err.message || err);
  }
}

const BATCH_STATUS_TEXT = {
  exact: "命中", partial: "模糊命中", none: "未命中", invalid: "格式异常",
};

function renderBatchSearch(d) {
  els.searchResults.innerHTML = "";
  batchResult = d && d.ok ? d : null;
  const hasGroups = !!(d && d.ok && Array.isArray(d.groups) && d.groups.length);
  els.searchBatchExport.disabled = !hasGroups;
  if (!d || !d.ok) {
    els.searchMeta.textContent = (d && d.error) || "暂无可查询的索引";
    return;
  }
  els.searchMeta.textContent =
    `共 ${d.query_count || 0} 个单号：命中 ${d.exact || 0} · 模糊命中 ${d.partial || 0} · ` +
    `未命中 ${d.none || 0} · 格式异常 ${d.invalid || 0}` +
    `（扫描 ${d.scanned || 0} 条${d.issues ? ` · 索引异常 ${d.issues} 项` : ""}）`;
  for (const g of d.groups || []) {
    const box = document.createElement("div");
    box.className = "batch-group";
    const head = document.createElement("div");
    head.className = "batch-head";
    head.innerHTML =
      `<span class="batch-query">${esc(g.query)}</span>` +
      `<span class="batch-badge ${esc(g.status)}">${BATCH_STATUS_TEXT[g.status] || esc(g.status)}</span>` +
      (g.total ? `<span class="batch-note">${g.total} 处</span>` : "");
    box.appendChild(head);
    if (g.status === "invalid") {
      const note = document.createElement("div");
      note.className = "batch-note";
      note.textContent = g.reason || "格式异常";
      box.appendChild(note);
    }
    if (g.status === "none") {
      // 未命中不等于不存在：索引异常/待核对值都可能漏查，口径与单个查询一致
      const note = document.createElement("div");
      note.className = "batch-note";
      note.textContent = "未命中不代表不存在" +
        (d.issues ? `：本次索引异常 ${d.issues} 项，待核对值也可能漏查` : "：可对照原图或在 Excel 中再确认");
      box.appendChild(note);
    }
    if (Array.isArray(g.hits) && g.hits.length) {
      const hits = document.createElement("div");
      hits.className = "batch-hits";
      for (const h of g.hits) {
        const row = document.createElement("div");
        row.className = "batch-hit";
        row.innerHTML =
          `<b>${esc(h.waybill)}</b>` +
          `<span class="pill ${SEARCH_STATUS[h.status] || "pill-pending"}">${esc(h.status || "—")}</span>` +
          `<span>${esc(h.source || "")}${h.page ? ` · 第 ${h.page} 页` : ""}` +
          `${h.sheet ? ` · ${esc(h.sheet)}` : ""}${h.cell ? ` · ${esc(h.cell)}` : ""}</span>`;
        hits.appendChild(row);
      }
      box.appendChild(hits);
      if (g.total > g.hits.length) {
        const more = document.createElement("div");
        more.className = "batch-note";
        more.textContent = `仅显示前 ${g.hits.length} 处，共 ${g.total} 处`;
        box.appendChild(more);
      }
    }
    els.searchResults.appendChild(box);
  }
}

function exportBatchCsv() {
  if (!batchResult || !Array.isArray(batchResult.groups)) return;
  const rows = [["查询单号", "结果", "匹配单号", "确认状态", "来源文件", "页码", "工作表", "单元格"]];
  for (const g of batchResult.groups) {
    const label = BATCH_STATUS_TEXT[g.status] || g.status || "";
    if (!Array.isArray(g.hits) || !g.hits.length) {
      rows.push([g.query, label, "", "", "", "", "", ""]);
      continue;
    }
    for (const h of g.hits) {
      rows.push([g.query, label, h.waybill || "", h.status || "",
                 h.source || "", h.page || "", h.sheet || "", h.cell || ""]);
    }
  }
  // 防 CSV 公式注入：以 = + - @ 或制表/回车开头的单元格加单引号前缀，
  // 避免单号/文件名被 Excel 当成公式执行（仅加引号并不能阻止）。
  const csvCell = (cell) => {
    let text = String(cell ?? "");
    if (/^[=+\-@\t\r]/.test(text)) text = "'" + text;
    return `"${text.replace(/"/g, '""')}"`;
  };
  const csv = "\ufeff" + rows.map((row) => row
    .map(csvCell)
    .join(",")).join("\r\n");
  const now = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  const name = `单号批量查询_${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}`
    + `_${pad(now.getHours())}${pad(now.getMinutes())}.csv`;
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = name;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
  toast(`已导出 ${name}`);
}

async function runSearch() {
  const q = els.searchInput.value.trim();
  if (!q) {
    els.searchMeta.textContent = "输入完整或部分单号，即时在当前结果里查找";
    els.searchResults.innerHTML = "";
    return;
  }
  els.searchMeta.textContent = "查询中…";
  try {
    const r = await fetch(
      `/api/search?q=${encodeURIComponent(q)}&job_id=${encodeURIComponent(state.jobId || "")}`
    );
    renderSearch(await r.json(), q);
  } catch (err) {
    els.searchMeta.textContent = "查询失败：" + (err.message || err);
  }
}

function renderSearch(d, q) {
  els.searchResults.innerHTML = "";
  if (!d.ok) {
    els.searchMeta.textContent = d.error || "暂无可查询的索引";
    return;
  }
  const total = d.total || 0;
  const issues = d.issues || 0;
  els.searchMeta.textContent = d.mode === "find"
    ? (total ? `显示 ${total} 处匹配 / 共扫描 ${d.scanned} 张图片或页`
      : `没有匹配「${q}」的结果 / 共扫描 ${d.scanned} 张图片或页`)
    : total
      ? `命中 ${total} 条 / 共扫描 ${d.scanned} 条` + (issues ? ` · 索引异常 ${issues} 项` : "")
      : `没有匹配「${q}」的单号 · 共扫描 ${d.scanned} 条` + (issues ? ` · 索引异常 ${issues} 项` : "");

  if (!total) {
    // 无结果不等于"不存在"：索引异常/待核对都可能漏查，必须如实说明
    const empty = document.createElement("div");
    empty.className = "search-empty";
    const lines = d.mode === "find"
      ? ["快速识别没有找到匹配。这不代表单号一定不存在：",
         "· 图片模糊、旋转或单号太小时可能会漏读；",
         "· 可以切换到「图片转 Excel」用完整识别进一步核对。"]
      : ["没有找到包含「" + q + "」的单号。这不代表它一定不存在："];
    if (d.mode !== "find") {
      if (issues) lines.push(`· 本次有 ${issues} 项索引异常，该页可能没识别到单号列；`);
      lines.push("· 单号可能被识别成待核对/格式异常值，可在结果条里查看待复核清单；");
      lines.push("· 也可以在导出的数据工作表里用筛选再确认一次。");
    }
    empty.innerHTML = lines.map((line) => esc(line)).join("<br>");
    els.searchResults.appendChild(empty);
    return;
  }

  for (const m of d.matches || []) {
    const item = document.createElement("div");
    item.className = "search-item";

    const top = document.createElement("div");
    top.className = "search-item-top";

    const waybill = document.createElement("span");
    waybill.className = "search-waybill";
    waybill.textContent = m.waybill;
    top.appendChild(waybill);

    const pill = document.createElement("span");
    pill.className = "pill " + (SEARCH_STATUS[m.status] || "pill-pending");
    pill.textContent = m.status || "待核对";
    top.appendChild(pill);

    if (m.duplicate) {
      const dup = document.createElement("span");
      dup.className = "search-flag";
      dup.textContent = "重复单号";
      top.appendChild(dup);
    }

    const copy = document.createElement("button");
    copy.className = "search-copy";
    const addr = m.sheet && m.cell ? `${m.sheet}!${m.cell}` : String(m.source || "");
    copy.textContent = m.sheet && m.cell ? "复制位置" : "复制文件名";
    copy.onclick = async () => {
      try {
        await navigator.clipboard.writeText(addr);
        toast(`已复制：${addr}`);
      } catch { toast(addr); }
    };
    top.appendChild(copy);

    if (m.image_url) {
      const locate = document.createElement("button");
      locate.className = "search-copy search-locate";
      locate.textContent = "原图定位";
      locate.title = "在原图上高亮这个单号的位置";
      locate.onclick = () => openWaybillImage(m);
      top.appendChild(locate);
    }
    item.appendChild(top);

    const loc = document.createElement("div");
    loc.className = "search-loc";
    const page = m.page ? `（第 ${m.page} 页）` : "";
    loc.textContent = d.mode === "find"
      ? `${m.source}${page}`
      : `${m.source}${page} · ${m.sheet}!${m.cell} · 第 ${m.row} 行`;
    item.appendChild(loc);

    if (m.summary) {
      const sum = document.createElement("div");
      sum.className = "search-sum";
      sum.textContent = "行摘要：" + m.summary;
      item.appendChild(sum);
    }
    if (m.note) {
      const note = document.createElement("div");
      note.className = "search-note";
      note.textContent = m.note;
      item.appendChild(note);
    }
    els.searchResults.appendChild(item);
  }
}

els.searchInput.addEventListener("input", () => {
  clearTimeout(searchTimer);
  searchTimer = setTimeout(runSearch, 250);
});
els.searchInput.addEventListener("keydown", (e) => {
  if (e.key === "Enter") { e.preventDefault(); clearTimeout(searchTimer); runSearch(); }
});
els.searchClear.addEventListener("click", () => {
  els.searchInput.value = "";
  els.searchInput.focus();
  els.searchResults.innerHTML = "";
  els.searchMeta.textContent = "输入完整或部分单号，即时在当前结果里查找";
});
els.searchClose.addEventListener("click", closeSearch);
els.searchModal.addEventListener("click", (e) => {
  if (e.target === els.searchModal) closeSearch();
});

// 单个 / 批量查询切换
els.searchModeSingle.addEventListener("click", () => setSearchMode("single"));
els.searchModeBatch.addEventListener("click", () => setSearchMode("batch"));
els.searchBatchRun.addEventListener("click", runBatchSearch);
els.searchBatchExport.addEventListener("click", exportBatchCsv);
els.searchBatchClear.addEventListener("click", () => {
  els.searchBatchInput.value = "";
  els.searchResults.innerHTML = "";
  batchResult = null;
  els.searchBatchExport.disabled = true;
  els.searchMeta.textContent = "每行一个单号，也可用逗号分隔；结果可导出 CSV";
  els.searchBatchInput.focus();
});
els.searchBatchInput.addEventListener("keydown", (e) => {
  // Ctrl+Enter 直接查询，避免长清单里还要去点按钮
  if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) {
    e.preventDefault();
    runBatchSearch();
  }
});

function friendlyError(text) {
  const t = text || "";
  if (t.includes("模型不存在")) {
    return "OCR 模型文件缺失。请确认安装目录下的 models\\PP-OCRv6 完整（det / rec 子目录里应有 .xml 和 .bin 文件）；若确实缺失，请重新获取完整安装包后重装。\n\n技术信息：" + t.slice(0, 300);
  }
  if (t.includes("DLL load failed") || t.includes("找不到指定的模块") || t.includes("The specified module could not be found")) {
    return "OCR 引擎依赖的组件加载失败（不是模型缺失）。常见原因与处理：\n· 杀毒软件误删/隔离了 runtime 里的文件：到杀软的「隔离区/恢复区」还原，并把安装目录加入白名单\n· 下载的文件被系统锁定：右键安装目录里的「解除锁定.bat」→ 以管理员身份运行\n· 仍失败时请重新获取完整安装包重装（新版已内置所需运行库）\n\n技术信息：" + t.slice(0, 300);
  }
  if (t.includes("OCR 引擎未就绪") || t.includes("OCR 引擎启动失败") || t.includes("无法执行 python") || t.includes("缺少 OCR 工作进程脚本")) {
    return "OCR 引擎未能启动（模型文件本身可能完好）。请按顺序检查：\n· 安装目录下的 runtime 文件夹是否完整（可能被杀毒软件清理）\n· 右键「解除锁定.bat」→ 以管理员身份运行，解除系统对下载文件的锁定\n· 过旧的电脑 CPU 或精简版系统可能不受支持\n· 以上无效时请重新获取完整安装包重装\n\n技术信息：" + t.slice(0, 300);
  }
  if (t.includes("所有文件都未识别出内容")) {
    return "所有文件都没能识别出内容。可以试试：\n· 把纸张拍正摆平，避免倾斜、阴影和反光\n· 换更清晰、光线更足的照片\n· 扫描件请使用 300 DPI 以上\n· 确认运单号区域没有被裁掉\n\n详细原因见运行日志。";
  }
  if (t.includes("未找到 OCR 运行环境")) {
    return "未找到本地 OCR 运行环境。请确认安装目录下的 runtime 文件夹完整，或重新获取安装包。\n\n技术信息：" + t.slice(0, 300);
  }
  return t;
}

function fail(message) {
  state.pendingSearchQuery = "";
  setStatus("转换失败", "err");
  setProgress(0, 0, false);
  log("错误：" + message);
  resetRunning();
  const friendly = friendlyError(message);
  showResult("error", `转换失败：${friendly.split("\n")[0].slice(0, 80)}`);
  if (friendly !== message) {
    setTimeout(() => alert(friendly), 120);
  }
}

function resetRunning() {
  state.running = false;
  stopElapsed();              // 保留最后一次耗时显示，只停表
  refreshStartButton();
  if (state.ws) { state.ws.close(); state.ws = null; }
}

/* 清场：把上一轮的结果卡、质量面板、失败条、待复核三组折叠与页签计数全部归零。
   start() 与 showResult() 共用，避免「转换中/失败/取消」时还挂着上一轮残留。 */
function resetRunOutput() {
  els.resultCard.hidden = true;
  els.resultCard.className = "result-card";
  els.rcTitle.textContent = "";
  els.rcSubtitle.textContent = "";
  els.rcSubtitle.removeAttribute("title");

  // 文件行 / 统计 / 备注
  els.rcFile.hidden = true;
  els.rcFileName.textContent = "";
  els.rcFileName.removeAttribute("title");
  els.rcMetrics.hidden = true;
  els.rcNote.hidden = true;
  els.rcNote.textContent = "";

  // 拍摄质量面板（含每图明细的展开态）
  els.rcQuality.hidden = true;
  els.rcQTier.className = "rcq-tier";
  els.rcQTier.textContent = "—";
  els.rcQScore.className = "rcq-score";
  els.rcQScore.textContent = "—";
  els.rcQWhy.hidden = true;
  els.rcQWhy.innerHTML = "";
  els.rcQList.hidden = true;
  els.rcQList.innerHTML = "";
  els.rcQToggle.hidden = true;
  els.rcQToggle.setAttribute("aria-expanded", "false");
  els.rcQToggle.textContent = "每图明细";
  els.rcQToggle.onclick = null;
  els.rcQSug.hidden = true;
  els.rcQSug.innerHTML = "";

  // 按文件汇总
  els.rcFiles.hidden = true;
  els.rcFilesBody.innerHTML = "";
  els.rcFilesCount.textContent = "";

  // 失败条
  els.rcFails.hidden = true;
  els.rcFailList.innerHTML = "";
  els.rcFailCount.textContent = "";

  // 操作按钮：全部隐藏并解除旧回调，避免点到上一轮的 onclick
  els.rcActions.hidden = true;
  for (const btn of [els.btnOpen, els.btnReveal, els.btnFlags,
    els.btnRetryFailed, els.btnSearch, els.btnFullCheck, els.btnDownload]) {
    btn.hidden = true;
    btn.onclick = null;
  }

  // 待复核：页签计数 + 三组折叠表一起归零
  state.flags = [];
  state.flagStates = {};
  els.flagPrioFilter.value = "";
  els.flagSearch.value = "";
  renderFlagPane([]);
  els.flagSummary.textContent = "";
  els.flagBodyModal.innerHTML = "";
  renderTabs();

  state.lastOutPath = "";
}

/* 失败 / 取消：同一张卡片只保留标题和原因，不显示统计与操作 */
function showResult(kind, text) {
  resetRunOutput();
  els.resultCard.hidden = false;
  els.resultCard.className = `result-card ${kind}`;
  els.rcTitle.textContent = kind === "error" ? "转换失败" : "已取消";
  els.rcSubtitle.textContent = text;
  els.rcSubtitle.title = text;
  scrollResultIntoView();
}
function hideResult() { els.resultCard.hidden = true; }
els.resultClose.addEventListener("click", hideResult);

els.btnStart.addEventListener("click", () => {
  if (state.files.length === 0) { els.fileInput.click(); return; }
  start();
});
els.waybillInput.addEventListener("input", refreshStartButton);
els.waybillInput.addEventListener("keydown", (e) => {
  if (e.key === "Enter" && !els.btnStart.disabled) {
    e.preventDefault();
    start();
  }
});
els.btnCancel.addEventListener("click", async () => {
  if (!state.running || !state.jobId) return;
  try {
    await fetch("/api/cancel", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ job_id: state.jobId }),
    });
    setStatus("正在取消…", "warn");
    log("已请求取消，正在等待当前步骤结束…");
  } catch { /* 忽略 */ }
});

/* ---------------- 查看图片（点击文件） ---------------- */
let previewUrl = null;          // 拖入文件的原始 objectURL
let previewRef = null;          // 当前预览的文件条目（有 path 或 file）
let previewRotate = 0;          // 当前预览方向，顺时针 0/90/180/270
let previewSourceUrl = "";      // 原图地址：objectURL 或本机预览接口
let previewRenderUrl = null;    // 旋转后渲染出的临时 objectURL
let previewRenderToken = 0;     // 连续点击旋转时丢弃过期的异步渲染
let previewZoom = 1;            // 当前缩放，1 = 适应窗口
let previewPanX = 0;            // 拖动位移：相对适应窗口时居中位置的横向偏移
let previewPanY = 0;            // 拖动位移：纵向偏移
let previewPan = null;          // 拖动中的起点与起始位移
let previewCopyTarget = null;   // 当前可复制的原图目标（原图定位时才有）
let previewAutoActualPending = false; // 打开预览时是否自动切到 1:1（避免大图缩到窗口后发虚）
const PREVIEW_MAX_EDGE = 4096;  // 预览渲染长边上限，旋转后仍保留足够细节供 1:1 查看
const PREVIEW_ZOOM_MIN = 1;     // 最小就是适应窗口
const PREVIEW_ZOOM_MAX = 8;
const PREVIEW_ZOOM_STEP = 1.25;

// 「复制图片」：普通文件预览复制当前原图；原图定位优先复制带标记的图
function syncPreviewCopyImageButton() {
  const hasTarget = !!previewCopyTarget;
  const src = previewSourceUrl || els.previewImg.getAttribute("src") || "";
  const hasImage = !!src;
  const ready = hasTarget || !!els.previewImg.naturalWidth;
  els.btnPreviewCopyImage.hidden = !(hasTarget || hasImage);
  els.btnPreviewCopyImage.disabled = !ready;
  els.btnPreviewCopyImage.textContent = "复制图片";
}

// 大图默认按 1:1 显示：文字/单号在缩略图尺寸下会发虚，和定位视图保持一致
function zoomPreviewToActualIfNeeded() {
  const img = els.previewImg;
  if (!img.naturalWidth || previewZoom !== 1) return;
  if (previewPanX || previewPanY) return;
  const fitW = img.clientWidth;
  if (!fitW) return;
  const actual = previewActualZoom();
  if (actual <= 1.2) return;
  setPreviewZoom(actual);
}

function updatePreviewAngle() {
  els.previewAngle.textContent = `方向 ${previewRotate}°`;
  els.btnRotateReset.disabled = previewRotate === 0;
}

// 预览用的旋转在浏览器本地完成：只影响画面与后续识别参数，不改动原文件
async function renderRotatedPreview() {
  const token = ++previewRenderToken;
  if (!previewSourceUrl) return;
  if (!previewRotate) {
    if (previewRenderUrl) { URL.revokeObjectURL(previewRenderUrl); previewRenderUrl = null; }
    els.previewImg.src = previewSourceUrl;
    return;
  }
  try {
    const img = new Image();
    img.src = previewSourceUrl;
    await img.decode();
    if (token !== previewRenderToken) return;
    const w = img.naturalWidth;
    const h = img.naturalHeight;
    if (!w || !h) throw new Error("图片尺寸无效");
    const scale = Math.min(1, PREVIEW_MAX_EDGE / Math.max(w, h));
    const dw = Math.max(1, Math.round(w * scale));
    const dh = Math.max(1, Math.round(h * scale));
    const swap = previewRotate === 90 || previewRotate === 270;
    const canvas = document.createElement("canvas");
    canvas.width = swap ? dh : dw;
    canvas.height = swap ? dw : dh;
    const ctx = canvas.getContext("2d");
    ctx.fillStyle = "#FFFFFF";
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    ctx.translate(canvas.width / 2, canvas.height / 2);
    ctx.rotate((previewRotate * Math.PI) / 180);
    ctx.drawImage(img, -dw / 2, -dh / 2, dw, dh);
    const blob = await new Promise((resolve) => canvas.toBlob(resolve, "image/jpeg", 0.95));
    if (token !== previewRenderToken) return;
    if (!blob) throw new Error("渲染失败");
    if (previewRenderUrl) URL.revokeObjectURL(previewRenderUrl);
    previewRenderUrl = URL.createObjectURL(blob);
    els.previewImg.src = previewRenderUrl;
  } catch {
    if (token === previewRenderToken) els.previewImg.src = previewSourceUrl;
  }
}

function applyPreviewRotate(next) {
  if (!previewRef) return;
  if (state.running) { toast("转换进行中，暂不能调整图片方向"); return; }
  previewRotate = ((next % 360) + 360) % 360;
  previewRef.rotate = previewRotate;
  // 方向变了，旧结果不再对应当前图片，退回「待转换」提醒重新跑一次
  if (previewRef.tag === "ok") previewRef.tag = "pending";
  updatePreviewAngle();
  resetPreviewZoom();
  previewAutoActualPending = true;
  renderRotatedPreview();
  renderFiles();
}

function clearPreviewMarks() {
  if (els.previewMarks) els.previewMarks.innerHTML = "";
}

function drawPreviewMarks(m) {
  clearPreviewMarks();
  if (!m || !m.image_url || !Array.isArray(m.box)) return;
  const w = Number(m.image_width) || 0;
  const h = Number(m.image_height) || 0;
  if (!w || !h) return;
  const [x1, y1, x2, y2] = m.box.map(Number);
  const mark = document.createElement("div");
  mark.className = "preview-mark" + (m.markClass ? ` ${m.markClass}` : "");
  mark.style.left = `${(x1 / w) * 100}%`;
  mark.style.top = `${(y1 / h) * 100}%`;
  mark.style.width = `${Math.max(0.4, ((x2 - x1) / w) * 100)}%`;
  mark.style.height = `${Math.max(0.4, ((y2 - y1) / h) * 100)}%`;
  els.previewMarks.appendChild(mark);
}

// 1:1 原图：让 1 个图像像素对应 1 个屏幕像素，看细节最清晰
function previewActualZoom() {
  const img = els.previewImg;
  const baseW = img.clientWidth || img.naturalWidth;
  if (!img.naturalWidth || !baseW) return 1;
  return Math.min(PREVIEW_ZOOM_MAX, Math.max(1, img.naturalWidth / baseW));
}

function togglePreviewActualZoom(anchor) {
  if (!els.previewImg.naturalWidth) return;
  const actual = previewActualZoom();
  if (actual <= 1.02) { resetPreviewZoom(); return; }  // 图比窗口还小，本身已是原图清晰度
  if (Math.abs(previewZoom - actual) < 0.02) resetPreviewZoom();
  else setPreviewZoom(actual, anchor);
}

function updatePreviewZoomUI() {
  const ready = !!(els.previewImg.naturalWidth);
  const actual = previewActualZoom();
  const atActual = ready && actual > 1.02 && Math.abs(previewZoom - actual) < 0.02;
  els.btnZoomReset.textContent = `${Math.round(previewZoom * 100)}%`;
  els.btnZoomReset.disabled = !ready;
  els.btnZoomOut.disabled = !ready || previewZoom <= PREVIEW_ZOOM_MIN + 1e-3;
  els.btnZoomIn.disabled = !ready || previewZoom >= PREVIEW_ZOOM_MAX - 1e-3;
  els.btnZoomReset.title = previewZoom === 1 ? "当前为适应窗口" : "恢复为适应窗口";
  els.btnZoomActual.disabled = !ready || actual <= 1.02;
  els.btnZoomActual.classList.toggle("on", atActual);
  els.btnZoomActual.title = !ready
    ? "图片加载中…"
    : actual <= 1.02
      ? "当前图片已按原图尺寸显示"
      : atActual
        ? "恢复为适应窗口"
        : `按原图像素显示（${Math.round(actual * 100)}%），最清晰`;
  // 任何时候都能按住拖动：放大看细节，适应窗口时也能把图挪到顺手的位置
  els.previewStage.classList.toggle("can-pan", ready);
}

// 拖动边界：无论怎么拖，都至少留 80px 的画面在预览区里，避免把图拖丢
function clampPreviewPan(panX, panY) {
  const img = els.previewImg;
  const stage = els.previewStage;
  const baseW = img.clientWidth || img.naturalWidth;
  const baseH = img.clientHeight || img.naturalHeight;
  if (!baseW || !baseH) return { x: 0, y: 0 };
  const dispW = baseW * previewZoom;
  const dispH = baseH * previewZoom;
  const keepX = Math.min(80, dispW);
  const keepY = Math.min(80, dispH);
  const limitX = Math.max(0, (stage.clientWidth + dispW) / 2 - keepX);
  const limitY = Math.max(0, (stage.clientHeight + dispH) / 2 - keepY);
  return {
    x: Math.min(limitX, Math.max(-limitX, panX)),
    y: Math.min(limitY, Math.max(-limitY, panY)),
  };
}

// 只改画布变换，不重建 DOM；位移先夹取到可视边界内
function applyPreviewTransform() {
  const pan = clampPreviewPan(previewPanX, previewPanY);
  previewPanX = pan.x;
  previewPanY = pan.y;
  const canvas = els.previewCanvas;
  if (previewZoom === 1 && previewPanX === 0 && previewPanY === 0) {
    canvas.style.transform = "";
  } else {
    canvas.style.transform =
      `translate(${previewPanX.toFixed(1)}px, ${previewPanY.toFixed(1)}px) scale(${previewZoom})`;
  }
  // 定位框叠在放大后的画布上：线宽按 1/缩放 反向补偿，保证任何倍率下都不压住文字
  if (els.previewMarks) {
    els.previewMarks.style.setProperty("--mark-scale", (1 / previewZoom).toFixed(3));
  }
}

function resetPreviewZoom() {
  previewZoom = 1;
  previewPanX = 0;
  previewPanY = 0;
  previewPan = null;
  els.previewCanvas.style.transform = "";
  if (els.previewMarks) els.previewMarks.style.setProperty("--mark-scale", "1");
  els.previewStage.classList.remove("panning");
  updatePreviewZoomUI();
}

// 设置缩放：anchor 是鼠标所在的窗口坐标，缩放时该点下方的画面保持不动
function setPreviewZoom(next, anchor) {
  const img = els.previewImg;
  const stage = els.previewStage;
  if (!img.naturalWidth) return;
  const target = Math.min(PREVIEW_ZOOM_MAX, Math.max(PREVIEW_ZOOM_MIN, next));
  if (target <= PREVIEW_ZOOM_MIN + 1e-3) {
    resetPreviewZoom();
    return;
  }
  const rect = stage.getBoundingClientRect();
  const ax = anchor ? anchor.x : rect.left + rect.width / 2;
  const ay = anchor ? anchor.y : rect.top + rect.height / 2;
  // 以预览区中心为原点：位移 p、缩放 z 时，画面点落在 p + u * z
  const dx = ax - (rect.left + rect.width / 2);
  const dy = ay - (rect.top + rect.height / 2);
  const ratio = target / previewZoom;
  previewPanX = dx - (dx - previewPanX) * ratio;
  previewPanY = dy - (dy - previewPanY) * ratio;
  previewZoom = target;
  applyPreviewTransform();
  updatePreviewZoomUI();
}

// 待复核红框定位：把目标单元格放大并移到预览区中心
function focusPreviewBox(m) {
  const [x1, y1, x2, y2] = (m.box || []).map(Number);
  if (![x1, y1, x2, y2].every(Number.isFinite)) return;
  const img = els.previewImg;
  const stage = els.previewStage;
  const w = Number(m.image_width) || img.naturalWidth;
  const h = Number(m.image_height) || img.naturalHeight;
  const baseW = img.clientWidth;
  const baseH = img.clientHeight;
  if (!w || !h || !baseW || !baseH) return;
  const boxW = Math.max(1, x2 - x1);
  const stageW = stage.clientWidth || 640;
  const scale = Math.min(6, Math.max(1.8, (stageW * 0.32) / (baseW * boxW / w)));
  previewZoom = scale;
  previewPanX = -(((x1 + x2) / 2) / w - 0.5) * baseW * scale;
  previewPanY = -(((y1 + y2) / 2) / h - 0.5) * baseH * scale;
  applyPreviewTransform();
  updatePreviewZoomUI();
}

function showLocatedImage(m) {
  if (!m || !m.image_url) { toast("这条没有可定位的原图"); return; }
  previewRef = null;
  previewSourceUrl = "";
  previewRenderToken += 1;
  previewRotate = 0;
  els.previewTools.hidden = true;
  if (previewRenderUrl) { URL.revokeObjectURL(previewRenderUrl); previewRenderUrl = null; }
  if (previewUrl) { URL.revokeObjectURL(previewUrl); previewUrl = null; }
  clearPreviewMarks();
  resetPreviewZoom();
  previewAutoActualPending = !m.zoomToBox;
  els.previewTitle.textContent = m.title || "原图定位";
  els.previewMeta.textContent = m.meta || "";
  els.btnPreviewCopy.disabled = true;
  els.btnPreviewDelete.disabled = true;
  els.btnPreviewCopy.hidden = true;
  els.btnPreviewDelete.hidden = true;
  // 定位模式下「复制图片」可用；没有 copy 目标（如旧接口数据）则保持隐藏
  previewCopyTarget = m.copy || null;
  syncPreviewCopyImageButton();
  els.previewHint.hidden = false;
  els.previewHint.textContent = m.hint || "";
  els.previewModal.hidden = false;
  const render = () => {
    drawPreviewMarks(m);
    if (m.zoomToBox) focusPreviewBox(m);
  };
  els.previewImg.onload = () => { els.previewImg.onload = null; render(); };
  els.previewImg.onerror = () => {
    els.previewImg.onerror = null;
    toast("原图缩略图已过期，请重新转换后再试");
  };
  els.previewImg.src = m.image_url;
  syncPreviewCopyImageButton();
  if (els.previewImg.complete && els.previewImg.naturalWidth) {
    if (m.zoomToBox) render();
    else requestAnimationFrame(() => {
      previewAutoActualPending = false;
      zoomPreviewToActualIfNeeded();
    });
  }
}

function openFlagImage(index) {
  const flag = state.flags[index];
  if (!flag) return;
  const pageIndex = Number(flag.page_index);
  if (!Number.isFinite(pageIndex) || !Array.isArray(flag.box) || !state.jobId) {
    toast("这条待复核没有可定位的原图（可能来自旧任务）");
    return;
  }
  const page = Number(flag.page_no) > 1 ? ` · 第 ${flag.page_no} 页` : "";
  showLocatedImage({
    image_url: `/api/job-image?job_id=${encodeURIComponent(state.jobId)}&idx=${pageIndex}`,
    title: `原图定位 · ${flag.file || ""}`,
    meta: `${flag.file || ""}${page} · 红框为待复核单元格`,
    hint: "红框是待复核单元格位置，请对照原图核对内容；识别值仅作参考。",
    box: flag.box,
    image_width: flag.image_width,
    image_height: flag.image_height,
    markClass: "flag",
    zoomToBox: true,
    copy: {
      job_id: state.jobId,
      idx: pageIndex,
      box: flag.box,
      image_width: flag.image_width,
      image_height: flag.image_height,
    },
  });
}

function openWaybillImage(m) {
  // 从「单号查询」里点定位时先收起查询弹窗，避免它盖在预览上面
  closeSearch();
  const page = m && m.page ? ` · 第 ${m.page} 页` : "";
  // 命中项来自 /api/search：用 image_url 反解出任务号与页码，供「复制图片」使用
  let copy = null;
  try {
    const url = new URL(m.image_url, location.origin);
    const jobId = url.searchParams.get("job_id") || "";
    const idx = Number(url.searchParams.get("idx"));
    if (jobId && Number.isFinite(idx) && Array.isArray(m.box)) {
      copy = {
        job_id: jobId,
        idx,
        box: m.box,
        image_width: m.image_width,
        image_height: m.image_height,
      };
    }
  } catch { copy = null; }
  showLocatedImage({
    image_url: m && m.image_url,
    title: `原图定位 · ${(m && m.waybill) || ""}`,
    meta: `${(m && m.source) || ""}${page} · 粉框为识别到的单号位置`,
    hint: "粉框是识别到的单号位置，已自动放大；滚轮可缩小查看整页。",
    box: m && m.box,
    image_width: m && m.image_width,
    image_height: m && m.image_height,
    zoomToBox: true,
    copy,
  });
}

function openPreview(f) {
  if (/\.pdf$/i.test(f.name || f.path || "")) {
    // PDF 不做图片预览：磁盘文件交给系统默认程序；拖入的文件没有磁盘路径，
    // 也无法在窗口里翻页，直接说明原因，避免弹出一个加载失败的空白预览框。
    if (f.path) {
      revealPath(f.path, "open").then((ok) =>
        toast(ok ? "已用默认程序打开 PDF" : "打开 PDF 失败，请从保存目录查看"));
    } else {
      toast("拖入的 PDF 无法在窗口内预览，可直接开始转换；如需打开原件，请用「添加文件」重新选择");
    }
    return;
  }
  previewRef = f;
  previewRotate = Number(f.rotate) || 0;
  clearPreviewMarks();
  resetPreviewZoom();
  previewAutoActualPending = true;
  els.previewTitle.textContent = f.name;
  els.previewMeta.textContent =
    `${humanSize(f.size)} · ${f.path ? "磁盘文件" : "拖入文件（无磁盘路径）"}`;
  els.previewHint.textContent =
    "这个文件是拖入的，浏览器没有给出磁盘路径，无法复制/删除；用「添加文件」重新选择后可用。";
  els.previewTools.hidden = false;
  updatePreviewAngle();
  els.previewImg.onload = null;
  els.previewImg.onerror = null;
  previewRenderToken += 1;
  if (previewRenderUrl) { URL.revokeObjectURL(previewRenderUrl); previewRenderUrl = null; }
  if (previewUrl) { URL.revokeObjectURL(previewUrl); previewUrl = null; }
  if (f.file) {
    previewUrl = URL.createObjectURL(f.file);
    previewSourceUrl = previewUrl;
  } else if (f.path) {
    previewSourceUrl = `/api/preview?path=${encodeURIComponent(f.path)}`;
  } else {
    previewSourceUrl = "";
  }
  if (!previewSourceUrl) {
    els.previewImg.removeAttribute("src");
  } else {
    renderRotatedPreview();
    if (els.previewImg.complete && els.previewImg.naturalWidth) {
      requestAnimationFrame(() => {
        previewAutoActualPending = false;
        zoomPreviewToActualIfNeeded();
      });
    }
  }
  const canManage = !!f.path;
  els.btnPreviewCopy.disabled = !canManage;
  els.btnPreviewDelete.disabled = !canManage;
  els.btnPreviewCopy.hidden = false;
  els.btnPreviewDelete.hidden = false;
  // 普通文件预览没有“带标记原图”目标，但可以直接复制当前原图
  previewCopyTarget = null;
  syncPreviewCopyImageButton();
  els.previewHint.hidden = canManage;
  els.previewModal.hidden = false;
}

function closePreview() {
  els.previewModal.hidden = true;
  previewRenderToken += 1;
  clearPreviewMarks();
  resetPreviewZoom();
  if (previewRenderUrl) { URL.revokeObjectURL(previewRenderUrl); previewRenderUrl = null; }
  if (previewUrl) { URL.revokeObjectURL(previewUrl); previewUrl = null; }
  previewSourceUrl = "";
  previewRotate = 0;
  previewRef = null;
  previewAutoActualPending = false;
  // 关闭后清空复制目标，避免下次打开残留
  previewCopyTarget = null;
  syncPreviewCopyImageButton();
}

async function fileAction(action) {
  if (!previewRef || !previewRef.path) return null;
  if (state.running) { toast("转换进行中，暂不能修改源文件"); return null; }
  try {
    const r = await fetch("/api/file-op", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ action, path: previewRef.path }),
    });
    return await r.json();
  } catch (err) {
    return { ok: false, error: err.message || String(err) };
  }
}

async function copyPreviewImageViaCanvas() {
  const img = els.previewImg;
  if (!img.naturalWidth) throw new Error("图片还没加载完");
  if (!navigator.clipboard || typeof window.ClipboardItem === "undefined") {
    throw new Error("当前环境不支持复制图片");
  }
  // previewImg 已经是当前方向（旋转后的图会重新渲染成 objectURL），
  // 这里按显示内容原样复制，不能再叠加一次旋转。
  const w = img.naturalWidth;
  const h = img.naturalHeight;
  const canvas = document.createElement("canvas");
  canvas.width = w;
  canvas.height = h;
  const ctx = canvas.getContext("2d");
  if (!ctx) throw new Error("浏览器不支持图片渲染");
  ctx.fillStyle = "#FFFFFF";
  ctx.fillRect(0, 0, w, h);
  ctx.drawImage(img, 0, 0, w, h);
  const blob = await new Promise((resolve) => canvas.toBlob(resolve, "image/png"));
  if (!blob) throw new Error("生成图片失败");
  await navigator.clipboard.write([new ClipboardItem({ "image/png": blob })]);
}

els.btnPreviewCopyImage.addEventListener("click", async () => {
  const btn = els.btnPreviewCopyImage;
  btn.disabled = true;
  btn.textContent = "复制中…";
  try {
    if (previewCopyTarget) {
      try {
        const res = await fetch("/api/copy-image", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(previewCopyTarget),
        });
        const data = await res.json().catch(() => ({}));
        if (!data.ok) throw new Error(data.error || "服务未返回原因");
      } catch (_) {
        // 带标记原图复制失败时，退回复制当前预览图
        await copyPreviewImageViaCanvas();
      }
    } else {
      await copyPreviewImageViaCanvas();
    }
    toast("图片已复制，可直接粘贴到聊天或文档");
  } catch (err) {
    toast(`复制图片失败：${err.message || "当前环境不支持"}`);
  } finally {
    // 还原按钮文字与可点状态（目标可能已被关闭弹窗清空）
    syncPreviewCopyImageButton();
  }
});

els.btnPreviewCopy.addEventListener("click", async () => {
  const d = await fileAction("copy");
  if (!d) return;
  toast(d.ok ? "文件已复制到剪贴板，可直接粘贴到文件夹" : `复制失败：${d.error}`);
});

els.btnPreviewDelete.addEventListener("click", async () => {
  if (!previewRef) return;
  const name = previewRef.name;
  if (!confirm(`删除文件？\n${name}\n\n文件会移到系统回收站，可从回收站还原。`)) return;
  const d = await fileAction("delete");
  if (!d) return;
  if (!d.ok) { toast(`删除失败：${d.error}`); return; }
  const targetPath = previewRef.path;
  state.files = state.files.filter((f) => f.path !== targetPath && f.id !== previewRef.id);
  state.selected = new Set([...state.selected].filter((id) => !state.files.some((f) => f.id === id)));
  renderFiles();
  closePreview();
  toast(`已删除：${name}（回收站可还原）`);
});

els.btnPreviewClose.addEventListener("click", closePreview);
els.btnRotateLeft.addEventListener("click", () => applyPreviewRotate(previewRotate - 90));
els.btnRotateRight.addEventListener("click", () => applyPreviewRotate(previewRotate + 90));
els.btnRotateReset.addEventListener("click", () => applyPreviewRotate(0));

/* 预览缩放：按钮、滚轮、双击；按住画面可自由拖到任意位置 */
els.btnZoomIn.addEventListener("click", () => setPreviewZoom(previewZoom * PREVIEW_ZOOM_STEP));
els.btnZoomOut.addEventListener("click", () => setPreviewZoom(previewZoom / PREVIEW_ZOOM_STEP));
els.btnZoomReset.addEventListener("click", () => resetPreviewZoom());
els.btnZoomActual.addEventListener("click", () => togglePreviewActualZoom());
els.previewImg.addEventListener("load", () => {
  updatePreviewZoomUI();
  syncPreviewCopyImageButton();
  if (previewAutoActualPending) {
    previewAutoActualPending = false;
    requestAnimationFrame(zoomPreviewToActualIfNeeded);
  }
});
els.previewStage.addEventListener("wheel", (e) => {
  if (!els.previewImg.naturalWidth) return;
  e.preventDefault();
  const factor = e.deltaY < 0 ? 1.12 : 1 / 1.12;
  setPreviewZoom(previewZoom * factor, { x: e.clientX, y: e.clientY });
}, { passive: false });
els.previewStage.addEventListener("dblclick", (e) => {
  if (!els.previewImg.naturalWidth || e.target.closest("button")) return;
  togglePreviewActualZoom({ x: e.clientX, y: e.clientY });
});

els.previewStage.addEventListener("pointerdown", (e) => {
  // 不再限制“放大后才能拖”：适应窗口时也可以把图挪到任意位置
  if (e.button !== 0 || !els.previewImg.naturalWidth) return;
  e.preventDefault();
  previewPan = {
    x: e.clientX, y: e.clientY,
    panX: previewPanX, panY: previewPanY,
  };
  els.previewStage.classList.add("panning");
  try { els.previewStage.setPointerCapture(e.pointerId); } catch (_) { /* 忽略 */ }
});
els.previewStage.addEventListener("pointermove", (e) => {
  if (!previewPan) return;
  previewPanX = previewPan.panX + (e.clientX - previewPan.x);
  previewPanY = previewPan.panY + (e.clientY - previewPan.y);
  applyPreviewTransform();
});
["pointerup", "pointercancel"].forEach((type) => {
  els.previewStage.addEventListener(type, () => {
    previewPan = null;
    els.previewStage.classList.remove("panning");
  });
});

els.previewModal.addEventListener("click", (e) => {
  if (e.target === els.previewModal) closePreview();
});
els.previewImg.addEventListener("error", () => {
  updatePreviewZoomUI();
  if (!els.previewModal.hidden) toast("图片预览加载失败（文件可能已被移动）");
});

/* ---------------- 弹窗：首次运行 + 关于 ---------------- */
async function initConfig() {
  try {
    const [cfgR, privR] = await Promise.all([
      fetch("/api/config"),
      fetch("/api/privacy"),
    ]);
    const cfg = await cfgR.json();
    const priv = await privR.json();

    els.firstRunNotice.textContent = priv.notice;
    els.firstRunVersion.textContent = `${priv.title} 版本 ${priv.version}`;
    els.aboutVersion.textContent = `版本 ${priv.version} · 本地 OCR 图片 / PDF 转 Excel 工具`;
    els.aboutNotice.textContent = priv.notice;
    const preferred = MODES.some((m) => m.value === cfg.default_mode)
      ? cfg.default_mode : cfg.last_mode;
    state.mode = MODES.some((m) => m.value === preferred) ? preferred : "full";
    const preferredLayout = cfg.default_layout || cfg.last_layout;
    state.layout = preferredLayout === "merged" ? "merged" : "per_image";
    state.autoOpen = !!cfg.auto_open;
    if (els.setDefaultMode) {
      els.setDefaultMode.value = MODES.some((m) => m.value === cfg.default_mode)
        ? cfg.default_mode : state.mode;
    }
    if (els.setAutoOpen) els.setAutoOpen.checked = state.autoOpen;
    if (els.setJobKeep) els.setJobKeep.value = cfg.job_keep || 10;
    if (els.setHoverActivate) els.setHoverActivate.checked = cfg.hover_activate !== false;
    if (els.setPasteHotkey) els.setPasteHotkey.checked = cfg.paste_hotkey !== false;
    state.checkUpdates = cfg.check_updates !== false;
    if (els.setCheckUpdates) els.setCheckUpdates.checked = state.checkUpdates;
    els.outPath.value = cfg.last_output || "";
    if (!els.outPath.value) {
      const d = await (await fetch("/api/default-output")).json();
      els.outPath.value = d.path;
    }
    renderModes();
    if (!cfg.privacy_ack) showModal(els.firstRunModal);
    else maybeCheckUpdates();
  } catch (e) {
    renderModes();
  }
}

function showModal(mask) { mask.hidden = false; }
function hideModal(mask) { mask.hidden = true; }

els.firstRunConfirm.addEventListener("click", async () => {
  if (els.firstRunAck.checked) {
    try {
      await fetch("/api/config", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ privacy_ack: true }),
      });
    } catch { /* 忽略 */ }
  }
  hideModal(els.firstRunModal);
  maybeCheckUpdates();
});

/* ---------------- 版本更新检查（只查询版本号，不上传数据） ---------------- */
const UPDATE_LAST_KEY = "fatduck.update.lastCheck";
const UPDATE_DISMISS_KEY = "fatduck.update.dismissed";
let updateTimer = 0;

function _openExternal(url) {
  if (!/^https:\/\//i.test(String(url || ""))) return;
  window.open(url, "_blank", "noopener");
}

function showUpdateBanner(d) {
  if (localStorage.getItem(UPDATE_DISMISS_KEY) === d.latest) return;
  state.lastUpdate = d;
  els.updateVer.textContent = `v${d.latest}`;
  els.updateBanner.hidden = false;
  els.btnAbout.classList.add("has-update");
}

async function checkForUpdates(manual = false) {
  if (manual && els.btnCheckUpdate) {
    els.btnCheckUpdate.disabled = true;
    els.updateStatus.textContent = "正在检查更新…";
  }
  try {
    const r = await fetch(`/api/update/check${manual ? "?force=1" : ""}`);
    const d = await r.json();
    if (!d.ok) {
      if (manual) els.updateStatus.textContent = "检查失败，请稍后再试";
      return;
    }
    localStorage.setItem(UPDATE_LAST_KEY, String(Date.now()));
    if (d.has_update) {
      showUpdateBanner(d);
      if (manual) {
        els.updateStatus.textContent = `发现新版本 v${d.latest}`;
        if (/^https:\/\//i.test(d.url || "")) {
          const a = document.createElement("a");
          a.href = d.url;
          a.target = "_blank";
          a.rel = "noopener";
          a.textContent = " 去下载";
          els.updateStatus.appendChild(a);
        }
      }
    } else if (manual) {
      els.updateStatus.textContent = `已是最新版本（v${d.current}）`;
    }
  } catch {
    if (manual) els.updateStatus.textContent = "检查失败，请检查网络";
  } finally {
    if (manual && els.btnCheckUpdate) els.btnCheckUpdate.disabled = false;
  }
}

function maybeCheckUpdates() {
  if (!state.checkUpdates) return;
  const last = Number(localStorage.getItem(UPDATE_LAST_KEY) || 0);
  if (Date.now() - last < 24 * 3600 * 1000) return;
  clearTimeout(updateTimer);
  updateTimer = setTimeout(() => checkForUpdates(false), 2500);
}

els.btnCheckUpdate?.addEventListener("click", () => checkForUpdates(true));
els.updateClose?.addEventListener("click", () => {
  els.updateBanner.hidden = true;
  if (state.lastUpdate?.latest) {
    localStorage.setItem(UPDATE_DISMISS_KEY, state.lastUpdate.latest);
  }
});
els.updateLink?.addEventListener("click", (e) => {
  e.preventDefault();
  _openExternal(state.lastUpdate?.url);
});

els.btnAbout.addEventListener("click", async () => {
  try {
    const r = await fetch("/api/privacy");
    const d = await r.json();
    els.aboutVersion.textContent = `版本 ${d.version} · 本地 OCR 图片 / PDF 转 Excel 工具`;
    els.aboutNotice.textContent = d.notice;
    if (els.updateStatus) {
      if (state.lastUpdate?.has_update) {
        els.updateStatus.textContent = `发现新版本 v${state.lastUpdate.latest}`;
        if (/^https:\/\//i.test(state.lastUpdate.url || "")) {
          const a = document.createElement("a");
          a.href = state.lastUpdate.url;
          a.target = "_blank";
          a.rel = "noopener";
          a.textContent = " 去下载";
          els.updateStatus.appendChild(a);
        }
      } else {
        els.updateStatus.textContent = `当前版本 v${d.version}`;
      }
    }
  } catch { /* 忽略 */ }
  showModal(els.aboutModal);
});
els.aboutConfirm.addEventListener("click", () => hideModal(els.aboutModal));

/* ---------------- 云端文本纠错（可选，默认关闭） ---------------- */
const KEY_SOURCE_TEXT = {
  environment: "已从环境变量 IMG2DOC_NVIDIA_KEY 读取到密钥",
  config: "已从本机配置读取到密钥（明文存储，注意风险）",
  missing: "未检测到密钥：请设置环境变量 IMG2DOC_NVIDIA_KEY",
};

async function refreshCloudStatus() {
  try {
    const d = await (await fetch("/api/config")).json();
    els.cloudToggle.checked = !!d.nv_fix_enabled;
    const src = d.nv_fix_key_source || "missing";
    const cls = src === "missing" ? "warn" : "ok";
    els.cloudStatus.innerHTML =
      `当前状态：<b class="${cls}">${d.nv_fix_enabled ? "已开启" : "已关闭"}</b><br>` +
      `模型：<b>${d.nv_fix_model || "-"}</b><br>` +
      `密钥：<span class="${cls}">${KEY_SOURCE_TEXT[src] || src}</span>`;
  } catch { /* 忽略 */ }
}

els.btnCloud.addEventListener("click", async () => {
  await refreshCloudStatus();
  showModal(els.cloudModal);
});

els.cloudConfirm.addEventListener("click", async () => {
  const enabled = els.cloudToggle.checked;
  try {
    await fetch("/api/config", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ nv_fix_enabled: enabled }),
    });
    log(`云端文本纠错已${enabled ? "开启" : "关闭"}`);
  } catch { /* 忽略 */ }
  hideModal(els.cloudModal);
});

els.cloudModal.addEventListener("click", (e) => {
  if (e.target === els.cloudModal) hideModal(els.cloudModal);
});

/* ---------------- 设置（默认工作方式 / 自动打开 / 历史保留） ---------------- */
els.btnSettings.addEventListener("click", () => {
  els.setOutPath.textContent = els.outPath.value || "—";
  showModal(els.settingsModal);
});

els.settingsConfirm.addEventListener("click", async () => {
  const mode = els.setDefaultMode.value;
  const auto = els.setAutoOpen.checked;
  let keep = Number(els.setJobKeep.value);
  if (!Number.isFinite(keep)) keep = 10;
  keep = Math.max(1, Math.min(50, Math.round(keep)));
  els.setJobKeep.value = keep;
  const hoverActivate = els.setHoverActivate ? els.setHoverActivate.checked : true;
  const pasteHotkey = els.setPasteHotkey ? els.setPasteHotkey.checked : true;
  const checkUpdates = els.setCheckUpdates ? els.setCheckUpdates.checked : true;
  try {
    await fetch("/api/config", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        default_mode: mode, auto_open: auto, job_keep: keep,
        hover_activate: hoverActivate, paste_hotkey: pasteHotkey,
        check_updates: checkUpdates,
      }),
    });
    state.autoOpen = auto;
    state.checkUpdates = checkUpdates;
    if (!checkUpdates) {
      clearTimeout(updateTimer);
      els.updateBanner.hidden = true;
    } else {
      maybeCheckUpdates();
    }
    toast("设置已保存");
  } catch {
    toast("设置保存失败，请重试");
  }
  hideModal(els.settingsModal);
});

els.btnSettingsCloud.addEventListener("click", async () => {
  hideModal(els.settingsModal);
  await refreshCloudStatus();
  showModal(els.cloudModal);
});

els.btnClearJobs.addEventListener("click", async () => {
  if (!confirm("确定清除本机全部任务历史吗？\n\n只删除历史记录与预览图，不影响已经导出的 Excel。")) return;
  try {
    await fetch("/api/jobs/clear", { method: "POST" });
    toast("任务历史已清除");
    if (!els.jobsModal.hidden) renderJobs([]);
  } catch {
    toast("清除失败，请重试");
  }
});

els.settingsModal.addEventListener("click", (e) => {
  if (e.target === els.settingsModal) hideModal(els.settingsModal);
});

[els.firstRunModal, els.aboutModal, els.cloudModal, els.settingsModal, els.jobsModal].forEach((mask) =>
  mask.addEventListener("click", (e) => { if (e.target === mask && mask !== els.firstRunModal) hideModal(mask); })
);

/* ---------------- 最近任务（M4 本机历史） ---------------- */
const JOB_MODE_NAME = { full: "全量识别", waybill: "单号清单", find: "查找单号" };

async function openJobs() {
  showModal(els.jobsModal);
  els.jobsList.innerHTML = '<div class="jobs-empty">正在载入…</div>';
  try {
    const d = await (await fetch("/api/jobs")).json();
    renderJobs(Array.isArray(d.jobs) ? d.jobs : []);
  } catch {
    els.jobsList.innerHTML = '<div class="jobs-empty">任务历史载入失败，请重试。</div>';
  }
}

function renderJobs(jobs) {
  els.jobsMeta.textContent = jobs.length
    ? `共 ${jobs.length} 个任务 · 只保存在本机，保留条数可在设置里调整`
    : "只保存在本机，保留条数可在设置里调整";
  els.jobsList.innerHTML = "";
  if (!jobs.length) {
    els.jobsList.innerHTML =
      '<div class="jobs-empty">还没有任务历史。转换完成后会自动保存在这里，隔天也能回来核对。</div>';
    return;
  }
  for (const job of jobs) {
    const row = document.createElement("div");
    row.className = "job-row";
    const files = Array.isArray(job.files) ? job.files : [];
    const fileText = files.length
      ? files[0] + (files.length > 1 ? ` 等 ${files.length} 个` : "")
      : "（未知文件）";
    const stats = [];
    if (job.total_pages) stats.push(`${job.total_pages} 页`);
    stats.push(job.mode === "find" ? `命中 ${job.waybills || 0} 处` : `单号 ${job.waybills || 0}`);
    if (job.mode === "full") stats.push(`待复核 ${job.flags_total || 0}`);
    const quality = job.quality_level
      ? `<span class="job-quality ${QUALITY_CLASS[job.quality_level] || ""}">${esc(job.quality_level)}</span>`
      : "";
    const outName = String(job.out_path || "").split(/[\\/]/).pop();
    row.innerHTML =
      `<div class="job-main">` +
      `<div class="job-line1"><b class="job-time">${esc(job.time || "")}</b>` +
      `<span class="job-mode">${esc(JOB_MODE_NAME[job.mode] || job.mode || "")}</span>${quality}</div>` +
      `<div class="job-line2" title="${esc(files.join("\n"))}">${esc(fileText)} · ${stats.join(" · ")}</div>` +
      (outName
        ? `<div class="job-out" title="${esc(job.out_path)}">${esc(outName)}` +
          `${job.out_exists ? "" : "（文件已移动或删除）"}</div>`
        : "") +
      `</div>` +
      `<div class="job-acts">` +
      `<button class="btn job-load" type="button" data-act="load" title="${job.has_flags ? "载入待复核快照与本机处理状态" : "查看该任务的结果汇总"}">${job.has_flags ? "载入复核" : "查看汇总"}</button>` +
      (job.out_path
        ? `<button class="btn" type="button" data-act="open" title="用默认程序打开结果 Excel"${job.out_exists ? "" : " disabled"}>打开 Excel</button>` +
          `<button class="btn" type="button" data-act="dir" title="在资源管理器中打开所在目录">目录</button>`
        : "") +
      `</div>`;
    row.querySelector('[data-act="load"]').addEventListener("click", () => loadHistoryJob(job.job_id));
    const openBtn = row.querySelector('[data-act="open"]');
    if (openBtn) openBtn.addEventListener("click", async () => {
      if (await revealPath(job.out_path, "open")) toast("已用默认程序打开");
      else toast("打开失败，文件可能已被移动或删除");
    });
    const dirBtn = row.querySelector('[data-act="dir"]');
    if (dirBtn) dirBtn.addEventListener("click", async () => {
      const dir = String(job.out_path || "").replace(/[\\/][^\\/]*$/, "");
      if (await revealPath(dir || job.out_path, "select")) toast("已在资源管理器中打开");
      else toast("打开目录失败，文件可能已被移动或删除");
    });
    els.jobsList.appendChild(row);
  }
}

// 载入历史任务：恢复待复核快照与本机处理状态，复用现有工作台与原图红框。
// 原图走 /api/job-image 的磁盘回退，程序重启后也能定位。
async function loadHistoryJob(jobId) {
  if (state.running) { toast("正在转换中，请完成后再载入历史任务"); return; }
  try {
    const [jobR, statesR] = await Promise.all([
      fetch(`/api/job?job_id=${encodeURIComponent(jobId)}`),
      fetch(`/api/job/flags?job_id=${encodeURIComponent(jobId)}`),
    ]);
    const jd = await jobR.json();
    if (!jd.ok) throw new Error(jd.error || "任务不存在或已清理");
    const sd = await statesR.json().catch(() => ({}));
    const meta = jd.job || {};
    const mode = meta.mode || "full";
    state.jobId = jobId;
    // 与 finish() 同口径：只有全量识别有待复核工作台
    state.flags = mode === "full" && Array.isArray(meta.flags) ? meta.flags : [];
    state.flagStates = (sd && sd.states) || {};
    state.lastOutPath = meta.out_path || "";
    els.flagFileFilter.value = "";
    els.flagPrioFilter.value = "";
    els.flagSearch.value = "";
    renderTabs();
    renderFlagPane(state.flags);
    showDoneCard({
      history: true,
      time: meta.time || "",
      mode,
      out_path: meta.out_path || "",
      sheets: meta.sheets || 0,
      tables: meta.tables || 0,
      waybills: meta.waybills || 0,
      unique_waybills: meta.unique_waybills || 0,
      confirmed_waybills: meta.confirmed_waybills || 0,
      pending_waybills: meta.pending_waybills || 0,
      format_invalid_waybills: meta.format_invalid_waybills || 0,
      flags: state.flags,
      quality: meta.quality_level
        ? { level: meta.quality_level, score: meta.quality_score } : null,
      file_summary: meta.file_summary || [],
      searchable: false,
      download_url: meta.out_path
        ? `/api/download?path=${encodeURIComponent(meta.out_path)}` : "",
      query: meta.query || "",
      matches: meta.matches || 0,
      exact_matches: meta.exact_matches || 0,
      review_matches: meta.review_matches || 0,
      suspect_matches: meta.suspect_matches || 0,
      scanned: meta.scanned || 0,
      failed_items: meta.failed_items || [],
    });
    hideModal(els.jobsModal);
    applyTab(state.flags.length ? "flags" : "files");
    log(`已载入历史任务（${meta.time || jobId}）：待复核 ${state.flags.length} 处，处理状态已恢复。`);
    toast(state.flags.length ? "已载入历史任务，可继续核对" : "已载入历史任务汇总");
  } catch (err) {
    toast(err.message || "载入历史任务失败");
  }
}

els.btnJobs.addEventListener("click", openJobs);
els.jobsClose.addEventListener("click", () => hideModal(els.jobsModal));

/* ---------------- 快捷键 ---------------- */
document.addEventListener("keydown", (e) => {
  const modalOpen = !els.firstRunModal.hidden || !els.aboutModal.hidden
    || !els.cloudModal.hidden || !els.searchModal.hidden || !els.previewModal.hidden
    || !els.jobsModal.hidden || !els.settingsModal.hidden || !els.flagModal.hidden;
  if (e.key === "F1") { e.preventDefault(); els.btnAbout.click(); return; }
  if (e.key === "Escape") {
    if (!els.firstRunModal.hidden && !els.firstRunAck.checked) { hideModal(els.firstRunModal); return; }
    if (!els.aboutModal.hidden) { hideModal(els.aboutModal); return; }
    if (!els.cloudModal.hidden) { hideModal(els.cloudModal); return; }
    if (!els.jobsModal.hidden) { hideModal(els.jobsModal); return; }
    if (!els.settingsModal.hidden) { hideModal(els.settingsModal); return; }
    if (!els.flagModal.hidden) { hideModal(els.flagModal); return; }
    if (!els.searchModal.hidden) { closeSearch(); return; }
    if (!els.previewModal.hidden) { closePreview(); return; }
    els.btnCancel.click(); return;
  }
  if (modalOpen) return;
  if (e.ctrlKey && e.key.toLowerCase() === "o") { e.preventDefault(); els.btnAdd.click(); }
  else if (e.ctrlKey && e.key === "Enter") { e.preventDefault(); start(); }
  else if (e.key === "Delete") {
    const el = document.activeElement;
    const tag = el && el.tagName;
    const type = ((el && el.type) || "").toLowerCase();
    // 输入框里按 Delete 是删文字；勾选框/单选框上按 Delete 仍应删除选中文件
    if ((tag === "INPUT" || tag === "TEXTAREA") && type !== "checkbox" && type !== "radio") return;
    if (!state.running) els.btnRemove.click();
  }
});

/* ---------------- 无边框窗口标题栏 ---------------- */
(function initWindowChrome() {
  const controls = document.getElementById("winControls");
  const topbar = document.getElementById("topbar");
  if (!controls || !topbar) return;
  const btnMin = document.getElementById("winMin");
  const btnMax = document.getElementById("winMax");
  const btnClose = document.getElementById("winClose");

  let maximized = false;
  let activated = false;

  const desktopApi = () => (window.pywebview && window.pywebview.api) || null;
  const call = (name, fallback, args) => {
    const api = desktopApi();
    if (!api || typeof api[name] !== "function") return Promise.resolve(fallback);
    try {
      const ret = args ? api[name](...args) : api[name]();
      return Promise.resolve(ret).catch(() => fallback);
    } catch (_) {
      return Promise.resolve(fallback);
    }
  };

  function applyState(max) {
    maximized = !!max;
    btnMax.classList.toggle("is-max", maximized);
    document.body.classList.toggle("is-max", maximized);
    const label = maximized ? "还原" : "最大化";
    btnMax.setAttribute("aria-label", label);
    btnMax.title = label;
  }
  window.__winState = applyState;  // 供 Python 窗口事件回调

  function toggleMaximize() {
    call("win_toggle_maximize", !maximized).then((next) => {
      applyState(typeof next === "boolean" ? next : !maximized);
    });
  }

  btnMin.addEventListener("click", () => call("win_minimize", null));
  btnMax.addEventListener("click", toggleMaximize);
  btnClose.addEventListener("click", () => call("win_close", null));

  // 别让标题栏的拖动逻辑把窗口按钮上的点击当成拖动
  ["pointerdown", "mousedown", "dblclick"].forEach((type) => {
    controls.addEventListener(type, (e) => e.stopPropagation());
  });

  /* 顶栏拖动：按下时先不动窗口，等鼠标移动超过 4px 才把拖动交给 Windows 原生
     移动循环（ReleaseCapture + WM_NCLBUTTONDOWN）。这样单击不会抢走鼠标，
     双击事件能正常落到页面；拖动本身由系统完成，贴边分屏、跨屏 DPI、跟手性
     都是系统级表现。按下时 preventDefault，避免拖动过程中选中页面文字。 */
  let drag = null;

  topbar.addEventListener("pointerdown", (e) => {
    if (e.button !== 0 || !desktopApi()) return;
    if (e.target.closest(".win-controls, button, a, input, select, textarea")) return;
    drag = { x: e.screenX, y: e.screenY, started: false };
  });

  // 只拦 mousedown：能阻止浏览器进入文字框选（蓝色高亮），click / dblclick 仍会触发
  topbar.addEventListener("mousedown", (e) => {
    if (e.button !== 0 || !desktopApi()) return;
    if (e.target.closest(".win-controls, button, a, input, select, textarea")) return;
    e.preventDefault();
  });

  topbar.addEventListener("pointermove", (e) => {
    if (!drag || drag.started) return;
    if (Math.abs(e.screenX - drag.x) < 4 && Math.abs(e.screenY - drag.y) < 4) return;
    drag.started = true;
    // 最大化时先还原并按鼠标位置摆好，再交给系统拖动，避免系统还原与 WinForms
    // 状态不一致（会把窗口缩成最小尺寸、按钮图标错乱）
    const ready = maximized
      ? call("win_drag_start", null, [e.screenX]).then((geo) => {
          if (geo && geo.restored) applyState(false);
        })
      : Promise.resolve();
    ready.then(() => call("win_drag_native", false)).catch(() => { /* 忽略 */ });
  });

  ["pointerup", "pointercancel", "lostpointercapture"].forEach((type) => {
    topbar.addEventListener(type, () => {
      drag = null;
      // 拖动时系统接管了鼠标，浏览器可能仍处于“按下”状态，收尾时清掉误选的文字
      try { const sel = window.getSelection(); if (sel) sel.removeAllRanges(); } catch (_) { /* 忽略 */ }
    });
  });

  function activate() {
    if (activated) return;
    activated = true;
    controls.hidden = false;
    document.body.classList.add("is-desktop");
    bindGrips();
    call("win_state", false).then(applyState);
  }

  /* 无边框窗口缩放：边缘 4px、四角 10px 热区。拖动时按逻辑像素算出新的
     窗口矩形，交给 Python SetWindowPos 执行；不使用系统 WS_THICKFRAME，
     所以网页顶栏直接贴到窗口最外沿，不会再多出一条系统色带。 */
  const MIN_W = 960, MIN_H = 680;
  let gripDrag = null;
  let gripsBound = false;

  function bindGrips() {
    if (gripsBound) return;
    gripsBound = true;
    document.querySelectorAll(".win-grip").forEach((grip) => {
      grip.addEventListener("pointerdown", (e) => {
        if (e.button !== 0 || !desktopApi() || maximized) return;
        e.preventDefault();
        try { grip.setPointerCapture(e.pointerId); } catch (_) { /* 忽略 */ }
        call("win_bounds", null).then((geo) => {
          if (!geo || geo.maximized) return;
          gripDrag = {
            grip, dir: grip.dataset.dir || "se", geo,
            sx: e.screenX, sy: e.screenY,
            last: { x: geo.x, y: geo.y, width: geo.width, height: geo.height },
            queued: false,
          };
          document.body.classList.add("is-resizing");
        });
      });

      grip.addEventListener("pointermove", (e) => {
        if (!gripDrag || gripDrag.grip !== grip) return;
        const dir = gripDrag.dir;
        const g = gripDrag.geo;
        const dx = e.screenX - gripDrag.sx;
        const dy = e.screenY - gripDrag.sy;
        let x = g.x, y = g.y, w = g.width, h = g.height;
        if (dir.includes("e")) w = Math.max(MIN_W, g.width + dx);
        if (dir.includes("s")) h = Math.max(MIN_H, g.height + dy);
        if (dir.includes("w")) { w = Math.max(MIN_W, g.width - dx); x = g.x + g.width - w; }
        if (dir.includes("n")) { h = Math.max(MIN_H, g.height - dy); y = g.y + g.height - h; }
        gripDrag.last = { x, y, width: w, height: h };
        if (gripDrag.queued) return;
        gripDrag.queued = true;
        requestAnimationFrame(() => {
          if (!gripDrag) return;
          gripDrag.queued = false;
          const b = gripDrag.last;
          call("win_resize", false, [Math.round(b.x), Math.round(b.y),
                                     Math.round(b.width), Math.round(b.height)]);
        });
      });

      const finish = () => {
        if (!gripDrag || gripDrag.grip !== grip) return;
        gripDrag = null;
        document.body.classList.remove("is-resizing");
      };
      grip.addEventListener("pointerup", finish);
      grip.addEventListener("pointercancel", finish);
      grip.addEventListener("lostpointercapture", finish);
    });
  }

  if (desktopApi()) activate();
  window.addEventListener("pywebviewready", activate);

  // 兜底：pywebview 的 API 注入可能晚于本脚本执行，而 pywebviewready 也可能已经
  // 在脚本执行前触发；浏览器模式轮询到超时也不会显示按钮。
  let tries = 0;
  const poll = setInterval(() => {
    if (activated || ++tries > 40) { clearInterval(poll); return; }
    if (desktopApi()) { clearInterval(poll); activate(); }
  }, 250);
})();

/* ---------------- 启动 ---------------- */
log("准备就绪。首次转换需要加载本地识别模型，可能等待几分钟。");
renderFiles();
applyTab(state.tab);
initConfig();

/* 心跳：本地服务据此判断窗口是否还开着，关闭窗口后自动退出程序 */
setInterval(() => { fetch("/api/ping").catch(() => {}); }, 5000);
fetch("/api/ping").catch(() => {});
