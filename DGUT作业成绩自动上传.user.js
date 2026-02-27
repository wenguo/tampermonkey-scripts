// ==UserScript==
// @name         作业成绩自动填写（多Sheet手动选择·最终版+延时+日志+抽样等级）
// @namespace    http://tampermonkey.net/
// @version      7.0
// @description  先从左侧列表抓取【姓名/列表等级/列表成绩】→只更新不一致的学生；多sheet手动选择；轻量MutationObserver捕捉iView message；用“点击前时间窗”等待“开始加载作业文件/加载完成”；成绩输入后延时2秒；设置抽样等级；提交后等待再次“加载完成”再处理下一个；日志CSV导出；防卡死
// @author       YourName
// @match        https://hw.dgut.edu.cn/teacher/homeworkPlan/*/mark
// @updateURL    https://cdn.jsdelivr.net/gh/wenguo/tampermonkey-scripts@main/DGUT%E4%BD%9C%E4%B8%9A%E6%88%90%E7%BB%A9%E8%87%AA%E5%8A%A8%E4%B8%8A%E4%BC%A0.user.js
// @downloadURL  https://cdn.jsdelivr.net/gh/wenguo/tampermonkey-scripts@main/DGUT%E4%BD%9C%E4%B8%9A%E6%88%90%E7%BB%A9%E8%87%AA%E5%8A%A8%E4%B8%8A%E4%BC%A0.user.js
// @grant        none
// @require      https://cdn.jsdelivr.net/npm/xlsx/dist/xlsx.full.min.js
// ==/UserScript==

(function () {
  "use strict";

  /***********************
   * 可调配置
   ***********************/
  const submitButtonSelector = "button.score-operation-button.ivu-btn-success";
  const studentCardSelector = ".ivu-card-body > div"; // 每个学生行
  const clickStudentLinkSelector = "a";

  // iView message：只抓 span（轻量）
  const messageSpanSelector = ".ivu-message-notice-content span";
  const messageRootSelector = "div.ivu-message";

  const EXCEL_COLS = {
    id: ["学号", "学生学号", "ID", "StudentID"],
    name: ["姓名", "学生姓名", "Name"],
    grade: ["自动总分_调整后", "成绩"],
    remark: ["评语", "详细评语"],
  };

  const PAGE_INPUT_HINTS = {
    grade: ["请输入作业成绩", "作业成绩", "成绩", "得分", "分数"],
    remark: ["评语", "详细评语", "批阅", "点评", "意见", "备注"],
  };

  // 站点消息文案（以你日志为准）
  const MSG_START_LOAD = "开始加载作业文件";
  const MSG_LOAD_DONE = "加载完成";
  const MSG_SAVE_UPLOADING = "保存并上传中";
  const MSG_SUCCEEDED = "Succeeded";

  const FORCE_OVERWRITE = false; // true=无论一致与否都重写
  const UPDATE_UNMARKED_ONLY = false; // true=只更新“未批阅”的（列表成绩为“未批阅”）

  const AFTER_CLICK_DELAY = 250;
  const STEP_DELAY_MS = 500;

  // ✅ 成绩输入后额外延时 2 秒
  const GRADE_INPUT_TEST_DELAY_MS = 2000;

  // 等消息超时
  const WAIT_START_LOAD_MAX_MS = 9000;
  const WAIT_LOAD_DONE_MAX_MS = 35000;

  // ✅ 加载完成后静默一小段
  const AFTER_LOAD_SILENCE_MS = 700;

  // ✅ 点击提交后：等待“Succeeded”（可选）+ 最终“加载完成”（必须）
  const WAIT_SUBMIT_SUCCEEDED_MAX_MS = 12000;
  const WAIT_AFTER_SUBMIT_LOAD_DONE_MAX_MS = 35000;

  // ✅ 防止单个学生卡死
  const PER_STUDENT_HARD_TIMEOUT_MS = 65000;

  const DEBUG = true;

  const buttonStyle =
    "width: 100%; height: 36px; font-size: 14px; border: none; padding: 8px 10px; cursor: pointer; box-sizing: border-box; border-radius: 8px;";

  let isPaused = false;
  let studentDataBySheet = {};
  let currentSheetName = "";

  /***********************
   * Dock UI (left panel + right debug)
   ***********************/
  const DOCK_UI_STORE_POS = "AUTO_GRADE_DOCK_POS_V1";
  const LEFT_UI_STORE_POS = "AUTO_GRADE_LEFT_UI_POS_V1";
  const DBG_UI_STORE_POS = "AUTO_GRADE_DBG_POS_V1";

  const dock = document.createElement("div");
  dock.style =
    "position: fixed; z-index: 10001; top: 10px; left: 10px; width: 960px; max-width: calc(100vw - 20px); max-height: calc(100vh - 20px); background: rgba(255,255,255,0.92); border: 1px solid rgba(0,0,0,0.18); border-radius: 12px; box-shadow: 0 10px 30px rgba(0,0,0,0.22); overflow: hidden;";

  const dockHeader = document.createElement("div");
  dockHeader.style =
    "padding: 8px 10px; font-size: 13px; font-weight: 700; color: #111; background: rgba(0,0,0,0.03); border-bottom: 1px solid rgba(0,0,0,0.08); cursor: move; user-select: none;";
  dockHeader.textContent = "AutoGrade";

  const dockBody = document.createElement("div");
  dockBody.style =
    "display: flex; align-items: stretch; gap: 12px; padding: 10px; box-sizing: border-box; max-height: calc(100vh - 70px);";

  const dockLeft = document.createElement("div");
  dockLeft.style = "width: 240px; flex: 0 0 240px; display: flex; flex-direction: column; gap: 8px;";

  const dockRight = document.createElement("div");
  dockRight.style = "flex: 1 1 auto; min-width: 320px; display: flex; flex-direction: column;";

  const dbgTitle = document.createElement("div");
  dbgTitle.textContent = "Debug";
  dbgTitle.style = "font-size: 12px; font-weight: 700; color: #111; margin-bottom: 8px;";

  const dbgBody = document.createElement("pre");
  dbgBody.style =
    "margin: 0; padding: 10px; flex: 1 1 auto; overflow: auto; white-space: pre-wrap; word-break: break-word; font-size: 12px; line-height: 1.45; background: rgba(0,0,0,0.72); color: #fff; border-radius: 10px; box-sizing: border-box;";
  dbgBody.textContent = "Debug: ready\n";

  dockRight.appendChild(dbgTitle);
  dockRight.appendChild(dbgBody);
  dockBody.appendChild(dockLeft);
  dockBody.appendChild(dockRight);
  dock.appendChild(dockHeader);
  dock.appendChild(dockBody);
  document.body.appendChild(dock);

  function loadPosFromStore(key) {
    try {
      const raw = localStorage.getItem(key);
      if (!raw) return null;
      const p = JSON.parse(raw);
      if (!p || typeof p.left !== "number" || typeof p.top !== "number") return null;
      return p;
    } catch (_) {
      return null;
    }
  }

  function saveDockPos(left, top) {
    try {
      localStorage.setItem(DOCK_UI_STORE_POS, JSON.stringify({ left, top }));
    } catch (_) {}
  }

  // Restore position: new dock pos -> old left panel pos -> old debug pos
  const dockSaved = loadPosFromStore(DOCK_UI_STORE_POS) || loadPosFromStore(LEFT_UI_STORE_POS) || loadPosFromStore(DBG_UI_STORE_POS);
  if (dockSaved) {
    dock.style.left = `${dockSaved.left}px`;
    dock.style.top = `${dockSaved.top}px`;
  }

  // Drag support (header only)
  let dockDragging = false;
  let dockDragOffX = 0;
  let dockDragOffY = 0;

  dockHeader.addEventListener("mousedown", (ev) => {
    if (ev.button !== 0) return;
    const rect = dock.getBoundingClientRect();
    dockDragging = true;
    dockDragOffX = ev.clientX - rect.left;
    dockDragOffY = ev.clientY - rect.top;
    ev.preventDefault();
  });

  window.addEventListener("mousemove", (ev) => {
    if (!dockDragging) return;
    const left = Math.max(0, ev.clientX - dockDragOffX);
    const top = Math.max(0, ev.clientY - dockDragOffY);
    dock.style.left = `${left}px`;
    dock.style.top = `${top}px`;
  });

  window.addEventListener("mouseup", () => {
    if (!dockDragging) return;
    dockDragging = false;
    const rect = dock.getBoundingClientRect();
    saveDockPos(rect.left, rect.top);
  });

  const dbgLines = [];
  function nowIso() {
    const d = new Date();
    const pad = (n) => String(n).padStart(2, "0");
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(
      d.getMinutes()
    )}:${pad(d.getSeconds())}`;
  }
  function dbg(msg) {
    if (!DEBUG) return;
    const line = `[${nowIso()}] ${msg}`;
    console.log("[AutoGrade]", msg);
    dbgLines.push(line);
    while (dbgLines.length > 90) dbgLines.shift();
    dbgBody.textContent = dbgLines.join("\n");
  }

  /***********************
   * Status UI
   ***********************/
  const statBox = document.createElement("div");
  statBox.style =
    "width: 100%; padding: 8px; font-size: 12px; background: rgba(0,0,0,0.75); color: #fff; border-radius: 8px; line-height: 1.5; box-sizing: border-box; white-space: pre-wrap;";
  statBox.textContent = "未加载Excel";
  const setStat = (t) => (statBox.textContent = t);

  /***********************
   * 工具
   ***********************/
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  function normText(x) {
    return String(x ?? "")
      .replace(/\u00A0/g, " ")
      .replace(/[\r\n\t]/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  }

  function remarkPreview(s, maxLen = 18) {
    const t = String(s ?? "").replace(/\s+/g, " ").trim();
    return t.length > maxLen ? t.slice(0, maxLen) + "..." : t;
  }

  function toNumStrMaybe(s) {
    const t = normText(s);
    if (!t) return "";
    const m = t.match(/-?\d+(\.\d+)?/);
    return m ? m[0] : "";
  }

  function toIntSafe(s) {
    const n = parseInt(toNumStrMaybe(s), 10);
    return Number.isFinite(n) ? n : NaN;
  }

  /***********************
   * ✅ 轻量消息捕捉：只处理新增 message 节点
   ***********************/
  const messageHistory = []; // {ts, tms, seq, text}
  const MESSAGE_HISTORY_MAX = 180;
  let lastMessageAt = 0;
  let msgSeq = 0;

  function pushMessage(text) {
    const t = normText(text);
    if (!t) return;
    const last = messageHistory.length ? messageHistory[messageHistory.length - 1].text : "";
    if (t === last) return;

    messageHistory.push({ ts: nowIso(), tms: Date.now(), seq: ++msgSeq, text: t });
    while (messageHistory.length > MESSAGE_HISTORY_MAX) messageHistory.shift();
    lastMessageAt = Date.now();

    dbg(`捕捉消息: ${t}`);
  }

  function extractMessageSpansFromNode(node) {
    if (!node || node.nodeType !== 1) return;
    const el = node;

    if (el.matches && el.matches(messageSpanSelector)) {
      pushMessage(el.innerText || el.textContent || "");
      return;
    }
    if (el.querySelectorAll) {
      el.querySelectorAll(messageSpanSelector).forEach((sp) => pushMessage(sp.innerText || sp.textContent || ""));
    }
  }

  function startMessageObserverLight() {
    document.querySelectorAll(messageSpanSelector).forEach((sp) => pushMessage(sp.innerText || sp.textContent || ""));

    const obs = new MutationObserver((mutations) => {
      for (const m of mutations) {
        if (m.type !== "childList") continue;
        for (const n of m.addedNodes) {
          if (n.nodeType !== 1) continue;
          const el = n;

          if (
            (el.matches && el.matches(messageRootSelector)) ||
            (el.querySelector && el.querySelector(messageRootSelector)) ||
            (el.matches && el.matches(messageSpanSelector)) ||
            (el.querySelector && el.querySelector(messageSpanSelector))
          ) {
            extractMessageSpansFromNode(el);
          }
        }
      }
    });

    obs.observe(document.body, { childList: true, subtree: true });
    dbg("MessageObserver 已启动（轻量）");
    return obs;
  }

  async function waitMessageSinceTime(substr, timeoutMs, sinceTimeMs, label) {
    dbg(
      `等待消息「${substr}」 label=${label || ""} timeout=${timeoutMs}ms sinceTime=${new Date(sinceTimeMs).toISOString()}`
    );

    const end = Date.now() + timeoutMs;
    while (Date.now() < end) {
      for (let i = 0; i < messageHistory.length; i++) {
        const m = messageHistory[i];
        if (m.tms >= sinceTimeMs && m.text.includes(substr)) {
          dbg(`命中消息「${substr}」 seq=${m.seq} time=${m.ts} text=${m.text}`);
          return { ok: true, hitSeq: m.seq };
        }
      }
      await sleep(120);
    }

    dbg(`等待「${substr}」超时。最近消息：${messageHistory.slice(-10).map((x) => x.text).join(" | ") || "无"}`);
    return { ok: false, hitSeq: -1 };
  }

  async function waitSilence(ms, timeoutMs) {
    dbg(`等待消息静默 ${ms}ms (timeout=${timeoutMs}ms)`);
    const end = Date.now() + timeoutMs;
    while (Date.now() < end) {
      if (lastMessageAt === 0) return true;
      const delta = Date.now() - lastMessageAt;
      if (delta >= ms) {
        dbg(`已静默 ${delta}ms`);
        return true;
      }
      await sleep(120);
    }
    dbg("等待静默超时");
    return false;
  }

  async function waitPdfLoadFlow(sinceTimeMs) {
    // start 可选，done 必须
    await waitMessageSinceTime(MSG_START_LOAD, WAIT_START_LOAD_MAX_MS, sinceTimeMs, "start");
    const done = await waitMessageSinceTime(MSG_LOAD_DONE, WAIT_LOAD_DONE_MAX_MS, sinceTimeMs, "done");
    if (!done.ok) throw new Error("等待PDF“加载完成”超时（看右上角 Debug 最近消息）");
    await waitSilence(AFTER_LOAD_SILENCE_MS, 6000);
  }

  function hasAnyMessageSince(tsMs) {
    for (let i = 0; i < messageHistory.length; i++) if (messageHistory[i].tms >= tsMs) return true;
    return false;
  }

  function isMarkingUiReady() {
    const gradeInput =
      document.querySelector("input.ivu-input-number-input[placeholder='请输入作业成绩']") ||
      findInputLike(PAGE_INPUT_HINTS.grade, false);
    const submitButton = document.querySelector(submitButtonSelector);

    if (!gradeInput || gradeInput.disabled || gradeInput.offsetParent === null) return false;
    if (!submitButton || submitButton.disabled || submitButton.offsetParent === null) return false;
    return true;
  }

  async function waitAfterSubmitFlow(sinceTimeMs) {
    // 可选消息：保存并上传中 / Succeeded
    await waitMessageSinceTime(MSG_SAVE_UPLOADING, 5000, sinceTimeMs, "uploading");
    await waitMessageSinceTime(MSG_SUCCEEDED, WAIT_SUBMIT_SUCCEEDED_MAX_MS, sinceTimeMs, "succeeded");
    // 必须：最终加载完成（提交后会重新加载PDF/刷新）
    await waitMessageSinceTime(MSG_START_LOAD, 15000, sinceTimeMs, "reload_start");
    const done = await waitMessageSinceTime(MSG_LOAD_DONE, WAIT_AFTER_SUBMIT_LOAD_DONE_MAX_MS, sinceTimeMs, "reload_done");
    if (!done.ok) throw new Error("提交后等待“加载完成”超时（可能保存未生效或页面未刷新）");
    await waitSilence(AFTER_LOAD_SILENCE_MS, 6000);
  }

  /***********************
   * 受控组件写入
   ***********************/
  function setValueByNativeSetter(el, value) {
    const v = value == null ? "" : String(value);
    const proto =
      el.tagName.toLowerCase() === "textarea"
        ? HTMLTextAreaElement.prototype
        : HTMLInputElement.prototype;
    const desc = Object.getOwnPropertyDescriptor(proto, "value");
    if (desc && desc.set) desc.set.call(el, v);
    else el.value = v;
  }

  function fillInputControlled(el, value) {
    el.focus();
    setValueByNativeSetter(el, value);
    el.dispatchEvent(new Event("input", { bubbles: true }));
    el.dispatchEvent(new Event("change", { bubbles: true }));
    el.dispatchEvent(new KeyboardEvent("keydown", { bubbles: true, key: "Enter" }));
    el.dispatchEvent(new KeyboardEvent("keyup", { bubbles: true, key: "Enter" }));
    el.blur();
    el.dispatchEvent(new Event("blur", { bubbles: true }));
  }

  function findInputLike(hints, preferTextarea = false) {
    const els = Array.from(document.querySelectorAll("input, textarea")).filter(
      (el) => !el.disabled && el.offsetParent !== null
    );

    const score = (el) => {
      const attrs = [
        el.getAttribute("placeholder"),
        el.getAttribute("aria-label"),
        el.getAttribute("name"),
        el.getAttribute("id"),
        el.className,
      ]
        .map(normText)
        .join(" | ");

      let s = 0;
      for (const h of hints) if (attrs.includes(h)) s += 6;

      const near = normText(el.parentElement?.innerText || "");
      for (const h of hints) if (near.includes(h)) s += 3;

      if (preferTextarea && el.tagName.toLowerCase() === "textarea") s += 3;
      if (!preferTextarea && el.tagName.toLowerCase() === "input") s += 1;
      return s;
    };

    els.sort((a, b) => score(b) - score(a));
    return els[0] && score(els[0]) > 0 ? els[0] : null;
  }

  /***********************
   * 抽样等级（由成绩推导）
   ***********************/
  function scoreToLevel(scoreNum) {
    if (scoreNum >= 90 && scoreNum <= 100) return "优";
    if (scoreNum >= 80 && scoreNum <= 89) return "良";
    if (scoreNum >= 70 && scoreNum <= 79) return "中";
    return "差";
  }

  function findSamplingSelectExact() {
    const ps = Array.from(document.querySelectorAll("p")).filter((p) => normText(p.innerText) === "抽样等级");
    for (const p of ps) {
      const next = p.nextElementSibling;
      if (next && next.classList && next.classList.contains("ivu-select")) return next;
    }
    return null;
  }

  function getSamplingLevelFromPage() {
    const sel = findSamplingSelectExact();
    if (!sel) return "";
    return normText(sel.querySelector(".ivu-select-selected-value")?.innerText || "");
  }

  async function setSamplingLevel(levelText) {
    const sel = findSamplingSelectExact();
    if (!sel) {
      dbg("未找到抽样等级 select");
      return false;
    }
    const current = getSamplingLevelFromPage();
    dbg(`抽样等级 当前=${current || "(空)"} 目标=${levelText}`);
    if (current === levelText) return true;

    const selection = sel.querySelector(".ivu-select-selection");
    if (!selection) return false;

    selection.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }));
    selection.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    await sleep(150);

    const dropdowns = Array.from(document.querySelectorAll(".ivu-select-dropdown"));
    const visible = dropdowns.find((d) => d.offsetParent !== null) || dropdowns[dropdowns.length - 1];
    if (!visible) return false;

    const items = Array.from(visible.querySelectorAll(".ivu-select-dropdown-list .ivu-select-item"));
    const target = items.find((li) => normText(li.innerText) === levelText);
    if (!target) return false;

    target.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }));
    target.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    target.dispatchEvent(new MouseEvent("mouseup", { bubbles: true }));

    await sleep(150);
    const after = getSamplingLevelFromPage();
    dbg(`抽样等级 设置后=${after || "(空)"}`);
    return after === levelText;
  }

  /***********************
   * Excel 解析（表头扫描）
   ***********************/
  function detectHeaderRow(rows, maxScan = 30) {
    const mustHaveAnyGrade = EXCEL_COLS.grade.map(normText);
    const mustHaveName = EXCEL_COLS.name.map(normText);

    const scanN = Math.min(maxScan, rows.length);
    for (let i = 0; i < scanN; i++) {
      const r = rows[i] || [];
      const normRow = r.map(normText).filter(Boolean);
      if (normRow.length < 2) continue;

      const hasName = normRow.some((x) => mustHaveName.includes(x));
      const hasGrade = normRow.some((x) => mustHaveAnyGrade.includes(x));
      if (hasName && hasGrade) return { headerRowIndex: i, headersRaw: r };
    }
    return null;
  }

  function pickHeaderIndexFuzzy(headersRaw, candidates) {
    const headersNorm = headersRaw.map(normText);
    const candNorm = candidates.map(normText);

    for (const c of candNorm) {
      const idx = headersNorm.findIndex((h) => h === c);
      if (idx !== -1) return idx;
    }
    for (const c of candNorm) {
      const idx = headersNorm.findIndex((h) => h.includes(c));
      if (idx !== -1) return idx;
    }
    return -1;
  }

  /***********************
   * UI：文件/Sheet/扫描差异/开始更新/暂停/导出日志
   ***********************/
  // Controls live in the dock's left column.
  const fileInput = document.createElement("input");
  fileInput.type = "file";
  fileInput.accept = ".xlsx,.xls";
  fileInput.style = `${buttonStyle} background: #fff; color: #111; border: 1px solid rgba(0,0,0,0.25);`;
  fileInput.addEventListener("change", handleFile);
  dockLeft.appendChild(fileInput);

  const sheetSelect = document.createElement("select");
  sheetSelect.style =
    "width: 100%; height: 36px; font-size: 14px; border: 1px solid rgba(0,0,0,0.25); background: #fff; border-radius: 8px; padding: 4px 8px; box-sizing: border-box;";
  sheetSelect.disabled = true;
  sheetSelect.addEventListener("change", () => {
    currentSheetName = sheetSelect.value || "";
    dbg(`选择sheet: ${currentSheetName}`);
  });
  dockLeft.appendChild(sheetSelect);

  const scanButton = document.createElement("button");
  scanButton.textContent = "扫描差异(预览)";
  scanButton.style = `${buttonStyle} background: #6f42c1; color: white; cursor: not-allowed;`;
  scanButton.disabled = true;
  scanButton.addEventListener("click", () => {
    if (!currentSheetName) return alert("请先选择sheet！");
    const plan = buildUpdatePlan(currentSheetName);
    showPlanPreview(plan);
  });
  dockLeft.appendChild(scanButton);

  const startButton = document.createElement("button");
  startButton.textContent = "开始更新(仅差异)";
  startButton.style = `${buttonStyle} background: gray; color: white; cursor: not-allowed;`;
  startButton.disabled = true;
  startButton.addEventListener("click", () => {
    if (!currentSheetName) return alert("请先在下拉框选择一个sheet！");
    isPaused = false;
    runUpdatePlan().catch((e) => console.error("[AutoGrade] 执行异常:", e));
  });
  dockLeft.appendChild(startButton);

  const pauseButton = document.createElement("button");
  pauseButton.textContent = "暂停";
  pauseButton.style = `${buttonStyle} background: #ffc107; color: black;`;
  pauseButton.addEventListener("click", () => {
    isPaused = true;
    progressState.phase = "已暂停";
    updateStatProgress();
    dbg("已暂停");
  });
  dockLeft.appendChild(pauseButton);

  const exportLogButton = document.createElement("button");
  exportLogButton.textContent = "导出日志CSV";
  exportLogButton.style = `${buttonStyle} background: #17a2b8; color: white;`;
  exportLogButton.addEventListener("click", () => exportLogsCsv());
  dockLeft.appendChild(exportLogButton);

  // Status box goes into the left panel (avoid overlap)
  dockLeft.appendChild(statBox);

  /***********************
   * 日志CSV
   ***********************/
  let runLogs = [];
  let runSummary = { sheet: "", total: 0, updated: 0, skipped: 0, notFound: 0, failed: 0, startedAt: "", finishedAt: "" };
  let lastPlan = null;

  let progressState = { sheet: "", idx: 0, total: 0, name: "", phase: "" };

  function updateStatProgress(note) {
    const lines = [];

    if (note) lines.push(String(note).replace(/[\r\n\t]+/g, " ").slice(0, 160));
    if (progressState.sheet) lines.push(`sheet: ${progressState.sheet}`);

    const total = Number(progressState.total) || 0;
    const idx = Number(progressState.idx) || 0;
    if (total > 0) lines.push(`进度: ${idx}/${total}`);
    if (progressState.name) lines.push(`当前: ${progressState.name}`);
    if (progressState.phase) lines.push(`阶段: ${progressState.phase}`);

    if (runSummary && (runSummary.total || runSummary.updated || runSummary.failed || runSummary.skipped || runSummary.notFound)) {
      lines.push(`已更新:${runSummary.updated}`);
      lines.push(`一致跳过:${runSummary.skipped}`);
      lines.push(`未找到:${runSummary.notFound} 失败:${runSummary.failed}`);
    }

    setStat(lines.join("\n"));
  }

  function csvEscape(v) {
    const s = String(v ?? "");
    if (/[",\n]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
    return s;
  }

  function exportLogsCsv() {
    if (!runLogs.length) return alert("当前没有日志可导出（请先执行一次）");
    const headers = ["ts", "sheet", "status", "name", "sid", "list_grade", "list_level", "excel_grade", "excel_level", "remark_preview", "note"];
    const lines = [headers.join(",")];
    for (const r of runLogs) lines.push(headers.map((k) => csvEscape(r[k])).join(","));

    const csv = "\ufeff" + lines.join("\n");
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });

    const a = document.createElement("a");
    const fnSheet = (runSummary.sheet || "sheet").replace(/[\\/:*?"<>|]/g, "_");
    const filename = `作业批阅日志_${fnSheet}_${new Date().toISOString().slice(0, 19).replace(/[:T]/g, "-")}.csv`;
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 1000);
  }

  function showPlanPreview(plan) {
    const lines = [];
    lines.push(`sheet=${plan.sheet}`);
    lines.push(`页面学生=${plan.pageTotal} | Excel记录=${plan.excelTotal}`);
    lines.push(`需要更新=${plan.toUpdate.length} | Excel未找到=${plan.notFound.length} | 跳过一致=${plan.same.length}`);
    lines.push("");
    lines.push("前20条待更新（姓名: 列表成绩/等级 -> 目标成绩/等级）");
    plan.toUpdate.slice(0, 20).forEach((x, idx) => {
      lines.push(`${idx + 1}. ${x.name}: ${x.listScore || "-"} / ${x.listLevel || "-"} -> ${x.excelScore || "-"} / ${x.excelLevel || "-"}`);
    });
    if (plan.toUpdate.length > 20) lines.push(`... 还有 ${plan.toUpdate.length - 20} 条`);
    alert(lines.join("\n"));
  }

  /***********************
   * 读取Excel
   ***********************/
  function handleFile(event) {
    const file = event.target.files[0];
    if (!file) return;

    setStat("正在读取Excel...");
    dbg(`读取Excel: ${file.name}`);

    const reader = new FileReader();
    reader.onload = function (e) {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });

      const sheetNames = workbook.SheetNames || [];
      if (!sheetNames.length) return alert("Excel中未发现任何sheet");

      const tmp = {};
      const skipped = [];

      for (const sn of sheetNames) {
        const ws = workbook.Sheets[sn];
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
        if (!rows || rows.length < 2) {
          skipped.push(`${sn}(空表/行数不足)`);
          continue;
        }

        const hdr = detectHeaderRow(rows, 30);
        if (!hdr) {
          skipped.push(`${sn}(未找到表头行)`);
          continue;
        }

        const headers = hdr.headersRaw;
        const startDataRow = hdr.headerRowIndex + 1;

        const idIdx = pickHeaderIndexFuzzy(headers, EXCEL_COLS.id);
        const nameIdx = pickHeaderIndexFuzzy(headers, EXCEL_COLS.name);
        const gradeIdx = pickHeaderIndexFuzzy(headers, EXCEL_COLS.grade);
        const remarkIdx = pickHeaderIndexFuzzy(headers, EXCEL_COLS.remark);

        if (nameIdx === -1 || gradeIdx === -1) {
          skipped.push(`${sn}(缺少姓名/成绩列)`);
          continue;
        }

        const list = rows
          .slice(startDataRow)
          .map((r) => {
            const g = normText(r[gradeIdx]);
            const gNumStr = toNumStrMaybe(g);
            const gNum = parseInt(gNumStr || "", 10);
            const level = Number.isFinite(gNum) ? scoreToLevel(gNum) : "";
            return {
              学号: idIdx !== -1 ? normText(r[idIdx]) : "",
              姓名: normText(r[nameIdx]),
              成绩: gNumStr || normText(g),
              评语: remarkIdx !== -1 ? normText(r[remarkIdx]) : "",
              等级: level,
            };
          })
          .filter((x) => x.姓名);

        tmp[sn] = list;
      }

      if (!Object.keys(tmp).length) {
        setStat("Excel解析失败");
        dbg(`Excel解析失败 skipped=${skipped.join(" | ")}`);
        return alert("未能解析到可用sheet数据（控制台查看 skipped）");
      }

      studentDataBySheet = tmp;

      sheetSelect.innerHTML = "";
      const opt0 = document.createElement("option");
      opt0.value = "";
      opt0.textContent = "请选择sheet";
      sheetSelect.appendChild(opt0);
      Object.keys(studentDataBySheet).forEach((sn) => {
        const opt = document.createElement("option");
        opt.value = sn;
        opt.textContent = `${sn}（${studentDataBySheet[sn].length}）`;
        sheetSelect.appendChild(opt);
      });
      sheetSelect.disabled = false;
      currentSheetName = "";

      scanButton.disabled = false;
      scanButton.style.cursor = "pointer";
      scanButton.style.background = "#6f42c1";

      startButton.disabled = false;
      startButton.style.background = "#28a745";
      startButton.style.cursor = "pointer";

      setStat(`Excel加载完成\n可用sheet: ${Object.keys(tmp).length}/${sheetNames.length}\n先点“扫描差异”`);
      dbg(`Excel加载完成 可用sheet=${Object.keys(tmp).length}/${sheetNames.length}`);
      if (skipped.length) dbg(`跳过sheet: ${skipped.join(" | ")}`);

      alert("Excel加载成功！请选择对应sheet后，先点“扫描差异(预览)”，确认无误再点“开始更新”。");
    };

    reader.readAsArrayBuffer(file);
  }

  /***********************
   * ✅ 从页面“列表”抓取：姓名/列表等级/列表成绩
   ***********************/
  function getStudentListFromPageWithGradeLevel() {
    const rows = Array.from(document.querySelectorAll(studentCardSelector)).filter((el) => el.offsetParent !== null);

    const list = rows
      .map((div) => {
        const link = div.querySelector(clickStudentLinkSelector);
        if (!link) return null;

        // 名字在 a 的第一个 text node
        const name = normText(link.childNodes?.[0]?.textContent || link.firstChild?.textContent || "");

        // a 内部的 span 是“优/良/中/差”（有时空）
        const levelSpan = link.querySelector("span");
        const listLevel = normText(levelSpan?.innerText || "");

        // 右侧 span float right 是成绩/未批阅
        const scoreSpan = div.querySelector("span[style*='float: right']");
        const listScoreRaw = normText(scoreSpan?.innerText || "");
        const listScoreNum = toNumStrMaybe(listScoreRaw);
        const listScore = listScoreNum || listScoreRaw;

        // 学号（可选：从整行 text 抓 6~12 位数字）
        const text = normText(div.textContent || "");
        const idMatch = text.match(/(\d{6,12})/);
        const sid = idMatch ? idMatch[1] : "";

        return { sid, name, listLevel, listScore };
      })
      .filter((x) => x && x.name);

    dbg(`列表抓取：学生数=${list.length}`);
    return list;
  }

  function findStudentRecord(sheetName, pageStudent) {
    const arr = studentDataBySheet[sheetName] || [];
    if (pageStudent.sid) {
      const hit = arr.find((x) => x.学号 && x.学号 === pageStudent.sid);
      if (hit) return hit;
    }
    return arr.find((x) => x.姓名 === pageStudent.name) || null;
  }

  /***********************
   * ✅ 先构建更新计划：只更新不一致的
   ***********************/
  function buildUpdatePlan(sheetName) {
    const pageList = getStudentListFromPageWithGradeLevel();
    const excelArr = studentDataBySheet[sheetName] || [];

    const plan = {
      sheet: sheetName,
      pageTotal: pageList.length,
      excelTotal: excelArr.length,
      toUpdate: [],
      same: [],
      notFound: [],
    };

    for (const ps of pageList) {
      const rec = findStudentRecord(sheetName, ps);
      if (!rec) {
        plan.notFound.push(ps);
        continue;
      }

      const excelScore = normText(rec.成绩);
      const excelLevel = normText(rec.等级);

      const listScore = normText(ps.listScore);
      const listLevel = normText(ps.listLevel);

      const isUnmarked = listScore.includes("未批阅") || listScore === "";
      if (UPDATE_UNMARKED_ONLY && !isUnmarked) {
        plan.same.push({ ...ps, excelScore, excelLevel, note: "仅更新未批阅，跳过已批阅" });
        continue;
      }

      const sameScore = normText(toNumStrMaybe(listScore) || listScore) === normText(excelScore);
      const sameLevel = listLevel === excelLevel;

      if (!FORCE_OVERWRITE && sameScore && sameLevel) {
        plan.same.push({ ...ps, excelScore, excelLevel });
      } else {
        plan.toUpdate.push({ ...ps, excelScore, excelLevel, rec });
      }
    }

    dbg(`扫描差异：toUpdate=${plan.toUpdate.length}, same=${plan.same.length}, notFound=${plan.notFound.length}`);
    return plan;
  }

  /***********************
   * ✅ 仅按计划更新
   ***********************/
  async function runUpdatePlan() {
    const sheetName = currentSheetName;
    if (!sheetName || !studentDataBySheet[sheetName]) return alert("请选择有效的sheet！");

    // 生成计划（保证“开始更新”永远只对差异动手）
    lastPlan = buildUpdatePlan(sheetName);

    runLogs = [];
    runSummary = {
      sheet: sheetName,
      total: lastPlan.toUpdate.length,
      updated: 0,
      skipped: lastPlan.same.length,
      notFound: lastPlan.notFound.length,
      failed: 0,
      startedAt: nowIso(),
      finishedAt: "",
    };

    progressState = { sheet: sheetName, idx: 0, total: runSummary.total, name: "", phase: "准备开始" };
    updateStatProgress();

    runLogs.push({
      ts: nowIso(),
      sheet: sheetName,
      status: "START",
      name: "",
      sid: "",
      list_grade: "",
      list_level: "",
      excel_grade: "",
      excel_level: "",
      remark_preview: "",
      note: `开始：待更新=${lastPlan.toUpdate.length}, 一致跳过=${lastPlan.same.length}, 未找到=${lastPlan.notFound.length}`,
    });

    if (!lastPlan.toUpdate.length) {
      runSummary.finishedAt = nowIso();
      runLogs.push({
        ts: nowIso(),
        sheet: sheetName,
        status: "END",
        name: "",
        sid: "",
        list_grade: "",
        list_level: "",
        excel_grade: "",
        excel_level: "",
        remark_preview: "",
        note: "无需更新（全部一致或无可用记录）",
      });
      alert("无需更新：页面列表与Excel一致。");
      return;
    }

    for (let i = 0; i < lastPlan.toUpdate.length; i++) {
      if (isPaused) break;

      const item = lastPlan.toUpdate[i];
      const ps = { sid: item.sid, name: item.name, listLevel: item.listLevel, listScore: item.listScore };
      const rec = item.rec;

      const excelGradeStr = item.excelScore;
      const excelLevel = item.excelLevel;

      progressState.idx = i + 1;
      progressState.total = runSummary.total;
      progressState.name = ps.name;
      progressState.phase = "开始处理";
      updateStatProgress();

      const perStudentStart = Date.now();
      try {
        await Promise.race([
          updateOneStudent(ps, rec, excelGradeStr, excelLevel),
          (async () => {
            while (Date.now() - perStudentStart < PER_STUDENT_HARD_TIMEOUT_MS) await sleep(200);
            throw new Error(`单个学生处理超时>${PER_STUDENT_HARD_TIMEOUT_MS}ms，自动跳过`);
          })(),
        ]);

        runSummary.updated++;
        runLogs.push({
          ts: nowIso(),
          sheet: sheetName,
          status: "UPDATED",
          name: ps.name,
          sid: ps.sid || "",
          list_grade: ps.listScore,
          list_level: ps.listLevel,
          excel_grade: excelGradeStr,
          excel_level: excelLevel,
          remark_preview: remarkPreview(rec.评语),
          note: "已提交并等待加载完成",
        });
      } catch (e) {
        runSummary.failed++;
        runLogs.push({
          ts: nowIso(),
          sheet: sheetName,
          status: "FAILED",
          name: ps.name,
          sid: ps.sid || "",
          list_grade: ps.listScore,
          list_level: ps.listLevel,
          excel_grade: excelGradeStr,
          excel_level: excelLevel,
          remark_preview: remarkPreview(rec.评语),
          note: String(e?.message || e),
        });
        dbg(`失败/超时：${String(e?.message || e)}`);
        dbg(`最近消息：${messageHistory.slice(-10).map((x) => x.text).join(" | ") || "无"}`);

        progressState.phase = "失败";
        updateStatProgress(String(e?.message || e));
      }

      if (isPaused) {
        progressState.phase = "已暂停";
      } else {
        progressState.phase = "完成当前";
      }
      updateStatProgress();
    }

    runSummary.finishedAt = nowIso();
    runLogs.push({
      ts: nowIso(),
      sheet: sheetName,
      status: "END",
      name: "",
      sid: "",
      list_grade: "",
      list_level: "",
      excel_grade: "",
      excel_level: "",
      remark_preview: "",
      note: `完成：updated=${runSummary.updated}, skipped=${runSummary.skipped}, notFound=${runSummary.notFound}, failed=${runSummary.failed}`,
    });

    dbg(`完成：updated=${runSummary.updated}, skipped=${runSummary.skipped}, notFound=${runSummary.notFound}, failed=${runSummary.failed}`);
    alert(
      `执行完成（仅差异）！\n待更新:${runSummary.total}\n已更新:${runSummary.updated}\n一致跳过:${runSummary.skipped}\nExcel未找到:${runSummary.notFound}\n失败:${runSummary.failed}\n可点击“导出日志CSV”下载日志。`
    );
  }

  /***********************
   * 更新单个学生（点击→等PDF→填→提交→等加载完成）
   ***********************/
  async function updateOneStudent(pageStudent, rec, excelGradeStr, excelLevel) {
    dbg(`更新学生: ${pageStudent.name} (${pageStudent.listScore}/${pageStudent.listLevel} -> ${excelGradeStr}/${excelLevel})`);

    progressState.phase = "点击学生";
    updateStatProgress();

    // 用时间窗：点击前
    const sinceTime = Date.now();

    const studentElement = Array.from(document.querySelectorAll(studentCardSelector)).find((div) =>
      normText(div.textContent || "").includes(pageStudent.name)
    );
    if (!studentElement) throw new Error("找不到该学生行（可能列表未渲染/滚动）");

    const link = studentElement.querySelector(clickStudentLinkSelector);
    if (!link) throw new Error("找不到学生链接a");

    link.dispatchEvent(new MouseEvent("click", { bubbles: true, cancelable: true, view: window }));
    await sleep(AFTER_CLICK_DELAY);

    dbg("等待PDF加载流程...");
    progressState.phase = "等待加载完成";
    updateStatProgress();

    try {
      await waitPdfLoadFlow(sinceTime);
    } catch (e) {
      // 首个学生常见：页面进入时已自动加载过一次PDF，点击第一个学生时不再弹出加载消息。
      // 这种情况下 sinceTime 之后没有新 message，但页面控件已就绪，可以直接继续。
      if (!hasAnyMessageSince(sinceTime) && isMarkingUiReady()) {
        dbg("未捕捉到新的加载消息，但页面控件已就绪，跳过等待");
      } else {
        throw e;
      }
    }
    dbg("PDF加载完成且静默，开始填写");
    await sleep(STEP_DELAY_MS);

    const gradeInput =
      document.querySelector("input.ivu-input-number-input[placeholder='请输入作业成绩']") ||
      findInputLike(PAGE_INPUT_HINTS.grade, false);
    if (!gradeInput) throw new Error("未找到成绩输入框");

    const remarkInput = findInputLike(PAGE_INPUT_HINTS.remark, true) || findInputLike(PAGE_INPUT_HINTS.remark, false);

    dbg(`填成绩: ${excelGradeStr}`);
    progressState.phase = "填写成绩";
    updateStatProgress();
    fillInputControlled(gradeInput, excelGradeStr);
    await sleep(GRADE_INPUT_TEST_DELAY_MS);

    if (excelLevel) {
      dbg(`设置等级: ${excelLevel}`);
      progressState.phase = "设置抽样等级";
      updateStatProgress();
      await setSamplingLevel(excelLevel);
      await sleep(STEP_DELAY_MS);
    }

    if (remarkInput) {
      dbg(`填评语: ${remarkPreview(rec.评语)}`);
      progressState.phase = "填写评语";
      updateStatProgress();
      fillInputControlled(remarkInput, rec.评语 || "");
      await sleep(STEP_DELAY_MS);
    } else {
      dbg("未找到评语输入框（跳过评语）");
    }

    gradeInput.dispatchEvent(new Event("change", { bubbles: true }));
    gradeInput.dispatchEvent(new Event("blur", { bubbles: true }));
    await sleep(250);

    const submitButton = document.querySelector(submitButtonSelector);
    if (!submitButton) throw new Error("未找到提交按钮");

    const submitSince = Date.now();
    dbg("点击提交");
    progressState.phase = "提交";
    updateStatProgress();
    submitButton.click();

    dbg("提交后等待消息（Succeeded/加载完成）...");
    progressState.phase = "等待提交完成";
    updateStatProgress();
    await waitAfterSubmitFlow(submitSince);

    await sleep(300);
    progressState.phase = "完成";
    updateStatProgress();
  }

  /***********************
   * 启动：只启动轻量消息监听
   ***********************/
  startMessageObserverLight();

})();
