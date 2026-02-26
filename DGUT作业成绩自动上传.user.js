// ==UserScript==
// @name         作业成绩自动填写（多Sheet手动选择·最终版+延时+日志+抽样等级）
// @namespace    http://tampermonkey.net/
// @version      6.6
// @description  多sheet手动选择；MutationObserver增量捕捉message；等待“加载完成”后再等待消息静默；比较成绩+抽样等级；步骤延时；日志CSV；防卡死超时跳过
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
  const studentCardSelector = ".ivu-card-body > div";
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

  // 站点真实提示文案（你日志里是“开始加载作业文件”）
  const MSG_START_LOAD = "开始加载作业文件";
  const MSG_LOAD_DONE = "加载完成";
  const MSG_SAVE_UPLOADING = "保存并上传中";
  const MSG_SUCCEEDED = "Succeeded";

  const FORCE_OVERWRITE = false;

  const AFTER_CLICK_DELAY = 300;
  const STEP_DELAY_MS = 600;

  // ✅ 成绩输入后额外延时 2 秒（测试稳定性）
  const GRADE_INPUT_TEST_DELAY_MS = 2000;

  // 等消息超时
  const WAIT_START_LOAD_MAX_MS = 8000;
  const WAIT_LOAD_DONE_MAX_MS = 30000;

  // ✅ 加载完成后静默一小段
  const AFTER_LOAD_SILENCE_MS = 700;

  const AFTER_SUBMIT_DELAY = 1000;

  // ✅ 点击提交后：等待“Succeeded”（可选）+ 最终“加载完成”（必须）
  const WAIT_SUBMIT_SUCCEEDED_MAX_MS = 12000; // 有时会很快出现
  const WAIT_AFTER_SUBMIT_LOAD_DONE_MAX_MS = 30000; // 提交后重新加载pdf

  // ✅ 防止单个学生卡死
  const PER_STUDENT_HARD_TIMEOUT_MS = 60000;

  const DEBUG = true;

  const buttonStyle =
    "position: fixed; z-index: 9999; width: 190px; height: 40px; font-size: 14px; border: none; padding: 10px; cursor: pointer;";

  let isPaused = false;
  let studentDataBySheet = {};
  let currentSheetName = "";

  /***********************
   * Debug UI
   ***********************/
  const dbgBox = document.createElement("div");
  dbgBox.style =
    "position: fixed; z-index: 9999; top: 10px; right: 10px; width: 560px; max-height: 360px; overflow: auto; padding: 10px; font-size: 12px; background: rgba(0,0,0,0.70); color: #fff; border-radius: 8px; line-height: 1.5; white-space: pre-wrap;";
  dbgBox.textContent = "Debug: ready\n";
  document.body.appendChild(dbgBox);

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
    while (dbgLines.length > 80) dbgLines.shift();
    dbgBox.textContent = dbgLines.join("\n");
  }

  /***********************
   * Status UI
   ***********************/
  const statBox = document.createElement("div");
  statBox.style =
    "position: fixed; z-index: 9999; top: 260px; left: 10px; width: 190px; padding: 8px; font-size: 12px; background: rgba(0,0,0,0.75); color: #fff; border-radius: 6px; line-height: 1.5;";
  statBox.textContent = "未加载Excel";
  document.body.appendChild(statBox);
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

  function remarkPreview(s, maxLen = 20) {
    const t = String(s ?? "").replace(/\s+/g, " ").trim();
    return t.length > maxLen ? t.slice(0, maxLen) + "..." : t;
  }

  /***********************
   * ✅ 轻量消息捕捉：只处理“新增的 message 节点”
   ***********************/
  const messageHistory = []; // {ts, tms, seq, text}
  const MESSAGE_HISTORY_MAX = 160;
  let lastMessageAt = 0;
  let msgSeq = 0;

  function pushMessage(text) {
    const t = normText(text);
    if (!t) return;

    const last = messageHistory.length ? messageHistory[messageHistory.length - 1].text : "";
    if (t === last) return; // 去重

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
      const spans = el.querySelectorAll(messageSpanSelector);
      spans.forEach((sp) => pushMessage(sp.innerText || sp.textContent || ""));
    }
  }

  function startMessageObserverLight() {
    // 初始已有 message
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
    dbg("MessageObserver 已启动（轻量：仅新增节点）");
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

  // ✅ 等待PDF加载（点击学生后）
  async function waitPdfLoadFlow(sinceTimeMs) {
    // start 是可选的：有时只弹“加载完成”
    await waitMessageSinceTime(MSG_START_LOAD, WAIT_START_LOAD_MAX_MS, sinceTimeMs, "start");
    const done = await waitMessageSinceTime(MSG_LOAD_DONE, WAIT_LOAD_DONE_MAX_MS, sinceTimeMs, "done");
    if (!done.ok) throw new Error("等待PDF“加载完成”超时（看右上角 Debug 最近消息）");
    await waitSilence(AFTER_LOAD_SILENCE_MS, 5000);
  }

  // ✅ 点击提交后：等待保存/成功（可选）+ 最终“加载完成”（必须）
  async function waitAfterSubmitFlow(sinceTimeMs) {
    // 有些时候会出现“保存并上传中...”“Succeeded”
    await waitMessageSinceTime(MSG_SAVE_UPLOADING, 5000, sinceTimeMs, "uploading"); // 可选，短等
    await waitMessageSinceTime(MSG_SUCCEEDED, WAIT_SUBMIT_SUCCEEDED_MAX_MS, sinceTimeMs, "succeeded"); // 可选

    // 提交后往往会重新触发加载pdf：开始加载作业文件 -> 加载完成
    await waitMessageSinceTime(MSG_START_LOAD, 12000, sinceTimeMs, "reload_start"); // 可选
    const done = await waitMessageSinceTime(MSG_LOAD_DONE, WAIT_AFTER_SUBMIT_LOAD_DONE_MAX_MS, sinceTimeMs, "reload_done");
    if (!done.ok) throw new Error("提交后等待“加载完成”超时（可能保存未生效或页面未刷新）");

    await waitSilence(AFTER_LOAD_SILENCE_MS, 5000);
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
   * 抽样等级
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
   * UI：文件/Sheet/开始/暂停/导出日志
   ***********************/
  const fileInput = document.createElement("input");
  fileInput.type = "file";
  fileInput.accept = ".xlsx,.xls";
  fileInput.style = `${buttonStyle} top: 10px; left: 10px; background: white; color: black; border: 1px solid black;`;
  fileInput.addEventListener("change", handleFile);
  document.body.appendChild(fileInput);

  const sheetSelect = document.createElement("select");
  sheetSelect.style =
    "position: fixed; z-index: 9999; top: 60px; left: 10px; width: 190px; height: 40px; font-size: 14px; border: 1px solid #333; background: #fff;";
  sheetSelect.disabled = true;
  sheetSelect.addEventListener("change", () => {
    currentSheetName = sheetSelect.value || "";
    dbg(`选择sheet: ${currentSheetName}`);
  });
  document.body.appendChild(sheetSelect);

  const startButton = document.createElement("button");
  startButton.textContent = "开始更新成绩";
  startButton.style = `${buttonStyle} top: 110px; left: 10px; background: gray; color: white; cursor: not-allowed;`;
  startButton.disabled = true;
  startButton.addEventListener("click", () => {
    if (!currentSheetName) return alert("请先在下拉框选择一个sheet！");
    isPaused = false;
    compareAndUpdatePage().catch((e) => console.error("[AutoGrade] 执行异常:", e));
  });
  document.body.appendChild(startButton);

  const pauseButton = document.createElement("button");
  pauseButton.textContent = "暂停";
  pauseButton.style = `${buttonStyle} top: 160px; left: 10px; background: #ffc107; color: black;`;
  pauseButton.addEventListener("click", () => {
    isPaused = true;
    dbg("已暂停");
  });
  document.body.appendChild(pauseButton);

  const exportLogButton = document.createElement("button");
  exportLogButton.textContent = "导出日志CSV";
  exportLogButton.style = `${buttonStyle} top: 210px; left: 10px; background: #17a2b8; color: white;`;
  exportLogButton.addEventListener("click", () => exportLogsCsv());
  document.body.appendChild(exportLogButton);

  /***********************
   * 日志CSV
   ***********************/
  let runLogs = [];
  let runSummary = { sheet: "", total: 0, updated: 0, skipped: 0, notFound: 0, failed: 0, startedAt: "", finishedAt: "" };

  function csvEscape(v) {
    const s = String(v ?? "");
    if (/[",\n]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
    return s;
  }

  function exportLogsCsv() {
    if (!runLogs.length) return alert("当前没有日志可导出（请先执行一次）");
    const headers = ["ts", "sheet", "status", "name", "sid", "page_grade", "excel_grade", "page_level", "excel_level", "remark_preview", "note"];
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
          .map((r) => ({
            学号: idIdx !== -1 ? normText(r[idIdx]) : "",
            姓名: normText(r[nameIdx]),
            成绩: normText(r[gradeIdx]),
            评语: remarkIdx !== -1 ? normText(r[remarkIdx]) : "",
          }))
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

      startButton.disabled = false;
      startButton.style.background = "#28a745";
      startButton.style.cursor = "pointer";

      setStat(`Excel加载完成\n可用sheet: ${Object.keys(tmp).length}/${sheetNames.length}\n请选择sheet后开始`);
      dbg(`Excel加载完成 可用sheet=${Object.keys(tmp).length}/${sheetNames.length}`);
      if (skipped.length) dbg(`跳过sheet: ${skipped.join(" | ")}`);

      alert("Excel加载成功！请选择对应sheet后点击开始。");
    };

    reader.readAsArrayBuffer(file);
  }

  /***********************
   * 页面学生列表 & 匹配
   ***********************/
  function getStudentListFromPage() {
    const elements = document.querySelectorAll(studentCardSelector);
    const list = Array.from(elements)
      .map((div) => {
        const link = div.querySelector(clickStudentLinkSelector);
        const name = link?.firstChild?.textContent ? normText(link.firstChild.textContent) : "";
        const text = normText(div.textContent || "");
        const idMatch = text.match(/(\d{6,12})/);
        const sid = idMatch ? idMatch[1] : "";
        return { sid, name };
      })
      .filter((x) => x.name);
    dbg(`页面学生数=${list.length}`);
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
   * 主流程
   ***********************/
  async function compareAndUpdatePage() {
    const sheetName = currentSheetName;
    if (!sheetName || !studentDataBySheet[sheetName]) return alert("请选择有效的sheet！");

    const pageStudents = getStudentListFromPage();

    runLogs = [];
    runSummary = { sheet: sheetName, total: pageStudents.length, updated: 0, skipped: 0, notFound: 0, failed: 0, startedAt: nowIso(), finishedAt: "" };
    runLogs.push({ ts: nowIso(), sheet: sheetName, status: "START", name: "", sid: "", page_grade: "", excel_grade: "", page_level: "", excel_level: "", remark_preview: "", note: `开始，总人数=${pageStudents.length}` });

    for (let i = 0; i < pageStudents.length; i++) {
      if (isPaused) break;

      const ps = pageStudents[i];
      const rec = findStudentRecord(sheetName, ps);

      if (!rec) {
        runSummary.notFound++;
        runLogs.push({ ts: nowIso(), sheet: sheetName, status: "NOT_FOUND", name: ps.name, sid: ps.sid || "", page_grade: "", excel_grade: "", page_level: "", excel_level: "", remark_preview: "", note: "Excel中未找到该学生" });
        continue;
      }

      const excelGradeStr = String(rec.成绩 ?? "").match(/-?\d+(\.\d+)?/)?.[0] ?? "";
      const excelScoreNum = Number(excelGradeStr);
      const excelLevel = Number.isNaN(excelScoreNum) ? "" : scoreToLevel(excelScoreNum);

      const perStudentStart = Date.now();
      try {
        const res = await Promise.race([
          openCompareAndMaybeUpdate(ps, rec, excelGradeStr, excelLevel),
          (async () => {
            while (Date.now() - perStudentStart < PER_STUDENT_HARD_TIMEOUT_MS) await sleep(200);
            throw new Error(`单个学生处理超时>${PER_STUDENT_HARD_TIMEOUT_MS}ms，自动跳过`);
          })(),
        ]);

        if (res.status === "SKIP") {
          runSummary.skipped++;
          runLogs.push({ ts: nowIso(), sheet: sheetName, status: "SKIP", name: ps.name, sid: ps.sid || "", page_grade: res.pageGrade, excel_grade: excelGradeStr, page_level: res.pageLevel, excel_level: excelLevel, remark_preview: remarkPreview(rec.评语), note: "成绩与等级均一致，跳过" });
        } else {
          runSummary.updated++;
          runLogs.push({ ts: nowIso(), sheet: sheetName, status: "UPDATED", name: ps.name, sid: ps.sid || "", page_grade: res.pageGrade, excel_grade: excelGradeStr, page_level: res.pageLevel, excel_level: excelLevel, remark_preview: remarkPreview(rec.评语), note: "已提交更新并等待加载完成" });
        }
      } catch (e) {
        runSummary.failed++;
        runLogs.push({ ts: nowIso(), sheet: sheetName, status: "FAILED", name: ps.name, sid: ps.sid || "", page_grade: "", excel_grade: excelGradeStr, page_level: "", excel_level: excelLevel, remark_preview: remarkPreview(rec.评语), note: String(e?.message || e) });
        dbg(`失败/超时：${String(e?.message || e)}`);
        dbg(`最近消息：${messageHistory.slice(-10).map((x) => x.text).join(" | ") || "无"}`);
      }

      setStat(`进行中 ${i + 1}/${runSummary.total}\n更新:${runSummary.updated} 跳过:${runSummary.skipped}\n未找到:${runSummary.notFound} 失败:${runSummary.failed}`);
    }

    runSummary.finishedAt = nowIso();
    runLogs.push({ ts: nowIso(), sheet: sheetName, status: "END", name: "", sid: "", page_grade: "", excel_grade: "", page_level: "", excel_level: "", remark_preview: "", note: `完成 updated=${runSummary.updated}, skipped=${runSummary.skipped}, notFound=${runSummary.notFound}, failed=${runSummary.failed}` });

    dbg(`完成：updated=${runSummary.updated}, skipped=${runSummary.skipped}, notFound=${runSummary.notFound}, failed=${runSummary.failed}`);
    alert(`执行完成！\n总:${runSummary.total}\n更新:${runSummary.updated}\n跳过:${runSummary.skipped}\n未找到:${runSummary.notFound}\n失败:${runSummary.failed}\n可点击“导出日志CSV”下载日志。`);
  }

  async function openCompareAndMaybeUpdate(pageStudent, rec, excelGradeStr, excelLevel) {
    dbg(`处理学生: ${pageStudent.name}`);

    // ✅ 关键：用时间窗（解决自动加载/竞态）
    const sinceTime = Date.now();

    const studentElement = Array.from(document.querySelectorAll(studentCardSelector)).find((div) =>
      normText(div.textContent || "").includes(pageStudent.name)
    );
    if (!studentElement) return { status: "SKIP", pageGrade: "", pageLevel: "" };

    const link = studentElement.querySelector(clickStudentLinkSelector);
    if (!link) return { status: "SKIP", pageGrade: "", pageLevel: "" };

    link.dispatchEvent(new MouseEvent("click", { bubbles: true, cancelable: true, view: window }));
    await sleep(AFTER_CLICK_DELAY);

    dbg("等待PDF加载流程...");
    await waitPdfLoadFlow(sinceTime);
    dbg("PDF加载完成且静默，开始填写");
    await sleep(STEP_DELAY_MS);

    const gradeInput =
      document.querySelector("input.ivu-input-number-input[placeholder='请输入作业成绩']") ||
      findInputLike(PAGE_INPUT_HINTS.grade, false);
    if (!gradeInput) throw new Error("未找到成绩输入框");

    const remarkInput = findInputLike(PAGE_INPUT_HINTS.remark, true) || findInputLike(PAGE_INPUT_HINTS.remark, false);

    const pageGradeNow = normText(gradeInput.value || "");
    const pageGradeNumStr = pageGradeNow.match(/-?\d+(\.\d+)?/)?.[0] ?? pageGradeNow;
    const pageLevel = getSamplingLevelFromPage();

    const sameGrade = normText(pageGradeNumStr) === normText(excelGradeStr);
    const sameLevel = normText(pageLevel) === normText(excelLevel);

    dbg(`比较：页面成绩=${pageGradeNumStr} 目标=${excelGradeStr} sameGrade=${sameGrade}`);
    dbg(`比较：页面等级=${pageLevel} 目标=${excelLevel} sameLevel=${sameLevel}`);

    if (!FORCE_OVERWRITE && sameGrade && sameLevel) {
      return { status: "SKIP", pageGrade: pageGradeNumStr, pageLevel };
    }

    dbg(`填成绩: ${excelGradeStr}`);
    fillInputControlled(gradeInput, excelGradeStr);
    await sleep(GRADE_INPUT_TEST_DELAY_MS);

    if (excelLevel) {
      dbg(`设置等级: ${excelLevel}`);
      await setSamplingLevel(excelLevel);
      await sleep(STEP_DELAY_MS);
    }

    if (remarkInput) {
      dbg(`填评语: ${remarkPreview(rec.评语)}`);
      fillInputControlled(remarkInput, rec.评语 || "");
      await sleep(STEP_DELAY_MS);
    } else {
      dbg("未找到评语输入框（跳过评语）");
    }

    gradeInput.dispatchEvent(new Event("change", { bubbles: true }));
    gradeInput.dispatchEvent(new Event("blur", { bubbles: true }));
    await sleep(200);

    const submitButton = document.querySelector(submitButtonSelector);
    if (!submitButton) throw new Error("未找到提交按钮");

    // ✅ 点击提交后，再等一轮“提交->重新加载->加载完成”
    const submitSince = Date.now();
    dbg("点击提交");
    submitButton.click();

    dbg("提交后等待消息（Succeeded/加载完成）...");
    await waitAfterSubmitFlow(submitSince);

    await sleep(AFTER_SUBMIT_DELAY);
    return { status: "UPDATED", pageGrade: pageGradeNumStr, pageLevel };
  }

  /***********************
   * 启动：只启动轻量消息监听（不触发任何等待/点击）
   ***********************/
  startMessageObserverLight();

})();