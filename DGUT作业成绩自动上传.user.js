// ==UserScript==
// @name         作业成绩自动填写（多Sheet手动选择·最终版+延时+日志）
// @namespace    http://tampermonkey.net/
// @version      6.0
// @description  多sheet读取并手动选择；自动扫描表头行解决“只识别课程设计”；成绩写入用原生setter+事件链解决“成绩不保存”；每次写入成绩后延时2秒测试；自动统计并可导出CSV日志
// @author       YourName
// @match        https://hw.dgut.edu.cn/teacher/homeworkPlan/*/mark
// @updateURL    https://cdn.jsdelivr.net/gh/wenguo/tampermonkey-scripts@main/DGUT%E4%BD%9C%E4%B8%9A%E6%88%90%E7%BB%A9%E8%87%AA%E5%8A%A8%E4%B8%8A%E4%BC%A0.user.js
// @downloadURL  https://cdn.jsdelivr.net/gh/wenguo/tampermonkey-scripts@main/DGUT%E4%BD%9C%E4%B8%9A%E6%88%90%E7%BB%A9%E8%87%AA%E5%8A%A8%E4%B8%8A%E4%BC%A0.user.js
// @grant        none
// @grant        none
// @require      https://cdn.jsdelivr.net/npm/xlsx/dist/xlsx.full.min.js
// ==/UserScript==

(function () {
  "use strict";

  /***********************
   * 可调配置
   ***********************/
  const submitButtonSelector = "button.score-operation-button.ivu-btn-success";
  const messageSelector = ".ivu-message-success";
  const studentCardSelector = ".ivu-card-body > div";
  const clickStudentLinkSelector = "a";

  const EXCEL_COLS = {
    id: ["学号", "学生学号", "ID", "StudentID"],
    name: ["姓名", "学生姓名", "Name"],
    grade: ["自动总分_调整后", "成绩"],
    remark: ["评语", "详细评语", "批阅"],
  };

  const PAGE_INPUT_HINTS = {
    grade: ["请输入作业成绩", "作业成绩", "成绩", "得分", "分数"],
    remark: ["评语", "详细评语", "批阅", "点评", "意见", "备注"],
  };

  // 是否强制覆盖（false=仅当页面成绩与Excel不一致才更新）
  const FORCE_OVERWRITE = false;

  // 点击学生后等待详情渲染
  const AFTER_CLICK_DELAY = 500;

  // 提交后等待返回/刷新
  const AFTER_SUBMIT_DELAY = 1000;

  // 等待“加载完成”提示最多次数/间隔
  const LOAD_MSG_MAX_RETRIES = 12;
  const LOAD_MSG_INTERVAL = 500;

  // ✅ 你要的：每次输入完成绩后额外延时2秒用于测试
  const GRADE_INPUT_TEST_DELAY_MS = 2000;

  const buttonStyle =
    "position: fixed; z-index: 9999; width: 190px; height: 40px; font-size: 14px; border: none; padding: 10px; cursor: pointer;";

  let isPaused = false;

  // { sheetName: [{学号,姓名,成绩,评语}, ...] }
  let studentDataBySheet = {};
  let currentSheetName = "";

  /***********************
   * 运行日志（统计+导出）
   ***********************/
  let runLogs = [];
  let runSummary = {
    sheet: "",
    total: 0,
    updated: 0,
    skipped: 0,
    notFound: 0,
    failed: 0,
    startedAt: "",
    finishedAt: "",
  };

  function nowIso() {
    const d = new Date();
    const pad = (n) => String(n).padStart(2, "0");
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(
      d.getMinutes()
    )}:${pad(d.getSeconds())}`;
  }

  function remarkPreview(s, maxLen = 20) {
    const t = String(s ?? "")
      .replace(/\s+/g, " ")
      .trim();
    return t.length > maxLen ? t.slice(0, maxLen) + "..." : t;
  }

  function addLog(row) {
    runLogs.push({
      ts: nowIso(),
      sheet: runSummary.sheet || "",
      ...row,
    });
  }

  /***********************
   * UI
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
    console.log("[AutoGrade] 当前选择sheet:", currentSheetName);
  });
  document.body.appendChild(sheetSelect);

  const startButton = document.createElement("button");
  startButton.textContent = "开始更新成绩";
  startButton.style = `${buttonStyle} top: 110px; left: 10px; background: gray; color: white; cursor: not-allowed;`;
  startButton.disabled = true;
  startButton.addEventListener("click", () => {
    if (!currentSheetName) {
      alert("请先在下拉框选择一个sheet！");
      return;
    }
    isPaused = false;
    compareAndUpdatePage().catch((e) => console.error("[AutoGrade] 执行异常:", e));
  });
  document.body.appendChild(startButton);

  const pauseButton = document.createElement("button");
  pauseButton.textContent = "暂停";
  pauseButton.style = `${buttonStyle} top: 160px; left: 10px; background: #ffc107; color: black;`;
  pauseButton.addEventListener("click", () => {
    isPaused = true;
    console.log("[AutoGrade] 更新已暂停");
  });
  document.body.appendChild(pauseButton);

  const exportLogButton = document.createElement("button");
  exportLogButton.textContent = "导出日志CSV";
  exportLogButton.style = `${buttonStyle} top: 210px; left: 10px; background: #17a2b8; color: white;`;
  exportLogButton.addEventListener("click", () => exportLogsCsv());
  document.body.appendChild(exportLogButton);

  const statBox = document.createElement("div");
  statBox.style =
    "position: fixed; z-index: 9999; top: 260px; left: 10px; width: 190px; padding: 8px; font-size: 12px; background: rgba(0,0,0,0.75); color: #fff; border-radius: 6px; line-height: 1.5;";
  statBox.textContent = "未加载Excel";
  document.body.appendChild(statBox);

  function setStat(text) {
    statBox.textContent = text;
  }

  /***********************
   * 基础工具
   ***********************/
  function sleep(ms) {
    return new Promise((r) => setTimeout(r, ms));
  }

  function normText(x) {
    return String(x ?? "")
      .replace(/\u00A0/g, " ") // NBSP
      .replace(/[\r\n\t]/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  }

  function fillSheetSelectWithCounts(dataBySheet) {
    sheetSelect.innerHTML = "";
    const opt0 = document.createElement("option");
    opt0.value = "";
    opt0.textContent = "请选择sheet";
    sheetSelect.appendChild(opt0);

    Object.keys(dataBySheet).forEach((sn) => {
      const opt = document.createElement("option");
      opt.value = sn;
      opt.textContent = `${sn}（${dataBySheet[sn].length}）`;
      sheetSelect.appendChild(opt);
    });

    sheetSelect.disabled = false;
    currentSheetName = "";
  }

  /***********************
   * ✅ 受控组件写入（解决成绩不保存）
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

    // 对 iView/ViewUI 等受控组件更友好
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

  function waitForLoadCompleteMessage(maxRetries = LOAD_MSG_MAX_RETRIES, interval = LOAD_MSG_INTERVAL) {
    return new Promise((resolve, reject) => {
      let attempts = 0;
      const timer = setInterval(() => {
        const messageElement = document.querySelector(messageSelector);
        const t = normText(messageElement?.textContent || "");
        if (messageElement && t.includes("加载完成")) {
          clearInterval(timer);
          resolve(true);
          return;
        }
        attempts++;
        if (attempts >= maxRetries) {
          clearInterval(timer);
          reject(new Error("加载完成消息超时"));
        }
      }, interval);
    });
  }

  /***********************
   * ✅ 表头扫描 + 模糊匹配（解决只识别课程设计）
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

    // 1) 严格等于
    for (const c of candNorm) {
      const idx = headersNorm.findIndex((h) => h === c);
      if (idx !== -1) return idx;
    }
    // 2) 包含（防止“姓名(必填)”等）
    for (const c of candNorm) {
      const idx = headersNorm.findIndex((h) => h.includes(c));
      if (idx !== -1) return idx;
    }
    return -1;
  }

  /***********************
   * Excel读取
   ***********************/
  function handleFile(event) {
    const file = event.target.files[0];
    if (!file) return;

    setStat("正在读取Excel...");

    const reader = new FileReader();
    reader.onload = function (e) {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });

      const sheetNames = workbook.SheetNames || [];
      if (sheetNames.length === 0) {
        alert("Excel中未发现任何sheet");
        setStat("Excel无sheet");
        return;
      }

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
        console.log(`[AutoGrade] sheet【${sn}】headerRow=${hdr.headerRowIndex + 1} records=${list.length}`);
      }

      if (Object.keys(tmp).length === 0) {
        alert("未能解析到可用sheet数据（打开控制台查看 skipped 原因）");
        setStat("Excel解析失败");
        console.warn("[AutoGrade] skipped:", skipped);
        return;
      }

      studentDataBySheet = tmp;

      fillSheetSelectWithCounts(studentDataBySheet);
      startButton.disabled = false;
      startButton.style.background = "#28a745";
      startButton.style.cursor = "pointer";

      setStat(`Excel加载完成\n可用sheet: ${Object.keys(tmp).length}/${sheetNames.length}\n请选择sheet后开始`);
      if (skipped.length) console.warn("[AutoGrade] skipped:", skipped);

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
        const right = div.querySelector("span[style*='float: right']");
        const grade = right?.textContent ? normText(right.textContent) : "";

        const text = normText(div.textContent || "");
        const idMatch = text.match(/(\d{6,12})/);
        const sid = idMatch ? idMatch[1] : "";

        return { sid, name, grade };
      })
      .filter((x) => x.name);

    console.log("[AutoGrade] 页面学生清单:", list);
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
   * 日志导出 CSV
   ***********************/
  function csvEscape(v) {
    const s = String(v ?? "");
    if (/[",\n]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
    return s;
  }

  function exportLogsCsv() {
    if (!runLogs.length) {
      alert("当前没有日志可导出（请先执行一次）");
      return;
    }

    const headers = ["ts", "sheet", "status", "name", "sid", "page_grade", "excel_grade", "remark_preview", "note"];
    const lines = [headers.join(",")];

    for (const r of runLogs) {
      lines.push(headers.map((k) => csvEscape(r[k])).join(","));
    }

    const csv = "\ufeff" + lines.join("\n"); // BOM for Excel UTF-8
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
   * 主流程
   ***********************/
  async function compareAndUpdatePage() {
    const sheetName = currentSheetName;
    if (!sheetName || !studentDataBySheet[sheetName]) {
      alert("请选择有效的sheet！");
      return;
    }

    const pageStudents = getStudentListFromPage();

    // init logs
    runLogs = [];
    runSummary = {
      sheet: sheetName,
      total: pageStudents.length,
      updated: 0,
      skipped: 0,
      notFound: 0,
      failed: 0,
      startedAt: nowIso(),
      finishedAt: "",
    };
    addLog({ status: "START", name: "", sid: "", page_grade: "", excel_grade: "", remark_preview: "", note: `开始，总人数=${runSummary.total}` });

    setStat(`开始执行\nsheet: ${sheetName}\n总人数: ${runSummary.total}`);

    for (let i = 0; i < pageStudents.length; i++) {
      if (isPaused) {
        setStat(`已暂停\n进度: ${i}/${runSummary.total}\n已更新: ${runSummary.updated}`);
        addLog({ status: "PAUSED", name: "", sid: "", page_grade: "", excel_grade: "", remark_preview: "", note: `暂停于 ${i}/${runSummary.total}` });
        break;
      }

      const ps = pageStudents[i];
      const rec = findStudentRecord(sheetName, ps);

      if (!rec) {
        runSummary.notFound++;
        addLog({
          status: "NOT_FOUND",
          name: ps.name,
          sid: ps.sid || "",
          page_grade: ps.grade || "",
          excel_grade: "",
          remark_preview: "",
          note: "Excel中未找到该学生",
        });
        setStat(`进行中 ${i + 1}/${runSummary.total}\n未找到: ${runSummary.notFound}\n已更新: ${runSummary.updated}`);
        continue;
      }

      const pageGrade = normText(ps.grade);
      const excelGrade = normText(rec.成绩);

      const shouldUpdate = FORCE_OVERWRITE ? true : pageGrade !== excelGrade;
      if (!shouldUpdate) {
        runSummary.skipped++;
        addLog({
          status: "SKIP",
          name: ps.name,
          sid: ps.sid || "",
          page_grade: pageGrade,
          excel_grade: excelGrade,
          remark_preview: remarkPreview(rec.评语),
          note: "成绩一致，跳过",
        });
        setStat(`进行中 ${i + 1}/${runSummary.total}\n跳过: ${runSummary.skipped}\n已更新: ${runSummary.updated}`);
        continue;
      }

      try {
        await updateStudent(ps, rec);
        runSummary.updated++;
        addLog({
          status: "UPDATED",
          name: ps.name,
          sid: ps.sid || "",
          page_grade: pageGrade,
          excel_grade: excelGrade,
          remark_preview: remarkPreview(rec.评语),
          note: "已提交更新",
        });
      } catch (err) {
        runSummary.failed++;
        addLog({
          status: "FAILED",
          name: ps.name,
          sid: ps.sid || "",
          page_grade: pageGrade,
          excel_grade: excelGrade,
          remark_preview: remarkPreview(rec.评语),
          note: String(err?.message || err || "unknown error"),
        });
      }

      setStat(`进行中 ${i + 1}/${runSummary.total}\n已更新: ${runSummary.updated}\n未找到: ${runSummary.notFound}\n失败: ${runSummary.failed}`);
    }

    runSummary.finishedAt = nowIso();
    addLog({
      status: "END",
      name: "",
      sid: "",
      page_grade: "",
      excel_grade: "",
      remark_preview: "",
      note: `完成 updated=${runSummary.updated}, skipped=${runSummary.skipped}, notFound=${runSummary.notFound}, failed=${runSummary.failed}`,
    });

    console.group("[AutoGrade] 本次执行统计");
    console.log(runSummary);
    console.table(runLogs);
    console.groupEnd();

    setStat(`完成\n总: ${runSummary.total}\n更新: ${runSummary.updated}\n跳过: ${runSummary.skipped}\n未找到: ${runSummary.notFound}\n失败: ${runSummary.failed}`);

    // 可选：完成后自动导出日志（不想自动导出就注释掉）
    // exportLogsCsv();

    alert(`执行完成！\n总: ${runSummary.total}\n更新: ${runSummary.updated}\n跳过: ${runSummary.skipped}\n未找到: ${runSummary.notFound}\n失败: ${runSummary.failed}\n可点击“导出日志CSV”下载日志。`);
  }

  async function updateStudent(pageStudent, rec) {
    return new Promise((resolve, reject) => {
      const studentElement = Array.from(document.querySelectorAll(studentCardSelector)).find((div) =>
        normText(div.textContent || "").includes(pageStudent.name)
      );

      if (!studentElement) return resolve();

      const link = studentElement.querySelector(clickStudentLinkSelector);
      if (!link) return resolve();

      link.dispatchEvent(new MouseEvent("click", { bubbles: true, cancelable: true, view: window }));

      setTimeout(async () => {
        try {
          try {
            await waitForLoadCompleteMessage();
          } catch (e) {
            console.warn("[AutoGrade] 未检测到加载完成提示，继续尝试填写:", pageStudent.name, e);
          }

          const gradeInput =
            document.querySelector("input.ivu-input-number-input[placeholder='请输入作业成绩']") ||
            findInputLike(PAGE_INPUT_HINTS.grade, false);

          const remarkInput =
            findInputLike(PAGE_INPUT_HINTS.remark, true) || findInputLike(PAGE_INPUT_HINTS.remark, false);

          if (!gradeInput) {
            console.error("[AutoGrade] 未找到成绩输入框，跳过:", pageStudent.name);
            return resolve();
          }

          // ✅ 成绩：抽取数字，避免“85分/良好(85)”导致InputNumber不认
          const gradeStr = String(rec.成绩 ?? "").match(/-?\d+(\.\d+)?/)?.[0] ?? "";
          fillInputControlled(gradeInput, gradeStr);

          // ✅ 你要的：写完成绩后延时2秒测试
          await sleep(GRADE_INPUT_TEST_DELAY_MS);

          // 评语：允许为空
          if (remarkInput) fillInputControlled(remarkInput, rec.评语 || "");

          const submitButton = document.querySelector(submitButtonSelector);
          if (submitButton) submitButton.click();
          else console.warn("[AutoGrade] 未找到提交按钮:", submitButtonSelector);

          setTimeout(resolve, AFTER_SUBMIT_DELAY);
        } catch (err) {
          reject(err);
        }
      }, AFTER_CLICK_DELAY);
    });
  }
})();