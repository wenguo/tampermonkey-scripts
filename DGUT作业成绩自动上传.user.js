// ==UserScript==
// @name         作业成绩自动填写（多Sheet适配版）
// @namespace    http://tampermonkey.net/
// @version      4.0
// @description  从Excel多sheet读取成绩/评语并自动填写到页面；按页面innerText自动匹配sheet；成绩列：自动总分_调整后 或 成绩；评语列：评语 或 详细评语
// @author       YourName
// @match        https://hw.dgut.edu.cn/teacher/homeworkPlan/*/mark
// @grant        none
// @require      https://cdn.jsdelivr.net/npm/xlsx/dist/xlsx.full.min.js
// ==/UserScript==

(function () {
  "use strict";

  /***********************
   * 用户可调配置
   ***********************/
  const submitButtonSelector = "button.score-operation-button.ivu-btn-success"; // 提交按钮
  const messageSelector = ".ivu-message-success"; // 成功消息框（用于检测“加载完成”）
  const studentCardSelector = ".ivu-card-body > div"; // 学生列表卡片
  const clickStudentLinkSelector = "a"; // 学生卡片内点击进入详情的链接

  // Excel列名优先级
  const EXCEL_COLS = {
    id: ["学号", "学生学号", "ID", "StudentID"],
    name: ["姓名", "学生姓名", "Name"],
    grade: ["自动总分_调整后", "成绩"],
    remark: ["评语", "详细评语"],
  };

  // 页面输入框查找关键字（placeholder/aria-label/name/id/class/附近文本）
  const PAGE_INPUT_HINTS = {
    grade: ["请输入作业成绩", "作业成绩", "成绩", "得分", "分数"],
    remark: ["评语", "批阅", "点评", "意见", "备注"],
  };

  const buttonStyle =
    "position: fixed; z-index: 9999; width: 170px; height: 40px; font-size: 14px; border: none; padding: 10px; cursor: pointer;";

  let isPaused = false;

  // 多sheet数据：{ sheetName: [{学号,姓名,成绩,评语}, ...] }
  let studentDataBySheet = {};
  let currentSheetName = "";

  /***********************
   * UI 控件
   ***********************/
  const fileInput = document.createElement("input");
  fileInput.type = "file";
  fileInput.accept = ".xlsx,.xls";
  fileInput.style = `${buttonStyle} top: 10px; left: 10px; background: white; color: black; border: 1px solid black;`;
  fileInput.addEventListener("change", handleFile);
  document.body.appendChild(fileInput);

  const sheetSelect = document.createElement("select");
  sheetSelect.style =
    "position: fixed; z-index: 9999; top: 60px; left: 10px; width: 170px; height: 40px; font-size: 14px; border: 1px solid #333; background: #fff;";
  sheetSelect.disabled = true;
  sheetSelect.addEventListener("change", () => {
    currentSheetName = sheetSelect.value || "";
    console.log("[AutoGrade] 手动切换sheet:", currentSheetName);
  });
  document.body.appendChild(sheetSelect);

  const startButton = document.createElement("button");
  startButton.textContent = "开始更新成绩";
  startButton.style = `${buttonStyle} top: 110px; left: 10px; background: gray; color: white; cursor: not-allowed;`;
  startButton.disabled = true;
  startButton.addEventListener("click", () => {
    isPaused = false;
    compareAndUpdatePage();
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

  /***********************
   * 工具函数
   ***********************/
  function normalizeText(s) {
    return String(s ?? "")
      .replace(/\s+/g, " ")
      .trim();
  }

  function pickFirstExistingHeader(headers, candidates) {
    for (const c of candidates) {
      const idx = headers.indexOf(c);
      if (idx !== -1) return { name: c, index: idx };
    }
    return null;
  }

  function getPageInnerText() {
    // innerText 比 textContent 更贴近用户看到的文本
    return normalizeText(document.body?.innerText || "");
  }

  function autoMatchSheetName(sheetNames) {
    const text = getPageInnerText();

    // 常见页面文本类似：'实验1-xxx的作业批阅'
    // 取 “的作业批阅” 之前的标题作为候选
    let candidate = "";
    const m = text.match(/([^\n]{2,80}?)的作业批阅/);
    if (m && m[1]) candidate = normalizeText(m[1]);

    // 若没匹配到，就退化为 document.title / URL 相关信息
    if (!candidate) candidate = normalizeText(document.title || "");

    // 1) 优先：sheetName 完整包含在 candidate 中
    for (const sn of sheetNames) {
      if (candidate.includes(sn)) return sn;
    }

    // 2) 次优：candidate 包含 “实验1-xxx”，sheetName 是“实验1-xxx”
    // 尝试抽取 “实验\d-xxx” 或 “课程设计”
    const m2 = candidate.match(/(实验\d+\s*[-—]\s*[^-—]{2,80}|课程设计)/);
    const key = m2 ? normalizeText(m2[1]).replace(/\s*[-—]\s*/g, "-") : "";
    if (key) {
      for (const sn of sheetNames) {
        const sn2 = normalizeText(sn).replace(/\s*[-—]\s*/g, "-");
        if (sn2 === key) return sn;
      }
    }

    // 3) 兜底：只要文本里出现 sheetName 的关键前缀（比如“实验1”）
    const m3 = candidate.match(/实验(\d+)/);
    if (m3) {
      const prefix = `实验${m3[1]}`;
      const hit = sheetNames.find((sn) => normalizeText(sn).startsWith(prefix));
      if (hit) return hit;
    }

    return ""; // 未匹配到
  }

  function fillSheetSelect(sheetNames, selected) {
    sheetSelect.innerHTML = "";
    const opt0 = document.createElement("option");
    opt0.value = "";
    opt0.textContent = selected ? `已自动匹配：${selected}` : "未匹配到，请手动选择";
    sheetSelect.appendChild(opt0);

    for (const sn of sheetNames) {
      const opt = document.createElement("option");
      opt.value = sn;
      opt.textContent = sn;
      sheetSelect.appendChild(opt);
    }

    sheetSelect.disabled = false;
    if (selected) sheetSelect.value = selected;
    currentSheetName = selected || "";
  }

  function findInputLike(hints, preferTextarea = false) {
    const els = Array.from(document.querySelectorAll("input, textarea"))
      .filter((el) => !el.disabled && el.offsetParent !== null); // 可见且可用

    const score = (el) => {
      const attrs = [
        el.getAttribute("placeholder"),
        el.getAttribute("aria-label"),
        el.getAttribute("name"),
        el.getAttribute("id"),
        el.className,
      ]
        .map((x) => normalizeText(x))
        .join(" | ");

      let s = 0;
      for (const h of hints) {
        if (attrs.includes(h)) s += 5;
      }

      // 再看一下附近文本（label/父节点）
      const near = normalizeText(el.parentElement?.innerText || "");
      for (const h of hints) {
        if (near.includes(h)) s += 2;
      }

      // textarea 适合评语
      if (preferTextarea && el.tagName.toLowerCase() === "textarea") s += 3;
      if (!preferTextarea && el.tagName.toLowerCase() === "input") s += 1;

      return s;
    };

    els.sort((a, b) => score(b) - score(a));
    return els[0] && score(els[0]) > 0 ? els[0] : null;
  }

  function setNativeValue(el, value) {
    const v = value == null ? "" : String(value);
    const tag = el.tagName.toLowerCase();
    if (tag === "input" || tag === "textarea") {
      el.focus();
      el.value = v;
      // 触发 Vue/React 等框架监听
      el.dispatchEvent(new Event("input", { bubbles: true }));
      el.dispatchEvent(new Event("change", { bubbles: true }));
      el.blur();
    }
  }

  function waitForLoadCompleteMessage(maxRetries = 12, interval = 500) {
    return new Promise((resolve, reject) => {
      let attempts = 0;
      const timer = setInterval(() => {
        const messageElement = document.querySelector(messageSelector);
        const text = normalizeText(messageElement?.textContent || "");
        if (messageElement && text.includes("加载完成")) {
          clearInterval(timer);
          console.log("[AutoGrade] 检测到加载完成消息");
          resolve();
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
   * Excel 读取（多Sheet）
   ***********************/
  function handleFile(event) {
    const file = event.target.files[0];
    if (!file) {
      alert("请上传有效的Excel文件");
      return;
    }

    const reader = new FileReader();
    reader.onload = function (e) {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });

      const sheetNames = workbook.SheetNames || [];
      if (sheetNames.length === 0) {
        alert("Excel中未发现任何sheet");
        return;
      }

      const tmp = {};
      for (const sn of sheetNames) {
        const ws = workbook.Sheets[sn];
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
        if (!rows || rows.length < 2) continue;

        const headers = rows[0].map((h) => normalizeText(h));
        const idH = pickFirstExistingHeader(headers, EXCEL_COLS.id);
        const nameH = pickFirstExistingHeader(headers, EXCEL_COLS.name);
        const gradeH = pickFirstExistingHeader(headers, EXCEL_COLS.grade);
        const remarkH = pickFirstExistingHeader(headers, EXCEL_COLS.remark);

        if (!nameH || !gradeH) {
          console.warn(`[AutoGrade] sheet【${sn}】缺少必要列：姓名/成绩（允许成绩列为：自动总分_调整后 或 成绩）`);
          continue;
        }

        const list = rows
          .slice(1)
          .map((r) => {
            const obj = {
              学号: idH ? normalizeText(r[idH.index]) : "",
              姓名: normalizeText(r[nameH.index]),
              成绩: normalizeText(r[gradeH.index]),
              评语: remarkH ? normalizeText(r[remarkH.index]) : "",
            };
            return obj;
          })
          .filter((x) => x.姓名);

        tmp[sn] = list;
        console.log(`[AutoGrade] sheet【${sn}】加载记录数:`, list.length);
      }

      if (Object.keys(tmp).length === 0) {
        alert("未能从Excel中解析到可用数据（请检查表头是否包含：姓名 + (自动总分_调整后/成绩) + (评语/详细评语)）");
        return;
      }

      studentDataBySheet = tmp;

      const autoSheet = autoMatchSheetName(Object.keys(studentDataBySheet));
      fillSheetSelect(Object.keys(studentDataBySheet), autoSheet);

      alert(
        autoSheet
          ? `Excel加载成功！已自动匹配sheet：${autoSheet}\n如不正确，可在下拉框手动切换。`
          : "Excel加载成功！但未自动匹配到sheet，请在下拉框手动选择。"
      );

      // 启用开始按钮
      startButton.disabled = false;
      startButton.style.background = "#28a745";
      startButton.style.cursor = "pointer";
    };

    reader.readAsArrayBuffer(file);
  }

  /***********************
   * 页面读取学生清单 & 更新
   ***********************/
  function getStudentListFromPage() {
    const studentElements = document.querySelectorAll(studentCardSelector);
    const studentList = Array.from(studentElements)
      .map((div) => {
        const link = div.querySelector(clickStudentLinkSelector);
        const right = div.querySelector("span[style*='float: right']");
        const text = normalizeText(div.textContent || "");

        // 尝试从卡片文本提取学号（如果页面里有的话）
        const idMatch = text.match(/(\d{6,12})/); // 6~12位数字
        const sid = idMatch ? idMatch[1] : "";

        const name = link && link.firstChild ? normalizeText(link.firstChild.textContent) : "";
        const grade = right ? normalizeText(right.textContent) : "";

        return { sid, name, grade };
      })
      .filter((x) => x.name);

    console.log("[AutoGrade] 页面学生清单:", studentList);
    return studentList;
  }

  function findStudentRecord(sheetName, pageStudent) {
    const arr = studentDataBySheet[sheetName] || [];
    // 优先按学号匹配（如果页面能解析出来）
    if (pageStudent.sid) {
      const hit = arr.find((x) => x.学号 && x.学号 === pageStudent.sid);
      if (hit) return hit;
    }
    // 再按姓名匹配
    return arr.find((x) => x.姓名 === pageStudent.name) || null;
  }

  async function compareAndUpdatePage() {
    const sheetName = currentSheetName || sheetSelect.value || "";
    if (!sheetName) {
      alert("请先在下拉框选择正确的sheet（或确保自动匹配成功）");
      return;
    }
    if (!studentDataBySheet[sheetName]) {
      alert(`未找到sheet数据：${sheetName}`);
      return;
    }

    console.log("[AutoGrade] 使用sheet:", sheetName, "记录数:", studentDataBySheet[sheetName].length);

    const pageStudents = getStudentListFromPage();
    for (const ps of pageStudents) {
      if (isPaused) {
        console.log("[AutoGrade] 更新已暂停");
        break;
      }

      const rec = findStudentRecord(sheetName, ps);
      if (!rec) {
        console.log(`[AutoGrade] 未找到学生: ${ps.name}（学号:${ps.sid || "-"}） 的Excel记录`);
        continue;
      }

      // 只对成绩不一致的更新（页面成绩可能为空/“-”）
      const pageGrade = normalizeText(ps.grade);
      const excelGrade = normalizeText(rec.成绩);

      if (pageGrade !== excelGrade) {
        console.log(`[AutoGrade] 成绩不一致 -> 更新: ${ps.name} | page=${pageGrade} | excel=${excelGrade}`);
        await updateStudent(ps, rec);
      } else {
        console.log(`[AutoGrade] 成绩一致 -> 跳过: ${ps.name} | ${pageGrade}`);
      }
    }

    console.log("[AutoGrade] 执行完成");
  }

  function updateStudent(pageStudent, rec) {
    return new Promise((resolve) => {
      const studentElement = Array.from(document.querySelectorAll(studentCardSelector)).find((div) =>
        normalizeText(div.textContent || "").includes(pageStudent.name)
      );

      if (!studentElement) return resolve();

      const link = studentElement.querySelector(clickStudentLinkSelector);
      if (!link) return resolve();

      link.dispatchEvent(new MouseEvent("click", { bubbles: true, cancelable: true, view: window }));

      setTimeout(async () => {
        try {
          await waitForLoadCompleteMessage();
        } catch (e) {
          console.warn("[AutoGrade] 未检测到加载完成，尝试继续填写（可能页面提示不同）", e);
        }

        // 智能找输入框（避免旧selector失效）
        const gradeInput =
          document.querySelector("input.ivu-input-number-input[placeholder='请输入作业成绩']") ||
          findInputLike(PAGE_INPUT_HINTS.grade, false);

        const remarkInput =
          findInputLike(PAGE_INPUT_HINTS.remark, true) || findInputLike(PAGE_INPUT_HINTS.remark, false);

        if (!gradeInput) {
          console.error("[AutoGrade] 未找到成绩输入框，跳过:", pageStudent.name);
          return resolve();
        }

        setNativeValue(gradeInput, rec.成绩);

        if (remarkInput) {
          setNativeValue(remarkInput, rec.评语 || "");
        } else {
          console.warn("[AutoGrade] 未找到评语输入框（可忽略）:", pageStudent.name);
        }

        const submitButton = document.querySelector(submitButtonSelector);
        if (submitButton) submitButton.click();
        else console.warn("[AutoGrade] 未找到提交按钮:", submitButtonSelector);

        // 给页面一点时间提交/返回
        setTimeout(resolve, 1000);
      }, 500);
    });
  }
})();