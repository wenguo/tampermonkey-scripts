// ==UserScript==
// @name         DGUT 作业：预检勾选 + 批量跳转翻页下载（合并版-修复）
// @namespace    http://tampermonkey.net/
// @version      1.0.1
// @description  homeplan页扫描作业ID供勾选；开始后按队列逐个进入作业页，等待DOM渲染后自动翻页下载原始作业/附件
// @author       You
// @match        https://hw.dgut.edu.cn/teacher/course/*/homeworkPlan*
// @match        https://hw.dgut.edu.cn/teacher/homeworkPlan/*/homework*
// @grant        none
// ==/UserScript==

(function () {
  'use strict';

  // ----------------------------
  // 配置区
  // ----------------------------
  const STORE_KEY = 'DGUT_HW_DL_QUEUE_V1';
  const STORE_OPTS_KEY = 'DGUT_HW_DL_OPTS_V1';

  const TXT_DOWNLOAD_ORIGINAL = '下载原始作业';
  const TXT_DOWNLOAD_ATTACH = '下载作业附件';

  const WAIT_BETWEEN_ORIGINAL_CLICKS_MS = 2000;
  const WAIT_BETWEEN_ATTACH_CLICKS_MS = 9000;
  const WAIT_PAGE_LOAD_MS = 1200;

  // 等待 SPA 渲染：最多等多久（毫秒）
  const WAIT_DOM_TIMEOUT_MS = 45000;

  // ----------------------------
  // 工具函数
  // ----------------------------
  const logPrefix = '[DGUT-HW-DL]';

  function log(...args) {
    console.log(logPrefix, ...args);
  }

  function sleep(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  function nowStr() {
    const d = new Date();
    return d.toLocaleString();
  }

  function getQueue() {
    try {
      const raw = sessionStorage.getItem(STORE_KEY);
      if (!raw) return { list: [], cursor: 0, startedAt: null, lastAt: null };
      return JSON.parse(raw);
    } catch (e) {
      return { list: [], cursor: 0, startedAt: null, lastAt: null };
    }
  }

  function setQueue(q) {
    sessionStorage.setItem(STORE_KEY, JSON.stringify(q));
  }

  function clearQueue() {
    sessionStorage.removeItem(STORE_KEY);
  }

  function getOpts() {
    try {
      const raw = sessionStorage.getItem(STORE_OPTS_KEY);
      if (!raw) return { doOriginal: true, doAttach: false };
      return JSON.parse(raw);
    } catch (e) {
      return { doOriginal: true, doAttach: false };
    }
  }

  function setOpts(opts) {
    sessionStorage.setItem(STORE_OPTS_KEY, JSON.stringify(opts));
  }

  function isHomeplanPage() {
    return /\/teacher\/course\/.*\/homeworkPlan/.test(location.href);
  }

  function isHomeworkPage() {
    return /\/teacher\/homeworkPlan\/.*\/homework/.test(location.href);
  }

  function buildHomeworkUrl(id) {
    return `https://hw.dgut.edu.cn/teacher/homeworkPlan/${encodeURIComponent(id)}/homework`;
  }

  function extractIdFromHomeworkUrl() {
    const m = location.pathname.match(/\/teacher\/homeworkPlan\/([^/]+)\/homework/);
    return m ? decodeURIComponent(m[1]) : null;
  }

  function escapeHtml(s) {
    return String(s || '')
      .replaceAll('&', '&amp;')
      .replaceAll('<', '&lt;')
      .replaceAll('>', '&gt;')
      .replaceAll('"', '&quot;')
      .replaceAll("'", '&#039;');
  }

  async function waitFor(fn, timeoutMs, intervalMs = 500) {
    const start = Date.now();
    while (Date.now() - start < timeoutMs) {
      try {
        const v = fn();
        if (v) return v;
      } catch (e) {}
      await sleep(intervalMs);
    }
    return null;
  }

  // ----------------------------
  // UI：通用悬浮面板（稳定更新，不删DOM）
  // ----------------------------
  function ensurePanel(id, titleText) {
    let panel = document.getElementById(id);
    if (panel) return panel;

    panel = document.createElement('div');
    panel.id = id;
    panel.style.cssText = `
      position: fixed;
      top: 10px;
      right: 10px;
      z-index: 10000;
      width: 380px;
      max-height: 85vh;
      overflow: auto;
      background: #fff;
      border: 2px solid #19be6b;
      border-radius: 10px;
      padding: 12px;
      box-shadow: 0 6px 18px rgba(0,0,0,0.18);
      font-family: -apple-system,BlinkMacSystemFont,Segoe UI,Roboto,Helvetica,Arial,sans-serif;
      font-size: 13px;
      line-height: 1.3;
    `;

    panel.innerHTML = `
      <div style="font-weight:700;color:#19be6b;font-size:14px;margin-bottom:8px;">${escapeHtml(titleText)}</div>
      <div id="${id}__body"></div>
    `;

    document.body.appendChild(panel);
    return panel;
  }

  function panelBody(id) {
    return document.getElementById(`${id}__body`);
  }

  function mkBtn(text, color, onClick) {
    const b = document.createElement('button');
    b.textContent = text;
    b.style.cssText = `
      width: 100%;
      padding: 8px 10px;
      margin-top: 8px;
      cursor: pointer;
      border: none;
      border-radius: 6px;
      color: #fff;
      background: ${color};
      font-weight: 600;
    `;
    b.addEventListener('click', onClick);
    return b;
  }

  function mkSmallBtn(text, onClick) {
    const b = document.createElement('button');
    b.textContent = text;
    b.style.cssText = `
      padding: 4px 8px;
      margin-right: 6px;
      cursor: pointer;
      border: 1px solid #dcdee2;
      border-radius: 6px;
      background: #f8f8f9;
    `;
    b.addEventListener('click', onClick);
    return b;
  }

  // ----------------------------
  // Part A：homeplan 页面扫描 + 勾选 + 启动队列
  // ----------------------------
  function scanIDsFromHomeplan() {
    const rows = document.querySelectorAll('tr.ivu-table-row');
    const items = [];

    rows.forEach((row) => {
      const idCell = row.querySelector('td:first-child p');
      const titleCell = row.querySelector('td:nth-child(2) p');

      const hasViewBtn = Array.from(row.querySelectorAll('span')).some(
        (s) => (s.innerText || '').trim() === '查看作业'
      );

      if (!idCell) return;
      const rawId = (idCell.innerText || '').trim();
      const cleanId = rawId.replace(/[^\d]/g, '');
      if (!cleanId) return;

      const title = titleCell ? (titleCell.innerText || '').trim() : '未知标题';

      items.push({ id: cleanId, title, hasViewBtn });
    });

    const seen = new Set();
    const dedup = [];
    for (const it of items) {
      if (seen.has(it.id)) continue;
      seen.add(it.id);
      dedup.push(it);
    }
    return dedup;
  }

  function renderHomeplanUI() {
    ensurePanel('dgut_hw_panel_home', '作业预检 & 批量下载（合并版-修复）');
    const body = panelBody('dgut_hw_panel_home');

    body.innerHTML = `
      <div style="color:#515a6e;margin-bottom:6px;">
        流程：<b>扫描</b> → <b>勾选</b> → <b>开始下载</b>（将按作业逐个跳转并自动翻页下载）
      </div>

      <div style="border:1px dashed #dcdee2;border-radius:8px;padding:8px;margin:8px 0;">
        <div style="font-weight:700;margin-bottom:6px;">下载内容</div>
        <label style="display:block;margin:4px 0;">
          <input type="checkbox" id="dgut-opt-original"/> 原始作业
        </label>
        <label style="display:block;margin:4px 0;">
          <input type="checkbox" id="dgut-opt-attach"/> 作业附件
        </label>
        <div style="color:#808695;font-size:12px;margin-top:6px;">
          浏览器可能会拦截“多文件自动下载”，需要允许本网站自动下载。
        </div>
      </div>

      <div style="margin:6px 0;" id="dgut-home-topbtns"></div>
      <div id="dgut-home-results" style="margin-top:8px;"><div style="color:#808695;">等待扫描...</div></div>
      <div id="dgut-home-actions"></div>
    `;

    const opts = getOpts();
    document.getElementById('dgut-opt-original').checked = !!opts.doOriginal;
    document.getElementById('dgut-opt-attach').checked = !!opts.doAttach;

    const topBtns = document.getElementById('dgut-home-topbtns');
    const results = document.getElementById('dgut-home-results');
    const actions = document.getElementById('dgut-home-actions');

    let lastItems = [];

    function refreshList(items) {
      lastItems = items;

      if (!items.length) {
        results.innerHTML = '<div style="color:#ed4014;">未识别到作业 ID：请检查页面表格是否已加载完成。</div>';
        return;
      }

      const table = document.createElement('table');
      table.style.cssText = 'width:100%;border-collapse:collapse;margin-top:6px;';
      table.innerHTML = `
        <thead>
          <tr style="background:#f8f8f9;">
            <th style="border:1px solid #dcdee2;padding:4px;width:34px;">选</th>
            <th style="border:1px solid #dcdee2;padding:4px;width:88px;">ID</th>
            <th style="border:1px solid #dcdee2;padding:4px;">名称</th>
          </tr>
        </thead>
        <tbody></tbody>
      `;

      const tb = table.querySelector('tbody');
      items.forEach((it) => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
          <td style="border:1px solid #dcdee2;padding:4px;text-align:center;">
            <input type="checkbox" class="dgut-hw-check" data-id="${it.id}" checked />
          </td>
          <td style="border:1px solid #dcdee2;padding:4px;color:#2d8cf0;font-weight:700;">${it.id}</td>
          <td style="border:1px solid #dcdee2;padding:4px;">
            <div style="font-weight:600;">${escapeHtml(it.title)}</div>
            <div style="color:#808695;font-size:12px;margin-top:2px;">
              ${it.hasViewBtn ? '可查看' : '未检测到“查看作业”按钮（仍可尝试下载）'}
            </div>
          </td>
        `;
        tb.appendChild(tr);
      });

      results.innerHTML = `<div>共识别到 <b>${items.length}</b> 个作业：</div>`;
      results.appendChild(table);
    }

    topBtns.appendChild(
      mkSmallBtn('全选', () => results.querySelectorAll('.dgut-hw-check').forEach((c) => (c.checked = true)))
    );
    topBtns.appendChild(
      mkSmallBtn('全不选', () => results.querySelectorAll('.dgut-hw-check').forEach((c) => (c.checked = false)))
    );
    topBtns.appendChild(
      mkSmallBtn('反选', () => results.querySelectorAll('.dgut-hw-check').forEach((c) => (c.checked = !c.checked)))
    );

    actions.appendChild(
      mkBtn('1) 扫描作业 ID 列表', '#19be6b', async () => {
        results.innerHTML = '<div style="color:#808695;">扫描中...</div>';
        await sleep(500);
        const items = scanIDsFromHomeplan();
        refreshList(items);
      })
    );

    actions.appendChild(
      mkBtn('2) 开始下载（按勾选队列逐个处理）', '#2d8cf0', () => {
        const doOriginal = !!document.getElementById('dgut-opt-original')?.checked;
        const doAttach = !!document.getElementById('dgut-opt-attach')?.checked;
        if (!doOriginal && !doAttach) {
          alert('请至少勾选一种下载内容：原始作业 / 作业附件');
          return;
        }
        setOpts({ doOriginal, doAttach });

        if (!lastItems.length) {
          lastItems = scanIDsFromHomeplan();
          refreshList(lastItems);
        }

        const checkedIds = Array.from(results.querySelectorAll('.dgut-hw-check'))
          .filter((c) => c.checked)
          .map((c) => c.getAttribute('data-id'))
          .filter(Boolean);

        if (!checkedIds.length) {
          alert('没有勾选任何作业。');
          return;
        }

        const q = { list: checkedIds, cursor: 0, startedAt: nowStr(), lastAt: nowStr() };
        setQueue(q);

        const firstId = q.list[0];
        alert(`即将开始下载：共 ${q.list.length} 个作业。\n将跳转到作业页：${firstId}`);
        location.href = buildHomeworkUrl(firstId);
      })
    );

    actions.appendChild(
      mkBtn('清空队列 / 停止（仅清 sessionStorage）', '#ed4014', () => {
        clearQueue();
        sessionStorage.removeItem(STORE_OPTS_KEY);
        alert('已清空队列与选项。');
      })
    );

    setTimeout(() => {
      const items = scanIDsFromHomeplan();
      refreshList(items);
    }, 2000);
  }

  // ----------------------------
  // Part B：作业页：等待DOM -> 翻页 -> 下载 -> 跳下一个
  // ----------------------------

  function renderHomeworkUI(statusText) {
    ensurePanel('dgut_hw_panel_hw', '作业下载队列执行中…（修复版）');
    const body = panelBody('dgut_hw_panel_hw');

    const q = getQueue();
    const opts = getOpts();
    const currentId = extractIdFromHomeworkUrl();
    const total = q.list.length || 0;
    const idx = (q.cursor || 0) + 1;

    body.innerHTML = `
      <div style="border:1px solid #dcdee2;border-radius:8px;padding:8px;">
        <div><b>队列进度：</b>${total ? `${idx}/${total}` : '-'}</div>
        <div><b>当前作业ID：</b><span style="color:#2d8cf0;font-weight:800;">${escapeHtml(currentId || '未知')}</span></div>
        <div><b>下载内容：</b>${opts.doOriginal ? '原始作业 ' : ''}${opts.doAttach ? '附件' : ''}</div>
        <div style="margin-top:6px;color:#515a6e;"><b>状态：</b>${escapeHtml(statusText || '准备中...')}</div>
        <div style="margin-top:6px;color:#808695;font-size:12px;">
          startedAt=${escapeHtml(q.startedAt || '-')}, lastAt=${escapeHtml(q.lastAt || '-')}
        </div>

        <div style="margin-top:10px;">
          <button id="dgut-resume" style="padding:6px 10px;border:1px solid #dcdee2;border-radius:6px;background:#f8f8f9;cursor:pointer;">
            继续/重试当前作业
          </button>
          <button id="dgut-stop" style="padding:6px 10px;border:1px solid #ed4014;border-radius:6px;background:#fff;color:#ed4014;cursor:pointer;margin-left:6px;">
            清空队列
          </button>
        </div>

        <div style="margin-top:8px;color:#808695;font-size:12px;">
          如果浏览器拦截多文件下载，请允许本站点“自动下载多个文件”。
        </div>
      </div>
    `;

    document.getElementById('dgut-resume')?.addEventListener('click', () => runQueueOnHomeworkPage(true));
    document.getElementById('dgut-stop')?.addEventListener('click', () => {
      clearQueue();
      alert('已清空队列。');
    });
  }

  function findPagination() {
    return document.querySelector('#app ul.ivu-page');
  }

  function findAnyDownloadButton() {
    const all = Array.from(document.querySelectorAll('button'));
    return all.find((b) => {
      const s = (b.innerText || '').trim();
      return s.includes(TXT_DOWNLOAD_ORIGINAL) || s.includes(TXT_DOWNLOAD_ATTACH);
    });
  }

  async function waitHomeworkDomReady(updateStatus) {
    updateStatus('等待页面内容渲染（分页/下载按钮）...');
    const ok = await waitFor(() => {
      // 只要分页或下载按钮任意一个出现，就认为DOM可用（有些作业只有一页可能没分页）
      return findPagination() || findAnyDownloadButton();
    }, WAIT_DOM_TIMEOUT_MS, 600);

    if (!ok) {
      throw new Error('等待作业页面DOM超时：未检测到分页或下载按钮（可能未登录/页面结构变更/接口失败）。');
    }

    // 再额外等一下让表格行完整出来
    await sleep(800);
  }

  async function waitForPageLoad() {
    await sleep(WAIT_PAGE_LOAD_MS);
    const loader = document.querySelector('.ivu-spin.ivu-spin-large.ivu-spin-fix');
    if (!loader) return;
    // 等 spinner 消失
    await waitFor(() => loader.offsetParent === null, 20000, 400);
  }

  async function goToFirstPageIfPossible() {
    const firstPageButton = Array.from(document.querySelectorAll('#app li.ivu-page-item'))
      .find(item => item.textContent.trim() === '1' && !item.classList.contains('ivu-page-item-active'));

    if (firstPageButton) {
      firstPageButton.click();
      await waitForPageLoad();
      return true;
    }
    return false;
  }

  async function goToNextPage() {
    const nextButton = document.querySelector('#app li.ivu-page-next:not(.ivu-page-disabled)');
    if (nextButton) {
      nextButton.click();
      await waitForPageLoad();
      return true;
    }
    return false;
  }

  async function downloadCurrentPageOriginal() {
    const buttons = document.querySelectorAll('button.ivu-btn.ivu-btn-info.ivu-btn-small, button.ivu-btn.ivu-btn-info');
    for (const button of buttons) {
      const t = (button.innerText || '').trim();
      if (t.includes(TXT_DOWNLOAD_ORIGINAL)) {
        button.click();
        await sleep(WAIT_BETWEEN_ORIGINAL_CLICKS_MS);
      }
    }
  }

  async function downloadCurrentPageAttachments() {
    const buttons = document.querySelectorAll('button.ivu-btn.ivu-btn-default.ivu-btn-small, button.ivu-btn.ivu-btn-default');
    for (const button of buttons) {
      const t = (button.innerText || '').trim();
      if (t.includes(TXT_DOWNLOAD_ATTACH)) {
        button.click();
        await sleep(WAIT_BETWEEN_ATTACH_CLICKS_MS);
      }
    }
  }

  async function downloadAllPagesForThisHomework(opts, updateStatus) {
    // 如果有分页，先跳到第一页（确保不会从中间页开始漏）
    if (findPagination()) {
      updateStatus('检测到分页：切到第一页...');
      await goToFirstPageIfPossible();
    } else {
      updateStatus('未检测到分页：按单页处理...');
    }

    // 当前页下载
    updateStatus('下载当前页...');
    if (opts.doOriginal) await downloadCurrentPageOriginal();
    if (opts.doAttach) await downloadCurrentPageAttachments();

    // 翻页循环（若无分页，next 会直接 false）
    while (true) {
      if (!findPagination()) break;
      updateStatus('翻到下一页...');
      const hasNext = await goToNextPage();
      if (!hasNext) break;

      updateStatus('下载当前页...');
      if (opts.doOriginal) await downloadCurrentPageOriginal();
      if (opts.doAttach) await downloadCurrentPageAttachments();
    }

    updateStatus('本作业下载完成。');
  }

  async function runQueueOnHomeworkPage(isManual = false) {
    const q = getQueue();
    const opts = getOpts();

    // 没队列就提示
    if (!q.list || !q.list.length) {
      renderHomeworkUI('未检测到队列（请先在 homeplan 页扫描并开始下载）。');
      return;
    }

    if (typeof q.cursor !== 'number') q.cursor = 0;
    if (q.cursor < 0) q.cursor = 0;

    if (q.cursor >= q.list.length) {
      clearQueue();
      renderHomeworkUI('队列已完成（已自动清空）。');
      alert('全部作业已下载完成！');
      return;
    }

    const currentId = extractIdFromHomeworkUrl();
    const targetId = q.list[q.cursor];

    let status = `准备下载（目标ID=${targetId}，当前页ID=${currentId || '未知'}）`;
    renderHomeworkUI(status);

    const updateStatus = (s) => {
      status = s;
      renderHomeworkUI(status);
      log('STATUS:', s);
    };

    log('Queue:', q, 'Opts:', opts, 'CurrentId:', currentId, 'TargetId:', targetId, 'Manual:', isManual);

    // 如果当前页不是队列目标，直接跳过去
    if (currentId && targetId && currentId !== targetId) {
      updateStatus(`当前页ID(${currentId})≠目标(${targetId})，跳转到目标页...`);
      await sleep(600);
      location.href = buildHomeworkUrl(targetId);
      return;
    }

    try {
      // 关键修复：等待 SPA DOM 渲染
      await waitHomeworkDomReady(updateStatus);

      updateStatus('开始执行：自动翻页下载中...');
      await downloadAllPagesForThisHomework(opts, updateStatus);

      // 本作业完成 -> cursor+1 -> 跳下一个
      q.cursor += 1;
      q.lastAt = nowStr();
      setQueue(q);

      if (q.cursor >= q.list.length) {
        clearQueue();
        updateStatus('队列全部完成（已清空）。');
        alert('全部作业已下载完成！');
        return;
      }

      const nextId = q.list[q.cursor];
      updateStatus(`切换到下一个作业：${nextId}`);
      await sleep(1200);
      location.href = buildHomeworkUrl(nextId);
    } catch (err) {
      console.error(logPrefix, err);
      updateStatus(`发生错误：${String(err)}（请打开控制台查看详情）`);
      if (!isManual) {
        // 自动模式下不给弹窗刷屏；面板上有“继续/重试”
        log('Auto mode error; user can click Resume.');
      } else {
        alert('下载过程中出现错误，请打开控制台查看。');
      }
    }
  }

  // ----------------------------
  // 启动
  // ----------------------------
  if (isHomeplanPage()) {
    window.addEventListener('load', renderHomeplanUI);
  } else if (isHomeworkPage()) {
    window.addEventListener('load', () => {
      // SPA：load 之后再等一点
      setTimeout(() => runQueueOnHomeworkPage(false), 1500);
    });
  }
})();
