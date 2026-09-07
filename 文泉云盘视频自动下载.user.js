 ==UserScript==
 @name         文泉云盘视频自动下载
 @namespace    httpswww.wqyunpan.com
 @version      0.16.0
 @description  双层点击进入原生 Preview，由网站自行鉴权；直接把视频按页面可读文件名写入用户选择的目录，无需 UUID 映射和后续重命名。
 @match        httpswww.wqyunpan.combook-resources-list
 @match        httpswww.wqyunpan.compreview.html
 @grant        GM_getValue
 @grant        GM_setValue
 @grant        GM_deleteValue
 @grant        GM_xmlhttpRequest
 @grant        GM_addStyle
 @grant        unsafeWindow
 @connect      .wqketang.com
 @connect      vodbj.wqketang.com
 @run-at       document-idle
 ==UserScript==

(function () {
    'use strict';

     ============================================================
     版本
     ============================================================

    const SCRIPT_VERSION = '0.16.0';

    
      继续沿用原来的 Job Key。
     
      这样之前已经成功失败的任务状态仍可读取。
      如果你希望完全重新开始，点击“清除任务”即可。
     
    const JOB_KEY = 'wq-video-download-job-v014';

     ============================================================
     保存目录 Handle 使用 IndexedDB 持久化
     ============================================================

    const HANDLE_DB_NAME = 'wqyunpan-video-downloader';
    const HANDLE_DB_VERSION = 1;
    const HANDLE_STORE_NAME = 'handles';
    const DIRECTORY_HANDLE_KEY = 'video-save-directory';

     仅用于界面显示目录名称
    const DIRECTORY_NAME_KEY = 'wq-video-save-directory-name';

     ============================================================
     时序配置
     ============================================================

     第一次点击后等待详情展开
    const EXPAND_TIMEOUT = 15000;

     返回列表以后，等待当前资源真正出现在 DOM
    const PRIMARY_WAIT_TIMEOUT = 20000;
    const PRIMARY_RETRY_INTERVAL = 300;

     点击第二层以后，仍未进入 Preview，则重新尝试
    const PREVIEW_START_TIMEOUT = 7000;
    const MAX_PREVIEW_OPEN_RETRY = 3;

     Preview 页面等待真实媒体 URL
    const MEDIA_TIMEOUT = 45000;

     opening 状态总保护
    const OPEN_TIMEOUT = 120000;

     下载最长时间
    const DOWNLOAD_TIMEOUT = 30  60  1000;

     每个视频完成后等待 2 秒
    const POST_DOWNLOAD_DELAY = 2000;

     返回列表固定等待
    const LIST_SETTLE_DELAY = 3000;

     SPA 路由检测
    const ROUTE_INTERVAL = 700;

     列表任务 watchdog
    const WATCHDOG_INTERVAL = 2000;

     顶层资源找不到时整页刷新次数
    const MAX_LIST_REFRESH_RETRY = 1;

     ============================================================
     页面运行状态
     ============================================================

    let queueLoopRunning = false;
    let listSettling = false;
    let lastRoute = '';
    let previewWorkerRoute = '';
    let panelDelegationInstalled = false;

     ============================================================
     页面判断
     ============================================================

    function isListPage() {
        return location.pathname.includes('book-resources-list');
    }

    function isPreviewPage() {
        return location.pathname.endsWith('preview.html');
    }

     ============================================================
     基础工具
     ============================================================

    function sleep(ms) {
        return new Promise(resolve = setTimeout(resolve, ms));
    }

    function normalizeText(text) {
        return String(text  '')
            .replace(s+g, ' ')
            .trim();
    }

    
      Windows 文件名安全处理。
     
      保留中文、括号、数字、小数点等。
      只删除 Windows 不允许的字符。
     
    function safeFilename(name) {
        let result = String(name  'video.mp4')
            .replace([x00-x1F]g, '_')
            .replace([]g, '_')
            .replace(s+g, ' ')
            .trim();

         Windows 不允许文件名以空格句点结束
        result = result.replace([. ]+$g, '');

        if (!.(mp4webm)$i.test(result)) {
            result += '.mp4';
        }

        
          避免路径过长。
          对你的课程文件名来说 180 已经非常充足。
         
        if (result.length  180) {
            const ext =
                result.toLowerCase().endsWith('.webm')
                     '.webm'
                     '.mp4';

            result =
                result.slice(
                    0,
                    180 - ext.length
                ) +
                ext;
        }

        return result;
    }

    function formatBytes(value) {
        let n = Number(value);

        if (!Number.isFinite(n)) {
            return '-';
        }

        const units = [
            'B',
            'KB',
            'MB',
            'GB',
            'TB'
        ];

        let index = 0;

        while (
            n = 1024 &&
            index  units.length - 1
        ) {
            n = 1024;
            index++;
        }

        return (
            n.toFixed(
                index  1  0
            ) +
            ' ' +
            units[index]
        );
    }

     ============================================================
     Job
     ============================================================

    function getJob() {
        return GM_getValue(
            JOB_KEY,
            null
        );
    }

    function saveJob(job) {
        GM_setValue(
            JOB_KEY,
            job
        );
    }

    function clearJob() {
        GM_deleteValue(
            JOB_KEY
        );
    }

    
      兼容前面版本已有任务。
     
    function normalizeJob(job) {
        if (
            !job 
            !Array.isArray(job.items)
        ) {
            return job;
        }

        job.items.forEach(item = {
            item.name = safeFilename(
                item.name 
                item.desiredFilename 
                'video.mp4'
            );

            item.desiredFilename =
                safeFilename(
                    item.desiredFilename 
                    item.name
                );

            item.status =
                item.status 
                'pending';

            item.openedAt =
                Number(
                    item.openedAt  0
                );

            item.previewStartedAt =
                Number(
                    item.previewStartedAt  0
                );

            item.previewOpenRetry =
                Number(
                    item.previewOpenRetry  0
                );

            item.downloadStartedAt =
                Number(
                    item.downloadStartedAt  0
                );

            item.completedAt =
                Number(
                    item.completedAt  0
                );

            item.listRefreshRetry =
                Number(
                    item.listRefreshRetry  0
                );

            item.savedBytes =
                Number(
                    item.savedBytes  0
                );

            item.saveMode =
                item.saveMode  '';

            item.error =
                item.error  '';
        });

        if (
            !Number.isInteger(job.current) 
            job.current  0
        ) {
            job.current = 0;
        }

        return job;
    }

    function getNormalizedJob() {
        const job =
            normalizeJob(
                getJob()
            );

        if (job) {
            saveJob(job);
        }

        return job;
    }

    function updateCurrentItem(updater) {
        const job =
            getNormalizedJob();

        if (
            !job 
            job.current =
            job.items.length
        ) {
            return null;
        }

        const item =
            job.items[job.current];

        updater(
            item,
            job
        );

        saveJob(job);

        return item;
    }

     ============================================================
     IndexedDB：保存目录 Handle
     ============================================================

    function getIndexedDBFactory() {
        try {
            if (
                typeof indexedDB !==
                'undefined'
            ) {
                return indexedDB;
            }
        } catch (_) {
        }

        try {
            if (
                typeof unsafeWindow !==
                    'undefined' &&
                unsafeWindow.indexedDB
            ) {
                return unsafeWindow.indexedDB;
            }
        } catch (_) {
        }

        return null;
    }

    function openHandleDB() {
        return new Promise(
            (resolve, reject) = {

                const factory =
                    getIndexedDBFactory();

                if (!factory) {
                    reject(
                        new Error(
                            '浏览器不支持 IndexedDB'
                        )
                    );

                    return;
                }

                const request =
                    factory.open(
                        HANDLE_DB_NAME,
                        HANDLE_DB_VERSION
                    );

                request.onupgradeneeded =
                    event = {

                        const db =
                            event.target.result;

                        if (
                            !db.objectStoreNames
                                .contains(
                                    HANDLE_STORE_NAME
                                )
                        ) {
                            db.createObjectStore(
                                HANDLE_STORE_NAME
                            );
                        }
                    };

                request.onsuccess =
                    () = {
                        resolve(
                            request.result
                        );
                    };

                request.onerror =
                    () = {
                        reject(
                            request.error 
                            new Error(
                                '打开 IndexedDB 失败'
                            )
                        );
                    };
            }
        );
    }

    async function storeDirectoryHandle(
        handle
    ) {
        const db =
            await openHandleDB();

        try {
            await new Promise(
                (resolve, reject) = {

                    const tx =
                        db.transaction(
                            HANDLE_STORE_NAME,
                            'readwrite'
                        );

                    tx.objectStore(
                        HANDLE_STORE_NAME
                    ).put(
                        handle,
                        DIRECTORY_HANDLE_KEY
                    );

                    tx.oncomplete =
                        () = resolve();

                    tx.onerror =
                        () = reject(
                            tx.error 
                            new Error(
                                '保存目录 Handle 失败'
                            )
                        );

                    tx.onabort =
                        () = reject(
                            tx.error 
                            new Error(
                                '保存目录 Handle 被中止'
                            )
                        );
                }
            );

        } finally {
            try {
                db.close();
            } catch (_) {
            }
        }
    }

    async function loadDirectoryHandle() {
        const db =
            await openHandleDB();

        try {
            return await new Promise(
                (resolve, reject) = {

                    const tx =
                        db.transaction(
                            HANDLE_STORE_NAME,
                            'readonly'
                        );

                    const request =
                        tx.objectStore(
                            HANDLE_STORE_NAME
                        ).get(
                            DIRECTORY_HANDLE_KEY
                        );

                    request.onsuccess =
                        () = {
                            resolve(
                                request.result 
                                null
                            );
                        };

                    request.onerror =
                        () = {
                            reject(
                                request.error 
                                new Error(
                                    '读取保存目录失败'
                                )
                            );
                        };
                }
            );

        } finally {
            try {
                db.close();
            } catch (_) {
            }
        }
    }

    async function forgetDirectoryHandle() {
        const db =
            await openHandleDB();

        try {
            await new Promise(
                (resolve, reject) = {

                    const tx =
                        db.transaction(
                            HANDLE_STORE_NAME,
                            'readwrite'
                        );

                    tx.objectStore(
                        HANDLE_STORE_NAME
                    ).delete(
                        DIRECTORY_HANDLE_KEY
                    );

                    tx.oncomplete =
                        () = resolve();

                    tx.onerror =
                        () = reject(
                            tx.error
                        );
                }
            );

        } finally {
            try {
                db.close();
            } catch (_) {
            }
        }

        GM_deleteValue(
            DIRECTORY_NAME_KEY
        );
    }

     ============================================================
     保存目录权限
     ============================================================

    async function queryDirectoryPermission(
        handle
    ) {
        if (!handle) {
            return 'denied';
        }

        try {
            if (
                typeof handle.queryPermission ===
                'function'
            ) {
                return await handle.queryPermission({
                    mode 'readwrite'
                });
            }
        } catch (_) {
        }

        return 'prompt';
    }

    async function getGrantedDirectoryHandle() {
        let handle;

        try {
            handle =
                await loadDirectoryHandle();
        } catch (error) {
            console.warn(
                '[WQ] loadDirectoryHandle',
                error
            );

            return null;
        }

        if (!handle) {
            return null;
        }

        const permission =
            await queryDirectoryPermission(
                handle
            );

        if (
            permission ===
            'granted'
        ) {
            return handle;
        }

        return null;
    }

    
      必须由用户主动点击按钮触发。
     
      注意：这个函数调用 picker 之前不要 await 其他异步操作，
      否则浏览器可能认为已经丢失“用户手势”。
     
    async function chooseSaveDirectory() {
        let picker = null;
        let pickerThis = window;

        try {
            if (
                typeof window.showDirectoryPicker ===
                'function'
            ) {
                picker =
                    window.showDirectoryPicker;

                pickerThis =
                    window;
            }
        } catch (_) {
        }

        if (!picker) {
            try {
                if (
                    typeof unsafeWindow !==
                        'undefined' &&
                    typeof unsafeWindow
                        .showDirectoryPicker ===
                        'function'
                ) {
                    picker =
                        unsafeWindow
                            .showDirectoryPicker;

                    pickerThis =
                        unsafeWindow;
                }
            } catch (_) {
            }
        }

        if (!picker) {
            throw new Error(
                '当前浏览器不支持目录直接写入。请使用最新版 Chrome 或 Edge。'
            );
        }

        
          第一件事就是打开目录选择器。
         
        const handle =
            await picker.call(
                pickerThis,
                {
                    mode
                        'readwrite'
                }
            );

        let permission =
            await queryDirectoryPermission(
                handle
            );

        if (
            permission !==
                'granted' &&
            typeof handle.requestPermission ===
                'function'
        ) {
            permission =
                await handle.requestPermission({
                    mode
                        'readwrite'
                });
        }

        if (
            permission !==
            'granted'
        ) {
            throw new Error(
                '没有获得保存目录的写入权限。'
            );
        }

        await storeDirectoryHandle(
            handle
        );

        GM_setValue(
            DIRECTORY_NAME_KEY,
            handle.name  '已选择目录'
        );

        setListStatus(
            '保存目录已选择：' +
            (
                handle.name 
                '已授权目录'
            )
        );

        await updateDirectoryUI();

        return handle;
    }

    async function updateDirectoryUI() {
        const el =
            document.getElementById(
                'wq-dir-status'
            );

        if (!el) {
            return;
        }

        const rememberedName =
            GM_getValue(
                DIRECTORY_NAME_KEY,
                ''
            );

        let handle = null;

        try {
            handle =
                await loadDirectoryHandle();
        } catch (_) {
        }

        if (!handle) {
            el.textContent =
                '保存目录：未选择';

            return;
        }

        const permission =
            await queryDirectoryPermission(
                handle
            );

        const name =
            handle.name 
            rememberedName 
            '已选择目录';

        if (
            permission ===
            'granted'
        ) {
            el.textContent =
                `保存目录：${name}（已授权）`;

        } else {
            el.textContent =
                `保存目录：${name}（需要重新授权）`;
        }
    }

     ============================================================
     DOM 工具
     ============================================================

    function isVisible(element) {
        if (!element) {
            return false;
        }

        try {
            const style =
                getComputedStyle(
                    element
                );

            if (
                style.display ===
                    'none' 
                style.visibility ===
                    'hidden' 
                Number(style.opacity) ===
                    0
            ) {
                return false;
            }

            const rect =
                element
                    .getBoundingClientRect();

            return (
                rect.width  0 &&
                rect.height  0
            );

        } catch (_) {
            return false;
        }
    }

    function isProbablyClickable(
        element
    ) {
        if (!element) {
            return false;
        }

        const tag =
            element.tagName
                .toLowerCase();

        if (
            tag === 'a' 
            tag === 'button'
        ) {
            return true;
        }

        const role =
            element.getAttribute.(
                'role'
            );

        if (
            role === 'button' 
            role === 'link'
        ) {
            return true;
        }

        if (
            element.hasAttribute.(
                'onclick'
            ) 
            element.hasAttribute.(
                'tabindex'
            )
        ) {
            return true;
        }

        try {
            return (
                getComputedStyle(
                    element
                ).cursor ===
                'pointer'
            );

        } catch (_) {
            return false;
        }
    }

    function findClickableAncestor(
        element
    ) {
        let current =
            element;

        for (
            let i = 0;
            i  10 &&
            current;
            i++
        ) {
            if (
                isProbablyClickable(
                    current
                )
            ) {
                return current;
            }

            current =
                current.parentElement;
        }

        current =
            element;

        for (
            let i = 0;
            i  5 &&
            current;
            i++
        ) {
            const text =
                normalizeText(
                    current.textContent
                );

            if (
                text &&
                text.length  280
            ) {
                return current;
            }

            current =
                current.parentElement;
        }

        return element;
    }

    function smartClick(element) {
        if (!element) {
            throw new Error(
                '点击目标为空'
            );
        }

        try {
            element.scrollIntoView({
                behavior
                    'auto',

                block
                    'center'
            });
        } catch (_) {
        }

        try {
            element.focus.();
        } catch (_) {
        }

        try {
            element.click();
            return;

        } catch (_) {
        }

        element.dispatchEvent(
            new MouseEvent(
                'click',
                {
                    bubbles
                        true,

                    cancelable
                        true,

                    view
                        window
                }
            )
        );
    }

     ============================================================
     扫描 MP4 文件名
     ============================================================

    function extractMp4Name(text) {
        const value =
            normalizeText(
                text
            );

        if (!value) {
            return null;
        }

        if (
            .mp4$i.test(value) &&
            value.length  230
        ) {
            return value;
        }

        return null;
    }

    function scanVideoNames() {
        const names =
            new Set();

        if (!document.body) {
            return [];
        }

        const walker =
            document.createTreeWalker(
                document.body,
                NodeFilter.SHOW_TEXT
            );

        let node;

        while (
            (
                node =
                    walker.nextNode()
            )
        ) {
            const parent =
                node.parentElement;

            if (!parent) {
                continue;
            }

            if (
                parent.closest(
                    '#wq-auto-panel'
                )
            ) {
                continue;
            }

            const name =
                extractMp4Name(
                    node.nodeValue
                );

            if (name) {
                names.add(name);
            }
        }

        return [
            ...names
        ];
    }

    function findFilenameCandidates(
        filename
    ) {
        const target =
            normalizeText(
                filename
            );

        const result =
            [];

        const seen =
            new Set();

        if (!document.body) {
            return result;
        }

        const walker =
            document.createTreeWalker(
                document.body,
                NodeFilter.SHOW_TEXT
            );

        let node;

        while (
            (
                node =
                    walker.nextNode()
            )
        ) {
            const parent =
                node.parentElement;

            if (
                !parent 
                parent.closest(
                    '#wq-auto-panel'
                )
            ) {
                continue;
            }

            const text =
                normalizeText(
                    node.nodeValue
                );

            if (
                text !== target
            ) {
                continue;
            }

            const clickable =
                findClickableAncestor(
                    parent
                );

            if (
                !clickable 
                seen.has(clickable)
            ) {
                continue;
            }

            seen.add(
                clickable
            );

            result.push({
                textElement
                    parent,

                clickable,

                visible
                    isVisible(
                        clickable
                    )
            });
        }

        return result;
    }

     ============================================================
     第一层  第二层识别
     ============================================================

    function previewScore(element) {
        if (!element) {
            return 0;
        }

        let score =
            0;

        let current =
            element;

        for (
            let level = 0;
            level  6 &&
            current;
            level++
        ) {
            const cls =
                String(
                    current.className 
                    ''
                ).toLowerCase();

            const title =
                String(
                    current.getAttribute.(
                        'title'
                    ) 
                    ''
                ).toLowerCase();

            const aria =
                String(
                    current.getAttribute.(
                        'aria-label'
                    ) 
                    ''
                ).toLowerCase();

            const text =
                normalizeText(
                    current.textContent
                );

            if (
                previewplayvideoeye.test(
                    cls
                )
            ) {
                score += 50;
            }

            if (
                预览播放查看.test(
                    title
                )
            ) {
                score += 80;
            }

            if (
                预览播放查看.test(
                    aria
                )
            ) {
                score += 80;
            }

            if (
                预览播放.test(
                    text
                )
            ) {
                score += 25;
            }

            if (
                current.querySelector.(
                    [
                        'svg',
                        'i',
                        '[class=icon]',
                        '[class=preview]',
                        '[class=play]',
                        '[class=eye]'
                    ].join(',')
                )
            ) {
                score += 15;
            }

            current =
                current.parentElement;
        }

        return score;
    }

    function findPrimaryItem(
        filename
    ) {
        const list =
            findFilenameCandidates(
                filename
            )
                .filter(
                    item =
                        item.visible
                );

        if (!list.length) {
            return null;
        }

        
          第一层通常 Preview 特征最低。
         
        list.sort(
            (a, b) =
                previewScore(
                    a.clickable
                ) -
                previewScore(
                    b.clickable
                )
        );

        return list[0];
    }

    async function waitForPrimaryItem(
        filename
    ) {
        const started =
            Date.now();

        while (
            Date.now() -
            started 
            PRIMARY_WAIT_TIMEOUT
        ) {
            const item =
                findPrimaryItem(
                    filename
                );

            if (item) {
                return item;
            }

            const elapsed =
                Math.floor(
                    (
                        Date.now() -
                        started
                    ) 
                    1000
                );

            setListStatus(
                `等待列表恢复 ${elapsed}s：${filename}`
            );

            await sleep(
                PRIMARY_RETRY_INTERVAL
            );
        }

        return null;
    }

    function findAlreadyVisibleDetail(
        filename
    ) {
        const list =
            findFilenameCandidates(
                filename
            )
                .filter(
                    item =
                        item.visible
                );

        if (
            list.length  2
        ) {
            return null;
        }

        const scored =
            list
                .map(
                    item = ({
                        ...item,

                        score
                            previewScore(
                                item.clickable
                            )
                    })
                )
                .sort(
                    (a, b) =
                        b.score -
                        a.score
                );

        if (
            scored[0].score 
            scored[
                scored.length - 1
            ].score
        ) {
            return scored[0];
        }

        return null;
    }

    function snapshotCandidates(
        filename
    ) {
        return new Set(
            findFilenameCandidates(
                filename
            )
                .filter(
                    item =
                        item.visible
                )
                .map(
                    item =
                        item.clickable
                )
        );
    }

    async function waitForDetailItem(
        filename,
        beforeSet,
        primaryElement
    ) {
        const started =
            Date.now();

        while (
            Date.now() -
            started 
            EXPAND_TIMEOUT
        ) {
            const list =
                findFilenameCandidates(
                    filename
                )
                    .filter(
                        item =
                            item.visible
                    );

            const candidates =
                list.filter(
                    item = {

                        if (
                            item.clickable ===
                            primaryElement
                        ) {
                            return false;
                        }

                        if (
                            !beforeSet.has(
                                item.clickable
                            )
                        ) {
                            return true;
                        }

                        return (
                            previewScore(
                                item.clickable
                            )  0
                        );
                    }
                );

            if (
                candidates.length
            ) {
                candidates.sort(
                    (a, b) =
                        previewScore(
                            b.clickable
                        ) -
                        previewScore(
                            a.clickable
                        )
                );

                return candidates[0];
            }

            await sleep(250);
        }

        return null;
    }

     ============================================================
     Preview 入口
     ============================================================

    function findPreviewHref(element) {
        let current =
            element;

        for (
            let level = 0;
            level  7 &&
            current;
            level++
        ) {
            if (
                current.tagName
                    .toLowerCase() ===
                'a'
            ) {
                const href =
                    current.getAttribute(
                        'href'
                    );

                if (href) {
                    try {
                        const url =
                            new URL(
                                href,
                                location.href
                            );

                        if (
                            url.pathname
                                .endsWith(
                                    'preview.html'
                                )
                        ) {
                            return url.href;
                        }

                    } catch (_) {
                    }
                }
            }

            const nested =
                current.querySelector.(
                    'a[href=preview.html]'
                );

            if (nested) {
                try {
                    return new URL(
                        nested.getAttribute(
                            'href'
                        ),
                        location.href
                    ).href;

                } catch (_) {
                }
            }

            current =
                current.parentElement;
        }

        return null;
    }

    async function enterPreview(
        detail,
        filename
    ) {
        const href =
            findPreviewHref(
                detail.clickable
            );

        if (href) {
            setListStatus(
                '② 进入 Preview：' +
                filename
            );

            location.href =
                href;

            return;
        }

        setListStatus(
            '② 点击 Preview：' +
            filename
        );

        smartClick(
            detail.clickable
        );
    }

    async function openResource(
        filename
    ) {
        
          如果当前项目已经展开，
          直接进入第二层。
         
        const existingDetail =
            findAlreadyVisibleDetail(
                filename
            );

        if (existingDetail) {
            setListStatus(
                '详情已展开：' +
                filename
            );

            await enterPreview(
                existingDetail,
                filename
            );

            return;
        }

        const primary =
            await waitForPrimaryItem(
                filename
            );

        if (!primary) {
            const error =
                new Error(
                    '等待20秒后仍找不到顶层资源项'
                );

            error.code =
                'PRIMARY_NOT_FOUND';

            throw error;
        }

        const before =
            snapshotCandidates(
                filename
            );

        setListStatus(
            '① 展开：' +
            filename
        );

        smartClick(
            primary.clickable
        );

        setListStatus(
            '等待详情：' +
            filename
        );

        const detail =
            await waitForDetailItem(
                filename,
                before,
                primary.clickable
            );

        if (!detail) {
            const error =
                new Error(
                    '详情展开后未找到第二个 MP4 项'
                );

            error.code =
                'DETAIL_NOT_FOUND';

            throw error;
        }

        await enterPreview(
            detail,
            filename
        );
    }

     ============================================================
     UI
     ============================================================

    function setListStatus(text) {
        const el =
            document.getElementById(
                'wq-auto-status'
            );

        if (el) {
            el.textContent =
                text;
        }
    }

    function updateListUI() {
        const count =
            document.getElementById(
                'wq-auto-count'
            );

        if (count) {
            count.textContent =
                `当前 DOM：${scanVideoNames().length} 个 MP4`;
        }

        const progress =
            document.getElementById(
                'wq-auto-progress'
            );

        if (!progress) {
            updateDirectoryUI();
            return;
        }

        const job =
            getNormalizedJob();

        if (!job) {
            progress.textContent =
                '无活动任务';

            updateDirectoryUI();
            return;
        }

        const success =
            job.items.filter(
                item =
                    item.status ===
                    'done'
            ).length;

        const failed =
            job.items.filter(
                item =
                    item.status ===
                    'failed'
            ).length;

        const pending =
            job.items.filter(
                item =
                    item.status ===
                    'pending'
            ).length;

        progress.textContent =
            `进度 ${Math.min(
                job.current + 1,
                job.items.length
            )}${job.items.length}` +
            ` · 成功 ${success}` +
            ` · 失败 ${failed}` +
            ` · 待处理 ${pending}`;

        updateDirectoryUI();
    }

     ============================================================
     面板事件
     ============================================================

    function installPanelDelegation() {
        if (
            panelDelegationInstalled
        ) {
            return;
        }

        panelDelegationInstalled =
            true;

        document.addEventListener(
            'click',
            async event = {

                const button =
                    event.target
                        .closest.(
                            '#wq-auto-panel [data-action]'
                        );

                if (!button) {
                    return;
                }

                event.preventDefault();
                event.stopPropagation();

                const action =
                    button.dataset.action;

                try {
                    switch (action) {
                        case 'choose-dir'
                            
                              这一调用必须直接处于 click 事件中。
                             
                            await chooseSaveDirectory();
                            break;

                        case 'start'
                            await startQueue();
                            break;

                        case 'retry-failed'
                            retryFailedItems();
                            break;

                        case 'resume'
                            resumeQueue();
                            break;

                        case 'stop'
                            stopQueue();
                            break;

                        case 'reset'
                            resetQueue();
                            break;

                        case 'forget-dir'
                            await forgetDirectoryHandle();

                            setListStatus(
                                '已清除保存目录授权'
                            );

                            await updateDirectoryUI();
                            break;
                    }

                } catch (error) {
                    console.error(
                        '[WQ] panel action',
                        action,
                        error
                    );

                    setListStatus(
                        '操作失败：' +
                        (
                            error.message 
                            error
                        )
                    );

                    alert(
                        '操作失败：n' +
                        (
                            error.message 
                            error
                        )
                    );
                }
            },
            true
        );
    }

    function installListPanel() {
        installPanelDelegation();

        const old =
            document.getElementById(
                'wq-auto-panel'
            );

        if (
            old &&
            old.dataset.version ===
            SCRIPT_VERSION
        ) {
            return;
        }

        if (old) {
            old.remove();
        }

        GM_addStyle(`
#wq-auto-panel {
    position fixed;
    right 18px;
    bottom 18px;
    z-index 2147483646;
    width 530px;
    padding 12px;
    border-radius 10px;
    background rgba(255,255,255,.98);
    color #222;
    box-shadow 0 6px 26px rgba(0,0,0,.28);
    font 13px -apple-system, BlinkMacSystemFont,
        Segoe UI, Microsoft YaHei, sans-serif;
}

#wq-auto-panel strong {
    display block;
    margin-bottom 7px;
    font-size 15px;
}

#wq-auto-count,
#wq-auto-progress,
#wq-auto-status,
#wq-dir-status {
    display block;
    margin-top 6px;
    color #666;
    line-height 1.55;
}

#wq-dir-status {
    color #1677ff;
}

#wq-auto-status {
    min-height 20px;
}

.wq-auto-btn {
    margin-top 10px;
    margin-right 5px;
    padding 6px 9px;
    border 1px solid #bbb;
    border-radius 5px;
    background white;
    cursor pointer;
}

.wq-auto-btnhover {
    background #f5f5f5;
}

.wq-auto-primary {
    background #1677ff;
    color white;
    border-color #1677ff;
}

.wq-auto-primaryhover {
    background #4096ff;
}

.wq-auto-dir {
    background #52c41a;
    color white;
    border-color #52c41a;
}

.wq-auto-dirhover {
    background #73d13d;
}
        `);

        const panel =
            document.createElement(
                'div'
            );

        panel.id =
            'wq-auto-panel';

        panel.dataset.version =
            SCRIPT_VERSION;

        panel.innerHTML = `
strong
    文泉视频自动下载 v${SCRIPT_VERSION}
strong

span id=wq-dir-status
    保存目录：检查中…
span

span id=wq-auto-count
    扫描中…
span

span id=wq-auto-progressspan

span id=wq-auto-status
    请先选择保存目录
span

button
    class=wq-auto-btn wq-auto-dir
    data-action=choose-dir

    ① 选择保存目录
button

button
    class=wq-auto-btn wq-auto-primary
    data-action=start

    ② 新建任务
button

br

button
    class=wq-auto-btn
    data-action=retry-failed

    仅重试失败项
button

button
    class=wq-auto-btn
    data-action=resume

    继续
button

button
    class=wq-auto-btn
    data-action=stop

    停止
button

button
    class=wq-auto-btn
    data-action=reset

    清除任务
button

button
    class=wq-auto-btn
    data-action=forget-dir

    清除目录授权
button
        `;

        document.body.appendChild(
            panel
        );

        updateListUI();
    }

     ============================================================
     新建任务
     ============================================================

    async function startQueue() {
        const dir =
            await getGrantedDirectoryHandle();

        if (!dir) {
            alert(
                '请先点击“① 选择保存目录”，选择视频保存文件夹。'
            );

            return;
        }

        const names =
            scanVideoNames();

        if (!names.length) {
            alert(
                '当前页面暂时没有识别到 MP4。n' +
                '请等资源列表完全显示以后再试。'
            );

            return;
        }

        if (
            !confirm(
                `识别到 ${names.length} 个视频。nn` +
                `保存目录：${dir.name}nn` +
                `视频将直接以页面中的原始可读文件名保存，` +
                `不再生成 UUID 映射，也不再需要 PowerShell 重命名。nn` +
                `注意：如果目标目录中已经存在同名文件，本版会覆盖它。nn` +
                `是否开始？`
            )
        ) {
            return;
        }

        const job = {
            version
                SCRIPT_VERSION,

            running
                true,

            createdAt
                Date.now(),

            listUrl
                location.href,

            current
                0,

            items
                names.map(
                    name = ({
                        name
                            safeFilename(
                                name
                            ),

                        desiredFilename
                            safeFilename(
                                name
                            ),

                        status
                            'pending',

                        openedAt
                            0,

                        previewStartedAt
                            0,

                        previewOpenRetry
                            0,

                        downloadStartedAt
                            0,

                        completedAt
                            0,

                        listRefreshRetry
                            0,

                        savedBytes
                            0,

                        saveMode
                            '',

                        error
                            ''
                    })
                )
        };

        saveJob(job);

        updateListUI();

        setListStatus(
            '任务已创建，开始处理…'
        );

        processQueue()
            .catch(
                console.error
            );
    }

     ============================================================
     停止  继续  重试
     ============================================================

    function stopQueue() {
        const job =
            getNormalizedJob();

        if (!job) {
            return;
        }

        job.running =
            false;

        saveJob(job);

        setListStatus(
            '任务已停止'
        );

        updateListUI();
    }

    function resumeQueue() {
        const job =
            getNormalizedJob();

        if (!job) {
            alert(
                '没有可以继续的任务。'
            );

            return;
        }

        if (
            job.current 
            job.items.length
        ) {
            const current =
                job.items[
                    job.current
                ];

            if (
                current.status ===
                    'opening' 
                current.status ===
                    'downloading'
            ) {
                current.status =
                    'pending';

                current.openedAt =
                    0;

                current.previewStartedAt =
                    0;

                current.downloadStartedAt =
                    0;

                current.error =
                    '';
            }
        }

        job.running =
            true;

        saveJob(job);

        setListStatus(
            '继续任务，重新处理当前资源…'
        );

        updateListUI();

        processQueue()
            .catch(
                console.error
            );
    }

    function retryFailedItems() {
        const job =
            getNormalizedJob();

        if (
            !job 
            !Array.isArray(
                job.items
            )
        ) {
            alert(
                '没有现有任务。'
            );

            return;
        }

        let count =
            0;

        job.items.forEach(
            item = {

                if (
                    item.status ===
                    'failed'
                ) {
                    item.status =
                        'pending';

                    item.error =
                        '';

                    item.openedAt =
                        0;

                    item.previewStartedAt =
                        0;

                    item.previewOpenRetry =
                        0;

                    item.downloadStartedAt =
                        0;

                    item.completedAt =
                        0;

                    item.listRefreshRetry =
                        0;

                    item.savedBytes =
                        0;

                    item.saveMode =
                        '';

                    count++;
                }
            }
        );

        if (!count) {
            alert(
                '当前任务没有失败项。'
            );

            return;
        }

        const firstPending =
            job.items.findIndex(
                item =
                    item.status ===
                    'pending'
            );

        if (
            firstPending  0
        ) {
            return;
        }

        job.current =
            firstPending;

        job.running =
            true;

        saveJob(job);

        setListStatus(
            `准备重试 ${count} 个失败项`
        );

        updateListUI();

        processQueue()
            .catch(
                console.error
            );
    }

    function resetQueue() {
        if (
            !confirm(
                '确定清除当前下载任务记录吗？n' +
                '已经写入磁盘的视频不会删除。'
            )
        ) {
            return;
        }

        clearJob();

        setListStatus(
            '任务已清除'
        );

        updateListUI();
    }

     ============================================================
     队列状态机
     ============================================================

    async function processQueue() {
        if (
            queueLoopRunning 
            !isListPage()
        ) {
            return;
        }

        queueLoopRunning =
            true;

        console.log(
            '[WQ] processQueue start'
        );

        try {
            while (true) {
                let job =
                    getNormalizedJob();

                if (
                    !job 
                    !job.running
                ) {
                    break;
                }

                 ------------------------------------------------
                 跳过 done  failed
                 ------------------------------------------------

                while (
                    job.current 
                    job.items.length &&
                    (
                        job.items[
                            job.current
                        ].status ===
                            'done' 
                        job.items[
                            job.current
                        ].status ===
                            'failed'
                    )
                ) {
                    job.current++;

                    saveJob(job);
                }

                 ------------------------------------------------
                 全部结束
                 ------------------------------------------------

                if (
                    job.current =
                    job.items.length
                ) {
                    job.running =
                        false;

                    job.finishedAt =
                        Date.now();

                    saveJob(job);

                    const success =
                        job.items.filter(
                            item =
                                item.status ===
                                'done'
                        ).length;

                    const failed =
                        job.items.filter(
                            item =
                                item.status ===
                                'failed'
                        ).length;

                    setListStatus(
                        `全部完成：成功 ${success}，失败 ${failed}`
                    );

                    updateListUI();

                    break;
                }

                const item =
                    job.items[
                        job.current
                    ];

                updateListUI();

                 =================================================
                 opening
                 =================================================

                if (
                    item.status ===
                    'opening'
                ) {
                    const elapsed =
                        Date.now() -
                        (
                            item.openedAt 
                            Date.now()
                        );

                    
                      正常情况下 Preview 一旦打开，
                      Preview 页面会把状态改为 downloading。
                     
                    if (
                        item.previewStartedAt
                    ) {
                        await sleep(
                            800
                        );

                        continue;
                    }

                    
                      仍然在列表页且超过 7 秒：
                      说明第二层点击没有真正打开 Preview。
                     
                    if (
                        isListPage() &&
                        elapsed 
                        PREVIEW_START_TIMEOUT
                    ) {
                        const retry =
                            item.previewOpenRetry 
                            0;

                        if (
                            retry 
                            MAX_PREVIEW_OPEN_RETRY
                        ) {
                            item.previewOpenRetry =
                                retry + 1;

                            item.status =
                                'pending';

                            item.openedAt =
                                0;

                            item.previewStartedAt =
                                0;

                            item.error =
                                '';

                            saveJob(job);

                            setListStatus(
                                `Preview 未打开，重新点击 ` +
                                `${item.previewOpenRetry}${MAX_PREVIEW_OPEN_RETRY}：` +
                                item.name
                            );

                            await sleep(
                                800
                            );

                            continue;
                        }

                        item.status =
                            'failed';

                        item.error =
                            `连续 ${MAX_PREVIEW_OPEN_RETRY} 次点击后仍未进入 Preview`;

                        saveJob(job);

                        continue;
                    }

                    if (
                        elapsed 
                        OPEN_TIMEOUT
                    ) {
                        item.status =
                            'failed';

                        item.error =
                            '等待 Preview 页面超时';

                        saveJob(job);

                        continue;
                    }

                    setListStatus(
                        `等待 Preview ${Math.floor(
                            elapsed  1000
                        )}s：${item.name}`
                    );

                    await sleep(
                        800
                    );

                    continue;
                }

                 =================================================
                 downloading
                 =================================================

                if (
                    item.status ===
                    'downloading'
                ) {
                    if (
                        item.downloadStartedAt &&
                        Date.now() -
                        item.downloadStartedAt 
                        DOWNLOAD_TIMEOUT
                    ) {
                        item.status =
                            'failed';

                        item.error =
                            '下载超时';

                        saveJob(job);

                        continue;
                    }

                    setListStatus(
                        '正在保存：' +
                        item.name
                    );

                    await sleep(
                        1200
                    );

                    continue;
                }

                 =================================================
                 pending
                 =================================================

                item.status =
                    'opening';

                item.openedAt =
                    Date.now();

                item.previewStartedAt =
                    0;

                item.error =
                    '';

                saveJob(job);

                setListStatus(
                    `准备 ${job.current + 1}${job.items.length}：${item.name}`
                );

                try {
                    await openResource(
                        item.name
                    );

                } catch (error) {
                    console.error(
                        '[WQ] openResource',
                        error
                    );

                    job =
                        getNormalizedJob();

                    if (
                        !job 
                        job.current =
                        job.items.length
                    ) {
                        break;
                    }

                    const current =
                        job.items[
                            job.current
                        ];

                    
                      顶层或详情偶发加载失败：
                      整个列表刷新一次再试。
                     
                    if (
                        (
                            error.code ===
                                'PRIMARY_NOT_FOUND' 
                            error.code ===
                                'DETAIL_NOT_FOUND'
                        ) &&
                        (
                            current.listRefreshRetry 
                            0
                        ) 
                        MAX_LIST_REFRESH_RETRY
                    ) {
                        current.listRefreshRetry =
                            (
                                current.listRefreshRetry 
                                0
                            ) +
                            1;

                        current.status =
                            'pending';

                        current.openedAt =
                            0;

                        current.previewStartedAt =
                            0;

                        current.error =
                            '页面资源未恢复，自动刷新重试';

                        saveJob(job);

                        setListStatus(
                            '页面状态异常，正在刷新后重试：' +
                            current.name
                        );

                        await sleep(
                            800
                        );

                        location.reload();

                        return;
                    }

                    current.status =
                        'failed';

                    current.error =
                        error.message;

                    saveJob(job);

                    setListStatus(
                        '失败：' +
                        current.name +
                        '  ' +
                        error.message
                    );
                }

                await sleep(
                    1000
                );
            }

        } finally {
            queueLoopRunning =
                false;

            console.log(
                '[WQ] processQueue stop'
            );
        }
    }

     ============================================================
     Preview 媒体 URL 识别
     ============================================================

    function mediaType(url) {
        const value =
            String(
                url  ''
            );

        if (
            .mp4([#]$)i.test(
                value
            )
        ) {
            return 'mp4';
        }

        if (
            .webm([#]$)i.test(
                value
            )
        ) {
            return 'webm';
        }

        return null;
    }

    function mediaScore(url) {
        const value =
            String(
                url  ''
            );

        let score =
            0;

        if (
            .mp4([#]$)i.test(
                value
            )
        ) {
            score += 100;
        }

        if (
            wqketang.comi.test(
                value
            )
        ) {
            score += 120;
        }

        if (
            auth_key=i.test(
                value
            )
        ) {
            score += 120;
        }

        if (
            -sd-i.test(
                value
            )
        ) {
            score += 10;
        }

        return score;
    }

    function collectVideoUrls() {
        const urls =
            [];

        document
            .querySelectorAll(
                'video'
            )
            .forEach(
                video = {

                    if (
                        video.currentSrc
                    ) {
                        urls.push(
                            video.currentSrc
                        );
                    }

                    if (
                        video.src
                    ) {
                        urls.push(
                            video.src
                        );
                    }

                    video
                        .querySelectorAll(
                            'source'
                        )
                        .forEach(
                            source = {

                                if (
                                    source.src
                                ) {
                                    urls.push(
                                        source.src
                                    );
                                }
                            }
                        );
                }
            );

        return urls;
    }

    function collectPerformanceUrls() {
        const urls =
            [];

        try {
            performance
                .getEntriesByType(
                    'resource'
                )
                .forEach(
                    entry = {

                        if (
                            mediaType(
                                entry.name
                            )
                        ) {
                            urls.push(
                                entry.name
                            );
                        }
                    }
                );

        } catch (_) {
        }

        return urls;
    }

    function findBestMedia() {
        const urls =
            [
                ...new Set([
                    ...collectVideoUrls(),
                    ...collectPerformanceUrls()
                ])
            ];

        const candidates =
            urls
                .filter(
                    url =
                        Boolean(
                            mediaType(
                                url
                            )
                        )
                )
                .sort(
                    (a, b) =
                        mediaScore(b) -
                        mediaScore(a)
                );

        if (!candidates.length) {
            return null;
        }

        return {
            url
                candidates[0],

            type
                mediaType(
                    candidates[0]
                )
        };
    }

     ============================================================
     Preview 状态提示
     ============================================================

    function showWorkerStatus(text) {
        let el =
            document.getElementById(
                'wq-worker-status'
            );

        if (!el) {
            el =
                document.createElement(
                    'div'
                );

            el.id =
                'wq-worker-status';

            el.style.cssText = `
                position fixed;
                right 16px;
                bottom 16px;
                z-index 2147483647;
                max-width 620px;
                padding 11px 15px;
                border-radius 8px;
                background rgba(0,0,0,.84);
                color white;
                font-size 14px;
                line-height 1.6;
            `;

            document.body.appendChild(
                el
            );
        }

        el.textContent =
            text;
    }

     ============================================================
     浏览器 fetch 流式写入
    
     成功时优点：
     视频不需要完整放进内存。
     ============================================================

    async function streamFetchToFile(
        url,
        fileHandle,
        filename
    ) {
        const controller =
            new AbortController();

        const timer =
            setTimeout(
                () = {
                    controller.abort();
                },
                DOWNLOAD_TIMEOUT
            );

        let writable = null;

        try {
            const response =
                await fetch(
                    url,
                    {
                        method
                            'GET',

                        cache
                            'no-store',

                        credentials
                            'omit',

                        signal
                            controller.signal
                    }
                );

            if (!response.ok) {
                throw new Error(
                    `HTTP ${response.status}`
                );
            }

            const total =
                Number(
                    response.headers.get(
                        'content-length'
                    )  0
                );

            writable =
                await fileHandle.createWritable();

            let loaded =
                0;

            if (
                response.body &&
                typeof response.body.getReader ===
                    'function'
            ) {
                const reader =
                    response.body.getReader();

                while (true) {
                    const {
                        done,
                        value
                    } =
                        await reader.read();

                    if (done) {
                        break;
                    }

                    if (value) {
                        await writable.write(
                            value
                        );

                        loaded +=
                            value.byteLength;

                        if (total  0) {
                            const percent =
                                Math.floor(
                                    loaded 
                                    100 
                                    total
                                );

                            showWorkerStatus(
                                `直接写入 ${percent}%：${filename}`
                            );

                        } else {
                            showWorkerStatus(
                                `已写入 ${formatBytes(
                                    loaded
                                )}：${filename}`
                            );
                        }
                    }
                }

            } else {
                const blob =
                    await response.blob();

                await writable.write(
                    blob
                );

                loaded =
                    blob.size;
            }

            await writable.close();

            writable =
                null;

            return {
                mode
                    'fetch-stream',

                bytes
                    loaded
            };

        } catch (error) {
            if (writable) {
                try {
                    if (
                        typeof writable.abort ===
                        'function'
                    ) {
                        await writable.abort();
                    }
                } catch (_) {
                }
            }

            throw error;

        } finally {
            clearTimeout(
                timer
            );
        }
    }

     ============================================================
     GM_xmlhttpRequest Blob 备用下载
    
     浏览器 fetch 如果因为 CDN CORS 失败，
     自动退到 Tampermonkey 后台下载。
     ============================================================

    function gmFetchBlob(
        url,
        filename
    ) {
        return new Promise(
            (
                resolve,
                reject
            ) = {

                let lastPercent =
                    -1;

                GM_xmlhttpRequest({
                    method
                        'GET',

                    url,

                    responseType
                        'blob',

                    timeout
                        DOWNLOAD_TIMEOUT,

                    onprogress
                        event = {

                            try {
                                if (
                                    event.total  0
                                ) {
                                    const percent =
                                        Math.floor(
                                            event.loaded 
                                            100 
                                            event.total
                                        );

                                    if (
                                        percent !==
                                        lastPercent
                                    ) {
                                        lastPercent =
                                            percent;

                                        showWorkerStatus(
                                            `备用下载 ${percent}%：${filename}`
                                        );
                                    }

                                } else if (
                                    event.loaded  0
                                ) {
                                    showWorkerStatus(
                                        `备用下载 ${formatBytes(
                                            event.loaded
                                        )}：${filename}`
                                    );
                                }
                            } catch (_) {
                            }
                        },

                    onload
                        response = {

                            if (
                                response.status  200 
                                response.status = 300
                            ) {
                                reject(
                                    new Error(
                                        `HTTP ${response.status}`
                                    )
                                );

                                return;
                            }

                            const blob =
                                response.response;

                            if (
                                !blob 
                                blob.size  1024
                            ) {
                                reject(
                                    new Error(
                                        '返回的视频数据为空或异常'
                                    )
                                );

                                return;
                            }

                            resolve(
                                blob
                            );
                        },

                    onerror
                        error = {

                            reject(
                                new Error(
                                    error.error 
                                    error.message 
                                    'GM_xmlhttpRequest 下载失败'
                                )
                            );
                        },

                    ontimeout
                        () = {

                            reject(
                                new Error(
                                    '备用下载超时'
                                )
                            );
                        }
                });
            }
        );
    }

    async function writeBlobToFile(
        blob,
        fileHandle,
        filename
    ) {
        let writable =
            null;

        try {
            writable =
                await fileHandle
                    .createWritable();

            showWorkerStatus(
                `正在写入磁盘：${filename}`
            );

            await writable.write(
                blob
            );

            await writable.close();

            writable =
                null;

            return {
                mode
                    'GM_xmlhttpRequest-blob',

                bytes
                    blob.size
            };

        } catch (error) {
            if (writable) {
                try {
                    if (
                        typeof writable.abort ===
                        'function'
                    ) {
                        await writable.abort();
                    }
                } catch (_) {
                }
            }

            throw error;
        }
    }

     ============================================================
     直接保存视频
     ============================================================

    async function saveVideoDirect(
        url,
        filename
    ) {
        filename =
            safeFilename(
                filename
            );

        const dir =
            await getGrantedDirectoryHandle();

        if (!dir) {
            const error =
                new Error(
                    '保存目录权限失效，请返回列表重新点击“选择保存目录”。'
                );

            error.code =
                'DIRECTORY_PERMISSION_REQUIRED';

            throw error;
        }

        
          createtrue：
          不存在则创建。
         
          createWritable() 默认会覆盖原文件。
         
        const fileHandle =
            await dir.getFileHandle(
                filename,
                {
                    create
                        true
                }
            );

         --------------------------------------------------------
         第一优先：
         浏览器原生 fetch 流式写入
         --------------------------------------------------------

        try {
            showWorkerStatus(
                '开始直接保存：' +
                filename
            );

            const result =
                await streamFetchToFile(
                    url,
                    fileHandle,
                    filename
                );

            return {
                ...result,
                filename
            };

        } catch (streamError) {
            console.warn(
                '[WQ] fetch stream failed, fallback to GM_xhr',
                streamError
            );

            showWorkerStatus(
                '流式下载不可用，切换备用方式：' +
                filename
            );
        }

         --------------------------------------------------------
         第二路线：
         Tampermonkey 下载 Blob 后写入目标文件
         --------------------------------------------------------

        try {
            const blob =
                await gmFetchBlob(
                    url,
                    filename
                );

            const result =
                await writeBlobToFile(
                    blob,
                    fileHandle,
                    filename
                );

            return {
                ...result,
                filename
            };

        } catch (fallbackError) {
            
              两种路线都失败，删除可能留下的空文件残缺文件。
             
            try {
                if (
                    typeof dir.removeEntry ===
                    'function'
                ) {
                    await dir.removeEntry(
                        filename
                    );
                }
            } catch (_) {
            }

            throw new Error(
                '视频保存失败：' +
                fallbackError.message
            );
        }
    }

     ============================================================
     Preview Worker
     ============================================================

    async function runPreviewWorker() {
        const job =
            getNormalizedJob();

        if (
            !job 
            !job.running 
            job.current =
            job.items.length
        ) {
            
              手工打开 Preview 页面时，不自动保存。
             
            return;
        }

        const item =
            job.items[
                job.current
            ];

        const filename =
            safeFilename(
                item.desiredFilename 
                item.name
            );

         --------------------------------------------------------
         首先确认保存目录仍然可写
         --------------------------------------------------------

        const dir =
            await getGrantedDirectoryHandle();

        if (!dir) {
            const latest =
                getNormalizedJob();

            if (
                latest &&
                latest.current 
                latest.items.length
            ) {
                const current =
                    latest.items[
                        latest.current
                    ];

                current.status =
                    'pending';

                current.openedAt =
                    0;

                current.previewStartedAt =
                    0;

                current.downloadStartedAt =
                    0;

                current.error =
                    '保存目录权限失效，需要重新选择目录';

                latest.running =
                    false;

                saveJob(
                    latest
                );
            }

            showWorkerStatus(
                '保存目录权限失效。返回列表后请重新选择保存目录，再点“继续”。'
            );

            await sleep(
                1500
            );

            returnToList();

            return;
        }

         --------------------------------------------------------
         Preview 页面真正启动
         --------------------------------------------------------

        updateCurrentItem(
            current = {

                current.previewStartedAt =
                    Date.now();

                current.status =
                    'downloading';

                current.downloadStartedAt =
                    Date.now();

                current.error =
                    '';
            }
        );

        showWorkerStatus(
            '等待网站正常加载视频：' +
            filename
        );

         --------------------------------------------------------
         等真实媒体 URL
         --------------------------------------------------------

        const started =
            Date.now();

        let media =
            null;

        while (
            Date.now() -
            started 
            MEDIA_TIMEOUT
        ) {
            media =
                findBestMedia();

            if (media) {
                break;
            }

            await sleep(
                500
            );
        }

        if (!media) {
            updateCurrentItem(
                current = {

                    current.status =
                        'failed';

                    current.error =
                        '没有找到网站已经正常加载的 MP4WebM';
                }
            );

            showWorkerStatus(
                '失败：没有找到视频地址'
            );

            await sleep(
                POST_DOWNLOAD_DELAY
            );

            returnToList();

            return;
        }

        console.log(
            '[WQ] media',
            media.url
        );

         --------------------------------------------------------
         直接以课程文件名写入磁盘
         --------------------------------------------------------

        try {
            const result =
                await saveVideoDirect(
                    media.url,
                    filename
                );

            updateCurrentItem(
                current = {

                    current.status =
                        'done';

                    current.completedAt =
                        Date.now();

                    current.savedBytes =
                        result.bytes  0;

                    current.saveMode =
                        result.mode;

                    current.error =
                        '';
                }
            );

            showWorkerStatus(
                '已保存：' +
                result.filename +
                ' · ' +
                formatBytes(
                    result.bytes
                ) +
                ' · ' +
                result.mode +
                '，2 秒后继续'
            );

            await sleep(
                POST_DOWNLOAD_DELAY
            );

        } catch (error) {
            console.error(
                '[WQ] save video',
                error
            );

            if (
                error.code ===
                'DIRECTORY_PERMISSION_REQUIRED'
            ) {
                const latest =
                    getNormalizedJob();

                if (latest) {
                    const current =
                        latest.items[
                            latest.current
                        ];

                    current.status =
                        'pending';

                    current.error =
                        error.message;

                    latest.running =
                        false;

                    saveJob(
                        latest
                    );
                }

                showWorkerStatus(
                    '保存目录权限失效，任务已暂停。'
                );

            } else {
                updateCurrentItem(
                    current = {

                        current.status =
                            'failed';

                        current.error =
                            error.message;
                    }
                );

                showWorkerStatus(
                    '保存失败：' +
                    error.message
                );
            }

            await sleep(
                POST_DOWNLOAD_DELAY
            );
        }

        returnToList();
    }

     ============================================================
     返回列表
     ============================================================

    function returnToList() {
        const job =
            getNormalizedJob();

        if (!job) {
            return;
        }

        
          Preview 如果是新窗口：
          直接关闭。
         
        try {
            if (
                window.opener &&
                !window.opener.closed
            ) {
                window.close();
                return;
            }

        } catch (_) {
        }

        
          当前标签页：
          重新载入列表，而不是 history.back，
          避免 BFCache 导致队列不恢复。
         
        if (
            job.listUrl
        ) {
            showWorkerStatus(
                '正在返回资源列表，准备下一项…'
            );

            setTimeout(
                () = {

                    location.replace(
                        job.listUrl
                    );

                },
                300
            );
        }
    }

     ============================================================
     列表恢复
     ============================================================

    async function prepareListPage() {
        if (
            listSettling
        ) {
            return;
        }

        listSettling =
            true;

        try {
            installListPanel();

            setListStatus(
                '等待资源列表恢复…'
            );

            await sleep(
                LIST_SETTLE_DELAY
            );

            updateListUI();

            const job =
                getNormalizedJob();

            if (
                job.running &&
                !queueLoopRunning
            ) {
                setListStatus(
                    '列表已恢复，继续任务…'
                );

                processQueue()
                    .catch(
                        console.error
                    );
            }

        } finally {
            listSettling =
                false;
        }
    }

     ============================================================
     路由
     ============================================================

    async function routeCheck() {
        const href =
            location.href;

        if (
            isListPage()
        ) {
            if (
                href !== lastRoute
            ) {
                lastRoute =
                    href;

                await prepareListPage();

            } else {
                installListPanel();

                const job =
                    getNormalizedJob();

                if (
                    job.running &&
                    !queueLoopRunning &&
                    !listSettling
                ) {
                    processQueue()
                        .catch(
                            console.error
                        );
                }
            }

            return;
        }

        if (
            isPreviewPage()
        ) {
            lastRoute =
                href;

            if (
                previewWorkerRoute !==
                href
            ) {
                previewWorkerRoute =
                    href;

                runPreviewWorker()
                    .catch(
                        error = {

                            console.error(
                                '[WQ Preview Worker]',
                                error
                            );

                            updateCurrentItem(
                                current = {

                                    current.status =
                                        'failed';

                                    current.error =
                                        error.message;
                                }
                            );
                        }
                    );
            }

            return;
        }

        lastRoute =
            href;
    }

     ============================================================
     BFCache 恢复保护
     ============================================================

    window.addEventListener(
        'pageshow',
        async event = {

            if (
                !isListPage()
            ) {
                return;
            }

            console.log(
                '[WQ] pageshow',
                event.persisted
            );

            queueLoopRunning =
                false;

            lastRoute =
                '';

            installListPanel();

            setListStatus(
                '页面已恢复，准备继续…'
            );

            await sleep(
                LIST_SETTLE_DELAY
            );

            updateListUI();

            const job =
                getNormalizedJob();

            if (
                job.running
            ) {
                processQueue()
                    .catch(
                        console.error
                    );
            }
        }
    );

     ============================================================
     Watchdog
     ============================================================

    setInterval(
        () = {

            if (
                !isListPage() 
                listSettling
            ) {
                return;
            }

            const job =
                getNormalizedJob();

            if (
                !job 
                !job.running 
                queueLoopRunning
            ) {
                return;
            }

            setListStatus(
                '检测到队列暂停，自动恢复…'
            );

            processQueue()
                .catch(
                    console.error
                );

        },
        WATCHDOG_INTERVAL
    );

     ============================================================
     初始化
     ============================================================

    installPanelDelegation();

    routeCheck()
        .catch(
            console.error
        );

    setInterval(
        () = {

            routeCheck()
                .catch(
                    console.error
                );

        },
        ROUTE_INTERVAL
    );

})();