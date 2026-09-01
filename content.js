(function () {
  'use strict';

  // ============================================================
  //  配置区 —— 根据实际情况调整（选择器如果变了只改这里）
  // ============================================================
  const CONFIG = {
    // Excel 列映射（按列名匹配，支持别名数组，取第一个匹配成功的）
    columns: {
      boxSku:   ['订单行/包装/Box SKU', 'Box SKU'],            // 整箱 SKU（包装含 box 时用）
      sku:      ['订单行/包装/SKU', 'SKU'],                    // 单件 SKU（包装为 piece 时用）
      innerRef: ['内部参考号', '订单行/产品/内部参考号'],       // 内部参考号/UPC（校验用）
      packaging:['订单行/包装', '包装'],                        // 包装类型（box / piece）
      quantity: ['订单行/包装数量', '包装数量'],                // 加购数量（重要！）
      batchNo:  ['订单关联'],                                   // 订单关联（日志命名用）
    },

    // B 站 DOM 选择器
    selectors: {
      // 规格选项容器中的可点击项
      specOption: 'div.Range.ng-binding',
      // 隐藏的 SKU select（Angular Material）
      hiddenSelect: 'select.md-visually-hidden',
      // 加购按钮（选规格前是 inactive，选中后变 active）
      addToBagBtn: 'span.inactive, span.active, [class*="add"] span, button[class*="add"]',
      // 数量输入框
      qtyInput: 'input[type="number"], input[ng-model*="qty"], input[ng-model*="quantity"]',
      // 成功弹窗检测关键词
      successText: 'added to your shopping bag',
      // 关闭弹窗按钮
      closeBtn: 'button, span, div',
      closeBtnText: 'CLOSE',
      // 商品页面上显示的 UPC / Catalog No
      upcElement: '[class*="upc"], [class*="UPC"], [class*="catalog"]',
      // 「Select size or color」触发器（多规格商品页面上的 div[ng-click=showProductOptionsDialog]）
      selectSizeTrigger: '[ng-click*="showProductOptionsDialog"]',
      // 规格弹窗（Angular Material md-dialog）
      selectSizeDialog: 'md-dialog, [role="dialog"], .md-dialog-container, .md-dialog',
      // 弹窗内规格选项按钮（ng-click 选择规格）
      dialogOptionBtn: 'button[ng-click*="onSelectProductOption"]',
      // 弹窗内 ADD TO BAG 按钮
      dialogAddToBagBtn: 'button[ng-click*="onAddSelectedProductOptionToBag"]',
    },

    // 行为参数
    behavior: {
      delayMin: 2000,       // 操作间最小延迟 (ms)
      delayMax: 4000,       // 操作间最大延迟 (ms)
      pageLoadTimeout: 8000,// 页面加载超时 (ms)
      modalWaitTimeout: 8000, // 等待加购弹窗超时 (ms)
      optionWaitTimeout: 8000, // 点击规格选项后等待后台页面刷新 UPC 的超时 (ms)
      retryCount: 2,        // 单条失败重试次数
      mergeSameSku: true,   // 同一 SKU+包装 的多行（不同订单关联）合并为一次加购，数量求和
    },
  };

  // ============================================================
  //  错误码常量（集中管理，扩展只需在此加一行）
  // ============================================================
  const ERRORS = {
    NO_SKU:          '无SKU，需人工加购',       // 单件 SKU（piece）缺失
    NO_BOX_SKU:      '无Box SKU，需人工加购',   // 整箱 SKU（box）缺失
    UNSUPPORTED_TYPE:'暂未支持该加购类型，需人工加购',
    OUT_OF_STOCK:    '缺货',
    PAGE_TIMEOUT:    '页面加载超时，需人工加购',
    CF_TIMEOUT:      'Cloudflare验证超时，需人工加购',
    UPC_MISMATCH:    '该SKU与该商品的UPC不匹配，请人工检测SKU与UPC进行加购',
  };

  // 未知错误前缀（与 ERRORS 区分，Excel 中标记为开发排查类）
  const UNKNOWN_ERR = '未知错误，请下载日志后交给开发人员排查';

  // ============================================================
  //  跨页面持久化（页面跳转刷新后面板/日志/任务全部恢复）
  // ============================================================
  // 批次键（挂到当前 runId 下的键，刷新/跳页不丢）
  const BATCH_BASE_KEYS = [
    'abw_logs', 'abw_parsed_items', 'abw_pending', 'abw_pending_results',
    'abw_xl_headers', 'abw_xl_rows', 'abw_batch_no', 'abw_file_name', 'abw_total_count',
  ];

  // 当前总任务标识（runId）：上传一次文件生成一个，刷新后据此恢复
  let currentRunId = (() => { try { return sessionStorage.getItem('abw_currentRunId') || ''; } catch (e) { return ''; } })();

  // 把逻辑键映射到物理键：批次键 → abw_run_<id>_<key>；全局键原样
  function realKey(key) {
    if (currentRunId && BATCH_BASE_KEYS.includes(key)) {
      return `abw_run_${currentRunId}_${key.replace(/^abw_/, '')}`;
    }
    return key;
  }
  function newRunId() {
    return 'run_' + Date.now().toString(36) + '_' + Math.random().toString(36).slice(2, 6);
  }
  // 清空当前批次的所有数据
  function clearCurrentBatch() {
    if (!currentRunId) return;
    for (const k of BATCH_BASE_KEYS) {
      try { sessionStorage.removeItem(`abw_run_${currentRunId}_${k.replace(/^abw_/, '')}`); } catch (e) {}
    }
  }
  // 开启新批次：清旧批次 → 切新 runId → 存文件名
  function startNewBatch(fileName) {
    clearCurrentBatch();
    currentRunId = newRunId();
    try { sessionStorage.setItem('abw_currentRunId', currentRunId); } catch (e) {}
    STORE.set(K.fileName, fileName || '');
  }

  const STORE = {
    get(key, fallback) {
      try { return JSON.parse(sessionStorage.getItem(realKey(key)) || 'null') ?? fallback; } catch (e) { return fallback; }
    },
    // 直接写入；返回是否成功（不内部打日志，避免空间满时日志本身递归失败）
    set(key, val) {
      try { sessionStorage.setItem(realKey(key), JSON.stringify(val)); return true; }
      catch (e) {
        console.warn('[ABW] sessionStorage 写入失败，可能存储空间不足', key, e);
        return false;
      }
    },
    // 带抢救的写入：首次失败 → 自动腾空间（清理可重建数据）→ 重试一次
    // 返回是否最终成功。调用方（关键数据落盘处）应检查返回值并决定是否停止任务
    setWithRetry(key, val) {
      if (this.set(key, val)) return true;
      try {
        // ① 清理续跑点 pending（剩余任务可从 items 重建，可丢弃）
        sessionStorage.removeItem(realKey(K.pending));
        // ② 截断日志到最近 200 条（日志可重建，是最大占用）
        const logsKey = realKey(K.logs);
        const logs = JSON.parse(sessionStorage.getItem(logsKey) || '[]');
        if (logs.length > 200) sessionStorage.setItem(logsKey, JSON.stringify(logs.slice(-200)));
        // ③ 清理所有非当前批次的残留键（旧 runId 数据，腾出空间）
        const keepPrefix = 'abw_run_' + currentRunId + '_';
        for (let i = sessionStorage.length - 1; i >= 0; i--) {
          const k = sessionStorage.key(i);
          if (k && k.startsWith('abw_run_') && !k.startsWith(keepPrefix)) sessionStorage.removeItem(k);
        }
      } catch (e2) { /* 腾空间失败也继续尝试写入 */ }
      return this.set(key, val);
    },
    remove(key) {
      try { sessionStorage.removeItem(realKey(key)); } catch (e) {}
    },
  };
  const K = {
    logs: 'abw_logs',        // 日志（最多 500 条）
    pos: 'abw_panel_pos',    // 面板位置
    collapsed: 'abw_panel_collapsed', // 折叠状态
    items: 'abw_parsed_items', // 解析的任务
    pending: 'abw_pending',     // 跳转中剩余任务（静默恢复，无日志）
    results: 'abw_pending_results',
    xlHeaders: 'abw_xl_headers', // 原始 Excel 表头（下载时用）
    xlRows: 'abw_xl_rows',       // 原始 Excel 全部行（下载时用）
    batchNo: 'abw_batch_no',     // 批次号（日志用）
    fileName: 'abw_file_name',   // 当前批次来源文件名
    totalCount: 'abw_total_count', // 本批次任务总数（进度条分母）
  };
  const MAX_LOGS = 5000; // ~350KB，覆盖1300+商品（日志过多会撑爆 sessionStorage，导致跳转后无法续跑）

  // ============================================================
  //  工具函数
  // ============================================================
  const log = (msg, type = 'info') => {
    // 持久化日志（跨页面保留）
    const logs = STORE.get(K.logs, []);
    const truncated = logs.length >= MAX_LOGS;
    logs.push({ t: Date.now(), msg, type });
    if (logs.length > MAX_LOGS) {
      logs.splice(0, logs.length - MAX_LOGS);
      if (!truncated) console.warn(`[ABW] 日志已达上限 ${MAX_LOGS} 条，早期日志将被丢弃`);
    }
    STORE.set(K.logs, logs);

    // 渲染到面板
    const el = document.getElementById('abw-log');
    if (el) {
      const time = new Date().toLocaleTimeString('zh-CN', { hour12: false });
      const colors = { info: '#888', success: '#2ecc71', warn: '#f39c12', error: '#e74c3c', action: '#3498db' };
      const entry = document.createElement('div');
      entry.style.cssText = `color:${colors[type]||colors.info};padding:2px 0;font-size:12px;border-bottom:1px solid rgba(255,255,255,0.05);`;
      entry.textContent = `[${time}] ${msg}`;
      el.appendChild(entry);
      el.scrollTop = el.scrollHeight;
    }
    console.log(`[ABW-${type.toUpperCase()}] ${msg}`);
  };

  // 恢复历史日志到面板
  function restoreLogs() {
    const el = document.getElementById('abw-log');
    if (!el) return;
    const logs = STORE.get(K.logs, []);
    const colors = { info: '#888', success: '#2ecc71', warn: '#f39c12', error: '#e74c3c', action: '#3498db' };
    logs.forEach(l => {
      const entry = document.createElement('div');
      const time = new Date(l.t).toLocaleTimeString('zh-CN', { hour12: false });
      entry.style.cssText = `color:${colors[l.type]||colors.info};padding:2px 0;font-size:12px;border-bottom:1px solid rgba(255,255,255,0.05);`;
      entry.textContent = `[${time}] ${l.msg}`;
      el.appendChild(entry);
    });
    el.scrollTop = el.scrollHeight;
  }

  // ============================================================
  //  停止控制（立即停止）
  // ============================================================
  let forceStop = false;
  const StopError = new Error('USER_STOP');
  const checkStop = () => { if (forceStop) throw StopError; };

  // ============================================================
  //  页面就绪检测（增强版：DOM 稳定 + 元素真正可见）
  // ============================================================

  // 严格判断元素是否真正可见、可交互（不只是存在于 DOM）
  // allowDisabled=true 时放行 disabled 按钮 —— 按钮可能因页面状态暂时禁用，仍需能找到
  function isElementReal(el, allowDisabled = false) {
    if (!el) return false;
    // 1) 必须挂载在 document 中
    if (!document.contains(el)) return false;
    // 2) offsetParent 为 null → 元素或祖先 display:none（最常见的 Angular 未渲染情况）
    if (el.offsetParent === null && el.tagName !== 'BODY') return false;
    // 3) 尺寸为 0 不可交互
    const rect = el.getBoundingClientRect();
    if (rect.width <= 1 || rect.height <= 1) return false;
    // 4) disabled（默认视为不可用；allowDisabled=true 时放行）
    if (!allowDisabled && el.disabled) return false;
    // 5) computed style: hidden / none / opacity:0
    const style = window.getComputedStyle(el);
    if (style.visibility === 'hidden') return false;
    if (style.display === 'none') return false;
    if (parseFloat(style.opacity) === 0) return false;
    return true;
  }

  // MutationObserver 等待 DOM 稳定（连续 stableMs 毫秒无变更）
  async function waitForDomStable(stableMs = 3000, timeout = 30000) {
    const start = Date.now();
    return new Promise((resolve) => {
      let stableTimer = null;
      let done = false;

      const finish = (ok) => {
        if (done) return;
        done = true;
        observer.disconnect();
        clearTimeout(stableTimer);
        resolve(ok);
      };

      const observer = new MutationObserver(() => {
        clearTimeout(stableTimer);
        stableTimer = setTimeout(() => finish(true), stableMs);
      });

      try {
        observer.observe(document.body, {
          childList: true,
          subtree: true,
          attributes: true,
          attributeFilter: ['class', 'style', 'ng-class', 'ng-style', 'ng-show', 'ng-hide'],
          characterData: true,
        });
      } catch (e) {
        // 极端情况 body 不存在，直接 resolve
        resolve(false);
        return;
      }

      // 初始计时器：如果 DOM 本来就不变，stableMs 后触发
      stableTimer = setTimeout(() => finish(true), stableMs);

      // 总超时
      setTimeout(() => finish(false), timeout);
    });
  }

  // 等待页面完全加载（readyState complete + 额外缓冲）
  async function waitForPageReady(timeout = 15000) {
    const start = Date.now();
    while (document.readyState !== 'complete') {
      if (Date.now() - start > timeout) break;
      await sleep(500);
    }
    await sleep(2000);
  }

  // 等待元素真正可见（不是只查 DOM 存在）
  async function waitForElement(selector, timeout = 10000) {
    const start = Date.now();
    while (Date.now() - start < timeout) {
      const el = document.querySelector(selector);
      if (el && isElementReal(el)) {
        log(`  ✓ 元素就绪: ${selector}`, 'info');
        return true;
      }
      await sleep(500);
    }
    return false;
  }

  // 组合等待：DOM 稳定 + 加购按钮真正可见可点击（并行检测，按钮先就绪则提前放行）
  // 返回 { ready: true } | { ready: false, reason: 'dom_timeout'|'no_button' }
  async function waitForProductPageReady(timeout = 40000) {
    log(`  ⏳ 等待商品页完全就绪...`, 'info');
    const start = Date.now();
    let lastMutation = Date.now();

    const observer = new MutationObserver(() => { lastMutation = Date.now(); });
    try {
      observer.observe(document.body, {
        childList: true, subtree: true,
        attributes: true, attributeFilter: ['class', 'style', 'ng-class', 'ng-style', 'ng-show', 'ng-hide'],
        characterData: true,
      });
    } catch (e) { /* body 不存在时降级 */ }

    while (Date.now() - start < timeout) {
      checkStop();
      const stableMs = Date.now() - lastMutation;
      const btn = await findAddToBagButton();

      // 按钮就绪 + DOM 稳定 1 秒 + readyState complete → 立即放行
      if (btn && document.readyState === 'complete' && stableMs >= 1000) {
        observer.disconnect();
        log(`  ✅ 商品页就绪 (${Date.now() - start}ms)`, 'success');
        return { ready: true };
      }

      // DOM 已稳定 3 秒但仍无按钮 → 缺货（最多再确认 8 秒）
      if (!btn && stableMs >= 3000) {
        observer.disconnect();
        log(`  ✓ DOM 已稳定，确认加购按钮...`, 'info');
        const btnReady = await waitForAddToBagBtnVisible(8000);
        if (!btnReady) {
          return { ready: false, reason: 'no_button' };
        }
        log(`  ✅ 商品页就绪 (${Date.now() - start}ms)`, 'success');
        return { ready: true };
      }

      await sleep(500);
    }

    observer.disconnect();
    log(`  ⚠️ DOM 未在超时内稳定`, 'warn');
    return { ready: false, reason: 'dom_timeout' };
  }

  // 等待加购按钮真正可见（轮询 findAddToBagButton + isElementReal）
  async function waitForAddToBagBtnVisible(timeout = 15000) {
    const start = Date.now();
    while (Date.now() - start < timeout) {
      checkStop();
      const btn = await findAddToBagButton();
      if (btn) {
        log(`  ✓ 加购按钮可见 (${Date.now() - start}ms)`, 'info');
        return true;
      }
      await sleep(1000);
    }
    return false;
  }

  // 可被打断的 sleep：点击「立即停止」后所有等待立即抛出 StopError
  const sleep = async (ms) => {
    checkStop();
    await new Promise(r => setTimeout(r, ms));
    checkStop();
  };
  const randomDelay = () => sleep(CONFIG.behavior.delayMin + Math.random() * (CONFIG.behavior.delayMax - CONFIG.behavior.delayMin));

  // 安全的 querySelector（带重试）
  const waitFor = async (selector, timeout = 5000, interval = 300) => {
    const start = Date.now();
    while (Date.now() - start < timeout) {
      // 先尝试主文档
      let el = document.querySelector(selector);
      if (el) return el;
      // 再尝试 iframe（B站有1个iframe）
      try {
        const iframes = document.querySelectorAll('iframe');
        for (const iframe of iframes) {
          try {
            el = iframe.contentDocument?.querySelector(selector);
            if (el) return el;
          } catch (e) { /* cross-origin iframe 跳过 */ }
        }
      } catch (e) {}
      await sleep(interval);
    }
    return null;
  };

  // const waitForAll = async (selector, timeout = 5000, interval = 300) => {
  //   const start = Date.now();
  //   while (Date.now() - start < timeout) {
  //     let els = document.querySelectorAll(selector);
  //     if (els.length > 0) return Array.from(els);
  //     try {
  //       const iframes = document.querySelectorAll('iframe');
  //       for (const iframe of iframes) {
  //         try {
  //           els = iframe.contentDocument?.querySelectorAll(selector);
  //           if (els && els.length > 0) return Array.from(els);
  //         } catch (e) {}
  //       }
  //     } catch (e) {}
  //     await sleep(interval);
  //   }
  //   return [];
  // };

  // // 查找包含指定文本的元素
  // const findByText = (selector, text) => {
  //   const els = document.querySelectorAll(selector);
  //   for (const el of els) {
  //     if (el.textContent.trim().toLowerCase().includes(text.toLowerCase())) return el;
  //   }
  //   // 也检查 iframe
  //   const iframes = document.querySelectorAll('iframe');
  //   for (const iframe of iframes) {
  //     try {
  //       const iframeEls = iframe.contentDocument?.querySelectorAll(selector);
  //       if (iframeEls) {
  //         for (const el of iframeEls) {
  //           if (el.textContent.trim().toLowerCase().includes(text.toLowerCase())) return el;
  //         }
  //       }
  //     } catch (e) {}
  //   }
  //   return null;
  // };

  // ============================================================
  //  Excel 解析器
  // ============================================================
  function parseExcel(buffer) {
    const workbook = XLSX.read(buffer, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];

    // 只展开实际有数据的行数：sheet 是稀疏存储（仅非空单元格有键），
    // 先扫描最后一数据行，再按该范围展开，避免 !ref 全表范围（如 A1:M1048576）
    // 展开出百万空行、生成几十 MB 数组，导致 abw_xl_rows 写入 sessionStorage 失败
    let lastDataRow = 0;
    for (const key of Object.keys(sheet)) {
      if (!key || key[0] === '!') continue;
      const m = key.match(/^([A-Z]+)(\d+)$/);
      if (m) lastDataRow = Math.max(lastDataRow, parseInt(m[2], 10));
    }
    const origRange = XLSX.utils.decode_range(sheet['!ref'] || 'A1');
    const range = `A1:${XLSX.utils.encode_col(origRange.e.c)}${Math.min(Math.max(lastDataRow, 1), origRange.e.r + 1)}`;
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', range });

    if (rows.length < 2) throw new Error('Excel 数据行不足');

    // 第一行是表头
    const headers = rows[0].map(h => String(h || '').trim());

    // 列名清洗：去掉所有非字母数字/中文的字符，小写
    const cleanHeader = (h) => String(h || '').replace(/[^a-zA-Z0-9\u4e00-\u9fa5]/g, '').toLowerCase();
    const cleanHeaders = headers.map(cleanHeader);

    // 找列索引：最短匹配优先（避免 '包装' 误匹配到 '包装/Box SKU' 列）
    function findColumnIndex(names) {
      for (const name of names) {
        const cleanName = cleanHeader(name);
        if (!cleanName) continue;
        // ① 精确相等
        let hit = cleanHeaders.findIndex(h => h === cleanName);
        if (hit !== -1) return hit;
        // ② 包含关系：取所有候选里"最短"的那个（最接近目标列名）
        const candidates = [];
        cleanHeaders.forEach((h, i) => { if (h && h.includes(cleanName)) candidates.push({ i, len: h.length }); });
        if (candidates.length > 0) {
          candidates.sort((a, b) => a.len - b.len);
          return candidates[0].i;
        }
      }
      return -1;
    }

    // 解析列映射
    const colIdx = {};
    const colMapLog = [];
    for (const [key, names] of Object.entries(CONFIG.columns)) {
      const idx = findColumnIndex(names);
      colIdx[key] = idx;
      colMapLog.push(`${key}→${idx >= 0 ? `col${idx}(${headers[idx]})` : '❌未找到'}`);
      if (idx === -1) console.warn(`[ABW] 未找到列 ${names[0]}，可用列: ${headers.join(' | ')}`);
    }
    console.log('[ABW] 列映射: ' + colMapLog.join(' | '));

    // 关键列缺失时直接报错（不再静默 fallback 成默认值！）
    const required = ['packaging', 'quantity'];
    const missing = required.filter(k => colIdx[k] === -1);
    if (missing.length > 0) {
      throw new Error(`关键列未找到: ${missing.join(', ')}（可用列: ${headers.join(' | ')}）`);
    }
    // SKU 列：Box SKU 与单件 SKU 至少有一个（按包装类型二选一）
    if (colIdx.boxSku === -1 && colIdx.sku === -1) {
      throw new Error(`SKU 列未找到（需包含"订单行/包装/Box SKU"或"订单行/包装/SKU"，可用列: ${headers.join(' | ')}）`);
    }

    // ---- 断点续跑检测 ----
    // 文件里存在"加购结果"列 → 说明是加购一半的结果文件：
    // 结果非空的行视为已处理（跳过），只加购结果为空的行
    const resultColIdx = findColumnIndex(['加购结果']);
    const resumeStats = {
      hadResultCol: resultColIdx >= 0,
      doneCount: 0,      // 已处理行数
      successCount: 0,   // 其中成功（含已在购物车）
      skipCount: 0,      // 其中缺货
      failCount: 0,      // 其中失败
    };

    // 解析数据行（跳过表头）
    const items = [];
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row || row.every(c => c == null || String(c || '').trim() === '')) continue;

      // 断点续跑：该行已有加购结果（非空）→ 已处理，跳过
      if (resultColIdx >= 0) {
        const doneLabel = String(row[resultColIdx] || '').trim();
        if (doneLabel) {
          resumeStats.doneCount++;
          if (/成功|已在购物车/.test(doneLabel)) resumeStats.successCount++;
          else if (/缺货/.test(doneLabel)) resumeStats.skipCount++;
          else resumeStats.failCount++;
          continue;
        }
      }

      const packaging = String(row[colIdx.packaging] || '').trim();
      // 按包装类型选 SKU 列：含 box → Box SKU；否则（piece）→ 单件 SKU
      const useBoxSku = /box/i.test(packaging);
      const skuCell = useBoxSku
        ? (colIdx.boxSku >= 0 ? row[colIdx.boxSku] : '')
        : (colIdx.sku >= 0 ? row[colIdx.sku] : '');
      const item = {
        _row: i + 1,
        boxSku: String(skuCell || '').trim(),
        innerRef: String(row[colIdx.innerRef] || '').trim(),
        packaging,
        quantity: parseInt(row[colIdx.quantity]) || 1,
        _noSku: false, // 标记无 SKU 的行
      };

      // 所选 SKU 列为空/'-' 时标记，运行时直接跳过（box 缺 Box SKU、piece 缺单件 SKU 都提示）
      if (!item.boxSku || item.boxSku === '-') {
        item._noSku = true;
        item._noSkuReason = useBoxSku ? ERRORS.NO_BOX_SKU : ERRORS.NO_SKU;
      }
      items.push(item);
    }

    console.log('[ABW] Excel 解析结果:', { headers, colIdx, itemCount: items.length, sample: items[0] });
    // 提取批次号（第一行数据）
    let batchNo = '';
    if (colIdx.batchNo >= 0 && items.length > 0) {
      const firstRowIdx = items[0]._row - 1; // _row 从1开始
      batchNo = String(rows[firstRowIdx][colIdx.batchNo] || '').trim();
    }
    return { items, headers, colMapLog, rawRows: rows, batchNo, resumeStats };
  }

  // ============================================================
  //  同 SKU 聚合（合并加购）
  //  同一「SKU+包装」出现在多行（分属不同订单关联）→ 合并为一个任务，
  //  数量求和、一次导航一次加购，避免重复导航/第二次加购覆盖数量。
  //  输入已含断点续跑过滤（只含本次要跑的行），组内部分行已处理时聚合的
  //  是「未处理行」的数量之和，购物车总量累计正确。
  // ============================================================
  function mergeSameSkuItems(items) {
    if (!CONFIG.behavior.mergeSameSku || items.length < 2) return items;

    const groups = new Map(); // key: sku|packaging
    const order = [];         // 聚合后的任务列表（保持 Excel 顺序）
    const upcWarnings = [];
    let mergedRows = 0;

    for (const it of items) {
      // 无 SKU 的行不聚合（运行时直接标记失败），保持单任务
      if (it._noSku || !it.boxSku) {
        order.push(it);
        continue;
      }
      const key = `${it.boxSku}|${it.packaging}`;
      const g = groups.get(key);
      if (!g) {
        const merged = { ...it, _rows: [it._row] };
        groups.set(key, merged);
        order.push(merged);
      } else {
        g.quantity += it.quantity;
        g._rows.push(it._row);
        // 内部参考号不一致 → 数据异常提示（UPC 校验仍用组内第一行）
        if (it.innerRef && g.innerRef && it.innerRef !== g.innerRef) {
          upcWarnings.push(`SKU=${it.boxSku}: Row${g._rows[0]} 与 Row${it._row} 的内部参考号不一致 (${g.innerRef} vs ${it.innerRef})，请注意核对`);
        }
        mergedRows++;
      }
    }

    // 统一补齐 _rows（未合并/无SKU 的任务也是单元素数组，恢复与回写逻辑统一）
    for (const it of order) {
      if (!it._rows) it._rows = [it._row];
    }

    if (mergedRows > 0) {
      log(`🔗 同SKU聚合: ${items.length} 行 → ${order.length} 个任务（合并 ${mergedRows} 行，数量已求和，一次导航一次加购）`, 'action');
      upcWarnings.forEach(w => log(`  ⚠️ ${w}`, 'warn'));
    }
    return order;
  }

  // 取一个结果对应的所有原始 Excel 行号（兼容无 _rows 的旧批次数据）
  function resultRows(r) {
    const item = r && r.item;
    if (!item) return [];
    if (item._rows && item._rows.length) return item._rows;
    return item._row ? [item._row] : [];
  }

  // 结果去重：同一任务（按主行号 _row）只保留最后一条结果
  // 兼容旧版本遗留的重复条目（navigating 跳页被重复 push，导致"已完成 N/总 M"虚高，如 14/7）
  function dedupeResults(list) {
    if (!list || list.length < 2) return list;
    const map = new Map();
    for (const r of list) {
      const key = r && r.item && r.item._row;
      if (key == null) continue;
      map.set(key, r);
    }
    return [...map.values()];
  }

  // ============================================================
  //  核心执行引擎
  // ============================================================
  let isRunning = false;
  const results = [];

  // 任务执行标识：统一驱动「立即停止 / 开始加购」按钮显示 + 持久化（刷新/跳页可恢复）
  const RUNNING_KEY = 'abw_running';
  function setRunning(running) {
    isRunning = running;
    try { sessionStorage.setItem(RUNNING_KEY, running ? '1' : '0'); } catch (e) {}
    const stopBtn = document.getElementById('abw-stop-btn');
    const startBtn = document.getElementById('abw-start-btn');
    if (stopBtn) stopBtn.style.display = running ? '' : 'none';
    if (startBtn) startBtn.style.display = running ? 'none' : '';
  }
  // 启动时读取持久化标识（跳页续跑时保持运行态，恢复逻辑会据 pending 再校正）
  isRunning = (() => { try { return sessionStorage.getItem(RUNNING_KEY) === '1'; } catch (e) { return false; } })();

  // 设置页面数量输入框（值相同则跳过）
  async function setQty(addQty) {
    const qtyInput = await waitFor(CONFIG.selectors.qtyInput, 3000);
    if (!qtyInput) {
      log(`  ⚠️ 未找到数量输入框，使用默认数量`, 'warn');
      return;
    }
    const target = String(addQty);
    if (qtyInput.value === target) return; // 已是目标值，避免多余事件刷新
    qtyInput.focus();
    qtyInput.select();
    const nativeInputValueSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
    nativeInputValueSetter.call(qtyInput, target);
    qtyInput.dispatchEvent(new Event('input', { bubbles: true }));
    qtyInput.dispatchEvent(new Event('change', { bubbles: true }));
    await sleep(500);
  }

  // 加购后半段（设数量 → 点击 → 校验弹窗）
  async function doAddToBag(_addBtn, item, addQty) {
    // Step 5: 设置数量
    log(`  → 设置数量: ${addQty}`, 'info');
    await setQty(addQty);
    // Step 6: 重新查找加购按钮（设数量后 Angular 可能替换了 DOM）
    const addBtn = await findAddToBagButton();
    if (!addBtn) throw new Error('设数量后找不到加购按钮');
    log(`  → 点击加购...`, 'info');
    addBtn.click();
    await sleep(1000);
    // Step 7: 检测是否弹出 "Select size or color" 弹窗（遍历规格选项，按 UPC 匹配加购）
    if (isSelectSizeDialogOpen()) {
      log(`  ⚠️ 检测到 "Select size or color" 弹窗，开始遍历规格选项匹配 UPC...`, 'warn');
      return await trySelectOptionAdd(item, addQty);
    }
    // Step 8: 等待弹窗并校验
    log(`  → 等待加购结果...`, 'info');
    const success = await waitForSuccessModal();
    if (!success) throw new Error('超时未检测到加购成功弹窗');
    // Step 9: 关闭弹窗
    await closeModal();
    await randomDelay();
    log(`  ✅ 加购成功!`, 'success');
    return { status: 'success', item };
  }

  // 从页面提取价格：根据包装规格匹配，返回 { unitPrice, boxPrice } 或 null
  function extractPriceFromPage(packaging) {
    if (!packaging) return null;
    const cleanPkg = packaging.replace(/\s+/g, ' ').trim().toLowerCase();
    const links = document.querySelectorAll('a');
    for (const link of links) {
      const text = (link.textContent || '').replace(/\s+/g, ' ').trim().toLowerCase();
      if (!text.includes('box') && !text.includes('piece')) continue;
      if (!text.includes(cleanPkg)) continue;
      // 提取第一个 HK$ 数字
      const firstMatch = text.match(/hk\$\s*([\d,]+\.?\d*)/i);
      if (!firstMatch) continue;
      const firstPrice = firstMatch[1].replace(/,/g, '');
      // 含 "average" → 整箱批发，boxPrice=第一个HK$, unitPrice=average后面的
      const avgMatch = text.match(/average\s*hk\$\s*([\d,]+\.?\d*)/i);
      if (avgMatch) {
        return { boxPrice: firstPrice, unitPrice: avgMatch[1].replace(/,/g, '') };
      }
      // 不含 "average" → 散件，无整箱价，单价就是第一个HK$
      return { boxPrice: '', unitPrice: firstPrice };
    }
    return null;
  }

  // 单条执行：SKU URL 直达商品页 → 检测加购按钮 → 设数量加购
  // allItems/curIdx: 当前批次列表与索引，供 CF 挑战场景保存续跑点（页面 reload 后自动续跑）
  async function executeItem(item, attempt = 1, allItems = null, curIdx = -1) {
    const { boxSku, innerRef, packaging, quantity, _row } = item;
    const addQty = quantity;

    // 无 SKU → 直接标记失败（box 缺 Box SKU / piece 缺单件 SKU）
    if (item._noSku) {
      const reason = item._noSkuReason || ERRORS.NO_SKU;
      log(`📦 第${_row}行: ${reason}`, 'warn');
      return { status: 'failed', item, reason };
    }

    // 聚合任务时日志显示"第N行"或"第N等M行"
    const rowLabel = item._rows && item._rows.length > 1 ? `${_row}等${item._rows.length}行` : `${_row}`;
    log(`📦 第${rowLabel}行: SKU=${boxSku} | 规格=${packaging} | 数量=${addQty} | 第${attempt}次尝试`, 'action');

    try {
      // ---- Step 1: 导航到商品页 ----
      // 匹配 pid.<SKU> 或 cpid=<SKU>（部分商品是"主商品pid + 变体cpid"结构，
      // 服务器会把 pid.<cpid> 请求 301 重定向到 pid.<主pid>?cpid=<变体>，只认 pid 会永远匹配不上 → 无限刷新）
      const productUrl = `/info.html/pid.${boxSku}`;
      const isOnProductPage = new RegExp(
        `/info\\.html/pid\\.${boxSku}(?:[/?#]|$)|[?&]cpid=${boxSku}(?:&|$)`,
        'i'
      ).test(location.href);
      if (!isOnProductPage) {
        // 防循环导航保险: 同一 SKU 连续导航 3 次仍未定位到目标页 → 放弃并报错，避免无限刷新
        const navKey = 'abw_nav_' + boxSku;
        const navCount = parseInt(sessionStorage.getItem(navKey) || '0', 10) + 1;
        if (navCount > 3) {
          sessionStorage.removeItem(navKey);
          log(`  ❌ 导航 ${navCount - 1} 次后仍未定位到该 SKU 页面（SKU 可能是变体ID/已失效，需人工加购）`, 'error');
          return { status: 'failed', item, reason: 'SKU定位失败，需人工加购' };
        }
        sessionStorage.setItem(navKey, String(navCount));
        log(`  → 导航到商品页 (第${navCount}次)...`, 'info');
        window.location.href = productUrl;
        return { status: 'navigating', item };
      }
      // 已到达目标页，清除该 SKU 的防循环标记
      sessionStorage.removeItem('abw_nav_' + boxSku);

      await waitForPageReady();

      // ---- Cloudflare 安全验证检测（必须在 waitForProductPageReady 之前，避免商品页文本误匹配）----
      // 注意: location.reload() 会重建页面并重新注入脚本，内存计数随旧实例销毁而归零，
      // 若 CF 持续不通过（挑战页保持原 URL）会陷入"检测→刷新→再检测"的无限循环，CF_TIMEOUT 永不触发。
      // 因此把重试计数持久化到 sessionStorage，跨刷新保留。
      const cfKey = 'abw_cf_' + boxSku;
      const cfMaxRetries = 5;
      let cfRetries = 0;
      try { cfRetries = parseInt(sessionStorage.getItem(cfKey) || '0', 10) || 0; } catch (e) {}
      const isCFPage = () => {
        const text = (document.body && document.body.innerText) || '';
        return /(checking your browser|just a moment|ddos protection|cf-browser-verification|本网站使用安全服务)/i.test(text)
            || !!document.querySelector('#challenge-form, #cf-challenge-running, .cf-browser-verification');
      };
      if (isCFPage()) {
        cfRetries++;
        try { sessionStorage.setItem(cfKey, String(cfRetries)); } catch (e) {}
        log(`  ⚠️ Cloudflare 安全验证 (${cfRetries}/${cfMaxRetries})，等待后重试...`, 'warn');
        if (cfRetries >= cfMaxRetries) {
          try { sessionStorage.removeItem(cfKey); } catch (e) {}
          log(`  ❌ Cloudflare 验证未通过，跳过此商品`, 'error');
          return { status: 'failed', item, reason: ERRORS.CF_TIMEOUT };
        }
        // ★ 保存续跑点（当前条目及之后全部）：CF 挑战通过后页面会自动 reload，此时脚本实例被销毁、
        // 无法走 runAll 的 navigating 返回路径保存 pending —— 若不在等待前主动保存，
        // 新页面恢复时 pending 为空 → 落入「停止态等手动点击」，表现为"CF 通过后脚本不加购"。
        // 保存后无论页面是自动 reload 还是下方手动 reload，新页面都会走「分支① 自动续跑」。
        if (allItems && curIdx >= 0) {
          STORE.setWithRetry(K.pending, allItems.slice(curIdx));
        }
        // CF 的 JS 挑战通过后会自动刷新放行；先给 15s 让挑战自动完成（期间页面自动 reload 则旧实例随页面销毁自然终止），
        // 仍未放行才强制 reload，并走 navigating 机制让新页面实例续跑（计数已持久化，不会被重置）
        await sleep(15000);
        if (isCFPage()) {
          location.reload();
          return { status: 'navigating', item };
        }
        // 挑战已通过且页面未自动 reload（内嵌 Turnstile 类验证）→ 清除本次续跑点，正常继续加购
        if (allItems && curIdx >= 0) {
          STORE.remove(K.pending);
        }
      } else {
        // 页面已放行，清除历史计数
        try { sessionStorage.removeItem(cfKey); } catch (e) {}
      }

      // ---- Step 1.1: 等待商品页关键元素就绪 ----
      const ready = await waitForProductPageReady(40000);
      if (!ready.ready) {
        if (ready.reason === 'no_button') {
          log(`  ⚠️ 页面已加载但无加购按钮，判定缺货`, 'warn');
          return { status: 'skipped', item, reason: ERRORS.OUT_OF_STOCK };
        }
        log(`  ⚠️ 商品页加载超时（DOM 未稳定），跳过`, 'warn');
        return { status: 'failed', item, reason: ERRORS.PAGE_TIMEOUT };
      }

      // ---- Step 3: 校验 UPC ----
      // 多规格商品（页面上有 "Select size or color" 触发器）跳过页面 UPC 预检，
      // 由规格弹窗流程（trySelectOptionAdd）逐个选项匹配 UPC；
      // 单规格商品保持原有页面 UPC 校验
      const hasOptionTrigger = !!findSelectSizeTrigger();
      const pageText = document.body.innerText;
      const upcMatch = pageText.match(/UPC[:\s]*(\d[\d\s-]{8,})/i);
      if (!hasOptionTrigger && upcMatch) {
        const pageUpc = upcMatch[1].replace(/\s/g, '');
        const refClean = innerRef.replace(/\s/g, '');
        if (refClean && !pageUpc.includes(refClean) && !refClean.includes(pageUpc)) {
          log(`  ❌ UPC 不匹配: 页面=${pageUpc} vs Excel=${refClean}`, 'error');
          return { status: 'failed', item, reason: ERRORS.UPC_MISMATCH };
        }
      }

      // ---- Step 2: 加购按钮二次确认（waitForProductPageReady 已确认过，这里兜底） ----
      const addBtn = await findAddToBagButton();
      if (!addBtn) {
        log(`  ⚠️ 无加购按钮（缺货），跳过`, 'warn');
        return { status: 'skipped', item, reason: ERRORS.OUT_OF_STOCK };
      }


      // 提取商品英文名
      const h1 = document.querySelector('h1');
      if (h1) item._abwProductName = h1.textContent.trim();

      // 提取价格（根据 Excel 包装列匹配页面链接）
      const priceInfo = extractPriceFromPage(packaging);
      if (priceInfo) {
        item._boxPrice = priceInfo.boxPrice;
        item._unitPrice = priceInfo.unitPrice;
        log(`  💰 价格: 整箱=HK$${priceInfo.boxPrice} 单价=HK$${priceInfo.unitPrice}`, 'info');
      } else {
        log(`  ⚠️ 未匹配到包装"${packaging}"的价格`, 'warn');
      }

      // ---- Step 3: 配送时间检测（只查商品购买区域，>=21 天才标记可能缺货）----
      const shipArea = document.querySelector('.buyingOption, .productInfo, .shippingGrid');
      const shipText = shipArea ? shipArea.innerText : pageText;
      const shipMatch = shipText.match(/usually ships within (\d+)(?:\s*(?:to|[-–])\s*\d+)?\s*days/i);
      if (shipMatch && parseInt(shipMatch[1], 10) >= 21) {
        item._shipWarning = `Usually ships within ${shipMatch[1]} days`;
        log(`  ⚠️ ${item._shipWarning}（可能缺货）`, 'warn');
      }

      // ---- Step 5-8: 加购 ----
      return await doAddToBag(addBtn, item, addQty);

    } catch (err) {
      if (err === StopError) throw err;
      log(`  ❌ 失败: ${err.message}`, 'error');
      if (attempt < CONFIG.behavior.retryCount) {
        log(`  🔄 重试中... (${attempt}/${CONFIG.behavior.retryCount})`, 'warn');
        await randomDelay();
        return executeItem(item, attempt + 1, allItems, curIdx);
      }
      return { status: 'failed', item, reason: UNKNOWN_ERR, error: `[Row${_row}|${boxSku}] ${err.message}` };
    }
  }

  // 查找加购按钮（多策略：文本 "add to bag" / class / active 态，主文档+iframe）
  async function findAddToBagButton() {
    // 在指定 document 中查找（主文档和 iframe 共用）
    function searchInDoc(rootDoc) {
      // 策略1: 匹配 "add to bag" 文本，优先按钮类型再兜底 span/div
      // 注意: 传 allowDisabled=true，按钮可能因页面状态暂时禁用，仍需能找到（点击失败时由后续报错体现）
      const _selectBest = (els, exactMatch) => {
        let best = null;
        for (const el of els) {
          const txt = (el.textContent || el.value || '').trim();
          const hits = exactMatch ? txt.toLowerCase() === 'add to bag' : /add\s*to\s*bag/i.test(txt);
          if (hits && isElementReal(el, true)) {
            if (!best || el.children.length < best.children.length) best = el;
          }
        }
        return best;
      };
      // 1a: 只匹配真正的按钮元素（regex 宽松）
      let btn = _selectBest(rootDoc.querySelectorAll('button, a, input[type="submit"]'), false);
      // 1b: 兜底 span/div（精确全字匹配，防止容器 div 误判）
      if (!btn) btn = _selectBest(rootDoc.querySelectorAll('span, div'), true);
      if (btn) return btn;

      // 策略2: class 包含 add/bag/cart 的元素
      const byClass = rootDoc.querySelector('[class*="add-to-bag"], [class*="addToBag"], [class*="add_bag"], [class*="btn-add"], [class*="addtobag"], [class*="addToCart"], [class*="add-to-cart"]');
      if (byClass && isElementReal(byClass, true)) return byClass;

      // 策略3: span.active（选规格后激活的加购按钮）
      const activeSpan = rootDoc.querySelector('span.active');
      if (activeSpan && isElementReal(activeSpan, true)) return activeSpan;

      return null;
    }

    // 主文档
    let btn = searchInDoc(document);
    if (btn) return btn;

    // iframe 兜底（B站商品页有 1 个 iframe）
    const iframes = document.querySelectorAll('iframe');
    for (const iframe of iframes) {
      try {
        btn = searchInDoc(iframe.contentDocument);
        if (btn) return btn;
      } catch (e) { /* cross-origin iframe 跳过 */ }
    }

    return null;
  }

  // 等待成功弹窗出现
  async function waitForSuccessModal() {
    const start = Date.now();
    while (Date.now() - start < CONFIG.behavior.modalWaitTimeout) {
      // 检查页面是否包含成功文本
      const bodyText = document.body.innerText;
      if (bodyText.includes('added to your shopping bag') || bodyText.includes('added to shopping bag')) {
        return true;
      }
      // 也检查是否有遮罩层/弹窗出现
      const overlay = document.querySelector('.md-dialog-container, [role="dialog"], .modal, .overlay, [class*="dialog"]');
      if (overlay && overlay.offsetParent !== null) {
        const overlayText = overlay.innerText || '';
        if (/added.*shopping\s*bag/i.test(overlayText)) return true;
      }
      await sleep(500);
    }
    return false;
  }

  // 检测是否弹出 "Select size or color" 弹窗（商品需人工选择规格）
  function isSelectSizeDialogOpen() {
    const dialogs = document.querySelectorAll('[role="dialog"], .md-dialog-container, .modal, [class*="dialog"]');
    for (const dlg of dialogs) {
      if (dlg.offsetParent === null) continue; // 不可见
      const text = (dlg.innerText || '').toLowerCase();
      if (text.includes('select size') || text.includes('select color')) return true;
    }
    return false;
  }

  // ============================================================
  //  多规格商品加购（Select size or color 弹窗）
  //  流程：点击弹窗内选项 → 后台页面 More Information 表格同步刷新 →
  //        提取 UPC(13位纯数字) 与 Excel 内部参考号比对 → 匹配则加购；
  //        全部选项无匹配 → 标记缺货
  // ============================================================

  // 查找页面上的 "Select size or color" 触发器（多规格商品才有）
  // 实测触发器是 div[ng-click="showProductOptionsDialog()"]（不是 <button>），
  // 已选中规格后其文本变为 "#25 Warm Beige x 6 ▼"，ng-click 保持不变
  function findSelectSizeTrigger() {
    const byNg = document.querySelector(CONFIG.selectors.selectSizeTrigger);
    if (byNg && isElementReal(byNg, true)) return byNg;
    // 兜底：文本匹配（button/div/span 都查）
    const els = document.querySelectorAll('button, div, span, a');
    for (const b of els) {
      if (!isElementReal(b, true)) continue;
      const t = (b.textContent || '').trim();
      if (/^select size or color$/i.test(t)) return b;
      if (/^#\d+/.test(t) && /▼/.test(t)) return b;
    }
    return null;
  }

  // 查找可见的 "Select size or color" 弹窗元素
  function findSelectSizeDialog() {
    const dialogs = document.querySelectorAll(CONFIG.selectors.selectSizeDialog);
    for (const d of dialogs) {
      if (d.offsetParent === null) continue; // 不可见
      const text = (d.innerText || '').toLowerCase();
      if (text.includes('select size') || text.includes('select color')) return d;
    }
    return null;
  }

  // 从 More Information 表格中按行标题取数值（UPC / Catalog No.）
  // ① 优先在包含 "More Information" 标题的表格中查找；② 兜底全局查找首个匹配行
  function extractDetailsRow(labelRe) {
    const tables = document.querySelectorAll('table');
    for (const table of tables) {
      const rows = table.querySelectorAll('tr');
      const hasSection = [...rows].some(r =>
        [...r.querySelectorAll('td,th')].some(c => /^more information$/i.test((c.textContent || '').trim()))
      );
      if (!hasSection) continue;
      for (const row of rows) {
        const cells = [...row.querySelectorAll('td')];
        if (cells.length < 2) continue;
        if (labelRe.test((cells[0].textContent || '').trim())) {
          return (cells[1].textContent || '').trim();
        }
      }
    }
    // 兜底：全局查找首个匹配行标题的 td（无 "More Information" 标题的表格）
    const tds = document.querySelectorAll('td');
    for (const td of tds) {
      if (labelRe.test((td.textContent || '').trim())) {
        const cells = [...td.parentElement.querySelectorAll('td')];
        if (cells.length >= 2) return (cells[1].textContent || '').trim();
      }
    }
    return '';
  }

  // 提取页面 More Information 中的 UPC 值（13位纯数字，兼容 "8800287116256 x 6" 格式）
  function extractUpcFromDetails() {
    const raw = extractDetailsRow(/^upc$/i);
    if (!raw) return null;
    const m = raw.match(/\b\d{13}\b/);
    return { raw, upc13: m ? m[0] : '' };
  }

  // 提取页面 More Information 中的 Catalog No.（用于判断选项点击后页面是否已刷新）
  function extractCatalogNoFromDetails() {
    return extractDetailsRow(/^catalog\s*no\.?$/i);
  }

  // Excel 内部参考号 → 13位 UPC（兼容 "8800287116256 x 6" 或纯13位数字）
  function extractRefUpc(innerRef) {
    const m = String(innerRef || '').match(/\b\d{13}\b/);
    return m ? m[0] : '';
  }

  // 获取规格弹窗内的选项名称列表（只取名称；按钮每次点击前重新定位，见 findDialogOptionBtn）
  function getDialogOptionLabels(dialog) {
    const labels = [];
    const items = dialog.querySelectorAll('md-list-item');
    for (const item of items) {
      const btn = item.querySelector(CONFIG.selectors.dialogOptionBtn);
      if (!btn) continue;
      // 选项名在 .infoCol 里的 span（如 "#25 Warm Beige x 6"），按钮自身文本为空
      const span = [...item.querySelectorAll('span')].find(s => /^#\d+/.test((s.textContent || '').trim()));
      labels.push(span ? span.textContent.trim()
        : (btn.getAttribute('aria-label') || item.textContent || '').trim().split('\n')[0].trim());
    }
    if (labels.length > 0) return labels;
    // 兜底：直接找选项按钮取名称
    return [...dialog.querySelectorAll(CONFIG.selectors.dialogOptionBtn)].map(btn =>
      (btn.textContent || '').trim() || (btn.getAttribute('aria-label') || '').trim().split('\n')[0].trim());
  }

  // 按选项名称在弹窗内重新定位选项按钮（返回新鲜引用）
  // 关键：Angular 选中一个选项后会重渲染弹窗 DOM，之前缓存的按钮引用会失效（isConnected=false），
  // 点击无效；因此每次点击前必须重新定位
  function findDialogOptionBtn(dialog, label) {
    const items = dialog.querySelectorAll('md-list-item');
    for (const item of items) {
      const btn = item.querySelector(CONFIG.selectors.dialogOptionBtn);
      if (!btn) continue;
      const span = [...item.querySelectorAll('span')].find(s => /^#\d+/.test((s.textContent || '').trim()));
      const itemLabel = span ? span.textContent.trim()
        : (btn.getAttribute('aria-label') || item.textContent || '').trim().split('\n')[0].trim();
      if (itemLabel === label) return btn;
    }
    // 兜底：直接匹配按钮文本
    return [...dialog.querySelectorAll(CONFIG.selectors.dialogOptionBtn)].find(btn =>
      (btn.textContent || '').trim() === label
      || ((btn.getAttribute('aria-label') || '').trim().split('\n')[0].trim() === label)) || null;
  }

  // 点击选项后等待后台页面刷新（UPC 或 Catalog No. 与点击前不同）
  async function waitForDetailsRefresh(prevSignal, timeout = 8000) {
    const start = Date.now();
    while (Date.now() - start < timeout) {
      checkStop();
      const upc = extractUpcFromDetails();
      const upc13 = upc ? upc.upc13 : '';
      const catalog = extractCatalogNoFromDetails();
      if (upc13 && upc13 !== prevSignal.upc13) return true;
      if (catalog && catalog !== prevSignal.catalog) return true;
      await sleep(300);
    }
    return false;
  }

  // 规格弹窗加购：依次点击选项 → 比对 UPC → 匹配则加购；全部不匹配 → 缺货
  async function trySelectOptionAdd(item, addQty) {
    const dialog = findSelectSizeDialog();
    if (!dialog) throw new Error('规格弹窗已消失');
    const optionLabels = getDialogOptionLabels(dialog);
    const refUpc = extractRefUpc(item.innerRef);
    log(`  ℹ️ 规格弹窗共 ${optionLabels.length} 个选项 | Excel内部参考号=${item.innerRef || '(空)'}${refUpc ? ` → UPC=${refUpc}` : ''}`, 'info');
    if (optionLabels.length === 0) {
      log('  ⚠️ 弹窗内未找到规格选项，标记缺货', 'warn');
      return { status: 'skipped', item, reason: ERRORS.OUT_OF_STOCK };
    }

    // ① 若当前已有选中规格（历史状态），先直接比对
    const cur = extractUpcFromDetails();
    if (cur && cur.upc13) {
      if (refUpc && cur.upc13 === refUpc) {
        log(`  ✅ 当前规格 UPC=${cur.upc13} 与内部参考号匹配，直接加购`, 'success');
        return await clickDialogAddToBag(item, addQty);
      }
      log(`  ⚠️ 当前规格 UPC=${cur.upc13} 不匹配，继续遍历其他选项`, 'info');
    }

    // ② 依次点击选项，等待后台刷新后读取 UPC 比对
    // 注意：Angular 选中选项后会重渲染弹窗 DOM，按钮引用会失效，
    // 因此每次点击前都用 findDialogOptionBtn 重新定位（实测修复后 8/8 选项均可点击）
    let prevSignal = {
      upc13: cur ? cur.upc13 : '',
      catalog: extractCatalogNoFromDetails() || '',
    };
    for (let i = 0; i < optionLabels.length; i++) {
      const label = optionLabels[i];
      log(`  → 点击选项 ${i + 1}/${optionLabels.length}: ${label}`, 'info');
      let btn = findDialogOptionBtn(dialog, label);
      if (!btn) {
        log(`  ⚠️ ${label}: 未找到选项按钮，尝试下一选项`, 'warn');
        continue;
      }
      btn.click();
      let refreshed = await waitForDetailsRefresh(prevSignal, CONFIG.behavior.optionWaitTimeout);
      if (!refreshed) {
        // 重试一次：重新定位（Angular 可能仍在重渲染）后再点击
        const btn2 = findDialogOptionBtn(dialog, label);
        if (btn2 && btn2 !== btn) {
          log(`  ↻ ${label}: 重新定位按钮后重试点击...`, 'info');
          btn2.click();
          refreshed = await waitForDetailsRefresh(prevSignal, CONFIG.behavior.optionWaitTimeout);
        }
      }
      prevSignal = {
        upc13: (extractUpcFromDetails() || {}).upc13 || '',
        catalog: extractCatalogNoFromDetails() || '',
      };
      if (!refreshed) {
        log(`  ⚠️ ${label}: 页面信息刷新超时，尝试下一选项`, 'warn');
        continue;
      }
      const upc13 = prevSignal.upc13;
      if (!upc13) {
        log(`  ⚠️ ${label}: 未提取到 UPC（More Information 无 UPC 行）`, 'warn');
        continue;
      }
      if (refUpc && upc13 === refUpc) {
        log(`  ✅ ${label}: UPC=${upc13} 匹配！开始加购`, 'success');
        item._selectedOption = label;
        return await clickDialogAddToBag(item, addQty);
      }
      log(`  ❌ ${label}: UPC=${upc13}${refUpc ? ` ≠ 期望 ${refUpc}` : '（Excel无13位UPC，无法匹配）'}`, 'warn');
    }

    // ③ 全部选项无匹配 → 缺货
    log('  🚫 所有规格选项均已点击且无 UPC 匹配，标记缺货', 'warn');
    return { status: 'skipped', item, reason: ERRORS.OUT_OF_STOCK };
  }

  // 点击规格弹窗内的 ADD TO BAG 并校验成功弹窗
  async function clickDialogAddToBag(item, addQty) {
    const dialog = findSelectSizeDialog();
    if (!dialog) throw new Error('规格弹窗已消失');
    // 关键：数量绑定在 selectedProductOption.quantity 上，切换选项会替换对象导致数量重置为 1，
    // 因此必须「先匹配到目标选项、再重设数量」，最后才点弹窗 ADD TO BAG
    log(`  → 重设数量: ${addQty}`, 'info');
    await setQty(addQty);
    const btn = dialog.querySelector(CONFIG.selectors.dialogAddToBagBtn)
      || [...dialog.querySelectorAll('button')].find(b => /add\s*to\s*b[ao]g/i.test((b.textContent || '').trim()));
    if (!btn) throw new Error('规格弹窗内未找到 ADD TO BAG 按钮');
    log('  → 点击弹窗内 ADD TO BAG...', 'info');
    btn.click();
    await sleep(1000);
    const success = await waitForSuccessModal();
    if (!success) throw new Error('超时未检测到加购成功弹窗');
    await closeModal();
    await randomDelay();
    log('  ✅ 加购成功!', 'success');
    return { status: 'success', item };
  }

  // 关闭弹窗
  async function closeModal() {
    // 查找 CLOSE 按钮
    const btns = document.querySelectorAll('button, span, a, div');
    for (const btn of btns) {
      const txt = (btn.textContent || '').trim().toUpperCase();
      if (txt === 'CLOSE' || txt === '关闭') {
        btn.click();
        log(`  → 已关闭弹窗`, 'info');
        await sleep(500);
        return;
      }
    }
    // 备选: 按 ESC 键
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', code: 'Escape', bubbles: true }));
    await sleep(300);
  }

  // 执行全部任务
  async function runAll(items) {
    setRunning(true);
    forceStop = false;
    const _startedAt = new Date().toISOString();
    // 全批次任务总数（进度条分母）；断点续跑/继续模式时 items 是剩余行，分母仍是全批次总数
    const total = STORE.get(K.totalCount, items.length);

    log(`🚀 开始执行，共 ${items.length} 条任务`, 'action');
    updateProgress(results.length, total);
    updateBatchInfo(); // 继续模式时已有历史结果，header 同步

    let stoppedByUser = false;
    let crashed = false;
    let navigating = false;
    let storageFull = false; // 存储空间不足导致停止（已加购结果需抢救）
    try {
      for (let i = 0; i < items.length; i++) {
        if (forceStop) {
          stoppedByUser = true;
          break;
        }

        let result;
        try {
          result = await executeItem(items[i], 1, items, i);
        } catch (err) {
          // 用户点击「立即停止」：直接终止，不保存状态
          if (err === StopError) {
            log('⏹️ 已立即停止（当前条目已中断）', 'warn');
            stoppedByUser = true;
            break;
          }
          throw err;
        }
        results.push(result);
        if (!results._startedAt) results._startedAt = _startedAt;
        // 实时落盘：每完成一条立即持久化，刷新后据此从断点恢复（方案成立的地基）
        // 保存失败（空间不足）时自动腾空间重试；仍失败则停止并抢救已加购结果
        if (!STORE.setWithRetry(K.results, results)) {
          storageFull = true;
          log('⛔ 存储空间不足：结果保存失败，已停止任务！已加购结果已尽力保存。', 'error');
          log('   请点击「下载结果Excel」保存进度 → 重新上传该文件即可断点续跑（已加购行自动跳过，无需清空购物车）', 'error');
          break;
        }
        updateProgress(results.length, total);
        updateBatchInfo();                       // 批次 header 实时刷新（已完成 N / 总 M）
        renderTaskList(items.slice(i + 1));      // 已完成的从任务列表移除（剩余待办）

        // 如果是导航状态（页面跳转），保存剩余任务，新页面自动续跑
        if (result.status === 'navigating') {
          // navigating 条目不进入 results：任务会在新页面重新执行并产生一条完成结果，
          // 若此时 push 会导致同一任务两条结果 → "已完成 N/总 M" 虚高（如 14/7）
          results.pop();
          STORE.set(K.results, results);
          updateProgress(results.length, total);
          updateBatchInfo();                     // 跳页前同步 header（navigating 结果不计入已完成）
          renderTaskList(items.slice(i + 1));    // 剩余任务列表随新页面续跑展示
          // 续跑点保存失败（空间不足）→ 停止并抢救，避免跳转后无法续跑（表现为任务莫名中断）
          if (!STORE.setWithRetry(K.pending, items.slice(i))) {
            storageFull = true;
            log('⛔ 存储空间不足：剩余任务保存失败，已停止任务！请点击「下载结果Excel」→ 重新上传该文件断点续跑（无需清空购物车）', 'error');
            break;
          }
          STORE.set(K.results, results);
          navigating = true;
          break;
        }

        // 条目间延迟
        if (i < items.length - 1) {
          await randomDelay();
        }
      }
    } catch (err) {
      // 引擎级异常：绝不能静默消失（否则任务卡死、按钮状态不恢复）
      crashed = true;
      log(`❌ 执行引擎异常，任务已终止: ${err.message}`, 'error');
      console.error('[ABW] runAll crashed', err);
    } finally {
      // 导航跳转 → 保持运行态（新页面自动续跑并显示「立即停止」）
      if (navigating) {
        setRunning(true);
        return;
      }
      // 任务结束/停止/异常 → 恢复为「开始加购」态
      setRunning(false);

      if (storageFull) {
        // 空间不足停止：已加购结果已尽力落盘，展示下载按钮，引导用户下载→重传→续跑
        document.getElementById('abw-dl-btn').style.display = '';
        const logEntries = STORE.get(K.logs, []);
        document.getElementById('abw-log-btn').style.display = logEntries.length > 0 ? '' : 'none';
        showSummary(results);
        log('⛔ 任务已停止（存储空间不足）', 'error');
        log('下一步操作：', 'info');
        log('  ① 点击「下载结果Excel」→ 保存已加购进度（已加购行会标结果）', 'info');
        log('  ② 重新上传刚下载的文件 → 脚本自动跳过已加购行，只处理剩余商品', 'info');
        log('  ③ 无需清空购物车，已加购的商品不会重复加购', 'info');
        notifyComplete();
      } else if (stoppedByUser) {
        log('⏹️ 已停止', 'warn');
        notifyComplete();
      } else if (crashed) {
        document.getElementById('abw-dl-btn').style.display = '';
        showSummary(results);
        log('⚠️ 任务被异常中断，请下载开发日志排查', 'error');
      } else {
        document.getElementById('abw-dl-btn').style.display = '';
        const logEntries = STORE.get(K.logs, []);
        document.getElementById('abw-log-btn').style.display = logEntries.length > 0 ? '' : 'none';
        showSummary(results);
        const sc = results.filter(r => r.status === 'success').length;
        const fa = results.filter(r => r.status === 'failed').length;
        const sk = results.filter(r => r.status === 'skipped').length;
        log(`✅ 全部完成! 成功:${sc} 失败:${fa}${sk > 0 ? ' 缺货:' + sk : ''}`, 'success');
        notifyComplete();
      }
    }
  }

  // ============================================================
  //  浮动面板 UI
  // ============================================================
  // 更新面板上的批次信息显示（文件名 + 已完成/总）
  function updateBatchInfo() {
    const el = document.getElementById('abw-batch-info');
    if (!el) return;
    const fileName = STORE.get(K.fileName, '');
    const total = STORE.get(K.totalCount, 0);
    if (!fileName && !total) { el.style.display = 'none'; return; }
    el.style.display = '';
    el.textContent = `📄 ${fileName || '(未知文件)'} · 已完成 ${results.length} / 总 ${total}`;
  }

  function createPanel() {
    const panel = document.createElement('div');
    panel.id = 'abw-panel';
    panel.innerHTML = `
      <style>
        #abw-panel {
          position: fixed; top: 60px; right: 20px; width: 380px; max-height: 80vh;
          background: #1a1a2e; border-radius: 16px; box-shadow: 0 8px 32px rgba(0,0,0,0.4), 0 0 0 1px rgba(255,255,255,0.08);
          font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; z-index: 999999;
          display: flex; flex-direction: column; overflow: hidden;
        }
        #abw-header {
          padding: 14px 18px; background: linear-gradient(135deg, #16213e, #0f3460);
          cursor: move; user-select: none; display: flex; justify-content: space-between; align-items: center;
          border-bottom: 1px solid rgba(255,255,255,0.06);
        }
        #abw-title { color: #eee; font-size: 14px; font-weight: 600; margin: 0; }
        #abw-title span { color: #e94560; }
        #abw-toggle { color: #888; cursor: pointer; font-size: 18px; line-height: 1; }
        #abw-body { padding: 16px; overflow-y: auto; flex: 1; }
        #abw-body.collapsed { display: none; }
        .abw-section { margin-bottom: 14px; }
        .abw-label { color: #aaa; font-size: 11px; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 6px; }
        #abw-batch-info { color: #e94560; font-size: 12px; margin: 0 0 10px; padding: 6px 10px; background: rgba(233,69,96,0.1); border: 1px solid rgba(233,69,96,0.25); border-radius: 6px; }
        .abw-file-wrap {
          position: relative; border: 2px dashed rgba(233,69,96,0.3); border-radius: 10px;
          padding: 20px; text-align: center; transition: all 0.3s; cursor: pointer;
          background: rgba(233,69,96,0.03);
        }
        .abw-file-wrap:hover { border-color: #e94560; background: rgba(233,69,96,0.08); }
        .abw-file-wrap.has-file { border-color: #2ecc71; border-style: solid; background: rgba(46,204,113,0.05); }
        #abw-file-input { position: absolute; inset: 0; opacity: 0; cursor: pointer; }
        .abw-file-text { color: #888; font-size: 13px; }
        .abw-file-name { color: #2ecc71; font-size: 13px; font-weight: 500; word-break: break-all; }
        .abw-task-list { max-height: 200px; overflow-y: auto; border-radius: 8px; background: rgba(0,0,0,0.2); }
        .abw-task-item { padding: 8px 12px; border-bottom: 1px solid rgba(255,255,255,0.04); font-size: 12px; color: #ccc; display: flex; justify-content: space-between; align-items: center; }
        .abw-task-item:last-child { border-bottom: none; }
        .abw-task-item .sku { color: #e94560; font-weight: 500; }
        .abw-task-item .spec { color: #f39c12; }
        .abw-task-item .qty { color: #3498db; }
        .abw-btn-row { display: flex; gap: 8px; margin-top: 12px; }
        .abw-btn {
          flex: 1; padding: 10px 16px; border: none; border-radius: 8px; font-size: 13px; font-weight: 600;
          cursor: pointer; transition: all 0.2s; text-align: center;
        }
        .abw-btn-primary { background: linear-gradient(135deg, #e94560, #c0392b); color: white; }
        .abw-btn-primary:hover:not(:disabled) { transform: translateY(-1px); box-shadow: 0 4px 12px rgba(233,69,96,0.4); }
        .abw-btn-primary:disabled { opacity: 0.4; cursor: not-allowed; }
        .abw-btn-stop { background: rgba(231,76,60,0.9); color: white; border: none; box-shadow: 0 4px 12px rgba(231,76,60,0.4); }
        .abw-btn-stop:hover { background: #c0392b; transform: translateY(-1px); }
        .abw-btn-sec { background: rgba(255,255,255,0.06); color: #aaa; border: 1px solid rgba(255,255,255,0.1); }
        .abw-btn-sec:hover { background: rgba(255,255,255,0.1); color: #ddd; }
        #abw-progress { height: 4px; background: rgba(255,255,255,0.06); border-radius: 2px; overflow: hidden; margin-top: 10px; }
        #abw-progress-bar { height: 100%; background: linear-gradient(90deg, #e94560, #f39c12); border-radius: 2px; transition: width 0.3s; width: 0%; }
        #abw-log { max-height: 180px; overflow-y: auto; font-family: 'Cascadia Code', 'Consolas', monospace;
                  background: rgba(0,0,0,0.3); border-radius: 8px; padding: 10px; margin-top: 10px;
                  font-size: 11.5px; line-height: 1.6; color: #bbb; }
        #abw-summary { display: none; padding: 12px; border-radius: 8px; margin-top: 10px; font-size: 12px; }
        #abw-summary.show { display: block; }
        .abw-stat { display: flex; justify-content: space-between; padding: 4px 0; }
        .abw-stat-success { color: #2ecc71; }
        .abw-stat-fail { color: #e74c3c; }
        .abw-stat-total { color: #f39c12; }
        #abw-footer { padding: 8px 16px; border-top: 1px solid rgba(255,255,255,0.06); font-size: 10px; color: #555; text-align: center; }
      </style>
      <div id="abw-header">
        <h1 id="abw-title">🛒 ABW <span>Auto</span> Purchase</h1>
        <span id="abw-toggle">−</span>
      </div>
      <div id="abw-body">
        <div id="abw-batch-info" style="display:none;"></div>
        <!-- 文件上传 -->
        <div class="abw-section">
          <div class="abw-label">📁 采购单 Excel</div>
          <div class="abw-file-wrap" id="abw-file-wrap">
            <input type="file" id="abw-file-input" accept=".xlsx,.xls,.csv">
            <div class="abw-file-text" id="abw-file-text">点击或拖拽上传 Excel 文件</div>
          </div>
        </div>

        <!-- 任务列表 -->
        <div class="abw-section" id="abw-task-section" style="display:none;">
          <div class="abw-label">📋 任务列表 (<span id="abw-task-count">0</span> 条)</div>
          <div class="abw-task-list" id="abw-task-list"></div>
          <div id="abw-progress"><div id="abw-progress-bar"></div></div>
        </div>

        <!-- 控制按钮 -->
        <div class="abw-btn-row">
          <button class="abw-btn abw-btn-primary" id="abw-start-btn" disabled>▶ 开始加购</button>
          <button class="abw-btn abw-btn-stop" id="abw-stop-btn" style="display:none;">⏹ 立即停止</button>
        </div>
        <div class="abw-btn-row">
          <button class="abw-btn abw-btn-sec" id="abw-clear-btn" disabled>🗑 清空日志</button>
        </div>
        <div class="abw-btn-row">
          <button class="abw-btn abw-btn-sec" id="abw-dl-btn">📥 下载结果 Excel</button>
        </div>
        <div class="abw-btn-row">
          <button class="abw-btn abw-btn-sec" id="abw-log-btn" style="display:none;">📋 下载开发日志</button>
        </div>

        <!-- 日志 -->
        <div class="abw-section">
          <div class="abw-label">📝 运行日志</div>
          <div id="abw-log"></div>
        </div>

        <!-- 汇总 -->
        <div id="abw-summary"></div>
      </div>
      <div id="abw-footer">ABW Auto Purchase v1.11.4 · 数据不出浏览器 · 跨页面保持</div>
    `;
    document.body.appendChild(panel);

    // 恢复面板位置（跨页面保持）
    const savedPos = STORE.get(K.pos, null);
    if (savedPos) {
      panel.style.left = savedPos.left + 'px';
      panel.style.top = savedPos.top + 'px';
      panel.style.right = 'auto';
    }

    // 恢复折叠状态
    const savedCollapsed = STORE.get(K.collapsed, false);
    if (savedCollapsed) {
      document.getElementById('abw-body').classList.add('collapsed');
      document.getElementById('abw-toggle').textContent = '+';
    }

    // 恢复历史日志
    restoreLogs();

    // 拖拽功能
    makeDraggable(panel, document.getElementById('abw-header'));

    // 折叠功能（状态跨页面保持）
    document.getElementById('abw-toggle').addEventListener('click', () => {
      const body = document.getElementById('abw-body');
      const toggle = document.getElementById('abw-toggle');
      body.classList.toggle('collapsed');
      toggle.textContent = body.classList.contains('collapsed') ? '+' : '−';
      STORE.set(K.collapsed, body.classList.contains('collapsed'));
    });

    // 文件上传
    const fileInput = document.getElementById('abw-file-input');
    fileInput.addEventListener('change', handleFileUpload);

    // 拖拽上传
    const wrap = document.getElementById('abw-file-wrap');
    wrap.addEventListener('dragover', e => { e.preventDefault(); wrap.style.borderColor = '#e94560'; });
    wrap.addEventListener('dragleave', () => { wrap.style.borderColor = ''; });
    wrap.addEventListener('drop', e => {
      e.preventDefault();
      wrap.style.borderColor = '';
      if (e.dataTransfer.files.length > 0) {
        fileInput.files = e.dataTransfer.files;
        handleFileUpload({ target: fileInput });
      }
    });

    // 按钮
    document.getElementById('abw-start-btn').addEventListener('click', () => {
      const items = STORE.get(K.items, []);
      if (items.length === 0) return;
      // 进入运行态：显示「立即停止」、隐藏「开始加购」（下载按钮常驻，执行中随时可导出进度）
      setRunning(true);
      document.getElementById('abw-log-btn').style.display = 'none';

      // 区分「继续」与「全新」：有已完成结果 → 从剩余行继续，不清空 results
      const savedResults = STORE.get(K.results, []);
      let toRun = items;
      if (savedResults.length > 0) {
        const doneRows = new Set();
        savedResults.forEach(r => resultRows(r).forEach(n => doneRows.add(n)));
        toRun = items.filter(it => !doneRows.has(it._row));
        if (toRun.length === 0) {
          log('✅ 所有任务均已完成，无需再跑', 'success');
          renderTaskList([]); // 全部完成 → 任务列表隐藏
          setRunning(false);
          return;
        }
        log(`▶️ 从断点继续：跳过已完成 ${savedResults.length} 条，剩余 ${toRun.length} 条`, 'action');
      } else {
        results.length = 0; // 全新执行
      }
      runAll(toRun);
    });
    // 「立即停止」：点击直接中断当前任务，不记录状态
    document.getElementById('abw-stop-btn').addEventListener('click', () => {
      forceStop = true;
      log('⏹️ 正在立即停止...', 'warn');
    });
    document.getElementById('abw-clear-btn').addEventListener('click', () => {
      document.getElementById('abw-log').innerHTML = '';
      STORE.remove(K.logs);
      results.length = 0;
      STORE.remove(K.results); // 同步清空落盘结果，避免刷新后恢复旧进度
      updateProgress(0, STORE.get(K.totalCount, 0));
      updateBatchInfo();
    });
    // 下载结果 Excel
    document.getElementById('abw-dl-btn').addEventListener('click', downloadResultExcel);
    // 开发日志下载按钮
    document.getElementById('abw-log-btn').addEventListener('click', downloadDevLog);


    // ===== 批次恢复（按优先级）=====
    const savedItems = STORE.get(K.items, []);
    const pendingItems = STORE.get(K.pending, null);
    const savedResults = STORE.get(K.results, []);
    const totalCount = STORE.get(K.totalCount, savedItems.length);

    if (pendingItems && pendingItems.length > 0) {
      // ① 跳页中断：自动续跑（串行导航必需机制，保持不动）→ 进入运行态，显示「立即停止」
      STORE.remove(K.pending);
      results.push(...dedupeResults(savedResults));
      renderTaskList(pendingItems);
      document.getElementById('abw-start-btn').disabled = false;
      document.getElementById('abw-clear-btn').disabled = false;
      setRunning(true);
      updateBatchInfo();
      setTimeout(() => runAll(pendingItems), 2000);
    } else {
      // 刷新/全新 → 停止态，显示「开始加购」
      setRunning(false);
      if (savedResults.length > 0 && savedItems.length > 0) {
        // ② 同页执行中被刷新/中断：恢复结果 + 剩余列表 + 进度，手动开始
        // 先对存量结果去重（兼容旧版本 navigating 重复条目导致的虚高，如 14/7）
        const restoredResults = dedupeResults(savedResults);
        results.push(...restoredResults);
        const doneRows = new Set();
        restoredResults.forEach(r => resultRows(r).forEach(n => doneRows.add(n)));
        const remaining = savedItems.filter(it => !doneRows.has(it._row));
        updateProgress(restoredResults.length, totalCount);
        updateBatchInfo();
        if (remaining.length === 0) {
          // 批次已全部完成：只恢复结果与进度，不提示「未完成批次」
          renderTaskList([]); // 空列表 → 任务列表隐藏
          document.getElementById('abw-start-btn').disabled = true;
          document.getElementById('abw-clear-btn').disabled = false;
          log(`✅ 检测到已完成批次：已完成 ${restoredResults.length}/${totalCount}，无剩余任务`, 'success');
        } else {
          renderTaskList(remaining);
          document.getElementById('abw-start-btn').disabled = false;
          document.getElementById('abw-clear-btn').disabled = false;
          log(`⏸️ 检测到未完成批次：已完成 ${restoredResults.length}/${totalCount}，剩余 ${remaining.length} 条待加购（点击「开始加购」继续）`, 'warn');
        }
      } else {
        // ③ 全新 / 无未完成任务
        if (savedItems.length > 0) {
          renderTaskList(savedItems);
          document.getElementById('abw-start-btn').disabled = false;
          document.getElementById('abw-clear-btn').disabled = false;
        }
        updateBatchInfo();
      }
    }
  }

  // 下载结果 Excel（原文件 + 加购结果列 + ABW 商品名）
  function downloadResultExcel() {
    const headers = STORE.get(K.xlHeaders, []);
    const rawRows = STORE.get(K.xlRows, []);
    if (!headers.length || !rawRows.length) { log('⚠️ 无原始数据，请重新上传 Excel', 'warn'); return; }

    // 按原始行号建结果映射（一个聚合任务 → 组内所有原始行标相同结果）
    const statusMap = {};
    const reasonMap = {};
    const shipMap = {};
    const nameMap = {};
    const unitPriceMap = {};
    const boxPriceMap = {};
    for (const r of results) {
      for (const rowNo of resultRows(r)) {
        statusMap[rowNo] = r.status;
        if (r.reason) reasonMap[rowNo] = r.reason;
        if (r.error) reasonMap[rowNo] = reasonMap[rowNo] || r.error;
        if (r.item._shipWarning) shipMap[rowNo] = r.item._shipWarning;
        if (r.item._abwProductName) nameMap[rowNo] = r.item._abwProductName;
        if (r.item._unitPrice) unitPriceMap[rowNo] = r.item._unitPrice;
        if (r.item._boxPrice) boxPriceMap[rowNo] = r.item._boxPrice;
      }
    }

    // 定位需回写的原始列：备注、整箱批发价（按表头名匹配，忽略大小写/分隔符）
    const cleanHeader = (h) => String(h || '').replace(/[^a-zA-Z0-9\u4e00-\u9fa5]/g, '').toLowerCase();
    const findColIdx = (names) => {
      const cleanNames = names.map(cleanHeader);
      return headers.findIndex(h => cleanNames.includes(cleanHeader(h)));
    };
    const remarkIdx = findColIdx(['订单行/备注', '备注']);
    const boxPriceColIdx = findColIdx(['订单行/整箱批发价', '整箱批发价']);
    // 派生"销售平台-区域"所需的列
    const shopIdx = findColIdx(['订单行/店铺']);
    const platformIdx = findColIdx(['订单行/店铺/电商平台']);
    const packagingColIdx = findColIdx(['订单行/包装']);
    // 订单关联列（结果需向下填充，与副本一致）
    const orderRelIdx = findColIdx(['订单关联']);

    // 表头："订单行/店铺/电商平台" 改名为 "销售平台-区域"
    const outHeaders = headers.slice();
    if (platformIdx >= 0) outHeaders[platformIdx] = '销售平台-区域';

    // 附加列：文件里已有该列则复用其位置（支持"结果文件再导入→再导出"往返），否则追加
    const extraColDefs = ['ABWproductname', '品牌名', '单价', '整箱批发价', '加购结果'];
    const extraColIdx = {};
    for (const name of extraColDefs) {
      let idx = findColIdx([name]);
      if (idx < 0) { idx = outHeaders.length; outHeaders.push(name); }
      extraColIdx[name] = idx;
    }
    // 导入文件里原有的"加购结果"列索引（未执行的行保留旧结果值）
    const prevResultIdx = findColIdx(['加购结果']);

    const outRows = [outHeaders.slice()];
    let currentOrderRel = ''; // 订单关联向下填充的当前值
    for (let i = 1; i < rawRows.length; i++) {
      const row = rawRows[i];
      if (!row || row.every(c => c == null || String(c || '').trim() === '')) continue;
      const st = statusMap[i + 1];
      const reason = reasonMap[i + 1] || '';
      const productName = nameMap[i + 1] || '';
      // 从 ABWproductname 提取品牌名（" - " 前面的部分）
      const brandName = productName ? productName.split(' - ')[0].trim() : '';
      const unitPrice = unitPriceMap[i + 1] || '';
      const boxPrice = boxPriceMap[i + 1] || '';
      // 结果标签：本次执行过 → 覆盖；未执行 → 保留导入文件里的旧结果（无则留空）
      const prevLabel = prevResultIdx >= 0 ? String(row[prevResultIdx] || '').trim() : '';
      let label = prevLabel;
      if (st === 'success') {
        label = shipMap[i + 1] ? `⏳成功(可能缺货,${shipMap[i + 1]})` : '✅成功';
      } else if (st === 'skipped') {
        label = '🚫缺货';
      } else if (st === 'failed') {
        label = `❌${reason || '失败'}`;
      }
      // st 为空（未执行）→ 保持 prevLabel（断点续跑的关键：不把未跑的行误标为失败）

      // 回写原始列：缺货/可能缺货 → 备注；有整箱批发价 → 覆盖整箱批发价列
      const outRow = row.slice();
      // 订单关联 向下填充（非空值向下传播，直到下一个非空值）
      if (orderRelIdx >= 0) {
        const rawRel = String(row[orderRelIdx] || '').trim();
        if (rawRel) currentOrderRel = rawRel;
        outRow[orderRelIdx] = currentOrderRel;
      }
      // 派生"销售平台-区域"：电商平台 + '-' + 店铺首词（piece 且首词为 MX 时改 BR）
      if (platformIdx >= 0) {
        const shopName = shopIdx >= 0 ? String(row[shopIdx] || '').trim() : '';
        const platform = String(row[platformIdx] || '').trim();
        const packagingVal = packagingColIdx >= 0 ? String(row[packagingColIdx] || '').trim() : '';
        let region = shopName.split(/\s+/)[0] || '';
        if (!/box/i.test(packagingVal) && region.toUpperCase() === 'MX') region = 'BR';
        outRow[platformIdx] = [platform, region].filter(Boolean).join('-');
      }
      if (st === 'skipped') {
        if (remarkIdx >= 0) outRow[remarkIdx] = '缺货，请帮忙查询是否有库存供下单';
      } else if (st === 'success' && shipMap[i + 1]) {
        if (remarkIdx >= 0) outRow[remarkIdx] = '可能缺货';
      }
      if (boxPrice && boxPriceColIdx >= 0) outRow[boxPriceColIdx] = parseFloat(boxPrice);

      // 写附加列（复用已有列位置，行不够长则补齐）
      while (outRow.length < outHeaders.length) outRow.push('');
      const extraVals = {
        'ABWproductname': productName,
        '品牌名': brandName,
        '单价': unitPrice,
        '整箱批发价': boxPrice,
        '加购结果': label,
      };
      for (const name of extraColDefs) outRow[extraColIdx[name]] = extraVals[name];
      outRows.push(outRow);
    }

    // 按 ABWproductname 列 A-Z 排序
    const nameIdx = extraColIdx['ABWproductname'];
    const headerRow = outRows.shift();
    outRows.sort((a, b) => {
      const va = String(a[nameIdx] || '').trim().toLowerCase();
      const vb = String(b[nameIdx] || '').trim().toLowerCase();
      return va < vb ? -1 : va > vb ? 1 : 0;
    });
    outRows.unshift(headerRow);

    const ws = XLSX.utils.aoa_to_sheet(outRows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '采购结果');
    const buf = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([buf], { type: 'application/octet-stream' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    const now = new Date();
    const pad = n => String(n).padStart(2, '0');
    a.download = `采购结果_${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}_${pad(now.getHours())}${pad(now.getMinutes())}.xlsx`;
    document.body.appendChild(a); a.click();
    document.body.removeChild(a); URL.revokeObjectURL(url);
    log('📥 结果 Excel 已下载（未执行行结果留空，可再次上传断点续跑）', 'success');
  }

  // 浏览器系统通知 + 标题闪烁
  function notifyComplete() {
    // 确保通知权限已授权
    if ('Notification' in window && Notification.permission !== 'granted' && Notification.permission !== 'denied') {
      Notification.requestPermission();
    }
    // 1) 系统通知（窗口后台时任务栏闪烁）
    if ('Notification' in window && Notification.permission === 'granted') {
      const sc = results.filter(r => r.status === 'success').length;
      const fa = results.filter(r => r.status === 'failed').length;
      const sk = results.filter(r => r.status === 'skipped').length;
      new Notification('🛒 ABW 任务已结束', {
        body: `成功 ${sc} · 失败 ${fa} · 缺货 ${sk}`,
        icon: 'favicon-32x32.png',
        requireInteraction: true
      });
    }
    // 2) 标题闪烁（前台也能看到）
    const origTitle = document.title;
    let blinks = 0;
    const timer = setInterval(() => {
      document.title = blinks % 2 === 0 ? '--- ABW 任务已结束 ---' : origTitle;
      if (++blinks >= 10) { clearInterval(timer); document.title = origTitle; }
    }, 400);
    // 3) 提示音
    try {
      const ctx = new (window.AudioContext || window.webkitAudioContext)();
      const osc = ctx.createOscillator();
      const gain = ctx.createGain();
      osc.connect(gain); gain.connect(ctx.destination);
      osc.type = 'sine';
      osc.frequency.setValueAtTime(880, ctx.currentTime);
      osc.frequency.setValueAtTime(660, ctx.currentTime + 0.1);
      gain.gain.setValueAtTime(0.3, ctx.currentTime);
      gain.gain.exponentialRampToValueAtTime(0.01, ctx.currentTime + 0.3);
      osc.start(ctx.currentTime + 0.1);
      osc.stop(ctx.currentTime + 0.3);
    } catch(e) { /* 音频不可用时静默 */ }
  }

  // 下载开发日志（任务结束自动触发）
  function downloadDevLog() {
    const batchNo = STORE.get(K.batchNo, '') || 'unknown';
    const now = new Date();
    const dateStr = now.toISOString().slice(0, 10);
    const timeStr = now.toISOString().slice(11, 19).replace(/:/g, '');
    const filename = `ABW_logs/abw_log_${dateStr}_${batchNo}_${timeStr}.json`;

    const logData = {
      batchNo,
      startedAt: results[0]?._startedAt || '',
      finishedAt: now.toISOString(),
      totalItems: results.length,
      summary: {
        success: results.filter(r => r.status === 'success').length,
        failed: results.filter(r => r.status === 'failed').length,
        skipped: results.filter(r => r.status === 'skipped').length,
      },
      items: results.map(r => ({
        row: r.item._row,
        rows: r.item._rows || [r.item._row],
        boxSku: r.item.boxSku,
        packaging: r.item.packaging,
        quantity: r.item.quantity,
        abwProductName: r.item._abwProductName || '',
        unitPrice: r.item._unitPrice || '',
        boxPrice: r.item._boxPrice || '',
        status: r.status,
        reason: r.reason || '',
        error: r.error || '',
      })),
      logs: STORE.get(K.logs, []),
    };

    const blob = new Blob([JSON.stringify(logData, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    // 扩展用 chrome.downloads 支持子目录；回落 <a> 兜底
    if (typeof chrome !== 'undefined' && chrome.downloads) {
      chrome.downloads.download({ url, filename, conflictAction: 'overwrite' });
    } else {
      const a = document.createElement('a');
      a.href = url; a.download = filename; document.body.appendChild(a); a.click();
      document.body.removeChild(a);
    }
    setTimeout(() => URL.revokeObjectURL(url), 30000);
    log('📋 开发日志已保存', 'info');
  }

  function makeDraggable(el, handle) {
    let offsetX = 0, offsetY = 0, isDragging = false;
    handle.addEventListener('mousedown', e => {
      isDragging = true;
      offsetX = e.clientX - el.offsetLeft;
      offsetY = e.clientY - el.offsetTop;
      handle.style.cursor = 'grabbing';
    });
    document.addEventListener('mousemove', e => {
      if (!isDragging) return;
      el.style.left = (e.clientX - offsetX) + 'px';
      el.style.top = (e.clientY - offsetY) + 'px';
      el.style.right = 'auto';
    });
    document.addEventListener('mouseup', () => {
      if (isDragging) {
        // 保存面板位置（跨页面保持）
        STORE.set(K.pos, { left: el.offsetLeft, top: el.offsetTop });
      }
      isDragging = false;
      handle.style.cursor = 'move';
    });
  }

  // 处理文件上传
  async function handleFileUpload(e) {
    const file = e.target.files[0];
    if (!file) return;

    // 开启新批次（总任务标识）：清旧批次 → 切新 runId → 存文件名，保证后续日志都写进新批次
    startNewBatch(file.name);

    const wrap = document.getElementById('abw-file-wrap');
    const textEl = document.getElementById('abw-file-text');

    wrap.classList.add('has-file');
    textEl.className = 'abw-file-name';
    textEl.textContent = `✅ ${file.name} (${(file.size / 1024).toFixed(1)} KB)`;

    log(`正在解析 Excel: ${file.name}...`, 'info');

    try {
      const buffer = await readFileAsArrayBuffer(file);
      const parsed = parseExcel(buffer);
      const rawItems = parsed.items;

      // 断点续跑且全部已处理 → 不报错，提示无需再跑
      const isAllDone = rawItems.length === 0
        && parsed.resumeStats && parsed.resumeStats.hadResultCol
        && parsed.resumeStats.doneCount > 0;
      if (rawItems.length === 0 && !isAllDone) {
        throw new Error('未能解析出有效数据行，请确认列名是否匹配');
      }
      if (isAllDone) {
        log('✅ 所有商品均已有加购结果，无需再跑（可直接下载结果 Excel）', 'success');
      }

      // 显示列映射，方便核对
      log(`📑 列映射: ${parsed.colMapLog.join(' | ')}`, 'info');

      // 断点续跑提示：文件带"加购结果"列 → 已处理行已跳过
      if (parsed.resumeStats && parsed.resumeStats.hadResultCol) {
        const rs = parsed.resumeStats;
        log(`🔁 断点续跑: 检测到"加购结果"列，已跳过已处理的 ${rs.doneCount} 行`, 'action');
        log(`    ├ 成功/已在购物车: ${rs.successCount}  |  缺货: ${rs.skipCount}  |  失败: ${rs.failCount}`, 'info');
        log(`    └ 本次仅加购 ${rawItems.length} 条未处理商品${rs.failCount > 0 ? '（如需重试失败行，请清空对应行的加购结果单元格）' : ''}`, 'info');
      }

      // 同 SKU 聚合：同一 SKU+包装 的多行（不同订单关联）合并为一次加购，数量求和
      // 聚合只发生在执行层；下载结果时仍按原始行逐行回写（行数/列结构不变）
      const items = mergeSameSkuItems(rawItems);

      // 存储解析结果 + 原始数据（下载结果时用）
      // 旧批次日志已由 startNewBatch 清除，这里只清面板 DOM
      document.getElementById('abw-log').innerHTML = '';
      results.length = 0;
      STORE.set(K.items, items);
      STORE.set(K.xlHeaders, parsed.headers);
      const xlRowsSaved = STORE.setWithRetry(K.xlRows, parsed.rawRows);
      STORE.set(K.batchNo, parsed.batchNo || '');
      STORE.set(K.totalCount, items.length); // 本批次任务总数（进度条分母）

      // 原始行数据保存失败 → 下载功能将不可用，必须在此拦截，避免"加购完才发现下载不了"
      if (!xlRowsSaved) {
        document.getElementById('abw-start-btn').disabled = true;
        log('⛔ 原始Excel数据保存失败（存储空间不足）：已阻止加购。请先点击「下载日志」，再清理浏览器存储（或关闭并重开标签页）后重新上传', 'error');
        throw new Error('存储空间不足，原始Excel数据保存失败；请清理浏览器存储（或关闭并重开标签页）后重新上传');
      }

      // 显示任务列表
      renderTaskList(items);

      // 有待跑任务才启用开始按钮（全部已处理时禁用）
      document.getElementById('abw-start-btn').disabled = items.length === 0;
      document.getElementById('abw-clear-btn').disabled = false;

      // 更新批次信息显示
      updateBatchInfo();

      if (items.length > 0) {
        log(`解析完成！共 ${items.length} 条有效任务`, 'success');
        log(`首条示例: SKU=${items[0].boxSku} | 规格=${items[0].packaging} | 数量=${items[0].quantity}`, 'info');
      }

    } catch (err) {
      log(`解析失败: ${err.message}`, 'error');
      wrap.classList.remove('has-file');
      textEl.className = 'abw-file-text';
      textEl.textContent = `❌ ${err.message} — 重新上传`;
      textEl.style.color = '#e74c3c';
    }
  }

  function readFileAsArrayBuffer(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = e => resolve(new Uint8Array(e.target.result));
      reader.onerror = reject;
      reader.readAsArrayBuffer(file);
    });
  }

  function renderTaskList(items) {
    const section = document.getElementById('abw-task-section');
    const list = document.getElementById('abw-task-list');
    const count = document.getElementById('abw-task-count');

    // 无剩余任务 → 隐藏任务列表（避免完成后最后一个任务残留）
    if (!items || items.length === 0) {
      section.style.display = 'none';
      return;
    }
    section.style.display = 'block';
    count.textContent = items.length;

    list.innerHTML = items.map((item, i) => {
      const merged = item._rows && item._rows.length > 1;
      const rowLabel = merged ? `Row${item._rows.join('+')}` : `Row${item._row}`;
      return `
        <div class="abw-task-item">
          <span>#${i+1} <span class="sku">${item.boxSku}</span> | <span class="spec">${item.packaging}</span> × <span class="qty">${item.quantity}</span>${merged ? ` <span style="color:#2ecc71;font-size:10px;">(合并${item._rows.length}行)</span>` : ''}</span>
          <span style="color:#666;font-size:10px;">${rowLabel}</span>
        </div>
      `;
    }).join('');
  }

  function updateProgress(current, total) {
    const pct = total > 0 ? Math.round((current / total) * 100) : 0;
    document.getElementById('abw-progress-bar').style.width = `${pct}%`;
  }

  function showSummary(results) {
    const summary = document.getElementById('abw-summary');
    const success = results.filter(r => r.status === 'success').length;
    const failed = results.filter(r => r.status === 'failed').length;
    const skipped = results.filter(r => r.status === 'skipped').length;
    summary.className = 'show';
    summary.style.background = failed > 0 ? 'rgba(231,76,60,0.08)' : 'rgba(46,204,113,0.08)';
    summary.innerHTML = `
      <div class="abw-label">📊 执行汇总</div>
      <div class="abw-stat abw-stat-total">总计: <b>${results.length}</b> 条</div>
      <div class="abw-stat abw-stat-success">✅ 成功: <b>${success}</b></div>
      <div class="abw-stat abw-stat-fail">❌ 失败: <b>${failed}</b></div>
      ${skipped > 0 ? '<div class="abw-stat" style="color:#e74c3c;">🚫 缺货: <b>' + skipped + '</b></div>' : ''}
    `;
  }


  // ============================================================
  //  启动
  // ============================================================
  // 等 DOM 完全就绪
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', createPanel);
  } else {
    // 小延迟确保页面脚本执行完毕
    setTimeout(createPanel, 1000);
  }

})();
