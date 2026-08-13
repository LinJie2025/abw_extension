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
    },

    // 行为参数
    behavior: {
      delayMin: 2000,       // 操作间最小延迟 (ms)
      delayMax: 4000,       // 操作间最大延迟 (ms)
      pageLoadTimeout: 8000,// 页面加载超时 (ms)
      modalWaitTimeout: 8000, // 等待加购弹窗超时 (ms)
      retryCount: 2,        // 单条失败重试次数
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
  const STORE = {
    get(key, fallback) {
      try { return JSON.parse(sessionStorage.getItem(key) || 'null') ?? fallback; } catch (e) { return fallback; }
    },
    set(key, val) {
      try { sessionStorage.setItem(key, JSON.stringify(val)); } catch (e) {
        console.warn('[ABW] sessionStorage 写入失败，可能存储空间不足', key, e);
      }
    },
    remove(key) {
      try { sessionStorage.removeItem(key); } catch (e) {}
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
  };
  const MAX_LOGS = 20000; // ~1.4MB，覆盖1300+商品，远低于 sessionStorage 5MB 上限

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
  function isElementReal(el) {
    if (!el) return false;
    // 1) 必须挂载在 document 中
    if (!document.contains(el)) return false;
    // 2) offsetParent 为 null → 元素或祖先 display:none（最常见的 Angular 未渲染情况）
    if (el.offsetParent === null && el.tagName !== 'BODY') return false;
    // 3) 尺寸为 0 不可交互
    const rect = el.getBoundingClientRect();
    if (rect.width <= 1 || rect.height <= 1) return false;
    // 4) disabled
    if (el.disabled) return false;
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
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

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

    // 解析数据行（跳过表头）
    const items = [];
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row || row.every(c => c == null || String(c || '').trim() === '')) continue;

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
    return { items, headers, colMapLog, rawRows: rows, batchNo };
  }

  // ============================================================
  //  核心执行引擎
  // ============================================================
  let isRunning = false;
  const results = [];

  // 加购后半段（设数量 → 点击 → 校验弹窗）
  async function doAddToBag(_addBtn, item, addQty) {
    // Step 5: 设置数量
    log(`  → 设置数量: ${addQty}`, 'info');
    const qtyInput = await waitFor(CONFIG.selectors.qtyInput, 3000);
    if (qtyInput) {
      qtyInput.focus();
      qtyInput.select();
      const nativeInputValueSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
      nativeInputValueSetter.call(qtyInput, String(addQty));
      qtyInput.dispatchEvent(new Event('input', { bubbles: true }));
      qtyInput.dispatchEvent(new Event('change', { bubbles: true }));
      await sleep(500);
    } else {
      log(`  ⚠️ 未找到数量输入框，使用默认数量`, 'warn');
    }
    // Step 6: 重新查找加购按钮（设数量后 Angular 可能替换了 DOM）
    const addBtn = await findAddToBagButton();
    if (!addBtn) throw new Error('设数量后找不到加购按钮');
    log(`  → 点击加购...`, 'info');
    addBtn.click();
    await sleep(1000);
    // Step 7: 检测是否弹出 "Select size or color" 弹窗（需人工选择规格）
    if (isSelectSizeDialogOpen()) {
      log(`  ⚠️ 检测到 "Select size or color" 弹窗，暂未支持该加购类型`, 'warn');
      return { status: 'failed', item, reason: ERRORS.UNSUPPORTED_TYPE };
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
  async function executeItem(item, attempt = 1) {
    const { boxSku, innerRef, packaging, quantity, _row } = item;
    const addQty = quantity;

    // 无 SKU → 直接标记失败（box 缺 Box SKU / piece 缺单件 SKU）
    if (item._noSku) {
      const reason = item._noSkuReason || ERRORS.NO_SKU;
      log(`📦 第${_row}行: ${reason}`, 'warn');
      return { status: 'failed', item, reason };
    }

    log(`📦 第${_row}行: SKU=${boxSku} | 规格=${packaging} | 数量=${addQty} | 第${attempt}次尝试`, 'action');

    try {
      // ---- Step 1: 导航到商品页 ----
      const productUrl = `/info.html/pid.${boxSku}`;
      if (!location.href.includes(boxSku)) {
        log(`  → 导航到商品页...`, 'info');
        window.location.href = productUrl;
        return { status: 'navigating', item };
      }

      await waitForPageReady();

      // ---- Cloudflare 安全验证检测（必须在 waitForProductPageReady 之前，避免商品页文本误匹配）----
      let cfRetries = 0;
      const cfMaxRetries = 5;
      while (cfRetries < cfMaxRetries) {
        const bodyText = document.body.innerText;
        const isCF = /(checking your browser|just a moment|ddos protection|cf-browser-verification|please wait.*seconds)/i.test(bodyText)
                  || /本网站使用安全服务/i.test(bodyText)  // Cloudflare 中文页面特征 短语
                  || document.querySelector('#challenge-form, #cf-challenge-running, .cf-browser-verification');
        if (!isCF) break;
        cfRetries++;
        log(`  ⚠️ Cloudflare 安全验证 (${cfRetries}/${cfMaxRetries})，等待后重试...`, 'warn');
        await sleep(5000);
        location.reload();
        await waitForPageReady();
      }
      if (cfRetries >= cfMaxRetries) {
        log(`  ❌ Cloudflare 验证未通过，跳过此商品`, 'error');
        return { status: 'failed', item, reason: ERRORS.CF_TIMEOUT };
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
      const pageText = document.body.innerText;
      const upcMatch = pageText.match(/UPC[:\s]*(\d[\d\s-]{8,})/i);
      if (upcMatch) {
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
        return executeItem(item, attempt + 1);
      }
      return { status: 'failed', item, reason: UNKNOWN_ERR, error: `[Row${_row}|${boxSku}] ${err.message}` };
    }
  }

  // 查找加购按钮（多策略：文本 "add to bag" / class / active 态，主文档+iframe）
  async function findAddToBagButton() {
    // 在指定 document 中查找（主文档和 iframe 共用）
    function searchInDoc(rootDoc) {
      // 策略1: 匹配 "add to bag" 文本，优先按钮类型再兜底 span/div
      const _selectBest = (els, exactMatch) => {
        let best = null;
        for (const el of els) {
          const txt = (el.textContent || el.value || '').trim();
          const hits = exactMatch ? txt.toLowerCase() === 'add to bag' : /add\s*to\s*bag/i.test(txt);
          if (hits && isElementReal(el)) {
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
      if (byClass && isElementReal(byClass)) return byClass;

      // 策略3: span.active（选规格后激活的加购按钮）
      const activeSpan = rootDoc.querySelector('span.active');
      if (activeSpan && isElementReal(activeSpan)) return activeSpan;

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
    isRunning = true;
    forceStop = false;
    const _startedAt = new Date().toISOString();

    log(`🚀 开始执行，共 ${items.length} 条任务`, 'action');
    updateProgress(0, items.length);

    let stoppedByUser = false;
    for (let i = 0; i < items.length; i++) {
      if (forceStop) {
        stoppedByUser = true;
        break;
      }

      let result;
      try {
        result = await executeItem(items[i]);
      } catch (err) {
        // 用户点击「立即停止」：直接终止，不保存状态
        if (err === StopError) {
          log('⏹️ 已立即停止（当前条目已中断）', 'warn');
          isRunning = false;
          document.getElementById('abw-stop-btn').style.display = 'none';
          document.getElementById('abw-start-btn').style.display = '';
          return;
        }
        throw err;
      }
      results.push(result);
      if (!results._startedAt) results._startedAt = _startedAt;
      updateProgress(i + 1, items.length);

      // 如果是导航状态（页面跳转），保存剩余任务，新页面自动续跑
      if (result.status === 'navigating') {
        STORE.set(K.pending, items.slice(i));
        STORE.set(K.results, results);
        isRunning = false;
        return;
      }

      // 条目间延迟
      if (i < items.length - 1) {
        await randomDelay();
      }
    }

    isRunning = false;
    document.getElementById('abw-stop-btn').style.display = 'none';
    document.getElementById('abw-start-btn').style.display = '';
    if (stoppedByUser) {
      log('⏹️ 已停止', 'warn');
      notifyComplete();
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

  // ============================================================
  //  浮动面板 UI
  // ============================================================
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
          <button class="abw-btn abw-btn-sec" id="abw-dl-btn" style="display:none;">📥 下载结果 Excel</button>
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
      <div id="abw-footer">ABW Auto Purchase v1.0.4 · 数据不出浏览器 · 跨页面保持</div>
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
      if (items.length > 0) {
        // 运行中：显示停止按钮，隐藏开始按钮
        document.getElementById('abw-stop-btn').style.display = '';
        document.getElementById('abw-start-btn').style.display = 'none';
        document.getElementById('abw-dl-btn').style.display = 'none';
        document.getElementById('abw-log-btn').style.display = 'none';
        results.length = 0;
        runAll(items);
      }
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
    });
    // 下载结果 Excel
    document.getElementById('abw-dl-btn').addEventListener('click', downloadResultExcel);
    // 开发日志下载按钮
    document.getElementById('abw-log-btn').addEventListener('click', downloadDevLog);


    // 恢复已解析的任务列表（跨页面保持）
    const savedItems = STORE.get(K.items, []);
    if (savedItems.length > 0) {
      renderTaskList(savedItems);
      document.getElementById('abw-start-btn').disabled = false;
      document.getElementById('abw-clear-btn').disabled = false;
    }

    // 静默恢复跳转中的任务（页面加载后自动续跑，无日志噪音）
    const pendingItems = STORE.get(K.pending, null);
    if (pendingItems && pendingItems.length > 0) {
      const prevResults = STORE.get(K.results, []);
      STORE.remove(K.pending);
      STORE.remove(K.results);
      results.push(...prevResults);
      renderTaskList(pendingItems);
      document.getElementById('abw-start-btn').disabled = false;
      setTimeout(() => runAll(pendingItems), 2000);
    }
  }

  // 下载结果 Excel（原文件 + 加购结果列 + ABW 商品名）
  function downloadResultExcel() {
    const headers = STORE.get(K.xlHeaders, []);
    const rawRows = STORE.get(K.xlRows, []);
    if (!headers.length || !rawRows.length) { log('⚠️ 无原始数据，请重新上传 Excel', 'warn'); return; }

    // 按原始行号建结果映射
    const statusMap = {};
    const reasonMap = {};
    const shipMap = {};
    for (const r of results) {
      statusMap[r.item._row] = r.status;
      if (r.reason) reasonMap[r.item._row] = r.reason;
      if (r.error) reasonMap[r.item._row] = reasonMap[r.item._row] || r.error;
      if (r.item._shipWarning) shipMap[r.item._row] = r.item._shipWarning;
    }

    // 收集 ABW 商品名
    const nameMap = {};
    const unitPriceMap = {};
    const boxPriceMap = {};
    for (const r of results) {
      if (r.item._abwProductName) nameMap[r.item._row] = r.item._abwProductName;
      if (r.item._unitPrice) unitPriceMap[r.item._row] = r.item._unitPrice;
      if (r.item._boxPrice) boxPriceMap[r.item._row] = r.item._boxPrice;
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

    // 构建输出：原表头 + ABWproductname + 品牌名 + 单价 + 整箱批发价 + 加购结果
    const outRows = [outHeaders.concat(['ABWproductname', '品牌名', '单价', '整箱批发价', '加购结果'])];
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
      let label;
      if (st === 'success') {
        label = shipMap[i + 1] ? `⏳成功(可能缺货,${shipMap[i + 1]})` : '✅成功';
      } else if (st === 'skipped') {
        label = '🚫缺货';
      } else if (reason) {
        label = `❌${reason}`;
      } else {
        label = '❌失败';
      }

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
        if (remarkIdx >= 0) outRow[remarkIdx] = '缺货';
      } else if (st === 'success' && shipMap[i + 1]) {
        if (remarkIdx >= 0) outRow[remarkIdx] = '可能缺货';
      }
      if (boxPrice && boxPriceColIdx >= 0) outRow[boxPriceColIdx] = parseFloat(boxPrice);

      outRows.push(outRow.concat([productName, brandName, unitPrice, boxPrice, label]));
    }

    // 按 ABWproductname 列 A-Z 排序
    const nameIdx = headers.length; // ABWproductname 在原始列之后
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
    a.href = url; a.download = '采购结果.xlsx'; document.body.appendChild(a); a.click();
    document.body.removeChild(a); URL.revokeObjectURL(url);
    log('📥 结果 Excel 已下载', 'success');
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

    const wrap = document.getElementById('abw-file-wrap');
    const textEl = document.getElementById('abw-file-text');

    wrap.classList.add('has-file');
    textEl.className = 'abw-file-name';
    textEl.textContent = `✅ ${file.name} (${(file.size / 1024).toFixed(1)} KB)`;

    log(`正在解析 Excel: ${file.name}...`, 'info');

    try {
      const buffer = await readFileAsArrayBuffer(file);
      const parsed = parseExcel(buffer);
      const items = parsed.items;

      if (items.length === 0) {
        throw new Error('未能解析出有效数据行，请确认列名是否匹配');
      }

      // 显示列映射，方便核对
      log(`📑 列映射: ${parsed.colMapLog.join(' | ')}`, 'info');

      // 存储解析结果 + 原始数据（下载结果时用）
      // 新文件上传 → 清空旧日志
      STORE.remove(K.logs);
      document.getElementById('abw-log').innerHTML = '';
      results.length = 0;
      STORE.set(K.items, items);
      STORE.set(K.xlHeaders, parsed.headers);
      STORE.set(K.xlRows, parsed.rawRows);
      STORE.set(K.batchNo, parsed.batchNo || '');

      // 显示任务列表
      renderTaskList(items);

      document.getElementById('abw-start-btn').disabled = false;
      document.getElementById('abw-clear-btn').disabled = false;

      log(`解析完成！共 ${items.length} 条有效任务`, 'success');
      log(`首条示例: SKU=${items[0].boxSku} | 规格=${items[0].packaging} | 数量=${items[0].quantity}`, 'info');

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

    section.style.display = 'block';
    count.textContent = items.length;

    list.innerHTML = items.map((item, i) => `
      <div class="abw-task-item">
        <span>#${i+1} <span class="sku">${item.boxSku}</span> | <span class="spec">${item.packaging}</span> × <span class="qty">${item.quantity}</span></span>
        <span style="color:#666;font-size:10px;">Row${item._row}</span>
      </div>
    `).join('');
  }

  function updateProgress(current, total) {
    const pct = Math.round((current / total) * 100);
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
