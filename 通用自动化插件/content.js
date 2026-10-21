// ═══════════════════════════════════════════════
//  Content Script — 录制 + 执行（注入所有 frames）
// ═══════════════════════════════════════════════

// ─── 选择器生成 ────────────────────────────────────
function buildCSSPath(el) {
  const parts = [];
  while (el && el !== document.documentElement && el !== document.body) {
    let part = el.tagName.toLowerCase();
    const stableClasses = [...el.classList].filter(c => !/\d/.test(c) && c.length > 2).slice(0, 2);
    if (stableClasses.length) part += '.' + stableClasses.join('.');
    const siblings = el.parentNode ? [...el.parentNode.children].filter(s => s.tagName === el.tagName) : [];
    if (siblings.length > 1) part += `:nth-of-type(${siblings.indexOf(el) + 1})`;
    parts.unshift(part);
    el = el.parentNode;
    if (parts.length > 5) break;
  }
  return parts.join(' > ');
}

function buildXPath(el) {
  const parts = [];
  let cur = el;
  while (cur && cur.nodeType === 1 && cur !== document.documentElement) {
    const tag = cur.tagName.toLowerCase();
    const siblings = cur.parentNode
      ? [...cur.parentNode.children].filter(s => s.tagName === cur.tagName)
      : [];
    const idx = siblings.indexOf(cur) + 1;
    parts.unshift(siblings.length > 1 ? `${tag}[${idx}]` : tag);
    cur = cur.parentNode;
    if (parts.length >= 6) break;
  }
  return '//' + parts.join('/');
}

// 生成多个备用选择器，按可靠性从高到低排列
function generateSelectors(el) {
  if (!el || el === document.body) return ['body'];
  const list = [];

  // 1. 稳定 ID
  if (el.id && !/^\d/.test(el.id) && !/^[a-f0-9-]{30,}$/.test(el.id) && el.id.length < 50) {
    list.push(`#${CSS.escape(el.id)}`);
  }

  // 2. FineReport widgetname
  const wn = el.getAttribute('widgetname');
  if (wn) list.push(`[widgetname="${wn}"]`);

  // 3. input / select 特征属性
  if (['INPUT', 'SELECT', 'TEXTAREA'].includes(el.tagName)) {
    const tag = el.tagName.toLowerCase();
    if (el.getAttribute('placeholder')) list.push(`${tag}[placeholder="${el.getAttribute('placeholder')}"]`);
    if (el.getAttribute('name'))        list.push(`${tag}[name="${el.getAttribute('name')}"]`);
    if (el.getAttribute('type') && el.getAttribute('type') !== 'text') {
      list.push(`${tag}[type="${el.getAttribute('type')}"]`);
    }
  }

  // 4. data-* 属性
  for (const attr of ['data-id', 'data-key', 'data-index', 'data-type']) {
    if (el.hasAttribute(attr)) {
      list.push(`${el.tagName.toLowerCase()}[${attr}="${el.getAttribute(attr)}"]`);
    }
  }

  // 5. CSS 路径
  const cssPath = buildCSSPath(el);
  if (cssPath && !list.includes(cssPath)) list.push(cssPath);

  // 6. XPath
  list.push(buildXPath(el));

  // 7. 文字（按钮/链接/可见文字）
  const text = el.textContent?.trim().substring(0, 50);
  if (text && ['BUTTON', 'A', 'SPAN', 'LI', 'TD', 'DIV', 'P'].includes(el.tagName)) {
    list.push(`text/${text}`);
  }

  return list.filter(Boolean);
}

// 兼容旧流程：返回最佳单选择器
function generateSelector(el) {
  return generateSelectors(el)[0] || 'body';
}

// 依次尝试选择器数组，返回第一个找到的元素
function findElementBySelectors(selectors, root = document) {
  const list = Array.isArray(selectors) ? selectors : (selectors ? [selectors] : []);
  for (const sel of list) {
    try {
      if (!sel) continue;
      if (sel.startsWith('text/')) {
        const text = sel.slice(5);
        const candidates = (root === document ? document : root).querySelectorAll(
          'button,a,span,li,td,div,p,input'
        );
        for (const e of candidates) {
          if (e.textContent.trim() === text && e.offsetParent !== null) return e;
        }
      } else if (sel.startsWith('//')) {
        const result = document.evaluate(
          sel, root === document ? document : root,
          null, XPathResult.FIRST_ORDERED_NODE_TYPE, null
        );
        if (result.singleNodeValue) return result.singleNodeValue;
      } else {
        const el = root.querySelector(sel);
        if (el) return el;
      }
    } catch { continue; }
  }
  return null;
}

function findElement(selector, root = document) {
  return findElementBySelectors(selector, root);
}

// ─── 录制状态 ──────────────────────────────────────
let recording = false;
let ignoredEl = null;

function flashElement(el) {
  const old = el.style.outline;
  el.style.outline = '2.5px solid #1a90ff';
  setTimeout(() => { el.style.outline = old; }, 600);
}

function makeLabel(type, el, extra) {
  const tag = el.tagName?.toLowerCase() || '';
  const text = el.textContent?.trim().substring(0, 30) || extra || '';
  switch (type) {
    case 'click':  return `点击 "${text}"`;
    case 'fill':   return `填写 [${el.getAttribute('placeholder') || el.getAttribute('name') || tag}] = "${extra}"`;
    case 'select': return `选择 "${extra}"`;
    default:       return type;
  }
}

// 事件监听（录制用）
const clickHandler = (e) => {
  if (!recording) return;
  if (e.target.closest('#auto-recorder-overlay')) return;
  if (e.target === ignoredEl || e.target.tagName === 'HTML') return;

  const selectors = generateSelectors(e.target);
  const text = e.target.textContent.trim().substring(0, 40);
  flashElement(e.target);

  const step = {
    type: 'click',
    selector: selectors[0],
    selectors,
    text,
    label: `点击 "${text || selectors[0]}"`,
    frameId: window === window.top ? 0 : undefined
  };
  chrome.runtime.sendMessage({ type: 'CS_STEP_RECORDED', step });
};

let fillTimer = null;
const inputHandler = (e) => {
  if (!recording || !['INPUT', 'TEXTAREA'].includes(e.target.tagName)) return;
  clearTimeout(fillTimer);
  fillTimer = setTimeout(() => {
    const selectors = generateSelectors(e.target);
    const val = e.target.value;
    const step = {
      type: 'fill',
      selector: selectors[0],
      selectors,
      value: val,
      label: `填写 "${val}"`,
      frameId: window === window.top ? 0 : undefined
    };
    chrome.runtime.sendMessage({ type: 'CS_STEP_RECORDED', step });
  }, 600);
};

const changeHandler = (e) => {
  if (!recording || e.target.tagName !== 'SELECT') return;
  const selectors = generateSelectors(e.target);
  const opt = e.target.options[e.target.selectedIndex]?.text || '';
  const step = {
    type: 'select_option',
    selector: selectors[0],
    selectors,
    option: opt,
    label: `选择 "${opt}"`,
    frameId: window === window.top ? 0 : undefined
  };
  chrome.runtime.sendMessage({ type: 'CS_STEP_RECORDED', step });
};

function attachRecordListeners() {
  document.addEventListener('click', clickHandler, true);
  document.addEventListener('input', inputHandler, true);
  document.addEventListener('change', changeHandler, true);
}
function detachRecordListeners() {
  document.removeEventListener('click', clickHandler, true);
  document.removeEventListener('input', inputHandler, true);
  document.removeEventListener('change', changeHandler, true);
}

// ─── 元素拾取器 ────────────────────────────────────
let pickerActive = false;
let pickerHovered = null;

const pickerMove = (e) => {
  if (!pickerActive) return;
  if (pickerHovered && pickerHovered !== e.target) {
    pickerHovered.style.outline = '';
  }
  pickerHovered = e.target;
  e.target.style.outline = '2px dashed #ff9800';
};

const pickerClick = (e) => {
  if (!pickerActive) return;
  e.preventDefault();
  e.stopPropagation();
  pickerActive = false;
  if (pickerHovered) pickerHovered.style.outline = '';
  document.removeEventListener('mousemove', pickerMove, true);
  document.removeEventListener('click', pickerClick, true);
  document.body.style.cursor = '';

  const selectors = generateSelectors(e.target);
  const label = e.target.textContent.trim().substring(0, 40) || selectors[0];
  chrome.runtime.sendMessage({ type: 'CS_PICKER_RESULT', selector: selectors[0], selectors, label });
};

// ─── 执行工具 ──────────────────────────────────────
function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

function triggerReactive(el, value) {
  try {
    const nativeSetter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set;
    if (nativeSetter) nativeSetter.call(el, value);
    else el.value = value;
  } catch { el.value = value; }
  ['input', 'change'].forEach(evt => el.dispatchEvent(new Event(evt, { bubbles: true })));
}

// 等待元素出现（接受选择器数组或字符串）
function waitFor(selectors, timeout = 15000, root = document) {
  return new Promise((resolve, reject) => {
    const start = Date.now();
    const primary = Array.isArray(selectors) ? selectors[0] : selectors;
    const check = () => {
      const el = findElementBySelectors(selectors, root);
      if (el) { resolve(el); return; }
      if (Date.now() - start > timeout) { reject(new Error(`等待超时: ${primary}`)); return; }
      setTimeout(check, 300);
    };
    check();
  });
}

function waitForLoad(timeout = 20000) {
  return new Promise((resolve) => {
    const done = () => { setTimeout(resolve, 800); };
    if (document.readyState === 'complete') { done(); return; }
    window.addEventListener('load', done, { once: true });
    setTimeout(done, timeout);
  });
}

// 带回退的点击（先按选择器数组，失败则按文字）
async function robustClick(selectors, fallbackText) {
  let el = findElementBySelectors(selectors);
  if (!el && fallbackText) {
    const all = document.querySelectorAll('button, a, span, div, li');
    for (const e of all) {
      if (e.textContent.trim().includes(fallbackText) && e.offsetParent !== null) {
        el = e; break;
      }
    }
  }
  const primary = Array.isArray(selectors) ? selectors[0] : selectors;
  if (!el) throw new Error(`找不到元素: ${primary}`);
  el.scrollIntoView({ block: 'center' });
  await sleep(100);
  el.click();
}

// ─── 消息响应（执行指令） ───────────────────────────
chrome.runtime.onMessage.addListener((msg, sender, reply) => {

  if (msg.type === 'START_RECORDING') {
    recording = true;
    attachRecordListeners();
    showOverlay('录制中 🔴');
  }
  if (msg.type === 'STOP_RECORDING') {
    recording = false;
    detachRecordListeners();
    hideOverlay();
  }

  if (msg.type === 'STOP') {
    window.__autoStop = true;
  }

  if (msg.type === 'ACTIVATE_PICKER') {
    pickerActive = true;
    document.body.style.cursor = 'crosshair';
    document.addEventListener('mousemove', pickerMove, true);
    document.addEventListener('click', pickerClick, true);
  }
  if (msg.type === 'CANCEL_PICKER') {
    pickerActive = false;
    if (pickerHovered) pickerHovered.style.outline = '';
    document.removeEventListener('mousemove', pickerMove, true);
    document.removeEventListener('click', pickerClick, true);
    document.body.style.cursor = '';
  }

  // 执行指令 ─────────────────────────────────────
  if (msg.type === 'EXEC_CLICK') {
    robustClick(msg.selectors || msg.selector, msg.text)
      .then(() => chrome.runtime.sendMessage({ type: 'CS_EXEC_DONE' }))
      .catch(e => chrome.runtime.sendMessage({ type: 'CS_EXEC_ERROR', error: e.message }));
  }

  if (msg.type === 'EXEC_FILL') {
    (async () => {
      const el = await waitFor(msg.selectors || msg.selector).catch(() => null);
      if (!el) return chrome.runtime.sendMessage({ type: 'CS_EXEC_ERROR', error: `填写：找不到 ${msg.selector}` });
      triggerReactive(el, msg.value);
      chrome.runtime.sendMessage({ type: 'CS_EXEC_DONE' });
    })();
  }

  if (msg.type === 'EXEC_SELECT') {
    (async () => {
      const sel = findElementBySelectors(msg.selectors || msg.selector);
      if (sel && sel.tagName === 'SELECT') {
        const opt = [...sel.options].find(o => o.text.trim() === msg.option);
        if (opt) { sel.value = opt.value; sel.dispatchEvent(new Event('change', { bubbles: true })); }
        return chrome.runtime.sendMessage({ type: 'CS_EXEC_DONE' });
      }
      try {
        const trigger = sel?.querySelector('input, .fr-trigger-button, .el-select__caret') || sel;
        if (trigger) trigger.click();
        await sleep(500);
        const items = document.querySelectorAll('.el-select-dropdown__item, .fr-combo-list-item, .fr-select-item');
        for (const item of items) {
          if ((item.getAttribute('title') || item.textContent.trim()) === msg.option) {
            item.click();
            return chrome.runtime.sendMessage({ type: 'CS_EXEC_DONE' });
          }
        }
        document.body.click();
        chrome.runtime.sendMessage({ type: 'CS_EXEC_ERROR', error: `找不到选项: ${msg.option}` });
      } catch (e) {
        chrome.runtime.sendMessage({ type: 'CS_EXEC_ERROR', error: e.message });
      }
    })();
  }

  if (msg.type === 'EXEC_WAIT_LOAD') {
    waitForLoad(msg.timeout)
      .then(() => chrome.runtime.sendMessage({ type: 'CS_EXEC_DONE' }));
  }

  if (msg.type === 'EXEC_WAIT_ELEMENT') {
    waitFor(msg.selectors || msg.selector, msg.timeout)
      .then(() => chrome.runtime.sendMessage({ type: 'CS_EXEC_DONE' }))
      .catch(e => chrome.runtime.sendMessage({ type: 'CS_EXEC_ERROR', error: e.message }));
  }

  if (msg.type === 'EXEC_GO_BACK') {
    history.back();
    const done = () => chrome.runtime.sendMessage({ type: 'CS_EXEC_DONE' });
    window.addEventListener('popstate', done, { once: true });
    setTimeout(done, 3000);
  }

  if (msg.type === 'EXEC_DOWNLOAD_FR') {
    try {
      if (typeof _g === 'function' && _g()) {
        _g().exportReportToExcel('simple');
        chrome.runtime.sendMessage({ type: 'CS_EXEC_DONE' });
      } else {
        const btn = [...document.querySelectorAll('button, a, span')].find(e => e.textContent.trim().includes('导出') && e.offsetParent);
        if (btn) { btn.click(); chrome.runtime.sendMessage({ type: 'CS_EXEC_DONE' }); }
        else chrome.runtime.sendMessage({ type: 'CS_EXEC_ERROR', error: '找不到导出按钮' });
      }
    } catch (e) {
      chrome.runtime.sendMessage({ type: 'CS_EXEC_ERROR', error: e.message });
    }
  }

  // 收集循环项 ─────────────────────────────────────
  if (msg.type === 'EXEC_COLLECT_ELEMENTS') {
    const els = [...document.querySelectorAll(msg.selector)];
    const items = els.map((el, i) => ({
      text: el.textContent.trim().substring(0, 100),
      selector: `${msg.selector}:nth-of-type(${i + 1})`,
      index: i
    }));
    chrome.runtime.sendMessage({ type: 'CS_LOOP_ITEMS', items });
  }

  if (msg.type === 'EXEC_COLLECT_DROPDOWN') {
    (async () => {
      const widget = findElementBySelectors(msg.selectors || msg.selector);
      if (!widget) return chrome.runtime.sendMessage({ type: 'CS_LOOP_ITEMS', items: [] });
      const trigger = widget.querySelector('.fr-trigger-button, .el-select__caret, input') || widget;
      trigger.click();
      await sleep(600);
      const items = [...document.querySelectorAll('.el-select-dropdown__item, .fr-combo-list-item, .fr-select-item')]
        .filter(el => el.offsetParent !== null)
        .map((el, i) => ({
          text: (el.getAttribute('title') || el.textContent.trim()).substring(0, 100),
          index: i,
          selector: msg.selector
        }))
        .filter(item => item.text && item.text !== '不选');
      document.body.click();
      chrome.runtime.sendMessage({ type: 'CS_LOOP_ITEMS', items });
    })();
  }
});

// ─── 录制悬浮条 ────────────────────────────────────
function showOverlay(text) {
  let el = document.getElementById('auto-recorder-overlay');
  if (!el) {
    el = document.createElement('div');
    el.id = 'auto-recorder-overlay';
    el.style.cssText = [
      'position:fixed', 'top:12px', 'right:12px', 'z-index:2147483647',
      'background:#e53935', 'color:#fff', 'padding:7px 14px',
      'border-radius:20px', 'font-size:13px', 'font-family:sans-serif',
      'box-shadow:0 2px 8px rgba(0,0,0,0.3)', 'pointer-events:none'
    ].join(';');
    document.body.appendChild(el);
  }
  el.textContent = text;
}
function hideOverlay() {
  document.getElementById('auto-recorder-overlay')?.remove();
}

chrome.runtime.sendMessage({ type: 'CS_FRAME_READY', url: location.href });
