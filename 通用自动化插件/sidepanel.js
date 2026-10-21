// ═══════════════════════════════════════════════
//  Side Panel — UI 逻辑
// ═══════════════════════════════════════════════

// ─── 状态 ──────────────────────────────────────
let steps = [];          // 当前流程步骤（可含 loop 嵌套）
let recording = false;
let running = false;
let selectedIdxs = new Set(); // 多选步骤（用于创建循环）

// ─── 获取当前激活 Tab ──────────────────────────
async function getActiveTab() {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  return tab;
}

// ─── 流程存储 ──────────────────────────────────
async function loadFlowList() {
  const data = await chrome.storage.local.get('flows');
  const flows = data.flows || {};
  const sel = document.getElementById('flowSelect');
  const curVal = sel.value;
  sel.innerHTML = '<option value="">— 选择已保存流程 —</option>';
  Object.keys(flows).forEach(name => {
    const opt = document.createElement('option');
    opt.value = name;
    opt.textContent = name;
    sel.appendChild(opt);
  });
  if (curVal) sel.value = curVal;
}

async function saveFlow() {
  const name = document.getElementById('flowName').value.trim();
  if (!name) return alert('请填写流程名称');
  const data = await chrome.storage.local.get('flows');
  const flows = data.flows || {};
  flows[name] = { name, steps, savedAt: Date.now() };
  await chrome.storage.local.set({ flows });
  await loadFlowList();
  document.getElementById('flowSelect').value = name;
  alert(`已保存: ${name}`);
}

async function loadFlow(name) {
  if (!name) return;
  const data = await chrome.storage.local.get('flows');
  const flow = (data.flows || {})[name];
  if (!flow) return;
  steps = flow.steps || [];
  document.getElementById('flowName').value = name;
  renderSteps();
}

async function deleteFlow() {
  const name = document.getElementById('flowSelect').value;
  if (!name) return;
  if (!confirm(`确定删除流程"${name}"？`)) return;
  const data = await chrome.storage.local.get('flows');
  delete (data.flows || {})[name];
  await chrome.storage.local.set({ flows: data.flows || {} });
  await loadFlowList();
  steps = [];
  renderSteps();
}

// ─── 步骤渲染 ───────────────────────────────────
const ICONS = {
  click: '👆', fill: '✏️', select_option: '📋', wait_load: '⏳',
  wait_time: '⏱️', wait_element: '🔍', go_back: '⬅️',
  download_fr: '📥', loop: '🔁', click_loop_item: '👆🔁', navigate: '🌐'
};

function stepIcon(type) { return ICONS[type] || '▪'; }

function stepDisplayLabel(step) {
  switch (step.type) {
    case 'click':          return `点击: ${step.text || step.selector}`;
    case 'fill':           return `填写: "${step.value}" → ${step.selector}`;
    case 'select_option':  return `选择: "${step.option}"`;
    case 'wait_load':      return `等待页面加载`;
    case 'wait_time':      return `等待 ${step.ms || 1000}ms`;
    case 'wait_element':   return `等待出现: ${step.selector}`;
    case 'go_back':        return `返回上一页`;
    case 'download_fr':    return `导出报表（FineReport）`;
    case 'loop':           return `循环 [${step.source}]: ${step.selector}`;
    case 'click_loop_item':return `点击当前循环项`;
    case 'navigate':       return `导航到 ${step.url || ''}`;
    default:               return step.type;
  }
}

function renderSteps() {
  const list = document.getElementById('stepList');
  if (steps.length === 0) {
    list.innerHTML = '<div class="empty-hint">录制或手动添加步骤后显示在此处</div>';
    document.getElementById('stepCount').textContent = '0 步';
    document.getElementById('btnRun').disabled = true;
    return;
  }
  list.innerHTML = '';
  steps.forEach((step, idx) => {
    list.appendChild(makeStepCard(step, idx, steps, null));
  });
  document.getElementById('stepCount').textContent = `${steps.length} 步`;
  document.getElementById('btnRun').disabled = false;
}

function makeStepCard(step, idx, parentArr, parentIdx) {
  const card = document.createElement('div');
  card.className = 'step-card' + (step.type === 'loop' ? ' loop-card' : '');
  if (selectedIdxs.has(idx)) card.classList.add('selected');

  // 点击选择（用于多选创建循环）
  card.addEventListener('click', (e) => {
    if (e.target.closest('.step-btns')) return;
    if (parentArr === steps) { // 只在顶层支持多选
      toggleSelect(idx);
    }
  });

  const header = document.createElement('div');
  header.className = 'step-header';

  const num = document.createElement('span');
  num.className = 'step-num';
  num.textContent = idx + 1;

  const icon = document.createElement('span');
  icon.className = 'step-icon';
  icon.textContent = stepIcon(step.type);

  const label = document.createElement('span');
  label.className = 'step-label';
  label.textContent = stepDisplayLabel(step);
  label.title = stepDisplayLabel(step);

  const btns = document.createElement('div');
  btns.className = 'step-btns';

  // 编辑
  const editBtn = document.createElement('button');
  editBtn.textContent = '编辑';
  editBtn.onclick = (e) => { e.stopPropagation(); openEditDialog(idx, parentArr); };

  // 上移
  const upBtn = document.createElement('button');
  upBtn.textContent = '↑';
  upBtn.disabled = idx === 0;
  upBtn.onclick = (e) => { e.stopPropagation(); moveStep(parentArr, idx, -1); };

  // 下移
  const dnBtn = document.createElement('button');
  dnBtn.textContent = '↓';
  dnBtn.disabled = idx === parentArr.length - 1;
  dnBtn.onclick = (e) => { e.stopPropagation(); moveStep(parentArr, idx, 1); };

  // 删除
  const delBtn = document.createElement('button');
  delBtn.textContent = '✕';
  delBtn.className = 'btn-del';
  delBtn.onclick = (e) => { e.stopPropagation(); deleteStep(parentArr, idx); };

  btns.append(editBtn, upBtn, dnBtn, delBtn);
  header.append(num, icon, label, btns);
  card.appendChild(header);

  // 循环子步骤
  if (step.type === 'loop') {
    // 元信息
    const meta = document.createElement('div');
    meta.className = 'loop-meta';
    const filterStr = step.filterCondition?.value ? `过滤: 含"${step.filterCondition.value}"` : '';
    const stopStr   = step.stopCondition?.value   ? `停止: 含"${step.stopCondition.value}"`   : '';
    meta.textContent = [filterStr, stopStr].filter(Boolean).join('  |  ') || '遍历全部';
    card.appendChild(meta);

    // 子步骤
    const subDiv = document.createElement('div');
    subDiv.className = 'sub-steps';
    (step.steps || []).forEach((sub, si) => {
      subDiv.appendChild(makeStepCard(sub, si, step.steps, idx));
    });
    // 在循环体内可以添加"点击循环项"
    const addLoopClick = document.createElement('button');
    addLoopClick.className = 'btn-xs';
    addLoopClick.textContent = '+ 点击循环项';
    addLoopClick.style.margin = '4px 0';
    addLoopClick.onclick = () => {
      step.steps.push({
        type: 'click_loop_item',
        label: '点击当前循环项',
        parentSelector: step.selector,
        frameId: step.frameId ?? 0
      });
      renderSteps();
    };
    subDiv.appendChild(addLoopClick);
    card.appendChild(subDiv);
  }

  return card;
}

function toggleSelect(idx) {
  if (selectedIdxs.has(idx)) selectedIdxs.delete(idx);
  else selectedIdxs.add(idx);
  renderSteps();
  document.getElementById('selectedCount').textContent = selectedIdxs.size;
}

function moveStep(arr, idx, dir) {
  const target = idx + dir;
  if (target < 0 || target >= arr.length) return;
  [arr[idx], arr[target]] = [arr[target], arr[idx]];
  renderSteps();
}

function deleteStep(arr, idx) {
  arr.splice(idx, 1);
  selectedIdxs.clear();
  renderSteps();
}

// ─── 编辑对话框 ─────────────────────────────────
let editTarget = null; // { arr, idx }

function openEditDialog(idx, arr) {
  editTarget = { arr, idx };
  const step = arr[idx];
  const fields = document.getElementById('editFields');
  fields.innerHTML = '';

  const addField = (label, key, type = 'text', value = '') => {
    const wrap = document.createElement('div');
    wrap.style.cssText = 'margin-bottom:8px';
    const lbl = document.createElement('label');
    lbl.textContent = label;
    lbl.style.cssText = 'display:block;font-size:11px;color:#555;margin-bottom:3px';
    const inp = document.createElement('input');
    inp.type = type;
    inp.value = value ?? '';
    inp.dataset.key = key;
    inp.style.cssText = 'width:100%;padding:5px 8px;border:1px solid #ccc;border-radius:4px;font-size:12px';
    wrap.append(lbl, inp);
    fields.appendChild(wrap);
  };

  switch (step.type) {
    case 'click':
      addField('选择器', 'selector', 'text', step.selector);
      addField('显示文字（备用识别）', 'text', 'text', step.text);
      addField('点击后等待（ms）', 'delay', 'number', step.delay || 400);
      addField('frameId（0=主框架，iframe需调整）', 'frameId', 'number', step.frameId ?? 0);
      break;
    case 'fill':
      addField('选择器', 'selector', 'text', step.selector);
      addField('填写内容（可用 {LOOP_TEXT} 引用循环项）', 'value', 'text', step.value);
      break;
    case 'select_option':
      addField('选择器', 'selector', 'text', step.selector);
      addField('选项文字', 'option', 'text', step.option);
      break;
    case 'wait_time':
      addField('等待时间（ms）', 'ms', 'number', step.ms || 1000);
      break;
    case 'wait_element':
      addField('等待出现的选择器', 'selector', 'text', step.selector);
      addField('超时（ms）', 'timeout', 'number', step.timeout || 15000);
      break;
    case 'download_fr':
      addField('选择器（留空用 FineReport API）', 'selector', 'text', step.selector || '');
      addField('下载后等待（ms）', 'delay', 'number', step.delay || 2000);
      break;
    case 'loop':
      addField('CSS 选择器', 'selector', 'text', step.selector);
      addField('过滤关键字（文字包含，留空不过滤）', 'filterValue', 'text', step.filterCondition?.value || '');
      addField('终止关键字（文字包含，留空不终止）', 'stopValue', 'text', step.stopCondition?.value || '');
      break;
    case 'navigate':
      addField('目标 URL', 'url', 'text', step.url || '');
      break;
    default:
      addField('（无可编辑字段）', '_', 'text', '');
  }

  document.getElementById('editDialog').classList.remove('hidden');
}

document.getElementById('btnEditConfirm').addEventListener('click', () => {
  if (!editTarget) return;
  const { arr, idx } = editTarget;
  const step = arr[idx];
  const inputs = document.querySelectorAll('#editFields input');
  inputs.forEach(inp => {
    const key = inp.dataset.key;
    const val = inp.type === 'number' ? Number(inp.value) : inp.value;
    if (key === 'filterValue') {
      step.filterCondition = val ? { type: 'text_contains', value: val } : null;
    } else if (key === 'stopValue') {
      step.stopCondition = val ? { type: 'text_contains', value: val } : null;
    } else {
      step[key] = val;
    }
  });
  // 同步 label
  step.label = stepDisplayLabel(step);
  editTarget = null;
  document.getElementById('editDialog').classList.add('hidden');
  renderSteps();
});

document.getElementById('btnEditCancel').addEventListener('click', () => {
  editTarget = null;
  document.getElementById('editDialog').classList.add('hidden');
});

// ─── 循环创建对话框 ──────────────────────────────
document.getElementById('btnMakeLoop').addEventListener('click', () => {
  if (selectedIdxs.size === 0) return alert('请先点击步骤卡片选中要放入循环体的步骤');
  document.getElementById('selectedCount').textContent = selectedIdxs.size;
  document.getElementById('loopDialog').classList.remove('hidden');
});

// 元素拾取器（用于循环选择器）
document.getElementById('btnPickLoop').addEventListener('click', async () => {
  const tab = await getActiveTab();
  chrome.runtime.sendMessage({ type: 'SP_ACTIVATE_PICKER', tabId: tab.id, callback: 'loop_selector' });
});

document.getElementById('btnLoopConfirm').addEventListener('click', () => {
  const selector = document.getElementById('loopSelector').value.trim();
  if (!selector) return alert('请填写循环目标选择器');

  const source = document.getElementById('loopSource').value;
  const clickTarget = document.getElementById('loopClickTarget').value.trim();
  const filterType  = document.getElementById('filterType').value;
  const filterValue = document.getElementById('filterValue').value.trim();
  const stopType    = document.getElementById('stopType').value;
  const stopValue   = document.getElementById('stopValue').value.trim();

  // 提取选中步骤
  const selectedArr = [...selectedIdxs].sort((a, b) => a - b);
  const loopBody = selectedArr.map(i => steps[i]);

  // 在第一个选中位置插入循环，移除原步骤
  const insertAt = selectedArr[0];
  selectedArr.reverse().forEach(i => steps.splice(i, 1));

  const loopStep = {
    type: 'loop',
    source,
    selector,
    frameId: Number(document.getElementById('loopSelector').dataset.frameId || 0),
    clickTarget: clickTarget || null,
    filterCondition: filterType && filterValue ? { type: filterType, value: filterValue } : null,
    stopCondition:   stopType  && stopValue   ? { type: stopType,   value: stopValue   } : null,
    steps: loopBody,
    label: `循环 [${source}]: ${selector}`
  };
  steps.splice(insertAt, 0, loopStep);

  selectedIdxs.clear();
  document.getElementById('loopDialog').classList.add('hidden');
  renderSteps();
});

document.getElementById('btnLoopCancel').addEventListener('click', () => {
  document.getElementById('loopDialog').classList.add('hidden');
});

// 循环来源切换时隐藏/显示点击目标行
document.getElementById('loopSource').addEventListener('change', () => {
  const src = document.getElementById('loopSource').value;
  document.getElementById('loopClickTargetRow').style.display = src === 'dropdown' ? 'none' : 'block';
});

// ─── 控制按钮 ─────────────────────────────────────
document.getElementById('btnRecord').addEventListener('click', async () => {
  const tab = await getActiveTab();
  if (!recording) {
    recording = true;
    document.getElementById('btnRecord').textContent = '■ 停止录制';
    document.getElementById('btnRecord').classList.add('active');
    chrome.runtime.sendMessage({ type: 'SP_START_RECORDING', tabId: tab.id });
  } else {
    recording = false;
    document.getElementById('btnRecord').textContent = '● 录制';
    document.getElementById('btnRecord').classList.remove('active');
    chrome.runtime.sendMessage({ type: 'SP_STOP_RECORDING' }, (resp) => {
      if (resp?.steps) {
        steps = resp.steps;
        renderSteps();
      }
    });
  }
});

document.getElementById('btnRun').addEventListener('click', async () => {
  if (steps.length === 0) return;
  const tab = await getActiveTab();
  running = true;
  setRunningUI(true);
  chrome.runtime.sendMessage({ type: 'SP_RUN_FLOW', flow: { steps }, tabId: tab.id });
});

document.getElementById('btnStop').addEventListener('click', () => {
  chrome.runtime.sendMessage({ type: 'SP_STOP_RUN' });
  setRunningUI(false);
});

document.getElementById('btnClear').addEventListener('click', () => {
  if (steps.length === 0 || confirm('清空所有步骤？')) {
    steps = [];
    selectedIdxs.clear();
    renderSteps();
  }
});

// 手动添加步骤
document.getElementById('btnAddWait').addEventListener('click', () => {
  steps.push({ type: 'wait_load', label: '等待页面加载' });
  renderSteps();
});
document.getElementById('btnAddBack').addEventListener('click', () => {
  steps.push({ type: 'go_back', label: '返回上一页' });
  renderSteps();
});

// 流程管理
document.getElementById('btnSave').addEventListener('click', saveFlow);
document.getElementById('btnDelete').addEventListener('click', deleteFlow);
document.getElementById('flowSelect').addEventListener('change', (e) => loadFlow(e.target.value));

// ─── UI 状态切换 ─────────────────────────────────
function setRunningUI(yes) {
  document.getElementById('btnRun').disabled = yes;
  document.getElementById('btnStop').disabled = !yes;
  document.getElementById('btnRecord').disabled = yes;
  document.getElementById('progressArea').classList.toggle('hidden', !yes);
}

// ─── background 消息监听 ──────────────────────────
chrome.runtime.onMessage.addListener((msg) => {
  switch (msg.type) {

    case 'STEP_ADDED':
      // 录制时步骤实时添加（steps 已在 STOP_RECORDING 中整体更新）
      document.getElementById('stepCount').textContent = `${msg.total} 步（录制中）`;
      document.getElementById('btnRun').disabled = true;
      break;

    case 'RECORDING_STOPPED':
      steps = msg.steps || steps;
      renderSteps();
      break;

    case 'PROGRESS': {
      const p = msg.progress;
      const total = p.stepTotal || 1;
      const pct = Math.round((p.stepIdx / total) * 100);
      document.getElementById('progBar').style.width = pct + '%';
      let txt = `步骤 ${p.stepIdx}/${p.stepTotal}`;
      if (p.loopTotal > 0) txt += `  |  循环 ${p.loopIdx}/${p.loopTotal}`;
      if (p.msg) txt += `  —  ${p.msg}`;
      document.getElementById('progText').textContent = txt;
      break;
    }

    case 'RUN_DONE':
      running = false;
      setRunningUI(false);
      document.getElementById('progText').textContent = '✅ 执行完成';
      break;

    case 'RUN_ERROR':
      running = false;
      setRunningUI(false);
      document.getElementById('progText').textContent = `❌ ${msg.error}`;
      break;

    case 'PICKER_RESULT':
      if (msg.callback === 'loop_selector') {
        document.getElementById('loopSelector').value = msg.selector;
        document.getElementById('loopSelector').dataset.frameId = String(msg.frameId ?? 0);
      }
      break;
  }
});

// ─── 导出 ────────────────────────────────────────
function downloadFile(filename, content, mimeType) {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

document.getElementById('btnExportJson').addEventListener('click', () => {
  if (steps.length === 0) return alert('没有可导出的步骤');
  const name = document.getElementById('flowName').value.trim() || '未命名流程';
  const payload = { name, exportedAt: new Date().toISOString(), steps };
  downloadFile(`${name}.json`, JSON.stringify(payload, null, 2), 'application/json');
});

document.getElementById('btnExportPy').addEventListener('click', () => {
  if (steps.length === 0) return alert('没有可导出的步骤');
  const name = document.getElementById('flowName').value.trim() || '未命名流程';
  downloadFile(`${name}.py`, buildPythonScript(name, steps), 'text/plain');
});

function buildPythonScript(name, allSteps) {
  const now = new Date().toLocaleString('zh-CN');
  const L = [];

  L.push(`# 流程: ${name}`);
  L.push(`# 导出时间: ${now}`);
  L.push(`# 由「网页自动化录制工具」导出`);
  L.push('');
  L.push('from selenium import webdriver');
  L.push('from selenium.webdriver.common.by import By');
  L.push('from selenium.webdriver.support.ui import WebDriverWait, Select');
  L.push('from selenium.webdriver.support import expected_conditions as EC');
  L.push('import time');
  L.push('');
  L.push('');
  L.push('def find_el(driver, selectors, timeout=10):');
  L.push('    """依次尝试多个选择器，返回第一个找到的可见元素"""');
  L.push('    for sel in selectors:');
  L.push('        try:');
  L.push('            if sel.startswith("text/"):');
  L.push('                text = sel[5:]');
  L.push('                els = driver.find_elements(By.XPATH, f\'//*/[normalize-space(.)="{text}"]\')');
  L.push('                visible = [e for e in els if e.is_displayed()]');
  L.push('                if visible: return visible[0]');
  L.push('            elif sel.startswith("//"):');
  L.push('                el = WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.XPATH, sel)))');
  L.push('                return el');
  L.push('            else:');
  L.push('                el = WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.CSS_SELECTOR, sel)))');
  L.push('                return el');
  L.push('        except Exception:');
  L.push('            continue');
  L.push('    raise Exception(f"找不到元素，已尝试: {selectors}")');
  L.push('');
  L.push('');
  L.push('driver = webdriver.Chrome()');
  L.push('driver.implicitly_wait(5)');
  L.push('');
  L.push('# ── 步骤 ─────────────────────────────────────');

  function emit(stepArr, indent) {
    stepArr.forEach((step, i) => {
      const sels = JSON.stringify(step.selectors || [step.selector].filter(Boolean));
      const comment = step.label || step.text || step.type;
      L.push('');
      L.push(`${indent}# ${i + 1}. ${comment}`);

      switch (step.type) {
        case 'click':
          L.push(`${indent}find_el(driver, ${sels}).click()`);
          L.push(`${indent}time.sleep(${((step.delay || 400) / 1000).toFixed(1)})`);
          break;

        case 'fill':
          L.push(`${indent}el = find_el(driver, ${sels})`);
          L.push(`${indent}el.clear()`);
          L.push(`${indent}el.send_keys(${JSON.stringify(step.value || '')})`);
          L.push(`${indent}time.sleep(0.3)`);
          break;

        case 'select_option':
          L.push(`${indent}Select(find_el(driver, ${sels})).select_by_visible_text(${JSON.stringify(step.option || '')})`);
          L.push(`${indent}time.sleep(0.4)`);
          break;

        case 'wait_load':
          L.push(`${indent}WebDriverWait(driver, 20).until(lambda d: d.execute_script("return document.readyState") == "complete")`);
          break;

        case 'wait_time':
          L.push(`${indent}time.sleep(${((step.ms || 1000) / 1000).toFixed(1)})`);
          break;

        case 'wait_element':
          L.push(`${indent}WebDriverWait(driver, ${Math.round((step.timeout || 15000) / 1000)}).until(`);
          L.push(`${indent}    EC.presence_of_element_located((By.CSS_SELECTOR, ${JSON.stringify(step.selector || '')})))`);
          break;

        case 'go_back':
          L.push(`${indent}driver.back()`);
          L.push(`${indent}time.sleep(1.0)`);
          break;

        case 'navigate':
          L.push(`${indent}driver.get(${JSON.stringify(step.url || '')})`);
          L.push(`${indent}WebDriverWait(driver, 20).until(lambda d: d.execute_script("return document.readyState") == "complete")`);
          L.push(`${indent}time.sleep(1.0)`);
          break;

        case 'loop': {
          const filterVal = step.filterCondition?.value;
          L.push(`${indent}loop_items = driver.find_elements(By.CSS_SELECTOR, ${JSON.stringify(step.selector || '')})`);
          if (filterVal) {
            L.push(`${indent}loop_items = [e for e in loop_items if ${JSON.stringify(filterVal)} in e.text]`);
          }
          L.push(`${indent}for _item in loop_items:`);
          if (step.steps && step.steps.length > 0) {
            emit(step.steps, indent + '    ');
          } else {
            L.push(`${indent}    pass`);
          }
          break;
        }

        default:
          L.push(`${indent}# [未支持: ${step.type}]`);
      }
    });
  }

  emit(allSteps, '');
  L.push('');
  L.push('# driver.quit()');
  return L.join('\n');
}

// ─── 启动系统 + 预置流程按钮 ─────────────────────

const SYSTEM_BASE_URL = 'http://11.210.4.110/fsdc2020/new/main/';

document.getElementById('btnLaunch').addEventListener('click', () => {
  chrome.runtime.sendMessage({ type: 'SP_OPEN_URL', url: SYSTEM_BASE_URL });
});

document.getElementById('btnPreset1').addEventListener('click', () => {
  const rec = PRESET_CR_RECORDINGS[0];
  steps = convertChromeRecording(rec.data);
  document.getElementById('flowName').value = rec.name;
  selectedIdxs.clear();
  renderSteps();
});

document.getElementById('btnPreset2').addEventListener('click', () => {
  const rec = PRESET_CR_RECORDINGS[1];
  steps = convertChromeRecording(rec.data);
  document.getElementById('flowName').value = rec.name;
  selectedIdxs.clear();
  renderSteps();
});

// ─── 导入 Chrome DevTools Recorder JSON ──────────

document.getElementById('btnImportCr').addEventListener('click', () => {
  document.getElementById('crFileInput').click();
});

document.getElementById('crFileInput').addEventListener('change', (e) => {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (ev) => {
    try {
      const data = JSON.parse(ev.target.result);
      const imported = convertChromeRecording(data);
      if (!imported.length) { alert('未能从文件中解析出任何可用步骤'); return; }

      if (steps.length > 0) {
        if (confirm(`当前已有 ${steps.length} 步。\n确定 → 追加到末尾\n取消 → 替换全部`)) {
          steps.push(...imported);
        } else {
          steps = imported;
        }
      } else {
        steps = imported;
      }

      if (data.title) document.getElementById('flowName').value = data.title;
      renderSteps();

      // 如果录制包含 iframe 步骤，提示用户
      const hasFrame = imported.some(s => s._hasFrame);
      imported.forEach(s => delete s._hasFrame);
      if (hasFrame) alert(`⚠️ 注意：此录制包含 iframe 内的操作（frameId 已设为 0）。\n若执行时找不到元素，请用"编辑"按钮调整 frameId。`);
    } catch (err) {
      alert('解析失败: ' + err.message);
    }
    e.target.value = '';
  };
  reader.readAsText(file, 'utf-8');
});

function convertChromeSelectors(rawSelectors) {
  const out = [];
  for (const group of (rawSelectors || [])) {
    for (const sel of (Array.isArray(group) ? group : [group])) {
      if (!sel || typeof sel !== 'string') continue;
      if (sel.startsWith('aria/') || sel.startsWith('pierce/')) continue;
      if (sel.startsWith('xpath/')) {
        out.push(sel.slice(5)); // "xpath//html/..." → "//html/..."
      } else {
        out.push(sel);
      }
    }
  }
  return out.filter((s, i) => out.indexOf(s) === i); // 去重保序
}

function convertChromeRecording(data) {
  const result = [];
  for (const step of (data.steps || [])) {
    if (step.type === 'setViewport') continue;

    if (step.type === 'navigate') {
      result.push({ type: 'navigate', url: step.url, label: `导航到 ${step.url}` });
      result.push({ type: 'wait_load', label: '等待页面加载' });
      continue;
    }

    if (step.type === 'click' || step.type === 'doubleClick') {
      const selectors = convertChromeSelectors(step.selectors);
      if (!selectors.length) continue;
      const textSel = selectors.find(s => s.startsWith('text/'));
      const text = textSel ? textSel.slice(5) : '';
      // frame[0] > 0 表示操作在 iframe 内；实际 frameId 在运行时才能确定
      const inFrame = step.frame && step.frame.length > 0 && step.frame[0] > 0;
      const converted = {
        type: 'click',
        selector: selectors[0],
        selectors,
        text,
        label: `点击 "${text || selectors[0]}"`,
        frameId: 0,
        delay: 400
      };
      if (inFrame) converted._hasFrame = true; // 临时标记，弹窗后删除
      result.push(converted);
    }

    // change 事件 → fill 或 select_option
    if (step.type === 'change') {
      const selectors = convertChromeSelectors(step.selectors);
      if (!selectors.length) continue;
      const inFrame = step.frame && step.frame.length > 0 && step.frame[0] > 0;
      // 根据 value 类型判断：有 value 字段用 fill，否则用 select_option
      if (step.value !== undefined) {
        const converted = {
          type: 'fill',
          selector: selectors[0],
          selectors,
          value: String(step.value),
          label: `填写 "${step.value}"`,
          frameId: 0
        };
        if (inFrame) converted._hasFrame = true;
        result.push(converted);
      }
    }
  }
  return result;
}

// ─── 初始化 ─────────────────────────────────────
loadFlowList();
renderSteps();
