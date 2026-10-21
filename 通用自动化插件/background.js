// ═══════════════════════════════════════════════
//  Background service worker — 状态机 + 消息路由
// ═══════════════════════════════════════════════

// 点击插件图标时打开侧边栏
chrome.sidePanel.setPanelBehavior({ openPanelOnActionClick: true });

// ─── 全局状态 ────────────────────────────────────
const G = {
  // 录制
  recording: false,
  recordingTabId: null,
  pendingSteps: [],       // 当前正在录制的步骤列表（未保存）

  // 元素拾取器
  pickMode: false,
  pickCallback: null,     // 'loop_selector' | 'step_selector'

  // 执行
  running: false,
  stopFlag: false,
  tabId: null,
  frameRegistry: {},      // frameId -> url (记录 iframe)
  progress: { stepIdx: 0, stepTotal: 0, loopIdx: 0, loopTotal: 0, msg: '' }
};

// ─── 消息路由 ─────────────────────────────────────
chrome.runtime.onMessage.addListener((msg, sender, reply) => {
  switch (msg.type) {

    // ── 侧边栏 → background ──
    case 'SP_START_RECORDING':
      startRecording(msg.tabId);
      break;

    case 'SP_STOP_RECORDING':
      stopRecording();
      reply({ steps: G.pendingSteps });
      return true;

    case 'SP_RUN_FLOW':
      runFlow(msg.flow, msg.tabId);
      break;

    case 'SP_STOP_RUN':
      G.stopFlag = true;
      broadcastToTab(G.tabId, { type: 'STOP' });
      break;

    case 'SP_GET_STATE':
      reply({
        recording: G.recording,
        running: G.running,
        progress: G.progress,
        pendingSteps: G.pendingSteps
      });
      return true;

    case 'SP_ACTIVATE_PICKER':
      activatePicker(msg.tabId, msg.callback);
      break;

    case 'SP_CANCEL_PICKER':
      cancelPicker(msg.tabId);
      break;

    case 'SP_OPEN_URL':
      chrome.tabs.create({ url: msg.url, active: true });
      break;

    // ── content.js → background ──
    case 'CS_STEP_RECORDED':
      if (G.recording) {
        const step = { ...(msg.step || {}) };
        // 录制消息来自哪个 frame，就把真实 frameId 固化到步骤里
        step.frameId = sender.frameId ?? 0;
        G.pendingSteps.push(step);
        notifySidepanel({ type: 'STEP_ADDED', step: msg.step, total: G.pendingSteps.length });
      }
      break;

    case 'CS_FRAME_READY':
      G.frameRegistry[sender.frameId] = { url: sender.url || '', tabId: sender.tab?.id };
      // 页面/iframe 初始化时，若录制进行中则立即恢复
      if (G.recording && sender.tab?.id === G.recordingTabId) {
        chrome.tabs.sendMessage(
          sender.tab.id,
          { type: 'START_RECORDING' },
          { frameId: sender.frameId },
          () => chrome.runtime.lastError
        );
      }
      break;

    case 'CS_PICKER_RESULT':
      if (G.pickMode) {
        G.pickMode = false;
        notifySidepanel({
          type: 'PICKER_RESULT',
          selector: msg.selector,
          label: msg.label,
          callback: G.pickCallback,
          frameId: sender.frameId ?? 0
        });
      }
      break;

    case 'CS_EXEC_DONE':
      // 单步执行完成，由 runFlow 的 await 处理
      if (G._execResolve) { G._execResolve({ ok: true, data: msg.data }); G._execResolve = null; }
      break;

    case 'CS_EXEC_ERROR':
      if (G._execResolve) { G._execResolve({ ok: false, error: msg.error }); G._execResolve = null; }
      break;

    case 'CS_LOOP_ITEMS':
      if (G._execResolve) { G._execResolve({ ok: true, data: msg.items }); G._execResolve = null; }
      break;
  }
});

// ─── 录制 ──────────────────────────────────────────
function startRecording(tabId) {
  G.recording = true;
  G.recordingTabId = tabId;
  G.pendingSteps = [];
  broadcastToTab(tabId, { type: 'START_RECORDING' });
  notifySidepanel({ type: 'RECORDING_STARTED' });
}

function stopRecording() {
  G.recording = false;
  broadcastToTab(G.recordingTabId, { type: 'STOP_RECORDING' });
  notifySidepanel({ type: 'RECORDING_STOPPED', steps: G.pendingSteps });
}

// ─── 元素拾取器 ─────────────────────────────────────
function activatePicker(tabId, callback) {
  G.pickMode = true;
  G.pickCallback = callback;
  chrome.tabs.sendMessage(tabId, { type: 'ACTIVATE_PICKER' }, { frameId: 0 });
}

function cancelPicker(tabId) {
  G.pickMode = false;
  chrome.tabs.sendMessage(tabId, { type: 'CANCEL_PICKER' }, { frameId: 0 });
}

// ─── 执行引擎 ─────────────────────────────────────
async function runFlow(flow, tabId) {
  G.running = true;
  G.stopFlag = false;
  G.tabId = tabId;
  G.progress = { stepIdx: 0, stepTotal: flow.steps.length, loopIdx: 0, loopTotal: 0, msg: '开始执行...' };
  notifyProgress();

  try {
    await executeSteps(flow.steps, 0);
    G.progress.msg = '✅ 执行完成';
  } catch (e) {
    G.progress.msg = `❌ 错误: ${e.message}`;
    notifySidepanel({ type: 'RUN_ERROR', error: e.message });
  }

  G.running = false;
  notifyProgress();
  notifySidepanel({ type: 'RUN_DONE' });
}

async function executeSteps(steps, depth) {
  for (let i = 0; i < steps.length; i++) {
    if (G.stopFlag) throw new Error('用户停止');
    const step = steps[i];

    if (depth === 0) {
      G.progress.stepIdx = i + 1;
      G.progress.stepTotal = steps.length;
    }
    G.progress.msg = stepLabel(step);
    notifyProgress();

    await executeOneStep(step);
  }
}

async function executeOneStep(step) {
  switch (step.type) {

    case 'navigate':
      await navigateTo(step.url || '');
      break;

    case 'wait_load':
      await sendAndWait({ type: 'EXEC_WAIT_LOAD', timeout: step.timeout || 20000 });
      break;

    case 'wait_time':
      await sleep(step.ms || 1000);
      break;

    case 'wait_element':
      await sendAndWait({ type: 'EXEC_WAIT_ELEMENT', selector: step.selector, selectors: step.selectors, timeout: step.timeout || 15000 });
      break;

    case 'go_back':
      await sendAndWait({ type: 'EXEC_GO_BACK' });
      await sleep(800);
      break;

    case 'click':
      await sendAndWait({ type: 'EXEC_CLICK', selector: step.selector, selectors: step.selectors, text: step.text, frameId: step.frameId });
      await sleep(step.delay || 400);
      break;

    case 'fill':
      await sendAndWait({ type: 'EXEC_FILL', selector: step.selector, selectors: step.selectors, value: step.value, frameId: step.frameId });
      await sleep(300);
      break;

    case 'select_option':
      await sendAndWait({ type: 'EXEC_SELECT', selector: step.selector, selectors: step.selectors, option: step.option, frameId: step.frameId });
      await sleep(400);
      break;

    case 'download_fr':
      await sendAndWait({ type: 'EXEC_DOWNLOAD_FR', selector: step.selector, selectors: step.selectors, frameId: step.frameId });
      await sleep(step.delay || 2000);
      break;

    case 'loop':
      await executeLoop(step);
      break;

    default:
      console.warn('[AutoTool] 未知步骤类型:', step.type);
  }
}

async function executeLoop(loopStep) {
  // 1. 收集循环项
  let items;
  if (loopStep.source === 'dropdown') {
    const r = await sendAndWait({
      type: 'EXEC_COLLECT_DROPDOWN',
      selector: loopStep.selector,
      frameId: loopStep.frameId
    });
    items = r.data || [];
  } else {
    // elements / table_rows
    const r = await sendAndWait({
      type: 'EXEC_COLLECT_ELEMENTS',
      selector: loopStep.selector,
      frameId: loopStep.frameId
    });
    items = r.data || [];
  }

  // 2. 过滤
  const filtered = items.filter(item => {
    if (!loopStep.filterCondition) return true;
    return checkCondition(item.text, loopStep.filterCondition);
  });

  G.progress.loopTotal = filtered.length;
  G.progress.loopIdx = 0;
  notifyProgress();

  // 3. 遍历
  for (let i = 0; i < filtered.length; i++) {
    if (G.stopFlag) break;

    // 停止条件检查（在已过滤结果中）
    if (loopStep.stopCondition && checkCondition(filtered[i].text, loopStep.stopCondition)) break;

    G.progress.loopIdx = i + 1;
    G.progress.msg = `循环 ${i + 1}/${filtered.length}: ${filtered[i].text.substring(0, 30)}`;
    notifyProgress();

    const currentItem = filtered[i];

    // 执行子步骤，将 currentItem 上下文注入
    for (const sub of loopStep.steps) {
      if (G.stopFlag) break;
      const resolved = resolveStep(sub, currentItem, loopStep.source);
      await executeOneStep(resolved);
    }
  }

  G.progress.loopIdx = 0;
  G.progress.loopTotal = 0;
}

// 将步骤中的 {LOOP_TEXT} 变量替换为当前项文字；将 click_loop_item 转换为实际 click
function resolveStep(step, item, source) {
  if (step.type === 'click_loop_item') {
    if (source === 'dropdown') {
      return {
        type: 'select_option',
        selector: step.parentSelector,
        option: item.text,
        frameId: step.frameId ?? 0
      };
    }
    return { type: 'click', selector: item.selector, text: item.text, frameId: step.frameId };
  }
  if (step.value && step.value.includes('{LOOP_TEXT}')) {
    return { ...step, value: step.value.replace(/\{LOOP_TEXT\}/g, item.text) };
  }
  return step;
}

function checkCondition(text, cond) {
  if (!cond || !cond.type) return false;
  switch (cond.type) {
    case 'text_contains':     return text.includes(cond.value);
    case 'text_not_contains': return !text.includes(cond.value);
    case 'date_contains':     return text.includes(cond.value); // 日期字符串匹配
    default: return false;
  }
}

// ─── 与 content.js 通信（带 await） ──────────────────
function sendAndWait(payload) {
  return new Promise((resolve, reject) => {
    G._execResolve = resolve;
    const frameId = payload.frameId ?? 0;
    delete payload.frameId;

    chrome.tabs.sendMessage(G.tabId, payload, { frameId }, (resp) => {
      if (chrome.runtime.lastError) {
        // content.js 可能还没准备好，重试一次
        setTimeout(() => {
          chrome.tabs.sendMessage(G.tabId, payload, { frameId }, (resp2) => {
            if (chrome.runtime.lastError) {
              G._execResolve = null;
              reject(new Error(chrome.runtime.lastError.message));
            }
          });
        }, 500);
      }
    });

    setTimeout(() => {
      if (G._execResolve) {
        G._execResolve = null;
        reject(new Error(`步骤超时: ${payload.type}`));
      }
    }, 60000);
  });
}

// ─── 工具函数 ──────────────────────────────────────
function broadcastToTab(tabId, msg) {
  if (!tabId) return;
  chrome.tabs.sendMessage(tabId, msg, { frameId: 0 }, () => chrome.runtime.lastError);
  // 也广播到所有 frames
  chrome.webNavigation.getAllFrames({ tabId }, (frames) => {
    if (!frames) return;
    frames.forEach(f => {
      if (f.frameId !== 0) {
        chrome.tabs.sendMessage(tabId, msg, { frameId: f.frameId }, () => chrome.runtime.lastError);
      }
    });
  });
}

function notifySidepanel(msg) {
  chrome.runtime.sendMessage(msg, () => chrome.runtime.lastError);
}

function notifyProgress() {
  notifySidepanel({ type: 'PROGRESS', progress: G.progress });
}

function stepLabel(step) {
  switch (step.type) {
    case 'click':         return `点击: ${step.text || step.selector}`;
    case 'fill':          return `填写: ${step.value}`;
    case 'select_option': return `选择: ${step.option}`;
    case 'wait_load':     return '等待页面加载...';
    case 'wait_time':     return `等待 ${step.ms}ms`;
    case 'wait_element':  return `等待元素: ${step.selector}`;
    case 'go_back':       return '返回上一页';
    case 'download_fr':   return '导出报表';
    case 'loop':          return `循环: ${step.selector}`;
    case 'navigate':      return `导航到: ${step.url}`;
    default:              return step.type;
  }
}

function navigateTo(url) {
  return new Promise((resolve, reject) => {
    if (!url) { resolve(); return; }
    chrome.tabs.update(G.tabId, { url }, () => {
      if (chrome.runtime.lastError) { reject(new Error(chrome.runtime.lastError.message)); return; }
      const onUpdated = (tabId, info) => {
        if (tabId !== G.tabId || info.status !== 'complete') return;
        chrome.tabs.onUpdated.removeListener(onUpdated);
        setTimeout(resolve, 1000);
      };
      chrome.tabs.onUpdated.addListener(onUpdated);
      // 30 秒兜底超时
      setTimeout(() => { chrome.tabs.onUpdated.removeListener(onUpdated); resolve(); }, 30000);
    });
  });
}

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }
