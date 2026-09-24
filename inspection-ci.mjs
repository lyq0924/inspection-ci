/**
 * inspection-ci.mjs
 * CI/CD 专用巡检脚本 — 无交互、headless 模式
 *
 * 用法（通过环境变量控制）:
 *   INSPECTION_ENV=123 node inspection-ci.mjs
 *   INSPECTION_ENV=123 SCENARIO_FILES=all node inspection-ci.mjs
 *   INSPECTION_ENV=123 SCENARIO_FILES=1,3 node inspection-ci.mjs
 *   INSPECTION_ENV=123 PRIORITY=P1 node inspection-ci.mjs
 *
 * 环境变量:
 *   INSPECTION_ENV     — 环境 key（如 123, 122），默认 123
 *   SCENARIO_FILES     — 场景文件选择: all | 1,3 | 具体文件名，默认 all
 *   PRIORITY           — 按优先级过滤: P0/P1/P2/P3，留空=全部
 *   MODULE             — 按模块过滤（如: 系统管理），留空=全部
 *   HEADLESS           — 是否无头模式: true/false，默认 true
 *   MIDSCENE_MODEL     — AI 模型名称（可选）
 */

import { PlaywrightAgent } from '@midscene/web';
import { chromium } from 'playwright';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import fs from 'node:fs';

const require = createRequire(import.meta.url);
const YAML = require('js-yaml');
const ExcelJS = require('exceljs');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType, ImageRun } = require('docx');
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// .env 加载
const envRoot = path.join(__dirname, '..', '.env');
if (fs.existsSync(envRoot)) require('dotenv').config({ path: envRoot });
else require('dotenv').config();

// ==================== CI 参数读取 ====================
const CI_ENV_KEY = process.env.INSPECTION_ENV || '123';
const CI_SCENARIO_FILES = process.env.SCENARIO_FILES || 'all';
const CI_PRIORITY = process.env.PRIORITY || '';
const CI_MODULE = process.env.MODULE || '';
const CI_HEADLESS = process.env.HEADLESS !== 'false';

// ==================== 工具函数 ====================
const localTS = () => {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}T${String(d.getHours()).padStart(2,'0')}-${String(d.getMinutes()).padStart(2,'0')}-${String(d.getSeconds()).padStart(2,'0')}`;
};
const wait = (ms) => new Promise((r) => setTimeout(r, ms));

// ==================== 配置 ====================
let RESULT_BASE = path.join(__dirname, 'results', localTS());
const DEFAULT_CONFIG = {
  page_load_wait: 1500,
  page_load_timeout: 30000,
  network_idle_wait: 3000,
  screenshot_quality: 'medium',
  record_network: true,
  ai_timeout: 45000,
  ai_max_replan: 5,
};

// ==================== 环境加载 ====================
function loadEnvironments() {
  const fp = path.join(__dirname, 'data', '0 shared', 'environments.json');
  if (!fs.existsSync(fp)) { console.error('❌ environments.json 未找到'); process.exit(1); }
  const raw = fs.readFileSync(fp, 'utf-8');
  const json = raw.replace(/^\s*\/\/[^\n]*/gm, '');
  let cfg;
  try { cfg = JSON.parse(json); } catch (e) { console.error('❌ environments.json 解析失败: ' + e.message); process.exit(1); }
  const envs = {};
  for (const [k, v] of Object.entries(cfg)) {
    if (k.startsWith('_')) continue;
    const adminUrl = v.admin?.url || '';
    envs[k] = {
      name: v.name || (k + '环境'),
      admin: { loginUrl: v.admin?.loginUrl || adminUrl, startUrl: adminUrl },
      apiBaseUrl: adminUrl.replace(/\/[^/]+$/, '') + '/backend-server',
    };
    if (v.token) process.env['BACKEND_TOKEN_' + k.toUpperCase()] = v.token;
  }
  return envs;
}
const ENVIRONMENTS = loadEnvironments();

// ==================== 菜单树加载 ====================
function listMenuTreeFiles() {
  const caseDir = path.join(__dirname, 'data', 'platform-inspection-case');
  if (!fs.existsSync(caseDir)) return [];
  return fs.readdirSync(caseDir).filter(f => f.endsWith('.txt') && f.includes('菜单结构'));
}
function loadMenuTree(selectedFile) {
  const caseDir = path.join(__dirname, 'data', 'platform-inspection-case');
  const fp = path.join(caseDir, selectedFile);
  if (fs.existsSync(fp)) return fs.readFileSync(fp, 'utf-8').trim();
  return '';
}

// ==================== YAML 文件加载 ====================
function listScenarioFiles() {
  const dir = path.join(__dirname, 'scenarios');
  if (!fs.existsSync(dir)) return [];
  return fs.readdirSync(dir).filter(f => (f.endsWith('.yml') || f.endsWith('.yaml')) && !f.startsWith('~$'));
}

// 解析环境变量 ${VAR}
function resolveEnvVar(text) {
  if (typeof text !== 'string') return String(text ?? '');
  const now = new Date();
  const builtins = {
    DATE: `${now.getFullYear()}${String(now.getMonth()+1).padStart(2,'0')}${String(now.getDate()).padStart(2,'0')}`,
    TIMESTAMP: now.toISOString().replace(/[:.]/g, '-'),
    TIME: `${String(now.getHours()).padStart(2,'0')}${String(now.getMinutes()).padStart(2,'0')}`,
  };
  return text.replace(/\$\{([^}]+)\}/g, (_, key) => builtins[key] || process.env[key] || '');
}

// 自动检测 YAML 格式
function detectYamlFormat(filePath) {
  const raw = fs.readFileSync(filePath, 'utf-8');
  const content = raw.replace(/^\s*#[^\n]*/gm, '');
  if (/^tasks:/m.test(content) && /^web:/m.test(content)) return 'task';
  if (/^scenarios:/m.test(content)) return 'scenario';
  return 'unknown';
}

// Task YAML 格式加载 (web: + tasks[].flow[])
function loadTaskYaml(filePath) {
  const raw = fs.readFileSync(filePath, 'utf-8');
  const doc = YAML.load(raw);
  if (!doc || !doc.tasks) return null;
  const webConfig = doc.web || {};
  const tasks = (doc.tasks || []).filter(t => t && t.name && t.flow && t.flow.length > 0).map((t, i) => ({
    id: i + 1,
    name: t.name || `任务${i + 1}`,
    flow: t.flow || [],
  }));
  return {
    web: {
      url: webConfig.url || '',
      viewportWidth: webConfig.viewportWidth || 1401,
      viewportHeight: webConfig.viewportHeight || 694,
      acceptInsecureCerts: webConfig.acceptInsecureCerts !== false,
    },
    tasks,
    file: path.basename(filePath),
  };
}

// 执行单个 flow 步骤
async function executeFlowStep(page, agent, step, stepIdx) {
  if (!step || typeof step !== 'object') return;
  if ('aiTap' in step) {
    const target = resolveEnvVar(String(step.aiTap));
    console.log(`     步骤${stepIdx}: 点击 [${target}]`);
    await agent.aiAct(`点击"${target}"`);
    await wait(1500);
  } else if ('aiInput' in step) {
    const target = resolveEnvVar(String(step.aiInput));
    const value = resolveEnvVar(String(step.value || ''));
    console.log(`     步骤${stepIdx}: 在 [${target}] 输入 [${value}]`);
    await agent.aiAct(`在"${target}"中输入"${value}"`);
    await wait(1000);
  } else if ('aiAction' in step) {
    let desc = step.aiAction;
    if (typeof desc === 'object' && desc !== null) { const vals = Object.values(desc); desc = vals.length > 0 ? String(vals[0]) : ''; }
    desc = resolveEnvVar(String(desc));
    console.log(`     步骤${stepIdx}: 执行 [${desc.length > 60 ? desc.substring(0, 60) + '...' : desc}]`);
    await agent.aiAct(desc);
    await wait(1500);
  } else if ('aiAssert' in step) {
    const assertion = resolveEnvVar(String(step.aiAssert));
    console.log(`     步骤${stepIdx}: 断言 [${assertion.length > 60 ? assertion.substring(0, 60) + '...' : assertion}]`);
    await agent.aiAssert(assertion);
  } else if ('aiWaitFor' in step) {
    const condition = resolveEnvVar(String(step.aiWaitFor));
    const timeout = step.timeout || 10000;
    console.log(`     步骤${stepIdx}: 等待 [${condition.length > 60 ? condition.substring(0, 60) + '...' : condition}] (超时${timeout}ms)`);
    try { await agent.aiWaitFor(condition, { timeoutMs: timeout }); }
    catch {
      const maxRetries = Math.ceil(timeout / 2000); let lastErr;
      for (let i = 0; i < maxRetries; i++) { try { await agent.aiAssert(condition); break; } catch (e) { lastErr = e; if (i === maxRetries - 1) throw lastErr; } await wait(2000); }
    }
  } else if ('sleep' in step) {
    const ms = parseInt(step.sleep);
    console.log(`     步骤${stepIdx}: 等待 ${ms}ms`);
    await wait(ms);
  }
}

// 执行 Task YAML 任务
async function executeTaskFlow(page, task, agent, config, resultDir, idx, total) {
  const tag = `[${idx}/${total}]`;
  console.log(`\n${tag} 📋 ${task.name}`);
  const startTime = Date.now();
  const shotPrefix = path.join(resultDir, `${String(idx).padStart(2,'0')}-${sanitizeFileName(task.name)}`);
  await page.screenshot({ path: shotPrefix + '-before.png', fullPage: false }).catch(() => {});
  const monitor = startErrorMonitor(page);
  let stepOk = true, stepError = '', aiIntent = '';
  for (let si = 0; si < task.flow.length; si++) {
    const step = task.flow[si]; if (!step) continue;
    aiIntent = step.aiTap || step.aiAction || step.aiInput || step.aiAssert || JSON.stringify(step);
    try {
      await executeFlowStep(page, agent, step, si + 1);
      if (step.aiTap || step.aiAction) {
        await page.waitForLoadState('networkidle', { timeout: config.network_idle_wait }).catch(() => {});
      }
    } catch (e) {
      stepOk = false;
      const errMsg = e.message?.substring(0, 120) || '未知错误';
      stepError = `步骤${si + 1}失败: ${errMsg}`;
      if (errMsg.includes('Replanned')) {
        stepError += ` [AI尝试寻找"${aiIntent}"但找不到目标元素]`;
      }
      console.log(`     ❌ ${stepError}`); break;
    }
  }
  if (stepOk) { await wait(800); monitor.snapshot(); }
  const monitorSnapshot = monitor.snapshot(); monitor.stop();
  await wait(300);
  await page.screenshot({ path: shotPrefix + '-after.png', fullPage: false }).catch(() => {});
  const pageName = await detectPageName(page);
  const elapsed = Date.now() - startTime;
  const result = {
    id: idx, name: task.name, module: 'YAML任务', priority: 'P1', tags: [],
    pageName, result: stepOk ? '通过' : '失败', stepError, aiIntent,
    assertions: [], errors: monitorSnapshot.errors, totalApiCalls: monitorSnapshot.totalApiCalls,
    screenshotBefore: shotPrefix + '-before.png', screenshotAfter: shotPrefix + '-after.png',
    elapsed, timestamp: new Date().toISOString(),
  };
  console.log(`     ${stepOk ? '✅' : '❌'} ${result.result} (${(elapsed / 1000).toFixed(1)}s)`);
  return result;
}
function loadScenario(filePath) {
  const raw = fs.readFileSync(filePath, 'utf-8');
  const doc = YAML.load(raw);
  if (!doc || !doc.scenarios) { console.error(`⚠️ ${path.basename(filePath)} 无 scenarios 字段`); return null; }
  const config = { ...DEFAULT_CONFIG, ...(doc.config || {}) };
  const prerequisite = doc.prerequisite || null;
  const scenarios = (doc.scenarios || []).filter(s => s.enabled !== false).map((s, i) => ({
    id: i + 1,
    name: s.name || `场景${i + 1}`,
    module: s.module || '未分类',
    priority: s.priority || 'P2',
    tags: s.tags || [],
    steps: s.steps || [],
    assertions: s.assertions || ['页面正常', '接口正常'],
  }));
  return { config, prerequisite, scenarios, file: path.basename(filePath) };
}

// ==================== 错误监控 ====================
function startErrorMonitor(page) {
  const errors = [];
  let totalApiCalls = 0;
  const onResponse = (response) => {
    totalApiCalls++;
    if (response.status() >= 400) {
      errors.push({ type: 'API', status: response.status(), url: response.url().substring(0, 200), time: Date.now() });
    }
  };
  const onPageError = (err) => {
    errors.push({ type: 'PAGE', message: err.message?.substring(0, 200), time: Date.now() });
  };
  const onConsole = (msg) => {
    if (msg.type() === 'error') {
      const text = msg.text();
      if (/WebSocket connection|websocket|wss?:\/\//i.test(text)) return;
      errors.push({ type: 'CONSOLE', message: text.substring(0, 200), time: Date.now() });
    }
  };
  page.on('response', onResponse);
  page.on('pageerror', onPageError);
  page.on('console', onConsole);
  return {
    errors,
    totalApiCalls: () => totalApiCalls,
    stop: () => {
      page.off('response', onResponse);
      page.off('pageerror', onPageError);
      page.off('console', onConsole);
    },
    snapshot: () => {
      const e = errors.splice(0, errors.length);
      const n = totalApiCalls;
      totalApiCalls = 0;
      return { errors: e, totalApiCalls: n };
    },
  };
}

// ==================== 页面名称检测 ====================
async function detectPageName(page) {
  try {
    await page.waitForTimeout(200);
    return await page.evaluate(() => {
      const isValid = (t) => {
        if (!t || t.length < 2 || t.length > 80) return false;
        if (/^[.•…·•\s]+$/.test(t)) return false;
        if (/^https?:\/\//i.test(t)) return false;
        if (/共\s*\d+\s*条|\d+\s*条\s*[\/／]\s*页|^\d+\s*条\/页$/.test(t)) return false;
        if (/^\d+$/.test(t)) return false;
        if (/[a-z]{9,}/i.test(t)) return false;
        return true;
      };
      const tryActive = ['.el-menu-item.is-active', '.el-menu-item.active', '.ant-menu-item-selected'];
      let activeItem = null;
      for (const s of tryActive) { activeItem = document.querySelector(s); if (activeItem) break; }
      if (activeItem) {
        const parts = [];
        const itemText = activeItem.textContent.replace(/\s+/g, ' ').trim();
        if (isValid(itemText)) parts.unshift(itemText);
        let el = activeItem.parentElement;
        const seen = new Set();
        while (el && el !== document.body) {
          const title = el.querySelector(':scope > .el-submenu__title, :scope > .ant-menu-submenu-title');
          if (title) { const tt = title.textContent.replace(/\s+/g, ' ').trim(); if (isValid(tt) && !seen.has(tt)) { seen.add(tt); parts.unshift(tt); } }
          if (el.tagName === 'LI' && (el.classList.contains('el-submenu') || el.classList.contains('ant-menu-submenu'))) {
            const st = el.querySelector(':scope > .el-submenu__title, :scope > .ant-menu-submenu-title');
            if (st) { const tt = st.textContent.replace(/\s+/g, ' ').trim(); if (isValid(tt) && !seen.has(tt)) { seen.add(tt); parts.unshift(tt); } }
          }
          el = el.parentElement;
        }
        if (parts.length > 0) return parts.join(' > ');
      }
      const bc = document.querySelector('.el-breadcrumb, .ant-breadcrumb');
      if (bc) {
        const items = bc.querySelectorAll('li, span, .el-breadcrumb__item, .ant-breadcrumb-link');
        const texts = Array.from(items).map(el => el.textContent.replace(/\s+/g, ' ').trim()).filter(Boolean);
        for (let i = texts.length - 1; i >= 0; i--) { if (isValid(texts[i])) return texts[i]; }
      }
      if (document.title) {
        const parts = document.title.split(/\s*[-|—]\s*/);
        for (const p of parts) { const t = p.replace(/\s+/g, ' ').trim(); if (isValid(t)) return t; }
      }
      return 'unknown-page';
    });
  } catch { return 'unknown-page'; }
}

function sanitizeFileName(name) {
  return (name || 'unknown').replace(/[\\/:*?"<>|]/g, '_').replace(/\s+/g, '-').substring(0, 40);
}

// ==================== 断言引擎 ====================
function parseAssertions(assertionList) {
  return (assertionList || []).map(a => {
    const text = typeof a === 'string' ? a.trim() : String(a);
    if (/页面正常/i.test(text)) return { type: 'page_ok', label: '页面正常' };
    if (/接口正常/i.test(text)) return { type: 'api_ok', label: '接口正常' };
    if (/控制台正常/i.test(text)) return { type: 'console_ok', label: '控制台正常' };
    if (/UI正常|渲染.*正常|页面.*渲染/i.test(text)) return { type: 'ui_ok', label: 'UI正常' };
    return { type: 'custom', label: text, text };
  });
}

async function evaluateAssertions(parsedAssertions, monitorSnapshot, agent, page) {
  const results = [];
  // 合并 UI 断言为一次 AI 调用，减少耗时
  const uiAssertions = parsedAssertions.filter(a => a.type === 'ui_ok');
  const otherAssertions = parsedAssertions.filter(a => a.type !== 'ui_ok');

  for (const a of otherAssertions) {
    switch (a.type) {
      case 'page_ok': {
        const pageErrors = monitorSnapshot.errors.filter(e => e.type === 'PAGE');
        results.push({ label: a.label, pass: pageErrors.length === 0, detail: pageErrors.length > 0 ? `${pageErrors.length}个JS异常` : '' });
        break;
      }
      case 'api_ok': {
        const apiErrors = monitorSnapshot.errors.filter(e => e.type === 'API');
        results.push({ label: a.label, pass: apiErrors.length === 0, detail: apiErrors.length > 0 ? `${apiErrors.length}个接口异常` : '' });
        break;
      }
      case 'console_ok': {
        const consoleErrors = monitorSnapshot.errors.filter(e => e.type === 'CONSOLE');
        results.push({ label: a.label, pass: consoleErrors.length === 0, detail: consoleErrors.length > 0 ? `${consoleErrors.length}个控制台错误` : '' });
        break;
      }
      case 'custom': {
        try {
          await agent.aiAssert(a.text, a.label);
          results.push({ label: a.label, pass: true, detail: '' });
        } catch (e) {
          results.push({ label: a.label, pass: false, detail: e.message?.substring(0, 100) || '断言失败' });
        }
        break;
      }
    }
  }

  // 批量执行 UI 断言（一次 AI 调用检查多个方面）
  if (uiAssertions.length > 0) {
    try {
      await agent.aiAssert('页面渲染正常，没有白屏、报错弹窗、布局错乱', '页面UI检查');
      for (const a of uiAssertions) {
        results.push({ label: a.label, pass: true, detail: '' });
      }
    } catch (e) {
      for (const a of uiAssertions) {
        results.push({ label: a.label, pass: false, detail: e.message?.substring(0, 100) || 'UI异常' });
      }
    }
  }

  return results;
}

// ==================== 场景执行 ====================
async function executeScenario(page, scenario, agent, config, resultDir, idx, total) {
  const tag = `[${idx}/${total}]`;
  console.log(`\n${tag} 📋 ${scenario.module} > ${scenario.name}`);
  const startTime = Date.now();
  const shotPrefix = path.join(resultDir, `${String(idx).padStart(2,'0')}-${sanitizeFileName(scenario.name)}`);

  await page.screenshot({ path: shotPrefix + '-before.png', fullPage: false }).catch(() => {});
  const monitor = startErrorMonitor(page);

  let stepOk = true;
  let stepError = '';
  let aiIntent = '';
  for (let si = 0; si < scenario.steps.length; si++) {
    const step = scenario.steps[si];
    aiIntent = step;
    try {
      console.log(`     步骤${si + 1}: ${step.substring(0, 50)}${step.length > 50 ? '...' : ''}`);
      await agent.aiAct(step);
      if (/加载完成|等待|点击|导航/i.test(step)) {
        await page.waitForLoadState('networkidle', { timeout: config.network_idle_wait }).catch(() => {});
        await wait(config.page_load_wait);
      } else {
        await wait(800);
      }
    } catch (e) {
      stepOk = false;
      const errMsg = e.message?.substring(0, 120) || '未知错误';
      stepError = `步骤${si + 1}失败: ${errMsg}`;
      // 增强错误信息：记录AI意图
      if (errMsg.includes('Replanned')) {
        stepError += ` [AI尝试寻找"${step}"但找不到目标元素]`;
      }
      console.log(`     ❌ ${stepError}`);
      break;
    }
  }

  if (stepOk) {
    await wait(1000);
    monitor.snapshot();
  }

  const monitorSnapshot = monitor.snapshot();
  monitor.stop();

  await wait(300);
  await page.screenshot({ path: shotPrefix + '-after.png', fullPage: false }).catch(() => {});

  const pageName = await detectPageName(page);
  const parsedAssertions = parseAssertions(scenario.assertions);
  const assertionResults = stepOk ? await evaluateAssertions(parsedAssertions, monitorSnapshot, agent, page) : [];

  const allPass = stepOk && assertionResults.every(r => r.pass);
  const elapsed = Date.now() - startTime;
  const result = {
    id: idx,
    name: scenario.name,
    module: scenario.module,
    priority: scenario.priority,
    tags: scenario.tags,
    pageName,
    result: allPass ? '通过' : '失败',
    stepError,
    aiIntent,
    assertions: assertionResults,
    errors: monitorSnapshot.errors,
    totalApiCalls: monitorSnapshot.totalApiCalls,
    screenshotBefore: shotPrefix + '-before.png',
    screenshotAfter: shotPrefix + '-after.png',
    elapsed,
    timestamp: new Date().toISOString(),
  };

  const icon = allPass ? '✅' : '❌';
  console.log(`     ${icon} ${result.result} (${(elapsed / 1000).toFixed(1)}s)`);
  if (!allPass) {
    const failures = assertionResults.filter(r => !r.pass);
    for (const f of failures) console.log(`       ⚠️ ${f.label}: ${f.detail}`);
  }
  return result;
}

// ==================== 报告生成: Excel ====================
async function generateResultExcel(results, envKey) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('巡检结果');
  [8, 20, 20, 10, 20, 10, 10, 12, 40].forEach((w, i) => ws.getColumn(i + 1).width = w);

  const headerRow = ws.addRow(['序号', '模块', '场景名称', '优先级', '页面名称', '结果', 'API调用', '耗时(s)', '异常详情']);
  headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } };
  headerRow.alignment = { horizontal: 'center', vertical: 'middle' };

  for (const r of results) {
    const errorDetail = [
      r.stepError,
      ...r.assertions.filter(a => !a.pass).map(a => `${a.label}: ${a.detail}`),
      ...r.errors.slice(0, 3).map(e => `[${e.type}] ${e.message || e.url || ''}`),
    ].filter(Boolean).join('; ');

    const row = ws.addRow([
      r.id, r.module, r.name, r.priority, r.pageName,
      r.result, r.totalApiCalls, (r.elapsed / 1000).toFixed(1), errorDetail,
    ]);
    if (r.result === '通过') {
      row.getCell(6).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2EFDA' } };
    } else {
      row.getCell(6).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFCE4EC' } };
      row.getCell(6).font = { color: { argb: 'FFC00000' }, bold: true };
    }
  }

  const summaryWs = wb.addWorksheet('汇总');
  const passCount = results.filter(r => r.result === '通过').length;
  const failCount = results.length - passCount;
  summaryWs.addRow(['巡检汇总']);
  summaryWs.addRow(['环境', envKey]);
  summaryWs.addRow(['时间', localTS()]);
  summaryWs.addRow(['总场景数', results.length]);
  summaryWs.addRow(['通过', passCount]);
  summaryWs.addRow(['失败', failCount]);
  summaryWs.addRow(['通过率', results.length > 0 ? ((passCount / results.length) * 100).toFixed(1) + '%' : 'N/A']);

  summaryWs.addRow([]);
  summaryWs.addRow(['模块', '总数', '通过', '失败']);
  const moduleStats = {};
  for (const r of results) {
    if (!moduleStats[r.module]) moduleStats[r.module] = { total: 0, pass: 0, fail: 0 };
    moduleStats[r.module].total++;
    if (r.result === '通过') moduleStats[r.module].pass++; else moduleStats[r.module].fail++;
  }
  for (const [mod, s] of Object.entries(moduleStats)) {
    summaryWs.addRow([mod, s.total, s.pass, s.fail]);
  }

  const excelPath = path.join(RESULT_BASE, `巡检结果-${localTS()}.xlsx`);
  await wb.xlsx.writeFile(excelPath);
  console.log(`📊 Excel: ${excelPath}`);
  return excelPath;
}

// ==================== 报告生成: HTML ====================
async function generateHtmlReport(results, envName, scenarioFiles) {
  const passCount = results.filter(r => r.result === '通过').length;
  const failCount = results.length - passCount;
  const passRate = results.length > 0 ? ((passCount / results.length) * 100).toFixed(1) : '0';

  const moduleStats = {};
  for (const r of results) {
    if (!moduleStats[r.module]) moduleStats[r.module] = { total: 0, pass: 0, fail: 0 };
    moduleStats[r.module].total++;
    if (r.result === '通过') moduleStats[r.module].pass++; else moduleStats[r.module].fail++;
  }

  const rows = results.map(r => {
    const cls = r.result === '通过' ? 'pass' : 'fail';
    const errors = [
      r.stepError,
      ...r.assertions.filter(a => !a.pass).map(a => `${a.label}: ${a.detail}`),
    ].filter(Boolean).join('<br>');
    return `<tr class="${cls}">
      <td>${r.id}</td><td>${r.module}</td><td>${r.name}</td><td>${r.priority}</td>
      <td>${r.pageName}</td><td class="result-${cls}">${r.result}</td>
      <td>${r.totalApiCalls}</td><td>${(r.elapsed/1000).toFixed(1)}s</td>
      <td>${errors || '-'}</td>
    </tr>`;
  }).join('\n');

  // 失败详情（增强版：包含AI意图）
  const failDetails = results.filter(r => r.result === '失败').map(r => {
    const intentInfo = r.aiIntent ? `<p style="margin:5px 0;font-size:13px;color:#666;"><b>AI意图：</b>尝试执行 "${r.aiIntent}"</p>` : '';
    const hint = r.stepError?.includes('找不到目标元素')
      ? '<p style="margin:5px 0;font-size:12px;color:#E65100;">💡 提示：菜单项可能已下线或名称变更，请检查 YAML 中的菜单名称是否与当前 UI 一致</p>'
      : '';
    return `<div style="margin:20px 0;padding:15px;background:#fff5f5;border-radius:8px;border-left:4px solid #C62828;">
        <h3 style="margin-top:0;color:#C62828;">${r.module} &gt; ${r.name}</h3>
        <p style="margin:5px 0;font-size:13px;color:#C62828;"><b>错误：</b>${r.stepError || '断言失败'}</p>
        ${intentInfo}
        ${hint}
        <div style="display:inline-block;margin-right:15px;"><p style="margin:5px 0;font-size:12px;color:#666;">执行前</p><a href="${path.basename(r.screenshotBefore)}" target="_blank"><img src="${path.basename(r.screenshotBefore)}" style="max-width:480px;border:1px solid #ddd;border-radius:4px;"></a></div>
        <div style="display:inline-block;"><p style="margin:5px 0;font-size:12px;color:#666;">执行后</p><a href="${path.basename(r.screenshotAfter)}" target="_blank"><img src="${path.basename(r.screenshotAfter)}" style="max-width:480px;border:1px solid #ddd;border-radius:4px;"></a></div>
      </div>`;
  }).join('\n');

  const moduleRows = Object.entries(moduleStats).map(([mod, s]) =>
    `<tr><td>${mod}</td><td>${s.total}</td><td>${s.pass}</td><td>${s.fail}</td>
     <td>${((s.pass/s.total)*100).toFixed(0)}%</td></tr>`
  ).join('\n');

  const html = `<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="UTF-8"><title>巡检报告 - ${envName}</title>
<style>
  body { font-family: -apple-system, 'Microsoft YaHei', sans-serif; margin: 20px; background: #f5f5f5; }
  .container { max-width: 1400px; margin: 0 auto; background: #fff; padding: 30px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }
  h1 { color: #333; border-bottom: 2px solid #4472C4; padding-bottom: 10px; }
  h2 { color: #4472C4; margin-top: 30px; }
  .summary { display: flex; gap: 20px; margin: 20px 0; }
  .summary-card { flex: 1; padding: 15px; border-radius: 8px; text-align: center; }
  .card-total { background: #E8EEF7; } .card-pass { background: #E2EFDA; } .card-fail { background: #FCE4EC; } .card-rate { background: #FFF3E0; }
  .summary-card .num { font-size: 32px; font-weight: bold; } .summary-card .label { color: #666; margin-top: 5px; }
  table { width: 100%; border-collapse: collapse; margin: 15px 0; font-size: 13px; }
  th { background: #4472C4; color: #fff; padding: 10px 8px; text-align: left; }
  td { padding: 8px; border-bottom: 1px solid #eee; }
  tr:hover { background: #f0f5ff; }
  .result-pass { color: #2E7D32; font-weight: bold; } .result-fail { color: #C62828; font-weight: bold; }
  .meta { color: #888; font-size: 13px; }
</style></head><body><div class="container">
  <h1>🔍 平台巡检报告 — ${envName}</h1>
  <p class="meta">生成时间: ${localTS()} | 场景文件: ${scenarioFiles.join(', ')}</p>
  <div class="summary">
    <div class="summary-card card-total"><div class="num">${results.length}</div><div class="label">总场景</div></div>
    <div class="summary-card card-pass"><div class="num">${passCount}</div><div class="label">通过</div></div>
    <div class="summary-card card-fail"><div class="num">${failCount}</div><div class="label">失败</div></div>
    <div class="summary-card card-rate"><div class="num">${passRate}%</div><div class="label">通过率</div></div>
  </div>
  <h2>模块汇总</h2>
  <table><tr><th>模块</th><th>总数</th><th>通过</th><th>失败</th><th>通过率</th></tr>${moduleRows}</table>
  <h2>巡检明细</h2>
  <table><tr><th>#</th><th>模块</th><th>场景</th><th>优先级</th><th>页面</th><th>结果</th><th>API</th><th>耗时</th><th>异常</th></tr>${rows}</table>
  ${failDetails ? `<h2>❌ 失败场景详情</h2>${failDetails}` : ''}
</div></body></html>`;

  const htmlPath = path.join(RESULT_BASE, `巡检报告-${localTS()}.html`);
  fs.writeFileSync(htmlPath, html, 'utf-8');
  console.log(`🌐 HTML: ${htmlPath}`);
  return htmlPath;
}

// ==================== 报告生成: Markdown ====================
async function generateMarkdownReport(results, envName, scenarioFiles) {
  const passCount = results.filter(r => r.result === '通过').length;
  const failCount = results.length - passCount;
  const passRate = results.length > 0 ? ((passCount / results.length) * 100).toFixed(1) : '0';

  let md = `# 平台巡检报告 — ${envName}\n\n`;
  md += `> 生成时间: ${localTS()} | 场景文件: ${scenarioFiles.join(', ')}\n\n`;
  md += `## 巡检概览\n\n| 指标 | 数值 |\n|---|---|\n`;
  md += `| 总场景 | ${results.length} |\n| 通过 | ${passCount} |\n| 失败 | ${failCount} |\n| 通过率 | ${passRate}% |\n\n`;
  md += `## 巡检明细\n\n| # | 模块 | 场景 | 优先级 | 页面 | 结果 | 耗时 | 异常 |\n`;
  md += `|---|---|---|---|---|---|---|---|\n`;
  for (const r of results) {
    const icon = r.result === '通过' ? '✅' : '❌';
    const errors = [r.stepError, ...r.assertions.filter(a => !a.pass).map(a => `${a.label}: ${a.detail}`)].filter(Boolean).join('; ');
    md += `| ${r.id} | ${r.module} | ${r.name} | ${r.priority} | ${r.pageName} | ${icon}${r.result} | ${(r.elapsed/1000).toFixed(1)}s | ${errors || '-'} |\n`;
  }

  const mdPath = path.join(RESULT_BASE, `巡检报告-${localTS()}.md`);
  fs.writeFileSync(mdPath, md, 'utf-8');
  console.log(`📄 Markdown: ${mdPath}`);
  return mdPath;
}

// ==================== 主流程（CI 模式，无交互） ====================
async function main() {
  console.log('╔══════════════════════════════════════╗');
  console.log('║   🔍 平台自动化巡检 — CI/CD 模式     ║');
  console.log('╚══════════════════════════════════════╝');

  // 1. 环境选择（从环境变量读取）
  const envKey = ENVIRONMENTS[CI_ENV_KEY] ? CI_ENV_KEY : Object.keys(ENVIRONMENTS)[0];
  const env = ENVIRONMENTS[envKey];
  console.log(`\n📡 环境: ${envKey} — ${env.name} → ${env.admin.startUrl}`);
  console.log(`🖥️  Headless: ${CI_HEADLESS}`);

  // 2. 加载菜单结构
  const menuFiles = listMenuTreeFiles();
  let menuTree = '';
  if (menuFiles.length > 0) {
    menuTree = loadMenuTree(menuFiles[0]);
    console.log(`📂 菜单结构: ${menuFiles[0]} (${menuTree.length} chars)`);
  }

  // 3. 加载场景文件
  const scenarioFiles = listScenarioFiles();
  if (scenarioFiles.length === 0) {
    console.error('❌ scenarios/ 目录下未找到场景文件 (.yml)');
    process.exit(1);
  }

  let selectedFiles = [];
  if (CI_SCENARIO_FILES === 'all') {
    selectedFiles = scenarioFiles;
  } else if (CI_SCENARIO_FILES.includes(',')) {
    const idxs = CI_SCENARIO_FILES.split(',').map(s => parseInt(s.trim())).filter(n => !isNaN(n));
    selectedFiles = idxs.map(i => scenarioFiles[i - 1]).filter(Boolean);
  } else if (/^\d+$/.test(CI_SCENARIO_FILES)) {
    const idx = parseInt(CI_SCENARIO_FILES) - 1;
    selectedFiles = [scenarioFiles[idx >= 0 && idx < scenarioFiles.length ? idx : 0]];
  } else {
    // 按文件名匹配
    selectedFiles = scenarioFiles.filter(f => f.includes(CI_SCENARIO_FILES));
    if (selectedFiles.length === 0) selectedFiles = scenarioFiles;
  }
  console.log(`📂 场景文件: ${selectedFiles.join(', ')}`);

  // 4. 分别加载 Task 和 Scenario 两种格式的文件
  const fileFormats = selectedFiles.map(f => ({
    file: f,
    format: detectYamlFormat(path.join(__dirname, 'scenarios', f)),
  }));

  let allScenarios = [];
  let allTasks = [];
  let mergedConfig = { ...DEFAULT_CONFIG };

  for (const { file, format } of fileFormats) {
    if (format === 'task') {
      const loaded = loadTaskYaml(path.join(__dirname, 'scenarios', file));
      if (loaded) {
        allTasks = allTasks.concat(loaded.tasks);
        console.log(`📝 Task格式: ${file} → ${loaded.tasks.length} 个任务`);
      }
    } else if (format === 'scenario') {
      const loaded = loadScenario(path.join(__dirname, 'scenarios', file));
      if (loaded) {
        mergedConfig = { ...mergedConfig, ...loaded.config };
        allScenarios = allScenarios.concat(loaded.scenarios);
        console.log(`📝 Scenario格式: ${file} → ${loaded.scenarios.length} 个场景`);
      }
    }
  }

  // 5. 过滤场景（优先级/模块，仅 Scenario 格式）
  let selectedScenarios = allScenarios;
  if (CI_PRIORITY) {
    selectedScenarios = selectedScenarios.filter(s => s.priority.toUpperCase() === CI_PRIORITY.toUpperCase());
    console.log(`🔍 按优先级 ${CI_PRIORITY} 过滤: ${selectedScenarios.length} 个场景`);
  }
  if (CI_MODULE) {
    selectedScenarios = selectedScenarios.filter(s => s.module.includes(CI_MODULE));
    console.log(`🔍 按模块 "${CI_MODULE}" 过滤: ${selectedScenarios.length} 个场景`);
  }

  const totalItems = selectedScenarios.length + allTasks.length;
  if (totalItems === 0) { console.log('没有可用的巡检项'); process.exit(0); }
  if (selectedScenarios.length > 0) console.log(`✅ 将执行 ${selectedScenarios.length} 个 Scenario 场景`);
  if (allTasks.length > 0) console.log(`✅ 将执行 ${allTasks.length} 个 Task 任务`);

  // 6. 创建结果目录
  RESULT_BASE = path.join(__dirname, 'results', `CI巡检-${envKey}-${localTS()}`);
  if (!fs.existsSync(RESULT_BASE)) fs.mkdirSync(RESULT_BASE, { recursive: true });

  // 7. 启动浏览器（headless 模式）
  console.log('\n🌐 启动浏览器 (headless)...');
  const browser = await chromium.launch({
    headless: CI_HEADLESS,
    args: ['--ignore-certificate-errors', '--no-sandbox', '--disable-setuid-sandbox'],
  });
  const context = await browser.newContext({
    viewport: { width: 1401, height: 900 },
    ignoreHTTPSErrors: true,
  });
  const page = await context.newPage();

  // 8. 导航
  console.log(`🔗 导航到: ${env.admin.startUrl}`);
  await page.goto(env.admin.startUrl, { waitUntil: 'domcontentloaded', timeout: 30000, ignoreHTTPSErrors: true }).catch(() => {});
  await wait(3000);

  // 9. 登录检测（CI 模式使用 Cookie 或 Token）
  const currentUrl = page.url();
  const COOKIE_FILE = path.join(__dirname, '.cache', `cookies-inspection-${envKey}.json`);

  if (currentUrl.includes('login') || currentUrl.includes('white')) {
    // 尝试使用 Cookie 恢复
    if (fs.existsSync(COOKIE_FILE)) {
      await context.addCookies(JSON.parse(fs.readFileSync(COOKIE_FILE, 'utf-8')));
      console.log('🍪 尝试 Cookie 恢复登录...');
      await page.reload({ waitUntil: 'domcontentloaded' }).catch(() => {});
      await wait(3000);
    }

    // 再次检查
    const url2 = page.url();
    if (url2.includes('login') || url2.includes('white')) {
      if (!CI_HEADLESS) {
        // 非无头模式：等待手动登录
        console.log('\n⏳ 请在浏览器中手动登录...（超时10分钟）');
        try {
          await page.waitForURL(u => !u.href.includes('login') && !u.href.includes('white'), { timeout: 600000 });
          console.log('✅ 登录成功');
          // 保存 Cookie
          if (!fs.existsSync(path.join(__dirname, '.cache'))) fs.mkdirSync(path.join(__dirname, '.cache'), { recursive: true });
          fs.writeFileSync(COOKIE_FILE, JSON.stringify(await context.cookies(), null, 2));
          console.log('🍪 Cookie 已保存');
        } catch {
          console.log('❌ 登录超时');
          await browser.close().catch(() => {});
          process.exit(1);
        }
      } else {
        console.error('❌ 无法自动登录，CI 无头模式不支持手动登录');
        console.error('   请先用 HEADLESS=false 运行一次完成登录，Cookie 会自动保存');
        await browser.close().catch(() => {});
        process.exit(1);
      }
    }
  }
  console.log('✅ 已登录');

  // 10. 创建 AI Agent
  const agent = new PlaywrightAgent(page, {
    ...(menuTree ? { aiActionContext: `当前平台菜单结构:\n${menuTree}` } : {}),
    generateReport: true,
    reportFileName: `midscene-report-${new Date().toISOString().replace(/[:.]/g, '-')}`,
    groupName: `CI巡检 — ${env.name}`,
    groupDescription: `环境: ${envKey} | 场景文件: ${selectedFiles.join(', ')}`,
  });

  // 11. 串行执行巡检（先 Scenario，后 Task）
  const results = [];
  const startTime = Date.now();
  let globalIdx = 0;

  // 先执行 Scenario 格式
  if (selectedScenarios.length > 0) {
    console.log(`\n🚀 开始 Scenario 巡检 (${selectedScenarios.length} 个场景)\n`);
    for (let i = 0; i < selectedScenarios.length; i++) {
      const scenario = selectedScenarios[i];
      globalIdx++;
      try {
        const result = await executeScenario(page, scenario, agent, mergedConfig, RESULT_BASE, globalIdx, totalItems);
        results.push(result);
      } catch (e) {
        console.log(`     ❌ 场景执行异常: ${e.message?.substring(0, 80)}`);
        results.push({
          id: globalIdx, name: scenario.name, module: scenario.module, priority: scenario.priority,
          tags: scenario.tags, pageName: 'error', result: '失败',
          stepError: e.message?.substring(0, 200), aiIntent: '', assertions: [], errors: [],
          totalApiCalls: 0, screenshotBefore: '', screenshotAfter: '',
          elapsed: 0, timestamp: new Date().toISOString(),
        });
      }
    }
  }

  // 再执行 Task 格式
  if (allTasks.length > 0) {
    console.log(`\n🚀 开始 Task 执行 (${allTasks.length} 个任务)\n`);
    for (let i = 0; i < allTasks.length; i++) {
      const task = allTasks[i];
      globalIdx++;
      try {
        const result = await executeTaskFlow(page, task, agent, mergedConfig, RESULT_BASE, globalIdx, totalItems);
        results.push(result);
      } catch (e) {
        console.log(`     ❌ 任务执行异常: ${e.message?.substring(0, 80)}`);
        results.push({
          id: globalIdx, name: task.name, module: 'YAML任务', priority: 'P1', tags: [],
          pageName: 'error', result: '失败', stepError: e.message?.substring(0, 200), aiIntent: '',
          assertions: [], errors: [], totalApiCalls: 0,
          screenshotBefore: '', screenshotAfter: '',
          elapsed: 0, timestamp: new Date().toISOString(),
        });
      }
    }
  }

  const totalTime = Date.now() - startTime;
  const passCount = results.filter(r => r.result === '通过').length;
  console.log(`\n${'='.repeat(50)}`);
  console.log(`✅ 巡检完成! ${passCount}/${results.length} 通过, 总耗时 ${(totalTime/1000/60).toFixed(1)} 分钟`);

  // 12. 生成报告
  console.log('\n📊 正在生成报告...');
  await generateResultExcel(results, envKey);
  await generateHtmlReport(results, env.name, selectedFiles);
  await generateMarkdownReport(results, env.name, selectedFiles);

  // 13. 复制报告到 output/latest/
  const outputDir = path.join(__dirname, 'output', 'latest');
  if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, { recursive: true });
  // 清空旧的 latest 报告
  fs.readdirSync(outputDir).filter(f => f !== '.gitkeep').forEach(f => fs.unlinkSync(path.join(outputDir, f)));
  // 复制新报告
  if (fs.existsSync(RESULT_BASE)) {
    fs.readdirSync(RESULT_BASE).forEach(f => {
      fs.copyFileSync(path.join(RESULT_BASE, f), path.join(outputDir, f));
    });
    console.log(`📂 报告已复制到: ${outputDir}`);
  }

  console.log(`\n📁 结果目录: ${RESULT_BASE}`);
  console.log(`📁 输出目录: ${outputDir}`);

  // CI 模式：直接关闭，不等待用户输入
  await browser.close().catch(() => {});

  // 如果有失败，以非零退出码退出（让 CI 知道失败了）
  if (passCount < results.length) {
    console.log(`\n⚠️ 有 ${results.length - passCount} 个场景失败`);
    process.exit(1);
  }
}

main().catch(async (err) => {
  console.error('FATAL: ' + err.message);
  console.error(err.stack);
  process.exit(1);
});
