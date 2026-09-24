/**
 * inspection.mjs
 * 平台自动化巡检脚本 — Playwright + Midscene AI
 *
 * 功能:
 *   - YAML 场景驱动，非技术人员可维护
 *   - 全量菜单遍历 + 指定页面巡检
 *   - 每页截图 + 错误监控（API/Console/PageError）
 *   - 多格式报告输出（Excel/Word/HTML/Markdown）
 *
 * 数据文件: scenarios/*.yml
 * 环境配置: data/0 shared/environments.json（公用）
 *
 * 用法: node inspection.mjs
 */

import { PlaywrightAgent } from '@midscene/web';
import { chromium } from 'playwright';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import fs from 'node:fs';
import readline from 'node:readline';

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

// ==================== 工具函数 ====================
const localTS = () => {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}T${String(d.getHours()).padStart(2,'0')}-${String(d.getMinutes()).padStart(2,'0')}-${String(d.getSeconds()).padStart(2,'0')}`;
};
const wait = (ms) => new Promise((r) => setTimeout(r, ms));
function ask(q) {
  return new Promise(r => {
    const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
    rl.question(q, a => { rl.close(); r(a.trim()); });
  });
}

// ==================== 配置 ====================
let RESULT_BASE = path.join(__dirname, 'results', localTS());
const DEFAULT_CONFIG = {
  page_load_wait: 3000,
  page_load_timeout: 30000,
  network_idle_wait: 5000,
  screenshot_quality: 'medium',
  record_network: true,
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
  return fs.readdirSync(caseDir)
    .filter(f => f.endsWith('.txt') && f.includes('菜单结构'));
}

function loadMenuTree(selectedFile) {
  const caseDir = path.join(__dirname, 'data', 'platform-inspection-case');
  const fp = path.join(caseDir, selectedFile);
  if (fs.existsSync(fp)) return fs.readFileSync(fp, 'utf-8').trim();
  return '';
}

// ==================== YAML 场景加载 ====================
function listScenarioFiles() {
  const dir = path.join(__dirname, 'scenarios');
  if (!fs.existsSync(dir)) return [];
  return fs.readdirSync(dir).filter(f => (f.endsWith('.yml') || f.endsWith('.yaml')) && !f.startsWith('~$'));
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

// ==================== Task YAML 格式加载 ====================
// 支持 web: + tasks[].flow[] 格式的 YAML 用例
// flow 步骤: aiTap / aiInput / aiAction / aiAssert / aiWaitFor / sleep
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

// 自动检测 YAML 格式
function detectYamlFormat(filePath) {
  const raw = fs.readFileSync(filePath, 'utf-8');
  const content = raw.replace(/^\s*#[^\n]*/gm, ''); // 去除注释行
  if (/^tasks:/m.test(content) && /^web:/m.test(content)) return 'task';
  if (/^scenarios:/m.test(content)) return 'scenario';
  return 'unknown';
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
      // 过滤已知的环境预期报错：WebSocket 连接失败
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
    await page.waitForTimeout(600);
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
      // 1. 左侧激活菜单
      const tryActive = ['.el-menu-item.is-active', '.el-menu-item.active', '.ant-menu-item-selected', 'li[class*="active"][class*="menu"]', '[class*="nav-item"][class*="active"]'];
      let activeItem = null;
      for (const s of tryActive) { activeItem = document.querySelector(s); if (activeItem) break; }
      if (activeItem) {
        const parts = [];
        const itemText = activeItem.textContent.replace(/\s+/g, ' ').trim();
        if (isValid(itemText)) parts.unshift(itemText);
        let el = activeItem.parentElement;
        const seen = new Set();
        while (el && el !== document.body) {
          const title = el.querySelector(':scope > .el-submenu__title, :scope > .ant-menu-submenu-title, :scope > [class*="submenu"][class*="title"]');
          if (title) { const tt = title.textContent.replace(/\s+/g, ' ').trim(); if (isValid(tt) && !seen.has(tt)) { seen.add(tt); parts.unshift(tt); } }
          if (el.tagName === 'LI' && (el.classList.contains('el-submenu') || el.classList.contains('ant-menu-submenu'))) {
            const st = el.querySelector(':scope > .el-submenu__title, :scope > .ant-menu-submenu-title');
            if (st) { const tt = st.textContent.replace(/\s+/g, ' ').trim(); if (isValid(tt) && !seen.has(tt)) { seen.add(tt); parts.unshift(tt); } }
          }
          el = el.parentElement;
        }
        if (parts.length > 0) return parts.join(' > ');
      }
      // 2. 面包屑
      const bc = document.querySelector('.el-breadcrumb, .ant-breadcrumb, [class*="breadcrumb"]');
      if (bc) {
        const items = bc.querySelectorAll('li, span, .el-breadcrumb__item, .ant-breadcrumb-link');
        const texts = Array.from(items).map(el => el.textContent.replace(/\s+/g, ' ').trim()).filter(Boolean);
        for (let i = texts.length - 1; i >= 0; i--) { if (isValid(texts[i])) return texts[i]; }
      }
      // 3. document.title
      if (document.title) {
        const parts = document.title.split(/\s*[-|—]\s*/);
        for (const p of parts) { const t = p.replace(/\s+/g, ' ').trim(); if (isValid(t)) return t; }
      }
      // 4. URL path
      const pathSeg = window.location.pathname.split('/').filter(Boolean).pop();
      if (pathSeg) { const t = decodeURIComponent(pathSeg); if (t.length >= 2 && t.length <= 60) return t; }
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
  for (const a of parsedAssertions) {
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
      case 'ui_ok': {
        try {
          await agent.aiAssert('页面渲染正常，没有白屏、报错弹窗、布局错乱', '页面UI检查');
          results.push({ label: a.label, pass: true, detail: '' });
        } catch (e) {
          results.push({ label: a.label, pass: false, detail: e.message?.substring(0, 100) || 'UI异常' });
        }
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
  return results;
}

// ==================== 场景执行 ====================
async function executeScenario(page, scenario, agent, config, resultDir, idx, total) {
  const tag = `[${idx}/${total}]`;
  console.log(`\n${tag} 📋 ${scenario.module} > ${scenario.name}`);
  const startTime = Date.now();
  const shotPrefix = path.join(resultDir, `${String(idx).padStart(2,'0')}-${sanitizeFileName(scenario.name)}`);

  // 执行前截图
  await page.screenshot({ path: shotPrefix + '-before.png', fullPage: false }).catch(() => {});

  // 启动错误监控
  const monitor = startErrorMonitor(page);

  // 执行步骤
  let stepOk = true;
  let stepError = '';
  for (let si = 0; si < scenario.steps.length; si++) {
    const step = scenario.steps[si];
    try {
      console.log(`     步骤${si + 1}: ${step.substring(0, 50)}${step.length > 50 ? '...' : ''}`);
      await agent.aiAct(step);
      // 等待页面稳定
      if (/加载完成|等待|点击|导航/i.test(step)) {
        await page.waitForLoadState('networkidle', { timeout: config.network_idle_wait }).catch(() => {});
        await wait(config.page_load_wait);
      } else {
        await wait(1500);
      }
    } catch (e) {
      stepOk = false;
      stepError = `步骤${si + 1}失败: ${e.message?.substring(0, 80)}`;
      console.log(`     ❌ ${stepError}`);
      break;
    }
  }

  // 页面稳定后，清除导航期间从上一页残留的错误（如后台长轮询API的延迟响应）
  // 这些错误属于上一页，不应归到当前场景
  if (stepOk) {
    await wait(1500);
    monitor.snapshot(); // 丢弃导航期间累积的错误
  }

  // 获取监控快照（仅包含目标页面加载后产生的错误）
  const monitorSnapshot = monitor.snapshot();
  monitor.stop();

  // 执行后截图
  await wait(500);
  await page.screenshot({ path: shotPrefix + '-after.png', fullPage: false }).catch(() => {});

  // 检测页面名称
  const pageName = await detectPageName(page);

  // 断言评估
  const parsedAssertions = parseAssertions(scenario.assertions);
  const assertionResults = stepOk ? await evaluateAssertions(parsedAssertions, monitorSnapshot, agent, page) : [];

  // 汇总结果
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

// ==================== Task YAML Flow 执行 ====================
// 解析环境变量 ${VAR}
function resolveEnvVar(text) {
  if (typeof text !== 'string') return String(text ?? '');
  return text.replace(/\$\{([^}]+)\}/g, (_, key) => process.env[key] || '');
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
    if (typeof desc === 'object' && desc !== null) {
      // 处理多行 YAML dict 格式: aiAction: { key: "value" }
      const vals = Object.values(desc);
      desc = vals.length > 0 ? String(vals[0]) : '';
    }
    desc = resolveEnvVar(String(desc));
    const display = desc.length > 60 ? desc.substring(0, 60) + '...' : desc;
    console.log(`     步骤${stepIdx}: 执行 [${display}]`);
    await agent.aiAct(desc);
    await wait(1500);

  } else if ('aiAssert' in step) {
    const assertion = resolveEnvVar(String(step.aiAssert));
    const display = assertion.length > 60 ? assertion.substring(0, 60) + '...' : assertion;
    console.log(`     步骤${stepIdx}: 断言 [${display}]`);
    await agent.aiAssert(assertion);

  } else if ('aiWaitFor' in step) {
    const condition = resolveEnvVar(String(step.aiWaitFor));
    const timeout = step.timeout || 10000;
    const display = condition.length > 60 ? condition.substring(0, 60) + '...' : condition;
    console.log(`     步骤${stepIdx}: 等待 [${display}] (超时${timeout}ms)`);
    try {
      await agent.aiWaitFor(condition, { timeoutMs: timeout });
    } catch {
      // 降级：轮询 aiAssert
      const maxRetries = Math.ceil(timeout / 2000);
      let lastErr;
      for (let i = 0; i < maxRetries; i++) {
        try { await agent.aiAssert(condition); break; }
        catch (e) { lastErr = e; if (i === maxRetries - 1) throw lastErr; }
        await wait(2000);
      }
    }

  } else if ('sleep' in step) {
    const ms = parseInt(step.sleep);
    console.log(`     步骤${stepIdx}: 等待 ${ms}ms`);
    await wait(ms);
  }
}

// 执行 Task YAML 格式的单个任务
async function executeTaskScenario(page, task, agent, config, resultDir, idx, total) {
  const tag = `[${idx}/${total}]`;
  console.log(`\n${tag} 📋 ${task.name}`);
  const startTime = Date.now();
  const shotPrefix = path.join(resultDir, `${String(idx).padStart(2,'0')}-${sanitizeFileName(task.name)}`);

  // 执行前截图
  await page.screenshot({ path: shotPrefix + '-before.png', fullPage: false }).catch(() => {});

  // 启动错误监控
  const monitor = startErrorMonitor(page);

  // 逐步执行 flow
  let stepOk = true;
  let stepError = '';
  const stepDetails = [];

  for (let si = 0; si < task.flow.length; si++) {
    const step = task.flow[si];
    if (!step) continue;
    try {
      await executeFlowStep(page, agent, step, si + 1);
      stepDetails.push({ index: si + 1, status: 'ok' });
      // 点击/导航类操作后等待页面稳定
      if (step.aiTap || step.aiAction) {
        await page.waitForLoadState('networkidle', { timeout: config.network_idle_wait }).catch(() => {});
      }
    } catch (e) {
      stepOk = false;
      stepError = `步骤${si + 1}失败: ${e.message?.substring(0, 120)}`;
      stepDetails.push({ index: si + 1, status: 'fail', error: stepError });
      console.log(`     ❌ ${stepError}`);
      break;
    }
  }

  // 清除导航期间的残留错误
  if (stepOk) {
    await wait(1500);
    monitor.snapshot();
  }

  const monitorSnapshot = monitor.snapshot();
  monitor.stop();

  // 执行后截图
  await wait(500);
  await page.screenshot({ path: shotPrefix + '-after.png', fullPage: false }).catch(() => {});

  const pageName = await detectPageName(page);
  const allPass = stepOk;
  const elapsed = Date.now() - startTime;

  const result = {
    id: idx,
    name: task.name,
    module: 'YAML任务',
    priority: 'P1',
    tags: [],
    pageName,
    result: allPass ? '通过' : '失败',
    stepError,
    stepDetails,
    assertions: [],
    errors: monitorSnapshot.errors,
    totalApiCalls: monitorSnapshot.totalApiCalls,
    screenshotBefore: shotPrefix + '-before.png',
    screenshotAfter: shotPrefix + '-after.png',
    elapsed,
    timestamp: new Date().toISOString(),
  };

  const icon = allPass ? '✅' : '❌';
  console.log(`     ${icon} ${result.result} (${(elapsed / 1000).toFixed(1)}s)`);
  return result;
}

// ==================== 中断处理 ====================
let PARTIAL_RESULTS = [];
let PARTIAL_ENV = null;
let PARTIAL_ENV_KEY = '';
let PARTIAL_FILES = [];
let PARTIAL_CONFIG = DEFAULT_CONFIG;
let PARTIAL_GENERATING = false;

// ==================== Midscene 报告复制 ====================
function copyMidsceneReport() {
  try {
    const midsceneDir = path.join(__dirname, 'midscene_run', 'report');
    if (!fs.existsSync(midsceneDir)) return;
    const midFiles = fs.readdirSync(midsceneDir)
      .filter(f => f.startsWith('midscene-report-') && f.endsWith('.html'))
      .sort().reverse();
    if (midFiles.length > 0) {
      const src = path.join(midsceneDir, midFiles[0]);
      const dest = path.join(RESULT_BASE, `Midscene巡检报告-${localTS()}.html`);
      fs.copyFileSync(src, dest);
      console.log(`🌐 Midscene报告: ${dest}`);
    }
  } catch (e) { console.log('⚠️ 复制Midscene报告失败:', e.message); }
}

async function handleInterrupt() {
  if (PARTIAL_GENERATING || PARTIAL_RESULTS.length === 0) { process.exit(0); }
  PARTIAL_GENERATING = true;
  console.log('\n\n⚠️ 脚本被中断，正在生成部分巡检报告...');
  try {
    const excelPath = await generateResultExcel(PARTIAL_RESULTS, PARTIAL_ENV_KEY);
    await generateWordReport(PARTIAL_RESULTS, PARTIAL_ENV?.name || '', PARTIAL_FILES, excelPath);
    await generateHtmlReport(PARTIAL_RESULTS, PARTIAL_ENV?.name || '', PARTIAL_FILES);
    await generateMarkdownReport(PARTIAL_RESULTS, PARTIAL_ENV?.name || '', PARTIAL_FILES);
    copyMidsceneReport();
  } catch (e) { console.error('生成报告失败:', e.message); }
  console.log('报告已生成，退出。');
  process.exit(0);
}
process.on('SIGINT', () => { handleInterrupt(); });
process.on('SIGTERM', () => { handleInterrupt(); });

// ==================== 报告生成: Excel ====================
async function generateResultExcel(results, envKey) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('巡检结果');
  [8, 20, 20, 10, 20, 10, 10, 12, 40].forEach((w, i) => ws.getColumn(i + 1).width = w);

  // 表头
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
    // 结果着色
    if (r.result === '通过') {
      row.getCell(6).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2EFDA' } };
    } else {
      row.getCell(6).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFCE4EC' } };
      row.getCell(6).font = { color: { argb: 'FFC00000' }, bold: true };
    }
  }

  // 汇总 sheet
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

  // 按模块汇总
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

// ==================== 报告生成: Word ====================
async function generateWordReport(results, envName, scenarioFiles, excelPath) {
  const passCount = results.filter(r => r.result === '通过').length;
  const failCount = results.length - passCount;
  const passRate = results.length > 0 ? ((passCount / results.length) * 100).toFixed(1) : '0';

  const children = [];
  // 标题
  children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(`平台巡检报告 — ${envName}`)] }));
  children.push(new Paragraph({ children: [new TextRun({ text: `生成时间: ${localTS()}`, color: '666666', size: 20 })] }));
  children.push(new Paragraph({ children: [new TextRun({ text: `场景文件: ${scenarioFiles.join(', ')}`, color: '666666', size: 20 })] }));

  // 概览
  children.push(new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun('巡检概览')] }));
  children.push(new Paragraph({ children: [new TextRun(`总场景: ${results.length}  |  通过: ${passCount}  |  失败: ${failCount}  |  通过率: ${passRate}%`)] }));

  // 按模块汇总
  children.push(new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun('模块汇总')] }));
  const moduleStats = {};
  for (const r of results) {
    if (!moduleStats[r.module]) moduleStats[r.module] = { total: 0, pass: 0, fail: 0 };
    moduleStats[r.module].total++;
    if (r.result === '通过') moduleStats[r.module].pass++; else moduleStats[r.module].fail++;
  }
  for (const [mod, s] of Object.entries(moduleStats)) {
    children.push(new Paragraph({ children: [new TextRun(`${mod}: ${s.total}个场景, 通过${s.pass}, 失败${s.fail}`)] }));
  }

  // 失败详情（含截图）
  const failedResults = results.filter(r => r.result === '失败');
  if (failedResults.length > 0) {
    children.push(new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun('失败场景详情')] }));
    for (const r of failedResults) {
      children.push(new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun(`${r.module} > ${r.name}`)] }));
      if (r.stepError) children.push(new Paragraph({ children: [new TextRun({ text: `步骤错误: ${r.stepError}`, color: 'CC0000' })] }));
      for (const a of r.assertions.filter(a => !a.pass)) {
        children.push(new Paragraph({ children: [new TextRun({ text: `断言失败 [${a.label}]: ${a.detail}`, color: 'CC0000' })] }));
      }
      if (r.errors.length > 0) {
        children.push(new Paragraph({ children: [new TextRun(`异常记录 (${r.errors.length}条):`)] }));
        for (const e of r.errors.slice(0, 5)) {
          children.push(new Paragraph({ children: [new TextRun({ text: `  [${e.type}] ${e.message || e.url || ''}`, size: 18 })] }));
        }
      }
      // 嵌入截图
      const beforeImg = r.screenshotBefore && fs.existsSync(r.screenshotBefore) ? fs.readFileSync(r.screenshotBefore) : null;
      const afterImg = r.screenshotAfter && fs.existsSync(r.screenshotAfter) ? fs.readFileSync(r.screenshotAfter) : null;
      if (beforeImg || afterImg) {
        children.push(new Paragraph({ children: [new TextRun({ text: '页面截图:', bold: true, size: 20 })] }));
        const imgChildren = [];
        if (beforeImg) {
          imgChildren.push(new ImageRun({ data: beforeImg, transformation: { width: 450, height: 253 }, type: 'png' }));
          imgChildren.push(new TextRun({ text: '  执行前  ', size: 18, color: '666666' }));
        }
        if (afterImg) {
          imgChildren.push(new ImageRun({ data: afterImg, transformation: { width: 450, height: 253 }, type: 'png' }));
          imgChildren.push(new TextRun({ text: '  执行后', size: 18, color: '666666' }));
        }
        children.push(new Paragraph({ children: imgChildren }));
      }
    }
  }

  // 全部结果列表
  children.push(new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun('巡检结果明细')] }));
  for (const r of results) {
    const icon = r.result === '通过' ? '✅' : '❌';
    children.push(new Paragraph({ children: [new TextRun(`${icon} #${r.id} ${r.module} > ${r.name} — ${r.pageName} (${(r.elapsed/1000).toFixed(1)}s)`)] }));
  }

  const doc = new Document({ sections: [{ children }] });
  const buffer = await Packer.toBuffer(doc);
  const wordPath = path.join(RESULT_BASE, `巡检报告-${localTS()}.docx`);
  fs.writeFileSync(wordPath, buffer);
  console.log(`📝 Word: ${wordPath}`);
  return wordPath;
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
    // 失败场景附加截图缩略图
    let shots = '-';
    if (r.result === '失败') {
      const b = r.screenshotBefore ? path.basename(r.screenshotBefore) : '';
      const a = r.screenshotAfter ? path.basename(r.screenshotAfter) : '';
      const thumbs = [];
      if (b) thumbs.push(`<a href="${b}" target="_blank"><img src="${b}" width="120" title="执行前"></a>`);
      if (a) thumbs.push(`<a href="${a}" target="_blank"><img src="${a}" width="120" title="执行后"></a>`);
      shots = thumbs.join(' ');
    }
    return `<tr class="${cls}">
      <td>${r.id}</td><td>${r.module}</td><td>${r.name}</td><td>${r.priority}</td>
      <td>${r.pageName}</td><td class="result-${cls}">${r.result}</td>
      <td>${r.totalApiCalls}</td><td>${(r.elapsed/1000).toFixed(1)}s</td>
      <td>${errors || '-'}</td><td>${shots}</td>
    </tr>`;
  }).join('\n');

  // 失败详情截图区
  let failDetails = '';
  const failedResults = results.filter(r => r.result === '失败');
  if (failedResults.length > 0) {
    failDetails = `<h2>❌ 失败场景截图详情</h2>`;
    for (const r of failedResults) {
      const b = r.screenshotBefore ? path.basename(r.screenshotBefore) : '';
      const a = r.screenshotAfter ? path.basename(r.screenshotAfter) : '';
      failDetails += `<div style="margin:20px 0;padding:15px;background:#fff5f5;border-radius:8px;border-left:4px solid #C62828;">
        <h3 style="margin-top:0;color:#C62828;">${r.module} &gt; ${r.name}</h3>`;
      if (b) failDetails += `<div style="display:inline-block;margin-right:15px;"><p style="margin:5px 0;font-size:12px;color:#666;">执行前</p><a href="${b}" target="_blank"><img src="${b}" style="max-width:480px;border:1px solid #ddd;border-radius:4px;"></a></div>`;
      if (a) failDetails += `<div style="display:inline-block;"><p style="margin:5px 0;font-size:12px;color:#666;">执行后</p><a href="${a}" target="_blank"><img src="${a}" style="max-width:480px;border:1px solid #ddd;border-radius:4px;"></a></div>`;
      failDetails += `</div>`;
    }
  }

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
  td img { border-radius: 4px; border: 1px solid #ddd; }
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
  <table><tr><th>#</th><th>模块</th><th>场景</th><th>优先级</th><th>页面</th><th>结果</th><th>API</th><th>耗时</th><th>异常</th><th>截图</th></tr>${rows}</table>
  ${failDetails}
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

  const moduleStats = {};
  for (const r of results) {
    if (!moduleStats[r.module]) moduleStats[r.module] = { total: 0, pass: 0, fail: 0 };
    moduleStats[r.module].total++;
    if (r.result === '通过') moduleStats[r.module].pass++; else moduleStats[r.module].fail++;
  }

  let md = `# 平台巡检报告 — ${envName}\n\n`;
  md += `> 生成时间: ${localTS()} | 场景文件: ${scenarioFiles.join(', ')}\n\n`;
  md += `## 巡检概览\n\n`;
  md += `| 指标 | 数值 |\n|---|---|\n`;
  md += `| 总场景 | ${results.length} |\n| 通过 | ${passCount} |\n| 失败 | ${failCount} |\n| 通过率 | ${passRate}% |\n\n`;
  md += `## 模块汇总\n\n| 模块 | 总数 | 通过 | 失败 |\n|---|---|---|---|\n`;
  for (const [mod, s] of Object.entries(moduleStats)) {
    md += `| ${mod} | ${s.total} | ${s.pass} | ${s.fail} |\n`;
  }
  md += `\n## 巡检明细\n\n| # | 模块 | 场景 | 优先级 | 页面 | 结果 | 耗时 | 异常 |\n`;
  md += `|---|---|---|---|---|---|---|---|\n`;
  for (const r of results) {
    const icon = r.result === '通过' ? '✅' : '❌';
    const errors = [r.stepError, ...r.assertions.filter(a => !a.pass).map(a => `${a.label}: ${a.detail}`)].filter(Boolean).join('; ');
    md += `| ${r.id} | ${r.module} | ${r.name} | ${r.priority} | ${r.pageName} | ${icon}${r.result} | ${(r.elapsed/1000).toFixed(1)}s | ${errors || '-'} |\n`;
  }

  // 失败详情（含截图）
  const failedResults = results.filter(r => r.result === '失败');
  if (failedResults.length > 0) {
    md += `\n## 失败场景详情\n\n`;
    for (const r of failedResults) {
      md += `### ${r.module} > ${r.name}\n\n`;
      if (r.stepError) md += `- **步骤错误**: ${r.stepError}\n`;
      for (const a of r.assertions.filter(a => !a.pass)) {
        md += `- **${a.label}**: ${a.detail}\n`;
      }
      if (r.errors.length > 0) {
        md += `- **异常记录** (${r.errors.length}条):\n`;
        for (const e of r.errors.slice(0, 5)) {
          md += `  - [${e.type}] ${e.message || e.url || ''}\n`;
        }
      }
      const b = r.screenshotBefore ? path.basename(r.screenshotBefore) : '';
      const a = r.screenshotAfter ? path.basename(r.screenshotAfter) : '';
      if (b || a) md += `- **截图**:\n`;
      if (b) md += `  - 执行前: ![before](${b})\n`;
      if (a) md += `  - 执行后: ![after](${a})\n`;
      md += '\n';
    }
  }

  const mdPath = path.join(RESULT_BASE, `巡检报告-${localTS()}.md`);
  fs.writeFileSync(mdPath, md, 'utf-8');
  console.log(`📄 Markdown: ${mdPath}`);
  return mdPath;
}

// ==================== 脚本备份 ====================
function backupScript() {
  try {
    const archiveDir = path.join(__dirname, 'archive');
    if (!fs.existsSync(archiveDir)) fs.mkdirSync(archiveDir, { recursive: true });
    const src = path.join(__dirname, 'inspection.mjs');
    if (fs.existsSync(src)) {
      const ts = localTS().replace(/:/g, '-');
      fs.copyFileSync(src, path.join(archiveDir, `inspection_${ts}.mjs`));
    }
  } catch {}
}

// ==================== 主流程 ====================
async function main() {
  backupScript();

  console.log('╔══════════════════════════════════════╗');
  console.log('║   🔍 平台自动化巡检 — Playwright+AI  ║');
  console.log('╚══════════════════════════════════════╝');

  // 1. 环境选择
  console.log('\n📡 Available environments:');
  for (const [k, v] of Object.entries(ENVIRONMENTS)) console.log(`  ${k} — ${v.name}`);
  const input = await ask('选择环境 (default 123): ');
  const envKey = ENVIRONMENTS[input] ? input : '123';
  const env = ENVIRONMENTS[envKey];
  PARTIAL_ENV = env;
  PARTIAL_ENV_KEY = envKey;
  console.log(`✅ 环境: ${env.name} → ${env.admin.startUrl}`);

  // 2. 加载菜单结构（供 AI 导航参考）
  const menuFiles = listMenuTreeFiles();
  let menuTree = '';
  if (menuFiles.length === 0) {
    console.log('\n⚠️ 未找到菜单结构文件（data/platform-inspection-case/）');
  } else if (menuFiles.length === 1) {
    menuTree = loadMenuTree(menuFiles[0]);
    console.log(`\n📂 菜单结构: ${menuFiles[0]} (${menuTree.length} chars)`);
  } else {
    console.log('\n📂 请选择菜单结构文件:');
    menuFiles.forEach((f, i) => console.log(`  ${i + 1}. ${f}`));
    const pick = parseInt(await ask('Select (default 1, 0=跳过): ')) - 1;
    if (pick >= 0 && pick < menuFiles.length) {
      menuTree = loadMenuTree(menuFiles[pick]);
      console.log(`✅ ${menuFiles[pick]} (${menuTree.length} chars)`);
    }
  }

  // 3. 选择场景文件
  const scenarioFiles = listScenarioFiles();
  if (scenarioFiles.length === 0) {
    console.error('❌ scenarios/ 目录下未找到场景文件 (.yml)');
    process.exit(1);
  }
  let selectedFiles = [];
  if (scenarioFiles.length === 1) {
    selectedFiles = [scenarioFiles[0]];
    console.log(`\n📂 场景文件: ${scenarioFiles[0]}`);
  } else {
    console.log('\n📂 可用场景文件:');
    scenarioFiles.forEach((f, i) => console.log(`  ${i + 1}. ${f}`));
    console.log('  a=all  1,3=list  1=单个');
    const pick = await ask('选择: ');
    const normalized = pick.replace(/，/g, ',').trim();
    if (normalized === 'a' || normalized === 'all') {
      selectedFiles = scenarioFiles;
    } else if (normalized.includes(',')) {
      const idxs = normalized.split(',').map(s => parseInt(s.trim())).filter(n => !isNaN(n));
      selectedFiles = idxs.map(i => scenarioFiles[i - 1]).filter(Boolean);
    } else {
      const idx = parseInt(normalized) - 1;
      selectedFiles = [scenarioFiles[idx >= 0 && idx < scenarioFiles.length ? idx : 0]];
    }
  }
  PARTIAL_FILES = selectedFiles;
  console.log(`✅ 已选: ${selectedFiles.join(', ')}`);

  // 3.5 自动检测 YAML 格式
  const scenarioDir = path.join(__dirname, 'scenarios');
  const formats = selectedFiles.map(f => detectYamlFormat(path.join(scenarioDir, f)));
  const hasTaskFormat = formats.includes('task');
  const hasScenarioFormat = formats.includes('scenario');

  if (hasTaskFormat && hasScenarioFormat) {
    console.log('\n⚠️ 不支持混用两种 YAML 格式，请选择同一格式的文件');
    process.exit(1);
  }

  const yamlFormat = hasTaskFormat ? 'task' : 'scenario';
  console.log(`\n📝 YAML 格式: ${yamlFormat === 'task' ? 'Task (web+tasks+flow)' : 'Scenario (scenarios+steps+assertions)'}`);

  // ==================== Task YAML 执行路径 ====================
  if (yamlFormat === 'task') {
    // 加载 Task YAML
    let allTasks = [];
    let taskWebConfig = {};
    for (const sf of selectedFiles) {
      const loaded = loadTaskYaml(path.join(scenarioDir, sf));
      if (!loaded) continue;
      taskWebConfig = { ...taskWebConfig, ...loaded.web };
      allTasks = allTasks.concat(loaded.tasks);
    }
    if (allTasks.length === 0) { console.log('没有可用的 YAML 任务'); process.exit(0); }

    console.log(`\n共 ${allTasks.length} 个 YAML 任务`);
    allTasks.forEach((t, i) => console.log(`  ${i + 1}. ${t.name} (${t.flow.length}步)`));
    console.log('\n选择方式: a=all / 1,3=list / 1-5=range / Enter=全部');
    const taskSel = await ask('选择任务: ');
    let selectedTasks = allTasks;
    const ts = taskSel.replace(/，/g, ',').trim().toLowerCase();
    if (ts === 'a' || ts === 'all') { /* 全部 */ }
    else if (/^\d+-\d+$/.test(ts)) { const [s, e] = ts.split('-').map(Number); selectedTasks = allTasks.filter((_, i) => (i+1) >= s && (i+1) <= e); }
    else if (ts.includes(',')) { const ids = ts.split(',').map(s => parseInt(s.trim())).filter(n => !isNaN(n)); selectedTasks = ids.map(id => allTasks[id - 1]).filter(Boolean); }
    else if (/^\d+$/.test(ts)) { const n = parseInt(ts); selectedTasks = allTasks.filter((_, i) => (i+1) === n); }
    if (selectedTasks.length === 0) { console.log('未选中任何任务'); process.exit(0); }
    console.log(`✅ 将执行 ${selectedTasks.length} 个任务`);

    // 创建结果目录
    RESULT_BASE = path.join(__dirname, 'results', `YAML任务-${envKey}-${localTS()}`);
    if (!fs.existsSync(RESULT_BASE)) fs.mkdirSync(RESULT_BASE, { recursive: true });

    // 启动浏览器
    console.log('\n🌐 启动浏览器...');
    const USER_DATA_DIR = path.join(__dirname, '.cache', 'chrome-profile-inspection');
    const context = await chromium.launchPersistentContext(USER_DATA_DIR, {
      headless: false,
      args: ['--ignore-certificate-errors', '--start-maximized'],
      ignoreHTTPSErrors: true,
    });
    const keepAlive = setInterval(() => { context.pages()[0]?.keyboard?.press('Shift').catch(() => {}); }, 4 * 60 * 1000);

    // Cookie 恢复
    const COOKIE_FILE = path.join(__dirname, '.cache', `cookies-inspection-${envKey}.json`);
    if (fs.existsSync(COOKIE_FILE)) {
      await context.addCookies(JSON.parse(fs.readFileSync(COOKIE_FILE, 'utf-8')));
      console.log('🍪 Cookie 已恢复');
    }
    const page = context.pages()[0] || await context.newPage();

    // 导航：使用 YAML 中的 URL 或环境默认 URL
    const taskUrl = taskWebConfig.url || env.admin.startUrl;
    const resolvedUrl = resolveEnvVar(taskUrl);
    console.log(`\n🔗 导航到: ${resolvedUrl}`);
    await page.goto(resolvedUrl, { waitUntil: 'domcontentloaded', timeout: 30000, ignoreHTTPSErrors: true }).catch(() => {});
    await wait(3000);

    // 登录检测
    const currentUrl = page.url();
    if (currentUrl.includes('login') || currentUrl.includes('white')) {
      console.log('\n⏳ 请在浏览器中手动登录...');
      try {
        await page.waitForURL(u => !u.href.includes('login') && !u.href.includes('white'), { timeout: 600000 });
        console.log('✅ 登录成功');
        fs.writeFileSync(COOKIE_FILE, JSON.stringify(await context.cookies(), null, 2));
      } catch {
        console.log('❌ 登录超时');
        clearInterval(keepAlive);
        await context.close().catch(() => {});
        process.exit(1);
      }
    } else {
      console.log('✅ 已登录');
    }

    // 创建 AI Agent
    const agentOpts = {
      ...(menuTree ? { aiActionContext: `当前平台菜单结构:\n${menuTree}` } : {}),
      generateReport: true,
      reportFileName: `midscene-report-${new Date().toISOString().replace(/[:.]/g, '-')}`,
      groupName: `YAML任务执行 — ${env.name}`,
      groupDescription: `环境: ${envKey} | 文件: ${selectedFiles.join(', ')}`,
    };
    const agent = new PlaywrightAgent(page, agentOpts);

    // 串行执行任务
    const mergedConfig = { ...DEFAULT_CONFIG };
    console.log(`\n🚀 开始执行 (${selectedTasks.length} 个任务, 串行模式)\n`);
    const results = [];
    const startTime = Date.now();

    for (let i = 0; i < selectedTasks.length; i++) {
      const task = selectedTasks[i];
      try {
        const result = await executeTaskScenario(page, task, agent, mergedConfig, RESULT_BASE, i + 1, selectedTasks.length);
        results.push(result);
        PARTIAL_RESULTS = [...results];
      } catch (e) {
        console.log(`     ❌ 任务执行异常: ${e.message?.substring(0, 80)}`);
        results.push({
          id: i + 1, name: task.name, module: 'YAML任务', priority: 'P1',
          tags: [], pageName: 'error', result: '失败',
          stepError: e.message?.substring(0, 200), assertions: [], errors: [],
          totalApiCalls: 0, screenshotBefore: '', screenshotAfter: '',
          elapsed: 0, timestamp: new Date().toISOString(),
        });
      }
    }

    const totalTime = Date.now() - startTime;
    const passCount = results.filter(r => r.result === '通过').length;
    console.log(`\n${'='.repeat(50)}`);
    console.log(`✅ 执行完成! ${passCount}/${results.length} 通过, 总耗时 ${(totalTime/1000/60).toFixed(1)} 分钟`);

    // 生成报告
    console.log('\n📊 正在生成报告...');
    const excelPath = await generateResultExcel(results, envKey);
    await generateWordReport(results, env.name, selectedFiles, excelPath);
    await generateHtmlReport(results, env.name, selectedFiles);
    await generateMarkdownReport(results, env.name, selectedFiles);
    copyMidsceneReport();

    console.log(`\n📁 结果目录: ${RESULT_BASE}`);
    console.log('\n按 Enter 关闭浏览器...');
    await ask('');
    clearInterval(keepAlive);
    try { await context.close(); } catch {}
    return; // Task 格式执行完毕
  }

  // ==================== Scenario YAML 执行路径（原有逻辑） ====================
  let allScenarios = [];
  let mergedConfig = { ...DEFAULT_CONFIG };
  let prerequisite = null;
  for (const sf of selectedFiles) {
    const loaded = loadScenario(path.join(__dirname, 'scenarios', sf));
    if (!loaded) continue;
    mergedConfig = { ...mergedConfig, ...loaded.config };
    if (loaded.prerequisite && !prerequisite) prerequisite = loaded.prerequisite;
    allScenarios = allScenarios.concat(loaded.scenarios);
  }
  if (allScenarios.length === 0) { console.log('没有可用的巡检场景'); process.exit(0); }

  // 5. 模块统计 & 场景选择
  const moduleStats = {};
  const priorityCount = {};
  for (const s of allScenarios) {
    moduleStats[s.module] = (moduleStats[s.module] || 0) + 1;
    priorityCount[s.priority] = (priorityCount[s.priority] || 0) + 1;
  }
  console.log(`\n共 ${allScenarios.length} 个巡检场景`);
  console.log('\n📊 模块分布:');
  for (const [mod, cnt] of Object.entries(moduleStats).sort((a, b) => b[1] - a[1])) {
    console.log(`  ${mod}: ${cnt}个`);
  }
  console.log('\n优先级: ' + Object.entries(priorityCount).map(([k, v]) => `${k}(${v})`).join(', '));
  console.log('\n选择方式:');
  console.log('  a/all — 全部');
  console.log('  P0/P1/P2/P3 — 按优先级');
  console.log('  M 模块名 — 按模块（如: M 系统管理）');
  console.log('  1-5 — 按序号范围');
  console.log('  1,3,5 — 按序号列表');
  const sel = await ask('选择场景: ');
  let selected = allScenarios;
  const ns = sel.replace(/，/g, ',').trim();
  const nsLower = ns.toLowerCase();
  if (nsLower === 'a' || nsLower === 'all') { /* 全部 */ }
  else if (/^p[0-3]$/i.test(nsLower)) { selected = allScenarios.filter(s => s.priority.toUpperCase() === nsLower.toUpperCase()); }
  else if (/^m\s+(.+)/i.test(nsLower)) { const modName = ns.replace(/^m\s+/i, '').trim(); selected = allScenarios.filter(s => s.module.includes(modName)); }
  else if (/^\d+-\d+$/.test(nsLower)) { const [s, e] = nsLower.split('-').map(Number); selected = allScenarios.filter(s => s.id >= s && s.id <= e); }
  else if (nsLower.includes(',')) { const ids = nsLower.split(',').map(s => parseInt(s.trim())).filter(n => !isNaN(n)); selected = allScenarios.filter(s => ids.includes(s.id)); }
  else if (/^\d+$/.test(nsLower)) { const n = parseInt(nsLower); selected = allScenarios.filter(s => s.id === n); }
  if (selected.length === 0) { console.log('未选中任何场景'); process.exit(0); }
  console.log(`✅ 将执行 ${selected.length} 个场景`);

  // 6. 创建结果目录
  RESULT_BASE = path.join(__dirname, 'results', `巡检-${envKey}-${localTS()}`);
  if (!fs.existsSync(RESULT_BASE)) fs.mkdirSync(RESULT_BASE, { recursive: true });

  // 7. 启动浏览器
  console.log('\n🌐 启动浏览器...');
  const USER_DATA_DIR = path.join(__dirname, '.cache', 'chrome-profile-inspection');
  const context = await chromium.launchPersistentContext(USER_DATA_DIR, {
    headless: false,
    args: ['--ignore-certificate-errors', '--start-maximized'],
    ignoreHTTPSErrors: true,
  });

  // 保活
  const keepAlive = setInterval(() => { context.pages()[0]?.keyboard?.press('Shift').catch(() => {}); }, 4 * 60 * 1000);

  // Cookie 恢复
  const COOKIE_FILE = path.join(__dirname, '.cache', `cookies-inspection-${envKey}.json`);
  if (fs.existsSync(COOKIE_FILE)) {
    await context.addCookies(JSON.parse(fs.readFileSync(COOKIE_FILE, 'utf-8')));
    console.log('🍪 Cookie 已恢复');
  }

  const page = context.pages()[0] || await context.newPage();

  // 8. 导航 & 登录
  console.log(`\n🔗 导航到: ${env.admin.startUrl}`);
  await page.goto(env.admin.startUrl, { waitUntil: 'domcontentloaded', timeout: 30000, ignoreHTTPSErrors: true }).catch(() => {});
  await wait(3000);

  const currentUrl = page.url();
  if (currentUrl.includes('login') || currentUrl.includes('white')) {
    console.log('\n⏳ 请在浏览器中手动登录...');
    console.log('   （等待登录完成，超时10分钟）');
    try {
      await page.waitForURL(u => !u.href.includes('login') && !u.href.includes('white'), { timeout: 600000 });
      console.log('✅ 登录成功');
      // 保存 Cookie
      fs.writeFileSync(COOKIE_FILE, JSON.stringify(await context.cookies(), null, 2));
    } catch {
      console.log('❌ 登录超时');
      clearInterval(keepAlive);
      await context.close().catch(() => {});
      process.exit(1);
    }
  } else {
    console.log('✅ 已登录');
  }

  // 9. 创建 AI Agent（注入菜单树上下文 + 启用 Midscene 报告）
  const agentOpts = {
    ...(menuTree ? { aiActionContext: `当前平台菜单结构:\n${menuTree}` } : {}),
    generateReport: true,
    reportFileName: `midscene-report-${new Date().toISOString().replace(/[:.]/g, '-')}`,
    groupName: `平台巡检 — ${env.name}`,
    groupDescription: `环境: ${envKey} | 场景文件: ${selectedFiles.join(', ')}`,
  };
  const agent = new PlaywrightAgent(page, agentOpts);

  // 10. 串行执行巡检
  console.log(`\n🚀 开始巡检 (${selected.length} 个场景, 串行模式)\n`);
  const results = [];
  const startTime = Date.now();

  for (let i = 0; i < selected.length; i++) {
    const scenario = selected[i];
    try {
      const result = await executeScenario(page, scenario, agent, mergedConfig, RESULT_BASE, i + 1, selected.length);
      results.push(result);
      PARTIAL_RESULTS = [...results]; // 用于中断时生成报告
    } catch (e) {
      console.log(`     ❌ 场景执行异常: ${e.message?.substring(0, 80)}`);
      results.push({
        id: i + 1, name: scenario.name, module: scenario.module, priority: scenario.priority,
        tags: scenario.tags, pageName: 'error', result: '失败',
        stepError: e.message?.substring(0, 200), assertions: [], errors: [],
        totalApiCalls: 0, screenshotBefore: '', screenshotAfter: '',
        elapsed: 0, timestamp: new Date().toISOString(),
      });
    }
  }

  const totalTime = Date.now() - startTime;
  const passCount = results.filter(r => r.result === '通过').length;
  console.log(`\n${'='.repeat(50)}`);
  console.log(`✅ 巡检完成! ${passCount}/${results.length} 通过, 总耗时 ${(totalTime/1000/60).toFixed(1)} 分钟`);

  // 11. 生成报告
  console.log('\n📊 正在生成报告...');
  const excelPath = await generateResultExcel(results, envKey);
  await generateWordReport(results, env.name, selectedFiles, excelPath);
  await generateHtmlReport(results, env.name, selectedFiles);
  await generateMarkdownReport(results, env.name, selectedFiles);

  // 12. 复制 Midscene 原生报告到结果目录
  copyMidsceneReport();

  console.log(`\n📁 结果目录: ${RESULT_BASE}`);
  console.log('\n按 Enter 关闭浏览器...');
  await ask('');
  clearInterval(keepAlive);
  try { await context.close(); } catch {}
}

main().catch(async (err) => {
  console.error('FATAL: ' + err.message);
  console.error(err.stack);
  await ask('').catch(() => {});
  process.exit(1);
});
