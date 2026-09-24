/**
 * inspection-task.mjs
 * Task YAML 用例执行器 — Playwright + Midscene AI
 *
 * 功能:
 *   - 执行 web: + tasks[].flow[] 格式的 YAML 用例
 *   - 支持: aiTap / aiInput / aiAction / aiAssert / aiWaitFor / sleep
 *   - 每任务截图 + 错误监控（API/Console/PageError）
 *   - 多格式报告输出（Excel/HTML/Markdown）
 *
 * 用例文件: scenarios/*.yaml (含 web: + tasks: 字段)
 * 环境配置: data/0 shared/environments.json
 *
 * 用法: node inspection-task.mjs
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
const { Document, Packer, Paragraph, TextRun, HeadingLevel, ImageRun } = require('docx');
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
function sanitizeFileName(name) {
  return (name || 'unknown').replace(/[\\/:*?"<>|]/g, '_').replace(/\s+/g, '-').substring(0, 40);
}
function resolveEnvVar(text) {
  if (typeof text !== 'string') return String(text ?? '');
  // 内置变量
  const now = new Date();
  const builtins = {
    DATE: `${now.getFullYear()}${String(now.getMonth()+1).padStart(2,'0')}${String(now.getDate()).padStart(2,'0')}`,
    TIMESTAMP: now.toISOString().replace(/[:.]/g, '-'),
    TIME: `${String(now.getHours()).padStart(2,'0')}${String(now.getMinutes()).padStart(2,'0')}`,
  };
  return text.replace(/\$\{([^}]+)\}/g, (_, key) => builtins[key] || process.env[key] || '');
}

// ==================== 配置 ====================
let RESULT_BASE = path.join(__dirname, 'results', localTS());
const DEFAULT_CONFIG = { page_load_wait: 3000, page_load_timeout: 30000, network_idle_wait: 5000 };

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
    envs[k] = { name: v.name || (k + '环境'), admin: { startUrl: v.admin?.url || '' } };
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
  const fp = path.join(__dirname, 'data', 'platform-inspection-case', selectedFile);
  return fs.existsSync(fp) ? fs.readFileSync(fp, 'utf-8').trim() : '';
}

// ==================== Task YAML 加载 ====================
function listTaskYamlFiles() {
  const dir = path.join(__dirname, 'scenarios');
  if (!fs.existsSync(dir)) return [];
  return fs.readdirSync(dir)
    .filter(f => (f.endsWith('.yml') || f.endsWith('.yaml')) && !f.startsWith('~$'))
    .filter(f => {
      try {
        const raw = fs.readFileSync(path.join(dir, f), 'utf-8');
        const content = raw.replace(/^\s*#[^\n]*/gm, '');
        return /^tasks:/m.test(content) && /^web:/m.test(content);
      } catch { return false; }
    });
}

function loadTaskYaml(filePath) {
  const raw = fs.readFileSync(filePath, 'utf-8');
  const doc = YAML.load(raw);
  if (!doc || !doc.tasks) return null;
  const webConfig = doc.web || {};
  const tasks = (doc.tasks || [])
    .filter(t => t && t.name && t.flow && t.flow.length > 0)
    .map((t, i) => ({ id: i + 1, name: t.name || `任务${i + 1}`, flow: t.flow || [] }));
  return {
    web: { url: webConfig.url || '', viewportWidth: webConfig.viewportWidth || 1401, viewportHeight: webConfig.viewportHeight || 694, acceptInsecureCerts: webConfig.acceptInsecureCerts !== false },
    tasks, file: path.basename(filePath),
  };
}

// ==================== 错误监控 ====================
function startErrorMonitor(page) {
  const errors = [];
  let totalApiCalls = 0;
  const onResponse = (response) => { totalApiCalls++; if (response.status() >= 400) errors.push({ type: 'API', status: response.status(), url: response.url().substring(0, 200), time: Date.now() }); };
  const onPageError = (err) => { errors.push({ type: 'PAGE', message: err.message?.substring(0, 200), time: Date.now() }); };
  const onConsole = (msg) => { if (msg.type() === 'error') { const t = msg.text(); if (/WebSocket|websocket|wss?:\/\//i.test(t)) return; errors.push({ type: 'CONSOLE', message: t.substring(0, 200), time: Date.now() }); } };
  page.on('response', onResponse); page.on('pageerror', onPageError); page.on('console', onConsole);
  return {
    errors, totalApiCalls: () => totalApiCalls,
    stop: () => { page.off('response', onResponse); page.off('pageerror', onPageError); page.off('console', onConsole); },
    snapshot: () => { const e = errors.splice(0); const n = totalApiCalls; totalApiCalls = 0; return { errors: e, totalApiCalls: n }; },
  };
}

// ==================== 页面名称检测 ====================
async function detectPageName(page) {
  try {
    await page.waitForTimeout(600);
    return await page.evaluate(() => {
      const ok = (t) => t && t.length >= 2 && t.length <= 80 && !/^[.•…·\s]+$/.test(t) && !/^https?:\/\//i.test(t) && !/^\d+$/.test(t);
      const sel = ['.el-menu-item.is-active', '.ant-menu-item-selected', 'li[class*="active"][class*="menu"]'];
      let el = null; for (const s of sel) { el = document.querySelector(s); if (el) break; }
      if (el) { const parts = []; const t = el.textContent.replace(/\s+/g, ' ').trim(); if (ok(t)) parts.unshift(t); let p = el.parentElement; const seen = new Set(); while (p && p !== document.body) { const ti = p.querySelector(':scope > .el-submenu__title, :scope > .ant-menu-submenu-title'); if (ti) { const tt = ti.textContent.replace(/\s+/g, ' ').trim(); if (ok(tt) && !seen.has(tt)) { seen.add(tt); parts.unshift(tt); } } p = p.parentElement; } if (parts.length) return parts.join(' > '); }
      const bc = document.querySelector('.el-breadcrumb, .ant-breadcrumb');
      if (bc) { const items = Array.from(bc.querySelectorAll('li, span')).map(e => e.textContent.replace(/\s+/g, ' ').trim()).filter(Boolean); for (let i = items.length - 1; i >= 0; i--) { if (ok(items[i])) return items[i]; } }
      if (document.title) { const parts = document.title.split(/\s*[-|—]\s*/); for (const p of parts) { const t = p.replace(/\s+/g, ' ').trim(); if (ok(t)) return t; } }
      return 'unknown-page';
    });
  } catch { return 'unknown-page'; }
}

// ==================== Flow 步骤执行 ====================
// 返回该步骤执行后是否需要等待页面稳定（aiAssert 前不等待，以捕获 Toast 等瞬态提示）
async function executeFlowStep(page, agent, step, stepIdx) {
  if (!step || typeof step !== 'object') return;

  if ('aiTap' in step) {
    const target = resolveEnvVar(String(step.aiTap));
    console.log(`     步骤${stepIdx}: 点击 [${target}]`);
    await agent.aiAct(`点击"${target}"`);

  } else if ('aiInput' in step) {
    const target = resolveEnvVar(String(step.aiInput));
    const value = resolveEnvVar(String(step.value || ''));
    console.log(`     步骤${stepIdx}: 在 [${target}] 输入 [${value}]`);
    await agent.aiAct(`在"${target}"中输入"${value}"`);

  } else if ('aiAction' in step) {
    let desc = step.aiAction;
    if (typeof desc === 'object' && desc !== null) { const vals = Object.values(desc); desc = vals.length > 0 ? String(vals[0]) : ''; }
    desc = resolveEnvVar(String(desc));
    console.log(`     步骤${stepIdx}: 执行 [${desc.length > 60 ? desc.substring(0, 60) + '...' : desc}]`);
    await agent.aiAct(desc);

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

// ==================== 任务执行 ====================
async function executeTask(page, task, agent, config, resultDir, idx, total) {
  console.log(`\n[${idx}/${total}] 📋 ${task.name}`);
  const startTime = Date.now();
  const shotPrefix = path.join(resultDir, `${String(idx).padStart(2,'0')}-${sanitizeFileName(task.name)}`);

  await page.screenshot({ path: shotPrefix + '-before.png', fullPage: false }).catch(() => {});
  const monitor = startErrorMonitor(page);

  let stepOk = true, stepError = '';
  for (let si = 0; si < task.flow.length; si++) {
    const step = task.flow[si]; if (!step) continue;
    const nextStep = task.flow[si + 1];
    const nextIsAssert = nextStep && typeof nextStep === 'object' && 'aiAssert' in nextStep;
    try {
      await executeFlowStep(page, agent, step, si + 1);
      // 智能等待：点击/导航后等待页面稳定，但如果下一步是 aiAssert 则不等待（以捕获 Toast 等瞬态提示）
      if ((step.aiTap || step.aiAction) && !nextIsAssert) {
        await page.waitForLoadState('networkidle', { timeout: config.network_idle_wait }).catch(() => {});
        await wait(1000);
      } else if (step.aiInput && !nextIsAssert) {
        await wait(500);
      }
      // aiAssert 前不加任何等待，紧接上一步立即执行
    } catch (e) {
      stepOk = false; stepError = `步骤${si + 1}失败: ${e.message?.substring(0, 120)}`;
      console.log(`     ❌ ${stepError}`); break;
    }
  }

  if (stepOk) { await wait(1500); monitor.snapshot(); }
  const snap = monitor.snapshot(); monitor.stop();
  await wait(500);
  await page.screenshot({ path: shotPrefix + '-after.png', fullPage: false }).catch(() => {});

  const pageName = await detectPageName(page);
  const elapsed = Date.now() - startTime;
  const result = {
    id: idx, name: task.name, module: 'YAML任务', priority: 'P1', tags: [],
    pageName, result: stepOk ? '通过' : '失败', stepError,
    assertions: [], errors: snap.errors, totalApiCalls: snap.totalApiCalls,
    screenshotBefore: shotPrefix + '-before.png', screenshotAfter: shotPrefix + '-after.png',
    elapsed, timestamp: new Date().toISOString(),
  };
  console.log(`     ${stepOk ? '✅' : '❌'} ${result.result} (${(elapsed / 1000).toFixed(1)}s)`);
  return result;
}

// ==================== 中断处理 ====================
let PARTIAL_RESULTS = [], PARTIAL_ENV = null, PARTIAL_ENV_KEY = '', PARTIAL_FILES = [], PARTIAL_GENERATING = false;
async function handleInterrupt() {
  if (PARTIAL_GENERATING || PARTIAL_RESULTS.length === 0) process.exit(0);
  PARTIAL_GENERATING = true;
  console.log('\n⚠️ 脚本被中断，正在生成部分报告...');
  try { await generateExcel(PARTIAL_RESULTS, PARTIAL_ENV_KEY); await generateWordReport(PARTIAL_RESULTS, PARTIAL_ENV?.name || '', PARTIAL_FILES); await generateHtmlReport(PARTIAL_RESULTS, PARTIAL_ENV?.name || '', PARTIAL_FILES); await generateMarkdownReport(PARTIAL_RESULTS, PARTIAL_ENV?.name || '', PARTIAL_FILES); } catch (e) { console.error('报告生成失败:', e.message); }
  process.exit(0);
}
process.on('SIGINT', () => handleInterrupt());
process.on('SIGTERM', () => handleInterrupt());

// ==================== 报告: Excel ====================
async function generateExcel(results, envKey) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('执行结果');
  [8, 20, 20, 10, 20, 10, 10, 12, 40].forEach((w, i) => ws.getColumn(i + 1).width = w);
  const hr = ws.addRow(['序号', '模块', '任务名称', '优先级', '页面名称', '结果', 'API调用', '耗时(s)', '异常详情']);
  hr.font = { bold: true, color: { argb: 'FFFFFFFF' } }; hr.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } }; hr.alignment = { horizontal: 'center', vertical: 'middle' };
  for (const r of results) {
    const err = [r.stepError, ...r.errors.slice(0, 3).map(e => `[${e.type}] ${e.message || e.url || ''}`)].filter(Boolean).join('; ');
    const row = ws.addRow([r.id, r.module, r.name, r.priority, r.pageName, r.result, r.totalApiCalls, (r.elapsed / 1000).toFixed(1), err]);
    if (r.result === '通过') row.getCell(6).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2EFDA' } };
    else { row.getCell(6).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFCE4EC' } }; row.getCell(6).font = { color: { argb: 'FFC00000' }, bold: true }; }
  }
  const sws = wb.addWorksheet('汇总');
  const pc = results.filter(r => r.result === '通过').length;
  sws.addRow(['YAML任务执行汇总', '', '']); sws.addRow(['环境', envKey]); sws.addRow(['时间', localTS()]);
  sws.addRow(['总任务数', results.length]); sws.addRow(['通过', pc]); sws.addRow(['失败', results.length - pc]);
  sws.addRow(['通过率', results.length > 0 ? ((pc / results.length) * 100).toFixed(1) + '%' : 'N/A']);
  const p = path.join(RESULT_BASE, `执行结果-${localTS()}.xlsx`);
  await wb.xlsx.writeFile(p); console.log(`📊 Excel: ${p}`); return p;
}

// ==================== 报告: Word ====================
async function generateWordReport(results, envName, files) {
  const pc = results.filter(r => r.result === '通过').length;
  const fc = results.length - pc;
  const pr = results.length > 0 ? ((pc / results.length) * 100).toFixed(1) : '0';
  const children = [];
  // 标题
  children.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(`YAML任务执行报告 — ${envName}`)] }));
  children.push(new Paragraph({ children: [new TextRun({ text: `生成时间: ${localTS()}`, color: '666666', size: 20 })] }));
  children.push(new Paragraph({ children: [new TextRun({ text: `用例文件: ${files.join(', ')}`, color: '666666', size: 20 })] }));
  // 概览
  children.push(new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun('执行概览')] }));
  children.push(new Paragraph({ children: [new TextRun(`总任务: ${results.length}  |  通过: ${pc}  |  失败: ${fc}  |  通过率: ${pr}%`)] }));
  // 失败详情
  const failed = results.filter(r => r.result === '失败');
  if (failed.length > 0) {
    children.push(new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun('失败任务详情')] }));
    for (const r of failed) {
      children.push(new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun(`#${r.id} ${r.name}`)] }));
      if (r.stepError) children.push(new Paragraph({ children: [new TextRun({ text: `步骤错误: ${r.stepError}`, color: 'CC0000' })] }));
      if (r.errors.length > 0) {
        children.push(new Paragraph({ children: [new TextRun(`异常记录 (${r.errors.length}条):`)] }));
        for (const e of r.errors.slice(0, 5)) {
          children.push(new Paragraph({ children: [new TextRun({ text: `  [${e.type}] ${e.message || e.url || ''}`, size: 18 })] }));
        }
      }
      // 截图
      const beforeImg = r.screenshotBefore && fs.existsSync(r.screenshotBefore) ? fs.readFileSync(r.screenshotBefore) : null;
      const afterImg = r.screenshotAfter && fs.existsSync(r.screenshotAfter) ? fs.readFileSync(r.screenshotAfter) : null;
      if (beforeImg || afterImg) {
        children.push(new Paragraph({ children: [new TextRun({ text: '页面截图:', bold: true, size: 20 })] }));
        const imgs = [];
        if (beforeImg) { imgs.push(new ImageRun({ data: beforeImg, transformation: { width: 450, height: 253 }, type: 'png' })); imgs.push(new TextRun({ text: '  执行前  ', size: 18, color: '666666' })); }
        if (afterImg) { imgs.push(new ImageRun({ data: afterImg, transformation: { width: 450, height: 253 }, type: 'png' })); imgs.push(new TextRun({ text: '  执行后', size: 18, color: '666666' })); }
        children.push(new Paragraph({ children: imgs }));
      }
    }
  }
  // 全部结果列表
  children.push(new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun('执行明细')] }));
  for (const r of results) {
    const icon = r.result === '通过' ? '✅' : '❌';
    children.push(new Paragraph({ children: [new TextRun(`${icon} #${r.id} ${r.name} — ${r.pageName} (${(r.elapsed/1000).toFixed(1)}s)`)] }));
  }
  const doc = new Document({ sections: [{ children }] });
  const buffer = await Packer.toBuffer(doc);
  const wordPath = path.join(RESULT_BASE, `执行报告-${localTS()}.docx`);
  fs.writeFileSync(wordPath, buffer);
  console.log(`📝 Word: ${wordPath}`);
  return wordPath;
}

// ==================== 报告: HTML ====================
async function generateHtmlReport(results, envName, files) {
  const pc = results.filter(r => r.result === '通过').length;
  const fc = results.length - pc;
  const pr = results.length > 0 ? ((pc / results.length) * 100).toFixed(1) : '0';
  const rows = results.map(r => {
    const cls = r.result === '通过' ? 'pass' : 'fail';
    const err = [r.stepError, ...r.errors.slice(0, 2).map(e => `[${e.type}] ${e.message || e.url || ''}`)].filter(Boolean).join('<br>');
    let shots = '-';
    if (r.result === '失败') { const b = r.screenshotBefore ? path.basename(r.screenshotBefore) : ''; const a = r.screenshotAfter ? path.basename(r.screenshotAfter) : ''; const t = []; if (b) t.push(`<a href="${b}" target="_blank"><img src="${b}" width="120" title="执行前"></a>`); if (a) t.push(`<a href="${a}" target="_blank"><img src="${a}" width="120" title="执行后"></a>`); shots = t.join(' '); }
    return `<tr class="${cls}"><td>${r.id}</td><td>${r.name}</td><td>${r.pageName}</td><td class="result-${cls}">${r.result}</td><td>${r.totalApiCalls}</td><td>${(r.elapsed/1000).toFixed(1)}s</td><td>${err||'-'}</td><td>${shots}</td></tr>`;
  }).join('\n');
  const html = `<!DOCTYPE html><html lang="zh-CN"><head><meta charset="UTF-8"><title>YAML任务执行报告</title>
<style>body{font-family:-apple-system,'Microsoft YaHei',sans-serif;margin:20px;background:#f5f5f5}.container{max-width:1400px;margin:0 auto;background:#fff;padding:30px;border-radius:8px;box-shadow:0 2px 8px rgba(0,0,0,.1)}h1{color:#333;border-bottom:2px solid #4472C4;padding-bottom:10px}.summary{display:flex;gap:20px;margin:20px 0}.summary-card{flex:1;padding:15px;border-radius:8px;text-align:center}.card-total{background:#E8EEF7}.card-pass{background:#E2EFDA}.card-fail{background:#FCE4EC}.card-rate{background:#FFF3E0}.summary-card .num{font-size:32px;font-weight:bold}.summary-card .label{color:#666;margin-top:5px}table{width:100%;border-collapse:collapse;margin:15px 0;font-size:13px}th{background:#4472C4;color:#fff;padding:10px 8px;text-align:left}td{padding:8px;border-bottom:1px solid #eee}tr:hover{background:#f0f5ff}.result-pass{color:#2E7D32;font-weight:bold}.result-fail{color:#C62828;font-weight:bold}.meta{color:#888;font-size:13px}</style></head>
<body><div class="container"><h1>📋 YAML任务执行报告 — ${envName}</h1><p class="meta">生成时间: ${localTS()} | 用例文件: ${files.join(', ')}</p>
<div class="summary"><div class="summary-card card-total"><div class="num">${results.length}</div><div class="label">总任务</div></div><div class="summary-card card-pass"><div class="num">${pc}</div><div class="label">通过</div></div><div class="summary-card card-fail"><div class="num">${fc}</div><div class="label">失败</div></div><div class="summary-card card-rate"><div class="num">${pr}%</div><div class="label">通过率</div></div></div>
<h2>执行明细</h2><table><tr><th>#</th><th>任务</th><th>页面</th><th>结果</th><th>API</th><th>耗时</th><th>异常</th><th>截图</th></tr>${rows}</table></div></body></html>`;
  const p = path.join(RESULT_BASE, `执行报告-${localTS()}.html`);
  fs.writeFileSync(p, html, 'utf-8'); console.log(`🌐 HTML: ${p}`); return p;
}

// ==================== 报告: Markdown ====================
async function generateMarkdownReport(results, envName, files) {
  const pc = results.filter(r => r.result === '通过').length;
  const fc = results.length - pc;
  const pr = results.length > 0 ? ((pc / results.length) * 100).toFixed(1) : '0';
  let md = `# YAML任务执行报告 — ${envName}\n\n> 生成时间: ${localTS()} | 用例文件: ${files.join(',')}\n\n`;
  md += `## 概览\n\n| 指标 | 数值 |\n|---|---|\n| 总任务 | ${results.length} |\n| 通过 | ${pc} |\n| 失败 | ${fc} |\n| 通过率 | ${pr}% |\n\n`;
  md += `## 执行明细\n\n| # | 任务 | 页面 | 结果 | 耗时 | 异常 |\n|---|---|---|---|---|---|\n`;
  for (const r of results) { const icon = r.result === '通过' ? '✅' : '❌'; md += `| ${r.id} | ${r.name} | ${r.pageName} | ${icon}${r.result} | ${(r.elapsed/1000).toFixed(1)}s | ${r.stepError || '-'} |\n`; }
  const p = path.join(RESULT_BASE, `执行报告-${localTS()}.md`);
  fs.writeFileSync(p, md, 'utf-8'); console.log(`📄 Markdown: ${p}`); return p;
}

// ==================== Midscene 报告复制 ====================
function copyMidsceneReport() {
  try {
    const dir = path.join(__dirname, 'midscene_run', 'report');
    if (!fs.existsSync(dir)) return;
    const files = fs.readdirSync(dir).filter(f => f.startsWith('midscene-report-') && f.endsWith('.html')).sort().reverse();
    if (files.length > 0) { const dest = path.join(RESULT_BASE, `Midscene报告-${localTS()}.html`); fs.copyFileSync(path.join(dir, files[0]), dest); console.log(`🌐 Midscene报告: ${dest}`); }
  } catch (e) { console.log('⚠️ 复制Midscene报告失败:', e.message); }
}

// ==================== 主流程 ====================
async function main() {
  console.log('╔══════════════════════════════════════════╗');
  console.log('║   📋 YAML任务执行器 — Playwright+AI     ║');
  console.log('╚══════════════════════════════════════════╝');

  // 1. 环境选择
  console.log('\n📡 Available environments:');
  for (const [k, v] of Object.entries(ENVIRONMENTS)) console.log(`  ${k} — ${v.name}`);
  const input = await ask('选择环境 (default 122): ');
  const envKey = ENVIRONMENTS[input] ? input : '122';
  const env = ENVIRONMENTS[envKey];
  PARTIAL_ENV = env; PARTIAL_ENV_KEY = envKey;
  console.log(`✅ 环境: ${env.name} → ${env.admin.startUrl}`);

  // 2. 菜单结构
  const menuFiles = listMenuTreeFiles();
  let menuTree = '';
  if (menuFiles.length > 0) { menuTree = loadMenuTree(menuFiles[0]); console.log(`📂 菜单结构: ${menuTree.length} chars`); }

  // 3. 选择 YAML 文件
  const yamlFiles = listTaskYamlFiles();
  if (yamlFiles.length === 0) { console.error('❌ scenarios/ 目录下未找到 Task YAML 文件 (需含 web: + tasks: 字段)'); process.exit(1); }
  let selectedFiles = [];
  if (yamlFiles.length === 1) { selectedFiles = [yamlFiles[0]]; console.log(`\n📂 用例文件: ${yamlFiles[0]}`); }
  else {
    console.log('\n📂 可用用例文件:');
    yamlFiles.forEach((f, i) => console.log(`  ${i + 1}. ${f}`));
    const pick = await ask('选择 (a=all, 1,3=list): ');
    const n = pick.replace(/，/g, ',').trim();
    if (n === 'a' || n === 'all') selectedFiles = yamlFiles;
    else if (n.includes(',')) { const ids = n.split(',').map(s => parseInt(s.trim())).filter(x => !isNaN(x)); selectedFiles = ids.map(i => yamlFiles[i - 1]).filter(Boolean); }
    else { const idx = parseInt(n) - 1; selectedFiles = [yamlFiles[idx >= 0 && idx < yamlFiles.length ? idx : 0]]; }
  }
  PARTIAL_FILES = selectedFiles;
  console.log(`✅ 已选: ${selectedFiles.join(', ')}`);

  // 4. 加载任务
  let allTasks = [], taskWebConfig = {};
  for (const sf of selectedFiles) {
    const loaded = loadTaskYaml(path.join(__dirname, 'scenarios', sf));
    if (!loaded) continue;
    taskWebConfig = { ...taskWebConfig, ...loaded.web };
    allTasks = allTasks.concat(loaded.tasks);
  }
  if (allTasks.length === 0) { console.log('没有可用的 YAML 任务'); process.exit(0); }

  console.log(`\n共 ${allTasks.length} 个 YAML 任务`);
  allTasks.forEach((t, i) => console.log(`  ${i + 1}. ${t.name} (${t.flow.length}步)`));
  const taskSel = await ask('选择任务 (a=all, 1-3=range, Enter=全部): ');
  let selectedTasks = allTasks;
  const ts = taskSel.replace(/，/g, ',').trim().toLowerCase();
  if (ts === 'a' || ts === 'all') {}
  else if (/^\d+-\d+$/.test(ts)) { const [s, e] = ts.split('-').map(Number); selectedTasks = allTasks.filter((_, i) => (i+1) >= s && (i+1) <= e); }
  else if (ts.includes(',')) { const ids = ts.split(',').map(s => parseInt(s.trim())).filter(n => !isNaN(n)); selectedTasks = ids.map(id => allTasks[id - 1]).filter(Boolean); }
  else if (/^\d+$/.test(ts)) { selectedTasks = allTasks.filter((_, i) => (i+1) === parseInt(ts)); }
  if (selectedTasks.length === 0) { console.log('未选中任何任务'); process.exit(0); }
  console.log(`✅ 将执行 ${selectedTasks.length} 个任务`);

  // 5. 结果目录
  RESULT_BASE = path.join(__dirname, 'results', `YAML任务-${envKey}-${localTS()}`);
  fs.mkdirSync(RESULT_BASE, { recursive: true });

  // 6. 启动浏览器
  console.log('\n🌐 启动浏览器...');
  const USER_DATA_DIR = path.join(__dirname, '.cache', 'chrome-profile-inspection');
  const context = await chromium.launchPersistentContext(USER_DATA_DIR, { headless: false, args: ['--ignore-certificate-errors', '--start-maximized'], ignoreHTTPSErrors: true });
  const keepAlive = setInterval(() => { context.pages()[0]?.keyboard?.press('Shift').catch(() => {}); }, 4 * 60 * 1000);

  const COOKIE_FILE = path.join(__dirname, '.cache', `cookies-inspection-${envKey}.json`);
  if (fs.existsSync(COOKIE_FILE)) { await context.addCookies(JSON.parse(fs.readFileSync(COOKIE_FILE, 'utf-8'))); console.log('🍪 Cookie 已恢复'); }
  const page = context.pages()[0] || await context.newPage();

  // 7. 导航（环境 URL 优先，YAML web.url 仅在环境无 startUrl 时备选）
  const envStartUrl = env.admin.startUrl;
  const taskUrl = resolveEnvVar(envStartUrl || taskWebConfig.url || '');
  if (!taskUrl) { console.error('❌ 无可用的导航地址'); process.exit(1); }
  console.log(`\n🔗 导航到: ${taskUrl}`);
  if (envStartUrl && taskWebConfig.url && taskWebConfig.url !== envStartUrl) {
    console.log(`   ℹ️ YAML 中的 url (${taskWebConfig.url}) 已被环境地址覆盖`);
  }
  await page.goto(taskUrl, { waitUntil: 'domcontentloaded', timeout: 30000, ignoreHTTPSErrors: true }).catch(() => {});
  await wait(3000);

  // 8. 登录检测
  const currentUrl = page.url();
  if (currentUrl.includes('login') || currentUrl.includes('white')) {
    console.log('\n⏳ 请在浏览器中手动登录...');
    try { await page.waitForURL(u => !u.href.includes('login') && !u.href.includes('white'), { timeout: 600000 }); console.log('✅ 登录成功'); fs.writeFileSync(COOKIE_FILE, JSON.stringify(await context.cookies(), null, 2)); }
    catch { console.log('❌ 登录超时'); clearInterval(keepAlive); await context.close().catch(() => {}); process.exit(1); }
  } else { console.log('✅ 已登录'); }

  // 9. AI Agent
  const agent = new PlaywrightAgent(page, {
    ...(menuTree ? { aiActionContext: `当前平台菜单结构:\n${menuTree}` } : {}),
    generateReport: true,
    reportFileName: `midscene-report-${new Date().toISOString().replace(/[:.]/g, '-')}`,
    groupName: `YAML任务执行 — ${env.name}`,
    groupDescription: `环境: ${envKey} | 文件: ${selectedFiles.join(', ')}`,
  });

  // 10. 串行执行
  console.log(`\n🚀 开始执行 (${selectedTasks.length} 个任务)\n`);
  const results = [];
  const startTime = Date.now();
  for (let i = 0; i < selectedTasks.length; i++) {
    try { const r = await executeTask(page, selectedTasks[i], agent, DEFAULT_CONFIG, RESULT_BASE, i + 1, selectedTasks.length); results.push(r); PARTIAL_RESULTS = [...results]; }
    catch (e) { console.log(`     ❌ 异常: ${e.message?.substring(0, 80)}`); results.push({ id: i+1, name: selectedTasks[i].name, module: 'YAML任务', priority: 'P1', tags: [], pageName: 'error', result: '失败', stepError: e.message?.substring(0, 200), assertions: [], errors: [], totalApiCalls: 0, screenshotBefore: '', screenshotAfter: '', elapsed: 0, timestamp: new Date().toISOString() }); }
  }

  const totalTime = Date.now() - startTime;
  const passCount = results.filter(r => r.result === '通过').length;
  console.log(`\n${'='.repeat(50)}`);
  console.log(`✅ 执行完成! ${passCount}/${results.length} 通过, 总耗时 ${(totalTime/1000/60).toFixed(1)} 分钟`);

  // 11. 报告
  console.log('\n📊 正在生成报告...');
  await generateExcel(results, envKey);
  await generateWordReport(results, env.name, selectedFiles);
  await generateHtmlReport(results, env.name, selectedFiles);
  await generateMarkdownReport(results, env.name, selectedFiles);
  copyMidsceneReport();

  console.log(`\n📁 结果目录: ${RESULT_BASE}`);
  console.log('\n按 Enter 关闭浏览器...');
  await ask('');
  clearInterval(keepAlive);
  try { await context.close(); } catch {}
}

main().catch(async (err) => { console.error('FATAL: ' + err.message); console.error(err.stack); await ask('').catch(() => {}); process.exit(1); });
