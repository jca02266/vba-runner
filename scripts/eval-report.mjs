#!/usr/bin/env node

/**
 * Generate aggregate and time-series reports from structured evaluations.
 *
 * Usage:
 *   node scripts/eval-report.mjs
 *   node scripts/eval-report.mjs --output /tmp/eval-report.md --html /tmp/eval-report.html --csv /tmp/eval-curve.csv
 */

import fs from 'node:fs';
import path from 'node:path';
import process from 'node:process';
import * as yaml from 'js-yaml';

const root = process.cwd();
const evaluationRoot = path.join(root, 'evaluation');
const recordsDir = path.join(evaluationRoot, 'evaluations');
const statesDir = path.join(evaluationRoot, 'states');
const defaultOutput = path.join(evaluationRoot, 'EVAL_REPORT.md');

const settledStatuses = new Set(['fixed', 'verified-no-bug', 'retired']);

function usage() {
  console.log('Usage: node scripts/eval-report.mjs [--output FILE] [--csv FILE]');
}

function parseArgs(argv) {
  const options = { output: defaultOutput, html: null, csv: null };
  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (arg === '--help' || arg === '-h') {
      usage();
      process.exit(0);
    }
    if (arg === '--output' || arg === '-o') {
      options.output = path.resolve(argv[++i] ?? '');
      continue;
    }
    if (arg === '--csv') {
      options.csv = path.resolve(argv[++i] ?? '');
      continue;
    }
    if (arg === '--html') {
      options.html = path.resolve(argv[++i] ?? '');
      continue;
    }
    throw new Error(`unknown argument: ${arg}`);
  }
  return options;
}

function listFiles(dir, suffix) {
  if (!fs.existsSync(dir)) return [];
  return fs.readdirSync(dir).filter((name) => name.endsWith(suffix));
}

function readFrontmatter(file) {
  const source = fs.readFileSync(file, 'utf8');
  const match = source.match(/^---\r?\n([\s\S]*?)\r?\n---(?:\r?\n|$)/);
  if (!match) throw new Error(`${file}: missing YAML frontmatter`);
  const data = yaml.load(match[1], { json: true });
  if (!data || typeof data !== 'object' || Array.isArray(data)) {
    throw new Error(`${file}: frontmatter must be an object`);
  }
  return data;
}

function readRecords() {
  return listFiles(recordsDir, '.md').map((name) => readFrontmatter(path.join(recordsDir, name)));
}

function readResults() {
  const results = new Map();
  for (const name of listFiles(statesDir, '.result.yml')) {
    const result = yaml.load(fs.readFileSync(path.join(statesDir, name), 'utf8'), { json: true });
    if (result?.evaluationId) results.set(result.evaluationId, result);
  }
  return results;
}

function findingCount(record) {
  return new Set(Array.isArray(record.findings) ? record.findings : []).size;
}

// A record is assigned to its first matching primary area so totals remain additive.
const areas = [
  ['Select Case', /Select Case/],
  ['Format', /Format/],
  ['数値精度', /LongLong|Decimal|Currency|数値String|算術|整数除算|指数|桁区切り|Scientific/],
  ['引数・ByRef・Property', /ByRef|Property|引数|CallByName/],
  ['ファイルI\/O・FSO', /FSO|ファイル|Append|Input|EOF|VFS/],
  ['配列・UDT', /配列|UDT|Binary/],
  ['Option Compare・Like', /Option Compare|Like/],
  ['Date', /日付|Date/],
  ['Err・On Error', /Err|エラー|On Error/],
  ['LSP・拡張機能', /LSP|拡張機能/],
  ['Parser・Lexer', /Lexer|文法|識別子/],
  ['評価基盤', /EVAL_LOG|カバレッジ|評価|テスト|Namespace|名前空間|リファクタリング|ミューテーション/],
];

function areaFor(focus) {
  return areas.find(([, pattern]) => pattern.test(String(focus ?? '')))?.[0] ?? 'その他';
}

function aggregate(records) {
  const groups = new Map();
  for (const record of records) {
    const area = areaFor(record.focus);
    const group = groups.get(area) ?? { area, evaluations: 0, bugs: 0, unresolved: 0 };
    group.evaluations += 1;
    group.bugs += findingCount(record);
    if (!settledStatuses.has(record.status)) group.unresolved += 1;
    groups.set(area, group);
  }
  return [...groups.values()].sort((a, b) => b.evaluations - a.evaluations || a.area.localeCompare(b.area));
}

function timeSeries(records, results) {
  const events = records
    .map((record) => ({ record, result: results.get(record.id) }))
    .filter(({ result }) => result?.completedAt && !Number.isNaN(Date.parse(result.completedAt)))
    .sort((a, b) => Date.parse(a.result.completedAt) - Date.parse(b.result.completedAt));
  let evaluations = 0;
  let bugs = 0;
  let unresolved = 0;
  return events.map(({ record, result }) => {
    evaluations += 1;
    bugs += findingCount(record);
    if (!settledStatuses.has(result.status)) unresolved += 1;
    return {
      date: new Date(result.completedAt).toISOString(),
      evaluationId: record.id,
      status: result.status,
      area: areaFor(record.focus),
      evaluations,
      bugs,
      unresolved,
    };
  });
}

function mdCell(value) {
  return String(value ?? '').replaceAll('|', '\\|').replaceAll('\n', ' ');
}

function renderMarkdown(records, summary, series) {
  const totalBugs = records.reduce((sum, record) => sum + findingCount(record), 0);
  const completed = series.length;
  const pending = records.length - completed;
  const lines = [
    '# 評価レポート',
    '',
    'このファイルは `scripts/eval-report.mjs` から生成されます。',
    '',
    `評価件数: ${records.length}、発見バグ件数: ${totalBugs}、完了日時付き: ${completed}、日時未登録: ${pending}`,
    '',
    '## 実装領域別集計',
    '',
    '| 実装領域 | 評価件数 | バグ件数 | 未収束件数 |',
    '|---|---:|---:|---:|',
    ...summary.map((row) => `| ${mdCell(row.area)} | ${row.evaluations} | ${row.bugs} | ${row.unresolved} |`),
    '',
    '## 時系列の収束状況',
    '',
    '完了状態ファイルの `completedAt` を基準にした累積値です。日時未登録の評価は含みません。',
    '',
    '| 完了日時 | 評価ID | 状態 | 実装領域 | 累積評価 | 累積バグ | 未収束 |',
    '|---|---|---|---|---:|---:|---:|',
    ...series.map((row) => `| ${row.date} | ${row.evaluationId} | ${row.status} | ${mdCell(row.area)} | ${row.evaluations} | ${row.bugs} | ${row.unresolved} |`),
  ];
  return `${lines.join('\n')}\n`;
}

function csvCell(value) {
  const text = String(value ?? '');
  return /[",\n]/.test(text) ? `"${text.replaceAll('"', '""')}"` : text;
}

function htmlCell(value) {
  return String(value ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;');
}

function renderConvergenceChart(series) {
  const labels = JSON.stringify(series.map((row) => row.date.slice(0, 10)));
  const values = (field) => JSON.stringify(series.map((row) => row[field]));
  return `<div class="chart-container"><canvas id="convergence-chart" role="img" aria-label="評価・バグ・未収束の累積推移"></canvas></div>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.4/dist/chart.umd.min.js"></script>
<script>
const convergenceLabels = ${labels};
new Chart(document.getElementById('convergence-chart'), {
  type: 'line',
  data: { labels: convergenceLabels, datasets: [
    { label: '累積評価', data: ${values('evaluations')}, borderColor: '#2563eb', backgroundColor: '#2563eb', tension: 0.15 },
    { label: '累積バグ', data: ${values('bugs')}, borderColor: '#dc2626', backgroundColor: '#dc2626', tension: 0.15 },
    { label: '未収束', data: ${values('unresolved')}, borderColor: '#d97706', backgroundColor: '#d97706', tension: 0.15 }
  ]},
  options: { responsive: true, interaction: { mode: 'index', intersect: false },
    plugins: { title: { display: true, text: 'バグ収束曲線' }, tooltip: { mode: 'index' } },
    scales: { x: { title: { display: true, text: '評価完了日' } }, y: { beginAtZero: true, title: { display: true, text: '累積件数' }, ticks: { precision: 0 } } }
  }
});
</script>`;
}

function renderHtml(records, summary, series) {
  const totalBugs = records.reduce((sum, record) => sum + findingCount(record), 0);
  const summaryRows = summary.map((row) => `<tr><td>${htmlCell(row.area)}</td><td>${row.evaluations}</td><td>${row.bugs}</td><td>${row.unresolved}</td></tr>`).join('\n');
  const seriesRows = series.map((row) => `<tr><td>${htmlCell(row.date)}</td><td>${htmlCell(row.evaluationId)}</td><td>${htmlCell(row.status)}</td><td>${htmlCell(row.area)}</td><td>${row.evaluations}</td><td>${row.bugs}</td><td>${row.unresolved}</td></tr>`).join('\n');
  return `<!doctype html>
<html lang="ja"><head><meta charset="utf-8">
<title>VBA Runner 評価レポート</title>
<style>body{font-family:system-ui,sans-serif;line-height:1.5;margin:2rem}table{border-collapse:collapse;margin:1rem 0 2rem}th,td{border:1px solid #bbb;padding:.35rem .6rem;text-align:left}th{background:#eee}td:not(:first-child){text-align:right}code{background:#f3f3f3;padding:.1rem .25rem}.chart-container{max-width:1000px;margin:1rem 0 2rem}</style>
</head><body><h1>評価レポート</h1>
<p>評価件数: ${records.length}、発見バグ件数: ${totalBugs}、完了日時付き: ${series.length}、日時未登録: ${records.length - series.length}</p>
<h2>実装領域別集計</h2><table><thead><tr><th>実装領域</th><th>評価件数</th><th>バグ件数</th><th>未収束件数</th></tr></thead><tbody>${summaryRows}</tbody></table>
<h2>時系列の収束状況</h2><p><code>completedAt</code>を基準にした累積値です。日時未登録の評価は含みません。</p>
${renderConvergenceChart(series)}
<table><thead><tr><th>完了日時</th><th>評価ID</th><th>状態</th><th>実装領域</th><th>累積評価</th><th>累積バグ</th><th>未収束</th></tr></thead><tbody>${seriesRows}</tbody></table>
</body></html>\n`;
}

function renderCsv(series) {
  const header = ['completedAt', 'evaluationId', 'status', 'area', 'cumulativeEvaluations', 'cumulativeBugs', 'unresolved'];
  return `${header.join(',')}\n${series.map((row) => [row.date, row.evaluationId, row.status, row.area, row.evaluations, row.bugs, row.unresolved].map(csvCell).join(',')).join('\n')}\n`;
}

function main() {
  const options = parseArgs(process.argv.slice(2));
  const records = readRecords();
  const results = readResults();
  const summary = aggregate(records);
  const series = timeSeries(records, results);
  fs.mkdirSync(path.dirname(options.output), { recursive: true });
  fs.writeFileSync(options.output, renderMarkdown(records, summary, series));
  if (options.html) {
    fs.mkdirSync(path.dirname(options.html), { recursive: true });
    fs.writeFileSync(options.html, renderHtml(records, summary, series));
  }
  if (options.csv) {
    fs.mkdirSync(path.dirname(options.csv), { recursive: true });
    fs.writeFileSync(options.csv, renderCsv(series));
  }
  console.log(options.output);
  if (options.html) console.log(options.html);
  if (options.csv) console.log(options.csv);
}

try {
  main();
} catch (error) {
  console.error(`eval-report: ${error.message}`);
  process.exitCode = 1;
}
