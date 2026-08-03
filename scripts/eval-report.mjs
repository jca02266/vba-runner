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
// Findings are considered resolved only when the evaluation explicitly closes
// the defect or classifies it as a known limit or withdrawn hypothesis.
// Blocked evaluations remain unconfirmed and therefore unresolved.
const resolvedFindingStatuses = new Set(['fixed']);
// Known limits and withdrawn candidates are evaluation outcomes, not findings
// to include in discovery or convergence totals.
const excludedFindingStatuses = new Set(['known-limit', 'retired']);
const pendingEvaluationStatuses = new Set(['needs-excel-probe', 'needs-excel', 'blocked', 'in-progress', 'claimed']);
const otherEvaluationStatuses = new Set(['known-limit', 'retired', 'abandoned', 'queued']);

function usage() {
  console.log('Usage: node scripts/eval-report.mjs [--output FILE] [--csv FILE]');
}

function parseArgs(argv) {
  const options = { output: null, html: null, csv: null };
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
  if (!options.output && !options.html && !options.csv) options.output = defaultOutput;
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
  return listFiles(recordsDir, '.md').map((name) => {
    const file = path.join(recordsDir, name);
    const source = fs.readFileSync(file, 'utf8');
    const data = readFrontmatter(file);
    const date = source.match(/^[-*]\s*評価日:\s*(\d{4}-\d{2}-\d{2})\s*$/m)?.[1];
    return { ...data, legacyCompletedAt: date ? `${date}T23:59:59+09:00` : undefined };
  });
}

function readResults() {
  const results = new Map();
  for (const name of listFiles(statesDir, '.result.yml')) {
    const result = yaml.load(fs.readFileSync(path.join(statesDir, name), 'utf8'), { json: true });
    if (result?.evaluationId) results.set(result.evaluationId, result);
  }
  return results;
}

function readStateEvents() {
  const events = [];
  for (const name of listFiles(statesDir, '.events.yml')) {
    const loaded = yaml.load(fs.readFileSync(path.join(statesDir, name), 'utf8'), { json: true });
    if (!Array.isArray(loaded)) continue;
    for (const event of loaded) {
      if (event?.candidateId && event?.status && event?.occurredAt
        && !Number.isNaN(Date.parse(event.occurredAt))) {
        events.push(event);
      }
    }
  }
  return events.sort((a, b) => Date.parse(a.occurredAt) - Date.parse(b.occurredAt));
}

function readFindings() {
  const findingsDir = path.join(evaluationRoot, 'findings');
  const findings = new Map();
  for (const name of listFiles(findingsDir, '.md')) {
    const finding = readFrontmatter(path.join(findingsDir, name));
    if (finding?.id) findings.set(finding.id, finding);
  }
  return findings;
}

function findingCount(record) {
  if (excludedFindingStatuses.has(record.status)) return 0;
  return new Set(Array.isArray(record.findings) ? record.findings : []).size;
}

function findingTypeSummary(records, findings) {
  const counts = new Map();
  for (const record of records) {
    if (excludedFindingStatuses.has(record.status)) continue;
    for (const id of new Set(record.findings ?? [])) {
      const type = findings.get(id)?.discoveryType ?? 'unknown';
      counts.set(type, (counts.get(type) ?? 0) + 1);
    }
  }
  return [...counts.entries()].sort(([a], [b]) => a.localeCompare(b)).map(([type, count]) => ({ type, count }));
}

function evaluationStatusCount(records, status) {
  return records.filter((record) => record.status === status).length;
}

function statusTableRecords(records, datedRecords) {
  const dated = new Set(datedRecords);
  const undatedPending = records.filter((record) =>
    !dated.has(record) && pendingEvaluationStatuses.has(record.status));
  return [...datedRecords, ...undatedPending];
}

function hasCompletedAt(record, results) {
  const result = results.get(record.id);
  const completedAt = result?.completedAt ?? record.legacyCompletedAt;
  return completedAt && !Number.isNaN(Date.parse(completedAt));
}

function formatLocalDateTime(value) {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return String(value ?? '');
  const pad = (number) => String(number).padStart(2, '0');
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`
    + ` ${pad(date.getHours())}:${pad(date.getMinutes())}:${pad(date.getSeconds())}`;
}

function localTimeZone() {
  return Intl.DateTimeFormat().resolvedOptions().timeZone || 'local time';
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

function timeSeries(records, results, findings, stateEvents) {
  const evaluationEvents = records
    .map((record) => ({
      record,
      result: results.get(record.id) ?? (record.legacyCompletedAt ? {
        status: record.status,
        completedAt: record.legacyCompletedAt,
        source: 'legacy evaluation date',
      } : null),
    }))
    .filter(({ result }) => result?.completedAt && !Number.isNaN(Date.parse(result.completedAt)))
    .sort((a, b) => Date.parse(a.result.completedAt) - Date.parse(b.result.completedAt)
      || Number(a.record.legacyNumber ?? 0) - Number(b.record.legacyNumber ?? 0));
  const recordsById = new Map(records.map((record) => [record.id, record]));
  const timeline = [
    ...evaluationEvents.map((event) => ({ ...event, kind: 'evaluation', at: event.result.completedAt })),
    ...stateEvents
      .filter((event) => recordsById.has(event.evaluationId))
      .map((event) => ({ event, kind: 'state', at: event.occurredAt })),
  ].sort((a, b) => Date.parse(a.at) - Date.parse(b.at) || (a.kind === 'evaluation' ? 1 : -1));
  let evaluations = 0;
  let discovered = 0;
  let fixed = 0;
  const discoveredFindingIds = new Set();
  const resolvedFindingIds = new Set();
  const openFindingIds = new Set();
  const candidateStatuses = new Map();
  const applyFindingStatus = (evaluationId, status) => {
    const evaluation = recordsById.get(evaluationId);
    if (!evaluation || excludedFindingStatuses.has(status)) return;
    for (const findingId of new Set(evaluation.findings ?? [])) {
      if (!findings.has(findingId)) continue;
      discoveredFindingIds.add(findingId);
      if (resolvedFindingStatuses.has(status)) {
        openFindingIds.delete(findingId);
        resolvedFindingIds.add(findingId);
      } else if (!resolvedFindingIds.has(findingId)) {
        openFindingIds.add(findingId);
      }
    }
    fixed = resolvedFindingIds.size;
  };
  const rows = timeline.map((item) => {
    let evaluationId;
    let status;
    let area;
    if (item.kind === 'state') {
      const stateEvent = item.event;
      evaluationId = stateEvent.evaluationId;
      status = stateEvent.status;
      const record = recordsById.get(evaluationId);
      area = areaFor(record?.focus);
      candidateStatuses.set(stateEvent.candidateId, status);
      applyFindingStatus(evaluationId, status);
    } else {
      const { record, result } = item;
      evaluationId = record.id;
      status = result.status ?? record.status;
      area = areaFor(record.focus);
      evaluations += 1;
      // The result snapshot is authoritative at this evaluation's own point.
      candidateStatuses.set(record.candidateId ?? record.id, status);
      applyFindingStatus(record.id, status);
    }
    discovered = discoveredFindingIds.size;
    const stateCounts = [...candidateStatuses.values()].reduce((counts, status) => {
      if (status === 'verified-no-bug') counts.nonBug += 1;
      else if (status === 'bug-found' || status === 'fixed') counts.bug += 1;
      else if (pendingEvaluationStatuses.has(status)) counts.pending += 1;
      else if (otherEvaluationStatuses.has(status)) counts.other += 1;
      return counts;
    }, { bug: 0, nonBug: 0, pending: 0, other: 0 });
    return {
      date: new Date(item.at).toISOString(),
      evaluationId,
      evaluationPoint: item.kind === 'evaluation',
      status,
      area,
      evaluations,
      discovered,
      fixed,
      openBugs: openFindingIds.size,
      pendingEvaluations: stateCounts.pending,
      nonBugEvaluations: stateCounts.nonBug,
      bugEvaluations: stateCounts.bug,
      otherEvaluations: stateCounts.other,
    };
  });
  return rows.filter((row, index) => {
    if (index === 0) return true;
    const previous = rows[index - 1];
    return [
      'evaluations', 'bugEvaluations', 'nonBugEvaluations', 'pendingEvaluations',
      'otherEvaluations', 'discovered', 'fixed', 'openBugs',
    ].some((key) => row[key] !== previous[key]);
  });
}

function mdCell(value) {
  return String(value ?? '').replaceAll('|', '\\|').replaceAll('\n', ' ');
}

function renderMarkdown(records, statusRecords, summary, series, findingTypes) {
  const totalBugs = records.reduce((sum, record) => sum + findingCount(record), 0);
  const completed = new Set(series.filter((row) => row.evaluationPoint).map((row) => row.evaluationId)).size;
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
    '## バグ発見種別',
    '',
    '`discoveryType: regression` はレグレッションテストまたは回帰試験で発見したデグレードを表します。',
    '',
    '| 発見種別 | Finding件数 |',
    '|---|---:|',
    ...findingTypes.map((row) => `| ${mdCell(row.type)} | ${row.count} |`),
    '',
    '',
    '## 状態の計上先',
    '',
    '完了日時付き評価を基本に集計し、日時未登録でも現在保留中の評価（needs-excel-probe等）は件数へ含めます。',
    '',
    '| 評価状態 | 件数 | 意味 | 評価分類 | Finding計上 |',
    '|---|---:|---|---|---|',
    `| \`verified-no-bug\` | ${evaluationStatusCount(statusRecords, 'verified-no-bug')} | バグを再現せず評価完了 | 非バグ評価 | なし |`,
    `| \`bug-found\` | ${evaluationStatusCount(statusRecords, 'bug-found')} | バグを確認し修正前 | バグ評価 | 発見Finding |`,
    `| \`fixed\` | ${evaluationStatusCount(statusRecords, 'fixed')} | 確認したバグを修正済み | バグ評価 | 解決済みFinding |`,
    `| \`needs-excel-probe\` | ${evaluationStatusCount(statusRecords, 'needs-excel-probe')} | Excelプローブ作成待ち | 判定保留評価 | 未確定 |`,
    `| \`needs-excel\` | ${evaluationStatusCount(statusRecords, 'needs-excel')} | 実Excel結果の反映待ち | 判定保留評価 | 未確定 |`,
    `| \`blocked\` | ${evaluationStatusCount(statusRecords, 'blocked')} | 外部要因・前提不足で進行不能 | 判定保留評価 | 未確定 |`,
    `| \`in-progress\` | ${evaluationStatusCount(statusRecords, 'in-progress')} | 評価または修正を実行中 | 判定保留評価 | 未確定 |`,
    `| \`claimed\` | ${evaluationStatusCount(statusRecords, 'claimed')} | 評価担当者が取得済み | 判定保留評価 | 未確定 |`,
    `| \`known-limit\` | ${evaluationStatusCount(statusRecords, 'known-limit')} | 既知の仕様上の制限 | その他（制限事項） | 対象外 |`,
    `| \`retired\` | ${evaluationStatusCount(statusRecords, 'retired')} | 仮説または候補を取り下げ | その他 | 対象外 |`,
    `| \`abandoned\` | ${evaluationStatusCount(statusRecords, 'abandoned')} | 評価を中止 | その他 | 対象外 |`,
    `| \`queued\` | ${evaluationStatusCount(statusRecords, 'queued')} | 評価待ち | その他 | 対象外 |`,
    '',
    '## 時系列の収束状況',
    '',
    '結果状態の `completedAt` または旧評価本文の `評価日` を基準にした値です。日時未登録の評価は含みません。評価分類は状態履歴からその時点で有効な候補数、累積評価は別途表示しません。Finding列はさらに別単位の収束指標です。',
    '',
    '| 完了日時 | 評価ID | 状態 | 実装領域 | 評価件数（累積） | バグ状態数 | 非バグ状態数 | 判定保留状態数 | その他状態数 | 発見Finding | 解決済みFinding | 未解決Finding |',
    '|---|---|---|---|---:|---:|---:|---:|---:|---:|---:|---:|',
    ...series.map((row) => `| ${row.date} | ${row.evaluationId} | ${row.status} | ${mdCell(row.area)} | ${row.evaluations} | ${row.bugEvaluations} | ${row.nonBugEvaluations} | ${row.pendingEvaluations} | ${row.otherEvaluations} | ${row.discovered} | ${row.fixed} | ${row.openBugs} |`),
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
  const labels = JSON.stringify(series.map((row) => formatLocalDateTime(row.date)));
  const values = (field) => JSON.stringify(series.map((row) => row[field]));
  return `<div class="chart-container"><canvas id="evaluation-chart" role="img" aria-label="評価分類の累積面グラフ"></canvas></div>
<div class="chart-container"><canvas id="finding-chart" role="img" aria-label="Findingの発見・改修推移"></canvas></div>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.4/dist/chart.umd.min.js"></script>
<script>
const convergenceLabels = ${labels};
new Chart(document.getElementById('evaluation-chart'), {
  type: 'line',
  data: { labels: convergenceLabels, datasets: [
    { label: 'バグ状態数', data: ${values('bugEvaluations')}, borderColor: '#2563eb', backgroundColor: 'rgba(37,99,235,.45)', fill: true, stack: 'evaluation', tension: 0.15 },
    { label: '非バグ状態数', data: ${values('nonBugEvaluations')}, borderColor: '#0891b2', backgroundColor: 'rgba(8,145,178,.45)', fill: true, stack: 'evaluation', tension: 0.15 },
    { label: '判定保留状態数', data: ${values('pendingEvaluations')}, borderColor: '#6b7280', backgroundColor: 'rgba(107,114,128,.45)', fill: true, stack: 'evaluation', tension: 0.15 },
    { label: 'その他状態数', data: ${values('otherEvaluations')}, borderColor: '#92400e', backgroundColor: 'rgba(146,64,14,.45)', fill: true, stack: 'evaluation', tension: 0.15 },
    { label: '累積評価数（面の合計確認用）', data: ${values('evaluations')}, borderColor: '#111827', backgroundColor: 'transparent', fill: false, borderWidth: 3, pointRadius: 0, yAxisID: 'yTotal', tension: 0.15 }
  ]},
  options: { responsive: true, interaction: { mode: 'index', intersect: false },
    plugins: { title: { display: true, text: '評価状態の時点別面グラフ（黒線＝累積評価件数）' }, tooltip: { mode: 'index' } },
    scales: { x: { title: { display: true, text: '評価完了日' } }, y: { beginAtZero: true, stacked: true, title: { display: true, text: 'その時点の状態数' }, ticks: { precision: 0 } }, yTotal: { beginAtZero: true, position: 'right', grid: { drawOnChartArea: false }, title: { display: true, text: '累積評価件数（黒線）' }, ticks: { precision: 0 } } }
  }
});
new Chart(document.getElementById('finding-chart'), {
  type: 'line', data: { labels: convergenceLabels, datasets: [
    { label: '累積発見Finding', data: ${values('discovered')}, borderColor: '#2563eb', backgroundColor: '#2563eb', tension: 0.15 },
    { label: '解決済みFinding', data: ${values('fixed')}, borderColor: '#16a34a', backgroundColor: '#16a34a', tension: 0.15 },
    { label: '未解決Finding', data: ${values('openBugs')}, borderColor: '#dc2626', backgroundColor: '#dc2626', tension: 0.15 }
  ]},
  options: { responsive: true, interaction: { mode: 'index', intersect: false }, plugins: { title: { display: true, text: 'Finding単位のバグ収束' }, tooltip: { mode: 'index' } }, scales: { x: { title: { display: true, text: '評価完了日' } }, y: { beginAtZero: true, title: { display: true, text: '累積Finding数' }, ticks: { precision: 0 } } } }
});
</script>`;
}

function renderHtml(records, statusRecords, summary, series, findingTypes) {
  const totalBugs = records.reduce((sum, record) => sum + findingCount(record), 0);
  const completed = new Set(series.filter((row) => row.evaluationPoint).map((row) => row.evaluationId)).size;
  const timeZone = localTimeZone();
  const summaryRows = summary.map((row) => `<tr><td>${htmlCell(row.area)}</td><td>${row.evaluations}</td><td>${row.bugs}</td><td>${row.unresolved}</td></tr>`).join('\n');
  const findingTypeRows = findingTypes.map((row) => `<tr><td><code>${htmlCell(row.type)}</code></td><td>${row.count}</td></tr>`).join('\n');
  // グラフは全期間を使い、一覧表だけ直近の評価に絞る。
  const recentSeries = series.slice(-10);
  const seriesRows = recentSeries.map((row) => `<tr><td>${htmlCell(formatLocalDateTime(row.date))}</td><td>${htmlCell(row.evaluationId)}</td><td>${htmlCell(row.status)}</td><td>${htmlCell(row.area)}</td><td>${row.evaluations}</td><td>${row.bugEvaluations}</td><td>${row.nonBugEvaluations}</td><td>${row.pendingEvaluations}</td><td>${row.otherEvaluations}</td><td>${row.discovered}</td><td>${row.fixed}</td><td>${row.openBugs}</td></tr>`).join('\n');
  return `<!doctype html>
<html lang="ja"><head><meta charset="utf-8">
<title>VBA Runner 評価レポート</title>
<style>body{font-family:system-ui,sans-serif;line-height:1.5;margin:2rem}table{border-collapse:collapse;margin:1rem 0 2rem}th,td{border:1px solid #bbb;padding:.35rem .6rem;text-align:left}th{background:#eee}td:not(:first-child){text-align:right}code{background:#f3f3f3;padding:.1rem .25rem}.chart-container{max-width:1000px;margin:1rem 0 2rem}</style>
</head><body><h1>評価レポート</h1>
<p>評価件数: ${records.length}、発見バグ件数: ${totalBugs}、完了日時付き: ${completed}、日時未登録: ${records.length - completed}</p>
<h2>実装領域別集計</h2><table><thead><tr><th>実装領域</th><th>評価件数</th><th>バグ件数</th><th>未収束件数</th></tr></thead><tbody>${summaryRows}</tbody></table>
<h2>バグ発見種別</h2><p><code>discoveryType: regression</code> はレグレッションテストまたは回帰試験で発見したデグレードを表します。</p><table><thead><tr><th>発見種別</th><th>Finding件数</th></tr></thead><tbody>${findingTypeRows}</tbody></table>
<h2>時系列の収束状況</h2><p>結果状態の <code>completedAt</code> または旧評価本文の <code>評価日</code> を基準にした値です。日時未登録の評価は含みません。表示日時は生成環境のローカルTZ（${htmlCell(timeZone)}）です。評価分類は評価単位の累積値、判定保留は状態履歴からその時点で有効な候補数、Finding列は別単位の収束指標です。</p>
${renderConvergenceChart(series)}
<h3>評価一覧（直近10件）</h3><table><thead><tr><th>完了日時（ローカルTZ）</th><th>評価ID</th><th>状態</th><th>実装領域</th><th>評価件数（累積）</th><th>バグ状態数</th><th>非バグ状態数</th><th>判定保留状態数</th><th>その他状態数</th><th>発見Finding</th><th>解決済みFinding</th><th>未解決Finding</th></tr></thead><tbody>${seriesRows}</tbody></table>
<h2>評価状態の意味と計上先</h2><p>完了日時付き評価を基本に集計し、日時未登録でも現在保留中の評価（needs-excel-probe等）は件数へ含めます。時系列グラフは完了日時付き評価だけを対象にします。</p><table><thead><tr><th>評価状態</th><th>件数</th><th>意味</th><th>評価分類</th><th>Finding計上</th></tr></thead><tbody>
<tr><td><code>verified-no-bug</code></td><td>${evaluationStatusCount(statusRecords, 'verified-no-bug')}</td><td>バグを再現せず評価完了</td><td>非バグ評価</td><td>なし</td></tr>
<tr><td><code>bug-found</code></td><td>${evaluationStatusCount(statusRecords, 'bug-found')}</td><td>バグを確認し修正前</td><td>バグ評価</td><td>発見Finding</td></tr>
<tr><td><code>fixed</code></td><td>${evaluationStatusCount(statusRecords, 'fixed')}</td><td>確認したバグを修正済み</td><td>バグ評価</td><td>解決済みFinding</td></tr>
<tr><td><code>needs-excel-probe</code></td><td>${evaluationStatusCount(statusRecords, 'needs-excel-probe')}</td><td>Excelプローブ作成待ち</td><td>判定保留評価</td><td>未確定</td></tr>
<tr><td><code>needs-excel</code></td><td>${evaluationStatusCount(statusRecords, 'needs-excel')}</td><td>実Excel結果の反映待ち</td><td>判定保留評価</td><td>未確定</td></tr>
<tr><td><code>blocked</code></td><td>${evaluationStatusCount(statusRecords, 'blocked')}</td><td>外部要因・前提不足で進行不能</td><td>判定保留評価</td><td>未確定</td></tr>
<tr><td><code>in-progress</code></td><td>${evaluationStatusCount(statusRecords, 'in-progress')}</td><td>評価または修正を実行中</td><td>判定保留評価</td><td>未確定</td></tr>
<tr><td><code>claimed</code></td><td>${evaluationStatusCount(statusRecords, 'claimed')}</td><td>評価担当者が取得済み</td><td>判定保留評価</td><td>未確定</td></tr>
<tr><td><code>known-limit</code></td><td>${evaluationStatusCount(statusRecords, 'known-limit')}</td><td>既知の仕様上の制限</td><td>その他（制限事項）</td><td>対象外</td></tr>
<tr><td><code>retired</code></td><td>${evaluationStatusCount(statusRecords, 'retired')}</td><td>仮説または候補を取り下げ</td><td>その他</td><td>対象外</td></tr>
<tr><td><code>abandoned</code></td><td>${evaluationStatusCount(statusRecords, 'abandoned')}</td><td>評価を中止</td><td>その他</td><td>対象外</td></tr>
<tr><td><code>queued</code></td><td>${evaluationStatusCount(statusRecords, 'queued')}</td><td>評価待ち</td><td>その他</td><td>対象外</td></tr>
</tbody></table>
</body></html>\n`;
}

function renderCsv(series) {
  const header = ['completedAt', 'evaluationId', 'status', 'area', 'cumulativeEvaluations', 'bugStateCount', 'nonBugStateCount', 'pendingStateCount', 'otherStateCount', 'cumulativeDiscoveredFindings', 'cumulativeFixedFindings', 'openFindings'];
  return `${header.join(',')}\n${series.map((row) => [row.date, row.evaluationId, row.status, row.area, row.evaluations, row.bugEvaluations, row.nonBugEvaluations, row.pendingEvaluations, row.otherEvaluations, row.discovered, row.fixed, row.openBugs].map(csvCell).join(',')).join('\n')}\n`;
}

function main() {
  const options = parseArgs(process.argv.slice(2));
  const records = readRecords();
  const results = readResults();
  const summary = aggregate(records);
  const findings = readFindings();
  const stateEvents = readStateEvents();
  const series = timeSeries(records, results, findings, stateEvents);
  const datedStatusRecords = records.filter((record) => hasCompletedAt(record, results));
  const statusRecords = statusTableRecords(records, datedStatusRecords);
  const findingTypes = findingTypeSummary(records, findings);
  if (options.output) {
    fs.mkdirSync(path.dirname(options.output), { recursive: true });
    fs.writeFileSync(options.output, renderMarkdown(records, statusRecords, summary, series, findingTypes));
  }
  if (options.html) {
    fs.mkdirSync(path.dirname(options.html), { recursive: true });
    fs.writeFileSync(options.html, renderHtml(records, statusRecords, summary, series, findingTypes));
  }
  if (options.csv) {
    fs.mkdirSync(path.dirname(options.csv), { recursive: true });
    fs.writeFileSync(options.csv, renderCsv(series));
  }
  if (options.output) console.log(options.output);
  if (options.html) console.log(options.html);
  if (options.csv) console.log(options.csv);
}

try {
  main();
} catch (error) {
  console.error(`eval-report: ${error.message}`);
  process.exitCode = 1;
}
