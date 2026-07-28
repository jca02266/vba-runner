#!/usr/bin/env node

import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import process from 'node:process';
import * as yaml from 'js-yaml';

const root = process.cwd();
const evalRoot = path.join(root, 'evaluation');
const recordsDir = path.join(evalRoot, 'evaluations');
const findingsDir = path.join(evalRoot, 'findings');
const campaignsDir = path.join(evalRoot, 'campaigns');
const statesDir = path.join(evalRoot, 'states');
const schemaFile = path.join(evalRoot, 'schema.yml');
const coverageIndexFile = path.join(evalRoot, 'coverage-index.yml');
const statuses = new Set([
  'queued', 'claimed', 'in-progress', 'verified-no-bug', 'bug-found',
  'fixed', 'blocked', 'abandoned', 'known-limit', 'needs-excel', 'retired',
]);

function fail(message) {
  console.error(`eval: ${message}`);
  process.exitCode = 1;
}

function files(dir, suffix) {
  if (!fs.existsSync(dir)) return [];
  return fs.readdirSync(dir).filter((name) => !suffix || name.endsWith(suffix));
}

function readRecord(file) {
  const source = fs.readFileSync(file, 'utf8');
  const match = source.match(/^---\r?\n([\s\S]*?)\r?\n---(?:\r?\n|$)/);
  if (!match) throw new Error(`${file}: missing YAML frontmatter`);
  const data = yaml.load(match[1], { json: true });
  if (!data || typeof data !== 'object' || Array.isArray(data)) {
    throw new Error(`${file}: frontmatter must be an object`);
  }
  return { file, source, data };
}

function readRecords() {
  return files(recordsDir, '.md').map((name) => readRecord(path.join(recordsDir, name)));
}

function readFindings() {
  return new Map(files(findingsDir, '.md').map((name) => {
    const record = readRecord(path.join(findingsDir, name));
    return [record.data.id, record];
  }));
}

function readSchema() {
  return yaml.load(fs.readFileSync(schemaFile, 'utf8'), { json: true });
}

function readCampaignItems() {
  const items = new Map();
  for (const name of files(campaignsDir, '.yml')) {
    const doc = yaml.load(fs.readFileSync(path.join(campaignsDir, name), 'utf8'), { json: true });
    if (typeof doc?.id !== 'string' || !Array.isArray(doc.items)) throw new Error(`${name}: invalid campaign file`);
    for (const item of doc.items) {
      if (typeof item.id !== 'string' || typeof item.status !== 'string' || !Number.isFinite(item.priority)) {
        throw new Error(`${name}: invalid campaign item`);
      }
      if (!statuses.has(item.status)) throw new Error(`${name}: invalid campaign status ${item.status}`);
      if (item.coverageTargets !== undefined && !Array.isArray(item.coverageTargets)) {
        throw new Error(`${name}: coverageTargets must be an array`);
      }
      if (items.has(item.id)) throw new Error(`duplicate campaign item ${item.id}`);
      items.set(item.id, { ...item, campaign: doc.id });
    }
  }
  return items;
}

function readClaims({ includeStale = false } = {}) {
  const now = Date.now();
  const claims = new Map();
  for (const name of files(statesDir, '.claim.yml')) {
    const file = path.join(statesDir, name);
    const state = yaml.load(fs.readFileSync(file, 'utf8'), { json: true });
    if (!state?.id || state.status !== 'claimed' || !state.token || !state.owner || !state.expiresAt) {
      throw new Error(`${file}: invalid claim state`);
    }
    const stale = !state?.expiresAt || Date.parse(state.expiresAt) <= now;
    if (includeStale || !stale) claims.set(state.id, { ...state, file, stale });
  }
  return claims;
}

function readResults() {
  const results = new Map();
  for (const name of files(statesDir, '.result.yml')) {
    const result = yaml.load(fs.readFileSync(path.join(statesDir, name), 'utf8'), { json: true });
    if (!result?.candidateId || results.has(result.candidateId)) throw new Error(`duplicate result ${result?.candidateId ?? name}`);
    if (!statuses.has(result.status) || !result.evaluationId) throw new Error(`invalid result ${name}`);
    results.set(result.candidateId, result);
  }
  return results;
}

function readCoverageTargets() {
  if (!fs.existsSync(coverageIndexFile)) return [];
  const index = yaml.load(fs.readFileSync(coverageIndexFile, 'utf8'), { json: true });
  const snapshots = index?.snapshots ?? [];
  const latest = [...snapshots].sort((a, b) => String(b.generatedAt ?? b.commit ?? '').localeCompare(String(a.generatedAt ?? a.commit ?? '')))[0];
  return (latest?.uncovered ?? []).map((entry) => typeof entry === 'string' ? entry : `${entry.file ?? ''}:${entry.line ?? ''}`);
}

function validateCoverageIndex() {
  if (!fs.existsSync(coverageIndexFile)) return;
  const index = yaml.load(fs.readFileSync(coverageIndexFile, 'utf8'), { json: true });
  if (!Array.isArray(index?.snapshots)) throw new Error('coverage-index.yml: snapshots must be an array');
  for (const snapshot of index.snapshots) {
    if (typeof snapshot.id !== 'string' || typeof snapshot.generatedAt !== 'string') {
      throw new Error('coverage snapshot requires id and generatedAt');
    }
    if (snapshot.uncovered !== undefined && !Array.isArray(snapshot.uncovered)) {
      throw new Error(`${snapshot.id}: uncovered must be an array`);
    }
    if (snapshot.report) {
      const report = path.resolve(root, snapshot.report);
      if (!fs.existsSync(report)) throw new Error(`${snapshot.id}: missing coverage report ${snapshot.report}`);
      if (snapshot.sha256 && hash(report) !== snapshot.sha256) throw new Error(`${snapshot.id}: coverage hash mismatch`);
    }
  }
}

function candidateScore(item, uncovered) {
  const targets = item.coverageTargets ?? [];
  const matched = targets.filter((target) => uncovered.some((entry) => String(entry).includes(target)));
  return { score: (item.priority ?? 99) - (matched.length ? 0.5 : 0), coverageMatched: matched };
}

function effectiveCandidate(item, results, claims) {
  const result = results.get(item.id);
  const claim = claims.get(item.id);
  return {
    ...item,
    effectiveStatus: result?.status ?? (claim ? 'claimed' : item.status),
    evaluationId: result?.evaluationId,
    claimOwner: claim?.owner,
  };
}

function validate(records = readRecords()) {
  const schema = readSchema();
  const findings = readFindings();
  const seen = new Set();
  const required = schema.record.required;
  for (const record of records) {
    const { file, data } = record;
    for (const key of required) {
      if (data[key] === undefined || data[key] === null || data[key] === '') {
        throw new Error(`${file}: missing ${key}`);
      }
    }
    if (seen.has(data.id)) throw new Error(`${file}: duplicate id ${data.id}`);
    seen.add(data.id);
    for (const key of Object.keys(data)) {
      if (!(schema.record.allowed ?? []).includes(key)) throw new Error(`${file}: unknown field ${key}`);
    }
    if (!(Number.isFinite(data.priority) || ['critical', 'high', 'medium', 'low'].includes(data.priority))) {
      throw new Error(`${file}: priority must be numeric or a known level`);
    }
    if (!schema.record.statuses.includes(data.status)) throw new Error(`${file}: invalid status ${data.status}`);
    for (const key of schema.record.stringFields ?? []) {
      if (data[key] !== undefined && data[key] !== null && typeof data[key] !== 'string') {
        throw new Error(`${file}: ${key} must be a string`);
      }
    }
    for (const key of schema.record.arrays ?? []) {
      if (data[key] !== undefined && !Array.isArray(data[key])) {
        throw new Error(`${file}: ${key} must be an array`);
      }
    }
    for (const findingId of data.findings ?? []) {
      const finding = findings.get(findingId);
      if (!finding) throw new Error(`${file}: missing finding ${findingId}`);
      if (finding.data.evaluation !== data.id) {
        throw new Error(`${file}: finding ${findingId} points to ${finding.data.evaluation}`);
      }
    }
    const audit = data.horizontalAudit;
    for (const key of schema.horizontalAudit.required) {
      if (!audit || !Array.isArray(audit[key])) throw new Error(`${file}: horizontalAudit.${key} must be an array`);
    }
  }
  readCampaignItems();
  validateCoverageIndex();
  return records;
}

function render(records, output) {
  const rows = records
    .sort((a, b) => String(a.data.id).localeCompare(String(b.data.id), undefined, { numeric: true }))
    .map(({ data }) => `| ${data.id} | ${data.campaign} | ${data.status} | ${data.focus} |`)
    .join('\n');
  const items = [...readCampaignItems().values()];
  const claims = readClaims();
  const results = readResults();
  const uncovered = readCoverageTargets();
  const queue = items
    .filter((item) => item.status === 'queued' && !claims.has(item.id) && !results.has(item.id))
    .map((item) => ({ item: effectiveCandidate(item, results, claims), ...candidateScore(item, uncovered) }))
    .sort((a, b) => a.score - b.score || String(a.item.id).localeCompare(String(b.item.id)))
    .map(({ item, coverageMatched }) => `| ${item.id} | ${item.priority ?? ''} | ${item.effectiveStatus} | ${item.focus ?? item.title} | ${coverageMatched.length ? 'yes' : ''} |`)
    .join('\n');
  const counts = [...new Set(records.map(({ data }) => data.status))]
    .sort().map((status) => `| ${status} | ${records.filter(({ data }) => data.status === status).length} |`).join('\n');
  const unresolved = records.reduce((sum, { data }) => sum + (data.horizontalAudit?.unresolved?.length ?? 0), 0);
  const excelPending = records.filter(({ data }) => data.status === 'needs-excel').length;
  const body = [
    '# 評価状態サマリー',
    '',
    'このファイルは `evaluation/evaluations/*.md` から生成されます。',
    '',
    '| ID | キャンペーン | 状態 | 対象 |',
    '|---|---|---|---|',
    rows,
    '',
    '## 状態集計',
    '',
    '| 状態 | 件数 |',
    '|---|---:|',
    counts,
    '',
    `横展開未解決経路: ${unresolved}、実Excel待ち評価: ${excelPending}`,
    '',
    '## 次の候補',
    '',
    '| ID | 優先度 | 実効状態 | 対象 | coverage一致 |',
    '|---|---:|---|---|---|',
    queue || '| (none) | | | | |',
    '',
  ].join('\n');
  fs.mkdirSync(path.dirname(output), { recursive: true });
  fs.writeFileSync(output, body);
}

function nextCandidate(limit = 1) {
  const candidates = [];
  const claims = readClaims();
  const results = readResults();
  const uncovered = readCoverageTargets();
  readCampaignItems().forEach((item) => {
    if (item.status === 'queued' && !claims.has(item.id) && !results.has(item.id)) {
      candidates.push({ item, ...candidateScore(item, uncovered) });
    }
  });
  candidates.sort((a, b) => a.score - b.score || String(a.item.id).localeCompare(String(b.item.id)));
  if (!candidates[0]) throw new Error('no queued candidate');
  const selected = candidates.slice(0, Math.max(1, Math.min(limit, 20)))
    .map(({ item, coverageMatched }) => ({ ...item, coverageMatched }));
  console.log(JSON.stringify(selected.length === 1 ? selected[0] : selected, null, 2));
}

function claim(id) {
  const item = readCampaignItems().get(id);
  if (!item) throw new Error(`unknown campaign item ${id}`);
  if (item.status !== 'queued') throw new Error(`${id} is not queued (${item.status})`);
  if (readResults().has(id)) throw new Error(`${id} already has a result`);
  fs.mkdirSync(statesDir, { recursive: true });
  const target = path.join(statesDir, `${id}.claim.yml`);
  const existing = readClaims({ includeStale: true }).get(id);
  if (existing && !existing.stale) throw new Error(`${id} is already claimed`);
  if (existing?.stale) fs.rmSync(existing.file);
  const state = {
    id,
    status: 'claimed',
    token: crypto.randomBytes(18).toString('hex'),
    owner: process.env.USER ?? 'unknown',
    claimedAt: new Date().toISOString(),
    expiresAt: new Date(Date.now() + 2 * 60 * 60 * 1000).toISOString(),
  };
  const contents = yaml.dump(state, { noRefs: true, lineWidth: 100 });
  try {
    const descriptor = fs.openSync(target, 'wx');
    fs.writeFileSync(descriptor, contents);
    fs.closeSync(descriptor);
  } catch (error) {
    throw new Error(`${id} claim raced with another process: ${error.message}`);
  }
  console.log(JSON.stringify({ id, token: state.token, expiresAt: state.expiresAt }, null, 2));
}

function ownedClaim(id, token) {
  const claim = readClaims({ includeStale: true }).get(id);
  if (!claim) throw new Error(`${id} is not claimed`);
  if (claim.stale) throw new Error(`${id} claim has expired`);
  if (!token || token !== claim.token) throw new Error(`${id} claim token is invalid`);
  return claim;
}

function audit() {
  const records = validate();
  const stale = [...readClaims({ includeStale: true }).values()].filter((state) => state.stale);
  const archiveDir = path.join(statesDir, 'archive');
  if (stale.length) fs.mkdirSync(archiveDir, { recursive: true });
  for (const state of stale) {
    const archive = path.join(archiveDir, `${path.basename(state.file, '.claim.yml')}.${Date.now()}.claim.yml`);
    fs.renameSync(state.file, archive);
  }
  const active = readClaims();
  console.log(`audit passed: ${records.length} records, ${active.size} active claims, recovered ${stale.length} stale claims`);
}

function release(id, token) {
  const target = path.join(statesDir, `${id}.claim.yml`);
  ownedClaim(id, token);
  fs.rmSync(target);
  console.log(`released ${id}`);
}

function complete(candidateId, evaluationId, status, token) {
  const item = readCampaignItems().get(candidateId);
  if (!item) throw new Error(`unknown campaign item ${candidateId}`);
  ownedClaim(candidateId, token);
  const evaluation = readRecords().find(({ data }) => data.id === evaluationId)?.data;
  if (!evaluation) throw new Error(`unknown evaluation ${evaluationId}`);
  if (evaluation.campaign !== item.campaign) throw new Error(`${candidateId} and ${evaluationId} have different campaigns`);
  if (evaluation.candidateId !== candidateId) {
    throw new Error(`${evaluationId} belongs to ${evaluation.candidateId}, not ${candidateId}`);
  }
  if (!statuses.has(status) || status === 'queued' || status === 'claimed') throw new Error(`invalid completion status ${status}`);
  if (evaluation.status !== status) throw new Error(`${evaluationId} status ${evaluation.status} cannot complete as ${status}`);
  if (status === 'fixed' && (!evaluation.commit || !Array.isArray(evaluation.tests) || evaluation.tests.length === 0)) {
    throw new Error(`${evaluationId} fixed result requires commit and regression tests`);
  }
  const target = path.join(statesDir, `${candidateId}.result.yml`);
  if (fs.existsSync(target)) throw new Error(`${candidateId} already has a result`);
  fs.mkdirSync(statesDir, { recursive: true });
  fs.writeFileSync(target, yaml.dump({ candidateId, evaluationId, status, completedAt: new Date().toISOString() }, { noRefs: true }));
  const claimFile = path.join(statesDir, `${candidateId}.claim.yml`);
  fs.rmSync(claimFile, { force: true });
  console.log(target);
}

function context(candidateId, limit = 5) {
  const item = readCampaignItems().get(candidateId);
  if (!item) throw new Error(`unknown campaign item ${candidateId}`);
  const records = validate();
  const results = readResults();
  const claims = readClaims();
  const related = records
    .filter(({ data }) => data.campaign === item.campaign || (item.priorCauseKey && data.causeKey === item.priorCauseKey))
    .sort((a, b) => String(b.data.id).localeCompare(String(a.data.id), undefined, { numeric: true }))
    .slice(0, Math.max(1, Math.min(limit, 10)));
  console.log(JSON.stringify({ candidate: effectiveCandidate(item, results, claims), relatedEvaluations: related.map(({ data }) => ({ id: data.id, focus: data.focus, status: data.status, causeKey: data.causeKey, horizontalAudit: data.horizontalAudit })), uncovered: readCoverageTargets().slice(0, 20) }, null, 2));
}

function record(file) {
  const parsed = readRecord(path.resolve(file));
  validate([parsed]);
  const target = path.join(recordsDir, `${parsed.data.id}.md`);
  if (fs.existsSync(target)) {
    const existing = fs.readFileSync(target, 'utf8');
    if (existing !== parsed.source) throw new Error(`${parsed.data.id} already exists with different content`);
    console.log(`${target} (unchanged)`);
    return;
  }
  fs.mkdirSync(recordsDir, { recursive: true });
  if (path.resolve(file) !== target) fs.copyFileSync(path.resolve(file), target);
  console.log(target);
}

function migrate(dryRun) {
  const legacy = path.join(root, 'EVAL_LOG.md');
  if (!fs.existsSync(legacy)) throw new Error('EVAL_LOG.md not found');
  const rows = fs.readFileSync(legacy, 'utf8').split(/\r?\n/).map((line) => {
    if (!/^\|\s*\d+\s*\|/.test(line)) return null;
    const body = line.slice(1, line.endsWith('|') ? -1 : undefined);
    const first = body.indexOf('|');
    const last = body.lastIndexOf('|');
    const second = body.indexOf('|', first + 1);
    if (first < 0 || second < 0 || last <= second) return null;
    return [
      null,
      body.slice(0, first).trim(),
      body.slice(first + 1, second).trim(),
      body.slice(second + 1, last).trim(),
      body.slice(last + 1).trim(),
    ];
  }).filter(Boolean);
  const candidates = rows.filter((row) => Number(row[1]) >= 100 && Number(row[1]) <= 188);
  if (dryRun) {
    console.log(JSON.stringify({ source: legacy, candidates: candidates.length }, null, 2));
    return;
  }
  const legacyDir = path.join(evalRoot, 'legacy');
  fs.mkdirSync(legacyDir, { recursive: true });
  fs.copyFileSync(legacy, path.join(legacyDir, 'EVAL_LOG.md'));
  fs.mkdirSync(recordsDir, { recursive: true });
  let created = 0;
  for (const [, number, title, description, date] of candidates) {
    const id = `EV-${String(number).padStart(5, '0')}`;
    const target = path.join(recordsDir, `${id}.md`);
    if (fs.existsSync(target)) continue;
    const text = `${title} ${description}`;
    const campaign = (text.match(/FZ-BUILTIN|FZ-GRAMMAR|MUT-ENGINE/) ?? ['LEGACY'])[0];
    const status = /未実装|恒久的制限/.test(text)
      ? 'known-limit'
      : (/修正済み|修正した|修正。|Bug [^ ]+.*修正/.test(text) ? 'fixed' : 'verified-no-bug');
    const record = [
      '---',
      `id: ${id}`,
      `legacyNumber: '${number}'`,
      `campaign: ${campaign}`,
      `status: ${status}`,
      'priority: medium',
      `focus: ${JSON.stringify(title)}`,
      'coverageSnapshot: null',
      'findings: []',
      'tests: []',
      'horizontalAudit:',
      '  confirmed: []',
      '  ruledOut: []',
      '  unresolved: []',
      '---',
      '',
      `# ${title}`,
      '',
      `- 旧評価番号: ${number}`,
      `- 評価日: ${date}`,
      '',
      description,
      '',
    ].join('\n');
    fs.writeFileSync(target, record);
    created++;
  }
  console.log(JSON.stringify({ source: legacy, candidates: candidates.length, created }, null, 2));
}

function hash(file) {
  return crypto.createHash('sha256').update(fs.readFileSync(file)).digest('hex');
}

function usage() {
  console.log('Usage: node scripts/eval.mjs <audit|validate|render|next|context|claim|release|complete> [args]');
}

try {
  const command = process.argv[2];
  if (command === 'validate') {
    const records = validate();
    console.log(`validated ${records.length} evaluation records`);
  } else if (command === 'audit') {
    audit();
  } else if (command === 'render') {
    const output = path.resolve(process.argv[3] ?? path.join(evalRoot, 'EVAL_LOG.generated.md'));
    render(validate(), output);
    console.log(output);
  } else if (command === 'next') {
    nextCandidate(process.argv[3] === '--limit' ? Number(process.argv[4]) || 1 : 1);
  } else if (command === 'claim' && process.argv[3]) {
    claim(process.argv[3]);
  } else if (command === 'release' && process.argv[3] && process.argv[4]) {
    release(process.argv[3], process.argv[4]);
  } else if (command === 'complete' && process.argv[3] && process.argv[4] && process.argv[5] && process.argv[6]) {
    complete(process.argv[3], process.argv[4], process.argv[5], process.argv[6]);
  } else if (command === 'context' && process.argv[3]) {
    context(process.argv[3], Number(process.argv[4]) || 5);
  } else if (command === 'record' && process.argv[3]) {
    record(process.argv[3]);
  } else if (command === 'migrate' && process.argv[3] === '--dry-run') {
    migrate(true);
  } else if (command === 'migrate' && process.argv[3] === '--apply') {
    migrate(false);
  } else if (command === 'hash' && process.argv[3]) {
    console.log(hash(path.resolve(process.argv[3])));
  } else {
    usage();
    process.exitCode = 1;
  }
} catch (error) {
  fail(error instanceof Error ? error.message : String(error));
}
