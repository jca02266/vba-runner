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
    for (const item of doc?.items ?? []) {
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
    const stale = !state?.expiresAt || Date.parse(state.expiresAt) <= now;
    if (includeStale || !stale) claims.set(state.id, { ...state, file, stale });
  }
  return claims;
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
  return records;
}

function render(records, output) {
  const rows = records
    .sort((a, b) => String(a.data.id).localeCompare(String(b.data.id), undefined, { numeric: true }))
    .map(({ data }) => `| ${data.id} | ${data.campaign} | ${data.status} | ${data.focus} |`)
    .join('\n');
  const items = [...readCampaignItems().values()];
  const claims = readClaims();
  const queue = items
    .filter((item) => item.status === 'queued' && !claims.has(item.id))
    .sort((a, b) => (a.priority ?? 99) - (b.priority ?? 99) || String(a.id).localeCompare(String(b.id)))
    .map((item) => `| ${item.id} | ${item.priority ?? ''} | ${item.focus ?? item.title} |`)
    .join('\n');
  const body = [
    '# 評価状態サマリー',
    '',
    'このファイルは `evaluation/evaluations/*.md` から生成されます。',
    '',
    '| ID | キャンペーン | 状態 | 対象 |',
    '|---|---|---|---|',
    rows,
    '',
    '## 次の候補',
    '',
    '| ID | 優先度 | 対象 |',
    '|---|---:|---|',
    queue || '| (none) | | |',
    '',
  ].join('\n');
  fs.mkdirSync(path.dirname(output), { recursive: true });
  fs.writeFileSync(output, body);
}

function nextCandidate() {
  const candidates = [];
  const claims = readClaims();
  readCampaignItems().forEach((item) => {
    if (item.status === 'queued' && !claims.has(item.id)) candidates.push(item);
  });
  candidates.sort((a, b) => (a.priority ?? 99) - (b.priority ?? 99) || String(a.id).localeCompare(String(b.id)));
  if (!candidates[0]) throw new Error('no queued candidate');
  console.log(JSON.stringify(candidates[0], null, 2));
}

function claim(id) {
  const item = readCampaignItems().get(id);
  if (!item) throw new Error(`unknown campaign item ${id}`);
  if (item.status !== 'queued') throw new Error(`${id} is not queued (${item.status})`);
  fs.mkdirSync(statesDir, { recursive: true });
  const target = path.join(statesDir, `${id}.claim.yml`);
  const existing = readClaims({ includeStale: true }).get(id);
  if (existing && !existing.stale) throw new Error(`${id} is already claimed`);
  if (existing?.stale) fs.rmSync(existing.file);
  const state = {
    id,
    status: 'claimed',
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
  console.log(target);
}

function audit() {
  const records = validate();
  const stale = [...readClaims({ includeStale: true }).values()].filter((state) => state.stale);
  const active = readClaims();
  console.log(`audit passed: ${records.length} records, ${active.size} active claims, ${stale.length} stale claims`);
}

function release(id) {
  const target = path.join(statesDir, `${id}.claim.yml`);
  if (!fs.existsSync(target)) throw new Error(`${id} is not claimed`);
  fs.rmSync(target);
  console.log(`released ${id}`);
}

function record(file) {
  const parsed = readRecord(path.resolve(file));
  validate([parsed]);
  const target = path.join(recordsDir, `${parsed.data.id}.md`);
  if (fs.existsSync(target) && path.resolve(file) !== target) {
    throw new Error(`${parsed.data.id} already exists`);
  }
  fs.mkdirSync(recordsDir, { recursive: true });
  if (path.resolve(file) !== target) fs.copyFileSync(path.resolve(file), target);
  console.log(target);
}

function migrate(dryRun) {
  const legacy = path.join(root, 'EVAL_LOG.md');
  if (!fs.existsSync(legacy)) throw new Error('EVAL_LOG.md not found');
  const rows = fs.readFileSync(legacy, 'utf8').split(/\r?\n/)
    .map((line) => line.match(/^\|\s*(\d+)\s*\|\s*(.*?)\s*\|\s*(.*?)\s*\|\s*(.*?)\s*\|\s*$/))
    .filter(Boolean);
  const candidates = rows.filter((row) => Number(row[1]) >= 100 && Number(row[1]) <= 188);
  if (!dryRun) throw new Error(`migration is intentionally dry-run only; found ${candidates.length} rows`);
  console.log(JSON.stringify({ source: legacy, candidates: candidates.length }, null, 2));
}

function hash(file) {
  return crypto.createHash('sha256').update(fs.readFileSync(file)).digest('hex');
}

function usage() {
  console.log('Usage: node scripts/eval.mjs <audit|validate|render|next|claim|release> [id]');
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
    nextCandidate();
  } else if (command === 'claim' && process.argv[3]) {
    claim(process.argv[3]);
  } else if (command === 'release' && process.argv[3]) {
    release(process.argv[3]);
  } else if (command === 'record' && process.argv[3]) {
    record(process.argv[3]);
  } else if (command === 'migrate' && process.argv[3] === '--dry-run') {
    migrate(true);
  } else if (command === 'hash' && process.argv[3]) {
    console.log(hash(path.resolve(process.argv[3])));
  } else {
    usage();
    process.exitCode = 1;
  }
} catch (error) {
  fail(error instanceof Error ? error.message : String(error));
}
