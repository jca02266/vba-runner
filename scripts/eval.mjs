#!/usr/bin/env node

import fs from 'node:fs';
import path from 'node:path';
import crypto from 'node:crypto';
import process from 'node:process';
import * as yaml from 'js-yaml';

const root = process.cwd();
const evalRoot = path.join(root, 'evaluation');
const recordsDir = path.join(evalRoot, 'evaluations');
const campaignsDir = path.join(evalRoot, 'campaigns');
const statesDir = path.join(evalRoot, 'states');
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

function validate(records = readRecords()) {
  const seen = new Set();
  const required = ['id', 'campaign', 'status', 'priority', 'focus'];
  for (const record of records) {
    const { file, data } = record;
    for (const key of required) {
      if (data[key] === undefined || data[key] === null || data[key] === '') {
        throw new Error(`${file}: missing ${key}`);
      }
    }
    if (seen.has(data.id)) throw new Error(`${file}: duplicate id ${data.id}`);
    seen.add(data.id);
    if (!statuses.has(data.status)) throw new Error(`${file}: invalid status ${data.status}`);
    if (typeof data.id !== 'string' || typeof data.campaign !== 'string') {
      throw new Error(`${file}: id and campaign must be strings`);
    }
    if (data.findings !== undefined && !Array.isArray(data.findings)) {
      throw new Error(`${file}: findings must be an array`);
    }
  }
  return records;
}

function render(records, output) {
  const rows = records
    .sort((a, b) => String(a.data.id).localeCompare(String(b.data.id), undefined, { numeric: true }))
    .map(({ data }) => `| ${data.id} | ${data.campaign} | ${data.status} | ${data.focus} |`)
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
  ].join('\n');
  fs.mkdirSync(path.dirname(output), { recursive: true });
  fs.writeFileSync(output, body);
}

function nextCandidate() {
  const candidates = [];
  for (const name of files(campaignsDir, '.yml')) {
    const doc = yaml.load(fs.readFileSync(path.join(campaignsDir, name), 'utf8'), { json: true });
    for (const item of doc?.items ?? []) {
      if (item.status === 'queued') candidates.push(item);
    }
  }
  candidates.sort((a, b) => (a.priority ?? 99) - (b.priority ?? 99) || String(a.id).localeCompare(String(b.id)));
  if (!candidates[0]) throw new Error('no queued candidate');
  console.log(JSON.stringify(candidates[0], null, 2));
}

function claim(id) {
  fs.mkdirSync(statesDir, { recursive: true });
  const target = path.join(statesDir, `${id}.claim.yml`);
  if (fs.existsSync(target)) throw new Error(`${id} is already claimed`);
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
  const now = Date.now();
  const stale = [];
  for (const name of files(statesDir, '.claim.yml')) {
    const state = yaml.load(fs.readFileSync(path.join(statesDir, name), 'utf8'), { json: true });
    if (state?.expiresAt && Date.parse(state.expiresAt) < now) stale.push(name);
  }
  if (stale.length) throw new Error(`stale claims require release: ${stale.join(', ')}`);
  console.log(`audit passed: ${records.length} records, ${files(statesDir, '.claim.yml').length} active claims`);
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
