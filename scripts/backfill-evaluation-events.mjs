#!/usr/bin/env node

import fs from 'node:fs';
import path from 'node:path';
import process from 'node:process';
import { execFileSync } from 'node:child_process';
import * as yaml from 'js-yaml';

const root = process.cwd();
const evaluationsDir = path.join(root, 'evaluation', 'evaluations');
const statesDir = path.join(root, 'evaluation', 'states');
const statuses = new Set([
  'queued', 'claimed', 'in-progress', 'verified-no-bug', 'bug-found',
  'fixed', 'blocked', 'abandoned', 'known-limit', 'needs-excel',
  'needs-excel-probe', 'retired',
]);
const apply = process.argv.includes('--apply');
const refine = process.argv.includes('--refine');

// A source commit is recorded only when its subject and changed files provide
// direct evidence for the legacy evaluation.  Unmatched records intentionally
// retain the migration timestamp instead of receiving a speculative time.
const historicalEvidence = new Map(Object.entries({
  'EV-00101': '12735b6', 'EV-00102': '82988a4', 'EV-00103': 'd7b2268',
  'EV-00104': 'ccb933b', 'EV-00105': '56576bc', 'EV-00106': 'f2898e2',
  'EV-00107': '7b1089b', 'EV-00108': '23a7a42', 'EV-00109': '856e0f8',
  'EV-00110': 'e4f7ebc', 'EV-00111': 'e6fcdc5', 'EV-00112': 'd324b7d',
  'EV-00113': '41829f2', 'EV-00114': '1a496c0', 'EV-00115': '9bd034b',
  'EV-00116': '2216ead', 'EV-00117': '0add27a', 'EV-00118': 'e767c05',
  'EV-00119': 'a048917', 'EV-00120': 'f6dbb75', 'EV-00121': 'd30ef37',
  'EV-00122': '278cdcf', 'EV-00123': 'c94bc8e', 'EV-00124': '1e96774',
  'EV-00125': '7f9f9aa', 'EV-00126': 'df9e0ed', 'EV-00127': '2f681c2',
  'EV-00128': '5cebc12', 'EV-00129': 'b0a8d61', 'EV-00130': 'e0933a7',
  'EV-00131': 'ea63dcd', 'EV-00132': '8e3dedb', 'EV-00133': '5a57497',
  'EV-00134': 'ede0a93', 'EV-00135': 'deaaceb', 'EV-00136': '6f77667',
  'EV-00137': '0d2f548', 'EV-00138': 'd37b610', 'EV-00139': '7dd769a',
  'EV-00140': '1ad6615', 'EV-00141': '5324bda', 'EV-00142': '1df69c3',
  'EV-00143': '5ec4a91', 'EV-00144': 'b00b9f8', 'EV-00145': '31b560b',
  'EV-00146': 'dced4cb', 'EV-00147': '427c948', 'EV-00148': 'c0f353b',
  'EV-00149': '8f1468c', 'EV-00150': '27d00d6', 'EV-00151': 'c3937d9',
  'EV-00152': '5bf1a84', 'EV-00153': '86ec0d1', 'EV-00154': 'aed8008',
  'EV-00155': '4d59991', 'EV-00156': '64ccd7f', 'EV-00157': '8ff9bce',
  'EV-00158': '20f236e',
  'EV-00159': 'dd7c918', 'EV-00160': '8391e09', 'EV-00161': '38ba66d',
  'EV-00162': 'b86d436', 'EV-00163': 'cc6239b', 'EV-00164': 'dd7c918',
  'EV-00165': 'dd7c918', 'EV-00166': '9cfe600', 'EV-00167': '3526ca4',
  'EV-00168': '2ba7eac', 'EV-00169': '3448535', 'EV-00170': '6064d4f',
  'EV-00171': 'c2681ae', 'EV-00172': '64ccd7f', 'EV-00173': '5c9dd85',
  'EV-00174': '93a8e5b',
  'EV-00175': '2d59435', 'EV-00176': '53fb289', 'EV-00177': 'e7f019b',
  'EV-00178': 'bafad22', 'EV-00179': 'b4279f0',
}));

function files(dir, suffix) {
  return fs.existsSync(dir) ? fs.readdirSync(dir).filter((name) => name.endsWith(suffix)) : [];
}

function frontmatter(file) {
  const source = fs.readFileSync(file, 'utf8');
  const match = source.match(/^---\r?\n([\s\S]*?)\r?\n---(?:\r?\n|$)/);
  if (!match) throw new Error(`${file}: missing YAML frontmatter`);
  const data = yaml.load(match[1], { json: true });
  if (!data || typeof data !== 'object' || Array.isArray(data)) {
    throw new Error(`${file}: frontmatter must be an object`);
  }
  return data;
}

function gitIntroductionDates() {
  const dates = new Map();
  try {
    const output = execFileSync('git', [
      'log', '--format=COMMIT%x09%H%x09%cI', '--diff-filter=A', '--name-only', '--',
      path.relative(root, evaluationsDir),
    ], { encoding: 'utf8' });
    let current = null;
    for (const line of output.split(/\r?\n/)) {
      const header = line.match(/^COMMIT\t([0-9a-f]+)\t(.+)$/);
      if (header) {
        current = { commit: header[1], date: header[2] };
      } else if (current && line.startsWith('evaluation/evaluations/') && !dates.has(line)) {
        dates.set(line, current);
      }
    }
  } catch {
    // Individual Git lookups below will report conflicts when necessary.
  }
  return dates;
}

function gitDate(commit, file, introductionDates) {
  if (commit && /^[0-9a-f]{7,40}$/i.test(String(commit))) {
    try {
      return execFileSync('git', ['show', '-s', '--format=%cI', commit], { encoding: 'utf8' }).trim();
    } catch {
      // Fall through to the file's introduction commit.
    }
  }
  const introduction = introductionDates.get(file);
  if (introduction) return introduction.date;
  return null;
}

function evidenceFor(evaluationId) {
  return historicalEvidence.get(evaluationId) ?? null;
}

function readEvents(candidateId) {
  const file = path.join(statesDir, `${candidateId}.events.yml`);
  if (!fs.existsSync(file)) return [];
  const value = yaml.load(fs.readFileSync(file, 'utf8'), { json: true });
  if (!Array.isArray(value)) throw new Error(`${file}: events must be an array`);
  return value;
}

function readResults() {
  const results = new Map();
  for (const name of files(statesDir, '.result.yml')) {
    const file = path.join(statesDir, name);
    const result = yaml.load(fs.readFileSync(file, 'utf8'), { json: true });
    if (!result?.candidateId) throw new Error(`${file}: missing candidateId`);
    results.set(result.candidateId, result);
  }
  return results;
}

function evNumber(id) {
  const value = Number.parseInt(String(id).replace(/^EV-/, ''), 10);
  return Number.isFinite(value) ? value : Number.MAX_SAFE_INTEGER;
}

function rebuildHistory(candidateId, events) {
  const sorted = [...events].sort((a, b) => Date.parse(a.occurredAt) - Date.parse(b.occurredAt)
    || evNumber(a.evaluationId) - evNumber(b.evaluationId));
  return sorted.map((event, index) => {
    const next = { ...event };
    if (index === 0) delete next.fromStatus;
    else next.fromStatus = sorted[index - 1].status;
    next.candidateId = candidateId;
    return next;
  });
}

function sameEvents(a, b) {
  return JSON.stringify(a) === JSON.stringify(b);
}

function main() {
  const introductionDates = gitIntroductionDates();
  const records = files(evaluationsDir, '.md').map((name) => {
    const file = path.join(evaluationsDir, name);
    const data = frontmatter(file);
    if (!data.id || !data.candidateId || !statuses.has(data.status)) {
      throw new Error(`${file}: id, candidateId, or status is invalid`);
    }
    return { file, data };
  });
  const results = readResults();
  const byCandidate = new Map();
  for (const record of records) {
    const list = byCandidate.get(record.data.candidateId) ?? [];
    list.push(record);
    byCandidate.set(record.data.candidateId, list);
  }

  const additions = [];
  const conflicts = [];
  const changes = [];
  for (const [candidateId, candidateRecords] of byCandidate) {
    const existing = readEvents(candidateId);
    const existingEvaluationIds = new Set(existing.map((event) => event.evaluationId));
    const next = [...existing];
    for (const { file, data } of candidateRecords) {
      const evidenceCommit = evidenceFor(data.id);
      const current = existing.find((event) => event.evaluationId === data.id);
      const recordCommit = data.commit && data.commit !== 'pending' ? data.commit : null;
      if (current && !(refine && (evidenceCommit || recordCommit) && current.inferred && !current.sourceCommit)) continue;
      const occurredAt = evidenceCommit
        ? gitDate(evidenceCommit, path.relative(root, file), introductionDates)
        : gitDate(data.commit, path.relative(root, file), introductionDates);
      if (!occurredAt || Number.isNaN(Date.parse(occurredAt))) {
        conflicts.push({ candidateId, evaluationId: data.id, reason: 'no Git timestamp' });
        continue;
      }
      const event = {
        candidateId,
        evaluationId: data.id,
        status: data.status,
        occurredAt,
        inferred: true,
        source: 'git-history',
        sourceCommit: evidenceCommit ?? recordCommit ?? undefined,
      };
      if (current) {
        const index = next.indexOf(current);
        next[index] = event;
        additions.push({ candidateId, evaluationId: data.id, status: data.status, occurredAt, refined: true });
      } else {
        next.push(event);
        additions.push({ candidateId, evaluationId: data.id, status: data.status, occurredAt });
      }
    }
    const rebuilt = rebuildHistory(candidateId, next);
    const result = results.get(candidateId);
    const last = rebuilt.at(-1);
    if (result && (!last || last.status !== result.status || last.evaluationId !== result.evaluationId)) {
      conflicts.push({ candidateId, reason: 'last event does not match result snapshot' });
      continue;
    }
    if (!sameEvents(existing, rebuilt)) changes.push({ candidateId, events: rebuilt });
  }

  if (apply && conflicts.length === 0) {
    for (const change of changes) {
      const target = path.join(statesDir, `${change.candidateId}.events.yml`);
      const temporary = `${target}.${process.pid}.tmp`;
      fs.writeFileSync(temporary, yaml.dump(change.events, { noRefs: true, lineWidth: 100 }));
      fs.renameSync(temporary, target);
    }
  }
  console.log(JSON.stringify({ mode: apply ? 'apply' : 'dry-run', additions, conflicts, changedCandidates: changes.map((change) => change.candidateId) }, null, 2));
  if (conflicts.length > 0) process.exitCode = 2;
}

try {
  main();
} catch (error) {
  console.error(`backfill-evaluation-events: ${error instanceof Error ? error.message : String(error)}`);
  process.exitCode = 1;
}
