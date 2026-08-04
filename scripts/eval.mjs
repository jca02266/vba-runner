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
const remediationsDir = path.join(evalRoot, 'remediations');
const rootCausesDir = path.join(evalRoot, 'root-causes');
const campaignsDir = path.join(evalRoot, 'campaigns');
const statesDir = path.join(evalRoot, 'states');
const schemaFile = path.join(evalRoot, 'schema.yml');
const coverageIndexFile = path.join(evalRoot, 'coverage-index.yml');
const excelQueueDir = path.join(root, 'tests', 'excel', 'queue');
const excelQueueSource = path.join(excelQueueDir, 'ExcelQueueVerification.bas');
const excelQueueResult = path.join(root, 'tests', 'excel', 'queue', 'ExcelQueueVerification.result');
const statuses = new Set([
  'queued', 'claimed', 'in-progress', 'verified-no-bug', 'bug-found',
  'fixed', 'blocked', 'abandoned', 'known-limit', 'needs-excel', 'needs-excel-probe', 'retired',
]);
const finalEvaluationStatuses = new Set(['verified-no-bug', 'fixed', 'known-limit', 'retired']);

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

function readRemediations() {
  return files(remediationsDir, '.md').map((name) => readRecord(path.join(remediationsDir, name)));
}

function readRootCauses() {
  return files(rootCausesDir, '.md').map((name) => readRecord(path.join(rootCausesDir, name)));
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
    results.set(result.candidateId, { ...result, file: path.join(statesDir, name) });
  }
  return results;
}

function eventsFile(candidateId) {
  return path.join(statesDir, `${candidateId}.events.yml`);
}

function readEvents(candidateId) {
  const file = eventsFile(candidateId);
  if (!fs.existsSync(file)) return [];
  const events = yaml.load(fs.readFileSync(file, 'utf8'), { json: true });
  if (!Array.isArray(events)) throw new Error(`${file}: events must be an array`);
  return events;
}

function validateEvents(candidateId, events, result) {
  let previousStatus = null;
  for (const [index, event] of events.entries()) {
    if (!event || event.candidateId !== candidateId || !event.evaluationId
      || !statuses.has(event.status) || !event.occurredAt
      || Number.isNaN(Date.parse(event.occurredAt))) {
      throw new Error(`${eventsFile(candidateId)}: invalid event ${index + 1}`);
    }
    if (event.fromStatus !== undefined && event.fromStatus !== previousStatus) {
      throw new Error(`${eventsFile(candidateId)}: event ${index + 1} fromStatus does not match history`);
    }
    if (previousStatus === event.status) {
      throw new Error(`${eventsFile(candidateId)}: event ${index + 1} repeats status ${event.status}`);
    }
    previousStatus = event.status;
  }
  if (result && events.length > 0) {
    const last = events.at(-1);
    if (last.status !== result.status || last.evaluationId !== result.evaluationId) {
      throw new Error(`${eventsFile(candidateId)}: last event does not match result snapshot`);
    }
  }
}

function appendEvent(candidateId, evaluationId, status, previousResult = null) {
  let events = readEvents(candidateId);
  validateEvents(candidateId, events);
  if (events.length === 0 && previousResult) {
    events = [{
      candidateId,
      evaluationId: previousResult.evaluationId,
      status: previousResult.status,
      occurredAt: previousResult.completedAt ?? new Date().toISOString(),
    }];
  }
  const event = {
    candidateId,
    evaluationId,
    status,
    ...(events.length ? { fromStatus: events.at(-1).status } : {}),
    occurredAt: new Date().toISOString(),
  };
  const next = [...events, event];
  validateEvents(candidateId, next);
  fs.mkdirSync(statesDir, { recursive: true });
  const target = eventsFile(candidateId);
  const temporary = `${target}.${process.pid}.${Date.now()}.tmp`;
  fs.writeFileSync(temporary, yaml.dump(next, { noRefs: true, lineWidth: 100 }));
  fs.renameSync(temporary, target);
  return event;
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

function validateExpectation(file, expectation, status) {
  if (expectation === undefined) return;
  if (!expectation || typeof expectation !== 'object' || Array.isArray(expectation)) {
    throw new Error(`${file}: expectation must be an object`);
  }
  const kinds = new Set(['spec', 'excel', 'hypothesis']);
  if (!kinds.has(expectation.kind)) throw new Error(`${file}: invalid expectation.kind ${expectation.kind}`);
  if (typeof expectation.statement !== 'string' || expectation.statement.trim() === '') {
    throw new Error(`${file}: expectation.statement must be a non-empty string`);
  }
  if (!Array.isArray(expectation.references)) throw new Error(`${file}: expectation.references must be an array`);
  for (const reference of expectation.references) {
    if (!reference || typeof reference !== 'object' || Array.isArray(reference) ||
        typeof reference.source !== 'string' || typeof reference.ref !== 'string' ||
        reference.source.trim() === '' || reference.ref.trim() === '') {
      throw new Error(`${file}: expectation references require source and ref`);
    }
  }
  const verification = expectation.verification;
  if (!['pending', 'completed', 'not-required'].includes(verification)) {
    throw new Error(`${file}: invalid expectation.verification ${verification}`);
  }
  if (['spec', 'excel'].includes(expectation.kind) &&
      (expectation.references.length === 0 || verification !== 'completed')) {
    throw new Error(`${file}: ${expectation.kind} expectation requires a completed reference`);
  }
  if (expectation.kind === 'hypothesis' && verification !== 'completed'
      && (finalEvaluationStatuses.has(status) || status === 'bug-found')) {
    throw new Error(`${file}: hypothesis expectation must be verified before terminal status ${status}`);
  }
}

function normalizedExcelSourceHash(directory = excelQueueDir) {
  const manifest = fs.readdirSync(directory)
    .filter((name) => /\.(?:bas|cls|frm)$/i.test(name))
    .sort()
    .map((name) => `${name}\n${fs.readFileSync(path.join(directory, name), 'utf8').replace(/\r\n/g, '\n')}\n`)
    .join('');
  return crypto.createHash('sha256').update(manifest, 'utf8').digest('hex');
}

function excelIds(source) {
  return new Set([...source.matchAll(/\bXL-\d{3}(?:-[A-Z0-9]+)*/g)].map((match) => match[0]));
}

function excelQueueState(data) {
  const requiredIds = data.excelProbeIds ?? [];
  const sourceText = fs.existsSync(excelQueueSource) ? fs.readFileSync(excelQueueSource, 'utf8') : '';
  const sourceIds = excelIds(sourceText);
  const missingSourceIds = requiredIds.filter((id) => !sourceIds.has(id));
  if (missingSourceIds.length > 0) {
    return { requiredState: 'needs-excel-probe', requiredIds, missingSourceIds, missingResultIds: [], complete: false, hashMatches: false, resultReady: false };
  }

  const resultText = fs.existsSync(excelQueueResult) ? fs.readFileSync(excelQueueResult, 'utf8') : '';
  const resultIds = new Set(resultText.split(/\r?\n/)
    .map((line) => line.match(/^(XL-\d{3}(?:-[A-Z0-9]+)*)\b/)?.[1])
    .filter(Boolean));
  const missingResultIds = requiredIds.filter((id) => !resultIds.has(id));
  const complete = /^QUEUE_COMPLETE=True\s*$/m.test(resultText);
  const recordedHash = resultText.match(/^QUEUE_SOURCE_SHA256=([0-9a-f]{64})\s*$/mi)?.[1]?.toLowerCase();
  const sourceHash = fs.existsSync(excelQueueSource) ? normalizedExcelSourceHash() : null;
  const hashMatches = Boolean(recordedHash && sourceHash && recordedHash === sourceHash);
  const resultReady = complete && hashMatches && missingResultIds.length === 0;
  return {
    requiredState: resultReady ? 'result-ready' : 'needs-excel',
    requiredIds,
    missingSourceIds,
    missingResultIds,
    complete,
    hashMatches,
    sourceHash,
    recordedHash: recordedHash ?? null,
    resultReady,
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
    validateExpectation(file, data.expectation, data.status);
    for (const key of schema.record.stringFields ?? []) {
      if (data[key] !== undefined && data[key] !== null && typeof data[key] !== 'string') {
        throw new Error(`${file}: ${key} must be a string`);
      }
    }
    if (data.directCauseKey !== undefined && !data.causeKey) {
      throw new Error(`${file}: directCauseKey requires root causeKey`);
    }
    const fixStatuses = new Set(['unassessed', 'not-required', 'planned', 'in-progress', 'fixed', 'blocked']);
    for (const key of ['directFixStatus', 'rootFixStatus']) {
      if (data[key] !== undefined && !fixStatuses.has(data[key])) {
        throw new Error(`${file}: invalid ${key} ${data[key]}`);
      }
    }
    if (data.rootFixStatus === 'fixed' &&
        (!(data.rootFixTaskId || data.rootFixCandidateId) || !data.rootFixCommit)) {
      throw new Error(`${file}: fixed root cause requires rootFixTaskId (or legacy rootFixCandidateId) and rootFixCommit`);
    }
    for (const key of schema.record.arrays ?? []) {
      if (data[key] !== undefined && !Array.isArray(data[key])) {
        throw new Error(`${file}: ${key} must be an array`);
      }
    }
    if (data.excelProbeIds !== undefined) {
      const invalidId = data.excelProbeIds.find((id) => typeof id !== 'string' || !/^XL-\d{3}(?:-[A-Z0-9]+)*$/.test(id));
      if (invalidId !== undefined) throw new Error(`${file}: invalid excelProbeIds entry ${invalidId}`);
      if (new Set(data.excelProbeIds).size !== data.excelProbeIds.length) {
        throw new Error(`${file}: duplicate excelProbeIds`);
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
    const evaluationNumber = Number.parseInt(String(data.id).replace(/^EV-/, ''), 10);
    const causeAnalysis = data.rootCauseAnalysis;
    if (causeAnalysis !== undefined) {
      if (![0, 1].includes(data.rootCauseProcedureVersion)) {
        throw new Error(`${file}: rootCauseProcedureVersion must be 0 or 1`);
      }
      if (data.rootCauseProcedureVersion === 1 && !data.rootCauseId) {
        throw new Error(`${file}: structured root cause analysis requires rootCauseId`);
      }
      if (data.rootCauseProcedureVersion === 0 && data.rootCauseId) {
        throw new Error(`${file}: procedure version 0 must not claim a rootCauseId`);
      }
    }
    if (Number.isFinite(evaluationNumber) && evaluationNumber >= (schema.record.rootCauseAnalysisFrom ?? Number.MAX_SAFE_INTEGER)) {
      for (const key of schema.rootCauseAnalysis.required) {
        if (!causeAnalysis || (typeof causeAnalysis[key] !== 'string' && !Array.isArray(causeAnalysis[key]))) {
          throw new Error(`${file}: rootCauseAnalysis.${key} is required`);
        }
      }
      if (!['confirmed', 'hypothesis', 'ruled-out', 'not-applicable'].includes(causeAnalysis.status)) {
        throw new Error(`${file}: invalid rootCauseAnalysis.status`);
      }
      for (const key of ['confirmed', 'ruledOut', 'unresolved']) {
        if (!Array.isArray(causeAnalysis[key])) throw new Error(`${file}: rootCauseAnalysis.${key} must be an array`);
      }
    }
    if (data.status === 'needs-excel' && audit.unresolved.length === 0) {
      throw new Error(`${file}: resolved Excel boundaries require a terminal evaluation status`);
    }
    if (['needs-excel-probe', 'needs-excel'].includes(data.status)) {
      if (!data.excelProbeIds?.length) throw new Error(`${file}: ${data.status} requires excelProbeIds`);
      const queue = excelQueueState(data);
      if (queue.resultReady) {
        throw new Error(`${file}: synchronized Excel result is ready; reconcile to a terminal status`);
      }
      if (queue.requiredState !== data.status) {
        throw new Error(`${file}: Excel queue requires ${queue.requiredState}, not ${data.status}`);
      }
    }
  }
  const rootCauses = validateRootCauses(schema, records);
  validateRemediations(schema, records, findings, rootCauses);
  const items = readCampaignItems();
  const recordsById = new Map(records.map(({ data }) => [data.id, data]));
  const results = readResults();
  for (const [candidateId, result] of results) {
    const expectedFile = `${candidateId}.result.yml`;
    if (path.basename(result.file) !== expectedFile) {
      throw new Error(`${result.file}: result filename does not match ${candidateId}`);
    }
    const item = items.get(candidateId);
    if (!item) throw new Error(`${result.file}: unknown candidate ${candidateId}`);
    const evaluation = recordsById.get(result.evaluationId);
    if (!evaluation) throw new Error(`${result.file}: unknown evaluation ${result.evaluationId}`);
    if (evaluation.candidateId !== candidateId) {
      throw new Error(`${result.file}: evaluation ${result.evaluationId} belongs to ${evaluation.candidateId}, not ${candidateId}`);
    }
    if (evaluation.campaign !== item.campaign) {
      throw new Error(`${result.file}: evaluation campaign ${evaluation.campaign} does not match ${item.campaign}`);
    }
    if (evaluation.status !== result.status) {
      throw new Error(`${result.file}: result status ${result.status} does not match evaluation ${evaluation.status}`);
    }
    const events = readEvents(candidateId);
    if (result.stateVersion === 1 && events.length === 0) {
      throw new Error(`${result.file}: stateVersion 1 requires an event history`);
    }
    validateEvents(candidateId, events, result);
  }
  for (const name of files(statesDir, '.events.yml')) {
    const candidateId = name.slice(0, -'.events.yml'.length);
    if (!results.has(candidateId)) throw new Error(`${path.join(statesDir, name)}: missing result snapshot`);
    validateEvents(candidateId, readEvents(candidateId), results.get(candidateId));
  }
  validateCoverageIndex();
  return records;
}

function validateRootCauses(schema, records) {
  const caseSchema = schema.rootCauseCase;
  const evaluations = new Map(records.map((record) => [record.data.id, record.data]));
  const cases = new Map();
  for (const record of readRootCauses()) {
    const { file, data } = record;
    for (const key of caseSchema.required) {
      if (data[key] === undefined || data[key] === null || data[key] === '') {
        throw new Error(`${file}: missing ${key}`);
      }
    }
    if (!/^RC-\d{5}$/.test(data.id)) throw new Error(`${file}: invalid root cause id ${data.id}`);
    if (![1].includes(data.procedureVersion)) throw new Error(`${file}: unsupported root cause procedureVersion ${data.procedureVersion}`);
    if (cases.has(data.id)) throw new Error(`${file}: duplicate root cause id ${data.id}`);
    if (path.basename(file) !== `${data.id}.md`) throw new Error(`${file}: filename does not match ${data.id}`);
    if (!caseSchema.statuses.includes(data.status)) throw new Error(`${file}: invalid root cause status ${data.status}`);
    for (const key of ['originEvaluations', 'evidence', 'alternatives', 'unresolved', 'reviews', 'acceptanceCriteria']) {
      if (!Array.isArray(data[key])) throw new Error(`${file}: ${key} must be an array`);
    }
    if (data.originEvaluations.length === 0) throw new Error(`${file}: originEvaluations must not be empty`);
    for (const id of data.originEvaluations) {
      const evaluation = evaluations.get(id);
      if (!evaluation) throw new Error(`${file}: unknown origin evaluation ${id}`);
      if (evaluation.causeKey !== data.causeKey) {
        throw new Error(`${file}: ${id} causeKey ${evaluation.causeKey} does not match ${data.causeKey}`);
      }
      if (evaluation.rootCauseProcedureVersion !== undefined &&
          evaluation.rootCauseProcedureVersion > data.procedureVersion) {
        throw new Error(`${file}: ${id} uses a newer root cause procedure than ${data.id}`);
      }
    }
    for (const [index, review] of data.reviews.entries()) {
      if (!review || typeof review !== 'object' || Array.isArray(review)
          || typeof review.reviewer !== 'string' || typeof review.summary !== 'string'
          || !caseSchema.reviewVerdicts.includes(review.verdict)) {
        throw new Error(`${file}: invalid review ${index + 1}`);
      }
      if (review.reviewer === data.proposedBy) {
        throw new Error(`${file}: proposer cannot review its own root cause hypothesis`);
      }
    }
    if (data.status === 'confirmed') {
      if (data.reviews.length === 0 || !data.reviews.some((review) => review.verdict === 'confirm')) {
        throw new Error(`${file}: confirmed root cause requires an independent confirming review`);
      }
      if (data.unresolved.length > 0) throw new Error(`${file}: confirmed root cause must not have unresolved questions`);
      if (data.acceptanceCriteria.length === 0) throw new Error(`${file}: confirmed root cause requires acceptanceCriteria`);
    }
    cases.set(data.id, record);
  }
  for (const record of records) {
    if (record.data.rootCauseId && !cases.has(record.data.rootCauseId)) {
      throw new Error(`${record.file}: unknown rootCauseId ${record.data.rootCauseId}`);
    }
  }
  return cases;
}

function validateRemediations(schema, records, findings, rootCauses) {
  const taskSchema = schema.remediationTask;
  const evaluations = new Map(records.map((record) => [record.data.id, record.data]));
  const seen = new Set();
  for (const task of readRemediations()) {
    const { file, data } = task;
    for (const key of taskSchema.required) {
      if (data[key] === undefined || data[key] === null || data[key] === '') {
        throw new Error(`${file}: missing ${key}`);
      }
    }
    if (!/^ROOT-\d{5}$/.test(data.id)) throw new Error(`${file}: invalid remediation id ${data.id}`);
    if (seen.has(data.id)) throw new Error(`${file}: duplicate remediation id ${data.id}`);
    seen.add(data.id);
    if (path.basename(file) !== `${data.id}.md`) throw new Error(`${file}: filename does not match ${data.id}`);
    if (!taskSchema.statuses.includes(data.status)) throw new Error(`${file}: invalid remediation status ${data.status}`);
    for (const key of taskSchema.arrays) {
      if (!Array.isArray(data[key])) throw new Error(`${file}: ${key} must be an array`);
    }
    for (const key of taskSchema.scopeRequired) {
      if (!data.scope || !Array.isArray(data.scope[key])) throw new Error(`${file}: scope.${key} must be an array`);
    }
    if (data.originEvaluations.length === 0) throw new Error(`${file}: originEvaluations must not be empty`);
    if (data.acceptanceCriteria.length === 0) throw new Error(`${file}: acceptanceCriteria must not be empty`);
    for (const id of data.originEvaluations) {
      const evaluation = evaluations.get(id);
      if (!evaluation) throw new Error(`${file}: unknown origin evaluation ${id}`);
      if (evaluation.causeKey !== data.causeKey) {
        throw new Error(`${file}: ${id} causeKey ${evaluation.causeKey} does not match ${data.causeKey}`);
      }
    }
    for (const id of data.originFindings) {
      const finding = findings.get(id)?.data;
      if (!finding) throw new Error(`${file}: unknown origin finding ${id}`);
      if (finding.causeKey !== data.causeKey) {
        throw new Error(`${file}: ${id} causeKey ${finding.causeKey} does not match ${data.causeKey}`);
      }
    }
    const rootCause = rootCauses.get(data.rootCauseId)?.data;
    if (!rootCause) throw new Error(`${file}: unknown rootCauseId ${data.rootCauseId}`);
    if (rootCause.status !== 'confirmed') throw new Error(`${file}: root cause ${data.rootCauseId} is not confirmed`);
    if (rootCause.causeKey !== data.causeKey) {
      throw new Error(`${file}: root cause ${data.rootCauseId} does not match ${data.causeKey}`);
    }
    if (data.procedureVersion !== rootCause.procedureVersion) {
      throw new Error(`${file}: procedureVersion does not match root cause ${data.rootCauseId}`);
    }
    if (data.status === 'completed') {
      if (typeof data.commit !== 'string' || data.commit.trim() === '') {
        throw new Error(`${file}: completed remediation requires commit`);
      }
      if (data.tests.length === 0) throw new Error(`${file}: completed remediation requires tests`);
    } else if (data.commit !== undefined) {
      throw new Error(`${file}: incomplete remediation must not have commit`);
    }
  }
  for (const record of records) {
    const taskId = record.data.rootFixTaskId;
    if (taskId && !seen.has(taskId)) throw new Error(`${record.file}: unknown rootFixTaskId ${taskId}`);
  }
}

function nextRemediation() {
  const task = readRemediations()
    .filter(({ data }) => data.status === 'queued')
    .sort((a, b) => {
      const rank = { critical: 0, high: 1, medium: 2, low: 3 };
      return (rank[a.data.priority] ?? 99) - (rank[b.data.priority] ?? 99)
        || String(a.data.id).localeCompare(String(b.data.id));
    })[0];
  if (!task) throw new Error('no queued remediation task');
  console.log(JSON.stringify(task.data, null, 2));
}

function remediationContext(id) {
  const task = readRemediations().find(({ data }) => data.id === id);
  if (!task) throw new Error(`unknown remediation task ${id}`);
  const records = readRecords();
  const findings = readFindings();
  console.log(JSON.stringify({
    task: task.data,
    evaluations: task.data.originEvaluations.map((evaluationId) =>
      records.find(({ data }) => data.id === evaluationId)?.data),
    findings: task.data.originFindings.map((findingId) => findings.get(findingId)?.data),
  }, null, 2));
}

function render(records, output) {
  const rows = records
    .sort((a, b) => String(a.data.id).localeCompare(String(b.data.id), undefined, { numeric: true }))
    .map(({ data }) => `| ${data.id} | ${data.campaign} | ${data.status} | ${data.rootCauseProcedureVersion ?? ''} | ${data.focus} |`)
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
  const excelPending = records.filter(({ data }) =>
    ['needs-excel-probe', 'needs-excel'].includes(data.status)).length;
  const body = [
    '# 評価状態サマリー',
    '',
    'このファイルは `evaluation/evaluations/*.md` から生成されます。',
    '',
    '| ID | キャンペーン | 状態 | 真因手順 | 対象 |',
    '|---|---|---|---:|---|',
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
  const existingResult = readResults().get(id);
  // A verified finding is intentionally resumable: the discovery loop records
  // the cause first, then a later remediation loop claims the same candidate
  // and completes it as fixed after implementation and regression testing.
  const resumable = existingResult && ['needs-excel', 'needs-excel-probe', 'blocked', 'in-progress', 'bug-found'].includes(existingResult.status);
  if (existingResult && !resumable) throw new Error(`${id} already has a result`);
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
  const previous = readResults().get(candidateId);
  if (previous && !['needs-excel-probe', 'needs-excel', 'blocked', 'in-progress', 'bug-found'].includes(previous.status)) {
    throw new Error(`${candidateId} already has a terminal result`);
  }
  fs.mkdirSync(statesDir, { recursive: true });
  const snapshot = { stateVersion: 1, candidateId, evaluationId, status };
  fs.writeFileSync(target, yaml.dump(snapshot, { noRefs: true }));
  appendEvent(candidateId, evaluationId, status, previous ?? null);
  const claimFile = path.join(statesDir, `${candidateId}.claim.yml`);
  fs.rmSync(claimFile, { force: true });
  console.log(target);
}

function transition(candidateId, evaluationId, status, token) {
  return complete(candidateId, evaluationId, status, token);
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
  console.log(JSON.stringify({ candidate: effectiveCandidate(item, results, claims), relatedEvaluations: related.map(({ data }) => ({ id: data.id, focus: data.focus, status: data.status, expectation: data.expectation, directCauseKey: data.directCauseKey, causeKey: data.causeKey, rootCauseProcedureVersion: data.rootCauseProcedureVersion, rootCauseId: data.rootCauseId, directFixStatus: data.directFixStatus, rootFixStatus: data.rootFixStatus, rootFixCandidateId: data.rootFixCandidateId, rootFixTaskId: data.rootFixTaskId, horizontalAudit: data.horizontalAudit })), uncovered: readCoverageTargets().slice(0, 20) }, null, 2));
}

function excelSync(evaluationId) {
  const record = readRecords().find(({ data }) => data.id === evaluationId);
  if (!record) throw new Error(`unknown evaluation ${evaluationId}`);
  if (!record.data.excelProbeIds?.length) throw new Error(`${evaluationId} has no excelProbeIds`);
  console.log(JSON.stringify({
    evaluationId,
    candidateId: record.data.candidateId,
    currentStatus: record.data.status,
    ...excelQueueState(record.data),
  }, null, 2));
}

function record(file) {
  const parsed = readRecord(path.resolve(file));
  const records = readRecords().filter(({ data }) => data.id !== parsed.data.id);
  validate([...records, parsed]);
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
  console.log('Usage: node scripts/eval.mjs <audit|validate|render|next|context|claim|release|complete|transition|excel-sync|remediation-next|remediation-context> [args]');
}

try {
  const command = process.argv[2];
  if (command === 'validate') {
    const records = validate();
    console.log(`validated ${records.length} evaluation records`);
  } else if (command === 'audit') {
    audit();
  } else if (command === 'render') {
    const output = path.resolve(process.argv[3] ?? path.join(root, 'EVAL_LOG.md'));
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
  } else if (command === 'transition' && process.argv[3] && process.argv[4] && process.argv[5] && process.argv[6]) {
    transition(process.argv[3], process.argv[4], process.argv[5], process.argv[6]);
  } else if (command === 'context' && process.argv[3]) {
    context(process.argv[3], Number(process.argv[4]) || 5);
  } else if (command === 'excel-sync' && process.argv[3]) {
    excelSync(process.argv[3]);
  } else if (command === 'remediation-next') {
    nextRemediation();
  } else if (command === 'remediation-context' && process.argv[3]) {
    remediationContext(process.argv[3]);
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
