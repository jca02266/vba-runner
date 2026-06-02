import * as fs from 'fs';
import * as path from 'path';
import { TableDrivenDetector } from '../../../test-libs/table-driven-detector';

// Load ApprovalRules_Before.bas
const samplePath = path.resolve(import.meta.dirname, '../../src/vba/ApprovalRules_Before.bas');
const sourceCode = fs.readFileSync(samplePath, 'utf-8');

console.log('=== Table-Driven Refactoring Detector Evaluation ===\n');
console.log(`Testing against: ${path.basename(samplePath)}`);
console.log(`Source length: ${sourceCode.length} chars, ${sourceCode.split('\n').length} lines\n`);

// Run the detector
const detector = new TableDrivenDetector();
const candidates = detector.detectFromCode(sourceCode);

if (candidates.length === 0) {
  console.error('❌ Detector did not identify any table-driven candidates');
  process.exit(1);
}

console.log(`✅ DETECTED ${candidates.length} TABLE-DRIVEN CANDIDATE(S)\n`);

for (let i = 0; i < candidates.length; i++) {
  const candidate = candidates[i];

  console.log(`\n═══════════════════ Candidate #${i + 1} ═══════════════════\n`);

  console.log('=== Function Metrics ===');
  console.log(`Function: ${candidate.functionName}`);
  console.log(`Location: lines ${candidate.startLine}-${candidate.endLine}`);
  console.log(`Current lines of code: ${candidate.currentLines}`);
  console.log(`Estimated after refactoring: ${candidate.estimatedAfterLines}`);
  console.log(`Potential lines saved: ${candidate.linesToSave} (${(candidate.reductionPercent * 100).toFixed(1)}% reduction)\n`);

  console.log('=== Branch Structure ===');
  console.log(`Nesting depth: ${candidate.nestingDepth}`);
  console.log(`Outer branch count (departments): ${candidate.outerBranchCount}`);
  console.log(`Inner branch count (thresholds per dept): ${candidate.innerBranchCount}`);
  console.log(`Total branch combinations: ${candidate.totalBranchCombinations}\n`);

  console.log('=== Table Structure ===');
  console.log(`Table rows (unique patterns): ${candidate.tableRows}`);
  console.log(`Table columns (parameters + result): ${candidate.tableColumns}`);
  console.log(`Repeating patterns detected: ${candidate.repeatingPatternCount}\n`);

  console.log('=== Analysis Details ===');
  console.log(`Can table-drive: ${candidate.canTableDrive}`);
  console.log(`Risk level: ${candidate.riskLevel}`);
  console.log(`Confidence score: ${candidate.confidenceScore}/100\n`);

  console.log('=== Assessment Criteria ===');
  console.log(`✓ Is repeating structure: ${candidate.reasons.isRepeatingStructure}`);
  console.log(`✓ Is simple assignment: ${candidate.reasons.isSimpleAssignment}`);
  console.log(`✓ Is simple condition: ${candidate.reasons.isSimpleCondition}`);
  console.log(`✓ Branch count >= 3: ${candidate.reasons.branchCountThreshold}`);
  console.log(`✓ Has no side effects: ${candidate.reasons.hasNoSideEffects}\n`);

  console.log('=== Detected Patterns ===');
  for (const pattern of candidate.diagnostics.detectedPatterns) {
    console.log(`  • ${pattern}`);
  }

  if (candidate.diagnostics.warnings.length > 0) {
    console.log('\n=== Warnings ===');
    for (const warning of candidate.diagnostics.warnings) {
      console.log(`  ⚠️  ${warning}`);
    }
  }

  console.log('\n=== Recommendation ===');
  console.log(candidate.recommendation);

  console.log('\n=== JSON Output ===');
  console.log(JSON.stringify(candidate, null, 2));
}
