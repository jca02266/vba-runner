import * as fs from 'fs';
import { TableDrivenDetector } from '../../../test-libs/table-driven-detector';

const samples = [
    {
        name: 'ApprovalRules_Before',
        path: '/Users/koji/src/github.com/jca02266/vba-compiler/sample/src/vba/ApprovalRules_Before.bas',
        description: 'Original: nested if-else pattern'
    },
    {
        name: 'ApprovalRules_After',
        path: '/Users/koji/src/github.com/jca02266/vba-compiler/sample/src/vba/ApprovalRules_After.bas',
        description: 'Refactored: table-driven with Type and initialization'
    },
    {
        name: 'ApprovalRules_Advanced',
        path: '/Users/koji/src/github.com/jca02266/vba-compiler/sample/src/vba/ApprovalRules_Advanced.bas',
        description: 'Advanced: Dictionary-based dynamic rules'
    }
];

const detector = new TableDrivenDetector();

console.log('=== Table-Driven Refactoring Detector - Multi-File Test ===\n');

for (const sample of samples) {
    console.log(`\n${'='.repeat(70)}`);
    console.log(`📄 ${sample.name}`);
    console.log(`${sample.description}`);
    console.log('='.repeat(70));

    try {
        const sourceCode = fs.readFileSync(sample.path, 'utf-8');
        const lines = sourceCode.split('\n').length;
        console.log(`Source: ${lines} lines`);

        const candidates = detector.detectFromCode(sourceCode);

        if (candidates.length === 0) {
            console.log('Result: ❌ No table-driven candidates detected');
        } else {
            console.log(`Result: ✅ ${candidates.length} candidate(s) found\n`);

            for (let i = 0; i < candidates.length; i++) {
                const cand = candidates[i];
                console.log(`  Candidate #${i + 1}: ${cand.functionName}`);
                console.log(`    Score: ${cand.confidenceScore}/100`);
                console.log(`    Risk: ${cand.riskLevel}`);
                console.log(`    Pattern: ${cand.outerBranchCount} × ${cand.innerBranchCount} combinations`);
                console.log(`    Impact: ${cand.currentLines} → ${cand.estimatedAfterLines} lines (save ${cand.linesToSave})`);
                if (cand.diagnostics.warnings.length > 0) {
                    console.log(`    ⚠️  Warnings:`);
                    for (const warn of cand.diagnostics.warnings) {
                        console.log(`      - ${warn}`);
                    }
                }
            }
        }
    } catch (err) {
        const error = err as any;
        console.log(`Result: ❌ Error`);
        console.log(`  ${error.message}`);
    }
}

console.log(`\n${'='.repeat(70)}`);
console.log('Analysis Summary');
console.log('='.repeat(70));

// Summary
const results = samples.map(sample => {
    try {
        const sourceCode = fs.readFileSync(sample.path, 'utf-8');
        const candidates = detector.detectFromCode(sourceCode);
        return {
            name: sample.name,
            candidates: candidates.length,
            maxScore: candidates.length > 0 ? Math.max(...candidates.map(c => c.confidenceScore)) : 0
        };
    } catch {
        return { name: sample.name, candidates: 0, maxScore: 0 };
    }
});

console.log('\n| File | Candidates | Max Score |');
console.log('|------|-------------|-----------|');
for (const result of results) {
    const status = result.candidates > 0 ? '✅' : '❌';
    console.log(`| ${status} ${result.name.padEnd(20)} | ${String(result.candidates).padEnd(11)} | ${String(result.maxScore).padEnd(9)} |`);
}

console.log('\n📊 Interpretation:');
console.log('  ✅ ApprovalRules_Before: Expected to be a strong candidate (nested if-else)');
console.log('  ❌ ApprovalRules_After: Should NOT be detected (already table-driven, loops instead of branches)');
console.log('  ❌ ApprovalRules_Advanced: May or may not detect (Dictionary-based, complex structures)');
