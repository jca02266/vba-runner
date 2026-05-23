/**
 * テーブル駆動リファクタリング候補の検出器（実装版）
 *
 * ApprovalRules_Before.bas のような大量の分岐パターンを検出し、
 * テーブル駆動化の有効性を評価する。
 */

import { Lexer } from '../src/engine/lexer';
import { Parser, Statement, IfStatement, Expression, Identifier } from '../src/engine/parser';

/**
 * テーブル駆動リファクタリング候補の評価結果
 */
export interface TableDrivenCandidate {
    functionName: string;
    startLine: number;
    endLine: number;

    nestingDepth: number;
    outerBranchCount: number;
    innerBranchCount: number;
    totalBranchCombinations: number;

    currentLines: number;
    estimatedAfterLines: number;
    linesToSave: number;
    reductionPercent: number;

    tableRows: number;
    tableColumns: number;
    repeatingPatternCount: number;

    canTableDrive: boolean;
    riskLevel: 'low' | 'medium' | 'high';
    confidenceScore: number;
    recommendation: string;

    reasons: {
        isRepeatingStructure: boolean;
        isSimpleAssignment: boolean;
        isSimpleCondition: boolean;
        branchCountThreshold: boolean;
        hasNoSideEffects: boolean;
    };

    // 診断情報
    diagnostics: {
        detectedPatterns: string[];
        warnings: string[];
    };
}

/**
 * テーブル駆動検出器
 */
export class TableDrivenDetector {
    /**
     * VBA コードを解析して、テーブル駆動の候補を検出
     */
    detectFromCode(code: string): TableDrivenCandidate[] {
        try {
            const tokens = new Lexer(code).tokenize();
            const ast = new Parser(tokens).parse();
            return this.detectFromAST(ast.body);
        } catch (e) {
            console.error('[TableDrivenDetector] Parse error:', e);
            return [];
        }
    }

    /**
     * AST から候補を検出
     */
    private detectFromAST(body: Statement[]): TableDrivenCandidate[] {
        const candidates: TableDrivenCandidate[] = [];

        // Step 1: 全関数を走査
        for (const stmt of body) {
            if (stmt.type === 'ProcedureDeclaration') {
                const proc = stmt as any;
                const procName = proc.name?.name || 'Unknown';
                const procBody = proc.body || [];

                // Step 2: ネストされた if を探す
                const nestedIfs = this.findNestedIfStatements(procBody);
                for (const ifStmt of nestedIfs) {
                    const candidate = this.analyzeIfStatement(
                        ifStmt as IfStatement,
                        procName,
                        procBody
                    );
                    if (candidate && candidate.confidenceScore >= 40) {
                        candidates.push(candidate);
                    }
                }
            }
        }

        return candidates;
    }

    /**
     * ネストされた if-else を検出
     */
    private findNestedIfStatements(body: Statement[]): Statement[] {
        const result: Statement[] = [];

        for (const stmt of body) {
            if (stmt.type === 'IfStatement') {
                if (this.hasDeepElseIfChain(stmt as IfStatement)) {
                    result.push(stmt);
                }
            }
        }

        return result;
    }

    /**
     * if-else-if チェーンが十分に深いか（3段階以上）
     */
    private hasDeepElseIfChain(stmt: IfStatement): boolean {
        let current: any = stmt;
        let depth = 0;

        while (current) {
            if (current.type === 'IfStatement') {
                depth++;
                current = current.alternate;
            } else {
                break;
            }
        }

        return depth >= 3;
    }

    /**
     * if ステートメントを詳しく分析
     */
    private analyzeIfStatement(
        ifStmt: IfStatement,
        funcName: string,
        procBody: Statement[]
    ): TableDrivenCandidate | null {
        const diagnostics = {
            detectedPatterns: [] as string[],
            warnings: [] as string[],
        };

        // Step 1: 外側の分岐チェーンを収集
        const outerChain = this.collectIfElseChain(ifStmt);
        const outerCount = outerChain.length;

        if (outerCount < 3) {
            return null;
        }
        diagnostics.detectedPatterns.push(`外側分岐数: ${outerCount}`);

        // Step 2: 内側の分岐を分析
        const innerStructures = outerChain.map((branch) => this.analyzeConsequent(branch));
        const innerCount = innerStructures[0]?.depth || 0;

        if (innerCount < 2) {
            return null;
        }
        diagnostics.detectedPatterns.push(`内側分岐数: ${innerCount}`);

        // Step 3: 形状の重複を検出
        const shapesMatch = this.shapesAreConsistent(innerStructures);
        if (!shapesMatch) {
            diagnostics.warnings.push('内側の分岐パターンが一致していない');
        }
        diagnostics.detectedPatterns.push(`形状一致: ${shapesMatch}`);

        // Step 4: 代入パターンを分析
        const assignmentAnalysis = this.analyzeAssignments(outerChain);
        diagnostics.detectedPatterns.push(
            `代入パターン: ${assignmentAnalysis.pattern} (複雑度: ${assignmentAnalysis.complexity})`
        );

        if (assignmentAnalysis.complexity > 50) {
            diagnostics.warnings.push('恣意的な修正が検出されました。テーブル化の効果が限定的です。');
        }

        // Step 5: 条件の複雑度
        const conditionComplexity = this.analyzeConditions(outerChain);
        if (conditionComplexity > 30) {
            diagnostics.warnings.push('条件式が複雑です。テーブル化に向きません。');
        }

        // Step 6: スコアを計算
        const confidenceScore = this.calculateConfidenceScore(
            outerCount,
            innerCount,
            shapesMatch,
            assignmentAnalysis.complexity,
            conditionComplexity,
            ifStmt
        );

        // Step 7: 信頼度が低ければ候補から除外
        if (confidenceScore < 40) {
            return null;
        }

        // Step 8: 結果を構築
        const currentLines = this.estimateCurrentLines(ifStmt);
        const estimatedAfterLines = 15 + outerCount * 3 + innerCount * 2;
        const linesToSave = currentLines - estimatedAfterLines;
        const reductionPercent = (linesToSave / currentLines) * 100;

        const riskLevel = this.assessRiskLevel(
            outerCount,
            innerCount,
            assignmentAnalysis.complexity,
            conditionComplexity
        );

        return {
            functionName: funcName,
            startLine: (ifStmt.loc?.start.line || 1) - 1,
            endLine: (ifStmt.loc?.end.line || 1) - 1,

            nestingDepth: this.estimateNestingDepth(ifStmt),
            outerBranchCount: outerCount,
            innerBranchCount: innerCount,
            totalBranchCombinations: outerCount * innerCount,

            currentLines,
            estimatedAfterLines,
            linesToSave,
            reductionPercent: Math.round(reductionPercent * 10) / 10,

            tableRows: outerCount,
            tableColumns: innerCount + 2,
            repeatingPatternCount: outerCount,

            canTableDrive: confidenceScore >= 60,
            riskLevel,
            confidenceScore: Math.round(confidenceScore),
            recommendation: this.generateRecommendation(
                confidenceScore,
                outerCount,
                linesToSave,
                assignmentAnalysis.complexity,
                riskLevel
            ),

            reasons: {
                isRepeatingStructure: shapesMatch,
                isSimpleAssignment: assignmentAnalysis.complexity < 30,
                isSimpleCondition: conditionComplexity < 50,
                branchCountThreshold: outerCount >= 3,
                hasNoSideEffects: assignmentAnalysis.hasSideEffects === false,
            },

            diagnostics,
        };
    }

    /**
     * if-else-if チェーンを配列として収集
     */
    private collectIfElseChain(stmt: IfStatement): IfStatement[] {
        const result: IfStatement[] = [stmt];
        let current: Statement[] | IfStatement | null = stmt.alternate;

        while (current && typeof current === 'object' && (current as any).type === 'IfStatement') {
            result.push(current as IfStatement);
            current = (current as IfStatement).alternate;
        }

        return result;
    }

    /**
     * consequent（分岐の本体）を分析
     */
    private analyzeConsequent(branch: IfStatement): any {
        const statements = branch.consequent;
        let depth = 0;

        // 内側の if を探す
        for (const stmt of statements) {
            if (stmt && stmt.type === 'IfStatement') {
                depth = this.countIfElseDepth(stmt as IfStatement);
                break;
            }
        }

        return { depth, statements };
    }

    /**
     * if-else-if チェーンの深さを計算
     */
    private countIfElseDepth(stmt: IfStatement): number {
        let depth = 0;
        let current: any = stmt;

        while (current && current.type === 'IfStatement') {
            depth++;
            current = current.alternate;
        }

        return depth;
    }

    /**
     * 複数の分岐の形状が一致しているか
     */
    private shapesAreConsistent(structures: any[]): boolean {
        if (structures.length < 2) return true;

        // 簡略化：全て同じ深さなら一致していると判定
        const firstDepth = structures[0].depth;
        return structures.every((s) => s.depth === firstDepth);
    }

    /**
     * 代入パターンを分析
     */
    private analyzeAssignments(
        outerChain: IfStatement[]
    ): { pattern: string; complexity: number; hasSideEffects: boolean } {
        // 簡略化：各分岐の本体を確認
        let complexity = 0;
        let hasSideEffects = false;
        const patterns: string[] = [];

        for (const branch of outerChain) {
            const consequent = branch.consequent;
            if (!consequent) continue;

            const stmts = Array.isArray(consequent) ? consequent : [consequent];
            for (const stmt of stmts) {
                if (!stmt) continue;

                // 関数呼び出しがあれば副作用の可能性
                if (this.hasFunctionCallInTree(stmt)) {
                    hasSideEffects = true;
                    complexity += 20;
                }

                // 複数ステートメントなら複雑度アップ
                if (Array.isArray(stmts) && stmts.length > 2) {
                    complexity += 10;
                }
            }
        }

        const pattern =
            complexity < 20
                ? 'Simple assignments'
                : complexity < 40
                  ? 'Mixed patterns'
                  : 'Complex logic';

        return { pattern, complexity: Math.min(complexity, 100), hasSideEffects };
    }

    /**
     * ツリーに関数呼び出しがあるかチェック
     */
    private hasFunctionCallInTree(node: any): boolean {
        if (!node) return false;

        if (node.type === 'CallExpression') {
            return true;
        }

        // 子ノードを再帰的に走査
        for (const key in node) {
            if (typeof node[key] === 'object' && node[key] !== null) {
                if (this.hasFunctionCallInTree(node[key])) {
                    return true;
                }
            }
        }

        return false;
    }

    /**
     * 条件式の複雑度を分析
     */
    private analyzeConditions(outerChain: IfStatement[]): number {
        let complexity = 0;

        for (const branch of outerChain) {
            const condition = branch.condition;
            if (!condition) continue;

            // BinaryExpression（=, <, >, etc.）なら低複雑度
            if ((condition as any).type === 'BinaryExpression' || (condition as any).type === 'BinaryOp') {
                complexity += 5;
            }
            // LogicalExpression（AND/OR）なら高複雑度
            else if ((condition as any).type === 'LogicalExpression') {
                complexity += 15;
            }
            // その他の複雑な式
            else {
                complexity += 10;
            }
        }

        return complexity;
    }

    /**
     * ネスト深度を推定
     */
    private estimateNestingDepth(stmt: IfStatement): number {
        let depth = 0;
        let current: any = stmt;

        while (current && current.type === 'IfStatement') {
            depth++;
            current = current.consequent;
            // 内側の if を探す
            if (Array.isArray(current)) {
                current = current.find((s) => s && s.type === 'IfStatement');
            } else if (current && current.type === 'BlockStatement') {
                current = current.body?.find((s) => s && s.type === 'IfStatement');
            }
        }

        return Math.min(depth, 3); // 最大 3
    }

    /**
     * 現在の推定行数
     */
    private estimateCurrentLines(stmt: IfStatement): number {
        const startLine = stmt.loc?.start.line || 1;
        const endLine = stmt.loc?.end.line || 1;
        return Math.max(endLine - startLine, 10);
    }

    /**
     * 信頼度スコアを計算
     */
    private calculateConfidenceScore(
        outerCount: number,
        innerCount: number,
        shapesMatch: boolean,
        assignmentComplexity: number,
        conditionComplexity: number,
        ifStmt: IfStatement
    ): number {
        // 基本スコア
        let score = 50;

        // 繰り返し度合い（最大 +30）
        const repetitionBonus = Math.min((outerCount / 10) * 30, 30);
        score += repetitionBonus;

        // 総パターン数（最大 +20）
        const patternBonus = Math.min(((outerCount * innerCount) / 25) * 20, 20);
        score += patternBonus;

        // 形状一致（+15 or -10）
        score += shapesMatch ? 15 : -10;

        // 代入複雑度ペナルティ（-0 to -30）
        const assignmentPenalty = (assignmentComplexity / 100) * 30;
        score -= assignmentPenalty;

        // 条件複雑度ペナルティ（-0 to -15）
        const conditionPenalty = (conditionComplexity / 100) * 15;
        score -= conditionPenalty;

        // 行数によるボーナス（削減が大きいほど+10）
        const currentLines = this.estimateCurrentLines(ifStmt);
        if (currentLines > 40) {
            score += 10;
        }

        return Math.max(0, Math.min(100, score));
    }

    /**
     * リスクレベルを評価
     */
    private assessRiskLevel(
        outerCount: number,
        innerCount: number,
        assignmentComplexity: number,
        conditionComplexity: number
    ): 'low' | 'medium' | 'high' {
        const complexityScore = assignmentComplexity + conditionComplexity;

        if (complexityScore > 60) {
            return 'high';
        } else if (complexityScore > 30) {
            return 'medium';
        }
        return 'low';
    }

    /**
     * 推奨テキストを生成
     */
    private generateRecommendation(
        score: number,
        outerCount: number,
        linesToSave: number,
        assignmentComplexity: number,
        riskLevel: string
    ): string {
        if (score >= 70) {
            return `強く推奨。${outerCount}個の分岐パターンが繰り返されており、最大${linesToSave}行の削減が見込める。${riskLevel === 'low' ? '実装リスクは低い。' : '実装時に注意が必要。'}`;
        } else if (score >= 50) {
            return `検討推奨。テーブル化で約${linesToSave}行削減可能。${assignmentComplexity > 30 ? 'ただし、恣意的修正が検出されたため、事前にロジック整理が必要です。' : ''}`;
        }
        return `検討の余地あり。スコア${Math.round(score)}は中程度です。ビジネスルール抽出やStrategy Patternも検討してください。`;
    }
}

// CLI code moved to separate script to avoid execution during module import
// See: table-driven-detector-cli.ts
