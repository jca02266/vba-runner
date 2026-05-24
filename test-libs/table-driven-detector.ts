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

    // デシジョンテーブル（抽出できた場合のみ）
    decisionTable?: DecisionTable;

    // 副作用の均一性分析
    sideEffects?: SideEffectAnalysis;

    // 条件式に使われているローカル変数
    conditionVariables?: LocalVariableInfo[];

    // T-04: 各ネストレベルの条件形状情報
    levelShapes?: LevelShapeInfo[];
}

/**
 * デシジョンテーブルの1ルール（1列分）
 */
export interface DecisionRule {
    condition: string;  // e.g., "amount < 50000"
    result: string;     // e.g., "Manager"
}

/**
 * デシジョンテーブルの1行（1キー分）
 */
export interface DecisionTableRow {
    key: string;            // e.g., "Sales" or composite "Capex|Emergency"
    keyPath?: string[];     // e.g., ["Engineering", "Capex", "Emergency"]
    rules: DecisionRule[];  // 各しきい値条件とその結果（最後のルールを除く）
    defaultResult: string;  // else の結果
}

/**
 * デシジョンテーブル全体
 */
export interface DecisionTable {
    keyVariable: string;      // e.g., "department"
    valueVariable: string;    // e.g., "amount"
    rows: DecisionTableRow[];
}

/**
 * 条件式に使われているローカル変数の情報
 */
export interface LocalVariableInfo {
    name: string;
    level?: number;                  // T-02: ネストレベル（0=外側キー, 1=第2キー, …）
    declaredAt?: number;             // Dim 宣言行（1始まり）。引数の場合は undefined
    assignedLines: number[];         // if チェーン開始より前にある代入文の行番号リスト
    assignmentAstShapes?: string[];  // T-01: 各代入文の RHS 形状（computeConditionShape 形式）
    hasUniformAssignment?: boolean;  // T-01: 全代入が同一形状か（代入部分もテーブル駆動化できるか）
}

/**
 * ネストレベルごとの条件形状情報（T-04）
 */
export interface LevelShapeInfo {
    level: number;
    shape: string;        // 代表形状（最初のブランチの条件形状）
    isUniform: boolean;   // 全ブランチで条件形状が一致しているか
    nonUniformBranches?: string[];  // 形状が異なるブランチのキー値
}

/**
 * 副作用（結果代入以外の文）の均一性分析
 */
export interface SideEffectAnalysis {
    hasUniformStructure: boolean;   // 全ブランチで副作用の構造が同一か
    sideEffectCount: number;        // 副作用文数（主代入を除く）
    structureSummary: string;       // 副作用の AST 形状サマリー
    mismatchedBranches?: string[];  // 不一致があったブランチのキーパス
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

        // Step 5: 条件の複雑度（ブランチあたり平均: 5=単純, 15=複雑）
        const conditionComplexity = this.analyzeConditions(outerChain);
        if (conditionComplexity > 8) {
            diagnostics.warnings.push('条件式が複雑です。テーブル化に向きません。');
        }

        // Step 6: スコアを計算
        const rawScore = this.calculateConfidenceScore(
            outerCount,
            innerCount,
            shapesMatch,
            assignmentAnalysis.complexity,
            conditionComplexity,
            ifStmt
        );

        // Step 6.5: 副作用の均一性を分析
        const sideEffects = this.analyzeSideEffects(outerChain, funcName);
        let confidenceScore = rawScore;
        if (sideEffects.sideEffectCount > 0) {
            if (sideEffects.hasUniformStructure) {
                diagnostics.detectedPatterns.push(
                    `副作用: ${sideEffects.sideEffectCount}文（構造均一 → テーブル化可能）`
                );
            } else {
                confidenceScore = Math.max(0, rawScore - 20);
                diagnostics.warnings.push(
                    `副作用の構造が不均一（${(sideEffects.mismatchedBranches ?? []).length}ブランチが不一致）`
                );
            }
        }

        // Step 6.6: 各レベルの条件形状を分析（T-04）
        const levelShapes = this.computeLevelShapes(outerChain);
        const nonUniformLevels = levelShapes.filter((ls) => !ls.isUniform);
        if (nonUniformLevels.length > 0) {
            confidenceScore = Math.max(0, confidenceScore - 15);
            for (const ls of nonUniformLevels) {
                diagnostics.warnings.push(
                    `Level ${ls.level} の条件形状が不均一（基準: "${ls.shape}"、異形ブランチ: ${(ls.nonUniformBranches ?? []).join(', ')}）`
                );
            }
        } else {
            diagnostics.detectedPatterns.push(
                `全レベルの条件形状が均一（${levelShapes.map((ls) => `L${ls.level}: ${ls.shape}`).join(', ')}）`
            );
        }

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
                isSimpleCondition: conditionComplexity <= 6,
                branchCountThreshold: outerCount >= 3,
                hasNoSideEffects: assignmentAnalysis.hasSideEffects === false,
            },

            diagnostics,

            decisionTable: this.extractDecisionTable(outerChain),
            sideEffects,
            conditionVariables: this.findConditionVariables(ifStmt, procBody),
            levelShapes,
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
                // ネストされた IfStatement は analyzeSideEffects / computeLevelShapes で
                // 別途分析するためここではスキップする（deep-walk による誤カウント防止）
                if ((stmt as any).type === 'IfStatement') continue;

                // AssignmentStatement / SetStatement の RHS 呼び出しはgetter（副作用なし）として扱う
                // 制限: RHS 関数が実際に副作用を持つ場合（ログ・DB更新など）は検出できない（T-07 参照）
                const sType = (stmt as any).type;
                const isAssignment = sType === 'AssignmentStatement' || sType === 'SetStatement';
                if (!isAssignment && this.hasFunctionCallInTree(stmt)) {
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
     * 条件式の複雑度を分析（ブランチあたりの平均値を返す）
     * 分岐数は repetitionBonus で別途評価するため、ここでは 1 ブランチあたりの
     * 条件複雑度を返す。BinaryExpression=5, LogicalExpression=15, その他=10。
     */
    private analyzeConditions(outerChain: IfStatement[]): number {
        if (outerChain.length === 0) return 0;
        let total = 0;

        for (const branch of outerChain) {
            const condition = branch.condition;
            if (!condition) continue;

            // BinaryExpression（=, <, >, etc.）なら低複雑度
            if ((condition as any).type === 'BinaryExpression' || (condition as any).type === 'BinaryOp') {
                total += 5;
            }
            // LogicalExpression（AND/OR）なら高複雑度
            else if ((condition as any).type === 'LogicalExpression') {
                total += 15;
            }
            // その他の複雑な式
            else {
                total += 10;
            }
        }

        // ブランチあたり平均（5=全BinaryExpression, 15=全LogicalExpression）
        return total / outerChain.length;
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
        // conditionComplexity は 1 ブランチ平均値（5=単純, 15=複雑）のため /15 で正規化
        const conditionPenalty = (conditionComplexity / 15) * 15;
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

    // -------------------------------------------------------------------------
    // デシジョンテーブル抽出
    // -------------------------------------------------------------------------

    /**
     * 各ネストレベルの条件形状を収集する（T-04）
     * 外側チェーン = level 0、その consequent 内の内側チェーン = level 1、以降再帰
     */
    private computeLevelShapes(outerChain: IfStatement[], level = 0): LevelShapeInfo[] {
        if (outerChain.length === 0) return [];

        // このレベルの全ブランチの条件形状を収集
        const branchShapes = outerChain.map((branch) => ({
            keyValue: this.extractKeyValue(branch.condition as any),
            shape: this.computeConditionShape(branch.condition as any),
        }));
        const firstShape = branchShapes[0].shape;
        const nonUniform = branchShapes.filter((b) => b.shape !== firstShape);

        const levelInfo: LevelShapeInfo = {
            level,
            shape: firstShape,
            isUniform: nonUniform.length === 0,
            ...(nonUniform.length > 0 && {
                nonUniformBranches: nonUniform.map((b) => b.keyValue),
            }),
        };

        // 内側レベルを再帰的に収集
        // 各外側ブランチの consequent から内側 if を取得し、内側チェーンを収集
        const innerChains: IfStatement[][] = [];
        for (const branch of outerChain) {
            const stmts: any[] = Array.isArray(branch.consequent) ? branch.consequent : [];
            const innerIf = stmts.find((s: any) => s?.type === 'IfStatement') as IfStatement | undefined;
            if (innerIf) {
                innerChains.push(this.collectIfElseChain(innerIf));
            }
        }

        const innerResults: LevelShapeInfo[] = [];
        if (innerChains.length > 0) {
            // 全ブランチの内側チェーンを統合してレベル+1 を再帰
            const allInnerBranches = innerChains.flat();
            innerResults.push(...this.computeLevelShapes(allInnerBranches, level + 1));
        }

        return [levelInfo, ...innerResults];
    }

    /**
     * 外側の if-else-if チェーンからデシジョンテーブルを構築する
     * キーレベルが複数ある場合（多段ネスト）は再帰的に処理する
     */
    private extractDecisionTable(outerChain: IfStatement[]): DecisionTable | undefined {
        if (outerChain.length === 0) return undefined;
        const keyVariable = this.extractKeyVariable(outerChain[0].condition as any);

        const nested = this.extractNestedRows(outerChain, []);
        if (nested.length === 0 || !keyVariable) return undefined;

        const valueVariable = nested.find((r) => r.rules.length > 0)?.innerVariable ?? '';
        const rows: DecisionTableRow[] = nested.map((r) => ({
            key: r.keyPath[r.keyPath.length - 1] ?? '?',
            keyPath: r.keyPath,
            rules: r.rules,
            defaultResult: r.defaultResult,
        }));

        return { keyVariable, valueVariable, rows };
    }

    /**
     * 条件式からキー式を取得する（T-03: 形状比較ファースト対応）
     * ブランチ間で変化しない「左辺」の式全体を返す。
     * 例: department = "Sales"            → "department"
     *     task.Category = "Infrastructure" → "task.Category"
     *     GetPriority(taskType) = "High"   → "GetPriority(taskType)"
     *     amount < 50000                   → "amount"
     */
    private extractKeyVariable(condition: any): string {
        if (!condition) return '';
        if (condition.type === 'BinaryExpression') {
            // 右辺がリテラル → 左辺がキー式
            if (condition.right?.type === 'StringLiteral' || condition.right?.type === 'NumberLiteral') {
                return this.formatExprFull(condition.left);
            }
            // 左辺がリテラル → 右辺がキー式
            if (condition.left?.type === 'StringLiteral' || condition.left?.type === 'NumberLiteral') {
                return this.formatExprFull(condition.right);
            }
            // しきい値比較（IDENT < IDENT など）→ 左辺を返す
            return this.formatExprFull(condition.left);
        }
        if (condition.type === 'CallExpression') return this.formatExprFull(condition);
        if (condition.type === 'UnaryExpression' && condition.operator === 'Not') {
            return this.formatExprFull(condition.argument);
        }
        return '';
    }

    /** 外側条件からキー値を取得（例: "Sales", "IsHighPriority", "GetLevel=1"） */
    private extractKeyValue(condition: any): string {
        if (!condition) return '?';
        if (condition.type === 'BinaryExpression') {
            // var = "literal"
            const lit = condition.right?.type === 'StringLiteral' ? condition.right
                      : condition.left?.type === 'StringLiteral'  ? condition.left : null;
            if (lit) return String(lit.value ?? '?');
            // FunctionCall(var) = literal  →  "FuncName=val"
            const call = condition.left?.type === 'CallExpression' ? condition.left : null;
            if (call) {
                const funcName = String((call.callee as any)?.name ?? '?');
                const val = this.formatExpr(condition.right);
                return `${funcName}=${val}`;
            }
            // var = ConstantIdentifier（定数識別子との比較: direction = xlUp）
            // NumericLiteral はしきい値として別途処理されるためここには来ない
            const rightType = condition.right?.type;
            if (rightType === 'Identifier' || rightType === 'MemberExpression') {
                return this.formatExprFull(condition.right);
            }
            const leftType = condition.left?.type;
            if (leftType === 'Identifier' || leftType === 'MemberExpression') {
                return this.formatExprFull(condition.left);
            }
        }
        // FunctionCall(var)  →  関数名をキー値に
        if (condition.type === 'CallExpression') {
            return String((condition.callee as any)?.name ?? '?');
        }
        // Not FunctionCall(var)  →  "Not FuncName"
        if (condition.type === 'UnaryExpression' && condition.operator === 'Not') {
            if (condition.argument?.type === 'CallExpression') {
                const name = String((condition.argument.callee as any)?.name ?? '?');
                return `Not ${name}`;
            }
        }
        return '?';
    }

    /** CallExpression の第1引数の識別子名を取得 */
    private extractFirstArgName(call: any): string {
        const args: any[] = call.args ?? [];
        const first = args[0];
        if (first?.type === 'Identifier') return String(first.name ?? '');
        return '';
    }

    /**
     * 再帰的にキーレベルを辿り、末端のしきい値ルールを収集する
     * - 条件が文字列等値なら「キーレベル」として再帰
     * - 条件が数値比較なら「しきい値レベル」としてルール抽出
     */
    private extractNestedRows(
        chain: IfStatement[],
        keyPath: string[]
    ): Array<{ keyPath: string[]; rules: DecisionRule[]; defaultResult: string; innerVariable: string }> {
        const result: Array<{ keyPath: string[]; rules: DecisionRule[]; defaultResult: string; innerVariable: string }> = [];

        for (const branch of chain) {
            const keyValue = this.extractKeyValue(branch.condition as any);
            const currentPath = [...keyPath, keyValue];

            const stmts = branch.consequent;
            const innerIf = stmts.find((s) => s?.type === 'IfStatement') as IfStatement | undefined;

            if (!innerIf) {
                // 葉ノード: 直接代入
                result.push({
                    keyPath: currentPath,
                    rules: [],
                    defaultResult: this.extractResult(stmts),
                    innerVariable: '',
                });
                continue;
            }

            if (this.isNumericThresholdCondition(innerIf.condition as any)) {
                // しきい値レベル: ルール抽出
                const extracted = this.extractThresholdRulesFromIfChain(innerIf);
                result.push({ keyPath: currentPath, ...extracted });
            } else {
                // キーレベル（文字列等値・関数呼び出し・その他）: 再帰
                const innerChain = this.collectIfElseChain(innerIf);
                result.push(...this.extractNestedRows(innerChain, currentPath));
            }
        }

        return result;
    }

    /**
     * 条件が数値しきい値比較（葉レベル）か判定
     * - <, <=, >, >= : 片側が NumberLiteral または Identifier（定数名）
     * - =            : 片側が NumberLiteral のみ
     *                  （= + Identifier は direction = xlUp のようなキー等値比較と区別できないため）
     * それ以外（文字列等値、関数呼び出し、論理演算など）はすべてキーレベルとして再帰する
     */
    private isNumericThresholdCondition(condition: any): boolean {
        if (condition?.type !== 'BinaryExpression') return false;
        const op: string = condition.operator ?? '';
        if (!['<', '<=', '>', '>=', '='].includes(op)) return false;
        const leftType  = condition.left?.type  ?? '';
        const rightType = condition.right?.type ?? '';
        if (op === '=') {
            // = は即値のみ（定数識別子はキーレベルとして処理）
            return leftType === 'NumberLiteral' || rightType === 'NumberLiteral';
        }
        // <, <=, >, >= は即値または定数識別子を許可
        const isNumericLike = (t: string) => t === 'NumberLiteral' || t === 'Identifier';
        return isNumericLike(leftType) || isNumericLike(rightType);
    }

    /**
     * 数値しきい値の if-else-if チェーンからルールを抽出する
     */
    private extractThresholdRulesFromIfChain(ifStmt: IfStatement): {
        rules: DecisionRule[];
        defaultResult: string;
        innerVariable: string;
    } {
        const rules: DecisionRule[] = [];
        let defaultResult = '?';
        let innerVariable = '';
        let current: any = ifStmt;

        while (current && current.type === 'IfStatement') {
            if (!innerVariable) innerVariable = this.extractKeyVariable(current.condition);
            rules.push({
                condition: this.formatCondition(current.condition),
                result: this.extractResult(current.consequent as Statement[]),
            });

            if (Array.isArray(current.alternate)) {
                defaultResult = this.extractResult(current.alternate as Statement[]);
                break;
            }
            current = current.alternate ?? null;
        }

        return { rules, defaultResult, innerVariable };
    }

    /** 条件式を文字列にフォーマット（例: "amount < 50000"） */
    private formatCondition(condition: any): string {
        if (!condition) return '?';
        if (condition.type === 'BinaryExpression') {
            const left  = this.formatExpr(condition.left);
            const op    = condition.operator ?? '?';
            const right = this.formatExpr(condition.right);
            return `${left} ${op} ${right}`;
        }
        return '?';
    }

    /** 式ノードを文字列にフォーマット（単純型のみ、後方互換用） */
    private formatExpr(expr: any): string {
        if (!expr) return '?';
        if (expr.type === 'Identifier')    return String(expr.name ?? '?');
        if (expr.type === 'NumberLiteral') return this.formatNumber(expr.value);
        if (expr.type === 'StringLiteral') return String(expr.value ?? '?');
        return '?';
    }

    /**
     * 式ノードを完全な文字列に展開する（T-03: MemberExpression・CallExpression に対応）
     * 例: task.Category              → "task.Category"
     *     GetPriority(taskType)      → "GetPriority(taskType)"
     *     Sheet1.Cells(row, 1).Value → "Sheet1.Cells(row, 1).Value"
     */
    private formatExprFull(expr: any): string {
        if (!expr) return '?';
        if (expr.type === 'Identifier')    return String(expr.name ?? '?');
        if (expr.type === 'StringLiteral') return String(expr.value ?? '?');
        if (expr.type === 'NumberLiteral') return this.formatNumber(expr.value);
        if (expr.type === 'MemberExpression') {
            return `${this.formatExprFull(expr.object)}.${expr.property?.name ?? '?'}`;
        }
        if (expr.type === 'CallExpression') {
            // callee が MemberExpression（Sheet1.Cells）の場合も再帰展開
            const callee = this.formatExprFull(expr.callee);
            const args = (expr.args ?? []).map((a: any) => this.formatExprFull(a)).join(', ');
            return `${callee}(${args})`;
        }
        return '?';
    }

    /**
     * 条件式・代入式を正規化した形状文字列に変換する（T-03/T-01 共用）
     * 変数（データ）は IDENT/STR/NUM に置換し、構造（関数名・プロパティ名）は保持する。
     * 例: task.Category = "Sales"             → "(IDENT.Category=STR)"
     *     GetPriority(taskType) = "High"       → "(GetPriority(IDENT)=STR)"
     *     amount < 50000                       → "(IDENT<NUM)"
     *     Sheet1.Cells(row, 1).Value（代入RHS）→ "IDENT.Cells(IDENT,NUM).Value"
     */
    private computeConditionShape(expr: any): string {
        if (!expr) return '?';
        if (expr.type === 'Identifier')    return 'IDENT';
        if (expr.type === 'StringLiteral') return 'STR';
        if (expr.type === 'NumberLiteral') return 'NUM';
        if (expr.type === 'BinaryExpression') {
            const l = this.computeConditionShape(expr.left);
            const r = this.computeConditionShape(expr.right);
            return `(${l}${expr.operator ?? '?'}${r})`;
        }
        if (expr.type === 'CallExpression') {
            // callee は関数名として保持（Identifier → 名前をそのまま、MemberExpression → 再帰）
            const calleeShape = this.computeCalleeShape(expr.callee);
            const args = (expr.args ?? []).map((a: any) => this.computeConditionShape(a)).join(',');
            return `${calleeShape}(${args})`;
        }
        if (expr.type === 'MemberExpression') {
            const obj = this.computeConditionShape(expr.object);
            const prop = String(expr.property?.name ?? '?');
            return `${obj}.${prop}`;
        }
        return expr.type ?? '?';
    }

    /**
     * CallExpression の callee を形状文字列に変換する（構造要素なので名前を保持）
     * Identifier  → 名前をそのまま（例: "GetPriority"）
     * MemberExpression → object を正規化し property 名を保持（例: "IDENT.Cells"）
     */
    private computeCalleeShape(callee: any): string {
        if (!callee) return '?';
        if (callee.type === 'Identifier') return String(callee.name ?? '?');
        if (callee.type === 'MemberExpression') {
            const obj = this.computeConditionShape(callee.object);
            const prop = String(callee.property?.name ?? '?');
            return `${obj}.${prop}`;
        }
        return '?';
    }

    /** 数値を千区切りでフォーマット（例: 50000 → "50,000"） */
    private formatNumber(n: any): string {
        const num = Number(n);
        if (isNaN(num)) return String(n ?? '?');
        return num.toLocaleString('en-US');
    }

    /** ステートメント列から代入の右辺値を取得（例: "Manager"） */
    private extractResult(stmts: Statement[]): string {
        if (!Array.isArray(stmts)) return '?';
        for (const stmt of stmts) {
            const s = stmt as any;
            if (s?.type === 'AssignmentStatement' && s.right) {
                return this.formatExpr(s.right);
            }
        }
        return '?';
    }

    // -------------------------------------------------------------------------
    // 条件変数の抽出
    // -------------------------------------------------------------------------

    /**
     * if 文全体を再帰的に走査し、条件式に現れる識別子を全て収集する
     */
    private walkConditions(node: any, callback: (condition: any) => void): void {
        if (!node) return;
        if (Array.isArray(node)) {
            for (const item of node) this.walkConditions(item, callback);
            return;
        }
        if (node.type === 'IfStatement') {
            callback(node.condition);
            this.walkConditions(node.consequent, callback);
            this.walkConditions(node.alternate, callback);
        }
    }

    /**
     * if 文全体を再帰的に走査し、条件式とそのネストレベルをコールバックに渡す（T-02）
     * - ElseIf チェーン（alternate）: 同じレベルを維持
     * - consequent 内の内側 if: レベル +1
     */
    private walkConditionsWithLevel(
        ifStmt: any,
        level: number,
        callback: (condition: any, level: number) => void
    ): void {
        let current: any = ifStmt;
        while (current && current.type === 'IfStatement') {
            callback(current.condition, level);

            // consequent 内の内側 if を探してレベル +1 で再帰
            const stmts: any[] = Array.isArray(current.consequent) ? current.consequent : [];
            const innerIf = stmts.find((s: any) => s?.type === 'IfStatement');
            if (innerIf) {
                this.walkConditionsWithLevel(innerIf, level + 1, callback);
            }

            // alternate が ElseIf なら同レベルで継続
            current = current.alternate?.type === 'IfStatement' ? current.alternate : null;
        }
    }

    /**
     * 条件式の AST を再帰的に走査して Identifier 名を収集する
     * CallExpression の callee（関数名）はスキップし、引数のみ収集する
     * MemberExpression は object 側の識別子のみ収集する（property は変数名ではない）
     */
    private collectIdentifiersFromExpr(expr: any, result: Set<string>): void {
        if (!expr) return;
        if (expr.type === 'Identifier') {
            result.add(String(expr.name ?? '').toLowerCase());
            return;
        }
        if (expr.type === 'CallExpression') {
            // callee（関数名）はスキップ、引数の識別子だけを収集
            for (const arg of (expr.args ?? [])) {
                this.collectIdentifiersFromExpr(arg, result);
            }
            return;
        }
        if (expr.type === 'MemberExpression') {
            // object 側のみ走査（property は変数名ではなくプロパティ名）
            this.collectIdentifiersFromExpr(expr.object, result);
            return;
        }
        // BinaryExpression / UnaryExpression / その他: 既知の子ノードを走査
        for (const key of ['left', 'right', 'argument', 'callee']) {
            if (expr[key]) this.collectIdentifiersFromExpr(expr[key], result);
        }
    }

    /**
     * if 文の全条件からキー式を収集し、関数本体から宣言・代入情報を付与して返す
     * T-03 対応: 単純識別子だけでなく MemberExpression や CallExpression も取得する
     */
    private findConditionVariables(
        ifStmt: IfStatement,
        procBody: Statement[]
    ): LocalVariableInfo[] {
        // Step 1: 全条件からキー式とレベルを収集（T-02: レベル情報付き）
        const keyExprLevels = new Map<string, number>();
        this.walkConditionsWithLevel(ifStmt, 0, (cond, level) => {
            const keyExpr = this.extractKeyVariable(cond);
            if (keyExpr && !keyExprLevels.has(keyExpr)) {
                keyExprLevels.set(keyExpr, level);
            }
        });

        if (keyExprLevels.size === 0) return [];

        const ifStartLine = ifStmt.loc?.start.line ?? 0;
        const infoMap = new Map<string, LocalVariableInfo>();
        for (const [keyExpr, level] of keyExprLevels) {
            infoMap.set(keyExpr, { name: keyExpr, level, assignedLines: [] });
        }

        // Step 2: 関数本体を走査して Dim 宣言と代入を収集
        // キー式が単純識別子の場合のみ宣言/代入を照合できる
        for (const stmt of procBody) {
            const s = stmt as any;

            // Dim 宣言
            if (s.type === 'VariableDeclaration') {
                for (const decl of (s.declarations ?? [])) {
                    const declName = String(decl.name?.name ?? '').toLowerCase();
                    for (const [keyExpr, info] of infoMap) {
                        if (keyExpr.toLowerCase() === declName) {
                            info.declaredAt = s.loc?.start.line ?? undefined;
                        }
                    }
                }
            }

            // 代入文（if チェーン開始より前のもの）
            if (s.type === 'AssignmentStatement') {
                const stmtLine: number = s.loc?.start.line ?? 0;
                if (stmtLine < ifStartLine) {
                    const target = s.left as any;
                    // 単純識別子の代入: `amount = ...`
                    if (target?.type === 'Identifier') {
                        const targetName = String(target.name ?? '').toLowerCase();
                        for (const [keyExpr, info] of infoMap) {
                            if (keyExpr.toLowerCase() === targetName) {
                                info.assignedLines.push(stmtLine);
                                // T-01: RHS の形状を収集
                                if (s.right) {
                                    const shape = this.computeConditionShape(s.right);
                                    if (!info.assignmentAstShapes) info.assignmentAstShapes = [];
                                    info.assignmentAstShapes.push(shape);
                                }
                            }
                        }
                    }
                    // メンバーアクセスの代入: `task.Category = ...`
                    if (target?.type === 'MemberExpression') {
                        const memberStr = this.formatExprFull(target);
                        if (infoMap.has(memberStr)) {
                            const info = infoMap.get(memberStr)!;
                            info.assignedLines.push(stmtLine);
                            // T-01: RHS の形状を収集
                            if (s.right) {
                                const shape = this.computeConditionShape(s.right);
                                if (!info.assignmentAstShapes) info.assignmentAstShapes = [];
                                info.assignmentAstShapes.push(shape);
                            }
                        }
                    }
                }
            }
        }

        // T-01: hasUniformAssignment を計算
        for (const info of infoMap.values()) {
            const shapes = info.assignmentAstShapes;
            if (shapes && shapes.length > 0) {
                info.hasUniformAssignment = shapes.every(s => s === shapes[0]);
            }
        }

        // ソート: 宣言行順（引数など宣言なしは末尾）
        return Array.from(infoMap.values()).sort((a, b) =>
            (a.declaredAt ?? 999999) - (b.declaredAt ?? 999999)
        );
    }

    // -------------------------------------------------------------------------
    // 副作用の均一性分析
    // -------------------------------------------------------------------------

    /**
     * 全ブランチの副作用文の構造を比較し、均一かどうかを判定する
     */
    private analyzeSideEffects(outerChain: IfStatement[], funcName: string): SideEffectAnalysis {
        const leafBlocks = this.collectLeafStatements(outerChain, []);
        if (leafBlocks.length === 0) {
            return { hasUniformStructure: true, sideEffectCount: 0, structureSummary: 'none' };
        }

        const shapeGroups = leafBlocks.map((lb) => this.getSideEffectShape(lb.stmts, funcName));
        const firstShape = JSON.stringify(shapeGroups[0]);
        const hasUniformStructure = shapeGroups.every((s) => JSON.stringify(s) === firstShape);
        const sideEffectCount = shapeGroups[0]?.length ?? 0;

        return {
            hasUniformStructure,
            sideEffectCount,
            structureSummary: (shapeGroups[0] ?? []).join(', ') || 'none',
            mismatchedBranches: hasUniformStructure
                ? undefined
                : leafBlocks
                      .filter((lb, i) => JSON.stringify(shapeGroups[i]) !== firstShape)
                      .map((lb) => lb.keyPath.join('|')),
        };
    }

    /**
     * ネストされた if チェーンから末端（葉）の statement ブロックを再帰的に収集する
     */
    private collectLeafStatements(
        chain: IfStatement[],
        keyPath: string[]
    ): Array<{ keyPath: string[]; stmts: Statement[] }> {
        const result: Array<{ keyPath: string[]; stmts: Statement[] }> = [];

        for (const branch of chain) {
            const keyValue = this.extractKeyValue(branch.condition as any);
            const currentPath = [...keyPath, keyValue];
            const stmts = branch.consequent;
            const innerIf = stmts.find((s) => s?.type === 'IfStatement') as IfStatement | undefined;

            if (!innerIf) {
                result.push({ keyPath: currentPath, stmts });
                continue;
            }

            // 外側ブランチの非 IfStatement 文（Set info.xxx = ... など）を保持する
            const outerNonIfStmts = stmts.filter((s: any) => s?.type !== 'IfStatement');

            if (this.isNumericThresholdCondition(innerIf.condition as any)) {
                // しきい値レベル: 各 branch の consequent と最後の else を収集
                // 外側非 if 文を各葉にマージして副作用を引き継ぐ
                let current: any = innerIf;
                while (current && current.type === 'IfStatement') {
                    result.push({ keyPath: currentPath, stmts: [...outerNonIfStmts, ...current.consequent as Statement[]] });
                    if (Array.isArray(current.alternate)) {
                        result.push({ keyPath: currentPath, stmts: [...outerNonIfStmts, ...current.alternate as Statement[]] });
                        break;
                    }
                    current = current.alternate ?? null;
                }
            } else {
                // キーレベル（文字列等値・関数呼び出し・その他）: 再帰
                // 外側非 if 文を各再帰葉にマージして副作用を引き継ぐ
                const recursed = this.collectLeafStatements(this.collectIfElseChain(innerIf), currentPath);
                for (const leaf of recursed) {
                    result.push({ keyPath: leaf.keyPath, stmts: [...outerNonIfStmts, ...leaf.stmts] });
                }
            }
        }

        return result;
    }

    /**
     * 葉ブロックの副作用文を AST 形状文字列の配列として返す
     * （funcName に代入する文は結果代入なので除外する）
     */
    private getSideEffectShape(stmts: Statement[], funcName: string): string[] {
        return stmts
            .filter((s) => {
                const stmt = s as any;
                // AssignmentStatement と SetStatement の両方で結果代入を除外する
                if (stmt?.type !== 'AssignmentStatement' && stmt?.type !== 'SetStatement') return true;
                const target = this.formatExprFull(stmt.left);
                return target.toLowerCase() !== funcName.toLowerCase();
            })
            .map((s) => this.computeStmtShape(s));
    }

    /** 文の AST 形状を文字列で返す（リテラル値を型プレースホルダーに置換） */
    private computeStmtShape(stmt: any): string {
        if (!stmt) return '?';
        // AssignmentStatement と SetStatement を同様に扱う
        // RHS 呼び出しはgetter（副作用なし）として computeConditionShape で形状を正規化する
        // 制限: RHS 関数が副作用を持つ場合は検出できない（T-07 参照）
        if (stmt.type === 'AssignmentStatement' || stmt.type === 'SetStatement') {
            const target = this.formatExprFull(stmt.left);
            const right = this.computeConditionShape(stmt.right);
            return `${target}=(${right})`;
        }
        return stmt.type ?? '?';
    }

    /** 式の AST 形状を文字列で返す（リテラルは STR/NUM に置換） */
    private computeExprShape(expr: any): string {
        if (!expr) return '?';
        switch (expr.type) {
            case 'StringLiteral': return 'STR';
            case 'NumberLiteral': return 'NUM';
            case 'Identifier': return String(expr.name ?? '?');
            case 'BinaryExpression': {
                const l = this.computeExprShape(expr.left);
                const op = expr.operator ?? '?';
                const r = this.computeExprShape(expr.right);
                return `${l}${op}${r}`;
            }
            default: return expr.type ?? '?';
        }
    }

    // -------------------------------------------------------------------------
    // デシジョンテーブル ASCII レンダリング
    // -------------------------------------------------------------------------

    /**
     * DecisionTableRow を Y/N/- 形式の ASCII テーブルとして描画する
     *
     * 例:
     *   renderRow(salesRow) →
     *   ┌──────────────────────┬─────────┬──────────┬────┬─────┐
     *   │ 条件                 │   R1    │    R2    │ R3 │  R4 │
     *   ├──────────────────────┼─────────┼──────────┼────┼─────┤
     *   │ amount < 50,000      │    Y    │    N     │  N │   N │
     *   │ amount < 500,000     │    -    │    Y     │  N │   N │
     *   │ amount < 2,000,000   │    -    │    -     │  Y │   N │
     *   ├──────────────────────┼─────────┼──────────┼────┼─────┤
     *   │ Approver（結果）     │ Manager │ Director │ VP │ CFO │
     *   └──────────────────────┴─────────┴──────────┴────┴─────┘
     */
    renderRow(row: DecisionTableRow): string {
        const n = row.rules.length;               // 条件数
        const ruleCount = n + 1;                  // ルール列数（条件数 + else）

        // 各列の結果値
        const results = [...row.rules.map((r) => r.result), row.defaultResult];

        // 列幅の計算（ヘッダー "R1".."Rn" と結果値の最大幅）
        const ruleColWidths = results.map((v, i) => Math.max(2, `R${i + 1}`.length, v.length));

        // 条件列の幅
        const condColWidth = Math.max(
            '条件'.length * 2,                   // 全角2文字 = 半角4
            ...row.rules.map((r) => r.condition.length),
            '結果'.length * 2 + 4,
        );

        const pad = (s: string, w: number) => s.padEnd(w);
        const center = (s: string, w: number) => {
            const space = Math.max(0, w - s.length);
            const l = Math.floor(space / 2);
            return ' '.repeat(l) + s + ' '.repeat(space - l);
        };

        const colSep = '┼';
        const hRule = (l: string, m: string, r: string) =>
            l + '─'.repeat(condColWidth + 2) + m +
            ruleColWidths.map((w) => '─'.repeat(w + 2)).join(m) + r;

        const dataRow = (label: string, cells: string[]) =>
            '│ ' + pad(label, condColWidth) + ' │' +
            cells.map((c, i) => ' ' + center(c, ruleColWidths[i]) + ' │').join('');

        const lines: string[] = [];
        lines.push(hRule('┌', '┬', '┐'));

        // ヘッダー行
        const headers = Array.from({ length: ruleCount }, (_, i) => `R${i + 1}`);
        lines.push(dataRow('条件', headers));
        lines.push(hRule('├', colSep, '┤'));

        // 条件行（Y/N/- マトリクス）
        for (let ci = 0; ci < n; ci++) {
            const cells = Array.from({ length: ruleCount }, (_, ri) => {
                if (ri < ci)  return '-';   // 前条件でマッチ済み → don't care
                if (ri === ci) return 'Y';  // このルールがマッチする条件
                return 'N';                 // まだマッチしていない
            });
            lines.push(dataRow(row.rules[ci].condition, cells));
        }

        lines.push(hRule('├', colSep, '┤'));

        // 結果行
        lines.push(dataRow('結果', results));
        lines.push(hRule('└', '┴', '┘'));

        return `${row.key}（${ruleCount}ルール）:\n` + lines.join('\n');
    }

    /**
     * DecisionTable 全体を描画する
     * 多段キー（keyPath.length > 1）の場合は第1キーでグループ化して表示する
     */
    renderDecisionTable(table: DecisionTable): string {
        if (table.rows.length === 0) return '';

        const isMultiLevel = table.rows.some((r) => (r.keyPath?.length ?? 1) > 1);

        if (!isMultiLevel) {
            return table.rows.map((row) => this.renderRow(row)).join('\n\n');
        }

        // 第1キーでグループ化
        const groups = new Map<string, DecisionTableRow[]>();
        for (const row of table.rows) {
            const groupKey = row.keyPath?.[0] ?? row.key;
            if (!groups.has(groupKey)) groups.set(groupKey, []);
            groups.get(groupKey)!.push(row);
        }

        const sections: string[] = [];
        for (const [groupKey, groupRows] of groups) {
            const lines: string[] = [`${groupKey}:`];
            for (const row of groupRows) {
                // 第1キーを除いたサブパスを表示キーに
                const subKey = row.keyPath ? row.keyPath.slice(1).join('|') : row.key;
                const displayRow = { ...row, key: subKey };
                const rendered = this.renderRow(displayRow)
                    .split('\n')
                    .map((l) => '  ' + l)
                    .join('\n');
                lines.push(rendered);
            }
            sections.push(lines.join('\n'));
        }

        return sections.join('\n\n');
    }
}

// CLI code moved to separate script to avoid execution during module import
// See: table-driven-detector-cli.ts
