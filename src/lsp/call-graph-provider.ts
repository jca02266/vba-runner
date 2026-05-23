import { Statement, CallExpression, CallStatement, Identifier, ProcedureDeclaration, TypeDeclaration, ClassDeclaration } from '../engine/parser';

export interface ProcNode {
    name: string;
    uri: string;
    line: number;
    scope: 'public' | 'private' | 'friend';
    referenceCount: number;
    isExcelDependent: boolean;
}

export interface CallEdge {
    from: string;  // lowercase
    to: string;    // lowercase
}

export interface CallGraph {
    nodes: Map<string, ProcNode>;  // key = name.toLowerCase()
    edges: CallEdge[];
}

const EXCEL_ROOT_OBJECTS = new Set([
    'sheets', 'range', 'cells', 'application',
    'activesheet', 'activeworkbook', 'activecell',
    'thisworkbook', 'workbook', 'workbooks',
    'worksheet', 'worksheets', 'columns', 'rows',
    'selection', 'workbookbeforesave',
]);

export class CallGraphProvider {
    buildCallGraph(fileMap: Map<string, { statements: Statement[], uri: string }>): CallGraph {
        const nodes = new Map<string, ProcNode>();
        const edges: CallEdge[] = [];
        const procsByFile = new Map<string, Map<string, ProcNode>>();  // uri -> procNameLower -> ProcNode

        // Phase 1: Collect all ProcedureDeclarations
        for (const [uri, { statements }] of fileMap) {
            const fileProcs = new Map<string, ProcNode>();
            this.collectProcedures(statements, uri, fileProcs);
            for (const proc of fileProcs.values()) {
                nodes.set(proc.name.toLowerCase(), proc);
            }
            procsByFile.set(uri, fileProcs);
        }

        // Phase 2 & 3: Collect calls and build edges
        for (const [_uri, { statements }] of fileMap) {
            const typeDefs = new Set<string>();
            this.collectTypeNames(statements, typeDefs);
            const procs = this.collectProceduresByUri(statements);

            for (const proc of procs) {
                const calls = this.collectCalls(proc.body || [], typeDefs);
                for (const calleeName of calls) {
                    const calleeKey = calleeName.toLowerCase();
                    if (nodes.has(calleeKey)) {
                        edges.push({
                            from: proc.name.toLowerCase(),
                            to: calleeKey,
                        });
                    }
                }
            }
        }

        // Phase 4: Calculate reference counts
        const refCount = new Map<string, number>();
        for (const edge of edges) {
            refCount.set(edge.to, (refCount.get(edge.to) ?? 0) + 1);
        }
        for (const [key, proc] of nodes) {
            proc.referenceCount = refCount.get(key) ?? 0;
        }

        // Phase 5: Detect Excel dependencies
        for (const { statements } of fileMap.values()) {
            const procs = this.collectProceduresByUri(statements);
            for (const proc of procs) {
                proc.isExcelDependent = this.isExcelDependentBody(proc.body || []);
                const procKey = proc.name.toLowerCase();
                const node = nodes.get(procKey);
                if (node) {
                    node.isExcelDependent = proc.isExcelDependent;
                }
            }
        }

        return { nodes, edges };
    }

    private collectProcedures(statements: Statement[], _uri: string, procs: Map<string, ProcNode>) {
        for (const stmt of statements) {
            if (stmt.type === 'ProcedureDeclaration') {
                const procDecl = stmt as ProcedureDeclaration;
                const scope = procDecl.scope || 'private';
                const line = procDecl.loc?.start.line ?? 0;
                const node: ProcNode = {
                    name: procDecl.name?.name || '',
                    uri: _uri,
                    line,
                    scope: scope as 'public' | 'private' | 'friend',
                    referenceCount: 0,
                    isExcelDependent: false,
                };
                procs.set(node.name.toLowerCase(), node);
            }
        }
    }

    private collectTypeNames(statements: Statement[], typeDefs: Set<string>) {
        for (const stmt of statements) {
            if (stmt.type === 'TypeDeclaration') {
                const typeDecl = stmt as TypeDeclaration;
                if (typeDecl.name) {
                    typeDefs.add(typeDecl.name.toLowerCase());
                }
            } else if (stmt.type === 'ClassDeclaration') {
                const classDecl = stmt as ClassDeclaration;
                if (classDecl.name) {
                    typeDefs.add(classDecl.name.toLowerCase());
                }
            }
        }
    }

    private collectProceduresByUri(statements: Statement[]): Array<{ name: string, body: Statement[], isExcelDependent?: boolean }> {
        const procs: Array<{ name: string, body: Statement[], isExcelDependent?: boolean }> = [];
        for (const stmt of statements) {
            if (stmt.type === 'ProcedureDeclaration') {
                const procDecl = stmt as ProcedureDeclaration;
                procs.push({
                    name: procDecl.name?.name || '',
                    body: procDecl.body || [],
                    isExcelDependent: false,
                });
            }
        }
        return procs;
    }

    private collectCalls(body: Statement[], typeDefs: Set<string>): Set<string> {
        const calls = new Set<string>();
        this.walkStatements(body, (node) => {
            if (node.type === 'CallExpression') {
                const callExpr = node as CallExpression;
                if (callExpr.callee?.type === 'Identifier') {
                    const name = (callExpr.callee as Identifier).name;
                    if (name && !typeDefs.has(name.toLowerCase())) {
                        calls.add(name);
                    }
                } else if ((callExpr.callee as any)?.name && typeof (callExpr.callee as any).name === 'string') {
                    const name = (callExpr.callee as any).name;
                    if (!typeDefs.has(name.toLowerCase())) {
                        calls.add(name);
                    }
                }
            } else if (node.type === 'CallStatement') {
                const callStmt = node as CallStatement;
                const expr = callStmt.expression as any;
                if (expr?.callee?.type === 'Identifier') {
                    const name = (expr.callee as Identifier).name;
                    if (name && !typeDefs.has(name.toLowerCase())) {
                        calls.add(name);
                    }
                }
            }
        });
        return calls;
    }

    private isExcelDependentBody(body: Statement[]): boolean {
        let found = false;
        this.walkStatements(body, (node) => {
            if (found) return;
            if (node.type === 'Identifier') {
                const name = (node as Identifier).name;
                if (name && EXCEL_ROOT_OBJECTS.has(name.toLowerCase())) {
                    found = true;
                }
            }
        });
        return found;
    }

    private walkStatements(statements: any[], callback: (node: any) => void): void {
        for (const stmt of statements) {
            this.walk(stmt, callback);
        }
    }

    private walk(node: any, callback: (node: any) => void): void {
        if (!node || typeof node !== 'object') {
            return;
        }

        callback(node);

        if (Array.isArray(node)) {
            for (const item of node) {
                this.walk(item, callback);
            }
        } else if (node.type) {
            for (const key of Object.keys(node)) {
                if (key === 'loc' || key === 'type') continue;
                const value = node[key];
                if (Array.isArray(value)) {
                    for (const item of value) {
                        this.walk(item, callback);
                    }
                } else if (value && typeof value === 'object') {
                    this.walk(value, callback);
                }
            }
        }
    }
}
