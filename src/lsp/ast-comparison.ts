import { Statement, Expression } from '../engine/parser';

/**
 * Compare two AST nodes for semantic equivalence (ignoring location info).
 * Returns true if both nodes represent the same expression structure and values.
 */
export function astEqual(node1: any, node2: any): boolean {
    if (node1 === node2) return true;
    if (!node1 || !node2) return node1 === node2;
    if (typeof node1 !== 'object' || typeof node2 !== 'object') return node1 === node2;

    // Compare structure by JSON representation, excluding 'loc'
    const key1 = serializeAst(node1);
    const key2 = serializeAst(node2);
    return key1 === key2;
}

/**
 * Serialize AST node to a canonical JSON string, excluding location info (recursively).
 */
export function serializeAst(node: any): string {
    const stripLoc = (n: any): any => {
        if (!n || typeof n !== 'object') return n;
        if (Array.isArray(n)) return n.map(stripLoc);
        const { loc, ...rest } = n;
        const result: any = {};
        for (const key of Object.keys(rest)) {
            result[key] = stripLoc(rest[key]);
        }
        return result;
    };
    return JSON.stringify(stripLoc(node));
}

/**
 * Recursively walk AST and find all expressions that match a given expression.
 * Returns array of matching nodes with their location info.
 */
export function findMatchingExpressions(
    body: any,
    targetExpr: any
): Array<{ node: any; start: { line: number; column: number }; end: { line: number; column: number } }> {
    const matches: Array<{ node: any; start: { line: number; column: number }; end: { line: number; column: number } }> = [];
    const targetKey = serializeAst(targetExpr);

    const walk = (node: any) => {
        if (!node || typeof node !== 'object') return;

        // Check if this node is an expression
        const isExpression = node.type && (
            node.type.includes('Expression') ||
            node.type === 'BinaryOp' ||
            node.type === 'Identifier' ||
            node.type === 'Literal' ||
            node.type === 'UnaryOp' ||
            node.type === 'MemberExpression' ||
            node.type === 'CallExpression'
        );

        if (isExpression && node.loc) {
            if (serializeAst(node) === targetKey) {
                const { start, end } = node.loc;
                matches.push({
                    node,
                    start: { line: start.line - 1, column: start.column - 1 },
                    end: { line: end.line - 1, column: end.column - 1 }
                });
            }
        }

        // Recurse into child nodes
        if (Array.isArray(node)) {
            for (const item of node) walk(item);
        } else {
            for (const key of Object.keys(node)) {
                if (key !== 'loc') {
                    walk(node[key]);
                }
            }
        }
    };

    walk(body);
    return matches;
}
