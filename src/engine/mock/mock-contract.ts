/** JavaScript モックの名前空間契約。実行時ローダーとLSP静的解析で共有する。 */
export const VBA_MOCK_NAMESPACES = ['constants', 'objects', 'procedures', 'comObjects'] as const;
export type VbaMockNamespace = typeof VBA_MOCK_NAMESPACES[number];

/** LSPが静的にキーを収集する名前空間。comObjectsのProgIDはVBA名ではない。 */
export const VBA_MOCK_IDENTIFIER_NAMESPACES = ['constants', 'objects', 'procedures'] as const;

export function isVbaMockNamespace(name: string): name is VbaMockNamespace {
    return (VBA_MOCK_NAMESPACES as readonly string[]).includes(name);
}

export function isVbaMockConstantValue(value: unknown): value is string | number | boolean {
    return typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean';
}
