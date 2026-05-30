/**
 * VBA の行継続文字 _ が必要かを判定するユーティリティ。
 * extension.ts の OnTypeFormattingProvider から利用する。
 */

const CONTINUATION_CHARS = /[+\-*\/\\^&,=(]$|[<>]$|<=$|>=$|<>$/;
const CONTINUATION_KEYWORDS = /\b(And|Or|Xor|Eqv|Imp|Mod|Like|Is)\s*$/i;

/**
 * インラインコメント（最初の unquoted '）より後ろを除去し、コード部分だけを返す。
 * コメントが存在する場合は { code, hasComment: true } を返す。
 */
export function stripInlineComment(line: string): { code: string; hasComment: boolean } {
    let inString = false;
    for (let i = 0; i < line.length; i++) {
        if (line[i] === '"') {
            // VBA の "" はエスケープされた引用符
            if (i + 1 < line.length && line[i + 1] === '"') { i++; continue; }
            inString = !inString;
        } else if (line[i] === "'" && !inString) {
            return { code: line.slice(0, i).trimEnd(), hasComment: true };
        }
    }
    return { code: line, hasComment: false };
}

/**
 * 前の行の末尾テキスト（trailing whitespace を除去済み）を受け取り、
 * 行継続文字 _ を挿入すべきかを返す。
 *
 * インラインコメントがある場合（例: Array(1,  ' note）は、
 * コメント末尾に _ を追加しても行継続として機能しないため false を返す。
 */
export function needsLineContinuation(trimmedLine: string): boolean {
    if (trimmedLine.length === 0) return false;
    if (trimmedLine.endsWith('_')) return false;          // already continued
    if (/^\s*'/.test(trimmedLine)) return false;          // comment-only line

    const { code, hasComment } = stripInlineComment(trimmedLine);

    // インラインコメントがある行は _ を挿入しない。
    // コード部分が行継続可能であっても、_ をコメント末尾に付けても無効であり、
    // _ をコメント前（' の直前）に置くと "_ + ' comment" の無効パターンになる。
    if (hasComment) return false;

    if (CONTINUATION_CHARS.test(code)) return true;
    if (CONTINUATION_KEYWORDS.test(code)) return true;
    return false;
}
