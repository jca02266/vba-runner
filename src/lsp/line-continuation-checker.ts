/**
 * VBA の行継続文字 _ が必要かを判定するユーティリティ。
 * extension.ts の OnTypeFormattingProvider から利用する。
 */

const CONTINUATION_CHARS = /[+\-*\/\\^&,=(]$|[<>]$|<=$|>=$|<>$/;
const CONTINUATION_KEYWORDS = /\b(And|Or|Xor|Eqv|Imp|Mod|Like|Is)\s*$/i;

/**
 * 前の行の末尾テキスト（trailing whitespace を除去済み）を受け取り、
 * 行継続文字 _ を挿入すべきかを返す。
 */
export function needsLineContinuation(trimmedLine: string): boolean {
    if (trimmedLine.length === 0) return false;
    if (trimmedLine.endsWith('_')) return false;          // already continued
    if (/^\s*'/.test(trimmedLine)) return false;          // comment line
    if (CONTINUATION_CHARS.test(trimmedLine)) return true;
    if (CONTINUATION_KEYWORDS.test(trimmedLine)) return true;
    return false;
}
