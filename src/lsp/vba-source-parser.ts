import { Lexer } from '../engine/lexer';
import { Parser } from '../engine/parser';
import { stripVBAFileHeader } from '../engine/preprocessor';

export interface VBAParseContext {
    moduleName: string;
    isClass?: boolean;
    errorRecovery?: boolean;
}

/**
 * Parse a complete VBA module through the same VBE-header and class-module
 * normalization used by LSP diagnostics.  Secondary entry points must use
 * this helper instead of sending raw .cls text to Lexer/Parser.
 */
export function parseVBAModule(source: string, context: VBAParseContext): ReturnType<Parser['parse']> {
    const tokens = new Lexer(stripVBAFileHeader(source)).tokenize();
    const options: {
        parseAsClass?: string;
        errorRecovery?: boolean;
    } = { errorRecovery: context.errorRecovery ?? true };
    if (context.isClass) options.parseAsClass = context.moduleName;
    return new Parser(tokens, options).parse();
}
