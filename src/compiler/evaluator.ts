import {
    Program,
    Statement,
    ForStatement,
    IfStatement,
    DoWhileStatement,
    WhileStatement,
    Expression,
    UnaryExpression,
    BinaryExpression,
    Identifier,
    NumberLiteral,
    StringLiteral,
    DateLiteral,
    AssignmentStatement,
    ProcedureDeclaration,
    VariableDeclaration,
    CallStatement,
    CallExpression,
    MemberExpression,
    ConstDeclaration,
    SetStatement,
    EraseStatement,
    ReDimStatement,
    ExitStatement,
    OnErrorStatement,
    TypeDeclaration,
    TypeMember,
    SelectCaseStatement,
    ForEachStatement,
    WithStatement,
    ImplicitWithObjectExpression,
    GoToStatement,
    StopStatement,
    EndStatement,
    GoSubStatement,
    ReturnStatement,
    OnGoToSubStatement,
    LSetStatement,
    RSetStatement,
    ErrorStatement,
    ClassDeclaration,
    NewExpression,
    Parser,
} from './parser';
import { Lexer, TokenType } from './lexer';

export class VbaBoolean {
    constructor(public value: -1 | 0) {}
    valueOf() { return this.value; }
    toString() { return this.value === -1 ? 'True' : 'False'; }
}

export class VbaDate {
    constructor(public value: number) {}
    valueOf() { return this.value; }
    toString() {
        // Simple conversion for now
        const d = new Date((this.value - 25569) * 86400000);
        return d.toLocaleDateString();
    }
}

export class VbaErrorValue {
    constructor(public code: number) {}
    valueOf() { return this.code; }
    toString() { return `Error ${this.code}`; }
}

export const vbaEmpty = null;

class VbaCollection {
    private _items: { value: any, key: string | null }[] = [];
    public readonly __isVbaCollection__ = true;

    public add(item: any, key?: string, before?: any, after?: any) {
        const keyLower = (key !== undefined && key !== vbaEmpty && key !== null) ? String(key).toLowerCase() : null;
        if (keyLower && this._items.some(i => i.key === keyLower)) {
            throw new Error("This key is already associated with an element of this collection");
        }
        
        const newItem = { value: item, key: keyLower };

        if (before !== undefined && before !== vbaEmpty && before !== null) {
            const idx = this.findIndex(before) - 1;
            this._items.splice(idx, 0, newItem);
        } else if (after !== undefined && after !== vbaEmpty && after !== null) {
            const idx = this.findIndex(after);
            this._items.splice(idx, 0, newItem);
        } else {
            this._items.push(newItem);
        }
    }

    private findIndex(id: any): number {
        if (typeof id === 'number') {
            if (id < 1 || id > this._items.length) throw new Error("Subscript out of range");
            return id;
        } else if (id !== undefined && id !== null && id !== vbaEmpty) {
            const k = String(id).toLowerCase();
            const idx = this._items.findIndex(i => i.key === k);
            if (idx === -1) throw new Error("Invalid procedure call or argument");
            return idx + 1;
        }
        throw new Error("Invalid procedure call or argument");
    }

    public count() {
        return this._items.length;
    }

    public item(id: any) {
        const idx = this.findIndex(id);
        return this._items[idx - 1].value;
    }

    public remove(id: any) {
        const idx = this.findIndex(id);
        this._items.splice(idx - 1, 1);
    }

    public get items(): any[] {
        return this._items.map(i => i.value);
    }
}
export const vbaNull = Symbol('vbaNull');
export const vbaNothing = Symbol('vbaNothing');
export const vbaMissing = Symbol('vbaMissing');
export const vbaTrue = new VbaBoolean(-1);
export const vbaFalse = new VbaBoolean(0);
export type VbaBooleanType = VbaBoolean;

export class Environment {
    private variables: Map<string, any> = new Map();
    private procedures: Map<string, ProcedureDeclaration> = new Map();
    private types: Map<string, TypeMember[]> = new Map();
    public enclosing?: Environment;

    constructor(enclosing?: Environment) {
        this.enclosing = enclosing;
    }

    set(name: string, value: any) {
        const key = name.toLowerCase();
        if (this.variables.has(key)) {
            this.variables.set(key, value);
            return;
        }
        let env: Environment | undefined = this.enclosing;
        while (env) {
            if (env.variables.has(key)) {
                env.variables.set(key, value);
                return;
            }
            env = env.enclosing;
        }
        this.variables.set(key, value);
    }

    setLocally(name: string, value: any) {
        this.variables.set(name.toLowerCase(), value);
    }

    get(name: string): any {
        const key = name.toLowerCase();
        if (this.variables.has(key)) {
            return this.variables.get(key);
        }
        let env: Environment | undefined = this.enclosing;
        while (env) {
            if (env.variables.has(key)) {
                return env.variables.get(key);
            }
            env = env.enclosing;
        }

        // Implicit initialization
        this.variables.set(key, 0);
        return 0;
    }

    setProcedure(name: string, proc: ProcedureDeclaration) {
        const key = name.toLowerCase();
        if (proc.isProperty && proc.propertyType) {
            this.procedures.set(`${key}:${proc.propertyType}`, proc);
        } else {
            this.procedures.set(key, proc);
        }
    }

    getProcedure(name: string, type?: 'get' | 'let' | 'set'): ProcedureDeclaration | undefined {
        const baseKey = name.toLowerCase();
        // If type is not specified, try baseKey then baseKey:get
        const keysToTry = type ? [`${baseKey}:${type}`] : [baseKey, `${baseKey}:get`];

        for (const key of keysToTry) {
            if (this.procedures.has(key)) {
                return this.procedures.get(key);
            }
        }

        let env: Environment | undefined = this.enclosing;
        while (env) {
            for (const key of keysToTry) {
                if (env.procedures.has(key)) {
                    return env.procedures.get(key);
                }
            }
            env = env.enclosing;
        }
        return undefined;
    }

    setType(name: string, members: TypeMember[]) {
        this.types.set(name.toLowerCase(), members);
    }

    getType(name: string): TypeMember[] | undefined {
        const key = name.toLowerCase();
        if (this.types.has(key)) {
            return this.types.get(key);
        }
        let env: Environment | undefined = this.enclosing;
        while (env) {
            if (env.types.has(key)) {
                return env.types.get(key);
            }
            env = env.enclosing;
        }
        return undefined;
    }
}

export type PrintCallback = (output: string) => void;

export class Evaluator {
    private env: Environment;
    private onPrint: PrintCallback;
    private errorHandlerLabel: string | null = null;
    private currentProcBody: Statement[] | null = null;
    private currentProcedureName: string | null = null;
    private currentProcedureType: string | null = null;
    private currentSourceModule: string = '';
    private executingModuleName: string = '';
    private withObjectStack: any[] = [];
    private gosubStack: number[] = []; // Stack of statement indices for GoSub/Return
    private staticVarStore: Map<string, any> = new Map(); // persistent store for Static variables
    private currentProcIsStatic: boolean = false;
    private arrayBase: number = 0;
    private staticVarsInCurrentProc: Set<string> = new Set();
    private classDefinitions: Map<string, ClassDeclaration> = new Map();
    private comparisonMode: 'Binary' | 'Text' = 'Binary';
    private errorHandlingMode: 'None' | 'Label' | 'ResumeNext' = 'None';
    private isInErrorHandler: boolean = false;
    private lastErrorIndex: number = -1;

    constructor(onPrint: PrintCallback) {
        this.env = new Environment();
        this.onPrint = onPrint;
        // Add built-in debug object
        this.env.set('debug', {
            print: (...args: any[]) => this.onPrint(args.join(' ')),
            assert: (condition: any) => {
                if (!this.isTrue(condition)) {
                    throw new Error('Execution error: Assertion failed');
                }
            }
        });

        // VBA date serial: days since 1899-12-30 (VBA epoch)
        const VBA_EPOCH = new Date(1899, 11, 30); // local time
        const MS_PER_DAY = 86400000;

        const toVbaDate = (d: Date): number =>
            (d.getTime() - VBA_EPOCH.getTime()) / MS_PER_DAY;

        const fromVbaDate = (serial: number): Date => {
            const ms = Math.round(serial * MS_PER_DAY);
            return new Date(VBA_EPOCH.getTime() + ms);
        };

        const parseVbaDate = (val: any): Date => {
            if (val === null || val === undefined) throw new Error('Execution error: Invalid date');
            if (typeof val === 'number') return fromVbaDate(val);
            const d = new Date(val);
            if (isNaN(d.getTime())) throw new Error(`Execution error: Type mismatch: '${val}'`);
            return d;
        };

        // Add typical VBA built-ins
        this.env.set('isempty', (val: any) => (val === undefined || val === null || val === '') ? vbaTrue : vbaFalse);
        this.env.set('ismissing', (val: any) => val === vbaMissing ? vbaTrue : vbaFalse);
        this.env.set('isnumeric', (val: any) => (!isNaN(parseFloat(val)) && isFinite(val)) ? vbaTrue : vbaFalse);
        this.env.set('cdbl', (val: any) => parseFloat(val) || 0);
        this.env.set('csng', (val: any) => {
            let num: number;
            if (val === vbaTrue) num = -1;
            else if (val === vbaFalse) num = 0;
            else num = parseFloat(val) || 0;
            const f32 = Math.fround(num);
            if (!isFinite(f32) && isFinite(num)) {
                this.throwVbaError(6, "Overflow");
            }
            return f32;
        });
        this.env.set('cdate', (val: any) => {
            if (val === null || val === vbaNull || val === vbaEmpty) throw new Error('Execution error: Type mismatch');
            if (val instanceof VbaDate) return val;
            if (typeof val === 'number') return new VbaDate(val);
            const d = new Date(val);
            if (isNaN(d.getTime())) throw new Error(`Execution error: Type mismatch: '${val}'`);
            return new VbaDate(toVbaDate(d));
        });
        this.env.set('cvdate', (val: any) => {
            if (val === vbaNull) return vbaNull;
            if (val === null || val === vbaEmpty) throw new Error('Execution error: Type mismatch');
            if (val instanceof VbaDate) return val;
            if (typeof val === 'number') return new VbaDate(val);
            const d = new Date(val);
            if (isNaN(d.getTime())) throw new Error(`Execution error: Type mismatch: '${val}'`);
            return new VbaDate(toVbaDate(d));
        });
        this.env.set('clng', (val: any) => {
            const n = this.vbaRound(parseFloat(val) || 0);
            if (n < -2147483648 || n > 2147483647) this.throwVbaError(6, "Overflow");
            return n;
        });
        this.env.set('ccur', (val: any) => {
            let num: number;
            if (val === vbaTrue) num = -1;
            else if (val === vbaFalse) num = 0;
            else num = parseFloat(val) || 0;
            
            const n = this.vbaRound(num, 4);
            // Currency range: -922,337,203,685,477.5808 to 922,337,203,685,477.5807
            if (n < -922337203685477.5808 || n > 922337203685477.5807) {
                this.throwVbaError(6, "Overflow");
            }
            return n;
        });
        this.env.set('cbyte', (val: any) => {
            let num: number;
            if (val instanceof VbaBoolean) {
                num = val.value === -1 ? 255 : 0;
            } else if (typeof val === 'string') {
                num = parseFloat(val);
                if (isNaN(num)) this.throwVbaError(13, "Type mismatch");
                num = this.vbaRound(num);
            } else if (typeof val === 'number') {
                num = this.vbaRound(val);
            } else {
                num = this.vbaRound(Number(val));
            }
            if (num < 0 || num > 255) this.throwVbaError(6, "Overflow");
            return num;
        });
        this.env.set('int', (val: any) => Math.floor(parseFloat(val)) || 0);
        this.env.set('ucase', (val: any) => val === vbaNull ? vbaNull : val === vbaEmpty ? "" : String(val).toUpperCase());
        this.env.set('lcase', (val: any) => val === vbaNull ? vbaNull : val === vbaEmpty ? "" : String(val).toLowerCase());
        this.env.set('trim', (val: any) => val === vbaNull ? vbaNull : val === vbaEmpty ? "" : String(val).trim());
        this.env.set('ltrim', (val: any) => val === vbaNull ? vbaNull : val === vbaEmpty ? "" : String(val).replace(/^\s+/, ''));
        this.env.set('rtrim', (val: any) => val === vbaNull ? vbaNull : val === vbaEmpty ? "" : String(val).replace(/\s+$/, ''));
        this.env.set('len', (val: any) => String(val || '').length);
        this.env.set('left', (val: any, len: number) => String(val || '').substring(0, len));
        this.env.set('right', (val: any, len: number) => {
            const s = String(val || '');
            return s.substring(s.length - len);
        });
        this.env.set('mid', (val: any, start: number, len?: number) => {
            const s = String(val || '');
            return len !== undefined ? s.substring(start - 1, start - 1 + len) : s.substring(start - 1);
        });
        this.env.set('format', (val: any, pattern?: string) => {
            if (val === null || val === vbaNull || val === vbaEmpty) return "";
            const fmt = pattern ? String(pattern) : "";
            if (fmt === "") return String(val);
            if (val instanceof VbaDate) {
                return this.formatDate(fromVbaDate(val.value), fmt);
            }
            if (typeof val === 'number') {
                return this.formatNumber(val, fmt);
            }
            return String(val);
        });
        this.env.set('format$', (val: any, pattern?: string) => {
            return this.env.get('format')(val, pattern);
        });
        this.env.set('instr', (...args: any[]) => {
            let start: number, s1: any, s2: any, compare: number | undefined;
            if (args.length >= 4) {
                [start, s1, s2, compare] = [args[0], args[1], args[2], args[3]];
            } else if (args.length === 3 && typeof args[0] === 'number') {
                [start, s1, s2] = [args[0], args[1], args[2]];
            } else {
                [start, s1, s2] = [1, args[0], args[1]];
            }
            if (s1 === vbaNull || s2 === vbaNull) return vbaNull;
            const str1 = String(s1 ?? '');
            const str2 = String(s2 ?? '');
            
            const isText = (compare === 1) || (compare === undefined && this.comparisonMode === 'Text');
            if (isText) {
                const idx = str1.toLowerCase().indexOf(str2.toLowerCase(), start - 1);
                return idx === -1 ? 0 : idx + 1;
            } else {
                const idx = str1.indexOf(str2, start - 1);
                return idx === -1 ? 0 : idx + 1;
            }
        });
        this.env.set('instrrev', (s1: any, s2: any, start: any = -1, compare: any = undefined) => {
            if (s1 === vbaNull || s2 === vbaNull) return vbaNull;
            const str = String(s1 ?? '');
            const find = String(s2 ?? '');
            if (str === "") return 0;
            if (find === "") return (start === -1 || start === undefined) ? str.length : Number(start);
            
            const effectiveStart = (start === -1 || start === undefined) ? str.length : Number(start);
            if (effectiveStart > str.length) return 0;

            const isText = (compare === 1) || (compare === undefined && this.comparisonMode === 'Text');
            let idx: number;
            if (isText) {
                idx = str.toLowerCase().lastIndexOf(find.toLowerCase(), effectiveStart - 1);
            } else {
                idx = str.lastIndexOf(find, effectiveStart - 1);
            }
            return idx === -1 ? 0 : idx + 1;
        });
        this.env.set('strcomp', (s1: any, s2: any, compare: number = 0) => {
            if (s1 === vbaNull || s2 === vbaNull) return vbaNull;
            let str1 = String(s1 ?? '');
            let str2 = String(s2 ?? '');
            if (compare === 1) { // vbTextCompare
                str1 = str1.toLowerCase();
                str2 = str2.toLowerCase();
            }
            if (str1 < str2) return -1;
            if (str1 > str2) return 1;
            return 0;
        });
        this.env.set('strreverse', (s: any) => {
            if (s === vbaNull) throw new Error('Execution error: Invalid use of Null');
            return String(s ?? '').split('').reverse().join('');
        });
        this.env.set('strconv', (s: any, conversion: any, lcid?: any) => {
            let str = String(s ?? '');
            const c = parseInt(conversion) || 0;

            const casing = c & 3;
            if (casing === 1) { // vbUpperCase
                str = str.toUpperCase();
            } else if (casing === 2) { // vbLowerCase
                str = str.toLowerCase();
            } else if (casing === 3) { // vbProperCase
                str = str.toLowerCase().replace(/(^|[^a-zA-Z0-9])([a-z])/g, (m, p1, p2) => p1 + p2.toUpperCase());
            }

            if (c & 4) { // vbWide
                str = str.replace(/[!-~]/g, m => String.fromCharCode(m.charCodeAt(0) + 0xFEE0));
                const map: any = { '｡': '。', '｢': '「', '｣': '」', '､': '、', '･': '・', 'ｦ': 'ヲ', 'ｧ': 'ァ', 'ｨ': 'ィ', 'ｩ': 'ゥ', 'ｪ': 'ェ', 'ｫ': 'ォ', 'ｬ': 'ャ', 'ｭ': 'ュ', 'ｮ': 'ョ', 'ｯ': 'ッ', 'ｰ': 'ー', 'ｱ': 'ア', 'ｲ': 'イ', 'ｳ': 'ウ', 'ｴ': 'エ', 'ｵ': 'オ', 'ｶ': 'カ', 'ｷ': 'キ', 'ｸ': 'ク', 'ｹ': 'ケ', 'ｺ': 'コ', 'ｻ': 'サ', 'ｼ': 'シ', 'ｽ': 'ス', 'ｾ': 'セ', 'ｿ': 'ソ', 'ﾀ': 'タ', 'ﾁ': 'チ', 'ﾂ': 'ツ', 'ﾃ': 'テ', 'ﾄ': 'ト', 'ﾅ': 'ナ', 'ﾆ': 'ニ', 'ﾇ': 'ヌ', 'ネ': 'ネ', 'ﾉ': 'ノ', 'ﾊ': 'ハ', 'ﾋ': 'ヒ', 'ﾌ': 'フ', 'ﾍ': 'ヘ', 'ﾎ': 'ホ', 'ﾏ': 'マ', 'ﾐ': 'ミ', 'ﾑ': 'ム', 'ﾒ': 'メ', 'ﾓ': 'モ', 'ﾔ': 'ヤ', 'ﾕ': 'ユ', 'ﾖ': 'ヨ', 'ﾗ': 'ラ', 'ﾘ': 'リ', 'ﾙ': 'ル', 'ﾚ': 'レ', 'ﾛ': 'ロ', 'ﾜ': 'ワ', 'ﾝ': 'ン', 'ﾞ': '゛', 'ﾟ': '゜' };
                str = str.replace(/[｡-ﾟ]/g, m => map[m] || m);
            }

            if (c & 8) { // vbNarrow
                str = str.replace(/[！-～]/g, m => String.fromCharCode(m.charCodeAt(0) - 0xFEE0));
                const map: any = { '。': '｡', '「': '｢', '」': '｣', '、': '､', '・': '･', 'ヲ': 'ｦ', 'ァ': 'ｧ', 'ィ': 'ｨ', 'ゥ': 'ｩ', 'ェ': 'ｪ', 'ォ': 'ｫ', 'ャ': 'ｬ', 'ュ': 'ｭ', 'ョ': 'ｮ', 'ッ': 'ｯ', 'ー': 'ｰ', 'ア': 'ｱ', 'イ': 'ｲ', 'ウ': 'ｳ', 'エ': 'ｴ', 'オ': 'ｵ', 'カ': 'ｶ', 'キ': 'ｷ', 'ク': 'ｸ', 'ケ': 'ｹ', 'コ': 'ｺ', 'サ': 'ｻ', 'シ': 'ｼ', 'ス': 'ｽ', 'セ': 'ｾ', 'ソ': 'ｿ', 'タ': 'ﾀ', 'チ': 'ﾁ', 'ツ': 'ﾂ', 'テ': 'ﾃ', 'ト': 'ﾄ', 'ナ': 'ﾅ', 'ニ': 'ﾆ', 'ヌ': 'ﾇ', 'ネ': 'ﾈ', 'ノ': 'ﾉ', 'ハ': 'ﾊ', 'ヒ': 'ﾋ', 'フ': 'ﾌ', 'ヘ': 'ﾍ', 'ホ': 'ﾎ', 'マ': 'ﾏ', 'ミ': 'ﾐ', 'ム': 'ﾑ', 'メ': 'ﾒ', 'モ': 'ﾓ', 'ヤ': 'ﾔ', 'ユ': 'ﾕ', 'ヨ': 'ﾖ', 'ラ': 'ﾗ', 'リ': 'ﾘ', 'ル': 'ﾙ', 'レ': 'ﾚ', 'ロ': 'ﾛ', 'ワ': 'ﾜ', 'ン': 'ﾝ', '゛': 'ﾞ', '゜': 'ﾟ' };
                str = str.replace(/[。-゜]/g, m => map[m] || m);
            }

            if (c & 16) { // vbKatakana
                str = str.replace(/[\u3041-\u3096]/g, m => String.fromCharCode(m.charCodeAt(0) + 0x60));
            }

            if (c & 32) { // vbHiragana
                str = str.replace(/[\u30A1-\u30FA]/g, m => String.fromCharCode(m.charCodeAt(0) - 0x60));
            }

            return str;
        });
        this.env.set('replace', (s: any, find: any, repl: any) => String(s || '').split(String(find || '')).join(String(repl || '')));

        this.env.set('filter', (sourceArray: any, match: any, include: any = vbaTrue, compare: any = 0) => {
            if (!Array.isArray(sourceArray)) {
                this.throwVbaError(13, "Type mismatch: SourceArray must be an array");
            }
            const matchStr = String(match ?? '');
            const includeBool = this.isTrue(include);
            
            const isText = (compare === 1) || (compare === undefined && this.comparisonMode === 'Text');
            
            const result = sourceArray.filter(item => {
                const itemStr = String(item ?? '');
                let found: boolean;
                if (isText) {
                    found = itemStr.toLowerCase().includes(matchStr.toLowerCase());
                } else {
                    found = itemStr.includes(matchStr);
                }
                return includeBool ? found : !found;
            });
            
            const resArray = [...result];
            (resArray as any).vbaBase = 0;
            return resArray;
        });

        this.env.set('cint', (val: any) => {
            const n = this.vbaRound(parseFloat(val) || 0);
            if (n < -32768 || n > 32767) this.throwVbaError(6, "Overflow");
            return n;
        });
        this.env.set('cstr', (val: any) => String(val === null ? '' : val));
        this.env.set('cbool', (val: any) => this.isTrue(val) ? vbaTrue : vbaFalse);
        this.env.set('fix', (val: any) => val > 0 ? Math.floor(val) : Math.ceil(val));
        this.env.set('Val', (s: any) => {
            if (typeof s !== 'string') return 0;
            const cleaned = s.replace(/ /g, '');
            const match = cleaned.match(/^[+-]?\d*(\.\d*)?/);
            return match ? parseFloat(match[0]) || 0 : 0;
        });

        this.env.set('Str', (n: any) => {
            const num = Number(n);
            if (isNaN(num)) return String(n);
            return num >= 0 ? ' ' + String(num) : String(num);
        });

        this.env.set('Oct', (n: any) => {
            const num = Math.floor(Number(n));
            return num.toString(8);
        });

        this.env.set('Hex', (n: any) => {
            const num = Math.floor(Number(n));
            return num.toString(16).toUpperCase();
        });

        this.env.set('TypeName', (val: any) => {
            if (val === vbaEmpty || val === undefined) return 'Empty';
            if (val === vbaNull) return 'Null';
            if (val === vbaNothing) return 'Nothing';
            if (val === vbaMissing) return 'Error';
            if (val instanceof VbaErrorValue) return 'Error';
            if (val instanceof VbaBoolean) return 'Boolean';
            if (val instanceof VbaDate) return 'Date';
            if (typeof val === 'number') return 'Double';
            if (typeof val === 'string') return 'String';
            if (typeof val === 'boolean') return 'Boolean';
            if (Array.isArray(val)) return 'Variant()';
            if (val.__isVbaDict__) return 'Dictionary';
            if (val.__isVbaCollection__) return 'Collection';
            if (val.__vbaClass__) return val.__className__;
            if (val.__vbaTypeName__) return val.__vbaTypeName__;
            if (typeof val === 'object') return 'Object';
            return 'Unknown';
        });

        this.env.set('abs', (val: any) => Math.abs(val));
        this.env.set('round', (val: any, digits: any = 0) => {
            return this.vbaRound(parseFloat(val) || 0, parseInt(digits) || 0);
        });
        this.env.set('sqr', (val: any) => Math.sqrt(val));

        // --- Math Module (§6.1.2.10) ---
        this.env.set('sgn', (val: any) => {
            const n = Number(val);
            return n > 0 ? 1 : n < 0 ? -1 : 0;
        });
        this.env.set('atn', (val: any) => Math.atan(Number(val)));
        this.env.set('cos', (val: any) => Math.cos(Number(val)));
        this.env.set('sin', (val: any) => Math.sin(Number(val)));
        this.env.set('tan', (val: any) => Math.tan(Number(val)));
        this.env.set('exp', (val: any) => {
            const n = Number(val);
            if (n > 709.782712893) throw new Error('Execution error: Overflow');
            return Math.exp(n);
        });
        this.env.set('log', (val: any) => {
            const n = Number(val);
            if (n <= 0) throw new Error('Execution error: Invalid procedure call or argument');
            return Math.log(n);
        });

        // Rnd / Randomize: shared state for the random sequence
        let rndSeed = 0.5;   // initial seed value
        let lastRnd = 0.5;   // most recently generated value
        let rndInitialized = false;

        // Simple LCG for seeded Rnd(< 0) — matches "same every time for given seed"
        const seededRnd = (seed: number): number => {
            // Use a deterministic formula: frac(abs(seed) * some_constant)
            const s = Math.abs(seed) * 9301 + 49297;
            return (s % 233280) / 233280;
        };

        this.env.set('rnd', (val?: any) => {
            if (!rndInitialized) {
                // Without Randomize, use a fixed initial sequence
                rndSeed = 0.5;
                rndInitialized = true;
            }
            if (val === undefined || val === null || (typeof val === 'number' && val > 0)) {
                // Next in sequence
                rndSeed = (rndSeed * 214013 + 2531011) % 4294967296;
                lastRnd = rndSeed / 4294967296;
                return lastRnd;
            } else if (typeof val === 'number' && val === 0) {
                // Return most recently generated number
                return lastRnd;
            } else if (typeof val === 'number' && val < 0) {
                // Same number every time for given seed
                const r = seededRnd(val);
                lastRnd = r;
                return r;
            }
            return lastRnd;
        });

        this.env.set('randomize', (val?: any) => {
            rndInitialized = true;
            if (val === undefined || val === null) {
                rndSeed = Date.now() % 4294967296;
            } else {
                // Seed from given number — deterministic
                const n = Math.abs(Number(val));
                rndSeed = Math.round(n * 1000) % 4294967296;
            }
            lastRnd = rndSeed / 4294967296;
        });

        this.env.set('isnull', (val: any) => val === vbaNull ? vbaTrue : vbaFalse);
        this.env.set('isdate', (val: any) => {
            if (val === null || val === undefined || val === vbaNull || val === vbaEmpty || val === vbaMissing) return vbaFalse;
            if (val instanceof VbaDate) return vbaTrue;
            if (typeof val === 'number') return vbaTrue;
            if (typeof val === 'string') {
                const d = new Date(val);
                return !isNaN(d.getTime()) ? vbaTrue : vbaFalse;
            }
            return vbaFalse;
        });
        const errorMessages: Record<number, string> = {
            3: "Return without GoSub",
            5: "Invalid procedure call or argument",
            6: "Overflow",
            7: "Out of memory",
            9: "Subscript out of range",
            10: "This array is fixed or temporarily locked",
            11: "Division by zero",
            13: "Type mismatch",
            14: "Out of string space",
            16: "Expression too complex",
            17: "Can't perform requested operation",
            18: "User interrupt occurred",
            20: "Resume without error",
            28: "Out of stack space",
            35: "Sub or Function not defined",
            48: "Error in loading DLL",
            51: "Internal error",
            52: "Bad file name or number",
            53: "File not found",
            54: "Bad file mode",
            55: "File already open",
            57: "Device I/O error",
            58: "File already exists",
            59: "Bad record length",
            61: "Disk full",
            62: "Input past end of file",
            63: "Bad record number",
            67: "Too many files",
            68: "Device unavailable",
            70: "Permission denied",
            71: "Disk not ready",
            74: "Can't rename with different drive",
            75: "Path/File access error",
            76: "Path not found",
            91: "Object variable or With block variable not set",
            92: "For loop not initialized",
            93: "Invalid pattern string",
            94: "Invalid use of Null",
            321: "Invalid file format",
            322: "Can't create necessary temporary file",
            325: "Invalid format in resource file",
            380: "Invalid property value",
            424: "Object required",
            429: "ActiveX component can't create object",
            430: "Class doesn't support Automation or doesn't support expected interface",
            438: "Object doesn't support this property or method",
            440: "Automation error",
            445: "Object doesn't support this action",
            446: "Object doesn't support named arguments",
            447: "Object doesn't support current locale settings",
            448: "Named argument not found",
            449: "Argument not optional",
            450: "Wrong number of arguments or invalid property assignment",
            451: "Property let procedure not defined and property get procedure did not return an object",
            453: "Specified DLL function not found",
            457: "This key is already associated with an element of this collection",
            458: "Variable uses an Automation type not supported in Visual Basic",
            459: "Object or class does not support the set of events",
            460: "Invalid clipboard format",
            461: "Method or data member not found",
            462: "The remote server machine does not exist or is unavailable",
            463: "Class not registered on local machine",
            481: "Invalid picture",
            482: "Printer error",
            735: "Can't save file to TEMP",
            744: "Search text not found",
            746: "Replacements too long"
        };

        const errorFunc = (errNum?: any) => {
            const num = (errNum === undefined) ? this.env.get('err').number : (parseInt(errNum) || 0);
            if (num === 0) return "";
            if (num > 65535) this.throwVbaError(6, "Overflow");
            return errorMessages[num] || "Application-defined or object-defined error";
        };
        (errorFunc as any).__vbaAutoCall__ = true;
        this.env.set('error', errorFunc);
        this.env.set('error$', errorFunc);

        this.env.set('iserror', (val: any) => (val instanceof VbaErrorValue) ? vbaTrue : vbaFalse);
        this.env.set('cverr', (val: any) => {
            if (val instanceof VbaErrorValue) return val;
            const code = parseInt(val) || 0;
            if (code < 0 || code > 65535) {
                this.throwVbaError(5, "Invalid procedure call or argument");
            }
            return new VbaErrorValue(code);
        });
        this.env.set('ismissing', (val: any) => (val === vbaMissing) ? vbaTrue : vbaFalse);
        this.env.set('isarray', (val: any) => Array.isArray(val) ? vbaTrue : vbaFalse);
        this.env.set('isobject', (val: any) => (val !== null && typeof val === 'object' && !Array.isArray(val)) ? vbaTrue : vbaFalse);
        this.env.set('choose', (index: any, ...choices: any[]) => {
            const idx = Math.floor(Number(index));
            if (idx >= 1 && idx <= choices.length) {
                return choices[idx - 1];
            }
            return vbaNull;
        });
        this.env.set('switch', (...args: any[]) => {
            for (let i = 0; i < args.length; i += 2) {
                if (i + 1 < args.length) {
                    const cond = args[i];
                    if (cond === vbaTrue || cond === true) {
                        return args[i + 1];
                    }
                }
            }
            return vbaNull;
        });

        this.env.set('vartype', (val: any) => {
            if (val === vbaEmpty || val === undefined) return 0; // vbEmpty
            if (val === vbaNull) return 1; // vbNull
            if (val instanceof VbaBoolean) return 11; // vbBoolean
            if (val instanceof VbaDate) return 7; // vbDate
            if (val === vbaMissing) return 10; // vbError
            if (Array.isArray(val)) return 8192 + 12; // vbArray + vbVariant
            if (typeof val === 'number') return 5; // vbDouble
            if (typeof val === 'string') return 8; // vbString
            if (typeof val === 'object') return 9; // vbObject
            return 12; // vbVariant
        });

        // VBA Type Constants (§6.1.1)
        this.env.set('vbempty', 0);
        this.env.set('vbnull', 1);
        this.env.set('vbinteger', 2);
        this.env.set('vblong', 3);
        this.env.set('vbsingle', 4);
        this.env.set('vbdouble', 5);
        this.env.set('vbcurrency', 6);
        this.env.set('vbdate', 7);
        this.env.set('vbstring', 8);
        this.env.set('vbobject', 9);
        this.env.set('vberror', 10);
        this.env.set('vboolean', 11);
        this.env.set('vbvariant', 12);
        this.env.set('vbdataobject', 13);
        this.env.set('vbdecimal', 14);
        this.env.set('vbbyte', 17);
        this.env.set('lbound', (arr: any) => {
            if (Array.isArray(arr)) {
                return (arr as any).vbaBase || 0;
            }
            return 0;
        });

        this.env.set('iif', (cond: any, truePart: any, falsePart: any) => cond ? truePart : falsePart);
        this.env.set('array', (...args: any[]) => [...args]);
        this.env.set('split', (s: any, delimiter: string = ' ') => String(s || '').split(delimiter));
        this.env.set('join', (arr: any[], delimiter: string = ' ') => Array.isArray(arr) ? arr.join(delimiter) : String(arr));
        this.env.set('asc', (s: any) => String(s || '').charCodeAt(0));
        this.env.set('chr', (n: number) => String.fromCharCode(n));
        this.env.set('space', (n: number) => ' '.repeat(n));
        this.env.set('string', (n: number, char: any) => {
            const s = String(char || '');
            return (s.length > 0 ? s[0] : '').repeat(n);
        });

        this.env.set('ubound', (arr: any, dimension?: number) => {
            if (Array.isArray(arr)) {
                const base = (arr as any).vbaBase || 0;
                if (dimension === 2 && arr.length > 0 && Array.isArray(arr[0])) {
                    const subBase = (arr[0] as any).vbaBase || 0;
                    return subBase + arr[0].length - 1;
                }
                return base + arr.length - 1;
            }
            return 0;
        });
        this.env.set('createobject', (progId: string) => {
            return this.createExternalObject(progId);
        });

        // Add VBA intrinsic constants
        this.env.set('true', vbaTrue);
        this.env.set('false', vbaFalse);
        this.env.set('empty', vbaEmpty);
        this.env.set('nothing', vbaNothing);
        this.env.set('null', vbaNull);



        // Weekday constants (vbSunday=1 ... vbSaturday=7)
        this.env.set('vbsunday', 1);
        this.env.set('vbmonday', 2);
        this.env.set('vbtuesday', 3);
        this.env.set('vbwednesday', 4);
        this.env.set('vbthursday', 5);
        this.env.set('vbfriday', 6);
        this.env.set('vbsaturday', 7);
        this.env.set('vbusesystemdayofweek', 0);

        // --- DateTime extraction functions ---
        this.env.set('year',   (d: any) => d === null ? null : parseVbaDate(d).getFullYear());
        this.env.set('month',  (d: any) => d === null ? null : parseVbaDate(d).getMonth() + 1);
        this.env.set('day',    (d: any) => d === null ? null : parseVbaDate(d).getDate());
        this.env.set('hour',   (d: any) => d === null ? null : parseVbaDate(d).getHours());
        this.env.set('minute', (d: any) => d === null ? null : parseVbaDate(d).getMinutes());
        this.env.set('second', (d: any) => d === null ? null : parseVbaDate(d).getSeconds());

        this.env.set('weekday', (d: any, firstDay: number = 1) => {
            if (d === null) return null;
            const jsDay = parseVbaDate(d).getDay(); // 0=Sun
            // VBA: firstDay=1(Sun)..7(Sat). JS getDay() 0=Sun..6=Sat
            // Shift so that firstDay maps to 1
            const offset = (firstDay <= 1) ? 0 : (firstDay - 1);
            return ((jsDay - offset + 7) % 7) + 1;
        });

        // --- DateSerial / TimeSerial ---
        this.env.set('dateserial', (year: number, month: number, day: number) => {
            const d = new Date(year, month - 1, day);
            if (year >= 0 && year <= 99) d.setFullYear(year); // preserve 2-digit years as-is
            return toVbaDate(d);
        });

        this.env.set('timeserial', (hour: number, minute: number, second: number) => {
            const totalSec = hour * 3600 + minute * 60 + second;
            return totalSec / 86400; // fraction of a day
        });

        // --- DateValue / TimeValue ---
        this.env.set('datevalue', (s: any) => {
            if (s === null) return null;
            const d = new Date(String(s));
            if (isNaN(d.getTime())) throw new Error(`Execution error: Type mismatch: '${s}'`);
            // Return date-only serial (strip time portion)
            return Math.floor(toVbaDate(d));
        });

        this.env.set('timevalue', (s: any) => {
            if (s === null) return null;
            const d = new Date(`1970-01-01T${String(s)}`);
            if (isNaN(d.getTime())) {
                const d2 = new Date(String(s));
                if (isNaN(d2.getTime())) throw new Error(`Execution error: Type mismatch: '${s}'`);
                return toVbaDate(d2) % 1;
            }
            return (d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds()) / 86400;
        });

        // --- DateAdd ---
        // Add months/years with VBA's month-end clamping behavior
        const addMonths = (d: Date, months: number): Date => {
            const day = d.getDate();
            const result = new Date(d);
            result.setDate(1); // avoid overflow when changing month
            result.setMonth(result.getMonth() + months);
            // Clamp to last day of the resulting month
            const lastDay = new Date(result.getFullYear(), result.getMonth() + 1, 0).getDate();
            result.setDate(Math.min(day, lastDay));
            return result;
        };

        this.env.set('dateadd', (interval: string, number: number, date: any) => {
            const d = parseVbaDate(date);
            const n = Math.round(number);
            const iv = String(interval).toLowerCase();
            let result: Date;
            switch (iv) {
                case 'yyyy': result = addMonths(d, n * 12); break;
                case 'q':    result = addMonths(d, n * 3); break;
                case 'm':    result = addMonths(d, n); break;
                case 'y': case 'd': case 'w':
                    result = new Date(d); result.setDate(result.getDate() + n); break;
                case 'ww':
                    result = new Date(d); result.setDate(result.getDate() + n * 7); break;
                case 'h':
                    result = new Date(d); result.setHours(result.getHours() + n); break;
                case 'n':
                    result = new Date(d); result.setMinutes(result.getMinutes() + n); break;
                case 's':
                    result = new Date(d); result.setSeconds(result.getSeconds() + n); break;
                default: throw new Error('Execution error: Invalid procedure call or argument (DateAdd interval)');
            }
            return toVbaDate(result);
        });

        // --- DateDiff ---
        this.env.set('datediff', (interval: string, date1: any, date2: any, firstDayOfWeek: number = 1) => {
            const d1 = parseVbaDate(date1);
            const d2 = parseVbaDate(date2);
            const iv = String(interval).toLowerCase();
            switch (iv) {
                case 'yyyy': return d2.getFullYear() - d1.getFullYear();
                case 'q':    return (d2.getFullYear() - d1.getFullYear()) * 4 +
                                    Math.floor(d2.getMonth() / 3) - Math.floor(d1.getMonth() / 3);
                case 'm':    return (d2.getFullYear() - d1.getFullYear()) * 12 +
                                    (d2.getMonth() - d1.getMonth());
                case 'y': case 'd': return Math.floor(toVbaDate(d2)) - Math.floor(toVbaDate(d1));
                case 'w':    return Math.floor((toVbaDate(d2) - toVbaDate(d1)) / 7);
                case 'ww': {
                    // Count week boundaries between d1 and d2
                    const offset = (firstDayOfWeek <= 1) ? 0 : firstDayOfWeek - 1;
                    const adj1 = Math.floor((Math.floor(toVbaDate(d1)) - offset + 7) / 7);
                    const adj2 = Math.floor((Math.floor(toVbaDate(d2)) - offset + 7) / 7);
                    return adj2 - adj1;
                }
                case 'h': return Math.floor((d2.getTime() - d1.getTime()) / 3600000);
                case 'n': return Math.floor((d2.getTime() - d1.getTime()) / 60000);
                case 's': return Math.floor((d2.getTime() - d1.getTime()) / 1000);
                default: throw new Error('Execution error: Invalid procedure call or argument (DateDiff interval)');
            }
        });

        // --- DatePart ---
        this.env.set('datepart', (interval: string, date: any, firstDayOfWeek: number = 1) => {
            const d = parseVbaDate(date);
            const iv = String(interval).toLowerCase();
            switch (iv) {
                case 'yyyy': return d.getFullYear();
                case 'q':    return Math.floor(d.getMonth() / 3) + 1;
                case 'm':    return d.getMonth() + 1;
                case 'y':    return Math.floor(toVbaDate(d)) - Math.floor(toVbaDate(new Date(d.getFullYear(), 0, 1))) + 1;
                case 'd':    return d.getDate();
                case 'w': {
                    const jsDay = d.getDay();
                    const offset = (firstDayOfWeek <= 1) ? 0 : (firstDayOfWeek - 1);
                    return ((jsDay - offset + 7) % 7) + 1;
                }
                case 'ww': return Math.ceil((Math.floor(toVbaDate(d)) - toVbaDate(new Date(d.getFullYear(), 0, 1)) + 1) / 7);
                case 'h':  return d.getHours();
                case 'n':  return d.getMinutes();
                case 's':  return d.getSeconds();
                default: throw new Error('Execution error: Invalid procedure call or argument (DatePart interval)');
            }
        });

        // --- Now / Date / Time / Timer ---
        this.env.set('now',   () => toVbaDate(new Date()));
        this.env.set('date',  () => Math.floor(toVbaDate(new Date())));
        this.env.set('date$', () => {
            const d = new Date();
            const mm = String(d.getMonth() + 1).padStart(2, '0');
            const dd = String(d.getDate()).padStart(2, '0');
            const yyyy = d.getFullYear();
            return `${mm}-${dd}-${yyyy}`;
        });
        this.env.set('time',  () => {
            const d = new Date();
            return (d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds()) / 86400;
        });
        this.env.set('time$', () => {
            const d = new Date();
            const hh = String(d.getHours()).padStart(2, '0');
            const mm = String(d.getMinutes()).padStart(2, '0');
            const ss = String(d.getSeconds()).padStart(2, '0');
            return `${hh}:${mm}:${ss}`;
        });
        this.env.set('timer', () => {
            const d = new Date();
            return d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds() + d.getMilliseconds() / 1000;
        });

        // Add common Excel VBA constants
        this.env.set('xlup', -4162);
        this.env.set('xldown', -4121);
        this.env.set('xltoleft', -4159);
        this.env.set('xltoright', -4161);

        // MsgBox constants
        this.env.set('vbokonly', 0);
        this.env.set('vboxcancel', 1);
        this.env.set('vbabortretryignore', 2);
        this.env.set('vbyesnocancel', 3);
        this.env.set('vbyesno', 4);
        this.env.set('vbretrycancel', 5);
        this.env.set('vbcritical', 16);
        this.env.set('vbquestion', 32);
        this.env.set('vboxclamation', 48);
        this.env.set('vbinformation', 64);
        this.env.set('vbdefaultbutton1', 0);
        this.env.set('vbdefaultbutton2', 256);

        this.env.set('vbok', 1);
        this.env.set('vbcancel', 2);
        this.env.set('vbabort', 3);
        this.env.set('vbretry', 4);
        this.env.set('vbignore', 5);
        this.env.set('vbyes', 6);
        this.env.set('vbno', 7);

        this.env.set('msgbox', (prompt: any, buttons: number = 0, title: string = "Microsoft Excel") => {
            this.onPrint(`[MSGBOX] ${title}: ${prompt} (Buttons: ${buttons})`);
            return 1; // vbOK
        });

        this.env.set('inputbox', (prompt: any, title: string = "Microsoft Excel", defaultVal: string = "") => {
            this.onPrint(`[INPUTBOX] ${title}: ${prompt} (Default: ${defaultVal})`);
            return defaultVal;
        });

        // Add VBA Err object
        this.env.set('err', {
            number: 0,
            source: '',
            description: '',
            clear: function () { this.number = 0; this.source = ''; this.description = ''; },
            raise: function (num: number, src?: string, desc?: string) {
                this.number = num;
                this.source = src || '';
                this.description = desc || '';
                throw { type: 'VbaError', number: num, message: desc || `VBA Error ${num}` };
            }
        });
    }

    public get(name: string): any {
        return this.env.get(name);
    }

    public set(name: string, value: any): void {
        this.env.set(name, value);
    }

    public callProcedure(name: string, args: any[], type?: 'get' | 'let' | 'set'): any {
        const procName = name.toLowerCase();
        const proc = this.env.getProcedure(procName, type);

        if (!proc) {
            // Fall back to built-in functions stored as closures in env
            const builtin = this.env.get(procName);
            if (typeof builtin === 'function') {
                return builtin(...args);
            }
            throw new Error(`Execution error: Procedure '${name}' not found`);
        }

        // Create a new local environment for the procedure call
        const localEnv = new Environment(this.env);

        // Map arguments to parameter names
        for (let i = 0; i < proc.parameters.length; i++) {
            const param = proc.parameters[i];
            const paramName = param.name;
            const argValue = i < args.length ? args[i] : (param.isOptional ? vbaMissing : vbaEmpty);
            localEnv.setLocally(paramName, argValue);
        }

        // Save current env and error handler state, swap to local
        const previousEnv = this.env;
        const previousErrorHandler = this.errorHandlerLabel;
        const previousErrorHandlingMode = this.errorHandlingMode;
        const previousIsInErrorHandler = this.isInErrorHandler;
        const previousLastErrorIndex = this.lastErrorIndex;

        const previousProcBody = this.currentProcBody;
        const previousProcedureName = this.currentProcedureName;
        const previousProcedureType = this.currentProcedureType;
        const previousExecutingModule = this.executingModuleName;
        const previousProcIsStatic = this.currentProcIsStatic;
        const previousStaticVars = this.staticVarsInCurrentProc;

        this.env = localEnv;
        this.errorHandlerLabel = null;
        this.errorHandlingMode = 'None';
        this.isInErrorHandler = false;
        this.lastErrorIndex = -1;

        this.currentProcBody = proc.body;
        this.currentProcedureName = proc.name.name;
        this.currentProcedureType = proc.propertyType || (proc.isFunction ? 'function' : 'sub');
        this.executingModuleName = proc.moduleName ?? '';
        this.currentProcIsStatic = proc.isStatic ?? false;
        this.staticVarsInCurrentProc = new Set();

        try {
            // Execute procedure body with error handling support
            this.executeStatements(proc.body, 0);
        } catch (e: any) {
            if (e && e.type === 'Exit') {
                if (
                    (e.target === 'Function' && proc.isFunction) ||
                    (e.target === 'Sub' && !proc.isFunction && !proc.isProperty) ||
                    (e.target === 'Property' && proc.isProperty)
                ) {
                    // Valid exit, caught and swallowed
                } else {
                    throw e; // Unhandled exit type
                }
            } else {
                throw e; // Real error
            }
        } finally {
            // Persist static variable values
            for (const varName of this.staticVarsInCurrentProc) {
                const key = `${procName}:${varName}`;
                this.staticVarStore.set(key, localEnv.get(varName));
            }
            // Restore previous environment and error handler state
            this.env = previousEnv;
            this.errorHandlerLabel = previousErrorHandler;
            this.errorHandlingMode = previousErrorHandlingMode;
            this.isInErrorHandler = previousIsInErrorHandler;
            this.lastErrorIndex = previousLastErrorIndex;

            this.currentProcBody = previousProcBody;
            this.currentProcedureName = previousProcedureName;
            this.currentProcedureType = previousProcedureType;
            this.executingModuleName = previousExecutingModule;
            this.currentProcIsStatic = previousProcIsStatic;
            this.staticVarsInCurrentProc = previousStaticVars;
        }

        // Return the function or property value
        if (proc.isFunction || proc.isProperty) {
            return localEnv.get(procName);
        }
        return vbaEmpty;
    }

    public setSourceModule(moduleName: string) {
        this.currentSourceModule = moduleName;
    }

    public evaluate(program: Program) {
        for (const stmt of program.body) {
            this.evaluateStatement(stmt);
        }
    }

    public evalExpression(exprString: string): any {
        const lexer = new Lexer(exprString);
        const tokens = lexer.tokenize();

        try {
            const exprParser = new Parser(tokens);
            const expr = exprParser.parseExpressionPublic();
            const nextToken = (exprParser as any).peek();

            // If the expression consumed the entire line (or stopped at EOF)
            if (!nextToken || nextToken.type === TokenType.EOF || nextToken.type === TokenType.Newline) {
                // If it is just an Identifier matching a Procedure, run it as a CallStatement
                if (expr.type === 'Identifier') {
                    const name = (expr as Identifier).name;
                    if (this.env.getProcedure(name)) {
                        this.callProcedure(name, []);
                        return undefined;
                    }
                }
                // Otherwise evaluate as a typical expression returning a value
                return this.evaluateExpression(expr);
            }
        } catch {
            // Ignored: fallback to full statement parsing
        }

        // Fallback: parse and evaluate as a Statement (returns undefined)
        const stmtParser = new Parser(tokens);
        const program = stmtParser.parse();
        this.evaluate(program);
        return undefined;
    }

    private evaluateStatement(stmt: Statement) {
        switch (stmt.type) {
            case 'ForStatement':
                this.evaluateForStatement(stmt as ForStatement);
                break;
            case 'ForEachStatement':
                this.evaluateForEachStatement(stmt as ForEachStatement);
                break;
            case 'IfStatement':
                this.evaluateIfStatement(stmt as IfStatement);
                break;
            case 'DoWhileStatement':
                this.evaluateDoWhileStatement(stmt as DoWhileStatement);
                break;
            case 'WhileStatement':
                this.evaluateWhileStatement(stmt as WhileStatement);
                break;
            case 'AssignmentStatement':
                this.evaluateAssignmentStatement(stmt as AssignmentStatement);
                break;
            case 'ProcedureDeclaration': {
                const procDecl = stmt as ProcedureDeclaration;
                procDecl.moduleName = this.currentSourceModule;
                this.env.setProcedure(procDecl.name.name, procDecl);
                break;
            }
            case 'VariableDeclaration':
                this.evaluateVariableDeclaration(stmt as VariableDeclaration);
                break;
            case 'CallStatement':
                this.evaluateCallStatement(stmt as CallStatement);
                break;
            case 'ConstDeclaration':
                this.evaluateConstDeclaration(stmt as ConstDeclaration);
                break;
            case 'SetStatement':
                this.evaluateSetStatement(stmt as SetStatement);
                break;
            case 'OnErrorStatement':
                this.evaluateOnErrorStatement(stmt as OnErrorStatement);
                break;
            case 'ResumeStatement':
                this.evaluateResumeStatement(stmt as ResumeStatement);
                break;
            case 'EraseStatement':
                this.evaluateEraseStatement(stmt as EraseStatement);
                break;
            case 'ReDimStatement':
                this.evaluateReDimStatement(stmt as ReDimStatement);
                break;
            case 'ExitStatement':
                this.evaluateExitStatement(stmt as ExitStatement);
                break;
            case 'EnumDeclaration':
                this.evaluateEnumDeclaration(stmt as EnumDeclaration);
                break;
            case 'TypeDeclaration':
                this.evaluateTypeDeclaration(stmt as TypeDeclaration);
                break;
            case 'ClassDeclaration':
                this.evaluateClassDeclaration(stmt as ClassDeclaration);
                break;
            case 'OptionCompareStatement':
                this.evaluateOptionCompareStatement(stmt as OptionCompareStatement);
                break;
            case 'AttributeStatement':
                this.evaluateAttributeStatement(stmt as AttributeStatement);
                break;
            case 'DeclareStatement':
                this.evaluateDeclareStatement(stmt as DeclareStatement);
                break;
            case 'LabelStatement':
                // No-op for now. Label execution just passes through.
                break;
            case 'SelectCaseStatement':
                this.evaluateSelectCaseStatement(stmt as SelectCaseStatement);
                break;
            case 'WithStatement':
                this.evaluateWithStatement(stmt as WithStatement);
                break;
            case 'GoToStatement':
                this.evaluateGoToStatement(stmt as GoToStatement);
                break;
            case 'StopStatement':
                this.evaluateStopStatement(stmt as StopStatement);
                break;
            case 'EndStatement':
                this.evaluateEndStatement(stmt as EndStatement);
                break;
            case 'GoSubStatement':
                this.evaluateGoSubStatement(stmt as GoSubStatement);
                break;
            case 'ReturnStatement':
                this.evaluateReturnStatement(stmt as ReturnStatement);
                break;
            case 'OnGoToSubStatement':
                this.evaluateOnGoToSubStatement(stmt as OnGoToSubStatement);
                break;
            case 'OptionExplicitStatement':
                this.evaluateOptionExplicitStatement(stmt as OptionExplicitStatement);
                break;
            case 'OptionBaseStatement':
                this.evaluateOptionBaseStatement(stmt as OptionBaseStatement);
                break;
            case 'OptionPrivateModuleStatement':
                this.evaluateOptionPrivateModuleStatement(stmt as OptionPrivateModuleStatement);
                break;
            case 'LSetStatement':
                this.evaluateLSetStatement(stmt as LSetStatement);
                break;
            case 'RSetStatement':
                this.evaluateRSetStatement(stmt as RSetStatement);
                break;
            case 'ErrorStatement':
                this.evaluateErrorStatement(stmt as ErrorStatement);
                break;
            default:
                throw new Error(`Execution error: Unknown statement type ${stmt.type}`);
        }
    }

    private evaluateForStatement(stmt: ForStatement) {
        const startValue = this.evaluateExpression(stmt.start);
        const endValue = this.evaluateExpression(stmt.end);
        const stepValue = stmt.step ? this.evaluateExpression(stmt.step) : 1;
        const varName = stmt.identifier.name;

        // Initialize block scope variable if it doesn't exist
        if (this.env.get(varName) === vbaEmpty) { // Check against vbaEmpty
            this.env.set(varName, startValue);
        } else {
            this.env.setLocally(varName, startValue);
        }

        const condition = () => stepValue > 0 ? this.env.get(varName) <= endValue : this.env.get(varName) >= endValue;

        while (condition()) {
            try {
                for (const bodyStmt of stmt.body) {
                    this.evaluateStatement(bodyStmt);
                }
            } catch (e: any) {
                if (e && e.type === 'Exit' && e.target === 'For') {
                    break;
                }
                throw e; // re-throw if it wasn't an Exit For
            }
            // Increment/decrement loop variable
            this.env.setLocally(varName, this.env.get(varName) + stepValue);
        }
    }

    private evaluateForEachStatement(stmt: ForEachStatement) {
        const collection = this.evaluateExpression(stmt.collection);
        const varName = stmt.variable.name;

        let elements: any[];
        if (Array.isArray(collection)) {
            elements = this.flattenArray(collection);
        } else if (collection && collection.__isVbaDict__) {
            elements = Array.from((collection.__map__ as Map<any, any>).keys());
        } else if (collection && typeof collection.items !== 'undefined') {
            elements = Array.isArray(collection.items) ? collection.items : [];
        } else {
            throw new Error(`Execution error: 'For Each' requires a collection or array`);
        }

        for (const element of elements) {
            this.env.set(varName, element);
            try {
                for (const bodyStmt of stmt.body) {
                    this.evaluateStatement(bodyStmt);
                }
            } catch (e: any) {
                if (e && e.type === 'Exit' && e.target === 'For') {
                    break;
                }
                throw e;
            }
        }
    }

    private flattenArray(arr: any[]): any[] {
        const result: any[] = [];
        for (const item of arr) {
            if (Array.isArray(item)) {
                result.push(...this.flattenArray(item));
            } else {
                result.push(item);
            }
        }
        return result;
    }

    private vbaRound(val: number, decimals: number = 0): number {
        const factor = Math.pow(10, decimals);
        // Use a small epsilon to handle floating point precision issues before rounding
        const scaled = val * factor;
        const i = Math.floor(scaled);
        const f = scaled - i;
        
        // Handle very close to midpoint due to float precision
        const epsilon = 1e-10;
        if (Math.abs(f - 0.5) < epsilon) {
            return ((i % 2 === 0) ? i : i + 1) / factor;
        }

        if (f < 0.5) return i / factor;
        if (f > 0.5) return (i + 1) / factor;
        return ((i % 2 === 0) ? i : i + 1) / factor;
    }

    private throwVbaError(number: number, message: string) {
        const err: any = new Error(message);
        err.type = 'VbaError';
        err.number = number;
        throw err;
    }

    private evaluateIfStatement(stmt: IfStatement) {
        const conditionVal = this.evaluateExpression(stmt.condition);
        if (this.isTrue(conditionVal)) {
            for (const bodyStmt of stmt.consequent) {
                this.evaluateStatement(bodyStmt);
            }
        } else if (stmt.alternate) {
            if (Array.isArray(stmt.alternate)) {
                for (const bodyStmt of stmt.alternate) {
                    this.evaluateStatement(bodyStmt);
                }
            } else {
                this.evaluateIfStatement(stmt.alternate as IfStatement);
            }
        }
    }

    private evaluateSelectCaseStatement(stmt: SelectCaseStatement) {
        const selectVal = this.evaluateExpression(stmt.expression);

        for (const caseClause of stmt.cases) {
            let matched = false;
            for (const range of caseClause.ranges) {
                if (range.kind === 'expression') {
                    matched = selectVal === this.evaluateExpression(range.value);
                } else if (range.kind === 'to') {
                    const start = this.evaluateExpression(range.start);
                    const end = this.evaluateExpression(range.end);
                    matched = selectVal >= start && selectVal <= end;
                } else {
                    // comparison: selectVal <op> value
                    const val = this.evaluateExpression(range.value);
                    switch (range.operator) {
                        case '=':  matched = selectVal === val; break;
                        case '<>': matched = selectVal !== val; break;
                        case '<':  matched = selectVal < val;   break;
                        case '>':  matched = selectVal > val;   break;
                        case '<=': matched = selectVal <= val;  break;
                        case '>=': matched = selectVal >= val;  break;
                    }
                }
                if (matched) break;
            }
            if (matched) {
                for (const s of caseClause.body) this.evaluateStatement(s);
                return;
            }
        }

        if (stmt.elseBody) {
            for (const s of stmt.elseBody) this.evaluateStatement(s);
        }
    }

    private evaluateDoWhileStatement(stmt: DoWhileStatement) {
        const checkCondition = (): boolean => {
            if (stmt.condition === undefined) return true; // infinite
            const val = this.evaluateExpression(stmt.condition);
            const truthy = this.isTrue(val);
            return stmt.conditionType === 'until' ? !truthy : truthy;
        };

        // eslint-disable-next-line no-constant-condition
        while (true) {
            // pre-condition check
            if (stmt.conditionPosition === 'pre' && !checkCondition()) break;

            try {
                for (const bodyStmt of stmt.body) {
                    this.evaluateStatement(bodyStmt);
                }
            } catch (e: any) {
                if (e && e.type === 'Exit' && e.target === 'Do') {
                    break;
                }
                throw e;
            }

            // post-condition check
            if (stmt.conditionPosition === 'post' && !checkCondition()) break;

            // infinite loop: no condition, no break unless Exit Do
            if (stmt.conditionPosition === undefined) continue;
        }
    }

    private evaluateWhileStatement(stmt: WhileStatement) {
        while (this.isTrue(this.evaluateExpression(stmt.condition))) {
            for (const bodyStmt of stmt.body) {
                this.evaluateStatement(bodyStmt);
            }
        }
    }

    private evaluateWithStatement(stmt: WithStatement) {
        const obj = this.evaluateExpression(stmt.expression);
        this.withObjectStack.push(obj);
        try {
            for (const bodyStmt of stmt.body) {
                this.evaluateStatement(bodyStmt);
            }
        } finally {
            this.withObjectStack.pop();
        }
    }

    private evaluateGoToStatement(stmt: GoToStatement) {
        throw { type: 'GoTo', label: stmt.label };
    }

    private evaluateStopStatement(stmt: StopStatement) {
        console.log('STOP Statement encountered');
        // implementation-defined: just log for now
    }

    private evaluateEndStatement(stmt: EndStatement) {
        throw { type: 'Terminate' };
    }

    private evaluateGoSubStatement(stmt: GoSubStatement) {
        throw { type: 'GoSub', label: stmt.label };
    }

    private evaluateReturnStatement(stmt: ReturnStatement) {
        throw { type: 'Return' };
    }

    private evaluateOnGoToSubStatement(stmt: OnGoToSubStatement) {
        const val = this.evaluateExpression(stmt.expression);
        const idx = Math.floor(Number(val));

        if (idx < 0 || idx > 255) {
            throw new Error(`Execution error: Invalid procedure call or argument (On...GoTo/GoSub index ${idx})`);
        }

        if (idx >= 1 && idx <= stmt.labels.length) {
            const label = stmt.labels[idx - 1];
            if (stmt.isGoSub) {
                throw { type: 'GoSub', label };
            } else {
                throw { type: 'GoTo', label };
            }
        }
    }

    private evaluateLSetStatement(stmt: LSetStatement) {
        const val = String(this.evaluateExpression(stmt.right) || '');
        if (stmt.left.type === 'Identifier') {
            const name = (stmt.left as Identifier).name;
            const target = String(this.env.get(name) || '');
            const result = val.padEnd(target.length, ' ').substring(0, target.length);
            this.env.set(name, result);
        } else {
            throw new Error("Execution error: LSet currently only supported for string variables");
        }
    }

    private evaluateRSetStatement(stmt: RSetStatement) {
        const val = String(this.evaluateExpression(stmt.right) || '');
        if (stmt.left.type === 'Identifier') {
            const name = (stmt.left as Identifier).name;
            const target = String(this.env.get(name) || '');
            const result = val.padStart(target.length, ' ').substring(0, target.length);
            this.env.set(name, result);
        } else {
            throw new Error("Execution error: RSet currently only supported for string variables");
        }
    }

    private evaluateErrorStatement(stmt: ErrorStatement) {
        const errNum = this.evaluateExpression(stmt.errorNumber);
        const errObj = this.env.get('err');
        if (errObj) {
            errObj.raise(errNum);
        } else {
            throw new Error(`VBA Error ${errNum}`);
        }
    }

    private evaluateAssignmentStatement(stmt: AssignmentStatement) {
        const val = this.evaluateExpression(stmt.right);

        if (stmt.left.type === 'Identifier') {
            const name = (stmt.left as Identifier).name;
            const procNameLower = name.toLowerCase();

            // Special case: assigning to current function/property name (return value)
            if (this.currentProcedureName && this.currentProcedureName.toLowerCase() === procNameLower) {
                if (this.currentProcedureType === 'function' || this.currentProcedureType === 'get') {
                    this.env.setLocally(name, val);
                    return;
                }
            }

            const proc = this.env.getProcedure(name, 'let');
            if (proc) {
                this.callProcedure(name, [val], 'let');
                return;
            }
            this.env.set(name, val);
        } else if (stmt.left.type === 'CallExpression') {
            // Array/Dictionary assignment: arr(0) = val OR dict("key") = val
            const call = stmt.left as CallExpression;
            if (call.callee.type === 'Identifier') {
                const name = (call.callee as Identifier).name;
                const lowerName = name.toLowerCase();

                if (lowerName === 'mid') {
                    // Mid(s, start, [len]) = val
                    const targetName = (call.args[0] as Identifier).name; // Must be identifier for assignment
                    const start = this.evaluateExpression(call.args[1]) as number;
                    const length = call.args.length > 2 ? this.evaluateExpression(call.args[2]) as number : -1;
                    const sourceStr = String(this.env.get(targetName) || '');
                    const replacement = String(val || '');

                    let replaceLen = length === -1 ? replacement.length : length;
                    // Mid statement rules: replacement is limited by length of source and specified length
                    replaceLen = Math.min(replaceLen, replacement.length, sourceStr.length - start + 1);

                    const head = sourceStr.substring(0, start - 1);
                    const tail = sourceStr.substring(start - 1 + replaceLen);
                    const mid = replacement.substring(0, replaceLen);

                    this.env.set(targetName, head + mid + tail);
                    return;
                }

                const target = this.env.get(name);
                
                if (Array.isArray(target)) {
                    // Support 1D or multi-dimensional array assignment arr(0, 1) = val -> arr[0][1] = val
                    let current = target;
                    for (let i = 0; i < call.args.length - 1; i++) {
                        const base = (current as any).vbaBase || 0;
                        const d = (this.evaluateExpression(call.args[i]) as number) - base;
                        if (!current[d]) {
                            current[d] = [];
                            (current[d] as any).vbaBase = base; // Inherit base for multi-dim? (VBA doesn't really have jagged with Base, but let's be consistent)
                        }
                        current = current[d];
                    }
                    const lastBase = (current as any).vbaBase || 0;
                    const lastIdx = (this.evaluateExpression(call.args[call.args.length - 1]) as number) - lastBase;
                    current[lastIdx] = val;
                } else if (target && target.__isVbaDict__) {
                    // Treat as Dictionary assignment dict("key") = val
                    const key = String(this.evaluateExpression(call.args[0]));
                    target.__map__.set(key, val);
                } else {
                    throw new Error(`Execution error: ${name} is not an array or dictionary`);
                }
            } else {
                throw new Error("Execution error: Complex left hand assignments not supported yet");
            }
        } else if (stmt.left.type === 'MemberExpression') {
            const member = stmt.left as MemberExpression;
            const obj = this.evaluateExpression(member.object);
            const propName = member.property.name.toLowerCase();
            if (obj && obj.__vbaClass__) {
                // VBA class instance: check for Property Let/Set, then fallback to field assignment
                const classDef = obj.__classDef__ as ClassDeclaration;
                const instanceEnv = obj.__instanceEnv__ as Environment;
                const setter = classDef.procedures.find(
                    p => p.isProperty && (p.propertyType === 'let' || p.propertyType === 'set') && p.name.name.toLowerCase() === propName
                );
                if (setter) {
                    this.callClassMethod(obj, setter, [val]);
                } else {
                    instanceEnv.set(propName, val);
                }
            } else if (obj && typeof obj === 'object') {
                obj[propName] = val;
            } else {
                throw new Error(`Execution error: Cannot assign property '${propName}' of undefined or primitive`);
            }
        } else if (stmt.left.type === 'ImplicitWithObjectExpression') {
            if (this.withObjectStack.length === 0) {
                throw new Error(`Execution error: '.' outside of With block`);
            }
            const obj = this.withObjectStack[this.withObjectStack.length - 1];
            const member = stmt.left as ImplicitWithObjectExpression;
            const propName = member.property.name.toLowerCase();
            if (obj && typeof obj === 'object') {
                obj[propName] = val;
            } else {
                throw new Error(`Execution error: Cannot assign property '${propName}' of undefined or primitive in With block`);
            }
        } else {
            throw new Error(`Execution error: Invalid assignment target`);
        }
    }

    private evaluateVariableDeclaration(stmt: VariableDeclaration) {
        const isStaticDecl = stmt.isStatic || this.currentProcIsStatic;
        for (const decl of stmt.declarations) {
            const varName = decl.name.name;
            const varKey = varName.toLowerCase();
            const staticKey = `${this.currentProcedureName?.toLowerCase()}:${varKey}`;

            // For static variables, restore persisted value if available
            if (isStaticDecl && this.staticVarStore.has(staticKey)) {
                this.env.set(varName, this.staticVarStore.get(staticKey));
                this.staticVarsInCurrentProc.add(varKey);
                continue;
            }

            let initialValue: any = vbaEmpty;
            // Typed numeric/string variables get VBA-spec default values
            if (decl.objectType) {
                const t = decl.objectType.toLowerCase();
                if (['integer', 'long', 'single', 'double', 'currency', 'byte'].includes(t)) {
                    initialValue = 0;
                } else if (t === 'string') {
                    initialValue = '';
                } else if (t === 'boolean') {
                    initialValue = 0; // vbaFalse
                }
            }
            if (decl.isArray) {
                if (decl.arraySize) {
                    const size = this.evaluateExpression(decl.arraySize);
                    const count = size - this.arrayBase + 1;
                    initialValue = new Array(count).fill(vbaEmpty);
                    (initialValue as any).vbaBase = this.arrayBase;
                } else {
                    initialValue = [];
                    (initialValue as any).vbaBase = this.arrayBase;
                }
            } else if (decl.isNew && decl.objectType === 'Collection') {
                initialValue = new VbaCollection();
            } else if (decl.isNew && decl.objectType && this.classDefinitions.has(decl.objectType.toLowerCase())) {
                initialValue = this.instantiateClass(decl.objectType);
            } else if (decl.objectType) {
                // Check if it's a user-defined Type
                const typeMembers = this.env.getType(decl.objectType);
                if (typeMembers) {
                    // Create an instance of the Type as a plain object with default values
                    const instance: any = {};
                    for (const member of typeMembers) {
                        const mt = member.memberType.toLowerCase();
                        if (mt === 'string') {
                            instance[member.name.toLowerCase()] = '';
                        } else {
                            instance[member.name.toLowerCase()] = 0;
                        }
                    }
                    initialValue = instance;
                }
            }
            this.env.set(varName, initialValue);
            if (isStaticDecl) {
                this.staticVarsInCurrentProc.add(varKey);
            }
        }
    }

    private evaluateTypeDeclaration(stmt: TypeDeclaration) {
        this.env.setType(stmt.name, stmt.members);
    }

    private evaluateOptionCompareStatement(stmt: OptionCompareStatement) {
        this.comparisonMode = stmt.mode;
    }

    private evaluateOptionExplicitStatement(stmt: OptionExplicitStatement) {
        // No-op at runtime. Parser should handle validation.
    }

    private evaluateOptionBaseStatement(stmt: OptionBaseStatement) {
        this.arrayBase = stmt.base;
    }

    private evaluateOptionPrivateModuleStatement(stmt: OptionPrivateModuleStatement) {
        // No-op for now. Affects visibility in multi-module environment.
    }

    private evaluateClassDeclaration(stmt: ClassDeclaration) {
        this.classDefinitions.set(stmt.name.toLowerCase(), stmt);
    }

    private instantiateClass(className: string): any {
        const classDef = this.classDefinitions.get(className.toLowerCase());
        if (!classDef) {
            throw new Error(`Execution error: Class '${className}' not found`);
        }

        // Create instance environment rooted at the global env
        const instanceEnv = new Environment(this.env);

        // Initialize public/private fields with default values
        for (const fieldDecl of classDef.fields) {
            for (const decl of fieldDecl.declarations) {
                const fieldKey = decl.name.name.toLowerCase();
                let defaultVal: any = vbaEmpty;
                const mt = (decl.objectType || '').toLowerCase();
                if (mt === 'string') defaultVal = '';
                else if (mt === 'integer' || mt === 'long' || mt === 'double' || mt === 'single') defaultVal = 0;
                instanceEnv.setLocally(decl.name.name, defaultVal);
            }
        }

        const instance: any = {
            __vbaClass__: true,
            __className__: classDef.name,
            __classDef__: classDef,
            __instanceEnv__: instanceEnv,
        };

        // Set Me in instance env pointing to the instance itself
        instanceEnv.setLocally('Me', instance);

        // Call Class_Initialize if defined
        const initProc = classDef.procedures.find(p => p.name.name.toLowerCase() === 'class_initialize');
        if (initProc) {
            this.callClassMethod(instance, initProc, []);
        }

        return instance;
    }

    private callClassMethod(instance: any, proc: ProcedureDeclaration, args: any[]): any {
        const instanceEnv = instance.__instanceEnv__ as Environment;
        const localEnv = new Environment(instanceEnv);

        // Map arguments to parameters
        for (let i = 0; i < proc.parameters.length; i++) {
            const paramName = proc.parameters[i].name;
            const argValue = i < args.length ? args[i] : vbaEmpty;
            localEnv.setLocally(paramName, argValue);
        }

        const previousEnv = this.env;
        const previousProcBody = this.currentProcBody;
        const previousProcedureName = this.currentProcedureName;
        const previousProcedureType = this.currentProcedureType;
        const previousProcIsStatic = this.currentProcIsStatic;
        const previousStaticVars = this.staticVarsInCurrentProc;
        this.env = localEnv;
        this.currentProcBody = proc.body;
        this.currentProcedureName = proc.name.name;
        this.currentProcedureType = proc.propertyType || (proc.isFunction ? 'function' : 'sub');
        this.currentProcIsStatic = false;
        this.staticVarsInCurrentProc = new Set();

        try {
            this.executeStatements(proc.body, 0);
        } catch (e: any) {
            if (e && e.type === 'Exit') {
                if (
                    (e.target === 'Function' && proc.isFunction) ||
                    (e.target === 'Sub' && !proc.isFunction && !proc.isProperty) ||
                    (e.target === 'Property' && proc.isProperty)
                ) {
                    // valid exit
                } else {
                    throw e;
                }
            } else {
                throw e;
            }
        } finally {
            this.env = previousEnv;
            this.currentProcBody = previousProcBody;
            this.currentProcedureName = previousProcedureName;
            this.currentProcedureType = previousProcedureType;
            this.currentProcIsStatic = previousProcIsStatic;
            this.staticVarsInCurrentProc = previousStaticVars;
        }

        if (proc.isFunction || proc.isProperty) {
            return localEnv.get(proc.name.name.toLowerCase());
        }
        return vbaEmpty;
    }

    private evaluateEnumDeclaration(stmt: EnumDeclaration) {
        let currentValue = 0;
        const enumObj: any = {};
        for (const member of stmt.members) {
            if (member.value) {
                currentValue = this.evaluateExpression(member.value);
            }
            const memberName = member.name.name;
            // Set directly in environment for flat access (common in VBA)
            this.env.set(memberName, currentValue);
            // Also store in enum object for EnumName.MemberName access
            enumObj[memberName] = currentValue;
            currentValue++;
        }
        // Register the Enum name itself
        this.env.set(stmt.name.name, enumObj);
    }

    private evaluateCallStatement(stmt: CallStatement) {
        this.evaluateExpression(stmt.expression);
    }

    private evaluateConstDeclaration(stmt: ConstDeclaration) {
        const value = this.evaluateExpression(stmt.value);
        this.env.set(stmt.name.name, value);
    }

    private evaluateSetStatement(stmt: SetStatement) {
        let value = this.evaluateExpression(stmt.right);

        // VBA requires Set target to be an object
        if (value !== null && typeof value !== 'object') {
            throw new Error(`Execution error: Object required`);
        }
        // If the right side evaluates to a variable name (string), resolve it
        if (typeof value === 'string' && stmt.right.type === 'Identifier') {
            value = this.env.get(value);
        }

        if (stmt.left.type === 'Identifier') {
            const name = (stmt.left as Identifier).name;
            const procNameLower = name.toLowerCase();

            // Special case: assigning to current function/property name (return value)
            if (this.currentProcedureName && this.currentProcedureName.toLowerCase() === procNameLower) {
                if (this.currentProcedureType === 'function' || this.currentProcedureType === 'get') {
                    this.env.setLocally(name, value);
                    return;
                }
            }

            const proc = this.env.getProcedure(name, 'set');
            if (proc) {
                this.callProcedure(name, [value], 'set');
                return;
            }
            this.env.set(name, value);
        } else {
            throw new Error(`Execution error: Unsupported Set target ${stmt.left.type}`);
        }
    }

    private evaluateResumeStatement(stmt: ResumeStatement) {
        if (!this.isInErrorHandler && this.errorHandlingMode !== 'ResumeNext') {
            // Technically Resume is only for error handlers, but let's be strict
            throw new Error('Execution error: Resume without active error handler');
        }
        throw { type: 'Resume', target: stmt.target };
    }

    private evaluateOnErrorStatement(stmt: OnErrorStatement) {
        const label = (stmt.label || '').toLowerCase().trim();
        if (label === '0') {
            // On Error GoTo 0 - disable error handling
            this.errorHandlerLabel = null;
            this.errorHandlingMode = 'None';
            this.isInErrorHandler = false; // Reset error handler state too
        } else if (label === 'resume next') {
            // On Error Resume Next
            this.errorHandlerLabel = null;
            this.errorHandlingMode = 'ResumeNext';
        } else {
            // On Error GoTo <label>
            this.errorHandlerLabel = stmt.label;
            this.errorHandlingMode = 'Label';
        }
    }

    private evaluateAttributeStatement(stmt: AttributeStatement) {
        // No-op: ignore Attributes
    }

    private evaluateDeclareStatement(stmt: DeclareStatement) {
        const name = stmt.name.toLowerCase();
        this.env.set(name, (...args: any[]) => {
            this.onPrint(`[DECLARE STUB] Calling ${stmt.isSub ? 'Sub' : 'Function'} ${stmt.name} from "${stmt.libName}" (Alias: ${stmt.aliasName || 'N/A'})`);
            return 0; // Dummy return
        });
    }

    private createExternalObject(progId: string): any {
        const id = progId.toLowerCase();
        if (id === 'scripting.dictionary') {
            const dict = new Map<any, any>();
            return {
                __isVbaDict__: true,
                __map__: dict,
                add: (k: any, v: any) => dict.set(k, v),
                exists: (k: any) => dict.has(k) ? vbaTrue : vbaFalse,
                remove: (k: any) => dict.delete(k),
                removeall: () => dict.clear(),
                count: () => dict.size,
                keys: () => Array.from(dict.keys()),
                items: () => Array.from(dict.values()),
                item: (k: any, v?: any) => {
                    if (v !== undefined) {
                        dict.set(k, v);
                    }
                    return dict.get(k);
                }
            };
        } else if (id === 'scripting.filesystemobject') {
            return {
                __isVbaFso__: true,
                fileexists: (path: string) => vbaFalse,
                folderexists: (path: string) => vbaFalse,
                createtextfile: (path: string) => ({
                    write: (s: string) => this.onPrint(`[FSO Write] ${path}: ${s}`),
                    writeline: (s: string) => this.onPrint(`[FSO WriteLine] ${path}: ${s}`),
                    close: () => {}
                }),
                opentextfile: (path: string) => ({
                    readall: () => "",
                    close: () => {}
                })
            };
        }
        throw new Error(`Execution error: Unsupported CreateObject '${progId}'`);
    }

    // Execute a sequence of statements starting from startIndex, with error handling support
    private executeStatements(body: Statement[], startIndex: number) {
        let i = startIndex;
        while (i < body.length) {
            const stmt = body[i];
            try {
                this.evaluateStatement(stmt);
                i++;
            } catch (e: any) {
                if (e && (e.type === 'Exit' || e.type === 'Terminate')) {
                    throw e;
                }

                if (e && e.type === 'GoTo') {
                    const labelName = e.label.toLowerCase();
                    const labelIndex = body.findIndex(s =>
                        s.type === 'LabelStatement' &&
                        (s as any).label.toLowerCase() === labelName
                    );
                    if (labelIndex >= 0) {
                        i = labelIndex;
                        continue;
                    }
                    throw e;
                }

                if (e && e.type === 'GoSub') {
                    const labelName = e.label.toLowerCase();
                    const labelIndex = body.findIndex(s =>
                        s.type === 'LabelStatement' &&
                        (s as any).label.toLowerCase() === labelName
                    );
                    if (labelIndex >= 0) {
                        this.gosubStack.push(i);
                        i = labelIndex;
                        continue;
                    }
                    throw e;
                }

                if (e && e.type === 'Return') {
                    if (this.gosubStack.length === 0) {
                        throw new Error('Execution error: Return without GoSub');
                    }
                    i = this.gosubStack.pop()! + 1;
                    continue;
                }

                if (e && e.type === 'Resume') {
                    this.isInErrorHandler = false;
                    if (e.target === '' || e.target === '0') {
                        i = this.lastErrorIndex;
                    } else if (e.target.toLowerCase() === 'next') {
                        i = this.lastErrorIndex + 1;
                    } else {
                        const labelName = e.target.toLowerCase();
                        const labelIndex = body.findIndex(s =>
                            s.type === 'LabelStatement' &&
                            (s as any).label.toLowerCase() === labelName
                        );
                        if (labelIndex >= 0) {
                            i = labelIndex;
                        } else {
                            throw new Error(`Execution error: Label '${e.target}' not found for Resume`);
                        }
                    }
                    // Clear Err object on Resume
                    const errObj = this.env.get('err');
                    if (errObj) { errObj.number = 0; errObj.description = ''; }
                    continue;
                }

                // Actual Error Handling
                if (this.isInErrorHandler) {
                    // Error inside error handler -> bubble up
                    this.isInErrorHandler = false;
                    throw e;
                }

                if (this.errorHandlingMode === 'ResumeNext') {
                    this.lastErrorIndex = i;
                    // Populate Err object
                    const errObj = this.env.get('err');
                    if (errObj) {
                        if (e && e.type === 'VbaError') {
                            errObj.number = e.number;
                            errObj.description = e.message;
                        } else {
                            errObj.number = 1000;
                            errObj.description = e.message || String(e);
                        }
                    }
                    i++;
                    continue;
                } else if (this.errorHandlingMode === 'Label' && this.errorHandlerLabel) {
                    const labelIndex = body.findIndex(s =>
                        s.type === 'LabelStatement' &&
                        (s as any).label.toLowerCase() === this.errorHandlerLabel!.toLowerCase()
                    );
                    if (labelIndex >= 0) {
                        this.isInErrorHandler = true;
                        this.lastErrorIndex = i;
                        // Populate Err object
                        const errObj = this.env.get('err');
                        if (errObj) {
                            if (e && e.type === 'VbaError') {
                                errObj.number = e.number;
                                errObj.description = e.message;
                            } else {
                                errObj.number = 1000;
                                errObj.description = e.message || String(e);
                            }
                        }
                        i = labelIndex;
                        continue;
                    }
                }

                // Default behavior: throw
                throw e;
            }
        }
    }

    private evaluateEraseStatement(stmt: EraseStatement) {
        this.env.set(stmt.name.name, []);
    }

    private evaluateReDimStatement(stmt: ReDimStatement) {
        // Evaluate bounds (just size for 1D for now)
        if (stmt.bounds.length > 0) {
            const size = this.evaluateExpression(stmt.bounds[stmt.bounds.length - 1]); // naive: takes last bound as size
            const count = size - this.arrayBase + 1;
            const arr = new Array(count).fill(0); // VBA numeric arrays default to 0
            (arr as any).vbaBase = this.arrayBase;
            this.env.set(stmt.name.name, arr);
        }
    }

    private evaluateExitStatement(stmt: ExitStatement) {
        throw { type: 'Exit', target: stmt.exitType };
    }

    private evaluateExpression(expr: Expression): any {
        switch (expr.type) {
            case 'NumberLiteral':
                return (expr as NumberLiteral).value;
            case 'StringLiteral':
                return (expr as StringLiteral).value;
            case 'DateLiteral':
                return this.evaluateDateLiteral(expr as DateLiteral);
            case 'Identifier':
                const idName = (expr as Identifier).name;
                const v = this.env.get(idName);
                if (typeof v === 'function' && (v as any).__vbaAutoCall__) {
                    return v();
                }
                const p = this.env.getProcedure(idName);
                if (p) {
                    // Only auto-call if it's a Function (not Sub) and has 0 required arguments
                    // For now, assume it's okay to call
                    return this.callProcedure(idName, []);
                }
                return v;
            case 'CallExpression':
                return this.evaluateCallExpression(expr as CallExpression);
            case 'MemberExpression':
                return this.evaluateMemberExpression(expr as MemberExpression);
            case 'DictionaryAccessExpression':
                return this.evaluateDictionaryAccessExpression(expr as DictionaryAccessExpression);
            case 'TypeOfIsExpression':
                return this.evaluateTypeOfIsExpression(expr as TypeOfIsExpression);
            case 'UnaryExpression':
                return this.evaluateUnaryExpression(expr as UnaryExpression);
            case 'BinaryExpression':
                return this.evaluateBinaryExpression(expr as BinaryExpression);
            case 'ImplicitWithObjectExpression':
                return this.evaluateImplicitWithObjectExpression(expr as ImplicitWithObjectExpression);
            case 'NewExpression':
                return this.instantiateClass((expr as NewExpression).className);
            default:
                throw new Error(`Execution error: Unknown expression type ${expr.type}`);
        }
    }

    private evaluateDateLiteral(expr: DateLiteral): any {
        const d = new Date(expr.value);
        if (isNaN(d.getTime())) {
            // Try fallback or simple parser if standard Date fails
            // VBA #mm/dd/yyyy# or #yyyy-mm-dd#
            // For now, let's assume JS Date can handle common formats
            throw new Error(`Execution error: Invalid date literal #${expr.value}#`);
        }
        // Excel/VBA serial number: days since 1899-12-30
        const baseDate = new Date(1899, 11, 30); // Dec 30, 1899
        const diff = d.getTime() - baseDate.getTime();
        const serial = diff / (24 * 60 * 60 * 1000);
        return new VbaDate(serial);
    }

    private evaluateCallExpression(expr: CallExpression): any {
        if (expr.callee.type === 'Identifier') {
            const name = (expr.callee as Identifier).name;
            const proc = this.env.getProcedure(name);

            if (proc) {
                // Cross-module Private access check
                if (
                    proc.scope === 'private' &&
                    proc.moduleName !== undefined &&
                    proc.moduleName !== '' &&
                    proc.moduleName !== this.executingModuleName
                ) {
                    throw new Error(
                        `Execution error: Cannot call Private procedure '${proc.name.name}' ` +
                        `from module '${this.executingModuleName || '(top-level)'}' ` +
                        `(defined in '${proc.moduleName}')`
                    );
                }

                // Procedure call (Function/Sub)
                const localEnv = new Environment(this.env);

                // Map arguments to parameters
                const byRefArgs: { paramName: string, identifierName: string }[] = [];
                for (let i = 0; i < proc.parameters.length; i++) {
                    const argVal = i < expr.args.length ? this.evaluateExpression(expr.args[i]) : 0;
                    localEnv.set(proc.parameters[i].name, argVal);

                    if (i < expr.args.length && expr.args[i].type === 'Identifier') {
                        // VBA Default is ByRef. If it is NOT explicitly ByVal, it is ByRef
                        if (!proc.parameters[i].isByVal) {
                            byRefArgs.push({
                                paramName: proc.parameters[i].name,
                                identifierName: (expr.args[i] as Identifier).name
                            });
                        }
                    }
                }

                if (proc.isFunction) {
                    // Implicit variable for function return value
                    localEnv.setLocally(proc.name.name, vbaEmpty);
                }

                const previousEnv = this.env;
                const previousErrorHandler = this.errorHandlerLabel;
                const previousProcBody = this.currentProcBody;
                const previousExecutingModule = this.executingModuleName;
                this.env = localEnv;
                this.errorHandlerLabel = null;
                this.currentProcBody = proc.body;
                this.executingModuleName = proc.moduleName ?? '';

                try {
                    this.executeStatements(proc.body, 0);
                } catch (e: any) {
                    if (e && e.type === 'Exit' && (e.target === 'Sub' || e.target === 'Function')) {
                        // Exit the procedure cleanly
                    } else {
                        throw e;
                    }
                }

                this.env = previousEnv; // Restore scope
                this.errorHandlerLabel = previousErrorHandler;
                this.currentProcBody = previousProcBody;
                this.executingModuleName = previousExecutingModule;

                // Synchronize ByRef arguments back to caller scope
                for (const ref of byRefArgs) {
                    const updatedVal = localEnv.get(ref.paramName);
                    this.env.setLocally(ref.identifierName, updatedVal);
                }

                if (proc.isFunction) {
                    return localEnv.get(proc.name.name);
                }
                return undefined;
            } else {
                // Might be an array access or built-in function
                const variable = this.env.get(name);
                if (typeof variable === 'function') {
                    const argsVals = expr.args.map(a => this.evaluateExpression(a));
                    return variable(...argsVals);
                } else if (Array.isArray(variable)) {
                    if (expr.args.length === 0) throw new Error(`Execution error: Missing index for array ${name}`);
                    // Support multi-dimensional array lookup arr(0, 1) -> arr[0][1]
                    let current = variable;
                    for (let i = 0; i < expr.args.length; i++) {
                        if (!current) return vbaEmpty; // Out of bounds or jagged array
                        const base = (current as any).vbaBase || 0;
                        const idx = (this.evaluateExpression(expr.args[i]) as number) - base;
                        current = current[idx];
                    }
                    if (current === undefined) return vbaEmpty;
                    return current;
                } else if (variable && variable.__isVbaDict__) {
                    // Dictionary read: dict("key")
                    if (expr.args.length === 0) throw new Error(`Execution error: Missing key for dictionary ${name}`);
                    const key = this.evaluateExpression(expr.args[0]);
                    return variable.__map__.get(key);
                }
                throw new Error(`Execution error: Cannot call unknown procedure or index unknown array '${name}'`);
            }
        } else if (expr.callee.type === 'MemberExpression' || expr.callee.type === 'ImplicitWithObjectExpression') {
            let obj: any;
            let methodNameOriginal: string;

            if (expr.callee.type === 'MemberExpression') {
                const member = expr.callee as MemberExpression;
                obj = this.evaluateExpression(member.object);
                methodNameOriginal = member.property.name;
            } else {
                if (this.withObjectStack.length === 0) {
                    throw new Error(`Execution error: '.' outside of With block`);
                }
                obj = this.withObjectStack[this.withObjectStack.length - 1];
                methodNameOriginal = (expr.callee as ImplicitWithObjectExpression).property.name;
            }

            const methodNameLower = methodNameOriginal.toLowerCase();

            // VBA class instance method call
            if (obj && obj.__vbaClass__) {
                const classDef = obj.__classDef__ as ClassDeclaration;
                const proc = classDef.procedures.find(p => p.name.name.toLowerCase() === methodNameLower);
                if (proc) {
                    const argsVals = expr.args.map(a => this.evaluateExpression(a));
                    return this.callClassMethod(obj, proc, argsVals);
                }
                throw new Error(`Execution error: Class '${obj.__className__}' has no method '${methodNameOriginal}'`);
            }

            if (obj) {
                // Try case-insensitive lookup first, then fallback to original casing
                let targetMethod = obj[methodNameLower];

                if (typeof targetMethod !== 'function') {
                    // Search object keys for case-insensitive match (for JS proxies/objects)
                    const keys = Object.keys(obj);
                    // If proxy, Object.keys might not work perfectly, so also check original
                    if (typeof obj[methodNameOriginal] === 'function') {
                        targetMethod = obj[methodNameOriginal];
                    } else {
                        const match = keys.find(k => k.toLowerCase() === methodNameLower);
                        if (match) targetMethod = obj[match];
                    }
                }

                if (typeof targetMethod === 'function') {
                    const argsVals = expr.args.map(a => this.evaluateExpression(a));
                    return targetMethod.apply(obj, argsVals);
                }
            }
            throw new Error(`Execution error: Object does not support property or method '${methodNameOriginal}'`);
        }


        // Generic fallback for calling result of an expression: (expr)(args)
        // e.g. Array(1, 2)(0)
        const target = this.evaluateExpression(expr.callee);
        if (Array.isArray(target)) {
            if (expr.args.length === 0) throw new Error(`Execution error: Missing index for array access`);
            let current = target;
            for (let i = 0; i < expr.args.length; i++) {
                if (!current) return vbaEmpty;
                const idx = this.evaluateExpression(expr.args[i]);
                current = current[idx];
            }
            return current === undefined ? vbaEmpty : current;
        } else if (target && target.__isVbaDict__) {
            if (expr.args.length === 0) throw new Error(`Execution error: Missing key for dictionary access`);
            const key = this.evaluateExpression(expr.args[0]);
            return target.__map__.get(key);
        } else if (typeof target === 'function') {
            const argsVals = expr.args.map(a => this.evaluateExpression(a));
            return target(...argsVals);
        }

        throw new Error(`Execution error: Unsupported call expression or target is not callable`);
    }

    private evaluateDictionaryAccessExpression(expr: DictionaryAccessExpression): any {
        const obj = this.evaluateExpression(expr.object);
        const property = expr.property.name;

        if (obj && obj.__isVbaDict__) {
            // VBA bang (!) access is essentially string key lookup: dict!Key -> dict("Key")
            return obj.__map__.get(property);
        }

        // Fallback or error (some objects might support ! besides Dictionary, but we only have Dictionary for now)
        throw new Error(`Execution error: Object does not support '!' access for property '${property}'`);
    }

    private evaluateTypeOfIsExpression(expr: TypeOfIsExpression): any {
        const obj = this.evaluateExpression(expr.expression);
        const typeName = expr.typeName.toLowerCase();

        if (obj === null || obj === undefined || typeof obj !== 'object') return vbaFalse;

        // Check for built-in types
        if (typeName === 'object') return vbaTrue; // Everything that reaches here is an object
        if (typeName === 'dictionary' && obj.__isVbaDict__) return vbaTrue;
        if (typeName === 'collection' && obj.__isVbaCollection__) return vbaTrue;

        // User defined types or classes (if we store metadata)
        if (obj.__vbaTypeName__ && obj.__vbaTypeName__.toLowerCase() === typeName) return vbaTrue;

        return vbaFalse;
    }

    private evaluateUnaryExpression(expr: UnaryExpression): any {
        const argument = this.evaluateExpression(expr.argument);
        switch (expr.operator.toLowerCase()) {
            case 'not':
                const res = ~argument;
                return (argument instanceof VbaBoolean) ? new VbaBoolean(res as any) : res;
            case '-':
                return -argument;
            case '+':
                return +argument;
            default:
                throw new Error(`Execution error: Unknown unary operator ${expr.operator}`);
        }
    }

    private evaluateMemberExpression(expr: MemberExpression): any {
        const obj = this.evaluateExpression(expr.object);
        const propName = expr.property.name.toLowerCase();

        // VBA class instance: look up field in instance environment or invoke Property Get
        if (obj && obj.__vbaClass__) {
            const classDef = obj.__classDef__ as ClassDeclaration;
            const instanceEnv = obj.__instanceEnv__ as Environment;

            // Check for Property Get
            const getter = classDef.procedures.find(
                p => p.isProperty && p.propertyType === 'get' && p.name.name.toLowerCase() === propName
            );
            if (getter) {
                return this.callClassMethod(obj, getter, []);
            }

            // No-arg Sub/Function access without parens (rare, treat as property-like call)
            const method = classDef.procedures.find(
                p => !p.isProperty && p.name.name.toLowerCase() === propName && p.parameters.length === 0 && p.isFunction
            );
            if (method) {
                return this.callClassMethod(obj, method, []);
            }

            // Field access
            return instanceEnv.get(propName);
        }

        if (obj && propName in obj) {
            const val = obj[propName];
            // Auto-call only zero-arg functions (VBA property/method without parens like col.Count, ws.Add)
            // Functions requiring args (like Worksheets(name)) are returned as references
            if (typeof val === 'function' && val.length === 0) {
                return val.call(obj);
            }
            return val;
        }
        // Case-insensitive fallback for JS objects
        if (obj && typeof obj === 'object') {
            const keys = Object.keys(obj);
            const match = keys.find(k => k.toLowerCase() === propName);
            if (match) {
                const val = obj[match];
                if (typeof val === 'function' && val.length === 0) {
                    return val.call(obj);
                }
                return val;
            }
        }
        throw new Error(`Execution error: Method or property not found '${propName}'`);
    }

    private evaluateBinaryExpression(expr: BinaryExpression): any {
        let leftVal = this.evaluateExpression(expr.left);
        let rightVal = this.evaluateExpression(expr.right);

        // Normalize booleans to VBA integers (-1, 0)
        if (leftVal === true) leftVal = vbaTrue;
        if (leftVal === false) leftVal = vbaFalse;
        if (rightVal === true) rightVal = vbaTrue;
        if (rightVal === false) rightVal = vbaFalse;

        switch (expr.operator.toLowerCase()) {
            case '+': 
                const sum = leftVal + rightVal;
                return (leftVal instanceof VbaDate || rightVal instanceof VbaDate) ? new VbaDate(sum) : sum;
            case '&': return String(leftVal) + String(rightVal);
            case '-': 
                const diff = leftVal - rightVal;
                if (leftVal instanceof VbaDate && rightVal instanceof VbaDate) return diff; // Date - Date = Number
                return (leftVal instanceof VbaDate) ? new VbaDate(diff) : diff;
            case '*': return leftVal * rightVal;
            case '/': 
                if (rightVal === 0) throw { type: 'VbaError', number: 11, message: 'Division by zero' };
                return leftVal / rightVal;
            case '\\': 
                if (rightVal === 0) throw { type: 'VbaError', number: 11, message: 'Division by zero' };
                return Math.floor(leftVal / rightVal);
            case 'mod': 
                if (rightVal === 0) throw { type: 'VbaError', number: 11, message: 'Division by zero' };
                return leftVal % rightVal;
            case '^': return Math.pow(leftVal, rightVal);
            case '=': 
                if (leftVal instanceof VbaErrorValue && rightVal instanceof VbaErrorValue) {
                    return leftVal.code === rightVal.code ? vbaTrue : vbaFalse;
                }
                if (typeof leftVal === 'string' && typeof rightVal === 'string' && this.comparisonMode === 'Text') {
                    return leftVal.toLowerCase() === rightVal.toLowerCase() ? vbaTrue : vbaFalse;
                }
                return leftVal === rightVal ? vbaTrue : vbaFalse;
            case '<>': 
                if (leftVal instanceof VbaErrorValue && rightVal instanceof VbaErrorValue) {
                    return leftVal.code !== rightVal.code ? vbaTrue : vbaFalse;
                }
                if (typeof leftVal === 'string' && typeof rightVal === 'string' && this.comparisonMode === 'Text') {
                    return leftVal.toLowerCase() !== rightVal.toLowerCase() ? vbaTrue : vbaFalse;
                }
                return leftVal !== rightVal ? vbaTrue : vbaFalse;
            case '<': 
                if (typeof leftVal === 'string' && typeof rightVal === 'string' && this.comparisonMode === 'Text') {
                    return leftVal.toLowerCase() < rightVal.toLowerCase() ? vbaTrue : vbaFalse;
                }
                return leftVal < rightVal ? vbaTrue : vbaFalse;
            case '>': 
                if (typeof leftVal === 'string' && typeof rightVal === 'string' && this.comparisonMode === 'Text') {
                    return leftVal.toLowerCase() > rightVal.toLowerCase() ? vbaTrue : vbaFalse;
                }
                return leftVal > rightVal ? vbaTrue : vbaFalse;
            case '<=': 
                if (typeof leftVal === 'string' && typeof rightVal === 'string' && this.comparisonMode === 'Text') {
                    return leftVal.toLowerCase() <= rightVal.toLowerCase() ? vbaTrue : vbaFalse;
                }
                return leftVal <= rightVal ? vbaTrue : vbaFalse;
            case '>=': 
                if (typeof leftVal === 'string' && typeof rightVal === 'string' && this.comparisonMode === 'Text') {
                    return leftVal.toLowerCase() >= rightVal.toLowerCase() ? vbaTrue : vbaFalse;
                }
                return leftVal >= rightVal ? vbaTrue : vbaFalse;
            case 'is': return leftVal === rightVal ? vbaTrue : vbaFalse;
            case 'like': return this.evaluateLike(leftVal, rightVal) ? vbaTrue : vbaFalse;
            case 'and':
                const andRes = leftVal & rightVal;
                return (leftVal instanceof VbaBoolean && rightVal instanceof VbaBoolean) ? new VbaBoolean(andRes as any) : andRes;
            case 'or':
                const orRes = leftVal | rightVal;
                return (leftVal instanceof VbaBoolean && rightVal instanceof VbaBoolean) ? new VbaBoolean(orRes as any) : orRes;
            case 'xor':
                const xorRes = leftVal ^ rightVal;
                return (leftVal instanceof VbaBoolean && rightVal instanceof VbaBoolean) ? new VbaBoolean(xorRes as any) : xorRes;
            case 'eqv':
                const eqvRes = ~(leftVal ^ rightVal);
                return (leftVal instanceof VbaBoolean && rightVal instanceof VbaBoolean) ? new VbaBoolean(eqvRes as any) : eqvRes;
            case 'imp':
                const impRes = (~leftVal) | rightVal;
                return (leftVal instanceof VbaBoolean && rightVal instanceof VbaBoolean) ? new VbaBoolean(impRes as any) : impRes;
            default:
                throw new Error(`Execution error: Unknown operator ${expr.operator}`);
        }
    }

    private evaluateImplicitWithObjectExpression(expr: ImplicitWithObjectExpression): any {
        if (this.withObjectStack.length === 0) {
            throw new Error(`Execution error: '.' outside of With block`);
        }
        const obj = this.withObjectStack[this.withObjectStack.length - 1];
        const propName = expr.property.name.toLowerCase();
        if (obj && typeof obj === 'object') {
            // Check for direct property access
            if (Object.prototype.hasOwnProperty.call(obj, propName)) {
                const val = obj[propName];
                if (typeof val === 'function') {
                    return val.bind(obj);
                }
                return val;
            }
            // Case-insensitive fallback
            const keys = Object.keys(obj);
            const match = keys.find(k => k.toLowerCase() === propName);
            if (match) {
                const val = obj[match];
                if (typeof val === 'function') {
                    return val.bind(obj);
                }
                return val;
            }
        }
        throw new Error(`Execution error: Cannot access property '${propName}' of undefined or primitive in With block`);
    }

    private evaluateLike(text: any, pattern: any): boolean {
        const textStr = String(text);
        const patternStr = String(pattern);
        
        // Convert VBA Like pattern to Regex
        let regexStr = '^';
        let i = 0;
        while (i < patternStr.length) {
            const char = patternStr[i];
            if (char === '*') {
                regexStr += '.*';
            } else if (char === '?') {
                regexStr += '.';
            } else if (char === '#') {
                regexStr += '\\d';
            } else if (char === '[') {
                regexStr += '[';
                i++;
                let negate = false;
                if (i < patternStr.length && patternStr[i] === '!') {
                    regexStr += '^';
                    negate = true;
                    i++;
                }
                while (i < patternStr.length && patternStr[i] !== ']') {
                    const charInList = patternStr[i];
                    if (charInList === '-' && i > 0 && i < patternStr.length - 1) {
                         // Range - keep as is in regex
                         regexStr += '-';
                    } else if ('\\^$-.[]'.indexOf(charInList) !== -1) {
                        regexStr += '\\' + charInList;
                    } else {
                        regexStr += charInList;
                    }
                    i++;
                }
                regexStr += ']';
            } else {
                // Escape regex special characters
                if ('\\.[]{}()^$+*?|'.indexOf(char) !== -1) {
                    regexStr += '\\' + char;
                } else {
                    regexStr += char;
                }
            }
            i++;
        }
        regexStr += '$';

        try {
            const flags = this.comparisonMode === 'Text' ? 'i' : '';
            const regex = new RegExp(regexStr, flags);
            return regex.test(textStr);
        } catch (e) {
            return false;
        }
    }

    private isTrue(val: any): boolean {
        if (val === undefined || val === null || val === vbaNull) return false;
        if (val instanceof VbaBoolean) return val.value !== 0;
        if (typeof val === 'number') return val !== 0;
        if (typeof val === 'boolean') return val;
        return !!val;
    }

    private formatDate(d: Date, pattern: string): string {
        let result = pattern;
        // Replace patterns from longest to shortest to avoid partial matches
        const replacements: [string | RegExp, string][] = [
            ['yyyy', String(d.getFullYear())],
            ['yy', String(d.getFullYear()).slice(-2)],
            ['mmmm', d.toLocaleString('en-US', { month: 'long' })],
            ['mmm', d.toLocaleString('en-US', { month: 'short' })],
            ['mm', String(d.getMonth() + 1).padStart(2, '0')],
            ['m', String(d.getMonth() + 1)],
            ['dddd', d.toLocaleString('en-US', { weekday: 'long' })],
            ['ddd', d.toLocaleString('en-US', { weekday: 'short' })],
            ['dd', String(d.getDate()).padStart(2, '0')],
            ['d', String(d.getDate())],
            ['hh', String(d.getHours()).padStart(2, '0')],
            ['h', String(d.getHours())],
            ['nn', String(d.getMinutes()).padStart(2, '0')],
            ['n', String(d.getMinutes())],
            ['ss', String(d.getSeconds()).padStart(2, '0')],
            ['s', String(d.getSeconds())],
        ];

        for (const [p, r] of replacements) {
            result = result.replace(new RegExp(String(p), 'g'), r);
        }
        return result;
    }

    private formatNumber(n: number, pattern: string): string {
        // Very basic implementation: handle "0.00", "0", "#,##0" etc.
        if (pattern.includes('.')) {
            const decimalPlaces = (pattern.split('.')[1] || '').length;
            return n.toLocaleString(undefined, {
                minimumFractionDigits: decimalPlaces,
                maximumFractionDigits: decimalPlaces,
                useGrouping: pattern.includes(',')
            });
        }
        return n.toLocaleString(undefined, {
            minimumFractionDigits: 0,
            maximumFractionDigits: 0,
            useGrouping: pattern.includes(',')
        });
    }
}
