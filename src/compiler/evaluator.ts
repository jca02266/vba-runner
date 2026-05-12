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
    ParenthesizedExpression,
    ConstDeclaration,
    SetStatement,
    EraseStatement,
    ReDimStatement,
    ArrayBound,
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
    ResumeStatement,
    OpenStatement,
    CloseStatement,
    PrintStatement,
    LineInputStatement,
    PutStatement,
    KillStatement,
    WriteStatement,
    InputStatement,
    GetStatement,
    SeekStatement,
    ResetStatement,
    AttributeStatement,
    DeclareStatement,
} from './parser';
import * as fs from 'fs';
import * as path from 'path';
import { Lexer, TokenType } from './lexer';
import { SandboxPath } from './sandbox';

export class VbaBoolean {
    public readonly __isVbaBoolean__ = true;
    constructor(public value: -1 | 0) {}
    valueOf() { return this.value; }
    toString() { return this.value === -1 ? 'True' : 'False'; }
}

// VBA date serial: days since 1899-12-30 (VBA epoch)
// Using UTC to avoid timezone-dependent serial numbers
const VBA_EPOCH = new Date(Date.UTC(1899, 11, 30));
const MS_PER_DAY = 86400000;

export const toVbaDate = (d: Date): number =>
    (d.getTime() - VBA_EPOCH.getTime()) / MS_PER_DAY;

export const fromVbaDate = (serial: number): Date => {
    const ms = Math.round(serial * MS_PER_DAY);
    return new Date(VBA_EPOCH.getTime() + ms);
};

export const parseVbaDate = (val: any): Date => {
    if (val === null || val === undefined) throw new Error('Execution error: Invalid date');
    if (val instanceof VbaDate || (val && val.__isVbaDate__)) return fromVbaDate(val.value);
    if (typeof val === 'number') return fromVbaDate(val);
    
    let str = String(val);
    // If it looks like a time-only string, prepend a base date for JS parser
    if (/^\d{1,2}:\d{1,2}(:\d{1,2})?$/.test(str)) {
        str = "1899/12/30 " + str;
    }
    
    const d = new Date(str);
    if (isNaN(d.getTime())) throw new Error(`Execution error: Type mismatch: '${val}'`);
    
    // Treat the parsed local time components as UTC to maintain timezone independence
    return new Date(Date.UTC(
        d.getFullYear(), d.getMonth(), d.getDate(),
        d.getHours(), d.getMinutes(), d.getSeconds(), d.getMilliseconds()
    ));
};

export class VbaDate {
    public __isVbaDate__ = true;
    constructor(public value: number) {}
    valueOf() { return this.value; }
    toString() {
        const d = fromVbaDate(this.value);
        if (Math.abs(this.value - Math.round(this.value)) < 1e-10) {
            return d.toLocaleDateString();
        }
        return d.toLocaleString();
    }
}

export class VbaDecimal {
    constructor(public value: number) {}
    toString() { return String(this.value); }
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

    public get items(): any[] {
        return this._items.map(i => i.value);
    }

    public remove(id: any) {
        const idx = this.findIndex(id);
        this._items.splice(idx - 1, 1);
    }
}

export class VbaErrObject {
    public number: number = 0;
    public source: string = "";
    public description: string = "";
    public helpfile: string = "";
    public helpcontext: number = 0;
    public lastdllerror: number = 0;

    public clear() {
        this.number = 0;
        this.source = "";
        this.description = "";
        this.helpfile = "";
        this.helpcontext = 0;
        this.lastdllerror = 0;
    }

    public raise(number: number, source?: any, description?: any, helpfile?: any, helpcontext?: any) {
        this.number = number;
        if (source !== undefined && source !== vbaEmpty && source !== null) this.source = String(source);
        if (description !== undefined && description !== vbaEmpty && description !== null) this.description = String(description);
        if (helpfile !== undefined && helpfile !== vbaEmpty && helpfile !== null) this.helpfile = String(helpfile);
        if (helpcontext !== undefined && helpcontext !== vbaEmpty && helpcontext !== null) this.helpcontext = Number(helpcontext);
        
        throw { type: 'VbaError', number: this.number, message: this.description };
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
    private withEventsVariables: Set<string> = new Set();
    private constantVariables: Set<string> = new Set();
    public enclosing?: Environment;

    constructor(enclosing?: Environment) {
        this.enclosing = enclosing;
    }

    set(name: string, value: any) {
        const key = name.toLowerCase();
        if (this.isConstant(key)) {
            throw new Error(`Execution error: Assignment to constant not allowed: '${name}'`);
        }
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

    setWithEvents(name: string) {
        this.withEventsVariables.add(name.toLowerCase());
    }

    isWithEvents(name: string): boolean {
        const key = name.toLowerCase();
        if (this.withEventsVariables.has(key)) return true;
        if (this.enclosing) return this.enclosing.isWithEvents(name);
        return false;
    }

    setConstant(name: string, value: any) {
        const key = name.toLowerCase();
        this.variables.set(key, value);
        this.constantVariables.add(key);
    }

    isConstant(name: string): boolean {
        const key = name.toLowerCase();
        if (this.constantVariables.has(key)) return true;
        if (this.enclosing) return this.enclosing.isConstant(name);
        return false;
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
    private fileHandles: Map<number, {
        fd: number,
        mode: 'Input' | 'Output' | 'Append' | 'Random' | 'Binary',
        path: string,
        buffer?: Buffer,
        pos?: number
    }> = new Map();
    private sandbox: SandboxPath = new SandboxPath();
    private onPrint: PrintCallback;
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
    private errObj: VbaErrObject = new VbaErrObject();
    private classDefinitions: Map<string, ClassDeclaration> = new Map();
    private comparisonMode: 'Binary' | 'Text' = 'Binary';
    private errorHandlerLabel: string | null = null;
    private errorHandlingMode: 'None' | 'ResumeNext' | 'GoTo' = 'None';
    private isInErrorHandler: boolean = false;
    private lastErrorIndex: number | null = null;
    private virtualRegistry: { [app: string]: { [section: string]: { [key: string]: string } } } = {};
    private dirIterator: string[] | null = null;
    private dirIndex: number = 0;

    constructor(onPrint: PrintCallback, config: { sandboxRoot?: string, env?: Record<string, string> } = {}) {
        this.env = new Environment();
        this.onPrint = onPrint;
        this.sandbox = new SandboxPath(config.sandboxRoot, config.env);
        this.registerStandardLibrary();
    }

    private registerStandardLibrary() {
        // --- Core Objects ---
        this.env.set('debug', {
            print: (...args: any[]) => this.onPrint(args.map(a => this.toDisplayString(a)).join(' ')),
            assert: (condition: any) => {
                if (!this.isTrue(condition)) {
                    throw new Error('Assertion failed');
                }
            }
        });
        this.env.set('err', this.errObj);
        this.env.set('application', {
            wait: (time: any) => {
                this.onPrint(`[APPLICATION.WAIT] ${time}`);
                return vbaTrue;
            },
            statusbar: "",
            displayalerts: true,
            screenupdating: true
        });

        // --- Information Module ---
        this.env.set('isempty', (val: any) => (val === undefined || val === null || val === vbaEmpty) ? vbaTrue : vbaFalse);
        this.env.set('ismissing', (val: any) => val === vbaMissing ? vbaTrue : vbaFalse);
        this.env.set('isnumeric', (val: any) => {
            if (val === vbaNull) return vbaFalse;
            if (val === vbaEmpty || val === undefined) return vbaTrue;
            if (typeof val === 'number' || typeof val === 'bigint' || val instanceof VbaDecimal || val instanceof VbaBoolean || val instanceof VbaDate) return vbaTrue;
            if (typeof val === 'string') {
                const s = val.trim();
                if (s === "") return vbaFalse;
                const cleaned = s.replace(/[$,]/g, '');
                return (!isNaN(Number(cleaned)) && isFinite(Number(cleaned))) ? vbaTrue : vbaFalse;
            }
            return vbaFalse;
        });
        this.env.set('isdate', (val: any) => {
            if (val instanceof VbaDate) return vbaTrue;
            if (typeof val === 'string') {
                const d = Date.parse(val);
                return !isNaN(d) ? vbaTrue : vbaFalse;
            }
            return vbaFalse;
        });
        this.env.set('isobject', (val: any) => (val && typeof val === 'object' && !Array.isArray(val) && val !== vbaNull) ? vbaTrue : vbaFalse);
        this.env.set('iserror', (val: any) => (val instanceof VbaErrorValue) ? vbaTrue : vbaFalse);
        this.env.set('isnull', (val: any) => (val === vbaNull) ? vbaTrue : vbaFalse);
        this.env.set('isarray', (val: any) => Array.isArray(val) ? vbaTrue : vbaFalse);

        this.env.set('vartype', (val: any) => {
            if (val === vbaEmpty || val === undefined) return 0; // vbEmpty
            if (val === vbaNull) return 1; // vbNull
            if (val instanceof VbaBoolean) return 11; // vbBoolean
            if (val instanceof VbaDate) return 7; // vbDate
            if (val === vbaMissing || val instanceof VbaErrorValue) return 10; // vbError
            if (Array.isArray(val)) return 8192 + 12; // vbArray + vbVariant
            if (typeof val === 'number') return 5; // vbDouble
            if (typeof val === 'string') return 8; // vbString
            if (val instanceof VbaDecimal) return 14; // vbDecimal
            if (typeof val === 'bigint') return 20; // vbLongLong
            if (val && val.__vbaTypeName__) return 36; // vbUserDefinedType
            if (val && (val.__vbaClass__ || val.__isVbaDict__ || val.__isVbaCollection__)) return 9; // vbObject
            if (typeof val === 'object' && val !== null) return 9; // vbObject
            return 12; // vbVariant
        });

        this.env.set('typename', (val: any) => {
            if (val === vbaEmpty || val === undefined) return 'Empty';
            if (val === vbaNull) return 'Null';
            if (val === vbaNothing) return 'Nothing';
            if (val === vbaMissing || val instanceof VbaErrorValue) return 'Error';
            if (val instanceof VbaBoolean) return 'Boolean';
            if (val instanceof VbaDate) return 'Date';
            if (typeof val === 'number') return 'Double';
            if (typeof val === 'string') return 'String';
            if (Array.isArray(val)) return 'Variant()';
            if (val && val.__vbaTypeName__) return val.__vbaTypeName__;
            if (val && val.__isVbaDict__) return 'Dictionary';
            if (val && val.__isVbaCollection__) return 'Collection';
            if (val && val.__vbaClass__) return val.__className__;
            if (val instanceof VbaDecimal) return 'Decimal';
            if (typeof val === 'bigint') return 'LongLong';
            return 'Object';
        });

        // --- Conversion Module ---
        this.env.set('cbyte', (val: any) => {
            if (val instanceof VbaBoolean) {
                return val.valueOf() ? 255 : 0;
            }
            const n = this.vbaRound(this.toVbaNumber(val));
            if (n < 0 || n > 255) this.throwVbaError(6, "Overflow");
            return n;
        });
        this.env.set('cint', (val: any) => {
            const n = this.vbaRound(this.toVbaNumber(val));
            if (n < -32768 || n > 32767) this.throwVbaError(6, "Overflow");
            return n;
        });
        this.env.set('clng', (val: any) => {
            const n = this.vbaRound(this.toVbaNumber(val));
            if (n < -2147483648 || n > 2147483647) this.throwVbaError(6, "Overflow");
            return n;
        });
        this.env.set('csng', (val: any) => {
            const n = this.toVbaNumber(val);
            const f32 = Math.fround(n);
            if (!isFinite(f32) && isFinite(n)) this.throwVbaError(6, "Overflow");
            return f32;
        });
        this.env.set('cdbl', (val: any) => this.toVbaNumber(val));
        this.env.set('cdate', (val: any) => {
            if (val === null || val === vbaNull || val === vbaEmpty) this.throwVbaError(13, "Type mismatch");
            if (val instanceof VbaDate) return val;
            if (typeof val === 'string' && !/^\d+(\.\d+)?$/.test(val)) {
                const d = Date.parse(val);
                if (isNaN(d)) this.throwVbaError(13, "Type mismatch");
                return new VbaDate(d / 86400000 + 25569);
            }
            return new VbaDate(this.toVbaNumber(val));
        });
        this.env.set('cvdate', (val: any) => {
            if (val === vbaNull) return vbaNull;
            return (this.env.get('cdate') as Function)(val);
        });
        this.env.set('cdec', (val: any) => new VbaDecimal(this.toVbaNumber(val)));
        this.env.set('ccur', (val: any) => {
            const n = this.vbaRound(this.toVbaNumber(val), 4);
            if (n < -922337203685477.5808 || n > 922337203685477.5807) this.throwVbaError(6, "Overflow");
            return n;
        });
        this.env.set('clnglng', (val: any) => {
            if (val === vbaNull || val === null) this.throwVbaError(13, "Type mismatch");
            if (typeof val === 'bigint') return val;
            if (typeof val === 'string') {
                const trimmed = val.trim();
                if (/^-?\d+$/.test(trimmed)) {
                    try {
                        const n = BigInt(trimmed);
                        if (n < -9223372036854775808n || n > 9223372036854775807n) this.throwVbaError(6, "Overflow");
                        return n;
                    } catch {
                        this.throwVbaError(13, "Type mismatch");
                    }
                }
            }
            const num = this.toVbaNumber(val);
            // Check range using number (approximate but necessary for large floats)
            // Note: 9223372036854775807 is exactly represented as a double but the next double is 9223372036854775807 + 1024
            if (num < -9223372036854775808 || num > 9223372036854775807) {
                // If it's very close to the edge, it might be a precision issue
                if (Math.abs(num - 9223372036854775807) < 1025 || Math.abs(num + 9223372036854775808) < 1025) {
                    // Keep going and let BigInt conversion handle it if it was a float that rounded to the edge
                } else {
                    this.throwVbaError(6, "Overflow");
                }
            }
            const n = BigInt(this.vbaRound(num));
            if (n < -9223372036854775808n || n > 9223372036854775807n) this.throwVbaError(6, "Overflow");
            return n;
        });
        this.env.set('clngptr', this.env.get('clnglng'));
        this.env.set('cstr', (val: any) => val === vbaNull ? this.throwVbaError(94, "Invalid use of Null") : String(val === vbaEmpty ? "" : val));
        this.env.set('cbool', (val: any) => this.isTrue(val) ? vbaTrue : vbaFalse);
        this.env.set('cvar', (val: any) => val);
        this.env.set('cverr', (val: any) => {
            if (val instanceof VbaErrorValue) return val;
            const code = this.toVbaNumber(val);
            if (code < 0 || code > 65535) this.throwVbaError(5, "Invalid procedure call or argument");
            return new VbaErrorValue(code);
        });

        this.env.set('hex', (n: any) => {
            if (n === vbaNull) return vbaNull;
            return (Math.floor(this.toVbaNumber(n)) >>> 0).toString(16).toUpperCase();
        });
        this.env.set('oct', (n: any) => {
            if (n === vbaNull) return vbaNull;
            return (Math.floor(this.toVbaNumber(n)) >>> 0).toString(8);
        });
        this.env.set('val', (s: any) => {
            if (typeof s !== 'string') return 0;
            const cleaned = s.trim().replace(/ /g, '');
            if (cleaned.toLowerCase().startsWith('&h')) return parseInt(cleaned.slice(2), 16) || 0;
            if (cleaned.toLowerCase().startsWith('&o')) return parseInt(cleaned.slice(2), 8) || 0;
            const match = cleaned.match(/^[+-]?\d*(\.\d*)?/);
            return match ? parseFloat(match[0]) || 0 : 0;
        });

        // Constants
        this.env.set('vbcrlf', "\r\n");
        this.env.set('vbcr', "\r");
        this.env.set('vblf', "\n");
        this.env.set('vbtab', "\t");
        this.env.set('vbnullstring', "");
        this.env.set('vbnullchar', "\0");
        this.env.set('vbok', 1);
        this.env.set('vbcancel', 2);
        this.env.set('vbabort', 3);
        this.env.set('vbretry', 4);
        this.env.set('vbignore', 5);
        this.env.set('vbyes', 6);
        this.env.set('vbno', 7);
        this.env.set('vbokonly', 0);
        this.env.set('vbokcancel', 1);
        this.env.set('vbabortretryignore', 2);
        this.env.set('vbyesnocancel', 3);
        this.env.set('vbyesno', 4);
        this.env.set('vbretrycancel', 5);
        this.env.set('vbcritical', 16);
        this.env.set('vbquestion', 32);
        this.env.set('vbexclamation', 48);
        this.env.set('vbinformation', 64);
        this.env.set('vbdefaultbutton1', 0);
        this.env.set('vbdefaultbutton2', 256);
        this.env.set('vbdefaultbutton3', 512);
        this.env.set('vbdefaultbutton4', 768);
        this.env.set('vbtextcompare', 1);
        this.env.set('vbbinarycompare', 0);
        this.env.set('vbuppercase', 1);
        this.env.set('vblowercase', 2);
        this.env.set('vbpropercase', 3);
        this.env.set('vbwide', 4);
        this.env.set('vbnarrow', 8);
        this.env.set('vbkatakana', 16);
        this.env.set('vbhiragana', 32);

        // --- Math Module ---
        this.env.set('abs', (val: any) => val === vbaNull ? vbaNull : Math.abs(this.toVbaNumber(val)));
        this.env.set('atn', (val: any) => val === vbaNull ? vbaNull : Math.atan(this.toVbaNumber(val)));
        this.env.set('cos', (val: any) => val === vbaNull ? vbaNull : Math.cos(this.toVbaNumber(val)));
        this.env.set('exp', (val: any) => {
            if (val === vbaNull) return vbaNull;
            const n = this.toVbaNumber(val);
            if (n > 709.782712893) this.throwVbaError(6, "Overflow");
            return Math.exp(n);
        });
        this.env.set('int', (val: any) => val === vbaNull ? vbaNull : Math.floor(this.toVbaNumber(val)));
        this.env.set('fix', (val: any) => {
            if (val === vbaNull) return vbaNull;
            const n = this.toVbaNumber(val);
            return n >= 0 ? Math.floor(n) : Math.ceil(n);
        });
        this.env.set('log', (val: any) => {
            if (val === vbaNull) return vbaNull;
            const n = this.toVbaNumber(val);
            if (n <= 0) this.throwVbaError(5, "Invalid procedure call or argument");
            return Math.log(n);
        });
        this.env.set('round', (val: any, digits: any = 0) => val === vbaNull ? vbaNull : this.vbaRound(this.toVbaNumber(val), Number(digits)));
        this.env.set('sgn', (val: any) => {
            if (val === vbaNull) return vbaNull;
            const n = this.toVbaNumber(val);
            return n > 0 ? 1 : n < 0 ? -1 : 0;
        });
        this.env.set('sin', (val: any) => val === vbaNull ? vbaNull : Math.sin(this.toVbaNumber(val)));
        this.env.set('sqr', (val: any) => {
            if (val === vbaNull) return vbaNull;
            const n = this.toVbaNumber(val);
            if (n < 0) this.throwVbaError(5, "Invalid procedure call or argument");
            return Math.sqrt(n);
        });
        this.env.set('tan', (val: any) => val === vbaNull ? vbaNull : Math.tan(this.toVbaNumber(val)));

        let rndSeed = 0.5;
        let lastRnd = 0.5;
        let rndInitialized = false;
        this.env.set('rnd', (val?: any) => {
            if (!rndInitialized) { rndSeed = 0.5; rndInitialized = true; }
            if (val === undefined || (typeof val === 'number' && val > 0)) {
                rndSeed = (rndSeed * 214013 + 2531011) % 4294967296;
                lastRnd = rndSeed / 4294967296;
                return lastRnd;
            } else if (val === 0) {
                return lastRnd;
            } else if (val < 0) {
                const s = Math.abs(Number(val)) * 9301 + 49297;
                lastRnd = (s % 233280) / 233280;
                return lastRnd;
            }
            return lastRnd;
        });
        this.env.set('randomize', (val?: any) => {
            rndInitialized = true;
            rndSeed = (val === undefined || val === null) ? (Date.now() % 4294967296) : (Math.round(Math.abs(Number(val)) * 1000) % 4294967296);
            lastRnd = rndSeed / 4294967296;
        });

        // --- String Module ---
        const ascFunc = (s: any) => String(s || '').charCodeAt(0);
        this.env.set('asc', ascFunc);
        this.env.set('ascw', ascFunc);
        const chrFunc = (n: any) => String.fromCharCode(Number(n));
        this.env.set('chr', chrFunc);
        this.env.set('chr$', chrFunc);
        this.env.set('chrw', chrFunc);
        this.env.set('instr', (...args: any[]) => {
            let start = 1, s1, s2, comp;
            if (args.length >= 4) [start, s1, s2, comp] = args;
            else if (args.length === 3 && typeof args[0] === 'number') [start, s1, s2] = args;
            else [s1, s2] = args;
            if (s1 === vbaNull || s2 === vbaNull) return vbaNull;
            const str1 = String(s1 ?? ''), str2 = String(s2 ?? '');
            const isText = (comp === 1) || (comp === undefined && this.comparisonMode === 'Text');
            const idx = isText ? str1.toLowerCase().indexOf(str2.toLowerCase(), start - 1) : str1.indexOf(str2, start - 1);
            return idx === -1 ? 0 : idx + 1;
        });
        this.env.set('instrrev', (s1: any, s2: any, start: any = -1, comp: any = undefined) => {
            if (s1 === vbaNull || s2 === vbaNull) return vbaNull;
            const str = String(s1 ?? ''), find = String(s2 ?? '');
            if (find === "") return (start === -1) ? str.length : Number(start);
            
            const startNum = Number(start);
            if (startNum !== -1 && startNum > str.length) return 0;
            
            const effStart = (start === -1) ? str.length : startNum;
            const isText = (comp === 1) || (comp === undefined && this.comparisonMode === 'Text');
            const idx = isText ? str.toLowerCase().lastIndexOf(find.toLowerCase(), effStart - 1) : str.lastIndexOf(find, effStart - 1);
            return idx === -1 ? 0 : idx + 1;
        });
        const lcaseFunc = (val: any) => val === vbaNull ? vbaNull : String(val ?? '').toLowerCase();
        this.env.set('lcase', lcaseFunc);
        this.env.set('lcase$', lcaseFunc);
        const strFunc = (val: any) => {
            if (val === vbaNull) return vbaNull;
            const n = this.toVbaNumber(val);
            return n >= 0 ? " " + n : String(n);
        };
        this.env.set('str', strFunc);
        this.env.set('str$', strFunc);

        const ucaseFunc = (val: any) => val === vbaNull ? vbaNull : String(val ?? '').toUpperCase();
        this.env.set('ucase', ucaseFunc);
        this.env.set('ucase$', ucaseFunc);
        const leftFunc = (val: any, len: any) => String(val ?? '').substring(0, Number(len));
        this.env.set('left', leftFunc);
        this.env.set('left$', leftFunc);
        const rightFunc = (val: any, len: any) => {
            const s = String(val ?? ''), l = Number(len);
            return s.substring(s.length - l);
        };
        this.env.set('right', rightFunc);
        this.env.set('right$', rightFunc);
        const midFunc = (val: any, start: any, len?: any) => {
            const s = String(val ?? ''), st = Number(start);
            return len !== undefined ? s.substring(st - 1, st - 1 + Number(len)) : s.substring(st - 1);
        };
        this.env.set('mid', midFunc);
        this.env.set('mid$', midFunc);
        this.env.set('len', (val: any) => val === vbaNull ? vbaNull : String(val ?? '').length);
        const ltrimFunc = (val: any) => val === vbaNull ? vbaNull : String(val ?? '').trimStart();
        this.env.set('ltrim', ltrimFunc);
        this.env.set('ltrim$', ltrimFunc);
        const rtrimFunc = (val: any) => val === vbaNull ? vbaNull : String(val ?? '').trimEnd();
        this.env.set('rtrim', rtrimFunc);
        this.env.set('rtrim$', rtrimFunc);
        const trimFunc = (val: any) => val === vbaNull ? vbaNull : String(val ?? '').trim();
        this.env.set('trim', trimFunc);
        this.env.set('trim$', trimFunc);
        const spaceFunc = (n: any) => ' '.repeat(Number(n));
        this.env.set('space', spaceFunc);
        this.env.set('space$', spaceFunc);
        const stringFunc = (n: any, char: any) => {
            const s = String(char ?? '');
            return (s.length > 0 ? s[0] : '').repeat(Number(n));
        };
        this.env.set('string', stringFunc);
        this.env.set('string$', stringFunc);
        this.env.set('split', (s: any, del: string = ' ') => String(s ?? '').split(del));
        this.env.set('join', (arr: any, del: string = ' ') => Array.isArray(arr) ? arr.join(del) : String(arr));
        this.env.set('replace', (s: any, f: any, r: any) => String(s ?? '').split(String(f ?? '')).join(String(r ?? '')));
        this.env.set('strcomp', (s1: any, s2: any, comp?: number) => {
            if (s1 === vbaNull || s2 === vbaNull) return vbaNull;
            let str1 = String(s1 ?? ''), str2 = String(s2 ?? '');
            const isText = (comp === 1) || (comp === undefined && this.comparisonMode === 'Text');
            if (isText) { str1 = str1.toLowerCase(); str2 = str2.toLowerCase(); }
            return str1 < str2 ? -1 : (str1 > str2 ? 1 : 0);
        });
        this.env.set('strconv', (s: any, conv: any) => {
            if (s === vbaNull) return vbaNull;
            let str = String(s ?? '');
            const c = Number(conv);
            const caseConv = c & 3;
            if (caseConv === 1) str = str.toUpperCase();
            else if (caseConv === 2) str = str.toLowerCase();
            else if (caseConv === 3) str = str.split(/\s+/).map(w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase()).join(' ');

            if (c & 16) {
                str = str.replace(/[\u3041-\u3096]/g, m => String.fromCharCode(m.charCodeAt(0) + 0x60));
            } else if (c & 32) {
                str = str.replace(/[\u30A1-\u30F6]/g, m => String.fromCharCode(m.charCodeAt(0) - 0x60));
            }

            if (c & 4) {
                str = str.replace(/[!-~]/g, m => String.fromCharCode(m.charCodeAt(0) + 0xFEE0)).replace(/ /g, '　');
            } else if (c & 8) {
                str = str.replace(/[！-～]/g, m => String.fromCharCode(m.charCodeAt(0) - 0xFEE0)).replace(/　/g, ' ');
            }
            return str;
        });
        this.env.set('strreverse', (s: any) => s === vbaNull ? vbaNull : String(s ?? '').split('').reverse().join(''));
        this.env.set('filter', (source: any, match: any, include: any = vbaTrue, compare: any = undefined) => {
            if (!Array.isArray(source)) this.throwVbaError(13, "Type mismatch");
            const find = String(match ?? '');
            const isInclude = this.isTrue(include);
            const isText = (compare === 1) || (compare === undefined && this.comparisonMode === 'Text');
            const result = source.filter(s => {
                const str = String(s ?? '');
                const found = isText ? str.toLowerCase().includes(find.toLowerCase()) : str.includes(find);
                return isInclude ? found : !found;
            });
            (result as any).vbaBase = 0;
            return result;
        });
        this.env.set('leftb', (val: any, len: any) => {
            const s = String(val ?? '');
            // In VBA, strings are UTF-16, so each char is 2 bytes. 
            // LeftB(s, 2) returns first char.
            return s.substring(0, Math.floor(Number(len) / 2));
        });
        this.env.set('rightb', (val: any, len: any) => {
            const s = String(val ?? '');
            const charLen = Math.floor(Number(len) / 2);
            return s.substring(s.length - charLen);
        });
        const midbFunc = (val: any, start: any, len?: any) => {
            const s = String(val ?? '');
            const charStart = Math.floor((Number(start) + 1) / 2);
            if (len === undefined) return s.substring(charStart - 1);
            const charLen = Math.floor(Number(len) / 2);
            return s.substring(charStart - 1, charStart - 1 + charLen);
        };
        this.env.set('midb', midbFunc);
        this.env.set('midb$', midbFunc);
        const formatFunc = (val: any, pattern?: string) => {
            if (val === null || val === vbaNull || val === vbaEmpty) return "";
            const fmt = pattern ? String(pattern) : "";
            if (fmt === "") return String(val);
            const fmtLower = fmt.toLowerCase();
            const namedFormats = ['general number', 'currency', 'fixed', 'standard', 'percent', 'scientific', 'true/false', 'yes/no', 'on/off'];
            const dateNamedFormats = ['general date', 'long date', 'medium date', 'short date', 'long time', 'medium time', 'short time'];
            if (namedFormats.includes(fmtLower)) {
                if (typeof val === 'number') return this.formatNumber(val, fmt);
                return String(val);
            }
            if (dateNamedFormats.includes(fmtLower)) {
                const dateVal = val instanceof VbaDate ? fromVbaDate(val.value) : (typeof val === 'number' ? fromVbaDate(val) : new Date(String(val)));
                return this.formatDate(dateVal, fmt);
            }
            const isDatePattern = /y|m|d|h|n|s|am\/pm/i.test(fmt);
            if (val instanceof VbaDate) return this.formatDate(fromVbaDate(val.value), fmt);
            if (typeof val === 'number') {
                if (isDatePattern && !/^[0#,.%]+$/.test(fmt)) return this.formatDate(fromVbaDate(val), fmt);
                return this.formatNumber(val, fmt);
            }
            return String(val);
        };
        this.env.set('format', formatFunc);
        this.env.set('format$', formatFunc);

        // --- Date/Time Module ---
        const nowFunc = () => new VbaDate(toVbaDate(new Date()));
        (nowFunc as any).__vbaAutoCall__ = true;
        this.env.set('now', nowFunc);

        const dateFunc = () => new VbaDate(Math.floor(toVbaDate(new Date())));
        (dateFunc as any).__vbaAutoCall__ = true;
        this.env.set('date', dateFunc);

        const timeFunc = () => {
            const serial = toVbaDate(new Date());
            return new VbaDate(serial - Math.floor(serial));
        };
        (timeFunc as any).__vbaAutoCall__ = true;
        this.env.set('time', timeFunc);

        const timerFunc = () => {
            const now = new Date();
            const midnight = new Date(now.getFullYear(), now.getMonth(), now.getDate());
            return (now.getTime() - midnight.getTime()) / 1000;
        };
        (timerFunc as any).__vbaAutoCall__ = true;
        this.env.set('timer', timerFunc);
        this.env.set('year', (d: any) => parseVbaDate(d).getUTCFullYear());
        this.env.set('month', (d: any) => parseVbaDate(d).getUTCMonth() + 1);
        this.env.set('day', (d: any) => parseVbaDate(d).getUTCDate());
        this.env.set('hour', (d: any) => parseVbaDate(d).getUTCHours());
        this.env.set('minute', (d: any) => parseVbaDate(d).getUTCMinutes());
        this.env.set('second', (d: any) => parseVbaDate(d).getUTCSeconds());
        this.env.set('dateserial', (y: any, m: any, d: any) => new VbaDate(toVbaDate(new Date(Date.UTC(Number(y), Number(m) - 1, Number(d))))));
        this.env.set('timeserial', (h: any, n: any, s: any) => new VbaDate(toVbaDate(new Date(Date.UTC(1899, 11, 30, Number(h), Number(n), Number(s))))));
        this.env.set('weekday', (d: any) => parseVbaDate(d).getUTCDay() + 1);
        this.env.set('dateadd', (interval: any, number: any, date: any) => {
            const d = parseVbaDate(date);
            const n = Number(number);
            const intv = String(interval).toLowerCase();
            
            if (intv === 'yyyy' || intv === 'q' || intv === 'm') {
                const oldDay = d.getUTCDate();
                if (intv === 'yyyy') d.setUTCFullYear(d.getUTCFullYear() + n);
                else if (intv === 'q') d.setUTCMonth(d.getUTCMonth() + n * 3);
                else if (intv === 'm') d.setUTCMonth(d.getUTCMonth() + n);
                
                // VBA behavior: if the day exceeds the last day of the month, 
                // it is set to the last day of that month.
                if (d.getUTCDate() !== oldDay) {
                    d.setUTCDate(0);
                }
            } else if (intv === 'y' || intv === 'd' || intv === 'w') {
                d.setUTCDate(d.getUTCDate() + n);
            } else if (intv === 'ww') {
                d.setUTCDate(d.getUTCDate() + n * 7);
            } else if (intv === 'h') {
                d.setUTCHours(d.getUTCHours() + n);
            } else if (intv === 'n') {
                d.setUTCMinutes(d.getUTCMinutes() + n);
            } else if (intv === 's') {
                d.setUTCSeconds(d.getUTCSeconds() + n);
            }
            return new VbaDate(toVbaDate(d));
        });
        this.env.set('datediff', (interval: any, date1: any, date2: any) => {
            const d1 = parseVbaDate(date1);
            const d2 = parseVbaDate(date2);
            const intv = String(interval).toLowerCase();
            const diffMs = d2.getTime() - d1.getTime();
            if (intv === 'yyyy') return d2.getUTCFullYear() - d1.getUTCFullYear();
            else if (intv === 'q') return (d2.getUTCFullYear() - d1.getUTCFullYear()) * 4 + Math.floor(d2.getUTCMonth() / 3) - Math.floor(d1.getUTCMonth() / 3);
            else if (intv === 'm') return (d2.getUTCFullYear() - d1.getUTCFullYear()) * 12 + d2.getUTCMonth() - d1.getUTCMonth();
            else if (intv === 'y' || intv === 'd' || intv === 'w') return Math.round(diffMs / 86400000);
            else if (intv === 'ww') return Math.round(diffMs / 604800000);
            else if (intv === 'h') return Math.round(diffMs / 3600000);
            else if (intv === 'n') return Math.round(diffMs / 60000);
            else if (intv === 's') return Math.round(diffMs / 1000);
            return 0;
        });
        this.env.set('datepart', (interval: any, date: any) => {
            const d = parseVbaDate(date);
            const intv = String(interval).toLowerCase();
            if (intv === 'yyyy') return d.getUTCFullYear();
            else if (intv === 'q') return Math.floor(d.getUTCMonth() / 3) + 1;
            else if (intv === 'm') return d.getUTCMonth() + 1;
            else if (intv === 'y') {
                const start = new Date(Date.UTC(d.getUTCFullYear(), 0, 0));
                const diff = d.getTime() - start.getTime();
                return Math.floor(diff / 86400000);
            }
            else if (intv === 'd') return d.getUTCDate();
            else if (intv === 'w') return d.getUTCDay() + 1;
            else if (intv === 'ww') {
                const start = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
                const diff = d.getTime() - start.getTime();
                return Math.floor((diff / 86400000 + start.getUTCDay() + 6) / 7);
            }
            else if (intv === 'h') return d.getUTCHours();
            else if (intv === 'n') return d.getUTCMinutes();
            else if (intv === 's') return d.getUTCSeconds();
            return 0;
        });
        this.env.set('datevalue', (val: any) => {
            const d = parseVbaDate(val);
            return new VbaDate(Math.floor(toVbaDate(d)));
        });
        this.env.set('timevalue', (val: any) => {
            const d = parseVbaDate(val);
            const serial = toVbaDate(d);
            return new VbaDate(serial - Math.floor(serial));
        });
        this.env.set('monthname', (month: any, abbreviate: any = vbaFalse) => {
            const m = Number(month);
            if (m < 1 || m > 12) this.throwVbaError(5, "Invalid procedure call or argument");
            const names = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
            const abbrs = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
            return this.isTrue(abbreviate) ? abbrs[m - 1] : names[m - 1];
        });
        this.env.set('weekdayname', (weekday: any, abbreviate: any = vbaFalse, firstdayofweek: any = 1) => {
            const w = Number(weekday);
            let first = Number(firstdayofweek);
            if (first === 0) first = 1;
            if (w < 1 || w > 7) this.throwVbaError(5, "Invalid procedure call or argument");
            const names = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
            const abbrs = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
            const idx = (w + first - 2) % 7;
            return this.isTrue(abbreviate) ? abbrs[idx] : names[idx];
        });

        // --- File System Module ---
        this.env.set('freefile', (range?: number) => {
            const start = (range === 1) ? 256 : 1;
            const end = (range === 1) ? 511 : 255;
            for (let i = start; i <= end; i++) if (!this.fileHandles.has(i)) return i;
            this.throwVbaError(67, "Too many files");
        });
        this.env.set('eof', (fn: any) => {
            const h = this.fileHandles.get(Number(fn));
            if (!h) this.throwVbaError(52, "Bad file name or number");
            return h.pos! >= fs.statSync(h.path).size ? vbaTrue : vbaFalse;
        });
        this.env.set('lof', (fn: any) => {
            const h = this.fileHandles.get(Number(fn));
            if (!h) this.throwVbaError(52, "Bad file name or number");
            return fs.statSync(h.path).size;
        });
        this.env.set('loc', (fn: any) => {
            const h = this.fileHandles.get(Number(fn));
            if (!h) this.throwVbaError(52, "Bad file name or number");
            return h.pos;
        });
        this.env.set('seek', (fn: any) => {
            const h = this.fileHandles.get(Number(fn));
            if (!h) this.throwVbaError(52, "Bad file name or number");
            return (h.pos || 0) + 1;
        });
        this.env.set('fileattr', (fn: any, info: any = 1) => {
            console.log(`[STUB] FileAttr #${fn}, ${info}`);
            return 1; // Default to Normal
        });
        this.env.set('chdrive', (drive: any) => {
            console.log(`[STUB] ChDrive "${drive}"`);
        });
        this.env.set('setattr', (path: any, attr: any) => {
            console.log(`[STUB] SetAttr "${path}", ${attr}`);
        });
        this.env.set('filedatetime', (path: any) => {
            const realPath = this.sandbox.toRealPath(String(path));
            const stats = fs.statSync(realPath);
            return new VbaDate(toVbaDate(stats.mtime));
        });

        this.env.set('curdir', (drive?: string) => this.sandbox.toVirtualPath(this.sandbox.getRoot()));
        this.env.set('curdir$', this.env.get('curdir'));
        this.env.set('dir', (pathName?: string, attributes?: number) => {
            if (pathName !== undefined && pathName !== null && pathName !== "") {
                try {
                    const realPath = this.sandbox.toRealPath(pathName);
                    const dirPath = path.dirname(realPath);
                    const filter = path.basename(realPath);
                    const files = fs.readdirSync(dirPath);
                    if (filter === "*.*" || filter === "*" || filter === "") {
                        this.dirIterator = files;
                    } else {
                        const regex = new RegExp('^' + filter.replace(/\./g, '\\.').replace(/\*/g, '.*').replace(/\?/g, '.') + '$', 'i');
                        this.dirIterator = files.filter((f: string) => regex.test(f));
                    }
                    this.dirIndex = 0;
                } catch (e) {
                    this.dirIterator = [];
                }
            }
            if (!this.dirIterator || this.dirIndex >= this.dirIterator.length) return "";
            return this.dirIterator[this.dirIndex++];
        });
        this.env.set('dir$', this.env.get('dir'));
        this.env.set('kill', (path: string) => fs.unlinkSync(this.sandbox.toRealPath(path)));
        this.env.set('filecopy', (src: string, dest: string) => fs.copyFileSync(this.sandbox.toRealPath(src), this.sandbox.toRealPath(dest)));
        this.env.set('mkdir', (path: string) => fs.mkdirSync(this.sandbox.toRealPath(path), { recursive: true }));
        this.env.set('rmdir', (path: string) => fs.rmdirSync(this.sandbox.toRealPath(path)));
        this.env.set('chdir', (path: string) => { /* Mock or Sandbox logic */ });
        this.env.set('filelen', (path: string) => fs.statSync(this.sandbox.toRealPath(path)).size);

        // --- Interaction Module ---
        this.env.set('shell', (cmd: any, style: any = 1) => { this.onPrint(`[SHELL] ${cmd} (Style: ${style})`); return 1; });
        this.env.set('msgbox', (msg: any, buttons: any = 0, title: any = "") => { this.onPrint(`[MSGBOX] ${title}: ${msg}`); return 1; });
        this.env.set('inputbox', (prompt: any, title: any = "", def: any = "") => { this.onPrint(`[INPUTBOX] ${prompt}`); return def; });
        this.env.set('appactivate', (title: string, wait?: boolean) => { this.onPrint(`[APPACTIVATE] ${title}`); });
        this.env.set('sendkeys', (keys: string, wait?: boolean) => { this.onPrint(`[SENDKEYS] ${keys}`); });
        this.env.set('doevents', () => 0);

        // --- Financial Module ---
        const getRateFactor = (rate: number, nper: number) => {
            if (rate === 0) return nper;
            return (Math.pow(1 + rate, nper) - 1) / rate;
        };
        this.env.set('fv', (rate: any, nper: any, pmt: any, pv: any = 0, type: any = 0) => {
            const r = Number(rate), n = Number(nper), p = Number(pmt), v = Number(pv), t = Number(type);
            const factor = getRateFactor(r, n);
            const result = -(v * Math.pow(1 + r, n) + p * (1 + r * t) * factor);
            return result;
        });
        this.env.set('pv', (rate: any, nper: any, pmt: any, fv: any = 0, type: any = 0) => {
            const r = Number(rate), n = Number(nper), p = Number(pmt), f = Number(fv), t = Number(type);
            if (r === 0) return -(f + p * n);
            const p1 = Math.pow(1 + r, n);
            const result = -(f + p * (1 + r * t) * ((p1 - 1) / r)) / p1;
            return result;
        });
        this.env.set('pmt', (rate: any, nper: any, pv: any, fv: any = 0, type: any = 0) => {
            const r = Number(rate), n = Number(nper), v = Number(pv), f = Number(fv), t = Number(type);
            if (r === 0) return -(v + f) / n;
            const p1 = Math.pow(1 + r, n);
            const result = -(v * p1 + f) / ((1 + r * t) * ((p1 - 1) / r));
            return result;
        });
        this.env.set('nper', (rate: any, pmt: any, pv: any, fv: any = 0, type: any = 0) => {
            const r = Number(rate), p = Number(pmt), v = Number(pv), f = Number(fv), t = Number(type);
            if (r === 0) return -(v + f) / p;
            const num = p * (1 + r * t) - f * r;
            const den = p * (1 + r * t) + v * r;
            return Math.log(num / den) / Math.log(1 + r);
        });
        this.env.set('rate', (nper: any, pmt: any, pv: any, fv: any = 0, type: any = 0, guess: any = 0.1) => {
            // Newton-Raphson approximation for Rate
            let r = Number(guess);
            const n = Number(nper), p = Number(pmt), v = Number(pv), f = Number(fv), t = Number(type);
            for (let i = 0; i < 20; i++) {
                const p1 = Math.pow(1 + r, n);
                const f_r = v * p1 + p * (1 + r * t) * ((p1 - 1) / r) + f;
                const df_r = v * n * Math.pow(1 + r, n - 1) + p * (t * ((p1 - 1) / r) + (1 + r * t) * (n * Math.pow(1 + r, n - 1) * r - (p1 - 1)) / (r * r));
                const newR = r - f_r / df_r;
                if (Math.abs(newR - r) < 1e-10) return newR;
                r = newR;
            }
            return r;
        });
        this.env.set('sln', (cost: any, salvage: any, life: any) => {
            return (Number(cost) - Number(salvage)) / Number(life);
        });
        this.env.set('syd', (cost: any, salvage: any, life: any, period: any) => {
            const c = Number(cost), s = Number(salvage), l = Number(life), p = Number(period);
            return ((c - s) * (l - p + 1) * 2) / (l * (l + 1));
        });
        this.env.set('ddb', (cost: any, salvage: any, life: any, period: any, factor: any = 2) => {
            let c = Number(cost), s = Number(salvage), l = Number(life), p = Number(period), f = Number(factor);
            if (p <= 0 || p > l) return 0;
            let book = c;
            let dep = 0;
            for (let i = 1; i <= p; i++) {
                dep = Math.min(book * (f / l), Math.max(0, book - s));
                book -= dep;
            }
            return dep;
        });
        this.env.set('irr', (values: any, guess: any = 0.1) => {
            if (!Array.isArray(values)) this.throwVbaError(13, "Type mismatch");
            const v = values.map(Number);
            let r = Number(guess);
            for (let i = 0; i < 100; i++) {
                let npv = 0;
                let dnpv = 0;
                for (let t = 0; t < v.length; t++) {
                    const p1 = Math.pow(1 + r, t);
                    npv += v[t] / p1;
                    if (t > 0) dnpv -= t * v[t] / (p1 * (1 + r));
                }
                const newR = r - npv / dnpv;
                if (Math.abs(newR - r) < 1e-10) return newR;
                r = newR;
            }
            return r;
        });
        this.env.set('mirr', (values: any, finance_rate: any, reinvest_rate: any) => {
            if (!Array.isArray(values)) this.throwVbaError(13, "Type mismatch");
            const v = values.map(Number);
            const fr = Number(finance_rate), rr = Number(reinvest_rate);
            const n = v.length - 1;
            let npv_neg = 0;
            let npv_pos = 0;
            for (let t = 0; t < v.length; t++) {
                if (v[t] < 0) npv_neg += v[t] / Math.pow(1 + fr, t);
                else npv_pos += v[t] / Math.pow(1 + rr, t);
            }
            const tv = npv_pos * Math.pow(1 + rr, n);
            return Math.pow(-tv / npv_neg, 1 / n) - 1;
        });
        this.env.set('npv', (rate: any, values: any) => {
            if (!Array.isArray(values)) this.throwVbaError(13, "Type mismatch");
            const r = Number(rate);
            const v = values.map(Number);
            let result = 0;
            for (let i = 0; i < v.length; i++) {
                result += v[i] / Math.pow(1 + r, i + 1);
            }
            return result;
        });
        this.env.set('ipmt', (rate: any, per: any, nper: any, pv: any, fv: any = 0, type: any = 0) => {
            const r = Number(rate), p = Number(per), n = Number(nper), v = Number(pv), f = Number(fv), t = Number(type);
            const pmt = Number(this.env.get('pmt')(r, n, v, f, t));
            let ipmt: number;
            if (p === 1) {
                ipmt = t === 1 ? 0 : -v * r;
            } else {
                const fv_prev = Number(this.env.get('fv')(r, p - 1, pmt, v, t));
                ipmt = -fv_prev * r;
            }
            return ipmt;
        });
        this.env.set('ppmt', (rate: any, per: any, nper: any, pv: any, fv: any = 0, type: any = 0) => {
            const r = Number(rate), p = Number(per), n = Number(nper), v = Number(pv), f = Number(fv), t = Number(type);
            const pmt = Number(this.env.get('pmt')(r, n, v, f, t));
            const ipmt = Number(this.env.get('ipmt')(r, p, n, v, f, t));
            return pmt - ipmt;
        });

        // --- Constants & Registry ---
        const errorMessages: Record<number, string> = { 5: "Invalid procedure call or argument", 6: "Overflow", 9: "Subscript out of range", 11: "Division by zero", 13: "Type mismatch", 52: "Bad file name or number", 53: "File not found", 58: "File already exists", 62: "Input past end of file", 70: "Permission denied", 76: "Path not found", 91: "Object variable not set", 94: "Invalid use of Null" };
        const errFunc = (n?: any) => errorMessages[n === undefined ? this.errObj.number : Number(n)] || "Application-defined or object-defined error";
        (errFunc as any).__vbaAutoCall__ = true;
        this.env.set('error', errFunc);
        this.env.set('error$', errFunc);
        this.env.set('vbsunday', 1);
        this.env.set('vbmonday', 2);
        this.env.set('vbtuesday', 3);
        this.env.set('vbwednesday', 4);
        this.env.set('vbthursday', 5);
        this.env.set('vbfriday', 6);
        this.env.set('vbsaturday', 7);
        this.env.set('vbusesystem', 0);
        this.env.set('vbbinarycompare', 0);
        this.env.set('vbtextcompare', 1);
        this.env.set('vbempty', 0); this.env.set('vbnull', 1); this.env.set('vbinteger', 2); this.env.set('vblong', 3); this.env.set('vbsingle', 4); this.env.set('vbdouble', 5); this.env.set('vbcurrency', 6); this.env.set('vbdate', 7); this.env.set('vbstring', 8); this.env.set('vbobject', 9); this.env.set('vberror', 10); this.env.set('vbboolean', 11); this.env.set('vbvariant', 12); this.env.set('vbbyte', 17); this.env.set('vblonglong', 20); this.env.set('vbarray', 8192);
        this.env.set('vbcrlf', "\r\n"); this.env.set('vbtab', "\t"); this.env.set('vbcr', "\r"); this.env.set('vblf', "\n");
        this.env.set('true', vbaTrue); this.env.set('false', vbaFalse); this.env.set('empty', vbaEmpty); this.env.set('nothing', vbaNothing); this.env.set('null', vbaNull);

        this.env.set('environ', (k: any) => this.sandbox.getEnv(k));
        this.env.set('environ$', this.env.get('environ'));
        this.env.set('createobject', (id: string) => this.createExternalObject(id));
        this.env.set('getobject', (pathname?: string, classId?: string) => {
            if (pathname) {
                return {
                    __vbaTypeName__: 'Object',
                    Path: pathname,
                    Name: pathname.split(/[\\\/]/).pop()
                };
            }
            if (classId) {
                return this.createExternalObject(classId);
            }
            return vbaNothing;
        });
        this.env.set('iif', (c: any, t: any, f: any) => this.isTrue(c) ? t : f);
        this.env.set('choose', (i: any, ...c: any[]) => { const idx = Math.floor(Number(i)); return (idx >= 1 && idx <= c.length) ? c[idx - 1] : vbaNull; });
        this.env.set('switch', (...args: any[]) => { for (let i = 0; i < args.length; i += 2) if (this.isTrue(args[i])) return args[i + 1]; return vbaNull; });
        this.env.set('array', (...args: any[]) => { const a = [...args]; (a as any).vbaBase = 0; return a; });
        this.env.set('lbound', (a: any, dim: any = 1) => {
            if (!Array.isArray(a)) return 0;
            const dimIndex = Number(dim) - 1;
            if ((a as any).__vbaDimensions__ && dimIndex >= 0 && dimIndex < (a as any).__vbaDimensions__.length) {
                return (a as any).__vbaDimensions__[dimIndex].lower;
            }
            let current = a;
            for (let i = 0; i < dimIndex; i++) {
                if (Array.isArray(current) && current.length > 0) current = current[0];
                else this.throwVbaError(9, "Subscript out of range");
            }
            return (current as any).vbaBase || 0;
        });
        this.env.set('ubound', (a: any, dim: any = 1) => {
            if (!Array.isArray(a)) return 0;
            const dimIndex = Number(dim) - 1;
            if ((a as any).__vbaDimensions__ && dimIndex >= 0 && dimIndex < (a as any).__vbaDimensions__.length) {
                return (a as any).__vbaDimensions__[dimIndex].upper;
            }
            let current = a;
            for (let i = 0; i < dimIndex; i++) {
                if (Array.isArray(current) && current.length > 0) current = current[0];
                else this.throwVbaError(9, "Subscript out of range");
            }
            return ((current as any).vbaBase || 0) + current.length - 1;
        });

        // --- Registry Module ---
        this.env.set('savesetting', (app: string, sec: string, key: string, val: any) => {
            if (!this.virtualRegistry[app]) this.virtualRegistry[app] = {};
            if (!this.virtualRegistry[app][sec]) this.virtualRegistry[app][sec] = {};
            this.virtualRegistry[app][sec][key] = String(val);
        });
        this.env.set('getsetting', (app: string, sec: string, key: string, def: any = "") => {
            return (this.virtualRegistry[app]?.[sec]?.[key]) ?? String(def);
        });
        this.env.set('getallsettings', (app: string, sec: string) => {
            const s = this.virtualRegistry[app]?.[sec];
            if (!s) return vbaEmpty;
            const res = Object.entries(s).map(([k, v]) => [k, v]);
            (res as any).vbaBase = 0;
            return res;
        });
        this.env.set('deletesetting', (app: string, sec?: string, key?: string) => {
            if (!sec) delete this.virtualRegistry[app];
            else if (!key) delete this.virtualRegistry[app]?.[sec];
            else delete this.virtualRegistry[app]?.[sec]?.[key];
        });
    }

    private print(msg: string) {
        this.onPrint(msg);
    }

    private triggerTerminate(obj: any) {
        if (obj && obj.__vbaClass__) {
            const classDef = obj.__classDef__ as ClassDeclaration;
            const terminateProc = classDef.procedures.find(p => p.name.name.toLowerCase() === 'class_terminate');
            if (terminateProc) {
                try {
                    this.callClassMethod(obj, terminateProc, []);
                } catch (e) {
                    // Errors in Terminate are typically suppressed or handled specially in VBA, 
                    // but for now let's just log or ignore to prevent crashing the lifecycle
                    console.error("Error in Class_Terminate:", e);
                }
            }
        }
    }

    private instantiateType(typeName: string): any {
        const typeMembers = this.env.getType(typeName);
        if (!typeMembers) return 0;

        const instance: any = {
            __vbaTypeName__: typeName
        };
        for (const member of typeMembers) {
            const mt = member.memberType;
            const mtLower = mt.toLowerCase();
            if (mtLower === 'string') {
                instance[member.name.toLowerCase()] = '';
            } else if (mtLower === 'boolean') {
                instance[member.name.toLowerCase()] = 0; // vbaFalse
            } else if (this.env.getType(mt)) {
                instance[member.name.toLowerCase()] = this.instantiateType(mt);
            } else {
                instance[member.name.toLowerCase()] = 0;
            }
        }
        return instance;
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

            if (param.isParamArray) {
                // Capture all remaining arguments into a Variant array
                const remainingArgs = args.slice(i);
                (remainingArgs as any).vbaBase = 0; // ParamArray is always 0-based
                localEnv.setLocally(paramName, remainingArgs);
                break;
            }

            let argValue: any;
            if (i < args.length) {
                argValue = args[i];
            } else if (param.defaultValue) {
                argValue = this.evaluateExpression(param.defaultValue);
            } else {
                argValue = (param.isOptional ? vbaMissing : vbaEmpty);
            }
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
        this.lastErrorIndex = null;

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
            case 'OpenStatement':
                this.evaluateOpenStatement(stmt as OpenStatement);
                break;
            case 'CloseStatement':
                this.evaluateCloseStatement(stmt as CloseStatement);
                break;
            case 'PrintStatement':
                this.evaluatePrintStatement(stmt as PrintStatement);
                break;
            case 'LineInputStatement':
                this.evaluateLineInputStatement(stmt as LineInputStatement);
                break;
            case 'PutStatement':
                this.evaluatePutStatement(stmt as PutStatement);
                break;
            case 'KillStatement':
                this.evaluateKillStatement(stmt as KillStatement);
                break;
            case 'WriteStatement':
                this.evaluateWriteStatement(stmt as WriteStatement);
                break;
            case 'InputStatement':
                this.evaluateInputStatement(stmt as InputStatement);
                break;
            case 'GetStatement':
                this.evaluateGetStatement(stmt as GetStatement);
                break;
            case 'SeekStatement':
                this.evaluateSeekStatement(stmt as SeekStatement);
                break;
            case 'ResetStatement':
                this.evaluateResetStatement(stmt as ResetStatement);
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
            case 'EventDeclaration':
                this.evaluateEventDeclaration(stmt as EventDeclaration);
                break;
            case 'RaiseEventStatement':
                this.evaluateRaiseEventStatement(stmt as RaiseEventStatement);
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
            case 'ImplementsDirective':
                this.evaluateImplementsDirective(stmt as ImplementsDirective);
                break;
            case 'AppActivateStatement':
                this.evaluateAppActivateStatement(stmt as AppActivateStatement);
                break;
            case 'SendKeysStatement':
                this.evaluateSendKeysStatement(stmt as SendKeysStatement);
                break;
            case 'LockStatement':
                this.evaluateLockStatement(stmt as LockStatement);
                break;
            case 'UnlockStatement':
                this.evaluateUnlockStatement(stmt as UnlockStatement);
                break;
            case 'WidthStatement':
                this.evaluateWidthStatement(stmt as WidthStatement);
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

    private toVbaNumber(val: any): number {
        if (val === vbaEmpty || val === undefined) return 0;
        if (val === vbaNull || val === null) this.throwVbaError(13, "Type mismatch");
        if (val instanceof VbaBoolean) return val.value;
        if (val instanceof VbaDate) return val.value;
        if (val instanceof VbaDecimal) return val.value;
        if (typeof val === 'number') return val;
        if (typeof val === 'bigint') return Number(val);
        if (typeof val === 'string') {
            const n = parseFloat(val);
            if (isNaN(n)) this.throwVbaError(13, "Type mismatch");
            return n;
        }
        if (val instanceof VbaErrorValue) this.throwVbaError(13, "Type mismatch");
        return Number(val);
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

    private evaluateEventDeclaration(stmt: EventDeclaration) {
        // Top-level event declarations are just metadata for the class
    }

    private evaluateRaiseEventStatement(stmt: RaiseEventStatement) {
        const eventName = stmt.eventName.name.toLowerCase();
        // Look for the instance in the current environment (Me)
        const me = this.env.get('Me');
        if (me && me.__events__) {
            const handlers = me.__events__.get(eventName);
            if (handlers) {
                const args = stmt.args.map(a => this.evaluateExpression(a));
                for (const handler of handlers) {
                    handler(...args);
                }
            }
        }
    }

    private evaluateImplementsDirective(stmt: ImplementsDirective) {
        // No-op for now. Used for interface compliance metadata.
    }

    private evaluateAppActivateStatement(stmt: AppActivateStatement) {
        const title = String(this.evaluateExpression(stmt.title));
        console.log(`[STUB] AppActivate "${title}"`);
    }

    private evaluateSendKeysStatement(stmt: SendKeysStatement) {
        const keys = String(this.evaluateExpression(stmt.keys));
        console.log(`[STUB] SendKeys "${keys}"`);
    }

    private evaluateLockStatement(stmt: LockStatement) {
        const fileNum = Number(this.evaluateExpression(stmt.fileNumber));
        console.log(`[STUB] Lock #${fileNum}`);
    }

    private evaluateUnlockStatement(stmt: UnlockStatement) {
        const fileNum = Number(this.evaluateExpression(stmt.fileNumber));
        console.log(`[STUB] Unlock #${fileNum}`);
    }

    private evaluateWidthStatement(stmt: WidthStatement) {
        const fileNum = Number(this.evaluateExpression(stmt.fileNumber));
        const width = Number(this.evaluateExpression(stmt.width));
        console.log(`[STUB] Width #${fileNum}, ${width}`);
    }

    private evaluateAssignmentStatement(stmt: AssignmentStatement) {
        const val = this.evaluateExpression(stmt.right);
        this.evaluateAssignmentToVariable(stmt.left, val);
    }

    private evaluateAssignmentToVariable(left: Expression, val: any) {
        if (left.type === 'Identifier') {
            const name = (left as Identifier).name;
            const oldVal = this.env.get(name);
            if (oldVal !== val) this.triggerTerminate(oldVal);

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
        } else if (left.type === 'CallExpression') {
            // Array/Dictionary assignment: arr(0) = val OR dict("key") = val
            const call = left as CallExpression;
            if (call.callee.type === 'Identifier') {
                const name = (call.callee as Identifier).name;
                const lowerName = name.toLowerCase();

                if (lowerName === 'mid' || lowerName === 'mid$' || lowerName === 'midb' || lowerName === 'midb$') {
                    // Mid(s, start, [len]) = val
                    const targetName = (call.args[0] as Identifier).name; // Must be identifier for assignment
                    let start = this.evaluateExpression(call.args[1]) as number;
                    let length = call.args.length > 2 ? this.evaluateExpression(call.args[2]) as number : -1;
                    
                    const isByte = lowerName.startsWith('midb');
                    if (isByte) {
                        // Convert byte position to char position (approximate: 2 bytes per char)
                        start = Math.floor((start + 1) / 2);
                        if (length !== -1) length = Math.floor(length / 2);
                    }

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
                            (current[d] as any).vbaBase = base;
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
        } else if (left.type === 'MemberExpression') {
            const member = left as MemberExpression;
            const obj = this.evaluateExpression(member.object);
            const propName = member.property.name.toLowerCase();
            if (obj && obj.__vbaClass__) {
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
        } else if (left.type === 'ImplicitWithObjectExpression') {
            if (this.withObjectStack.length === 0) {
                throw new Error(`Execution error: '.' outside of With block`);
            }
            const obj = this.withObjectStack[this.withObjectStack.length - 1];
            const member = left as ImplicitWithObjectExpression;
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
            if (decl.isWithEvents) {
                this.env.setWithEvents(varName);
            }
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
                if (decl.arrayBounds && decl.arrayBounds.length > 0) {
                    initialValue = this.createMultiDimArray(decl.arrayBounds, initialValue);
                    (initialValue as any).vbaFixed = true;
                } else {
                    initialValue = [];
                    (initialValue as any).vbaBase = this.arrayBase;
                    (initialValue as any).vbaFixed = false;
                }
            } else if (decl.isNew && decl.objectType === 'Collection') {
                initialValue = new VbaCollection();
            } else if (decl.isNew && decl.objectType && this.classDefinitions.has(decl.objectType.toLowerCase())) {
                initialValue = this.instantiateClass(decl.objectType);
            } else if (decl.objectType) {
                const t = decl.objectType.toLowerCase();
                if (!['integer', 'long', 'single', 'double', 'currency', 'byte', 'string', 'boolean'].includes(t)) {
                    initialValue = this.instantiateType(decl.objectType);
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

    public registerClass(name: string, classDef: ClassDeclaration) {
        this.classDefinitions.set(name.toLowerCase(), classDef);
    }

    private evaluateClassDeclaration(stmt: ClassDeclaration) {
        this.registerClass(stmt.name, stmt);
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
            __vbaTypeName__: classDef.name,
            __classDef__: classDef,
            __instanceEnv__: instanceEnv,
            __events__: new Map<string, ((...args: any[]) => void)[]>(),
        };

        // Initialize Events
        for (const stmt of classDef.body) {
            if (stmt.type === 'EventDeclaration') {
                const eventDecl = stmt as EventDeclaration;
                instance.__events__.set(eventDecl.name.name.toLowerCase(), []);
            }
        }

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
            const param = proc.parameters[i];
            const paramName = param.name;

            if (param.isParamArray) {
                const remainingArgs = args.slice(i);
                (remainingArgs as any).vbaBase = 0;
                localEnv.setLocally(paramName, remainingArgs);
                break;
            }

            let argValue: any;
            if (i < args.length) {
                argValue = args[i];
            } else if (param.defaultValue) {
                argValue = this.evaluateExpression(param.defaultValue);
            } else {
                argValue = (param.isOptional ? vbaMissing : vbaEmpty);
            }
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
        this.env.setConstant(stmt.name.name, value);
    }

    private evaluateSetStatement(stmt: SetStatement) {
        let value = this.evaluateExpression(stmt.right);

        // VBA requires Set target to be an object (or Nothing)
        if (value !== null && value !== vbaNothing && typeof value !== 'object') {
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

            const oldVal = this.env.get(name);
            if (oldVal !== value) this.triggerTerminate(oldVal);

            const proc = this.env.getProcedure(name, 'set');
            if (proc) {
                this.callProcedure(name, [value], 'set');
                return;
            }
            this.env.set(name, value);

            // Check for WithEvents binding
            if (this.env.isWithEvents(name) && value && value.__events__) {
                // Binding: look for VarName_EventName subs in the current environment
                for (const eventName of value.__events__.keys()) {
                    const handlerName = `${name}_${eventName}`;
                    const handler = this.env.getProcedure(handlerName);
                    if (handler) {
                        const handlers = value.__events__.get(eventName);
                        handlers.push((...args: any[]) => {
                            this.callProcedure(handlerName, args);
                        });
                    }
                }
            }
        } else {
            throw new Error(`Execution error: Unsupported Set target ${stmt.left.type}`);
        }
    }

    private evaluateResumeStatement(stmt: ResumeStatement) {
        const target = (stmt.target || '').toLowerCase().trim();
        if (target === '') {
            throw { type: 'Resume', mode: 'Current' };
        } else if (target === 'next') {
            throw { type: 'Resume', mode: 'Next' };
        } else {
            throw { type: 'Resume', mode: 'Label', label: stmt.target };
        }
    }

    private evaluateOnErrorStatement(stmt: OnErrorStatement) {
        const label = (stmt.label || '').toLowerCase().trim();
        if (label === '0') {
            this.errorHandlerLabel = null;
            this.errorHandlingMode = 'None';
        } else if (label === '-1') {
            this.isInErrorHandler = false;
        } else if (label === 'resume next') {
            this.errorHandlerLabel = null;
            this.errorHandlingMode = 'ResumeNext';
        } else {
            this.errorHandlerLabel = stmt.label;
            this.errorHandlingMode = 'GoTo';
        }
    }

    private evaluateOpenStatement(stmt: OpenStatement) {
        const vbaPath = String(this.evaluateExpression(stmt.path));
        const realPath = this.sandbox.toRealPath(vbaPath);
        const fileNum = Number(this.evaluateExpression(stmt.fileNumber));

        if (this.fileHandles.has(fileNum)) {
            this.throwVbaError(55, "File already open");
        }

        let flags = '';
        switch (stmt.mode) {
            case 'Input': flags = 'r'; break;
            case 'Output': flags = 'w'; break;
            case 'Append': flags = 'a'; break;
            case 'Random':
            case 'Binary': 
                if (!fs.existsSync(realPath)) {
                    fs.writeFileSync(realPath, "");
                }
                flags = 'r+'; break;
        }

        try {
            // Ensure directory exists for write modes
            if (flags === 'w' || flags === 'a') {
                const dir = path.dirname(realPath);
                if (!fs.existsSync(dir)) {
                    fs.mkdirSync(dir, { recursive: true });
                }
            } else if (flags === 'r' && !fs.existsSync(realPath)) {
                this.throwVbaError(53, "File not found");
            }

            const fd = fs.openSync(realPath, flags);
            this.fileHandles.set(fileNum, {
                fd,
                mode: stmt.mode,
                path: realPath,
                pos: 0
            });
        } catch (e: any) {
            if (e.code === 'ENOENT') this.throwVbaError(53, "File not found");
            if (e.code === 'EACCES') this.throwVbaError(75, "Path/File access error");
            throw e;
        }
    }

    private evaluateCloseStatement(stmt: CloseStatement) {
        const nums = stmt.fileNumbers.length > 0 
            ? stmt.fileNumbers.map(n => Number(this.evaluateExpression(n)))
            : Array.from(this.fileHandles.keys());

        for (const num of nums) {
            const handle = this.fileHandles.get(num);
            if (handle) {
                fs.closeSync(handle.fd);
                this.fileHandles.delete(num);
            }
        }
    }

    private evaluatePrintStatement(stmt: PrintStatement) {
        const fileNum = Number(this.evaluateExpression(stmt.fileNumber));
        const handle = this.fileHandles.get(fileNum);
        if (!handle) this.throwVbaError(52, "Bad file name or number");
        if (handle.mode === 'Input') this.throwVbaError(54, "Bad file mode");

        let output = "";
        for (const expr of stmt.expressions) {
            if (expr === 'Comma') {
                const currentLen = output.length;
                const target = Math.ceil((currentLen + 1) / 14) * 14;
                output += " ".repeat(target - currentLen);
            } else if (expr === 'Semicolon') {
                // Continue
            } else if (typeof expr === 'object' && expr !== null && 'type' in expr) {
                 if (expr.type === 'Spc') {
                     const n = Number(this.evaluateExpression((expr as any).val));
                     output += " ".repeat(Math.max(0, n));
                 } else if (expr.type === 'Tab') {
                     const n = Number(this.evaluateExpression((expr as any).val));
                     output += " ".repeat(Math.max(0, n - output.length));
                 } else {
                     output += String(this.evaluateExpression(expr as any));
                 }
            } else {
                const val = this.evaluateExpression(expr as any);
                output += String(val === vbaNull ? "Null" : (val === vbaEmpty ? "" : val));
            }
        }
        
        const last = stmt.expressions[stmt.expressions.length - 1];
        if (last !== 'Semicolon' && last !== 'Comma') {
            output += "\r\n";
        }

        fs.writeSync(handle.fd, output);
        handle.pos! += output.length;
    }

    private evaluateLineInputStatement(stmt: LineInputStatement) {
        const fileNum = Number(this.evaluateExpression(stmt.fileNumber));
        const handle = this.fileHandles.get(fileNum);
        if (!handle) this.throwVbaError(52, "Bad file name or number");
        
        const buffer = Buffer.alloc(1);
        let line = "";
        let bytesRead = 0;

        while (true) {
            bytesRead = fs.readSync(handle.fd, buffer, 0, 1, handle.pos);
            if (bytesRead === 0) break;
            const char = buffer.toString('utf8', 0, 1);
            handle.pos!++;
            if (char === '\n') break;
            if (char !== '\r') line += char;
        }

        this.env.setLocally(stmt.variable.name, line);
    }

    private evaluatePutStatement(stmt: PutStatement) {
        const fileNum = Number(this.evaluateExpression(stmt.fileNumber));
        const handle = this.fileHandles.get(fileNum);
        if (!handle) this.throwVbaError(52, "Bad file name or number");

        const data = this.evaluateExpression(stmt.data);
        const s = String(data);
        const buffer = Buffer.from(s);
        
        let position: number | null = handle.pos;
        if (stmt.recordNumber) {
            position = (Number(this.evaluateExpression(stmt.recordNumber)) - 1);
        }

        fs.writeSync(handle.fd, buffer, 0, buffer.length, position);
        handle.pos! += buffer.length;
    }

    private evaluateKillStatement(stmt: KillStatement) {
        const vbaPath = String(this.evaluateExpression(stmt.path));
        const realPath = this.sandbox.toRealPath(vbaPath);
        fs.unlinkSync(realPath);
    }

    private evaluateWriteStatement(stmt: WriteStatement) {
        const fileNum = Number(this.evaluateExpression(stmt.fileNumber));
        const handle = this.fileHandles.get(fileNum);
        if (!handle) throw new Error(`Execution error: File number ${fileNum} not open`);

        const output = stmt.items.map(item => {
            const val = this.evaluateExpression(item);
            if (typeof val === 'string') return `"${val}"`;
            if (val instanceof VbaDate) return `#${val.toString()}#`;
            if (val === null) return "#NULL#";
            return String(val);
        }).join(",");

        const lineOutput = output + "\n";
        fs.writeSync(handle.fd, lineOutput);
        handle.pos! += lineOutput.length;
    }

    private evaluateInputStatement(stmt: InputStatement) {
        const fileNum = Number(this.evaluateExpression(stmt.fileNumber));
        const handle = this.fileHandles.get(fileNum);
        if (!handle) throw new Error(`Execution error: File number ${fileNum} not open`);

        // Simple line-based implementation for now.
        // Real VBA Input # parses delimiters even across lines.
        let content = "";
        const buf = Buffer.alloc(1024);
        let bytesRead = 0;
        let readPos = handle.pos || 0;
        while ((bytesRead = fs.readSync(handle.fd, buf, 0, 1024, readPos)) > 0) {
            content += buf.toString('utf8', 0, bytesRead);
            if (content.includes('\n')) break;
            readPos += bytesRead;
        }
        
        const lineEnd = content.indexOf('\n');
        const line = lineEnd === -1 ? content : content.slice(0, lineEnd);
        handle.pos = (handle.pos || 0) + line.length + (lineEnd === -1 ? 0 : 1);
        if (content[lineEnd - 1] === '\r') {
            // handle CRLF
        }
        
        const values = line.trim().split(",").map(v => v.trim().replace(/^"|"$/g, ''));
        for (let i = 0; i < stmt.variables.length; i++) {
            if (i < values.length) {
                this.evaluateAssignmentToVariable(stmt.variables[i], values[i]);
            }
        }
    }

    private evaluateGetStatement(stmt: GetStatement) {
        const fileNum = Number(this.evaluateExpression(stmt.fileNumber));
        const handle = this.fileHandles.get(fileNum);
        if (!handle) throw new Error(`Execution error: File number ${fileNum} not open`);

        // Basic implementation: read up to 1024 bytes or until EOF
        const buffer = Buffer.alloc(1024);
        let position: number | null = handle.pos;
        if (stmt.recordNumber) {
            position = (Number(this.evaluateExpression(stmt.recordNumber)) - 1);
        }
        
        const bytesRead = fs.readSync(handle.fd, buffer, 0, buffer.length, position);
        const s = buffer.toString('utf8', 0, bytesRead);
        this.evaluateAssignmentToVariable(stmt.variable, s);
        handle.pos! += bytesRead;
    }

    private evaluateSeekStatement(stmt: SeekStatement) {
        const fileNum = Number(this.evaluateExpression(stmt.fileNumber));
        const handle = this.fileHandles.get(fileNum);
        if (!handle) throw new Error(`Execution error: File number ${fileNum} not open`);

        const pos = Number(this.evaluateExpression(stmt.position));
        // Node doesn't have seekSync on FD directly without lseek, 
        // but we can track it in our handle if we use it for subsequent read/write.
        handle.pos = Math.max(0, pos - 1); 
    }

    private evaluateResetStatement(stmt: ResetStatement) {
        for (const [num, handle] of this.fileHandles) {
            fs.closeSync(handle.fd);
        }
        this.fileHandles.clear();
    }

    private evaluateAttributeStatement(stmt: AttributeStatement) {
        if (stmt.name.toLowerCase() === 'vb_name') {
            const val = String(this.evaluateExpression(stmt.value)).replace(/"/g, '');
            this.executingModuleName = val;
        }
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
                fileexists: (p: string) => {
                    try {
                        const full = this.sandbox.toRealPath(p);
                        return (fs.existsSync(full) && fs.statSync(full).isFile()) ? vbaTrue : vbaFalse;
                    } catch { return vbaFalse; }
                },
                folderexists: (p: string) => {
                    try {
                        const full = this.sandbox.toRealPath(p);
                        return (fs.existsSync(full) && fs.statSync(full).isDirectory()) ? vbaTrue : vbaFalse;
                    } catch { return vbaFalse; }
                },
                createtextfile: (p: string, overwrite: boolean = true) => {
                    const full = this.sandbox.toRealPath(p);
                    if (!overwrite && fs.existsSync(full)) this.throwVbaError(58, "File already exists");
                    const fd = fs.openSync(full, 'w');
                    return {
                        write: (s: string) => fs.writeSync(fd, s),
                        writeline: (s: string) => fs.writeSync(fd, s + "\r\n"),
                        close: () => fs.closeSync(fd)
                    };
                },
                opentextfile: (p: string, iomode: number = 1) => {
                    const full = this.sandbox.toRealPath(p);
                    const mode = iomode === 1 ? 'r' : (iomode === 2 ? 'w' : 'a');
                    const fd = fs.openSync(full, mode);
                    let pos = 0;
                    return {
                        readall: () => {
                            const content = fs.readFileSync(full, 'utf8');
                            return content;
                        },
                        readline: () => {
                            // Simple readline implementation
                            const content = fs.readFileSync(full, 'utf8');
                            const lines = content.slice(pos).split(/\r?\n/);
                            if (lines.length > 0) {
                                const line = lines[0];
                                pos += line.length + (content[pos + line.length] === '\r' ? 2 : 1);
                                return line;
                            }
                            return "";
                        },
                        write: (s: string) => fs.writeSync(fd, s),
                        close: () => fs.closeSync(fd)
                    };
                },
                createfolder: (p: string) => {
                    const full = this.sandbox.toRealPath(p);
                    fs.mkdirSync(full, { recursive: true });
                    return { path: p };
                },
                deletefile: (p: string) => {
                    const full = this.sandbox.toRealPath(p);
                    fs.unlinkSync(full);
                },
                deletefolder: (p: string) => {
                    const full = this.sandbox.toRealPath(p);
                    fs.rmSync(full, { recursive: true, force: true });
                },
                getfile: (p: string) => {
                    const full = this.sandbox.toRealPath(p);
                    const stats = fs.statSync(full);
                    return {
                        path: p,
                        size: stats.size,
                        datecreated: new VbaDate(toVbaDate(stats.birthtime)),
                        datelastmodified: new VbaDate(toVbaDate(stats.mtime)),
                        attributes: stats.mode
                    };
                },
                getfolder: (p: string) => ({ path: p }),
                getbasename: (p: string) => path.basename(p),
                getextensionname: (p: string) => path.extname(p).replace('.', ''),
                getparentfoldername: (p: string) => path.dirname(p),
                getabsolutepathname: (p: string) => p // already handled by sandbox resolve internally
            };
        } else if (id === 'msxml2.xmlhttp' || id === 'microsoft.xmlhttp') {
            let responseText = "";
            let status = 0;
            return {
                open: (method: string, url: string, async: boolean = true) => {
                    this.onPrint(`[XMLHTTP] Open: ${method} ${url}`);
                },
                send: (body?: any) => {
                    this.onPrint(`[XMLHTTP] Send`);
                    responseText = "Mock Response";
                    status = 200;
                },
                setrequestheader: (h: string, v: string) => { /* Mock */ },
                getresponsetext: () => responseText,
                responsetext: responseText, // Property access
                getstatus: () => status,
                status: status,
                readystate: 4
            };
        } else if (id === 'adodb.stream') {
            let content = "";
            let streamPos = 0;
            return {
                open: () => { streamPos = 0; },
                close: () => { },
                write: (data: any) => { content += String(data); },
                writetext: (text: string) => { content += text; },
                read: (len: number) => { const r = content.slice(streamPos, streamPos + len); streamPos += len; return r; },
                readtext: () => { const r = content.slice(streamPos); streamPos = content.length; return r; },
                savetofile: (p: string, mode: number = 1) => {
                    const full = this.sandbox.toRealPath(p);
                    fs.writeFileSync(full, content);
                },
                loadfromfile: (p: string) => {
                    const full = this.sandbox.toRealPath(p);
                    content = fs.readFileSync(full, 'utf8');
                    streamPos = 0;
                },
                type: 2, // adTypeText
                charset: 'utf-8',
                position: streamPos,
                size: content.length
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
                if (e && (e.type === 'Exit' || e.type === 'Terminate' || e.type === 'ProcedureReturn')) {
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
                    throw new Error(`Execution error: Label '${e.label}' not found`);
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
                    if (e.mode === 'Current') {
                        i = this.lastErrorIndex !== null ? this.lastErrorIndex : i;
                    } else if (e.mode === 'Next') {
                        i = this.lastErrorIndex !== null ? this.lastErrorIndex + 1 : i + 1;
                    } else if (e.mode === 'Label') {
                        const labelName = e.label.toLowerCase();
                        const labelIndex = body.findIndex(s =>
                            s.type === 'LabelStatement' &&
                            (s as any).label.toLowerCase() === labelName
                        );
                        if (labelIndex >= 0) {
                            i = labelIndex;
                        } else {
                            throw new Error(`Execution error: Label '${e.label}' not found for Resume`);
                        }
                    }
                    this.errObj.clear();
                    continue;
                }

                // Handle VbaError or standard JS Error
                if (!this.isInErrorHandler) {
                    const errorNumber = e.type === 'VbaError' ? e.number : 1000;
                    const errorMessage = e.message || String(e);
                    
                    this.errObj.number = errorNumber;
                    this.errObj.description = errorMessage;

                    if (this.errorHandlingMode === 'ResumeNext') {
                        this.lastErrorIndex = i;
                        i++;
                        continue;
                    }

                    if (this.errorHandlingMode === 'GoTo' && this.errorHandlerLabel) {
                        const labelName = this.errorHandlerLabel.toLowerCase();
                        const labelIndex = body.findIndex(s =>
                            s.type === 'LabelStatement' &&
                            (s as any).label.toLowerCase() === labelName
                        );
                        if (labelIndex >= 0) {
                            this.lastErrorIndex = i;
                            this.isInErrorHandler = true;
                            i = labelIndex;
                            continue;
                        }
                    }
                }

                // If we reach here, either we are already in an error handler,
                // or there's no handler configured. Bubble up.
                this.isInErrorHandler = false;
                throw e;
            }
        }
    }

    private evaluateEraseStatement(stmt: EraseStatement) {
        const varName = stmt.name.name;
        const arr = this.env.get(varName);
        if (Array.isArray(arr)) {
            if ((arr as any).vbaFixed) {
                // Fixed array: re-initialize elements
                const defaultValue = (arr as any).__vbaDefaultValue__ ?? 0;
                this.reinitializeArray(arr, defaultValue);
            } else {
                // Dynamic array: de-allocate
                const newArr: any[] = [];
                (newArr as any).vbaBase = (arr as any).vbaBase ?? this.arrayBase;
                (newArr as any).vbaFixed = false;
                this.env.set(varName, newArr);
            }
        }
    }

    private reinitializeArray(arr: any[], defaultValue: any) {
        for (let i = 0; i < arr.length; i++) {
            if (Array.isArray(arr[i]) && !(arr[i] as any).__vbaClass__) {
                this.reinitializeArray(arr[i], defaultValue);
            } else {
                arr[i] = defaultValue;
            }
        }
    }

    private createMultiDimArray(bounds: ArrayBound[], initialValue: any): any[] {
        const dimensions: { lower: number, upper: number }[] = [];
        
        for (const bound of bounds) {
            const upper = Number(this.evaluateExpression(bound.upper));
            const lower = bound.lower ? Number(this.evaluateExpression(bound.lower)) : this.arrayBase;
            dimensions.push({ lower, upper });
        }

        const buildArray = (dimIdx: number): any[] => {
            const { lower, upper } = dimensions[dimIdx];
            const count = upper - lower + 1;
            if (count < 0) throw new Error("Execution error: Subscript out of range");

            const arr = new Array(count);
            if (dimIdx < dimensions.length - 1) {
                for (let i = 0; i < count; i++) {
                    arr[i] = buildArray(dimIdx + 1);
                }
            } else {
                arr.fill(initialValue);
            }
            return arr;
        };

        const result = buildArray(0);
        (result as any).__vbaDimensions__ = dimensions;
        (result as any).vbaBase = dimensions[0].lower; // Fallback for 1D code
        (result as any).__vbaDefaultValue__ = initialValue;
        return result;
    }

    private copyPreservedData(oldArr: any, newArr: any, dimensions: { lower: number, upper: number }[], dimIdx: number = 0) {
        if (!Array.isArray(oldArr) || !Array.isArray(newArr)) return;

        const copyLen = Math.min(oldArr.length, newArr.length);

        if (dimIdx < dimensions.length - 1) {
            for (let i = 0; i < copyLen; i++) {
                this.copyPreservedData(oldArr[i], newArr[i], dimensions, dimIdx + 1);
            }
        } else {
            for (let i = 0; i < copyLen; i++) {
                newArr[i] = oldArr[i];
            }
        }
    }

    private evaluateReDimStatement(stmt: ReDimStatement) {
        const varName = stmt.name.name;
        const oldArr = this.env.get(varName);
        
        let defaultValue: any = 0;
        if (stmt.objectType) {
            const t = stmt.objectType.toLowerCase();
            if (t === 'string') defaultValue = '';
            else if (t === 'boolean') defaultValue = 0;
        } else if (Array.isArray(oldArr)) {
            defaultValue = (oldArr as any).__vbaDefaultValue__ ?? 0;
        }

        if (stmt.bounds.length > 0) {
            const arr = this.createMultiDimArray(stmt.bounds, defaultValue);
            (arr as any).vbaFixed = false;

            if (stmt.isPreserve && Array.isArray(oldArr)) {
                this.copyPreservedData(oldArr, arr, (arr as any).__vbaDimensions__);
            }

            this.env.set(varName, arr);
        }
    }

    private evaluateExitStatement(stmt: ExitStatement) {
        throw { type: 'Exit', target: stmt.exitType };
    }

    private toNumber(val: any): number {
        if (val === vbaNull || val === vbaEmpty || val === vbaNothing || val === vbaMissing) {
            // Symbols cannot be converted to number directly in JS
            return 0;
        }
        if (val instanceof VbaBoolean) return val.value;
        if (val instanceof VbaDate) return val.value;
        if (val instanceof VbaDecimal) return val.value;
        if (typeof val === 'bigint') return Number(val);
        const n = parseFloat(val);
        return isNaN(n) ? 0 : n;
    }

    private toDisplayString(val: any): string {
        if (val === vbaEmpty || val === undefined) return "";
        if (val === vbaNull) return "Null";
        if (val === vbaNothing) return "Nothing";
        if (val === vbaTrue || (val instanceof VbaBoolean && val.value === -1)) return "True";
        if (val === vbaFalse || (val instanceof VbaBoolean && val.value === 0)) return "False";
        if (val && typeof val === 'object' && (val instanceof VbaDate || (val as any).__isVbaDate__)) {
            return val.toString();
        }
        if (val instanceof VbaErrorValue) return val.toString();
        if (val instanceof VbaDecimal) return val.toString();
        if (typeof val === 'bigint') return val.toString();
        return String(val);
    }

    private evaluateExpression(expr: Expression): any {
        switch (expr.type) {
            case 'NumberLiteral':
                return (expr as NumberLiteral).value;
            case 'StringLiteral':
                return (expr as StringLiteral).value;
            case 'AddressOfExpression':
                return (expr as AddressOfExpression).procedureName.name;
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
                    // Only auto-call if it's a Function/Property and has 0 required arguments
                    const requiredCount = p.parameters.filter(param => !param.isOptional && !param.isParamArray).length;
                    if (requiredCount === 0) {
                        return this.callProcedure(idName, []);
                    }
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
            case 'ParenthesizedExpression':
                return this.evaluateExpression((expr as ParenthesizedExpression).expression);
            case 'NewExpression':
                return this.instantiateClass((expr as NewExpression).className);
            default:
                throw new Error(`Execution error: Unknown expression type ${expr.type}`);
        }
    }

    private evaluateDateLiteral(expr: DateLiteral): any {
        // Ensure UTC for consistency in serial numbers
        let d = new Date(expr.value);
        if (isNaN(d.getTime())) {
            throw new Error(`Execution error: Invalid date literal #${expr.value}#`);
        }
        // If it's a date-only string like "yyyy-mm-dd", JS Date(string) might treat it as UTC already,
        // but it's better to be explicit.
        if (/^\d{4}-\d{2}-\d{2}$/.test(expr.value)) {
            d = new Date(expr.value + "T00:00:00Z");
        }
        return new VbaDate(toVbaDate(d));
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
                const byRefArgs: { paramName: string, originalExpr: Expression }[] = [];
                const namedArgs = new Map<string, any>();
                const positionalArgs: any[] = [];

                for (const argExpr of expr.args) {
                    if (argExpr.type === 'NamedArgument') {
                        const namedArg = argExpr as NamedArgument;
                        namedArgs.set(namedArg.name.toLowerCase(), this.evaluateExpression(namedArg.value));
                    } else {
                        positionalArgs.push(this.evaluateExpression(argExpr));
                    }
                }

                for (let i = 0; i < proc.parameters.length; i++) {
                    const param = proc.parameters[i];
                    const paramNameLower = param.name.toLowerCase();

                    if (param.isParamArray) {
                        const remainingArgs = positionalArgs.slice(i);
                        (remainingArgs as any).vbaBase = 0;
                        localEnv.set(param.name, remainingArgs);
                        break;
                    }

                    let argVal: any;
                    if (namedArgs.has(paramNameLower)) {
                        argVal = namedArgs.get(paramNameLower);
                    } else if (i < positionalArgs.length) {
                        argVal = positionalArgs[i];
                    } else if (param.defaultValue) {
                        argVal = this.evaluateExpression(param.defaultValue);
                    } else {
                        argVal = param.isOptional ? vbaMissing : 0;
                    }
                    localEnv.setLocally(param.name, argVal);

                    // ByRef handling
                    if (i < expr.args.length) {
                        const argExpr = expr.args[i];
                        if (!param.isByVal) {
                            byRefArgs.push({
                                paramName: param.name,
                                originalExpr: argExpr
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
                const previousErrorHandlingMode = this.errorHandlingMode;
                const previousIsInErrorHandler = this.isInErrorHandler;
                const previousLastErrorIndex = this.lastErrorIndex;

                const previousProcBody = this.currentProcBody;
                const previousExecutingModule = this.executingModuleName;
                this.env = localEnv;
                this.errorHandlerLabel = null;
                this.errorHandlingMode = 'None';
                this.isInErrorHandler = false;
                this.lastErrorIndex = -1;
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
                } finally {
                    this.env = previousEnv; // Restore scope
                    
                    // Synchronize ByRef arguments back to caller scope (even if an error occurred)
                    for (const ref of byRefArgs) {
                        const updatedVal = localEnv.get(ref.paramName);
                        try {
                            this.evaluateAssignmentToVariable(ref.originalExpr, updatedVal);
                        } catch {
                            // If it's an r-value (like a function call, literal, or expression), VBA silently discards the ByRef update
                        }
                    }

                    this.errorHandlerLabel = previousErrorHandler;
                    this.errorHandlingMode = previousErrorHandlingMode;
                    this.isInErrorHandler = previousIsInErrorHandler;
                    this.lastErrorIndex = previousLastErrorIndex;

                    this.currentProcBody = previousProcBody;
                    this.executingModuleName = previousExecutingModule;
                }

                if (proc.isFunction) {
                    return localEnv.get(proc.name.name);
                }
                return undefined;
            } else {
                // Might be an array access, built-in function, or variable reference
                const variable = this.env.get(name);
                if (expr.args.length === 0 && typeof variable !== 'function' && !Array.isArray(variable) && !(variable && variable.__isVbaDict__)) {
                    return variable;
                }
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
            case '&': return this.toDisplayString(leftVal) + this.toDisplayString(rightVal);
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
                return (leftVal instanceof VbaBoolean && rightVal instanceof VbaBoolean) ? (andRes === -1 ? vbaTrue : vbaFalse) : andRes;
            case 'or':
                const orRes = leftVal | rightVal;
                return (leftVal instanceof VbaBoolean && rightVal instanceof VbaBoolean) ? (orRes === -1 ? vbaTrue : vbaFalse) : orRes;
            case 'xor':
                const xorRes = leftVal ^ rightVal;
                return (leftVal instanceof VbaBoolean && rightVal instanceof VbaBoolean) ? (xorRes === -1 ? vbaTrue : vbaFalse) : xorRes;
            case 'eqv':
                const eqvRes = ~(leftVal ^ rightVal);
                return (leftVal instanceof VbaBoolean && rightVal instanceof VbaBoolean) ? (eqvRes === -1 ? vbaTrue : vbaFalse) : eqvRes;
            case 'imp':
                const impRes = (~leftVal) | rightVal;
                return (leftVal instanceof VbaBoolean && rightVal instanceof VbaBoolean) ? (impRes === -1 ? vbaTrue : vbaFalse) : impRes;
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
        // Special characters in VBA Like: *, ?, #, [
        // Regex special characters to escape: \, ^, $, ., |, (, ), [, ], {, }, +, * , ?
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
                if (i < patternStr.length && patternStr[i] === '!') {
                    regexStr += '^';
                    i++;
                }
                while (i < patternStr.length && patternStr[i] !== ']') {
                    const charInList = patternStr[i];
                    // Escape regex special chars inside [] if they are not part of range
                    if ('\\^$[]{}|()+.'.includes(charInList) && charInList !== '-') {
                        regexStr += '\\' + charInList;
                    } else {
                        regexStr += charInList;
                    }
                    i++;
                }
                if (i < patternStr.length) {
                    regexStr += ']';
                }
            } else {
                // Escape regex special characters
                if ('\\^$[]{}|()+.'.includes(char)) {
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
        const pLower = pattern.toLowerCase();
        if (pLower === 'general date') return d.toLocaleString();
        if (pLower === 'long date') return d.toLocaleDateString(undefined, { dateStyle: 'full' });
        if (pLower === 'medium date') return d.toLocaleDateString(undefined, { dateStyle: 'medium' });
        if (pLower === 'short date') return d.toLocaleDateString(undefined, { dateStyle: 'short' });
        if (pLower === 'long time') return d.toLocaleTimeString(undefined, { timeStyle: 'medium' });
        if (pLower === 'medium time') return d.toLocaleTimeString(undefined, { timeStyle: 'short' });
        if (pLower === 'short time') {
            return d.getHours().toString().padStart(2, '0') + ':' + d.getMinutes().toString().padStart(2, '0');
        }

        const months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
        const monthsShort = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
        const days = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
        const daysShort = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];

        return pattern.replace(/yyyy|yy|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|nn|n|ss|s|AM\/PM|am\/pm/g, (match) => {
            const mLower = match.toLowerCase();
            switch (mLower) {
                case 'yyyy': return String(d.getFullYear());
                case 'yy': return String(d.getFullYear()).slice(-2);
                case 'mmmm': return months[d.getMonth()];
                case 'mmm': return monthsShort[d.getMonth()];
                case 'mm': return String(d.getMonth() + 1).padStart(2, '0');
                case 'm': return String(d.getMonth() + 1);
                case 'dddd': return days[d.getDay()];
                case 'ddd': return daysShort[d.getDay()];
                case 'dd': return String(d.getDate()).padStart(2, '0');
                case 'd': return String(d.getDate());
                case 'hh': return String(d.getHours()).padStart(2, '0');
                case 'h': return String(d.getHours());
                case 'nn': return String(d.getMinutes()).padStart(2, '0');
                case 'n': return String(d.getMinutes());
                case 'ss': return String(d.getSeconds()).padStart(2, '0');
                case 's': return String(d.getSeconds());
                case 'am/pm': return d.getHours() >= 12 ? (match === 'AM/PM' ? 'PM' : 'pm') : (match === 'AM/PM' ? 'AM' : 'am');
                default: return match;
            }
        });
    }

    private formatNumber(n: number, pattern: string): string {
        const pLower = pattern.toLowerCase();
        // Handle named formats
        switch (pLower) {
            case 'general number': return String(n);
            case 'currency': return n.toLocaleString(undefined, { style: 'currency', currency: 'USD' });
            case 'fixed': return n.toFixed(2);
            case 'standard': return n.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
            case 'percent': return (n * 100).toFixed(2) + '%';
            case 'scientific': return n.toExponential(2);
            case 'true/false': return n !== 0 ? 'True' : 'False';
            case 'yes/no': return n !== 0 ? 'Yes' : 'No';
            case 'on/off': return n !== 0 ? 'On' : 'Off';
        }

        // Custom patterns: #, 0, ., ,, %
        let fmt = pattern;
        let isPercent = false;
        if (fmt.includes('%')) {
            isPercent = true;
            n *= 100;
            fmt = fmt.replace('%', '');
        }

        const hasThousands = fmt.includes(',');
        const parts = fmt.split('.');
        const integerPart = parts[0];
        const decimalPart = parts.length > 1 ? parts[1] : '';

        // Determine decimal places
        const minDecimals = (decimalPart.match(/0/g) || []).length;
        const maxDecimals = decimalPart.length;

        let result = n.toLocaleString(undefined, {
            minimumFractionDigits: minDecimals,
            maximumFractionDigits: maxDecimals,
            useGrouping: hasThousands
        });

        if (isPercent) result += '%';
        return result;
    }
}
