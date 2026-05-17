import { useState, useRef, useCallback, useMemo } from 'react'
import { Play, Eraser, TerminalSquare, FlaskConical, Github } from 'lucide-react'
import { Lexer, TokenType } from './engine/lexer'
import { Parser } from './engine/parser'
import { Evaluator } from './engine/evaluator'

// ======= VBA Syntax Highlighting =======
function tokenToClass(type: TokenType): string {
  switch (type) {
    case TokenType.KeywordFor: case TokenType.KeywordTo: case TokenType.KeywordNext:
    case TokenType.KeywordIf: case TokenType.KeywordThen: case TokenType.KeywordElseIf:
    case TokenType.KeywordElse: case TokenType.KeywordEnd: case TokenType.KeywordDo:
    case TokenType.KeywordWhile: case TokenType.KeywordLoop: case TokenType.KeywordSub:
    case TokenType.KeywordFunction: case TokenType.KeywordDim: case TokenType.KeywordAs:
    case TokenType.KeywordNew: case TokenType.KeywordConst: case TokenType.KeywordSet:
    case TokenType.KeywordCall: case TokenType.KeywordOn: case TokenType.KeywordError:
    case TokenType.KeywordGoTo: case TokenType.KeywordErase: case TokenType.KeywordReDim:
    case TokenType.KeywordStep: case TokenType.KeywordExit: case TokenType.KeywordByRef:
    case TokenType.KeywordByVal: case TokenType.KeywordOption: case TokenType.KeywordExplicit:
    case TokenType.KeywordType: case TokenType.KeywordEmpty:
    case TokenType.KeywordAnd: case TokenType.KeywordOr: case TokenType.KeywordNot:
    case TokenType.KeywordMod:
      return 'syn-keyword';
    case TokenType.KeywordCollection:
      return 'syn-type';
    case TokenType.String:
      return 'syn-string';
    case TokenType.Number:
      return 'syn-number';
    default:
      return '';
  }
}

function highlightVBA(code: string): string {
  try {
    const lexer = new Lexer(code);
    const tokens = lexer.tokenize();
    let result = '';
    let srcPos = 0;
    const src = code;
    for (const token of tokens) {
      if (token.type === TokenType.EOF) break;
      const val = token.type === TokenType.String ? `"${token.value}"` : token.value;
      const idx = src.indexOf(val, srcPos);
      if (idx === -1) continue;
      if (idx > srcPos) result += escapeHtml(src.substring(srcPos, idx));
      const cls = tokenToClass(token.type);
      const escaped = escapeHtml(val);
      result += cls ? `<span class="${cls}">${escaped}</span>` : escaped;
      srcPos = idx + val.length;
    }
    if (srcPos < src.length) result += escapeHtml(src.substring(srcPos));
    return result;
  } catch { return escapeHtml(code); }
}

// Apply comment highlighting on top of lexer output
function applyVBAComments(code: string, highlighted: string): string {
  const lines = code.split('\n');
  const hLines = highlighted.split('\n');
  const result: string[] = [];
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const hLine = i < hLines.length ? hLines[i] : '';
    const cIdx = findVBACommentStart(line);
    if (cIdx !== -1) {
      const pos = mapToHighlightedPos(hLine, cIdx);
      result.push(hLine.substring(0, pos) + `<span class="syn-comment">${escapeHtml(line.substring(cIdx))}</span>`);
    } else {
      result.push(hLine);
    }
  }
  return result.join('\n');
}

function findVBACommentStart(line: string): number {
  let inStr = false;
  for (let i = 0; i < line.length; i++) {
    if (line[i] === '"') inStr = !inStr;
    if (!inStr && line[i] === "'") return i;
    if (!inStr && line.substring(i, i + 3).toLowerCase() === 'rem' &&
      (i === 0 || /\s/.test(line[i - 1])) &&
      (i + 3 >= line.length || /\s/.test(line[i + 3]))) return i;
  }
  return -1;
}

function highlightVBAFull(code: string): string {
  return applyVBAComments(code, highlightVBA(code));
}

// ======= TypeScript Syntax Highlighting =======
const TS_KEYWORDS = new Set([
  'import', 'export', 'from', 'const', 'let', 'var', 'function', 'return',
  'if', 'else', 'for', 'while', 'do', 'switch', 'case', 'break', 'continue',
  'new', 'this', 'class', 'extends', 'interface', 'type', 'enum',
  'async', 'await', 'try', 'catch', 'finally', 'throw',
  'typeof', 'instanceof', 'in', 'of', 'as', 'null', 'undefined', 'true', 'false',
  'void', 'any', 'string', 'number', 'boolean', 'Map', 'Array',
]);

const TS_TYPES = new Set([
  'string', 'number', 'boolean', 'any', 'void', 'Map', 'Array', 'Object',
]);

function highlightTS(code: string): string {
  let result = '';
  let i = 0;
  while (i < code.length) {
    // Line comment
    if (code[i] === '/' && code[i + 1] === '/') {
      const end = code.indexOf('\n', i);
      const slice = end === -1 ? code.substring(i) : code.substring(i, end);
      result += `<span class="syn-comment">${escapeHtml(slice)}</span>`;
      i += slice.length;
      continue;
    }
    // Block comment
    if (code[i] === '/' && code[i + 1] === '*') {
      const end = code.indexOf('*/', i + 2);
      const slice = end === -1 ? code.substring(i) : code.substring(i, end + 2);
      result += `<span class="syn-comment">${escapeHtml(slice)}</span>`;
      i += slice.length;
      continue;
    }
    // String (single/double/backtick)
    if (code[i] === '"' || code[i] === "'" || code[i] === '`') {
      const q = code[i];
      let j = i + 1;
      while (j < code.length && code[j] !== q) {
        if (code[j] === '\\') j++; // skip escaped
        j++;
      }
      if (j < code.length) j++; // consume closing quote
      result += `<span class="syn-string">${escapeHtml(code.substring(i, j))}</span>`;
      i = j;
      continue;
    }
    // Number
    if (/[0-9]/.test(code[i]) && (i === 0 || !/[a-zA-Z_$]/.test(code[i - 1]))) {
      let j = i;
      while (j < code.length && /[0-9.eExX_a-fA-F]/.test(code[j])) j++;
      result += `<span class="syn-number">${escapeHtml(code.substring(i, j))}</span>`;
      i = j;
      continue;
    }
    // Identifier / keyword
    if (/[a-zA-Z_$]/.test(code[i])) {
      let j = i;
      while (j < code.length && /[a-zA-Z0-9_$]/.test(code[j])) j++;
      const word = code.substring(i, j);
      if (TS_TYPES.has(word) && /[<:\s]/.test(code[j] || ' ')) {
        result += `<span class="syn-type">${escapeHtml(word)}</span>`;
      } else if (TS_KEYWORDS.has(word)) {
        result += `<span class="syn-keyword">${escapeHtml(word)}</span>`;
      } else {
        result += escapeHtml(word);
      }
      i = j;
      continue;
    }
    result += escapeHtml(code[i]);
    i++;
  }
  return result;
}

// ======= Shared Utilities =======
function escapeHtml(str: string): string {
  return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function mapToHighlightedPos(hLine: string, plainPos: number): number {
  let plainCount = 0;
  let inTag = false;
  for (let i = 0; i < hLine.length; i++) {
    if (hLine[i] === '<') { inTag = true; continue; }
    if (hLine[i] === '>') { inTag = false; continue; }
    if (inTag) continue;
    if (hLine[i] === '&') {
      const semi = hLine.indexOf(';', i);
      if (semi !== -1 && semi - i < 8) {
        if (plainCount === plainPos) return i;
        plainCount++;
        i = semi;
        continue;
      }
    }
    if (plainCount === plainPos) return i;
    plainCount++;
  }
  return hLine.length;
}

// ======= Highlighted Editor Component =======
interface HighlightedEditorProps {
  value: string;
  onChange: (val: string) => void;
  highlighter: (code: string) => string;
  placeholder?: string;
}

function HighlightedEditor({ value, onChange, highlighter, placeholder }: HighlightedEditorProps) {
  const textareaRef = useRef<HTMLTextAreaElement>(null);
  const highlightRef = useRef<HTMLPreElement>(null);
  const lineNumRef = useRef<HTMLDivElement>(null);

  const handleScroll = useCallback(() => {
    if (textareaRef.current && highlightRef.current && lineNumRef.current) {
      highlightRef.current.scrollTop = textareaRef.current.scrollTop;
      highlightRef.current.scrollLeft = textareaRef.current.scrollLeft;
      lineNumRef.current.scrollTop = textareaRef.current.scrollTop;
    }
  }, []);

  const lines = useMemo(() => {
    const count = value.split('\n').length;
    return Array.from({ length: Math.max(1, count) }, (_, i) => i + 1);
  }, [value]);

  const highlighted = useMemo(() => highlighter(value), [value, highlighter]);

  return (
    <div className="flex-1 relative flex overflow-hidden">
      <div
        className="w-12 bg-neutral-900 text-neutral-500 text-right pr-3 py-4 font-mono text-sm select-none overflow-hidden leading-[1.65] shrink-0"
        ref={lineNumRef}
      >
        {lines.map((n) => <div key={n}>{n}</div>)}
      </div>
      <div className="flex-1 relative min-w-0">
        <pre
          ref={highlightRef}
          className="absolute inset-0 m-0 p-4 font-mono text-sm leading-[1.65] whitespace-pre overflow-hidden pointer-events-none text-neutral-200"
          aria-hidden="true"
          dangerouslySetInnerHTML={{ __html: highlighted + '\n' }}
        />
        <textarea
          ref={textareaRef}
          value={value}
          onChange={(e) => onChange(e.target.value)}
          onScroll={handleScroll}
          spellCheck={false}
          wrap="off"
          className="absolute inset-0 w-full h-full bg-transparent text-transparent caret-neutral-200 font-mono text-sm p-4 focus:outline-none resize-none leading-[1.65] whitespace-pre overflow-auto"
          placeholder={placeholder}
        />
      </div>
    </div>
  );
}

// ======= Browser Test Runner =======
function createBrowserTestRunner(vbaSource: string, outputFn: (msg: string) => void) {
  // Compile VBA source once
  const lexer = new Lexer(vbaSource);
  const tokens = lexer.tokenize();
  const parser = new Parser(tokens);
  const ast = parser.parse();

  // Only evaluate declarations (Functions, Subs, Types, Consts, Variables) 
  // to avoid executing top-level subroutine calls (e.g. `MainLoop`) during initialization
  const declarationAst = {
    ...ast,
    body: ast.body.filter(stmt =>
      stmt.type === 'ProcedureDeclaration' ||
      stmt.type === 'TypeDeclaration' ||
      stmt.type === 'VariableDeclaration' ||
      stmt.type === 'ConstDeclaration'
    )
  };

  const evaluator = new Evaluator((out) => { outputFn(out) });
  evaluator.evaluate(declarationAst);

  // vbaTest object: matches VBARunner class interface
  const vbaTest = {
    run(procedureName: string, args: any[]): any {
      return evaluator.callProcedure(procedureName, args);
    },
    eval(exprString: string): any {
      return evaluator.evalExpression(exprString);
    }
  };

  // assert.strictEqual
  const assert = {
    strictEqual: (actual: any, expected: any, message?: string) => {
      const label = message || `${actual} === ${expected}`;
      if (actual !== expected) {
        const msg = `[FAIL] ${label} - Expected ${expected} but got ${actual}`;
        outputFn(msg);
        throw new Error(msg);
      }
      outputFn(`[PASS] ${label}`);
    }
  };

  return { vbaTest, assert };
}

// ======= Main App =======
function App() {
  const defaultSnippet = `Function AddNumbers(a, b)
  AddNumbers = a + b
End Function

Sub MainLoop()
  ' This loop prints 11, 12, 13
  for i = 1 to 3
    Rem A VBA comment
    debug.print AddNumbers(i, 10)
  next i
End Sub

MainLoop`

  const defaultTestCode = `// eval example: evaluate VBA expression directly
assert.strictEqual(vbaTest.eval("AddNumbers(3, 4)"), 7);
assert.strictEqual(vbaTest.eval("AddNumbers(0, 0)"), 0);
assert.strictEqual(vbaTest.eval("AddNumbers(-5, 10)"), 5);

// run example: call VBA procedure with JS arguments
const r = vbaTest.run("AddNumbers", [100, 200]);
assert.strictEqual(r, 300);`;

  const [code, setCode] = useState(defaultSnippet)
  const [testCode, setTestCode] = useState(defaultTestCode)
  const [runOutput, setRunOutput] = useState<string[]>([])
  const [testOutput, setTestOutput] = useState<string[]>([])
  const [consoleTab, setConsoleTab] = useState<'run' | 'test'>('run')

  const handleRun = () => {
    setConsoleTab('run')
    setRunOutput(['> Compiling and running...'])
    try {
      const lexer = new Lexer(code)
      const tokens = lexer.tokenize()
      const parser = new Parser(tokens)
      const ast = parser.parse()
      const newOutputs: string[] = []
      const evaluator = new Evaluator((out) => { newOutputs.push(out) })
      evaluator.evaluate(ast)
      setRunOutput(prev => [...prev, ...newOutputs, '> Execution finished.'])
    } catch (err: any) {
      const msg = err.number ? `Error: ${err.number} ${err.message}` : `Error: ${err.message}`;
      setRunOutput(prev => [...prev, msg])
    }
  }

  const handleRunTests = () => {
    setConsoleTab('test')
    const outputs: string[] = ['> Compiling VBA source...'];
    let passCount = 0, failCount = 0;

    try {
      // Create test runner with VBA source
      const { vbaTest, assert } = createBrowserTestRunner(code, (msg) => {
        outputs.push(msg);
        if (msg.includes('[PASS]')) passCount++;
        if (msg.includes('[FAIL]')) failCount++;
      });

      outputs.push('> Running TypeScript tests...');
      outputs.push('');

      // Prepare console.log to capture output
      const testConsole = {
        log: (...args: any[]) => {
          outputs.push(args.map(a => typeof a === 'object' ? JSON.stringify(a) : String(a)).join(' '));
        },
        error: (...args: any[]) => {
          outputs.push('ERROR: ' + args.map(a => typeof a === 'object' ? JSON.stringify(a) : String(a)).join(' '));
        }
      };

      // Strip TypeScript type annotations for browser execution
      const jsCode = stripTypeAnnotations(testCode);

      // Execute test code with vbaTest pre-created
      const testFn = new Function('vbaTest', 'assert', 'console', 'Map', jsCode);
      testFn(vbaTest, assert, testConsole, Map);

      const summary = failCount === 0
        ? `\n✅ All ${passCount} tests passed!`
        : `\n❌ ${failCount} failed, ${passCount} passed`;
      outputs.push(summary);
    } catch (err: any) {
      const msg = err.number ? `Error: ${err.number} ${err.message}` : `Error: ${err.message}`;
      outputs.push(msg);
      if (failCount > 0 || passCount > 0) {
        outputs.push(`\n❌ ${failCount} failed, ${passCount} passed (aborted)`);
      }
    }

    setTestOutput(outputs);
  }

  const handleClear = () => {
    if (consoleTab === 'run') setRunOutput([]);
    else setTestOutput([]);
  }

  const currentOutput = consoleTab === 'run' ? runOutput : testOutput;

  const renderLine = (line: string, i: number) => (
    <div
      key={i}
      className={`whitespace-pre-wrap leading-relaxed ${line.includes('[PASS]') ? 'text-emerald-400' :
        line.includes('[FAIL]') ? 'text-red-400' :
          line.startsWith('✅') ? 'text-emerald-400 font-semibold mt-2' :
            line.startsWith('❌') ? 'text-red-400 font-semibold mt-2' :
              line.startsWith('Error:') ? 'text-red-400' :
                ''
        }`}
    >{line}</div>
  );

  return (
    <div className="flex flex-col h-screen bg-neutral-950 text-neutral-100 font-sans">
      {/* Header */}
      <header className="flex items-center px-6 py-3 bg-neutral-900 border-b border-neutral-800 shadow-sm shrink-0">
        <div className="flex items-center gap-3">
          <TerminalSquare className="w-6 h-6 text-blue-500" />
          <h1 className="text-xl font-semibold tracking-tight">VBA Web Runner</h1>
          <a href="https://github.com/jca02266/vba-runner" target="_blank" rel="noopener noreferrer" className="flex items-center gap-1.5 text-neutral-500 hover:text-white transition-colors text-sm">
            <Github className="w-4 h-4" />
            GitHub
          </a>
        </div>
      </header>

      {/* Editors (top row) */}
      <div className="flex flex-1 min-h-0 overflow-hidden">
        {/* Source Editor - VBA (left) */}
        <section className="flex-1 flex flex-col border-r border-neutral-800 bg-neutral-950 min-w-0">
          <div className="px-4 py-1.5 bg-neutral-900 border-b border-neutral-800 text-xs font-mono text-neutral-400 flex items-center gap-3 shrink-0">
            <button
              onClick={handleRun}
              className="flex items-center gap-1 px-2.5 py-1 text-xs font-semibold text-white bg-blue-600 rounded hover:bg-blue-500 transition-colors cursor-pointer"
            >
              <Play className="w-3 h-3" />
              Run
            </button>
            <button
              onClick={() => setRunOutput([])}
              className="flex items-center gap-1 px-2 py-1 text-xs text-neutral-400 hover:text-white hover:bg-neutral-700 rounded transition-colors cursor-pointer"
            >
              <Eraser className="w-3 h-3" />
              Clear
            </button>
            <span className="w-2 h-2 rounded-full bg-green-500 ml-auto"></span>
            main.vba
          </div>
          <HighlightedEditor value={code} onChange={setCode} highlighter={highlightVBAFull} placeholder="Type your VBA code here..." />
        </section>

        {/* Test Editor - TypeScript (right) */}
        <section className="flex-1 flex flex-col bg-neutral-950 min-w-0">
          <div className="px-4 py-1.5 bg-neutral-900 border-b border-neutral-800 text-xs font-mono text-neutral-400 flex items-center gap-3 shrink-0">
            <button
              onClick={handleRunTests}
              className="flex items-center gap-1 px-2.5 py-1 text-xs font-semibold text-white bg-emerald-600 rounded hover:bg-emerald-500 transition-colors cursor-pointer"
            >
              <FlaskConical className="w-3 h-3" />
              Test
            </button>
            <span className="w-2 h-2 rounded-full bg-emerald-500 ml-auto"></span>
            test.ts
          </div>
          <HighlightedEditor value={testCode} onChange={setTestCode} highlighter={highlightTS} placeholder="Write your TypeScript test code here..." />
        </section>
      </div>

      {/* Console (bottom) */}
      <section className="h-56 flex flex-col bg-black border-t border-neutral-800 shrink-0">
        <div className="flex items-center gap-2 bg-neutral-900 border-b border-neutral-800 shrink-0 px-3">
          {/* Action Buttons */}
          <button
            onClick={handleClear}
            className="flex items-center gap-1 px-2 py-1 text-xs text-neutral-400 hover:text-white hover:bg-neutral-700 rounded transition-colors cursor-pointer"
          >
            <Eraser className="w-3 h-3" />
            Clear
          </button>
          <button
            onClick={handleRun}
            className="flex items-center gap-1 px-2.5 py-1 text-xs font-semibold text-white bg-blue-600 rounded hover:bg-blue-500 transition-colors cursor-pointer"
          >
            <Play className="w-3 h-3" />
            Run
          </button>
          <button
            onClick={handleRunTests}
            className="flex items-center gap-1 px-2.5 py-1 text-xs font-semibold text-white bg-emerald-600 rounded hover:bg-emerald-500 transition-colors cursor-pointer"
          >
            <FlaskConical className="w-3 h-3" />
            Test
          </button>

          {/* Console Tabs */}
          <button
            onClick={() => setConsoleTab('run')}
            className={`ml-3 px-4 py-2 text-xs font-mono flex items-center gap-2 transition-colors cursor-pointer border-b-2 ${consoleTab === 'run'
              ? 'text-neutral-200 border-blue-500 bg-black'
              : 'text-neutral-500 border-transparent hover:text-neutral-300 hover:bg-neutral-800'
              }`}
          >
            <span className={`w-2 h-2 rounded-full ${consoleTab === 'run' ? 'bg-blue-500' : 'bg-neutral-600'}`}></span>
            Console
          </button>
          <button
            onClick={() => setConsoleTab('test')}
            className={`px-4 py-2 text-xs font-mono flex items-center gap-2 transition-colors cursor-pointer border-b-2 ${consoleTab === 'test'
              ? 'text-neutral-200 border-emerald-500 bg-black'
              : 'text-neutral-500 border-transparent hover:text-neutral-300 hover:bg-neutral-800'
              }`}
          >
            <span className={`w-2 h-2 rounded-full ${consoleTab === 'test' ? 'bg-emerald-500' : 'bg-neutral-600'}`}></span>
            Test Results
          </button>
        </div>
        <div className="flex-1 p-4 overflow-y-auto font-mono text-sm text-neutral-300">
          {currentOutput.length === 0 ? (
            <span className="text-neutral-600 italic">Ready...</span>
          ) : (
            currentOutput.map(renderLine)
          )}
        </div>
      </section >
    </div >
  )
}

// ======= TypeScript to JS Stripping =======
// Simple transformer: remove type annotations, generics, and TS-specific syntax
function stripTypeAnnotations(tsCode: string): string {
  let result = tsCode;

  // Remove import statements
  result = result.replace(/^import\s+.*$/gm, '');

  // Remove "export" keyword
  result = result.replace(/\bexport\s+/g, '');

  // Remove ": type" annotations from variable declarations and function params
  // e.g., "const x: string" -> "const x", "(a: number)" -> "(a)"
  result = result.replace(/:\s*(?:any|string|number|boolean|void|Map<[^>]*>|Array<[^>]*>|[A-Z][a-zA-Z0-9]*(?:\[\])?)\s*(?=[,=)\];}]|$)/gm, '');

  // Remove generic type parameters on function calls: e.g., Map<string, number>() -> Map()
  result = result.replace(/<(?:string|number|boolean|any|void|[A-Z][a-zA-Z0-9]*)(?:\s*,\s*(?:string|number|boolean|any|void|[A-Z][a-zA-Z0-9]*))*>/g, '');

  // Remove "as TypeName" casts
  result = result.replace(/\bas\s+[A-Z][a-zA-Z0-9]*/g, '');

  // Remove async/await for simplicity (our tests are sync)
  result = result.replace(/\basync\s+/g, '');
  result = result.replace(/\bawait\s+/g, '');

  return result;
}

export default App
