# vba-runner

A TypeScript-based VBA execution engine that runs and tests VBA code without Excel.
Load `.bas` files from TypeScript and call VBA procedures directly, or use the bundled CLI tools for static analysis, formatting, and syntax checking.

**[日本語](./README.ja.md)** | **[CHANGELOG](./CHANGELOG.md)** | **[VBA Runner Project](../../README.md)**

## Installation

```bash
npm install vba-runner
```

## Usage

> [!TIP]
> Try it in the browser before writing any code: [Web UI Demo](https://vba-web-runner.netlify.app/) lets you run VBA snippets and check `Debug.Print` output instantly.

### 1. `eval`: Evaluate a VBA expression or code fragment inline

Pass VBA syntax as a string and get the result back. The quickest way to try the engine.

```typescript
import { VBARunner } from 'vba-runner';
const vbaRunner = new VBARunner(); // create an empty environment

// Evaluate a VBA expression directly
const sum = vbaRunner.eval("1 + 2 + 3"); // => 6

// Multi-line code with variable declarations works too
vbaRunner.eval("Dim x : x = 10 : Debug.Print x * 2");
```

### 2. `run`: Call a procedure defined in a loaded VBA file

Load an existing `.bas` file and call its procedures with arguments. Ideal for unit-testing complex business logic.

```typescript
import { VBARunner, assert } from 'vba-runner';

// 1. Load the VBA file under test
const vbaRunner = new VBARunner('src/vba/Sample.bas');

// 2. Call a procedure by name (pass arguments as a JavaScript array)
const result1 = vbaRunner.run('Add', [1, 2]);
const result2 = vbaRunner.run('Multiply', [result1, 2]);

// 3. Assert the results
assert.strictEqual(result1, 3);
assert.strictEqual(result2, 6);
```

### 3. Load files with glob patterns

Pass a glob pattern string (or an array of patterns) to select files. Supports `*`, `**`, and `{a,b}` brace expansion — same as bash glob.

```typescript
import { VBARunner, assert } from 'vba-runner';

// Single glob pattern
const vbaRunner = new VBARunner('src/vba/*.{bas,cls}');

// Recursive glob
const vbaRunner2 = new VBARunner('src/**/*.bas');

// Mix of patterns and literal paths in an array
const vbaRunner3 = new VBARunner([
    'src/vba/core/*.bas',
    'src/vba/models/Account.cls',
]);
```

Glob patterns are resolved relative to `process.cwd()`. No mock directory scanning is performed (mocks are a directory-load feature).

### 4. Load specific files by array

Pass an array of literal file paths to load exactly the files you want:

```typescript
import { VBARunner, assert } from 'vba-runner';

const vbaRunner = new VBARunner([
    'src/vba/Calc.cls',
    'src/vba/Utils.bas',
]);

const result = vbaRunner.run('DoubleIt', [21]);
assert.strictEqual(result, 42);
```

### 5. Load an entire directory

Point `VBARunner` at a directory to load all `.bas` and `.cls` files at once. Use this for larger VBA projects with multiple modules, including **Class modules** (`.cls`).

```typescript
import { VBARunner, assert } from 'vba-runner';

// 1. Load all VBA files in the directory (.bas and .cls)
const vbaRunner = new VBARunner('src/vba');

// 2. Call any procedure defined across those files
const result = vbaRunner.run('CalcTotal', [100, 200, 300]);
assert.strictEqual(result, 600);
```

Class modules work out of the box — VBA standard `.cls` headers are stripped automatically. `Class_Initialize` / `Class_Terminate`, `Property Get/Let/Set`, and `Private`/`Public` members are all supported.

```typescript
// src/vba/BankAccount.cls
// VERSION 1.0 CLASS
// ...
// Attribute VB_Name = "BankAccount"
// Private m_balance As Double
// ...

const vbaRunner = new VBARunner('src/vba');
vbaRunner.eval('Dim acct As New BankAccount');
vbaRunner.eval('acct.Deposit 1000');
assert.strictEqual(vbaRunner.eval('acct.Balance'), 1000);
```

> **State persists across `eval()` and `run()` calls** within the same `VBARunner` instance. Variables and objects declared with `eval()` remain accessible in subsequent `eval()` or `run()` calls — useful for setting up test fixtures step by step.

### 4. Comparing Boolean values

`run()` and `eval()` automatically convert VBA `Boolean` results to plain JS `true`/`false`,
so standard assertions work directly:

```typescript
import { VBARunner, assert } from 'vba-runner';

const vbaRunner = new VBARunner('src/vba/Sample.bas');
assert.strictEqual(vbaRunner.run('IsPositive', [5]), true);  // OK
```

### 5. Mocking Excel-dependent objects (Application / ActiveSheet, etc.)

VBA code that uses `ActiveSheet.Range(...)` or `Cells(...)` can be tested with the
built-in mock by enabling the `excelStub` option.

```typescript
import { VBARunner, assert } from 'vba-runner';

const vbaRunner = new VBARunner('src/vba/Sheet1Logic.bas', { excelStub: true });

// Seed a cell value before running
vbaRunner.excelStub.ActiveSheet.setCellValue('A1', 100);

vbaRunner.run('DoubleA1ToB1', []);

// Check the resulting cell value
assert.strictEqual(vbaRunner.excelStub.ActiveSheet.getCellValue('B1'), 200);
```

> **Note 1: `Cells(row, col)` shares the same storage as `Range(address)`.**
> Numeric addressing like `Cells(13, 1)` is converted internally to the equivalent of
> `Range("A13")`, so it reads/writes the same cell as `setCellValue`/`getCellValue`. Seed test
> data with string addresses (`setCellValue('A1', ...)`) and it works whether the VBA code
> under test uses `Cells` or `Range`.

The built-in mock only persists `Value` reads/writes. Formatting properties like
`Interior.Color` are no-ops that don't retain state, and `Application.OnKey` /
`Application.OnTime` aren't implemented (calling them throws). To test code that
depends on these, **extend the mock rather than replacing `Application` wholesale**
— a full replacement discards `ActiveSheet`/`Sheets`/`Range` for any code that also
needs those.

> **Note 2: code that passes a sub-second offset to `Application.OnTime` (e.g.
> `Now + 0.4/86400`) can fall into an infinite loop on a real Excel session.** VBA's `Now()`
> truncates to whole seconds, so the computed target time can already be in the past by the
> time `Now()` is evaluated — and `Application.OnTime` fires immediately when given a past
> time. If the timer handler reschedules itself the same way, this repeats endlessly and pegs
> the UI thread (CPU usage stays low, so it looks like a hang rather than a busy loop — easy to
> miss). Always use an offset of at least one second.

The cleanest way is to subclass `MockApplication` (also exported by this package)
and pass an instance to `excelStub` instead of `true`:

```typescript
import { VBARunner, MockApplication, assert } from 'vba-runner';

class AppWithOnKey extends MockApplication {
  onKeyLog: string[] = [];
  OnKey(key: string, procedureName?: string) {
    this.onKeyLog.push(`${key}=${procedureName ?? ''}`);
  }
}

const vbaRunner = new VBARunner('src/vba/KeyHandler.bas', { excelStub: new AppWithOnKey() });

vbaRunner.run('SetupKeyHandlers', []);
assert.ok(vbaRunner.excelStub.onKeyLog.length > 0);
```

If you don't need a reusable class, you can instead add the missing method directly
onto the `excelStub` instance after construction — it's the exact object already
wired up as `Application`/`ActiveSheet`, so extending it in place works too:

```typescript
const vbaRunner = new VBARunner('src/vba/KeyHandler.bas', { excelStub: true });

let registered: [string, string] | null = null;
(vbaRunner.excelStub as any).OnKey = (key: string, procedureName: string) => {
  registered = [key, procedureName];
};

vbaRunner.run('SetupKeyHandlers', []);
assert.ok(registered !== null);
```

Only reach for `vbaRunner.evaluator.setBuiltinOverride(name, value)` when you want to
replace a builtin (e.g. `Application`) entirely with a different object — note that
this discards the rest of the default Excel stub (`ActiveSheet`, `Sheets`, etc.)
unless your VBA code never touches them through that same name.

To inject any other named variable or constant directly, `vbaRunner.set(name, value)` is also available.

See [`docs/MOCK_GUIDE.md`](https://github.com/jca02266/vba-runner/blob/main/docs/MOCK_GUIDE.md) for more advanced mocking patterns.

### 6. Quiet mode / routing `Debug.Print` output separately

By default, `run()` logs every call (`[PASS] ProcName(...) -> result (Nms)`) and
`Debug.Print` goes to `console.log`, so both end up interleaved on stdout. Pass
`quiet: true` to suppress the `[PASS]` logs, and `onPrint` to redirect
`Debug.Print` output independently (e.g. to `stderr`, or to a collector array):

```typescript
import { VBARunner } from 'vba-runner';

const vbaRunner = new VBARunner('src/vba/Sample.bas', {
  quiet: true,                                    // suppress [PASS] logs
  onPrint: (s) => process.stderr.write(s + '\n'),  // Debug.Print -> stderr
});

vbaRunner.run('SeedGlider', [1, 1]); // no [PASS] log printed
```

### 8. Virtual filesystem (`vbaRunner.fs`)

VBA file I/O (`Open`, `Print #`, `Line Input #`, `Close`, `Scripting.FileSystemObject`, etc.)
runs against an **in-memory virtual filesystem** — nothing touches your real disk.
Every `VBARunner` instance gets its own independent VFS.

Windows-style paths used inside VBA code are automatically mapped into the VFS:

| VBA path | VFS path (default `sandboxRoot: '/sandbox'`) |
|---|---|
| `C:\data\input.txt` | `/sandbox/c/data/input.txt` |
| `D:\report.csv` | `/sandbox/d/report.csv` |
| `relative\path.txt` | `/sandbox/relative/path.txt` |

Use `vbaRunner.fs` to pre-populate input files or read back output files from TypeScript:

```typescript
import { VBARunner } from 'vba-runner';

const vbaRunner = new VBARunner('src/vba/Report.bas');

// Pre-populate an input file the VBA code will read
vbaRunner.fs.mkdirSync('/sandbox/c/data', { recursive: true });
vbaRunner.fs.writeFileSync('/sandbox/c/data/input.csv', 'Alice,100\nBob,200\n');

// Run the VBA procedure that reads the file and writes a summary
vbaRunner.run('GenerateSummary', []);

// Read back the output file produced by VBA
const output = vbaRunner.fs.readFileSync('/sandbox/c/data/summary.txt', 'utf8');
console.log(output);
```

To use a different VFS root, pass `sandboxRoot` to the constructor:

```typescript
const vbaRunner = new VBARunner('src/vba/Report.bas', { sandboxRoot: '/myapp' });
// Now C:\file.txt maps to /myapp/c/file.txt
vbaRunner.fs.writeFileSync('/myapp/c/file.txt', 'data');
```

> File I/O done with `Do While Not ts.AtEndOfStream` (FSO TextStream) or `Do While Not EOF(n)` (native `Open`) both work correctly in the VFS.

### 9. `ByRef` out-parameters

VBA's default argument-passing mode is `ByRef`. When a called `Sub`/`Function`
assigns to a `ByRef` parameter, `run()` writes the final value back into the
JS array you passed in — useful for the common "status + message" pattern:

```typescript
import { VBARunner } from 'vba-runner';

const vbaRunner = new VBARunner('src/vba/Sample.bas');

const args: any[] = [/* plate */ 'ABC-123', /* outMessage */ ''];
vbaRunner.run('ParkCar', args);
console.log(args[1]); // the Sub's ByRef outMessage parameter, after the call
```

`ByVal` parameters are left untouched (the caller's original value is kept).

### 10. Conditional compilation (`#If` / `#Const`)

vba-runner supports VBA conditional compilation directives (`#If`, `#Else`, `#ElseIf`, `#End If`, `#Const`).

**Default compiler constants** mirror a modern 64-bit Windows VBA environment (Office 2010+, 64-bit):

| Constant | Default | Meaning |
|---|---|---|
| `VBA7` | `True` (-1) | VBA version 7 (Office 2010+) |
| `Win64` | `True` (-1) | 64-bit process |
| `Win32` | `True` (-1) | Windows platform (always True on Windows, even in 64-bit) |
| `Mac` | `False` (0) | Mac platform |

These defaults mean that code guarded by `#If VBA7 Then` or `#If Win64 Then` is **included** by default, which matches the most common modern Office environment.

**Overriding compiler constants** — pass `compilerConstants` to the constructor to simulate a different environment:

```typescript
// Simulate 32-bit Office (pre-2010 or 32-bit Office on 64-bit Windows)
const vbaRunner = new VBARunner('src/vba/Module.bas', {
  compilerConstants: { VBA7: 0, Win64: 0 },
});
```

```typescript
// Simulate Mac VBA
const vbaRunner = new VBARunner('src/vba/Module.bas', {
  compilerConstants: { Mac: -1, Win32: 0, Win64: 0 },
});
```

**Note:** `#Const` directives inside the VBA source file take priority over `compilerConstants` passed to the constructor, which is consistent with VBA specification.

## CLI Tools

| Command | Description |
|---|---|
| `vba-run <file.bas>` | Execute a VBA file and print `Debug.Print` output |
| `vba-analyzer <file.bas>` | Static analysis: outline, reference counts, duplicate detection |
| `vba-formatter <file.bas>` | Format VBA code (indentation, spacing) |
| `vba-parse-check <file.bas>` | Syntax check (detect parse errors) |

```bash
# Show outline of a VBA file
vba-analyzer --outline src/vba/Module1.bas

# Detect duplicate code blocks
vba-analyzer --diff src/vba/

# Syntax check
vba-parse-check src/vba/Module1.bas
```

## Documentation

- [REFERENCE.md](https://github.com/jca02266/vba-runner/blob/main/REFERENCE.md) — Type system, mock registration, Sandbox policy, VFS details
- [README.md](https://github.com/jca02266/vba-runner/blob/main/README.md) — Project overview and quick start
