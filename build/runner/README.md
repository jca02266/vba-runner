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

### 3. Load an entire directory

Point `VBARunner` at a directory to load all `.bas` files at once. Use this for larger VBA projects with multiple modules.

```typescript
import { VBARunner, assert } from 'vba-runner';

// 1. Load all VBA files in the directory
const vbaRunner = new VBARunner('src/vba');

// 2. Call any procedure defined across those files
const result = vbaRunner.run('CalcTotal', [100, 200, 300]);
assert.strictEqual(result, 600);
```

### 4. Comparing Boolean values

Comparing a VBA `Boolean` result directly against a plain JS `true`/`false` via
`assert.strictEqual(result, true)` fails (VBA `Boolean` values are returned as a
wrapper object). Use the dedicated `assert.isTrue` / `assert.isFalse` instead.

```typescript
import { VBARunner, assert } from 'vba-runner';

const vbaRunner = new VBARunner('src/vba/Sample.bas');
const result = vbaRunner.run('IsPositive', [5]);

assert.isTrue(result);                  // OK
// assert.strictEqual(result, true);    // fails
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
