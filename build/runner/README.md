# vba-runner

A TypeScript-based VBA execution engine that runs and tests VBA code without Excel.
Load `.bas` files from TypeScript and call VBA procedures directly, or use the bundled CLI tools for static analysis, formatting, and syntax checking.

> 日本語ドキュメントは [README.ja.md](./README.ja.md) をご覧ください。

[CHANGELOG](https://github.com/jca02266/vba-runner/blob/main/build/runner/CHANGELOG.md)

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

The built-in mock only persists `Value` reads/writes. Formatting properties like
`Interior.Color` are no-ops that don't retain state, and `Application.OnKey` /
`Application.OnTime` aren't implemented (calling them throws). To test code that
depends on these, replace `Application` (or any other builtin) entirely with your
own mock object via `vbaRunner.evaluator.setBuiltinOverride(name, value)`.

```typescript
const vbaRunner = new VBARunner('src/vba/KeyHandler.bas');

let registered: [string, string] | null = null;
const customApplication = {
  OnKey: (key: string, procedureName: string) => { registered = [key, procedureName]; },
};
vbaRunner.evaluator.setBuiltinOverride('Application', customApplication);

vbaRunner.run('SetupKeyHandlers', []);
assert.ok(registered !== null);
```

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
