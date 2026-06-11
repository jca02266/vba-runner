# vba-runner

A TypeScript-based VBA execution engine that runs and tests VBA code without Excel.
Load `.bas` files from TypeScript and call VBA procedures directly, or use the bundled CLI tools for static analysis, formatting, and syntax checking.

> 日本語ドキュメントは [README.ja.md](./README.ja.md) をご覧ください。

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
