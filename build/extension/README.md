# VBA Runner — VS Code Extension

A VS Code extension for VBA development. When you open `.bas` / `.cls` / `.frm` files, rich editing support powered by LSP (Language Server Protocol) becomes available.

**[日本語](./README.ja.md)** | **[CHANGELOG](./CHANGELOG.md)**

## Installation

Install from the VS Code Marketplace by searching for **VBA Runner**, or run:

```bash
code --install-extension jca02266.vba-runner
```

**For development (when cloning the repository)**:

1. Open this repository in VS Code
2. Press `F5` to launch the Extension Development Host
3. Open a `.bas` / `.cls` / `.frm` file in the new window

## Supported File Types

| Extension | Type |
|---|---|
| `.bas` | Standard Module |
| `.cls` | Class Module |
| `.frm` | Form Module |

## Editor Support (LSP)

### Hover

Hovering over a symbol (Sub / Function / variable / constant / class / event) shows its signature in a popup.

```vb
Sub CalcSum(a As Integer, b As Integer)
'         ↑ Hover here → "Sub CalcSum(a As Integer, b As Integer)"
```

### Go to Definition

Place the cursor on a symbol and press `F12` to jump to its declaration.

| Symbol | Example |
|---|---|
| Sub / Function / Property | `Sub Foo()`, `Function Bar() As Long` |
| Variable | `Dim x As Integer`, `Public count As Long` |
| Constant | `Const MAX As Long = 100` |
| Class | `Class MyClass` |
| Event | `Event DataChanged(newVal As Variant)` |

### Find References

Place the cursor on a symbol and press `Shift+F12` to list all locations where it is referenced.

### Code Completion

Suggestions for VBA keywords, built-in functions, and procedures defined in your source files appear as you type. Also triggered by `.`.

## Code Lens

Inline action buttons appear above each procedure declaration.

| Button | Action |
|---|---|
| `▶ Run` | Run the procedure with VBA Runner |
| `🐛 Debug` | Step through with the debugger |
| `N references` | List all reference sites |
| `Untested` / `✓ Tested` | Generate a test stub / Jump to the test function |
| `📊 Show in Call Graph` | Highlight in the call graph |

## Refactoring

Available from the Command Palette (`Ctrl+Shift+P`) or via Code Lens.

| Command | Description |
|---|---|
| **Refactor: Introduce Variable** | Extract a selected expression into a variable |
| **Refactor: Extract Function** | Extract selected code into a new procedure |
| **Refactor: Extract Constant** | Extract a selected literal into a constant |
| **Refactor: Inline Variable** | Inline a variable into all its usage sites |
| **Refactor: Introduce With** | Wrap repeated object references in a `With` block |
| **Refactor: Remove Unused Variables** | Delete declarations of unused variables |
| **Refactor: Organize Declarations** | Move variable declarations to the top of the procedure |

## Call Graph

Open the Command Palette (`Ctrl+Shift+P`) and run:

- **VBA: Show Call Graph** — Display the call graph starting from the procedure at the cursor
- **VBA: Show in Call Graph** — Highlight the procedure at the cursor in the full graph

## Testing

### Test Stub Generation

Click the **Untested** Code Lens button to generate a `Test_<ProcedureName>` stub. On first use, you will be asked where to place tests:

- **Same file** — Appends the stub to the end of the current `.bas` file
- **Separate file** — Creates `<FileName>Test.bas` and appends there

The choice is saved to workspace settings (`vba-runner.test.location`).

### Mock Skeleton Generation

Run **VBA: Generate Mocks** from the Command Palette to analyze Excel object dependencies (`Worksheet`, `Range`, etc.) in the source file and generate a mock skeleton at `__mocks__/ExcelObjects.bas`.

## VBA Debugger Integration

With a `.bas` file open, press `F5` to launch the file with the VBA Runner debugger (no `launch.json` required).

## Settings

| Setting | Default | Description |
|---|---|---|
| `vba-runner.lint.enabled` | `false` | Enable all lint diagnostics (VBA001, etc.) |
| `vba-runner.lint.enabledCodes` | `[]` | Enable specific lint codes (e.g. `["VBA009"]`) |
| `vba-runner.editor.autoLineContinuation` | `true` | Auto-insert line continuation `_` when pressing Enter mid-expression |
| `vba-runner.editor.autoKeywordCasing` | `true` | Auto-correct keyword casing on confirm (like VBE behavior) |
| `vba-runner.test.location` | _(unset)_ | Where to place test stubs (`sameFile` / `separateFile`). Prompted on first use if unset. |

## Repository

[github.com/jca02266/vba-runner](https://github.com/jca02266/vba-runner)
