# VBA Runner — VS Code Extension

A VS Code extension for VBA development. When you open `.bas` / `.cls` / `.frm` files, rich editing support powered by LSP (Language Server Protocol) becomes available.

> 日本語ドキュメントは [README.ja.md](./README.ja.md) をご覧ください。

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

## Features

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

### Code Completion

Suggestions for VBA keywords, built-in functions, and procedures defined in your source files appear as you type. Also triggered by `.`.

### Call Graph

Open the Command Palette (`Ctrl+Shift+P`) and run:

- **VBA: Show Call Graph** — Display the call graph starting from the procedure at the cursor
- **VBA: Show in Call Graph** — Highlight the procedure at the cursor in the full graph

### Introduce Variable Refactoring

Run **Refactor: Introduce Variable** from the Command Palette to extract a selected expression into a variable.

### VBA Debugger Integration

With a `.bas` file open, press `F5` to launch the file with the VBA Runner debugger (no `launch.json` required).

## Settings

| Setting | Default | Description |
|---|---|---|
| `vba-runner.lint.enabled` | `false` | Enable all lint diagnostics (VBA001, etc.) |
| `vba-runner.lint.enabledCodes` | `[]` | Enable specific lint codes (e.g. `["VBA009"]`) |
| `vba-runner.editor.autoLineContinuation` | `true` | Auto-insert line continuation `_` when pressing Enter mid-expression |
| `vba-runner.editor.autoKeywordCasing` | `true` | Auto-correct keyword casing on confirm (like VBE behavior) |

## Repository

[github.com/jca02266/vba-runner](https://github.com/jca02266/vba-runner)
