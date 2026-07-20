# VBA Runner — VS Code Extension

A VS Code extension for VBA development. When you open `.bas` / `.cls` / `.frm` files, rich editing support powered by LSP (Language Server Protocol) becomes available.

**[日本語](./README.ja.md)** | **[CHANGELOG](./CHANGELOG.md)** | **[VBA Runner Project](../../README.md)** | **[Documentation hub](../../docs/README.md)**

> Audience: package users. For goal-oriented guides, see the [documentation hub](../../docs/README.md).

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

### Signature Help

When you type `(` or `,` inside a function call, a popup shows the parameter names and the current parameter position. Works for built-in VBA functions (`MsgBox`, `Format`, etc.) as well as your own Sub / Function declarations.

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

### Rename Symbol

Place the cursor on a symbol and press `F2` to rename it across the file. All references are updated at once.

### Code Completion

Suggestions for VBA keywords, built-in functions, and procedures defined in your source files appear as you type.

#### Member Completion (`.`)

Typing `.` after a variable shows members specific to its declared type. Supports:

| Type | Trigger |
|---|---|
| `Scripting.Dictionary` | `Dim d As Scripting.Dictionary` → `d.` |
| `Scripting.FileSystemObject` | `Dim fso As Scripting.FileSystemObject` → `fso.` |
| `ADODB.Recordset` / `ADODB.Connection` | `Dim rs As ADODB.Recordset` → `rs.` |
| `RegExp` / `VBScript.RegExp` | `Dim re As RegExp` → `re.` |
| `Collection` | `Dim col As Collection` → `col.` |
| `Range` / `Worksheet` / `Workbook` | `Dim ws As Worksheet` → `ws.` |
| `Sheets` / `Application` | `Dim app As Application` → `app.` |
| User-defined classes | `Dim obj As MyClass` → `obj.` |

Cross-module completion is also supported: classes declared in other open `.bas` / `.cls` files in the workspace are recognized automatically.

#### Chain Access Resolution

Member completion works across chained accesses. The return type of each member is tracked, so subsequent `.` completions resolve correctly:

```vb
Dim ws As Worksheet
ws.Cells.         ' → Range members (Cells returns Range)
ws.Range("A1").Offset(1, 0).  ' → Range members (Offset returns Range)
ws.Parent.        ' → Workbook members (Parent returns Workbook)
```

#### With Block Completion

Inside a `With` block, typing `.` at the start of a line shows the members of the `With` object:

```vb
With ws
    .Cells.       ' → Range members
    .Name         ' → "Name" property of Worksheet
End With
```

#### Snippets

| Prefix | Expands to |
|---|---|
| `fe` | `For Each ... In ... Next` |
| `for` | `For ... To ... Next` |
| `sc` | `Select Case ... Case Else ... End Select` |
| `if` | `If ... Then ... Else ... End If` |
| `oeg` | `On Error GoTo ... ErrHandler` pattern |
| `wi` | `With ... End With` |
| `sub` | `Sub ... End Sub` |
| `fn` | `Function ... End Function` |
| `do` | `Do While ... Loop` |
| `dim` | `Dim` variable declaration |
| `dict` | `Dim … As Object` + `Set … = CreateObject("Scripting.Dictionary")` |
| `fso` | `Dim … As Object` + `Set … = CreateObject("Scripting.FileSystemObject")` |
| `regex` | `Dim … As Object` + `Set … = CreateObject("VBScript.RegExp")` + Pattern / Global |
| `adors` | `Dim … As Object` + `Set … = CreateObject("ADODB.Recordset")` |
| `adocn` | `Dim … As Object` + `Set … = CreateObject("ADODB.Connection")` + ConnectionString |
| `cobj` | `Dim … As Object` + `Set … = CreateObject("ProgID")` (汎用) |

### Document Symbols (Outline)

The outline panel (`Ctrl+Shift+O`) and the workspace symbol search (`Ctrl+T`) list all Sub / Function / Property / class members defined in your VBA files.

Section divider comments using `' --- Name ---` or `' === Name ===` are also recognized as Namespace symbols in the outline, making large modules easier to navigate.

```vb
' ─── Initialization ───────────────────────────────
Public Sub Initialize()
    ...
End Sub

' === Data Processing ===
Public Function Process(data) As Long
    ...
End Function
```

## Diagnostics

The following diagnostic rules are reported as you type:

| Code | Severity | Rule | Condition |
|---|---|---|---|
| — | Error | Parse error | Syntax error detected by the parser |
| VBA001 | Warning | `ByVal`/`ByRef` missing | Parameter has no explicit passing modifier |
| VBA009 | Warning | Dead store | Variable is assigned but never read |
| VBA011 | Hint | Range access | `Sheets("name")` should use a typed variable |
| VBA013 | Warning | `Option Explicit` missing | File lacks `Option Explicit` |
| VBA014 | Warning | Unused variable | Variable declared but never referenced |
| VBA016 | Warning | Unknown type | `Dim x As UnknownType` — type is not recognized |

**Quick Fixes** are provided for:

- **VBA013** — *Add 'Option Explicit'*: inserts `Option Explicit` at the top of the file
- **VBA016** — *Add 'TypeName' to vba-types.json*: appends a placeholder entry for the type
- **VBA016** — *Initialize vba-types.json with all COM type definitions*: creates `vba-types.json` pre-populated with all built-in COM types (shown when the file does not yet exist)

Lint rules (VBA001, VBA009, VBA014, etc.) are off by default. Enable them via settings:

```json
{ "vba-runner.lint.enabled": true }
```

or selectively:

```json
{ "vba-runner.lint.enabledCodes": ["VBA009", "VBA014"] }
```

## External Type Definitions (`vba-types.json`)

To add member completion for types not built into the extension — such as custom COM objects, mock classes, or Excel types not yet listed — create `vba-types.json` in your workspace root:

```json
{
  "MyComObject": [
    { "label": "DoWork",  "kind": "Method",   "detail": "DoWork(arg As String) As Boolean" },
    { "label": "Status",  "kind": "Property", "detail": "Status As Long" }
  ],
  "MyHelper": [
    { "label": "Compute", "kind": "Function", "detail": "Compute(x As Long) As Double", "returnType": "myresult" }
  ]
}
```

**Fields:**

| Field | Values | Description |
|---|---|---|
| `label` | string | Member name |
| `kind` | `"Method"` / `"Function"` / `"Property"` / `"Variable"` / `"Constant"` | Icon and category |
| `detail` | string | Signature shown in the completion popup |
| `returnType` | string (lowercase type name) | Return type for chain access resolution |

**Initialization via Quick Fix:**

When a VBA016 diagnostic appears for an unknown type, the Quick Fix menu offers:

- *Add 'TypeName' to vba-types.json* — adds a placeholder entry for that type
- *Initialize vba-types.json with all COM type definitions* — creates the file pre-populated with all built-in types (Scripting.Dictionary, Range, Worksheet, etc.)

`vba-types.json` entries take **priority over** the built-in definitions, so you can override any built-in type's member list.

The file is reloaded automatically whenever it changes.

## Code Lens

Inline action buttons appear above each procedure declaration.

| Button | Action |
|---|---|
| `▶ Run` | Run the procedure with VBA Runner |
| `🐛 Debug` | Step through with the debugger |
| `N references` | List all reference sites |
| `Untested` / `✓ Tested` | Generate a test stub / Jump to the test function |
| `📊 Show in Call Graph` | Highlight in the call graph |
| `✓ Nms` | Test passed (shown after running tests, e.g. `✓ 3ms`) |
| `✗ message` | Test failed with the first line of the error message |

## Formatting

Press `Shift+Alt+F` (or right-click → **Format Document**) to auto-format the file. Also works with `"editor.formatOnSave": true` in VS Code settings.

Formatting rules include:
- Consistent indentation for `Sub` / `Function` / `If` / `For` / `With` / `Select Case`
- `Case` labels are aligned with the `Select Case` keyword
- Keyword casing is normalized to the standard VBA style

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

### Inline Test Results

After running tests with the **▶ Run** Code Lens on a test procedure, the result is shown inline:

- `✓ 3ms` — test passed in 3 ms
- `✗ Expected 1 but got 2` — test failed with the first line of the assertion message

### Mock Skeleton Generation

Run **VBA: Generate Mocks** from the Command Palette to analyze Excel object dependencies (`Worksheet`, `Range`, etc.) in the source file and generate a mock skeleton at `__mocks__/ExcelObjects.bas`.

## VBA Debugger Integration

With a `.bas` file open, press `F5` to launch the file with the VBA Runner debugger (no `launch.json` required).

### `Debug.Print` Output

`Debug.Print` output is directed to a dedicated **VBA Debug** Output Channel (separate from the VBA Runner log), which opens automatically when output is produced.

## Settings

| Setting | Default | Description |
|---|---|---|
| `vba-runner.lint.enabled` | `false` | Enable all lint diagnostics (VBA001, etc.) |
| `vba-runner.lint.enabledCodes` | `[]` | Enable specific lint codes (e.g. `["VBA009"]`) |
| `vba-runner.editor.autoLineContinuation` | `true` | Auto-insert line continuation `_` when pressing Enter mid-expression |
| `vba-runner.editor.autoKeywordCasing` | `true` | Auto-correct keyword casing on confirm (like VBE behavior) |
| `vba-runner.test.location` | _(unset)_ | Where to place test stubs (`sameFile` / `separateFile`). Prompted on first use if unset. |

## Documentation

- [Documentation hub](../../docs/README.md) — Goal-oriented guides (users / developers)
- [LSP.md](../../LSP.md) — LSP design and implementation (for developers)
- [REFERENCE.md](../../REFERENCE.md) — Detailed specs
- [README.md](../../README.md) — Project overview

## Repository

[github.com/jca02266/vba-runner](https://github.com/jca02266/vba-runner)
