# Changelog

All notable changes to the VBA Runner extension are documented here.

## [0.3.2] - 2026-07-19

### Fixed

- **Real-VBA engine compatibility updates** — the bundled evaluator now includes the latest fixes for `Null`/error handling, date and numeric type behavior, class and UDT semantics, file I/O edge cases, and built-in function compatibility.
- **VBA language edge cases** — logical operators now preserve Boolean-side numeric subtypes, and `StrConv` now handles kana conversion consistently with real VBA.

## [0.3.1] - 2026-07-15

### Added

- **`Next i, j` multi-variable support** — nested `For` loops can now share a single `Next` statement listing multiple loop variables (e.g. `Next j, i`). Works for both `For…Next` and `For Each…Next`.
- **Fixed-length string declarations** — `Dim s As String * N` is now parsed and evaluated correctly per §5.2.3.1.4.
- **`LenB` / `AscB` / `ChrB`** built-in functions added for byte-oriented string operations.
- **`Split` limit argument** — the optional third argument of `Split(expr, delim, limit)` is now respected.
- **`DatePart` firstdayofweek argument** — the optional `firstdayofweek` parameter is now handled.
- **`Global` keyword** — `Global x As Integer` / `Global Const` / `Global Sub` / `Global Enum` are now recognised as aliases for `Public` (§5.2.3.1).
- **UDT / Enum LSP support** — user-defined `Type` and `Enum` members now appear in completion lists and hover tooltips.

### Fixed

- **`TypeOf x Is ADODB.Recordset`** — dotted library type names in `TypeOf … Is` expressions now parse correctly.
- **`On Error GoTo -1`** — the `-1` label was split into two tokens; now parsed as a single token.
- **`Erase arr1, arr2`** — multiple variables in a single `Erase` statement now work correctly.
- **`ReDim arr1(5), arr2(10)`** — multiple variables in a single `ReDim` statement now supported.
- **`ReDim obj.Arr(n)` / `ReDim .Items(n)`** — member-access and With-expression targets are now accepted.
- **`Const A = 1, B = 2`** — multiple constants in a single `Const` declaration now supported.
- **`Open "f.txt" As #1`** — the `For mode` clause is now optional (defaults to Random).
- **File I/O `#` optional** — `Print`, `Write`, `Input`, `Put`, `Get`, `Seek`, `Line Input`, and `Width` now accept file numbers without the `#` prefix.
- **`Open "f.txt" For Random Len = N`** — `Len` clause with spaces around `=` was rejected; now handled correctly.
- **`Currency` / `Decimal` type overhaul** — internal representation switched to `BigInt` fixed-point arithmetic, fixing rounding errors and comparison bugs.
- **`NPV` function** — returned `NaN` when the values array used 1-based VBA indexing; now fixed.
- **`TypeName` / `VarType`** — subtype information was lost for `Variant` variables; type-promotion rules now applied correctly to binary expressions.
- **With block `Property Get`** — calling a user-defined class `Property Get` from inside a `With` block raised Error 424; fixed.
- **Extract Function** — the refactoring tool generated duplicate `Dim` lines, causing VBA compile errors; fixed.
- **Hover display** — scope labels were incorrectly uppercased; `Const` hover showed wrong type/value; both fixed.

## [0.3.0] - 2026-07-05

### Added

- **Project-wide diagnostics** — VBA warnings and errors are now shown for all `.bas`/`.cls` files in the workspace folder, not just the currently open document.
- **Enhanced hover info** — hovering over a variable now shows its kind (local variable / parameter / constant), scope, declaring file, and inferred type.
- **ByRef / ByVal inlay hints** — Sub/Function parameter declarations display `ByRef` or `ByVal` labels inline to make pass-by-reference explicit.
- **Object variable member completion** — `Dim x As Object` followed by `Set x = ws.Cells(1, 1)` (or another chain expression) now resolves to the actual return type and shows correct member completions.
- **CreateObject inlay hint** — variables assigned via `CreateObject("Scripting.Dictionary")` etc. show the concrete type (e.g. `As Dictionary`) as an inlay hint.

### Fixed

- **Syntax highlighting stopped working** — a regex parenthesis mismatch in the TextMate grammar (`Me)|End)` → `Me|End)`) broke all keyword highlighting.
- **VBA003 diagnostic severity** was incorrectly reported as Warning; it is now shown as a Hint.
- **Indentation: `End` no longer triggered dedent in `Select Case`** — the `decreaseIndentPattern` was changed to require a full keyword (`End If`, `End Sub`, …) which prevented immediate dedent when typing `End`. Restored to `End\b`.
- **Indentation: `End If` over-dedented when an `Else` block was present** — `Else` and `ElseIf` are now in the `increaseIndentPattern` as well, so the code inside them is properly indented.
- **Indentation: `Case` and `Else` are now "double-duty"** — they dedent their own line and indent the code that follows.
- **Inlay hint showed wrong type when a variable had two different `CreateObject` assignments** — the first assignment's type was always used; now returns no hint when types differ.
- **Signature help disappeared while typing inside a string literal argument.**
- **`.cls` files in the same workspace directory were flagged as VBA016 (unknown type)** — the extension now parses co-located `.cls` files and recognises their class names.
- **F12 on a type name (`As MyClass`) now jumps to the `.cls` file definition** — previously opened the wrong file or did nothing.
- **`.cls` file features broken** — definition jump, hover, and Find References did not work because `parseClassBody` did not set location info; fixed.
- **`Test_*` Code Lens always showed "✓ Tested"** — any procedure starting with `Test_` was counted as tested even when it had never run; fixed by requiring an actual test result.
- **VBA mock last-wins** — when multiple `__mocks__` files define the same class, the last-loaded file now correctly replaces earlier stubs.
- **`vba-types.json` stub cache not cleared on deletion** — cached type stubs are now evicted when the file is deleted.
- **Duplicate inlay hints for `As Variant` parameters** — explicitly typed `As Variant` parameters no longer receive a redundant hint.
- **`this module` appeared in hover for local variables and parameters** — removed; only shown for module-level symbols.
- **Multiple additional hover display issues** fixed (scope labels, type name capitalisation, etc.).
- **Inline Variable incorrectly inlined assignments with side effects** (e.g., function calls with observable effects).
- **Reference search scope fixes** — Private Sub/Const/Dim, local variable shadowing per procedure, and cross-file candidate filtering.

### Changed

- Snippets reorganised: `oeg` and `oes` triggers no longer emit a trailing `End Sub` (was redundant when the block closer is auto-inserted by indentation rules).
- `Do … Loop` snippet now also expands on the `wh` trigger.

## [0.2.10] - 2026-07-03

### Fixed

- **Time-only date literals (`#HH:MM:SS#`) raised Error 13 (Type mismatch)** — these literals are now parsed correctly without error.
- **`Class_Terminate` fired prematurely when an object was still referenced by a Dictionary or another variable** — reference counting (`__refCount__`) is now used so `Class_Terminate` is only invoked when the last reference is released.
- **`Class_Terminate` fired prematurely on the return value of a `Function` at scope exit** — the returned object is now protected from early termination when the function's local scope is cleaned up.

## [0.2.9] - 2026-06-28

### Added

- **`Format()` full implementation** — numeric sections (`;` delimiter), `#`/`0` placeholders with zero-suppression, literal characters, percent, thousands separator, scientific notation (`E+`/`E-`). String formatting: `>`/`<`/`@`/`&`/`!`. Date tokens: `q` (quarter), `y` (day of year), `w`/`ww` (weekday/week), `ddddd`/`ttttt` (short date/long time).

### Fixed

- **`eval()` with `Exit Sub`/`Exit Function` at the top level crashed with an unhandled JS exception** — the signal is now caught and suppressed when `eval()` runs outside a procedure body.
- **Opening the same file twice with `Open` did not raise Error 55 (File already open)** — duplicate `Open` on the same path now correctly raises `Error 55`.
- **`FSO TextStream.AtEndOfStream` raised Error 438 (not implemented)** — now returns `True` when all content has been read.
- **`Write #` wrote `True`/`False` for `Boolean` instead of `#TRUE#`/`#FALSE#`** — booleans are now formatted in VBA's canonical CSV representation.
- **`FSO TextStream.ReadAll()` returned the full file after `ReadLine()` advanced the position** — now returns only the remaining unread content.
- **`Scripting.Dictionary.Item` on a non-existent key silently created an empty entry** — missing key access now logs a warning and returns `Empty` without mutating the dictionary.
- **`eval()` raised Error 424 when using arithmetic on a built-in function's return value** — expressions like `UBound(arr) + 1` in `eval()` now evaluate correctly.
- **`Nothing`-typed variables were not correctly converted in some assignment paths** — `Is Nothing` checks now work consistently.
- **`Dim As New` objects did not call `Class_Terminate` when going out of scope** — auto-instantiated class instances now invoke `Class_Terminate` on cleanup.

## [0.2.8] - 2026-06-26

### Fixed

- **Class fields of fixed-size array type (`Dim arr(N) As T`) were not initialized** — accessing `Items(0)` inside a class method raised Error 9, and external access `obj.Items(0)` raised Error 438. Fixed-size array fields are now correctly initialized on construction.
- **Class fields of `Boolean`, `Currency`, `Byte`, `LongLong`, `LongPtr` type defaulted to `Empty`** — these fields now initialize to `False`/`0` as real VBA does.
- **Class fields of UDT type defaulted to `Empty`** — member assignment in `Class_Initialize` raised Error 91. UDT fields are now initialized to a proper UDT instance.
- **Class fields with object/class types defaulted to `Empty` instead of `Nothing`** — `Is Nothing` checks now work correctly on unset object fields.
- **`WithEvents` field member-access event wiring was not connected** — event handlers were not hooked up to the source object, so raising an event never triggered the handler.
- **`Open ... For Append` silently overwrote the file** — writes now correctly append to the end of an existing file.
- **`Err.Raise` with `vbObjectError` offset put internal framework text in `Err.Description`** — the user-supplied description is now preserved. `vbObjectError` (−2147221504) is now a recognized constant.
- **`eval()` fast-path swallowed runtime errors as syntax errors** — runtime errors (division by zero, type mismatch, etc.) are now reported correctly instead of being relabelled as `Compile error: Expected expression`.

## [0.2.7] - 2026-06-24

### Fixed

- **Running/debugging a `Dim` array sized by a variable (e.g. `Dim arr(n)`) silently executed instead of raising the compile error real VBA produces** — both module-level and procedure-level `Dim` declarations now report `Compile error: Constant expression required` before execution. `ReDim` is unaffected, since variable bounds are valid there.

## [0.2.6] - 2026-06-22

### Added

- **Generate Mock Stubs** is now available directly from the Source Actions menu (right-click → Source Action → Generate Mock Stubs), in addition to the existing command palette entry.

### Fixed

- **Generate Mock Stubs command failed after packaging** — the command was launching `npx tsx test-libs/vba-analyzer.ts` as a child process, which is not bundled in the `.vsix`. It now imports the analyzer functions directly so they are included in the esbuild bundle.
- **Extract Method included `End Sub` in the extracted range** — VS Code's line-drag selection sets `range.end` to the next line at character 0, causing `End Sub` to be incorporated into the extracted body. The range is now normalized before processing.
- **README relative links were rewritten incorrectly by vsce** — added `--baseContentUrl` / `--baseImagesUrl` flags to the `vsce package` / `vsce publish` commands so relative links resolve to the correct GitHub paths on the Marketplace page.

## [0.2.5] - 2026-06-21

### Added

- **Built-in functions now validate argument count and support named arguments (`:=`)** — around 200 built-in functions (string/math/conversion/date-time/file-system/registry functions, `MsgBox`, `InputBox`, `Shell`, `InStr`/`InStrB`, etc.) now raise VBA's standard "Argument not optional" / "Wrong number of arguments" errors when called with too few or too many arguments, and accept named arguments the same way user-defined `Sub`/`Function` procedures do (e.g. `InStr(Start:=1, String1:="abc", String2:="a")`).
- **`MockWorksheet.Range` / `MockApplication.Range` now support the 2-argument form `Range(Cell1, Cell2)`** — returns the bounding rectangle between two corner cells (accepting address strings or `Range`/`Cells()` results), normalizing the corners regardless of which order they're passed in.

### Fixed

- **A positional argument gap landing on a required parameter was silently treated as `0`/`Empty`** — `Foo(1, , 3)` where the skipped slot corresponds to a required (non-`Optional`) parameter now raises "Argument not optional" instead of silently passing `0`/`Empty`, matching real VBA. This applies to all call forms: standalone procedures, class methods, interface methods, default properties, and module-qualified calls (`Module1.Foo(1, , 3)`).

## [0.2.4] - 2026-06-21

### Fixed

- **`Rnd`/`DoEvents`/`FreeFile`/`GetObject` couldn't be called without parentheses** — `Int(Rnd * 4) + 1` and similar bare references returned the function object instead of invoking it (`Type mismatch`). These builtins are now auto-invoked like `Now`/`Date`/`Time`/`Timer`.
- **Mocked unset cells returned `0` instead of `Empty`** — `MockWorksheet`'s `.Value` getter (and related cell-read paths) now return `Empty` for unset cells, matching real Excel. `IsEmpty()` and string-context coercion (`""` vs `"0"`) now behave correctly against mocked sheets.
- **README relative links broken on the Marketplace page** — added the missing `repository.directory` field so npm/vsce stop rewriting relative links against the repo root.

## [0.2.3] - 2026-06-21

### Added

- **New compile error: module-level declarations after a procedure** — `Dim`/`Const`/`Type`/`Enum` (and other module-level statements) placed after a `Sub`/`Function`/`Property` are now flagged when running, testing, or debugging a procedure, matching standard VBA's "Only comments may appear after End Sub, End Function, or End Property" rule.

## [0.2.2] - 2026-06-20

### Changed

- **README**: Added a link to the CHANGELOG on GitHub.

## [0.2.1] - 2026-06-20

### Fixed

- **Module-level array bounds referencing `Const` now resolve correctly** — `Dim a(0 To W - 1)` where `W` is a `Const` previously caused `UBound = -1` and a `Subscript out of range` error at runtime. Pass 1 now defers array-bound declarations to Pass 2 (after all constants are resolved), eliminating the incorrect empty-array creation.

## [0.2.0] - 2026-06-14

### Added

- **Auto-insert `()` on Enter** — When pressing Enter after `Sub Foo`, `Function Bar`, or `Property Get/Let/Set` declarations that are missing their parentheses, the parentheses are automatically inserted (VBE-compatible behavior).
- **Auto-insert `End Sub/Function/Property` on Enter** — Pressing Enter on a `Sub`/`Function`/`Property` header line automatically appends the matching `End` statement on the next line (experimental, VBE-compatible).
- **Auto-insert `End If/For/With/Do/While/Select Case` on Enter** — Pressing Enter on `If`, `For`, `For Each`, `With`, `Do`, `Do While`, `Do Until`, `While`, and `Select Case` header lines automatically inserts the corresponding `End` or `Loop`/`Wend` statement.
- **Auto-indent `Else`/`ElseIf`/`Case`** — Pressing Enter on these mid-block keywords correctly adjusts indentation to match the enclosing block.
- **Internationalization (l10n)** — All user-visible messages in the extension are now localizable via `@vscode/l10n`. The Japanese locale bundle (`bundle.l10n.ja.json`) is included. Bundles are loaded at activation for both packaged and F5 dev-mode runs.

### Fixed

- **F5 dev mode l10n bundle not loaded** — When running the extension via F5 (Extension Development Host), `vscode.l10n.bundle` is `undefined`. The extension now manually loads the locale bundle via `@vscode/l10n`'s `l10n.config()` at activation time.
- **Dead store false positive: class module fields** — Module-level variables (fields) assigned inside one procedure and read inside another were incorrectly flagged as dead stores. Fields are now treated as always-live across all procedures in a module.
- **Dead store false positive: `Set var = Nothing`** — COM object release (`Set x = Nothing`) was incorrectly flagged as a dead store. These assignments are now recognized as side-effecting operations and excluded from dead store analysis.
- **Parser: `Event` declaration rejects reserved-identifier names** — `Public Event Open()` now correctly produces a compile error per MS-VBAL §3.3.5.2. Previously, the permissive `isNameToken` check allowed any keyword as an event name.
- **Auto-format on Enter: `autoParentheses` and `autoEndBlock` are now mutually exclusive** — Previously, both handlers could fire on the same keystroke, resulting in double edits.

## [0.1.9] - 2026-06-13

### Added

- **VBA013 lint rule: missing `Option Explicit`** — Files without `Option Explicit` are now flagged as an Error (red icon) since undeclared variable names silently become `Variant` and typos go undetected until runtime.
- **Lint severity tiers** — Lint diagnostics are now classified into three tiers: Error (severity 1, red), Warning (severity 2, yellow), and Information (severity 3, blue). `Select Case` without `Case Else` (VBA005) and unreachable code (VBA010) have been reclassified from Information to Warning.
- **`Class_Terminate` fires on scope exit** — Local objects created with `Set x = New ClassName` now automatically trigger `Class_Terminate` when the enclosing Sub/Function returns, without requiring an explicit `Set x = Nothing`.
- **`VarPtr` / `StrPtr` / `ObjPtr`** — These memory-address functions now return a unique non-zero `Long` value (stub implementation sufficient for code that only checks IsNull/non-zero).
- **`AddressOf` with module member syntax** — `AddressOf Module.Procedure` (§5.6.16.8) is now parsed correctly.
- **`Mid` assignment statement** — `Mid(str, start, length) = value` (§5.4.3.8), including `MidB`, `Mid$`, and `MidB$` variants, is now fully implemented.
- **`Def`-Directives** — `DefInt`, `DefStr`, `DefDbl`, `DefBool`, `DefLng`, `DefCur`, `DefSng`, `DefDate`, `DefObj`, `DefVar` (§5.2.2) are now parsed and applied as implicit type declarations for variables whose names begin with the specified letters.
- **`Optional` parameter order check** — A compile error (§5.3.1.5) is now raised when a non-optional parameter follows an optional one in the same parameter list.

### Fixed

- **Code Lens "n references"** — Clicking the reference count in the Code Lens no longer searches for the `Function`/`Sub` keyword. The cursor position is now placed on the procedure name, so the correct references are found.

## [0.1.8] - 2026-06-13

### Fixed

- **Parser: `statement-keyword` correctly rejected as module-level procedure names** — `Function Open()`, `Sub Close()`, `Function Print()` etc. now produce a compile error per §3.3.5.2 of the VBA Language Specification. Class member methods with reserved-word names (e.g. `obj.Open`, `ws.Close`) continue to work correctly via `unrestricted-name` rules.
- **Parser: dead-code lookahead removed** — The `hasFileOpenSyntaxAhead()` pre-scan that guarded `Open`/`Close` dispatch was replaced with direct unconditional dispatch, since `Open` and `Close` are reserved-identifiers and cannot appear as identifiers at statement position.
- **Parser: file I/O statement dispatch cleaned up** — `Print`, `Write`, `Input`, `Put`, `Seek` are reserved-identifiers and are now dispatched unconditionally to their file I/O parse functions. Previously a missing `#` would silently fall through to the identifier branch.

## [0.1.7] - 2026-06-12

### Added

- **Cross-file Go to Definition (F12)** — Jumping to a Sub/Function definition now searches all `.bas`, `.cls`, and `.frm` files in the same directory, not just the currently open file. Files are scanned lazily on the first F12 press so there is no startup overhead.
- **Cross-file Find All References (Shift+F12)** — Reference search now spans the entire directory. Module-level procedures and variables are searched across all VBA files; local variables remain scoped to their own file.

## [0.1.6] - 2026-06-12

### Fixed

- **Contextual keywords as parameter names** — Keywords such as `text`, `binary`, `compare`, `output`, `lib`, `alias` etc. can now be used as Sub/Function parameter names without triggering a "Expected parameter name" diagnostic.
- **`!` and `^` as type-declaration characters** — `Dim x!` (Single) and `Dim x^` (LongLong) now work correctly. Previously they conflicted with the bang member-access operator (`dict!Key`) and the exponentiation operator (`x^2`). A lookahead disambiguates: `!` and `^` are consumed as type suffixes only when not followed by an alphanumeric character or `_`.
- **Contextual keyword labels** — Labels such as `Error:` inside procedure bodies no longer produce a spurious "unexpected token in expression ':'" diagnostic.

## [0.1.5] - 2026-06-11

### Added

- **Generate Test stub** — Clicking the "Untested" Code Lens button generates a `Test_<ProcedureName>` stub. On first use, you choose whether to place tests in the same file or a separate `<FileName>Test.bas`. The choice is saved to workspace settings (`vba-runner.test.location`). If the test already exists, the cursor jumps to it instead.
- **`vba-runner.test.location` setting** — Workspace setting to control where test stubs are placed (`sameFile` / `separateFile`).

### Changed

- Updated README with complete feature overview: Code Lens buttons, all refactoring commands, test generation, mock generation, and Find References.

## [0.1.4] - 2026-06-10

### Fixed

- **Code Lens "Run" now uses unsaved editor content** — Previously, clicking "Run" above a Sub/Function read the file from disk, ignoring any unsaved edits in the editor. The procedure appeared in Code Lens (parsed from the live buffer) but was not found at runtime, resulting in "Sub or Function not defined". Now the in-memory document text is used whenever available.

## [0.1.3] - 2026-06-10

### Changed

- Added Marketplace icon.

## [0.1.0] - 2026-06-09

### Added

- Initial release: VBA execution engine, Code Lens (Run / Debug / Test), LSP-based diagnostics, syntax highlighting, and basic refactoring commands.
