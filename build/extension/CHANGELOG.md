# Changelog

All notable changes to the VBA Runner extension are documented here.

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
