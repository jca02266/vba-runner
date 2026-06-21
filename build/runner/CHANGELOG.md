# Changelog

All notable changes to the `vba-runner` npm package are documented here.

## [0.1.1-alpha.7] - 2026-06-21

### Fixed

- **`Rnd`/`DoEvents`/`FreeFile`/`GetObject` couldn't be called without parentheses** — `Int(Rnd * 4) + 1` and similar bare references returned the function object instead of invoking it (`Type mismatch`). These builtins are now auto-invoked like `Now`/`Date`/`Time`/`Timer`.
- **Mocked unset cells returned `0` instead of `Empty`** — `MockWorksheet`'s `.Value` getter (and related cell-read paths) now return `Empty` for unset cells, matching real Excel. `IsEmpty()` and string-context coercion (`""` vs `"0"`) now behave correctly against mocked sheets.
- **README relative links broken on the npm package page** — added the missing `repository.directory` field so npm stops rewriting relative links against the repo root.

### Changed

- **`docs/MOCK_GUIDE.md` Part 0 now covers missing built-in functions, not just Excel/Windows-dependent ones** — broadened the intro and added a "0-5. Stubbing an unimplemented built-in function (RGB example)" section, since `__mocks__/` works just as well for plain unimplemented builtins (e.g. `RGB`) as it does for environment-dependent ones (`MsgBox`, `Range`, etc.).

## [0.1.1-alpha.6] - 2026-06-21

### Added

- **`excelStub` now accepts a `MockApplication` subclass instance** — `new VBARunner(path, { excelStub: myCustomApp })` lets you extend the built-in mock (e.g. to add `Application.OnKey`/`OnTime`) without losing the built-in `ActiveSheet`/`Sheets`/`Range` behavior.
- **`MockApplication`, `MockWorksheet`, `MockRange`, `MockRows`, `MockColumns`, `MockWorkbook` are now exported** from the package, along with the `VbaType`, `VbaDefaultProperty`, `VbaIterable`, `VbaComObject` types.
- **New `allowTopLevelStatements` option** for `evalVBASingle`/`evalVBAModules` (default `true`) to toggle the module-level-statement extension on or off.

### Changed

- **`VBARunner` now enforces standard VBA's module-level declaration ordering** — `Dim`/`Const`/`Type`/`Enum` (and other statements) placed after a `Sub`/`Function`/`Property` now raise a compile error matching real VBA ("Only comments may appear after End Sub, End Function, or End Property") instead of being silently tolerated. If an existing `.bas`/`.cls` file relies on the previous permissive ordering, move module-level declarations above all procedures.

## [0.1.1-alpha.5] - 2026-06-20

### Changed

- **README**: Added a link to the CHANGELOG on GitHub.

## [0.1.1-alpha.4] - 2026-06-20

### Fixed

- **Module-level array bounds referencing `Const` now resolve correctly** — `Dim a(0 To W - 1)` where `W` is a `Const` previously caused `UBound = -1` and a `Subscript out of range` error at runtime. Pass 1 now defers array-bound declarations to Pass 2 (after all constants are resolved), eliminating the incorrect empty-array creation.

## [0.1.1-alpha.3] - 2026-06-14

### Fixed

- **Dead store false positive: class module fields** — Module-level variables assigned in one procedure and read in another were incorrectly flagged as dead stores. Fields are now treated as always-live.
- **Dead store false positive: `Set var = Nothing`** — COM object release was incorrectly flagged as a dead store. These assignments are now excluded from dead store analysis.
- **Parser: `Event` declaration rejects reserved-identifier names** — `Public Event Open()` now correctly produces a compile error per MS-VBAL §3.3.5.2.
- **`.cls` file header stripping** — `vba-analyzer` and `vba-formatter` now correctly strip the `VERSION 1.0 CLASS` / `BEGIN` / `END` header block from Excel-exported `.cls` files before parsing.

## [0.1.1-alpha.2] - 2026-06-13

### Added

- **`Class_Terminate` fires on scope exit** — Local objects created with `Set x = New ClassName` now automatically trigger `Class_Terminate` when the enclosing Sub/Function returns, without requiring an explicit `Set x = Nothing`.
- **`VarPtr` / `StrPtr` / `ObjPtr`** — These memory-address functions now return a unique non-zero `Long` value (stub implementation sufficient for code that only checks IsNull/non-zero).
- **`AddressOf` with module member syntax** — `AddressOf Module.Procedure` (§5.6.16.8) is now parsed correctly.
- **`Mid` assignment statement** — `Mid(str, start, length) = value` (§5.4.3.8), including `MidB`, `Mid$`, and `MidB$` variants, is now fully implemented.
- **`Def`-Directives** — `DefInt`, `DefStr`, `DefDbl`, `DefBool`, `DefLng`, `DefCur`, `DefSng`, `DefDate`, `DefObj`, `DefVar` (§5.2.2) are now parsed and applied as implicit type declarations.
- **`Optional` parameter order check** — A compile error (§5.3.1.5) is now raised when a non-optional parameter follows an optional one.

### Fixed

- **`Class_Terminate` early termination** — Objects borrowed from a Dictionary or returned from a function are no longer incorrectly terminated at the callee's scope exit.

## [0.1.1-alpha.1] - 2026-06-13

### Added

- **Unified CLI** — `vba-runner` now provides a single entry point with subcommands: `run`, `analyze`, `format`, `parse-check`.

## [0.1.1-alpha.0] - 2026-06-12

### Fixed

- **Contextual keywords as parameter names** — Keywords such as `text`, `binary`, `compare`, `output`, `lib`, `alias` etc. can now be used as Sub/Function parameter names without a parse error.
- **`!` and `^` as type-declaration characters** — `Dim x!` (Single) and `Dim x^` (LongLong) now work correctly without conflicting with the bang member-access operator (`dict!Key`) or the exponentiation operator (`x^2`).
- **Contextual keyword labels** — Labels such as `Error:` inside procedure bodies are now parsed correctly.

## [0.1.0-alpha.0] - 2026-06-12

Initial alpha release.
