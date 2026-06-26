# Changelog

All notable changes to the `vba-runner` npm package are documented here.

## [0.1.1-alpha.10] - 2026-06-26

### Added

- **`quiet` and `onPrint` options for `VBARunner`** — pass `{ quiet: true }` to suppress the `[PASS]` log lines that `run()` emits, and `{ onPrint: (s) => ... }` to redirect `Debug.Print` output to a custom handler (e.g. `stderr` or a collector array) instead of `console.log`.

### Changed

- **`run()` and `eval()` now return plain JS `boolean` for VBA `Boolean` results** — previously a `VbaBoolean` wrapper object was returned, which failed `assert.strictEqual(result, true)` and caused JS truthiness bugs (`if (vbaFalse)` was truthy). Both methods now convert `VbaBoolean` to `true`/`false` automatically. The engine internally still uses `VbaBoolean` so VBA arithmetic (`True + 1 = 0`) is unaffected.

### Removed

- **`assert.isTrue` and `assert.isFalse` have been removed** — they were the workaround for the `VbaBoolean` wrapper. Use `assert.strictEqual(result, true)` / `assert.strictEqual(result, false)` or any standard assertion library directly.

### Fixed

- **`eval()` with multi-line code only executed the first line** — subsequent lines were silently dropped when input contained newlines. All lines are now executed in order.
- **`eval()` with a bare assignment statement (`x = 10`, `arr(1) = "a"`) performed no assignment** — the statement was parsed as an expression and the value was discarded. Bare assignments at the top level are now correctly executed.
- **`eval('x + 1')` returned a different result than `eval('... : x + 1')` (trailing expression in multi-statement input)** — the two evaluation paths now agree.
- **`eval()` did not process `On Error` directives and leaked `On Error` state from class methods** — each `eval()` call now runs in a clean error-handler context, and class method calls no longer leave behind a stale `On Error` handler in the caller's frame.
- **`eval()` did not enforce `Option Explicit`** — undeclared variable references in `eval()` code were silently treated as `Empty` even when `Option Explicit` was declared in the loaded module.
- **`eval()` lost module-level `Dim` declarations after a file was loaded** — variables declared with `Dim` at module scope were invisible to `eval()` calls made after constructing `VBARunner` from a file or directory.
- **`run()` did not write `ByRef` parameter changes back to the caller's `args` array** — mutations to `ByRef` parameters inside a called `Sub`/`Function` are now reflected in the JS array passed as `args`, consistent with VBA semantics.
- **`Function` declared `As Type()` (array return type) was treated as a scalar** — the trailing `()` was silently discarded, causing `IsArray()` to return `False` for the result and array-element access to fail.
- **Class method `On Error` state leaked into the caller** — an `On Error GoTo` inside a class method that was never triggered left the handler active when control returned to the caller.
- **`Join()` on an array with a non-zero lower bound raised an error** — `Join(Array(1, 2), ",")` where `Array()` produces a 0-based array worked, but arrays declared with `Dim arr(1 To 3)` caused `Subscript out of range`. `Join` now handles any lower bound.
- **`Open ... For Append` silently overwrote the file** — opening an existing file in Append mode now correctly positions writes at the end of the file instead of truncating it.
- **Class fields of type `Boolean`, `Currency`, `Byte`, `LongLong`, `LongPtr` defaulted to `Empty` instead of `0`/`False`** — these fields now match real VBA's default-initialization rules.
- **Class fields of UDT type defaulted to `Empty`** — member assignment inside `Class_Initialize` raised Error 91. UDT fields are now initialized to a proper UDT instance on construction.
- **Class fields with object types defaulted to `Empty` instead of `Nothing`** — `Is Nothing` checks and `WithEvents` wiring that depend on the Nothing default now work correctly.
- **`WithEvents` field member-access event wiring was not connected** — event handlers declared with `WithEvents` were not hooked up to the source object's event sink, so raising an event on the source object never called the handler.
- **Class fields declared as fixed-size arrays (`Dim arr(N) As T`) were not initialized** — accessing `obj.Items(0)` inside a class method raised Error 9 (Subscript out of range), and `obj.Items(0) = val` from outside raised Error 438. Fixed-size array fields are now initialized to a properly-dimensioned array on construction, and external indexed access (`obj.Field(i)`) is now supported for both reads and writes.
- **`Err.Description` was populated with internal framework text instead of the user's message** — `Err.Raise vbObjectError + 1, , "my message"` set `Err.Description` to the error-construction code path's label instead of `"my message"`. `vbObjectError` (−2147221504) was also unrecognized and reported as an unknown error number.
- **`eval()` fast-path silently swallowed runtime errors as syntax errors** — a runtime error thrown during expression evaluation (e.g. division by zero, type mismatch) was caught by the expression-parser fallback and re-reported as `Compile error: Expected expression`, hiding the real error.

## [0.1.1-alpha.9] - 2026-06-24

### Fixed

- **`Dim` array bounds accepted a variable without raising VBA's compile error** — real VBA requires `Dim arr(n)` bounds to be a constant expression; passing a variable previously executed without error instead of failing at compile time. Both module-level and procedure-level `Dim` declarations now raise `Compile error: Constant expression required` when the bound isn't constant. `ReDim` is unaffected, since variable bounds are valid there.

## [0.1.1-alpha.8] - 2026-06-21

### Added

- **Built-in functions now validate argument count and support named arguments (`:=`)** — around 200 built-in functions (string/math/conversion/date-time/file-system/registry functions, `MsgBox`, `InputBox`, `Shell`, `InStr`/`InStrB`, etc.) now raise VBA's standard "Argument not optional" / "Wrong number of arguments" errors when called with too few or too many arguments, and accept named arguments the same way user-defined `Sub`/`Function` procedures do (e.g. `InStr(Start:=1, String1:="abc", String2:="a")`).
- **`MockWorksheet.Range` / `MockApplication.Range` now support the 2-argument form `Range(Cell1, Cell2)`** — returns the bounding rectangle between two corner cells (accepting address strings or `Range`/`Cells()` results), normalizing the corners regardless of which order they're passed in.

### Fixed

- **A positional argument gap landing on a required parameter was silently treated as `0`/`Empty`** — `Foo(1, , 3)` where the skipped slot corresponds to a required (non-`Optional`) parameter now raises "Argument not optional" instead of silently passing `0`/`Empty`, matching real VBA. This applies to all call forms: standalone procedures, class methods, interface methods, default properties, and module-qualified calls (`Module1.Foo(1, , 3)`).

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
