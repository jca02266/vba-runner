# Changelog

All notable changes to the `vba-runner` npm package are documented here.

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
