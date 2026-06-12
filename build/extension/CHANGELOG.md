# Changelog

All notable changes to the VBA Runner extension are documented here.

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
