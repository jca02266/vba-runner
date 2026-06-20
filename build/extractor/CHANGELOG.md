# Changelog

All notable changes to the `vba-extractor` npm package are documented here.

## [Unreleased]

### Added

- **`scripts/Build-Xlsm.ps1`** — Bundled PowerShell script to bootstrap a brand-new `.xlsm` from a directory of `.bas` files via Excel COM automation, since `import` requires an existing `.xlsm` and can't create one from scratch. Documented in the README under "Creating a new .xlsm from scratch".

## [0.1.1-alpha.5] - 2026-06-20

### Changed

- **README**: Added a link to the CHANGELOG on GitHub.

## [0.1.1-alpha.4] - 2026-06-16

### Fixed

- **`import` output now opens, recompiles, and supports sheet copy in Excel.** Earlier alpha builds could produce a `vbaProject.bin` that Excel rejected ("unreadable content") or that triggered "Sheet1 cannot be copied", especially after adding or deleting modules. Two root causes were fixed:
  - The cfb.js signature stream (`Sh33tJ5`) was being "removed" by zeroing its CFB directory-entry type after write. That left dangling red-black tree pointers, corrupting the directory. It is now left as a valid stream — Office ignores unknown CFB streams.
  - `_VBA_PROJECT` is now replaced with the canonical 7-byte source-only header (`CC 61 FF FF 00 00 00`, Version `0xFFFF`) and every `__SRP_*` performance-cache stream is deleted, instead of zero-filling the cache at its original size. This matches the EPPlus approach and makes Excel cleanly recompile from source. Module edits — including document modules such as worksheets and `ThisWorkbook` — are reflected correctly.

### Changed

- Adding a brand-new **document/designer module** (UserForm, worksheet, or `ThisWorkbook`) is now rejected, not just UserForms. These are bound to host objects that a `.cls` source cannot recreate; updating their existing code-behind remains supported.
- **Unknown command-line options now error** instead of being silently treated as a positional argument (e.g. `--output foo.xlsm` previously wrote a file literally named `--output`).

## [0.1.1-alpha.3] - 2026-06-15

### Added

- **Full module add/delete support in `import`** — The `import` command now handles adding new standard modules (`.bas`) and class modules (`.cls`) as well as deleting modules that are absent from the source directory. UserForm deletion is also supported (including removal of the binary form-layout storage).
- **P-code cache invalidation** — `_VBA_PROJECT` and all `__SRP_*` streams are now invalidated on import so that Excel recompiles macros from source, preventing crashes from stale bytecode offsets.
- **dir stream rebuild** — The `dir` stream is fully reconstructed after import to reflect module additions and deletions, keeping `MODULEOFFSET` consistent.
- **PROJECT / PROJECTwm update** — Module declarations and workspace entries in the `PROJECT` text stream, and the `PROJECTwm` binary stream, are updated to match the new module list.
- **UserForm add rejected with clear error** — Attempting to add a brand-new UserForm via `import` exits with an informative error message, since the binary form-layout storage cannot be reconstructed from a `.cls` source alone.

## [0.1.1-alpha.2] - 2026-06-14

### Added

- **Confirmation prompt and automatic backup on `import`** — The `import` command now displays a warning and requires explicit `y` confirmation before modifying the Excel file. When the output path is the same as the input, a `.bak` backup is automatically created before overwriting.

### Changed

- All CLI messages are now in English.

## [0.1.1-alpha.1] - 2026-06-14

### Fixed

- **Japanese module names extracted correctly** — Module names containing multi-byte characters (e.g. Japanese) were garbled due to codepage encoding. The extractor now reads the `MODULENAMEUNICODE` record (0x0047) and the Unicode variant of `MODULESTREAMNAME` (0x001A) from the DIR stream, ensuring non-ASCII names are preserved correctly.

## [0.1.0-alpha.0] - 2026-06-12

Initial alpha release.
