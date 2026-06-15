# Changelog

All notable changes to the `vba-extractor` npm package are documented here.

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
