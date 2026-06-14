# Changelog

All notable changes to the `vba-extractor` npm package are documented here.

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
