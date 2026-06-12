# Changelog

All notable changes to the `vba-runner` npm package are documented here.

## [0.1.1-alpha.0] - 2026-06-12

### Fixed

- **Contextual keywords as parameter names** — Keywords such as `text`, `binary`, `compare`, `output`, `lib`, `alias` etc. can now be used as Sub/Function parameter names without a parse error.
- **`!` and `^` as type-declaration characters** — `Dim x!` (Single) and `Dim x^` (LongLong) now work correctly without conflicting with the bang member-access operator (`dict!Key`) or the exponentiation operator (`x^2`).
- **Contextual keyword labels** — Labels such as `Error:` inside procedure bodies are now parsed correctly.

## [0.1.0-alpha.0] - 2026-06-12

Initial alpha release.
