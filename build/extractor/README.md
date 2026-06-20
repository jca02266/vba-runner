# vba-extractor

A CLI tool to export and import VBA source code from Office files (`.xlsm`, etc.) without Excel.
Manage VBA modules as plain text files under version control.

> 日本語ドキュメントは [README.ja.md](./README.ja.md) をご覧ください。

[CHANGELOG](https://github.com/jca02266/vba-runner/blob/main/build/extractor/CHANGELOG.md)

## Installation

```bash
npm install -g vba-extractor
```

## Usage

### export — Extract VBA source from an Office file

```bash
vba-extractor export <input.xlsm> [output-dir] [--encoding <cp>]
```

| Argument | Description |
|---|---|
| `<input.xlsm>` | Target Office file |
| `[output-dir]` | Output directory (defaults to the same directory as the input file) |
| `--encoding <cp>` | Explicit code page (e.g. `cp932`, `cp1252`). Defaults to the PROJECTCODEPAGE stored in the file |

```bash
# Export .bas / .cls files to the current directory
vba-extractor export MyBook.xlsm

# Export to a specific directory
vba-extractor export MyBook.xlsm src/vba
```

### import — Write VBA source back into an Office file

```bash
vba-extractor import <input.xlsm> <source-dir> [output.xlsm] [--encoding <cp>]
```

| Argument | Description |
|---|---|
| `<input.xlsm>` | Target Office file |
| `<source-dir>` | Directory containing the `.bas` / `.cls` files to write back |
| `[output.xlsm]` | Output file (overwrites the input file if omitted) |
| `--encoding <cp>` | Explicit code page. Defaults to the PROJECTCODEPAGE stored in the file |

```bash
# Write edited .bas / .cls files back into the xlsm
vba-extractor import MyBook.xlsm src/vba

# Save to a new file, keeping the original intact
vba-extractor import MyBook.xlsm src/vba MyBook_updated.xlsm
```

## Output format

| Module type | Extension |
|---|---|
| Standard module | `.bas` |
| Class module / Form / Sheet object | `.cls` |

## Related packages

- [vba-runner](https://www.npmjs.com/package/vba-runner) — Run and test VBA code without Excel
