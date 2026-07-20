# vba-extractor

A CLI tool to export and import VBA source code from Office files (`.xlsm`, etc.) without Excel.
Manage VBA modules as plain text files under version control.

**[日本語](./README.ja.md)** | **[CHANGELOG](./CHANGELOG.md)** | **[VBA Runner Project](../../README.md)** | **[Documentation hub](../../docs/README.md)**

> Audience: package users. For goal-oriented guides, see the [documentation hub](../../docs/README.md).

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
vba-extractor import <input.xlsm> <source-dir> [output.xlsm] [--encoding <cp>] [--yes]
```

| Argument | Description |
|---|---|
| `<input.xlsm>` | Target Office file |
| `<source-dir>` | Directory containing the `.bas` / `.cls` files to write back |
| `[output.xlsm]` | Output file (overwrites the input file if omitted) |
| `--encoding <cp>` | Explicit code page. Defaults to the PROJECTCODEPAGE stored in the file |
| `--yes`, `-y` | Skip the confirmation prompt (`Create a backup and run import? [y/N]`). Useful for scripts/AI agents calling `import` repeatedly |

```bash
# Write edited .bas / .cls files back into the xlsm
vba-extractor import MyBook.xlsm src/vba

# Save to a new file, keeping the original intact
vba-extractor import MyBook.xlsm src/vba MyBook_updated.xlsm

# Skip the confirmation prompt (for automation)
vba-extractor import MyBook.xlsm src/vba --yes
```

`.cls` files exported by the VBE's own "Export File..." menu carry a `VERSION 1.0 CLASS` /
`BEGIN ... END` header block (holding component properties like `MultiUse`) that `export`
does not produce and `import` does not expect in the module source itself. `import`
automatically detects and strips this header, so `.cls` files from either source work
as input without manual editing.

### Creating a new .xlsm from scratch (Windows + Excel)

`import` requires an existing `.xlsm` with an existing VBA project (`xl/vbaProject.bin`) - it cannot create one from nothing. An `.xlsm` merely saved as macro-enabled but with no VBA project yet does not contain that part. If you're
starting from `.bas` files only (e.g. AI-generated or written by hand, with no Excel
workbook yet), use the bundled PowerShell script to bootstrap the first `.xlsm` via Excel
COM automation, then `import` to sync in everything else (class modules, sheet code, etc.):

```powershell
# Requires Windows + a locally installed Excel, and Excel's "Trust access to the
# VBA project object model" (AccessVBOM) setting enabled.
powershell -File node_modules/vba-extractor/scripts/Build-Xlsm.ps1 -SourceDir src/vba -OutputPath MyBook.xlsm

# Sync .cls files (and re-sync .bas) into the newly created workbook
vba-extractor import MyBook.xlsm src/vba
```

> If `-OutputPath` is relative, the script resolves it to an absolute path (based on its own
> working directory) before calling `Workbook.SaveAs` — this avoids it being resolved against
> Excel's own default working folder instead, since Excel runs as a separate COM-automated process.

See the script's own `Get-Help -Full` (or its header comment) for details and caveats.

## Output format

| Module type | Extension |
|---|---|
| Standard module | `.bas` |
| Class module / Form / Sheet object | `.cls` |

## Related packages

- [Documentation hub](../../docs/README.md) — Goal-oriented guides (users / developers)
- [vba-runner](https://www.npmjs.com/package/vba-runner) — Run and test VBA code without Excel
