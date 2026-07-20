# VBA Runner — VBA Execution Engine + Refactoring Toolkit

Run, test, and statically analyze VBA code without Excel.
Combine it with AI to turn refactoring proposals from "hypotheses" into **verified, running code** on the spot.

> 日本語ドキュメントは [README.ja.md](./README.ja.md) をご覧ください。

## Why VBA Runner?

AI can read VBA code and suggest refactoring improvements, but without VBA Runner you still need to open Excel to verify anything — making each iteration slow and heavy.

VBA Runner addresses the following problems:

| Problem | Without VBA Runner | With VBA Runner |
|---|---|---|
| **Verify proposals** | Manual check in Excel (slow) | Run immediately, verify on the spot |
| **Navigate large codebases** | Token limits prevent loading everything | Pass `vba-runner analyze` outlines to the AI |
| **Safety net for refactoring** | None (gut feeling) | Run before changing → auto-generate snapshot tests |
| **Iteration speed** | Edit → paste into Excel → run manually | Edit → run tests (seconds) |

## What do you want to do?

| Goal | Start here |
|---|---|
| Run and test VBA from TypeScript | [vba-runner](build/runner/README.md) |
| Extract / write back VBA from Excel files | [vba-extractor](build/extractor/README.md) |
| Edit and inspect VBA in VS Code | [VS Code extension](build/extension/README.md) |
| Try VBA in the browser | [VBA Web Runner](build/playground/README.md) |
| Improve legacy VBA with AI | [Hands-on tutorial](docs/TUTORIAL.md) |
| Browse all guides by goal | [Documentation hub](docs/README.md) |
| Contribute to this repository | [Contributing guide](CONTRIBUTING.md) |

## Packages

| Package | Description | Canonical docs |
|---|---|---|
| `vba-runner` (npm library) | VBA execution engine + test runner + CLI | [build/runner/README.md](build/runner/README.md) |
| `vba-extractor` (npm CLI) | Export / import VBA source from Office files | [build/extractor/README.md](build/extractor/README.md) |
| VBA Runner (VS Code extension) | LSP, debugger, call hierarchy | [build/extension/README.md](build/extension/README.md) |
| VBA Web Runner (Web UI demo) | Try VBA in the browser | [build/playground/README.md](build/playground/README.md) |

For install steps and APIs, follow each package README.
To clone and develop this repository, see [CONTRIBUTING.md](CONTRIBUTING.md).

## Related docs

| Document | Role |
|---|---|
| [docs/README.md](docs/README.md) | Goal-oriented documentation hub |
| [FOR_AI.md](FOR_AI.md) | Instructions for AI working on a user's VBA |
| [REFERENCE.md](REFERENCE.md) | Detailed specs (sandbox, types, mocks, etc.) |
| [CONTRIBUTING.md](CONTRIBUTING.md) | Entry point for repository developers |
