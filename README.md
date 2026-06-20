# VBA Runner — VBA Execution Engine + Refactoring Toolkit

Run, test, and statically analyze VBA code without Excel.
Combine it with AI to turn refactoring proposals from "hypotheses" into **verified, running code** on the spot.

> 日本語ドキュメントは [README.ja.md](./README.ja.md) をご覧ください。

## Why VBA Runner?

AI can read VBA code and suggest refactoring improvements, but without VBA Runner you still need to open Excel to verify anything — making each iteration slow and heavy.

| Problem | Without VBA Runner | With VBA Runner |
|---|---|---|
| **Verify proposals** | Manual check in Excel (slow) | Run immediately, verify on the spot |
| **Navigate large codebases** | Token limits prevent loading everything | Pass `vba-analyzer` outlines to the AI |
| **Safety net for refactoring** | None (gut feeling) | Run before changing → auto-generate snapshot tests |
| **Iteration speed** | Edit → paste into Excel → run manually | Edit → run tests (seconds) |

## Packages

| Package | Description | Details |
|---|---|---|
| `vba-runner` (npm library) | VBA execution engine + test runner + CLI tools | [build/runner/README.md](build/runner/README.md) |
| `vba-extractor` (npm CLI) | Export / import VBA source from Office files | [build/extractor/README.md](build/extractor/README.md) |
| VBA Runner (VS Code extension) | LSP integration, debugger, call graph | [build/extension/README.md](build/extension/README.md) |
| VBA Web Runner (Web UI demo) | Try VBA in the browser instantly | [build/playground/README.md](build/playground/README.md) |

## Getting Started

| Goal | Where to start |
|---|---|
| Test VBA functions from TypeScript | [build/runner/README.md](build/runner/README.md) |
| Export VBA source from an Excel file | [build/extractor/README.md](build/extractor/README.md) |
| Ask AI to refactor legacy VBA | [FOR_AI.md](FOR_AI.md) |
| Walk through the full flow: extract → review in VS Code → AI refactor/feature add → write back | [docs/TUTORIAL.md](docs/TUTORIAL.md) |
| Do the same, but by only giving instructions to an AI agent (no command details) | [docs/TUTORIAL_AI.md](docs/TUTORIAL_AI.md) |
| Learn refactoring techniques yourself | [docs/REFACTORING_GUIDE.md](docs/REFACTORING_GUIDE.md) → [docs/INDEX.md](docs/INDEX.md) |
| Edit VBA comfortably in VS Code | [build/extension/README.md](build/extension/README.md) |

## Quick Start

```bash
# 1. Clone the repository
git clone https://github.com/jca02266/vba-runner.git
cd vba-runner

# 2. Install dependencies
npm install

# 3. Run the sample unit tests to verify the setup
npx tsx sample/tests/ts/TaskScheduler_Core.test.ts

# 4. Run all spec tests
npm test
```

Place your VBA source alongside `vba-runner/` and start writing tests.

## Directory Layout

```
vba-runner/
├── src/engine/          VBA Lexer / Parser / Evaluator (core engine)
├── src/extension.ts     VS Code extension entry point
├── src/lsp/             LSP implementation
├── src/App.tsx          Web UI (React)
├── test-libs/           VBARunner class and CLI tools (vba-analyzer, etc.)
├── tests/spec/          VBA language spec tests
├── tests/vba/           Tests for comparing behaviour against real VBA
├── sample/              Refactoring example (TaskScheduler)
│   ├── src/vba/         Production VBA code
│   └── tests/ts/        Sample unit tests (TypeScript)
├── tools/extractor/     vba-extractor CLI implementation
├── build/
│   ├── runner/          npm package distribution
│   ├── extractor/       vba-extractor npm package distribution
│   ├── extension/       VS Code extension distribution (.vsix)
│   └── playground/      VBA Web Runner distribution (Vite)
└── docs/                Tutorials, design guides, and practical examples (14 docs)
```

Example layout when placing VBA source alongside the engine:

```
project-dir/
├── vba-runner/          cloned here
├── massive-vba-project/ VBA source to refactor
│   ├── Module1.bas
│   └── Module2.bas
└── test/                TypeScript test code
    ├── Module1Test/test.ts
    └── Module2Test/test.ts
```

## Documentation

| Document | Audience | Content |
|---|---|---|
| [REFERENCE.md](REFERENCE.md) | Users / developers | Detailed specs: sandbox policy, type system, mocks, extension |
| [FOR_AI.md](FOR_AI.md) | AI-assisted VBA refactoring | AI operation guide, refactoring cycle (Phases 1–7) |
| [LSP.md](LSP.md) | VS Code extension developers | LSP design and implementation spec |
| [docs/](docs/) | VBA testing / refactoring practitioners | End-to-end tutorials, test design principles, refactoring techniques, mock guide (14 docs) |
| [INDEX.md](INDEX.md) | — | Full document index |
