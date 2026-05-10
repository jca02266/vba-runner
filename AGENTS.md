# Agent Information

## How to Start the Web App (for Development)

If you need to verify operations on the Web UI or test editor features, start the development server using the following commands:

```bash
# Install dependencies (only required the first time)
npm install

# Start the local development server (default: http://localhost:5173/)
npm run dev
```

## How to Run Automated Tests

This project includes a TypeScript-based unit testing environment for verifying the logic of VBA code.
When an agent makes changes to the code, please run the following command to ensure there are no regressions:

```bash
# Bundle and run tests (AST construction and verification)
npx esbuild sample/tests/ts/TaskScheduler_Core.test.ts --bundle --outfile=sample/tests/ts/TaskScheduler_Core.test.cjs --platform=node && node sample/tests/ts/TaskScheduler_Core.test.cjs
```

## File I/O Sandbox Policy

When implementing or modifying file-related functions (e.g., `Open`, `Kill`, `Dir`, `Environ`), you **MUST** follow the Sandbox rules described in `README.md`:

- **All file paths** must be resolved through the `SandboxPath` class in `src/compiler/sandbox.ts`.
- **Sandbox Root**: File operations are restricted to the sandbox root (default: `workspace/`).
- **No Traversal**: Any attempt to access paths outside the sandbox root (e.g., using `../`) must result in a runtime error.
- **Path Virtualization**: Windows-style absolute paths are mapped to subdirectories under the sandbox root (e.g., `C:\foo` -> `{sandboxRoot}/c/foo`).
- **Environment Variables**: The `Environ` function must only access the sandbox's virtual environment, not the host OS environment variables.
